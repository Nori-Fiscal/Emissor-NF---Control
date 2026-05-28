from __future__ import annotations

import math
import re
import unicodedata
from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
from typing import Dict, Iterable, List, Optional, Tuple
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# Configuração geral
# ============================================================
MOEDA_Q = Decimal("0.01")
PRECO_Q = Decimal("0.0001")

CANON_PEDIDO = {
    "numero_pedido": ["Número pedido", "Numero pedido", "Pedido", "Nº Pedido", "N Pedido"],
    "nome_comprador": ["Nome Comprador", "Comprador"],
    "sku": ["SKU", "Código", "Codigo", "Cód. Produto", "Cod Produto"],
    "produto": ["Produto", "Descrição", "Descricao"],
    "unidade": ["Un", "Unidade"],
    "quantidade": ["Quantidade", "Qtd", "Qtde"],
    "valor_unitario": ["Valor Unitário", "Valor Unitario", "Preço", "Preco", "Valor Unit"],
    "valor_total": ["Valor Total", "Total Item", "Total do Item"],
    "total_pedido": ["Total Pedido", "Valor Pedido", "Total Geral"],
    "observacoes": ["Observações", "Observacoes", "Obs"],
}

CANON_PRECO = {
    "nome_lista": ["Nome lista", "Lista", "Tabela", "Nome da lista"],
    "sku": ["SKU", "Código", "Codigo", "Cód. Produto", "Cod Produto"],
    "preco": ["Preço", "Preco", "Valor", "Valor Unitário", "Valor Unitario"],
}


@dataclass
class ItemPedido:
    indice: int
    linha_excel: int
    valores_linha: List[object]
    sku: str
    quantidade: int
    preco_unitario: Decimal
    valor_total: Decimal


@dataclass
class ResultadoDivisao:
    sucesso: bool
    mensagem: str
    alocacao: List[Dict[int, int]]
    totais: List[Decimal]
    avisos: List[str]


# ============================================================
# Utilitários de leitura / normalização
# ============================================================
def normalizar_texto(valor: object) -> str:
    if valor is None:
        return ""
    texto = str(valor).strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = texto.lower()
    texto = re.sub(r"[^a-z0-9]+", "", texto)
    return texto


def mapa_alias(canon: Dict[str, List[str]]) -> Dict[str, str]:
    saida = {}
    for chave, aliases in canon.items():
        for alias in aliases + [chave]:
            saida[normalizar_texto(alias)] = chave
    return saida


ALIAS_PEDIDO = mapa_alias(CANON_PEDIDO)
ALIAS_PRECO = mapa_alias(CANON_PRECO)


def para_decimal(valor: object, campo: str = "valor") -> Decimal:
    if valor is None or valor == "":
        raise ValueError(f"Campo {campo} vazio.")
    if isinstance(valor, Decimal):
        return valor
    texto = str(valor).strip()
    # Aceita formatos pt-BR simples: 1.234,56 ou 1234,56.
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return Decimal(texto)
    except InvalidOperation as exc:
        raise ValueError(f"Não foi possível converter {campo}: {valor!r}") from exc


def fmt_moeda(valor: Decimal) -> str:
    valor = valor.quantize(MOEDA_Q, rounding=ROUND_HALF_UP)
    s = f"{valor:,.2f}"
    return "R$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")


def localizar_cabecalho(ws, alias_map: Dict[str, str], obrigatorios: Iterable[str], max_linhas: int = 12) -> Tuple[int, Dict[str, int]]:
    melhor_linha = None
    melhor_mapa = {}
    melhor_score = -1
    obrigatorios = set(obrigatorios)

    for row in range(1, min(ws.max_row, max_linhas) + 1):
        mapa = {}
        for col in range(1, ws.max_column + 1):
            chave = alias_map.get(normalizar_texto(ws.cell(row, col).value))
            if chave and chave not in mapa:
                mapa[chave] = col
        score = len(obrigatorios.intersection(mapa.keys())) * 10 + len(mapa)
        if score > melhor_score:
            melhor_score = score
            melhor_linha = row
            melhor_mapa = mapa

    faltantes = obrigatorios.difference(melhor_mapa.keys())
    if faltantes:
        faltantes_txt = ", ".join(sorted(faltantes))
        raise ValueError(f"Cabeçalho não localizado ou colunas obrigatórias ausentes: {faltantes_txt}.")

    return melhor_linha or 1, melhor_mapa


def listar_abas_e_listas_preco(preco_bytes: bytes) -> Tuple[List[str], List[str]]:
    wb = load_workbook(BytesIO(preco_bytes), data_only=True, read_only=True)
    nomes_listas = set()
    for ws in wb.worksheets:
        try:
            header_row, colmap = localizar_cabecalho(ws, ALIAS_PRECO, ["sku", "preco"])
        except Exception:
            continue
        col_lista = colmap.get("nome_lista")
        if col_lista:
            for row in range(header_row + 1, ws.max_row + 1):
                valor = ws.cell(row, col_lista).value
                if valor not in (None, ""):
                    nomes_listas.add(str(valor).strip())
    return wb.sheetnames, sorted(nomes_listas)


def carregar_precos(preco_bytes: bytes, aba: str, nome_lista: Optional[str]) -> Dict[str, Decimal]:
    wb = load_workbook(BytesIO(preco_bytes), data_only=True, read_only=True)
    ws = wb[aba]
    header_row, colmap = localizar_cabecalho(ws, ALIAS_PRECO, ["sku", "preco"])

    precos: Dict[str, Decimal] = {}
    for row in range(header_row + 1, ws.max_row + 1):
        sku_raw = ws.cell(row, colmap["sku"]).value
        preco_raw = ws.cell(row, colmap["preco"]).value
        if sku_raw in (None, "") and preco_raw in (None, ""):
            continue

        if nome_lista and colmap.get("nome_lista"):
            lista_atual = ws.cell(row, colmap["nome_lista"]).value
            if str(lista_atual).strip() != nome_lista:
                continue

        if sku_raw in (None, ""):
            continue
        sku = str(sku_raw).strip()
        if preco_raw in (None, ""):
            continue
        preco = para_decimal(preco_raw, f"preço do SKU {sku}").quantize(PRECO_Q, rounding=ROUND_HALF_UP)
        precos[sku] = preco

    if not precos:
        raise ValueError("Nenhum preço válido foi encontrado na lista selecionada.")
    return precos


def carregar_itens_pedido(pedido_bytes: bytes, aba: str, precos: Dict[str, Decimal]) -> Tuple[int, Dict[str, int], List[ItemPedido], List[Dict[str, object]]]:
    wb = load_workbook(BytesIO(pedido_bytes), data_only=False, read_only=False)
    ws = wb[aba]
    header_row, colmap = localizar_cabecalho(ws, ALIAS_PEDIDO, ["numero_pedido", "sku", "quantidade"])

    itens: List[ItemPedido] = []
    pendencias: List[Dict[str, object]] = []

    for row in range(header_row + 1, ws.max_row + 1):
        sku_raw = ws.cell(row, colmap["sku"]).value
        qtd_raw = ws.cell(row, colmap["quantidade"]).value

        if sku_raw in (None, "") and qtd_raw in (None, ""):
            continue
        if sku_raw in (None, ""):
            pendencias.append({"Linha": row, "SKU": "", "Problema": "Linha com quantidade, mas sem SKU."})
            continue

        sku = str(sku_raw).strip()
        if sku not in precos:
            pendencias.append({"Linha": row, "SKU": sku, "Problema": "SKU não encontrado na lista de preços selecionada."})
            continue

        try:
            qtd_dec = para_decimal(qtd_raw, f"quantidade da linha {row}")
        except Exception as exc:
            pendencias.append({"Linha": row, "SKU": sku, "Problema": str(exc)})
            continue

        if qtd_dec <= 0:
            pendencias.append({"Linha": row, "SKU": sku, "Problema": "Quantidade menor ou igual a zero."})
            continue

        if qtd_dec != qtd_dec.to_integral_value():
            pendencias.append({"Linha": row, "SKU": sku, "Problema": "Quantidade fracionária; ajuste manualmente ou trate em pedido separado."})
            continue

        quantidade = int(qtd_dec)
        preco = precos[sku]
        total = (preco * Decimal(quantidade)).quantize(MOEDA_Q, rounding=ROUND_HALF_UP)
        valores_linha = [ws.cell(row, col).value for col in range(1, ws.max_column + 1)]
        itens.append(
            ItemPedido(
                indice=len(itens),
                linha_excel=row,
                valores_linha=valores_linha,
                sku=sku,
                quantidade=quantidade,
                preco_unitario=preco,
                valor_total=total,
            )
        )

    return header_row, colmap, itens, pendencias


# ============================================================
# Algoritmo de divisão operacional
# ============================================================
def ceil_decimal(a: Decimal, b: Decimal) -> int:
    return int((a / b).to_integral_value(rounding="ROUND_CEILING"))


def floor_decimal(a: Decimal, b: Decimal) -> int:
    return int((a / b).to_integral_value(rounding="ROUND_FLOOR"))


def totais_por_alocacao(alocacao: List[Dict[int, int]], itens: List[ItemPedido]) -> List[Decimal]:
    totais = []
    for pedido in alocacao:
        total = Decimal("0")
        for idx, qtd in pedido.items():
            total += itens[idx].preco_unitario * Decimal(qtd)
        totais.append(total.quantize(MOEDA_Q, rounding=ROUND_HALF_UP))
    return totais


def distribuir_itens(itens: List[ItemPedido], minimo: Decimal, maximo: Decimal) -> ResultadoDivisao:
    avisos: List[str] = []
    if not itens:
        return ResultadoDivisao(False, "Nenhum item válido para dividir.", [], [], avisos)
    if minimo <= 0 or maximo <= 0 or minimo > maximo:
        return ResultadoDivisao(False, "Informe uma faixa válida: mínimo maior que zero e menor ou igual ao máximo.", [], [], avisos)

    total_geral = sum((item.preco_unitario * Decimal(item.quantidade) for item in itens), Decimal("0")).quantize(MOEDA_Q, rounding=ROUND_HALF_UP)
    if total_geral < minimo:
        return ResultadoDivisao(
            False,
            f"Total do pedido ({fmt_moeda(total_geral)}) é menor que o mínimo informado ({fmt_moeda(minimo)}).",
            [],
            [],
            avisos,
        )

    item_unitario_maior = [item for item in itens if item.preco_unitario > maximo]
    if item_unitario_maior:
        skus = ", ".join(sorted({item.sku for item in item_unitario_maior})[:10])
        return ResultadoDivisao(False, f"Há SKU com preço unitário acima do máximo por pedido/NF: {skus}.", [], [], avisos)

    n_min = ceil_decimal(total_geral, maximo)
    n_max = floor_decimal(total_geral, minimo)
    if n_min > n_max:
        return ResultadoDivisao(
            False,
            "Não existe uma quantidade de pedidos/NFs que deixe todos os totais dentro da faixa informada. "
            f"Total geral: {fmt_moeda(total_geral)}; faixa: {fmt_moeda(minimo)} a {fmt_moeda(maximo)}.",
            [],
            [],
            avisos,
        )

    alvo = (minimo + maximo) / Decimal("2")
    n_sugerido = int((total_geral / alvo).to_integral_value(rounding="ROUND_HALF_UP"))
    n = max(n_min, min(n_sugerido, n_max))
    n = max(1, n)

    # Alocação inicial proporcional: preserva preços e quantidades totais.
    alocacao: List[Dict[int, int]] = [defaultdict(int) for _ in range(n)]
    totais = [Decimal("0") for _ in range(n)]

    # Ordena por maior impacto para balancear melhor desde o início.
    itens_ordenados = sorted(itens, key=lambda x: (x.preco_unitario, x.valor_total, x.sku), reverse=True)
    for item in itens_ordenados:
        base = item.quantidade // n
        resto = item.quantidade % n
        if base:
            for b in range(n):
                alocacao[b][item.indice] += base
                totais[b] += item.preco_unitario * Decimal(base)
        for _ in range(resto):
            b = min(range(n), key=lambda i: totais[i])
            alocacao[b][item.indice] += 1
            totais[b] += item.preco_unitario

    totais = [t.quantize(MOEDA_Q, rounding=ROUND_HALF_UP) for t in totais]

    # Rebalanceamento por movimentação de unidades entre pedidos.
    # Objetivo: manter todos dentro da faixa sem alterar preço unitário nem quantidade total.
    max_iter = 50000
    for _ in range(max_iter):
        totais = totais_por_alocacao(alocacao, itens)
        baixos = [i for i, t in enumerate(totais) if t < minimo]
        altos = [i for i, t in enumerate(totais) if t > maximo]
        if not baixos and not altos:
            return ResultadoDivisao(True, f"Divisão concluída em {n} pedido(s)/NF(s).", [dict(x) for x in alocacao], totais, avisos)

        movimento = False

        # Primeiro corrige pedidos abaixo do mínimo.
        if baixos:
            destino = min(baixos, key=lambda i: totais[i])
            necessidade = minimo - totais[destino]
            doadores = sorted([i for i in range(n) if i != destino and totais[i] > minimo], key=lambda i: totais[i], reverse=True)
            melhor = None
            melhor_score = None
            for origem in doadores:
                for idx, qtd_origem in alocacao[origem].items():
                    if qtd_origem <= 0:
                        continue
                    p = itens[idx].preco_unitario
                    if totais[destino] + p > maximo:
                        continue
                    if totais[origem] - p < minimo and totais[origem] <= maximo:
                        continue
                    # Aproxima o destino do mínimo sem estourar o máximo.
                    score = abs(necessidade - p)
                    if melhor is None or score < melhor_score:
                        melhor = (origem, destino, idx)
                        melhor_score = score
            if melhor:
                origem, destino, idx = melhor
                alocacao[origem][idx] -= 1
                if alocacao[origem][idx] == 0:
                    del alocacao[origem][idx]
                alocacao[destino][idx] += 1
                movimento = True

        if movimento:
            continue

        # Depois corrige pedidos acima do máximo.
        if altos:
            origem = max(altos, key=lambda i: totais[i])
            destinos = sorted([i for i in range(n) if i != origem and totais[i] < maximo], key=lambda i: totais[i])
            melhor = None
            melhor_score = None
            for destino in destinos:
                capacidade = maximo - totais[destino]
                for idx, qtd_origem in alocacao[origem].items():
                    if qtd_origem <= 0:
                        continue
                    p = itens[idx].preco_unitario
                    if p > capacidade:
                        continue
                    score = abs((totais[origem] - maximo) - p)
                    if melhor is None or score < melhor_score:
                        melhor = (origem, destino, idx)
                        melhor_score = score
            if melhor:
                origem, destino, idx = melhor
                alocacao[origem][idx] -= 1
                if alocacao[origem][idx] == 0:
                    del alocacao[origem][idx]
                alocacao[destino][idx] += 1
                movimento = True

        if not movimento:
            break

    totais = totais_por_alocacao(alocacao, itens)
    fora = [fmt_moeda(t) for t in totais if not (minimo <= t <= maximo)]
    return ResultadoDivisao(
        False,
        "Não foi possível balancear todos os pedidos dentro da faixa usando quantidades inteiras. "
        f"Totais fora da faixa: {', '.join(fora[:8])}.",
        [dict(x) for x in alocacao],
        totais,
        avisos,
    )


# ============================================================
# Geração dos arquivos Excel
# ============================================================
def copiar_estilo(celula_origem, celula_destino) -> None:
    if celula_origem.has_style:
        celula_destino._style = copy(celula_origem._style)
    if celula_origem.number_format:
        celula_destino.number_format = celula_origem.number_format
    if celula_origem.alignment:
        celula_destino.alignment = copy(celula_origem.alignment)


def nome_pedido(base: str, seq: int, mascara: str) -> str:
    base = str(base).strip() if base not in (None, "") else "PEDIDO"
    try:
        return mascara.format(base=base, seq=seq, seq2=f"{seq:02d}", seq3=f"{seq:03d}")
    except Exception:
        return f"{base}-{seq:02d}"


def gerar_planilha_pedido(
    pedido_bytes: bytes,
    aba: str,
    header_row: int,
    colmap: Dict[str, int],
    itens: List[ItemPedido],
    alocacao_pedido: Dict[int, int],
    total_pedido: Decimal,
    numero_pedido_saida: str,
) -> bytes:
    wb = load_workbook(BytesIO(pedido_bytes), data_only=False)
    ws = wb[aba]
    max_col = ws.max_column
    linha_modelo = header_row + 1

    estilo_modelo = []
    for col in range(1, max_col + 1):
        estilo_modelo.append(copy(ws.cell(linha_modelo, col)._style))
    altura_modelo = ws.row_dimensions[linha_modelo].height

    if ws.max_row > header_row:
        ws.delete_rows(header_row + 1, ws.max_row - header_row)

    linha_destino = header_row + 1
    for idx in sorted(alocacao_pedido.keys(), key=lambda i: itens[i].linha_excel):
        qtd = alocacao_pedido[idx]
        if qtd <= 0:
            continue
        item = itens[idx]
        valores = list(item.valores_linha[:max_col])
        while len(valores) < max_col:
            valores.append(None)

        valores[colmap["numero_pedido"] - 1] = numero_pedido_saida
        valores[colmap["sku"] - 1] = item.sku
        valores[colmap["quantidade"] - 1] = qtd
        if colmap.get("valor_unitario"):
            valores[colmap["valor_unitario"] - 1] = float(item.preco_unitario)
        if colmap.get("valor_total"):
            valores[colmap["valor_total"] - 1] = float((item.preco_unitario * Decimal(qtd)).quantize(MOEDA_Q, rounding=ROUND_HALF_UP))
        if colmap.get("total_pedido"):
            valores[colmap["total_pedido"] - 1] = float(total_pedido.quantize(MOEDA_Q, rounding=ROUND_HALF_UP))

        for col, valor in enumerate(valores, start=1):
            cell = ws.cell(linha_destino, col)
            cell.value = valor
            if col - 1 < len(estilo_modelo):
                cell._style = copy(estilo_modelo[col - 1])
        if altura_modelo:
            ws.row_dimensions[linha_destino].height = altura_modelo
        linha_destino += 1

    # Formatos de número relevantes.
    if colmap.get("valor_unitario"):
        col = get_column_letter(colmap["valor_unitario"])
        for row in range(header_row + 1, linha_destino):
            ws[f"{col}{row}"].number_format = 'R$ #,##0.0000'
    if colmap.get("valor_total"):
        col = get_column_letter(colmap["valor_total"])
        for row in range(header_row + 1, linha_destino):
            ws[f"{col}{row}"].number_format = 'R$ #,##0.00'
    if colmap.get("total_pedido"):
        col = get_column_letter(colmap["total_pedido"])
        for row in range(header_row + 1, linha_destino):
            ws[f"{col}{row}"].number_format = 'R$ #,##0.00'
    if colmap.get("quantidade"):
        col = get_column_letter(colmap["quantidade"])
        for row in range(header_row + 1, linha_destino):
            ws[f"{col}{row}"].number_format = '0'

    saida = BytesIO()
    wb.save(saida)
    return saida.getvalue()


def gerar_auditoria(
    itens: List[ItemPedido],
    alocacao: List[Dict[int, int]],
    totais: List[Decimal],
    minimo: Decimal,
    maximo: Decimal,
    numeros_pedido: List[str],
    pendencias: Optional[List[Dict[str, object]]] = None,
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo_NFs"
    ws.append(["Pedido/NF", "Valor total", "Dentro da faixa", "Qtd. linhas", "Qtd. peças"])
    for i, pedido in enumerate(alocacao):
        qtd_pecas = sum(pedido.values())
        dentro = "Sim" if minimo <= totais[i] <= maximo else "Não"
        ws.append([numeros_pedido[i], float(totais[i]), dentro, len(pedido), qtd_pecas])
    ws.append([])
    ws.append(["Total geral", float(sum(totais, Decimal("0"))), "", "", sum(sum(p.values()) for p in alocacao)])

    ws2 = wb.create_sheet("Itens")
    ws2.append(["Pedido/NF", "Linha original", "SKU", "Quantidade", "Preço unitário", "Valor total item"])
    for i, pedido in enumerate(alocacao):
        for idx in sorted(pedido.keys(), key=lambda x: itens[x].linha_excel):
            qtd = pedido[idx]
            item = itens[idx]
            ws2.append([
                numeros_pedido[i],
                item.linha_excel,
                item.sku,
                qtd,
                float(item.preco_unitario),
                float((item.preco_unitario * Decimal(qtd)).quantize(MOEDA_Q, rounding=ROUND_HALF_UP)),
            ])

    if pendencias:
        ws3 = wb.create_sheet("Pendencias")
        ws3.append(["Linha", "SKU", "Problema"])
        for p in pendencias:
            ws3.append([p.get("Linha"), p.get("SKU"), p.get("Problema")])

    for wsx in wb.worksheets:
        for row in wsx.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="center")
        for cell in wsx[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        for col in range(1, wsx.max_column + 1):
            letra = get_column_letter(col)
            largura = max(12, min(38, max(len(str(wsx.cell(row, col).value or "")) for row in range(1, wsx.max_row + 1)) + 2))
            wsx.column_dimensions[letra].width = largura
        for row in range(2, wsx.max_row + 1):
            for col in range(1, wsx.max_column + 1):
                header = str(wsx.cell(1, col).value or "").lower()
                if "valor" in header or "preço" in header:
                    wsx.cell(row, col).number_format = 'R$ #,##0.00'
        wsx.freeze_panes = "A2"

    saida = BytesIO()
    wb.save(saida)
    return saida.getvalue()


def gerar_pendencias_xlsx(pendencias: List[Dict[str, object]]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Pendencias"
    ws.append(["Linha", "SKU", "Problema"])
    for p in pendencias:
        ws.append([p.get("Linha"), p.get("SKU"), p.get("Problema")])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="C00000")
    for col in range(1, ws.max_column + 1):
        letra = get_column_letter(col)
        ws.column_dimensions[letra].width = 22 if col < 3 else 70
    saida = BytesIO()
    wb.save(saida)
    return saida.getvalue()


def montar_zip(
    pedido_bytes: bytes,
    aba_pedido: str,
    header_row: int,
    colmap: Dict[str, int],
    itens: List[ItemPedido],
    resultado: ResultadoDivisao,
    base_numero_pedido: str,
    mascara: str,
    minimo: Decimal,
    maximo: Decimal,
) -> bytes:
    zip_buffer = BytesIO()
    numeros = [nome_pedido(base_numero_pedido, i + 1, mascara) for i in range(len(resultado.alocacao))]
    with ZipFile(zip_buffer, "w", compression=ZIP_DEFLATED) as zf:
        for i, aloc in enumerate(resultado.alocacao):
            arquivo = gerar_planilha_pedido(
                pedido_bytes=pedido_bytes,
                aba=aba_pedido,
                header_row=header_row,
                colmap=colmap,
                itens=itens,
                alocacao_pedido=aloc,
                total_pedido=resultado.totais[i],
                numero_pedido_saida=numeros[i],
            )
            nome_arquivo = f"pedido_{i + 1:02d}_{numeros[i]}.xlsx".replace("/", "-").replace("\\", "-")
            zf.writestr(nome_arquivo, arquivo)

        auditoria = gerar_auditoria(
            itens=itens,
            alocacao=resultado.alocacao,
            totais=resultado.totais,
            minimo=minimo,
            maximo=maximo,
            numeros_pedido=numeros,
        )
        zf.writestr("AUDITORIA_desmembramento.xlsx", auditoria)
    return zip_buffer.getvalue()


# ============================================================
# Interface Streamlit
# ============================================================
st.set_page_config(page_title="Gerador de pedidos desmembrados", layout="wide")

st.title("Gerador de pedidos desmembrados por faixa de valor")
st.caption("Mantém preço unitário, quantidade total e layout do modelo. Gera uma trilha de auditoria para conferência.")

with st.expander("Critérios usados pelo sistema", expanded=False):
    st.write(
        "O sistema divide o pedido por composição de itens, sem alterar preço unitário e sem alterar a quantidade total. "
        "A faixa de valor é validada matematicamente; quando não for possível atingir todos os pedidos dentro do intervalo informado, "
        "o sistema bloqueia a geração e mostra o motivo."
    )

col1, col2 = st.columns(2)
with col1:
    pedido_file = st.file_uploader("Planilha de pedido padronizada (.xlsx)", type=["xlsx"])
with col2:
    preco_file = st.file_uploader("Planilha de preços (.xlsx)", type=["xlsx"])

min_col, max_col, suf_col = st.columns([1, 1, 2])
with min_col:
    valor_min = st.number_input("Valor mínimo por pedido/NF", min_value=0.01, value=110000.00, step=1000.00, format="%.2f")
with max_col:
    valor_max = st.number_input("Valor máximo por pedido/NF", min_value=0.01, value=125000.00, step=1000.00, format="%.2f")
with suf_col:
    mascara_pedido = st.text_input(
        "Máscara do número do pedido de saída",
        value="{base}-{seq2}",
        help="Use {base}, {seq}, {seq2} ou {seq3}. Ex.: {base}-{seq2} gera 37-01, 37-02...",
    )

aba_pedido = None
aba_preco = None
nome_lista = None

if pedido_file:
    pedido_bytes_ui = pedido_file.getvalue()
    wb_pedido_ro = load_workbook(BytesIO(pedido_bytes_ui), read_only=True, data_only=True)
    aba_pedido = st.selectbox("Aba do pedido", wb_pedido_ro.sheetnames, index=0)

if preco_file:
    preco_bytes_ui = preco_file.getvalue()
    abas_preco, nomes_listas = listar_abas_e_listas_preco(preco_bytes_ui)
    aba_preco = st.selectbox("Aba da lista de preços", abas_preco, index=0)
    if nomes_listas:
        opcoes = ["Todas as listas"] + nomes_listas
        escolha = st.selectbox("Nome da lista de preços", opcoes, index=1 if len(opcoes) > 1 else 0)
        nome_lista = None if escolha == "Todas as listas" else escolha

processar = st.button("Gerar pedidos", type="primary", disabled=not (pedido_file and preco_file and aba_pedido and aba_preco))

if processar:
    try:
        pedido_bytes = pedido_file.getvalue()
        preco_bytes = preco_file.getvalue()
        minimo = para_decimal(valor_min, "valor mínimo").quantize(MOEDA_Q, rounding=ROUND_HALF_UP)
        maximo = para_decimal(valor_max, "valor máximo").quantize(MOEDA_Q, rounding=ROUND_HALF_UP)

        precos = carregar_precos(preco_bytes, aba_preco, nome_lista)
        header_row, colmap, itens, pendencias = carregar_itens_pedido(pedido_bytes, aba_pedido, precos)

        st.subheader("Conferência inicial")
        total_lido = sum((i.valor_total for i in itens), Decimal("0")).quantize(MOEDA_Q, rounding=ROUND_HALF_UP)
        c1, c2, c3 = st.columns(3)
        c1.metric("Itens válidos", len(itens))
        c2.metric("Total precificado", fmt_moeda(total_lido))
        c3.metric("Pendências", len(pendencias))

        if pendencias:
            st.error("Existem pendências antes da divisão. Corrija a lista de preços ou o pedido e tente novamente.")
            df_pend = pd.DataFrame(pendencias)
            st.dataframe(df_pend, use_container_width=True)
            st.download_button(
                "Baixar pendências em Excel",
                data=gerar_pendencias_xlsx(pendencias),
                file_name="pendencias_skus.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.stop()

        resultado = distribuir_itens(itens, minimo, maximo)
        if not resultado.sucesso:
            st.error(resultado.mensagem)
            if resultado.totais:
                st.dataframe(
                    pd.DataFrame({
                        "Pedido/NF": [i + 1 for i in range(len(resultado.totais))],
                        "Total": [fmt_moeda(t) for t in resultado.totais],
                        "Dentro da faixa": ["Sim" if minimo <= t <= maximo else "Não" for t in resultado.totais],
                    }),
                    use_container_width=True,
                )
            st.stop()

        # Número base do pedido original.
        col_num = colmap["numero_pedido"]
        wb_tmp = load_workbook(BytesIO(pedido_bytes), data_only=True, read_only=True)
        ws_tmp = wb_tmp[aba_pedido]
        base_numero = ws_tmp.cell(header_row + 1, col_num).value or "PEDIDO"

        numeros = [nome_pedido(str(base_numero), i + 1, mascara_pedido) for i in range(len(resultado.alocacao))]
        resumo = pd.DataFrame({
            "Pedido/NF": numeros,
            "Valor total": [fmt_moeda(t) for t in resultado.totais],
            "Qtd. linhas": [len(a) for a in resultado.alocacao],
            "Qtd. peças": [sum(a.values()) for a in resultado.alocacao],
            "Dentro da faixa": ["Sim" if minimo <= t <= maximo else "Não" for t in resultado.totais],
        })

        st.success(resultado.mensagem)
        st.dataframe(resumo, use_container_width=True)

        zip_bytes = montar_zip(
            pedido_bytes=pedido_bytes,
            aba_pedido=aba_pedido,
            header_row=header_row,
            colmap=colmap,
            itens=itens,
            resultado=resultado,
            base_numero_pedido=str(base_numero),
            mascara=mascara_pedido,
            minimo=minimo,
            maximo=maximo,
        )
        st.download_button(
            "Baixar ZIP com pedidos e auditoria",
            data=zip_bytes,
            file_name="pedidos_desmembrados.zip",
            mime="application/zip",
        )

    except Exception as exc:
        st.exception(exc)
