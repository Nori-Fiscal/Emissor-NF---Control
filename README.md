# Sistema Streamlit - Gerador de pedidos desmembrados por faixa de valor

Este sistema recebe:

1. a planilha de pedido no modelo padronizado;  
2. a planilha de preços com `SKU` e `Preço`.

Ele gera um arquivo ZIP com pedidos em `.xlsx`, mantendo o layout do modelo enviado e preenchendo os campos de valor com base na lista de preços.

## Regras implementadas

- Mantém o preço unitário da lista de preços.
- Mantém a quantidade total por SKU/linha do pedido original.
- Calcula `Valor Unitário`, `Valor Total` e `Total Pedido` nos arquivos de saída.
- Divide os itens em pedidos/NFs dentro da faixa configurada, por padrão de R$ 110.000,00 a R$ 125.000,00.
- Bloqueia a geração quando houver SKU sem preço, quantidade inválida ou quando a faixa for matematicamente impossível para o total do pedido.
- Gera `AUDITORIA_desmembramento.xlsx` com resumo dos pedidos e itens alocados.

O sistema não altera o arquivo original enviado; ele usa o arquivo como modelo e cria novas cópias.

## Instalação

Crie uma pasta para o projeto, copie os arquivos deste pacote e execute:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

No Linux/Mac:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Como usar

1. Abra o Streamlit no navegador.
2. Faça upload do pedido padronizado.
3. Faça upload da lista de preços.
4. Selecione a aba do pedido, a aba da lista e o nome da lista, quando houver.
5. Confirme a faixa mínima e máxima por pedido/NF.
6. Clique em **Gerar pedidos**.
7. Baixe o ZIP com os arquivos `.xlsx` e a auditoria.

## Colunas esperadas

### Pedido

- `Número pedido`
- `SKU`
- `Quantidade`
- Opcional, mas preenchidas quando existirem: `Valor Unitário`, `Valor Total`, `Total Pedido`

### Lista de preços

- `SKU`
- `Preço`
- Opcional: `Nome lista`

## Observação operacional e fiscal

Use o sistema apenas para divisão operacional legítima de pedidos, com preços e quantidades preservados e com trilha de auditoria. Valide os arquivos de saída com o responsável fiscal/contábil antes de emissão de NF.
