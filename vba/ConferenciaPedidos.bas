Option Explicit

' Macro opcional para consolidar os arquivos de pedido gerados pelo Streamlit.
' Cole este módulo em um arquivo Excel habilitado para macro (.xlsm).

Public Sub ConsolidarPedidosGerados()
    Dim pasta As String
    Dim arquivo As String
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinhaDestino As Long
    Dim ultimaLinhaOrigem As Long
    Dim ultimaColunaOrigem As Long
    Dim primeiraCopia As Boolean
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Selecione a pasta com os pedidos gerados"

    If fd.Show <> -1 Then Exit Sub
    pasta = fd.SelectedItems(1)
    If Right$(pasta, 1) <> "\" Then pasta = pasta & "\"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error Resume Next
    Set wsDestino = ThisWorkbook.Worksheets("Conferencia_Pedidos")
    On Error GoTo 0

    If wsDestino Is Nothing Then
        Set wsDestino = ThisWorkbook.Worksheets.Add
        wsDestino.Name = "Conferencia_Pedidos"
    Else
        wsDestino.Cells.Clear
    End If

    primeiraCopia = True
    ultimaLinhaDestino = 1
    arquivo = Dir$(pasta & "pedido_*.xlsx")

    Do While arquivo <> ""
        Set wbOrigem = Workbooks.Open(Filename:=pasta & arquivo, ReadOnly:=True)
        Set wsOrigem = wbOrigem.Worksheets(1)

        ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, 1).End(xlUp).Row
        ultimaColunaOrigem = wsOrigem.Cells(1, wsOrigem.Columns.Count).End(xlToLeft).Column

        If primeiraCopia Then
            wsDestino.Cells(1, 1).Value = "Arquivo Origem"
            wsOrigem.Range(wsOrigem.Cells(1, 1), wsOrigem.Cells(1, ultimaColunaOrigem)).Copy wsDestino.Cells(1, 2)
            ultimaLinhaDestino = 2
            primeiraCopia = False
        End If

        If ultimaLinhaOrigem >= 2 Then
            wsDestino.Range(wsDestino.Cells(ultimaLinhaDestino, 1), wsDestino.Cells(ultimaLinhaDestino + ultimaLinhaOrigem - 2, 1)).Value = arquivo
            wsOrigem.Range(wsOrigem.Cells(2, 1), wsOrigem.Cells(ultimaLinhaOrigem, ultimaColunaOrigem)).Copy wsDestino.Cells(ultimaLinhaDestino, 2)
            ultimaLinhaDestino = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).Row + 1
        End If

        wbOrigem.Close SaveChanges:=False
        arquivo = Dir$()
    Loop

    wsDestino.Columns.AutoFit
    wsDestino.Rows(1).Font.Bold = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Conferência consolidada com sucesso.", vbInformation, "Pedidos"
End Sub
