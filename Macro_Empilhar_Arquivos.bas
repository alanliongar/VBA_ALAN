Attribute VB_Name = "M�dulo1"
'Macro para compilar todas as abas de arquivos Excel em um �nico arquivo, empilhando os dados
Sub EmpilharTodosArquivos()
    Dim pasta As String, arquivo As String
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinha As Long, linhaInicio As Long
    Dim maxLinhas As Long: maxLinhas = 1048576 ' Limite da aba no Excel
    Dim abaIndex As Integer: abaIndex = 1

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Pasta que cont�m os arquivos Excel
    pasta = "C:\Users\alugomes\OneDrive - rd.com.br\�rea de Trabalho\Compilado da jadlog\"
    arquivo = Dir(pasta & "*.xls*")

    ' Criar a primeira aba de destino
    Set wsDestino = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
    wsDestino.Name = "Empilhado_" & abaIndex
    linhaInicio = 1

    Do While arquivo <> ""
        If arquivo <> ThisWorkbook.Name Then
            Set wbOrigem = Workbooks.Open(pasta & arquivo)
            Set wsOrigem = wbOrigem.Sheets(1) ' Considera apenas a primeira aba de cada arquivo

            ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, 1).End(xlUp).Row

            ' Se a aba atual n�o tiver espa�o suficiente, cria nova aba
            If linhaInicio + ultimaLinha - 1 > maxLinhas Then
                abaIndex = abaIndex + 1
                Set wsDestino = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
                wsDestino.Name = "Empilhado_" & abaIndex
                linhaInicio = 1
            End If

            ' Copiar e colar os dados
            wsOrigem.Range("A1").Resize(ultimaLinha, wsOrigem.UsedRange.Columns.Count).Copy
            wsDestino.Cells(linhaInicio, 1).PasteSpecial Paste:=xlPasteValues

            linhaInicio = linhaInicio + ultimaLinha
            wbOrigem.Close SaveChanges:=False
        End If
        arquivo = Dir
    Loop

    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Dados empilhados com sucesso!", vbInformation
End Sub

