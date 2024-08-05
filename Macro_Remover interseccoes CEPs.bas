Attribute VB_Name = "Módulo1"
Sub Ajuste_de_Faixa_REMOVER()
Sheets("Remover").Select
Dim i As Long
Dim j As Long
i = 2
While Cells(i, 1) <> ""
i = i + 1
If Cells(i, 2).Value < Cells(i, 1).Value Then
MsgBox ("interrompa a macro e corrija o erro na linha " + CStr(i) + " da prioridade. A Coluna B é menor que a coluna A.")
End If
Wend
i = i - 1
'código da testeeeeeeeeeee aqui



j = 2
While Cells(j, 4) <> ""
j = j + 1
If Cells(j, 5).Value < Cells(j, 4).Value Then
MsgBox ("interrompa a macro e corrija o erro na linha " + CStr(j) + " da alteração. A Coluna E é menor que a coluna D.")
End If
Wend
j = j - 1
'término do ordenamento
        i = 2
        While Cells(i, 1) <> ""
        j = 2
                While Cells(j, 4) <> ""

If Cells(i, 1).Value <= Cells(j, 4).Value And Cells(i, 2).Value < Cells(j, 5).Value And Cells(i, 2).Value >= Cells(j, 4).Value Then ' BB>CC Then 'AA <= CC And BB < DD Then
    Cells(j, 4) = Cells(i, 2).Value + 1
    'Cells(j, 6) = "Linha modificada na coluna D, linha " + CStr(i)
    'avaliado
End If


If Cells(i, 1).Value <= Cells(j, 4).Value And Cells(i, 2).Value >= Cells(j, 5).Value Then 'AA <= CC And BB >= DD Then
    Range("C" + CStr(j) + ":G" + CStr(j)).Select 'C, D e E - se colocar o código da filial.
    Selection.Delete Shift:=xlUp
    j = j - 1
    'Cells(j, 6) = "faixa inteira contida na faixa da linha " + CStr(i)
End If


If Cells(i, 1).Value > Cells(j, 4).Value And Cells(i, 2).Value >= Cells(j, 5).Value And Cells(i, 1).Value <= Cells(j, 5).Value Then 'AA > CC And BB >= DD and AA <= DD Then
    Cells(j, 5) = Cells(i, 1).Value - 1
    'Cells(j, 6) = "Linha modificada na coluna E, linha " + CStr(i)
    'avaliado
End If


If Cells(i, 1).Value > Cells(j, 4).Value And Cells(i, 2).Value < Cells(j, 5).Value Then 'AA > CC And BB < DD Then
    Range("C" + CStr(j) + ":G" + CStr(j)).Select 'C, D e E
    Selection.Insert Shift:=xlDown
    
    Cells(j, 4) = Cells(j + 1, 4).Value
    
    Cells(j, 3) = Cells(j + 1, 3).Value 'coluna C

    Cells(j, 5) = Cells(i, 1).Value - 1

    Cells(j + 1, 4) = Cells(i, 2).Value + 1

End If
     j = j + 1
     Wend
        i = i + 1
        Wend
End Sub

Sub definir_colunas()
ThisWorkbook.Sheets(1).Name = "Remover"
ThisWorkbook.Sheets(1).Cells(1, 1) = "CEPI prioridade"
ThisWorkbook.Sheets(1).Cells(1, 2) = "CEPF prioridade"
ThisWorkbook.Sheets(1).Cells(1, 3) = "Método (opcional)"
ThisWorkbook.Sheets(1).Cells(1, 4) = "CEPI - Alteração"
ThisWorkbook.Sheets(1).Cells(1, 5) = "CEPF - Alteração"
ThisWorkbook.Sheets(1).Cells(1, 6) = "QTD_DIAS_UTEIS (opcional)"
ThisWorkbook.Sheets(1).Cells(1, 7) = "Preço (opcional)"
End Sub
