


Attribute VB_Name = "Módulo1"
Sub limites_MetAntigo()
Dim i As Long
Dim j As Long
Dim cepo As Long
Dim cepi As Long
Dim cepf As Long
Dim lim As Long
Dim uf As String
Dim time1 As Double

Sheets("tabela").Select
i = 2
While Cells(i, 2) <> ""
time1 = Timer
cepo = Cells(i, 2).Value
lim = Cells(i, 5).Value
uf = CStr(Cells(i, 9).Value)
Cells(i, 2).Copy
Sheets("ceps1").Select
Cells(2, 17).Select
ActiveSheet.Paste
Calculate
cepi = 99999999
cepf = 0
j = 2
While Cells(j, 2) <> ""
If Cells(j, 2).Value <> 16 Then
    If CLng(Cells(j, 13).Value) <= CLng(lim) And uf = CStr(Cells(j, 6).Value) Then
        If Cells(j, 1).Value < cepi And uf = CStr(Cells(j, 6).Value) Then
        cepi = Cells(j, 1).Value
        End If
        If Cells(j, 1).Value > cepf And uf = CStr(Cells(j, 6).Value) Then
        cepf = Cells(j, 1).Value
        End If
    End If
End If
j = j + 1
Wend
j = 2
Sheets("ceps2").Select
While Cells(j, 2) <> ""
If Cells(j, 2).Value <> 16 Then
    If CLng(Cells(j, 13).Value) <= CLng(lim) And uf = CStr(Cells(j, 6).Value) Then
        If Cells(j, 1).Value < cepi And uf = CStr(Cells(j, 6).Value) Then
        cepi = Cells(j, 1).Value
        End If
        If Cells(j, 1).Value > cepf And uf = CStr(Cells(j, 6).Value) Then
        cepf = Cells(j, 1).Value
        End If
    End If
End If
j = j + 1
Wend
Sheets("tabela").Select
Cells(i, 3) = cepi
Cells(i, 4) = cepf
If i Mod 10 = 0 Then
ActiveWorkbook.Save
End If
Cells(i, 12) = Timer - time1

i = i + 1
Wend
End Sub
 
Function uf(a As Long) As String
If a >= 69900000 And a <= 69999999 Then
uf = "AC"
End If

If a >= 57000000 And a <= 57999999 Then
uf = "AL"
End If

If a >= 69000000 And a <= 69299999 Then
uf = "AM"
End If

If a >= 69400000 And a <= 69899999 Then
uf = "AM"
End If

If a >= 68900000 And a <= 68999999 Then
uf = "AP"
End If

If a >= 40000000 And a <= 48999999 Then
uf = "BA"
End If

If a >= 60000000 And a <= 63999999 Then
uf = "CE"
End If

If a >= 70000000 And a <= 72799999 Then
uf = "DF"
End If

If a >= 73000000 And a <= 73699999 Then
uf = "DF"
End If

If a >= 29000000 And a <= 29999999 Then
uf = "ES"
End If

If a >= 72800000 And a <= 72999999 Then
uf = "GO"
End If

If a >= 73700000 And a <= 76799999 Then
uf = "GO"
End If

If a >= 65000000 And a <= 65999999 Then
uf = "MA"
End If

If a >= 30000000 And a <= 39999999 Then
uf = "MG"
End If

If a >= 79000000 And a <= 79999999 Then
uf = "MS"
End If

If a >= 78000000 And a <= 78899999 Then
uf = "MT"
End If

If a >= 66000000 And a <= 68899999 Then
uf = "PA"
End If

If a >= 58000000 And a <= 58999999 Then
uf = "PB"
End If

If a >= 50000000 And a <= 56999999 Then
uf = "PE"
End If

If a >= 64000000 And a <= 64999999 Then
uf = "PI"
End If

If a >= 80000000 And a <= 87999999 Then
uf = "PR"
End If

If a >= 20000000 And a <= 28999999 Then
uf = "RJ"
End If

If a >= 59000000 And a <= 59999999 Then
uf = "RN"
End If

If a >= 76800000 And a <= 76999999 Then
uf = "RO"
End If

If a >= 69300000 And a <= 69399999 Then
uf = "RR"
End If

If a >= 90000000 And a <= 99999999 Then
uf = "RS"
End If

If a >= 88000000 And a <= 89999999 Then
uf = "SC"
End If

If a >= 49000000 And a <= 49999999 Then
uf = "SE"
End If

If a >= 1000000 And a <= 19999999 Then
uf = "SP"
End If

If a >= 77000000 And a <= 77999999 Then
uf = "TO"
End If
End Function
Sub Definir_Areas_MET_Nova()
Application.ScreenUpdating = False
Application.Calculation = xlManual
Dim i As Long
Dim j As Long
Dim cepo As Long
Dim lim As Long
Dim uf As String
Dim java As Long
Dim srv As String
Dim sit As String
Dim trn As String
Dim time1 As Double
Dim k As Long
Dim ini As Long
Dim fin As Long

Sheets("tabela").Select
i = 2
While Cells(i, 2) <> ""
time1 = Timer
cepo = Cells(i, 2).Value
lim = Cells(i, 5).Value
uf = CStr(Cells(i, 9).Value)
java = CStr(Cells(i, 1).Value)
sit = CStr(Cells(i, 10).Value)
trn = CStr(Cells(i, 11).Value)
srv = CStr(Cells(i, 8).Value)

Cells(i, 2).Copy
Sheets("ceps1").Select
Cells(2, 17).Select
ActiveSheet.Paste
ThisWorkbook.Sheets("ceps1").Calculate
Sheets("ceps1").Select
    Columns("A:M").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A:$M").AutoFilter Field:=13, Criteria1:="<=" + CStr(lim), Operator:=xlAnd
    ActiveSheet.Range("$A:$M").AutoFilter Field:=6, Criteria1:="=" + CStr(uf), Operator:=xlAnd
    Cells.Select
    Selection.Copy
    Sheets("ceps1r").Select
    Range("A1").Select
    ActiveSheet.Paste

Sheets("ceps2").Select
ThisWorkbook.Sheets("ceps2").Calculate
    Columns("A:M").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A:$M").AutoFilter Field:=13, Criteria1:="<=" + CStr(lim), Operator:=xlAnd
    ActiveSheet.Range("$A:$M").AutoFilter Field:=6, Criteria1:="=" + CStr(uf), Operator:=xlAnd
    Cells.Select
    Selection.Copy
    Sheets("ceps2r").Select
    Range("A1").Select
    ActiveSheet.Paste

Sheets("ceps1r").Select
j = 2
While Cells(j, 1) <> ""
j = j + 1
Wend
j = j - 1
If j >= 2 Then
Range("A2:A" + CStr(j)).Copy
Sheets("faixas").Select
k = 2
While Cells(k, 3) <> ""
k = k + 1
Wend
Cells(k, 3).Select
ActiveSheet.Paste
Cells(k, 4).Select
ActiveSheet.Paste
Else
Sheets("faixas").Select
k = 2
While Cells(k, 3) <> ""
k = k + 1
Wend
End If
ini = k
'colocar os outros valores aqui
Cells(ini, 1) = java
Cells(ini, 2) = cepo
Cells(ini, 5) = lim
Cells(ini, 6) = srv
Cells(ini, 7) = uf
Cells(ini, 8) = sit
Cells(ini, 9) = trn
Sheets("ceps2r").Select
j = 2
While Cells(j, 1) <> ""
j = j + 1
Wend
j = j - 1
If j >= 2 Then
Range("A2:A" + CStr(j)).Copy
Sheets("faixas").Select
k = 2
While Cells(k, 3) <> ""
k = k + 1
Wend
Cells(k, 3).Select
ActiveSheet.Paste
Cells(k, 4).Select
ActiveSheet.Paste
Else
Sheets("faixas").Select
k = 2
While Cells(k, 3) <> ""
k = k + 1
Wend
End If


Sheets("faixas").Select
k = 2
While Cells(k, 3) <> ""
k = k + 1
Wend
k = k - 1
fin = k

Range("A" + CStr(ini) + ":B" + CStr(ini)).Copy
Range("A" + CStr(ini) + ":B" + CStr(fin)).Select
ActiveSheet.Paste
Range("E" + CStr(ini) + ":I" + CStr(ini)).Copy
Range("E" + CStr(ini) + ":I" + CStr(fin)).Select
ActiveSheet.Paste


Sheets("ceps1").Select
    Columns("A:M").Select
    Selection.AutoFilter
    
Sheets("ceps2").Select
    Columns("A:M").Select
    Selection.AutoFilter


Sheets("ceps1r").Select
Cells.Clear

Sheets("ceps2r").Select
Cells.Clear

Sheets("tabela").Select
If i Mod 5 = 0 Then
ActiveWorkbook.Save
End If



Cells(i, 12) = Timer - time1

i = i + 1
Wend
ActiveWorkbook.Save
Sheets("faixas").Select
Call Ajeitar_IP
ActiveWorkbook.Save
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
End Sub

Sub Ajeitar_IP()
Dim i As Long
i = 2
While Cells(i, 3) <> ""
i = i + 1
Wend
i = i - 1
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-7]<10000000,CONCATENATE(LEFT(RC[-7],4),""000""),CONCATENATE(LEFT(RC[-7],5),""000""))"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-7]<10000000,CONCATENATE(LEFT(RC[-7],4),""999""),CONCATENATE(LEFT(RC[-7],5),""999""))"
    Range("J2:K2").Select
    Selection.AutoFill Destination:=Range("J2:K" + CStr(i))
    ThisWorkbook.Sheets("faixas").Calculate
    Range("A1").Select
    
    Range("J1").Select

    Range("J1:K" + CStr(i)).Select
    Selection.Copy
    Range("J1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("J:J").Select
    Selection.TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("K:K").Select
    Selection.TextToColumns Destination:=Range("K1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("A1").Select
End Sub


Function TiraAcento(Palavra)
 CAcento = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
 SAcento = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
 Texto = ""
 If Palavra <> "" Then
 For X = 1 To Len(Palavra)
 Letra = Mid(Palavra, X, 1)
 Pos_Acento = InStr(CAcento, Letra)
 If Pos_Acento > 0 Then
 Letra = Mid(SAcento, Pos_Acento, 1)
 End If
 Texto = Texto & Letra
 Next
 TiraAcento = Texto
 End If
 End Function

Function VerificaPalavra(atributo)

Dim i
 Dim id
 Dim Auxiliar
 Dim Resultado

Auxiliar = Split(atributo, " ", -1, vbBinaryCompare)

For i = LBound(Auxiliar) To UBound(Auxiliar)
 Resultado = Resultado & " " & TiraAcento(Auxiliar(i))
 Next

VerificaPalavra = Trim(Resultado)
 End Function
'código criado pelo Alan
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


Function cidade(a As Long) As String
Dim limInf As Long
Dim limSup As Long
Dim sidade As String
On Error Resume Next
limInf = Application.WorksheetFunction.VLookup(a, ThisWorkbook.Sheets("Cidade").Columns("C:E"), 1, 1)
limSup = Application.WorksheetFunction.VLookup(a, ThisWorkbook.Sheets("Cidade").Columns("C:E"), 2, 1)
sidade = Application.WorksheetFunction.VLookup(a, ThisWorkbook.Sheets("Cidade").Columns("C:E"), 3, 1)
If a >= limInf And a <= limSup Then
cidade = sidade
Else
cidade = "Cidade nao definida"
End If
On Error GoTo 0
End Function



'Distancias - '=6378,137*ACOS(COS((PI()/180)*K2)*COS((PI()/180)*$O$2)*COS((PI()/180)*$P$2-(PI()/180)*L2)+SEN((PI()/180)*K2)*SEN((PI()/180)*$O$2))*1000
' O2: '=SEERRO(PROCV(Q2;A:L;11;0);PROCV(Q2;ceps2!A:L;11;0))
' P2: '=SEERRO(PROCV(Q2;A:L;12;0);PROCV(Q2;ceps2!A:L;12;0))
'ceps2: N2 '=ceps1!N2
'=ceps1!O2
'=ceps1!P2
'=ceps1!Q2
