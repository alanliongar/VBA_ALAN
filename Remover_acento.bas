Attribute VB_Name = "Módulo1"
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
