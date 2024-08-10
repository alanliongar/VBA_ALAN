Attribute VB_Name = "Temporizador"
Function tempuu(t As Variant) As String
Dim dias As Double
Dim hora As Double
Dim minutos As Double
Dim segundos As Double
Dim tempo As Double
tempo = CDbl(t)
dias = Application.RoundDown(tempo, 0)
hora = Application.RoundDown(tempo * 24, 0) Mod 24
minutos = Application.RoundDown(tempo * 24 * 60, 0) Mod 60
segundos = Application.RoundDown((tempo * 24 * 60 - Application.RoundDown(tempo * 24 * 60, 0)) * 60, 0) Mod 60
tempuu = "A macro demorou " + CStr(dias) + " dias, " + CStr(hora) + " horas, " + CStr(minutos) + " minutos e " + CStr(segundos) + " segundos para rodar."
End Function