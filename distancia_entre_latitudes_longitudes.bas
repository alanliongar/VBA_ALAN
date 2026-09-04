Attribute VB_Name = "Módulo5"
Public Function distancia(lat1 As Double, long1 As Double, lat2 As Double, long2 As Double) As Double

    Dim p As Double
    
    p = Application.WorksheetFunction.Pi() / 180
    
    distancia = 6378.137 * Application.WorksheetFunction.Acos( _
        Cos(p * lat1) * _
        Cos(p * lat2) * _
        Cos(p * long2 - p * long1) + _
        Sin(p * lat1) * _
        Sin(p * lat2) _
    )

End Function
