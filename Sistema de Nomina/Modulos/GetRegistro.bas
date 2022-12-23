Attribute VB_Name = "GetRegistro"
Public Function GetUltimoR(Hoja As Worksheet) As Integer
    GetUltimoR = GetNuevoR(Hoja) - 1
End Function

Public Function GetNuevoR(Hoja As Worksheet) As Integer
    
    Dim Fila As Long
    Fila = 2
    
    Do While Hoja.Cells(Fila, 1) <> ""
        Fila = Fila + 1
    Loop
    
    GetNuevoR = Fila
    
End Function

