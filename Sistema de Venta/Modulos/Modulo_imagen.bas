Attribute VB_Name = "Modulo_imagen"
Option Explicit

Function nAnimales(nombre As String) As Integer
    Application.ScreenUpdating = False

    Hoja29.Select
    Hoja29.Cells(2, 5).Select

    nAnimales = 0

    Do While Not IsEmpty(ActiveCell)
        If nombre = ActiveCell Then
            nAnimales = ActiveCell.Row
        End If
        ActiveCell.Offset(1, 0).Select
    Loop

    Application.ScreenUpdating = True

End Function



