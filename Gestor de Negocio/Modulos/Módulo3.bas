Attribute VB_Name = "M�dulo3"
Option Private Module
Sub MostrarHojas()

    Dim Hoja As Worksheet
    
    For Each Hoja In Worksheets
        If Hoja.CodeName <> "Hoja20" Then
            Hoja.Visible = xlSheetVisible
      End If
    Next Hoja
    
End Sub
Sub OcultarHojas()

    Dim Hoja As Worksheet
    
    For Each Hoja In Worksheets
        If Hoja.CodeName <> "Hoja20" Then
            Hoja.Visible = xlSheetVeryHidden
      End If
    Next Hoja
    
End Sub
