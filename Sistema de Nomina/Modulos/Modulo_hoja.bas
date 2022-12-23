Attribute VB_Name = "Modulo_hoja"
Option Private Module
Sub MostrarHojas()

    Dim Hoja As Worksheet
    
    For Each Hoja In Worksheets
        If Hoja.CodeName <> "Hoja0" Then
            Hoja.Visible = xlSheetVisible
      End If
    Next Hoja
    
End Sub
Sub OcultarHojas()

    Dim Hoja As Worksheet
    
    For Each Hoja In Worksheets
        If Hoja.CodeName <> "Hoja0" Then
            Hoja.Visible = xlSheetVeryHidden
      End If
    Next Hoja
    
End Sub
