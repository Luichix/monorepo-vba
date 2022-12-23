Attribute VB_Name = "Módulo3"
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

Public Sub Reportes_Inventario()

' Reportes_Inventario
Application.ScreenUpdating = False
    Application.EnableEvents = False

    Cells.Select
    Selection.Copy
    Cells(1, 1).Select
    Hoja14.Select
    Cells.Select
    ActiveSheet.Paste
    ActiveWindow.DisplayGridlines = False
    Columns("B:I").Select
    Selection.EntireColumn.Hidden = True
    Columns("K:K").Select
    Selection.EntireColumn.Hidden = True
    Cells(1, 1).Select
    Columns("A:A").ColumnWidth = 16
    Columns("J:J").ColumnWidth = 8
    Columns("A:A").Select
    Selection.Font.Size = 8
    
    Hoja14.Select
    Hoja14.Cells(1, 1).Select
    
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False

    Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub


