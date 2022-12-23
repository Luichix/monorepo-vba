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

    Cells.Select
    Selection.Copy
    Cells(1, 1).Select
    Workbooks.Add
    Cells.Select
    ActiveSheet.Paste
    ActiveWindow.DisplayGridlines = False
    Columns("B:I").Select
    Selection.EntireColumn.Hidden = True
    Columns("K:K").Select
    Selection.EntireColumn.Hidden = True
    Cells(1, 1).Select
End Sub

Sub OcultarDisplays()
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayFormulaBar = False
End Sub
Sub MostrarDisplays()
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayGridlines = True
End Sub
Sub xddx()

FilaActivo = Hoja45.Range("A" & Rows.Count).End(xlUp).Row
FilaPasivo = Hoja45.Range("E" & Rows.Count).End(xlUp).Row

If FilaActivo > FilaPasivo Then
    MsgBox FilaActivo
ElseIf FilaPasivo > FilaActivo Then
    MsgBox FilaPasivo
End If

End Sub
