Attribute VB_Name = "Reporte_Personal"
Sub Reporte_Planilla()
Attribute Reporte_Planilla.VB_ProcData.VB_Invoke_Func = " \n14"
    Cells.Select
    Selection.Copy
    Workbooks.Add
    Cells.Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-36
    Columns("E:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("F:F").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("H:H").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("K:M").Select
    Selection.EntireColumn.Hidden = True
    Columns("N:N").Select
    Selection.EntireColumn.Hidden = True
    Columns("R:R").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    Columns("R:V").Select
    Range("V1").Activate
    Selection.EntireColumn.Hidden = True
    Range("X4").Select
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select
    ActiveWindow.DisplayGridlines = False
End Sub
