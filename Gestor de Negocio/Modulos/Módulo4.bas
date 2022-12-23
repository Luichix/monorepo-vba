Attribute VB_Name = "Módulo4"
Sub REPORTE1()
Attribute REPORTE1.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False
    Hoja12.Select
    Range("Existencias[[#Headers],[DESCRIPCIÓN]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveWindow.SmallScroll Down:=3
    Selection.Copy
    Application.WindowState = xlNormal
    Workbooks.Add
    ActiveSheet.Paste
    Columns("B:L").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Range("E6").Select
    Columns("A:A").EntireColumn.AutoFit
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveWindow.DisplayGridlines = False
    Application.ScreenUpdating = True
    
End Sub
