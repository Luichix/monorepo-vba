Attribute VB_Name = "Módulo3"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveWindow.SmallScroll Down:=-12
    Range("A1:O36").Select
    Range("O36").Activate
    Selection.Copy
    ActiveWindow.SmallScroll Down:=24
    Range("A40").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=36
    Application.CutCopyMode = False
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Rows("42:42").Select
    Selection.Cut
    Rows("43:43").Select
    ActiveSheet.Paste
    Range("B47").Select
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    Range("A17:E17").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("D12").Select
End Sub
