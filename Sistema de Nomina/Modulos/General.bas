Attribute VB_Name = "General"
Option Explicit
Sub Reporte_General()
Attribute Reporte_General.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Id As String
Dim Colaborador As String
Dim Fila As Long
Dim Final As Long
Dim TOTAL As String
Dim separador As String
Dim CK As String
Dim ACH As String
Dim xCol As Long


CK = Hoja81.Range("G2").Text
ACH = Hoja81.Range("G3").Text

separador = Application.International(xlListSeparator)

Id = "ID"
Colaborador = "COLABORADOR"
TOTAL = "TOTAL"

    Hoja16.Activate
    Hoja16.Cells.Select
    Selection.Clear
    
    Hoja4.Activate
    Hoja4.Cells.Select
    Application.CutCopyMode = False
    Hoja4.Cells.Copy
    
    Hoja16.Activate
    Hoja16.Cells(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Hoja16.Columns("C:BF").Select
    Selection.Delete
    
    Hoja16.Columns("S:AC").Select
    Selection.Delete
    
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Rows("1:2").Select
    Selection.RowHeight = 25
    
    Rows("3:500").Select
    Selection.RowHeight = 20
    
    Fila = Hoja16.Range("A" & Rows.Count).End(xlUp).Row
        Final = Fila + 1

    
    Hoja16.Activate
    Hoja16.Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection
         .VerticalAlignment = xlCenter
    End With
    
    
    Hoja16.Activate
    Hoja16.Cells(2, 1).Select
    Hoja16.Cells(2, 1) = Id
    
    Hoja16.Cells(2, 2).Select
    Hoja16.Cells(2, 2) = Colaborador
    
    Hoja16.Columns("A:A").Select
    Selection.ColumnWidth = 7
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Hoja16.Columns("C:D").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Hoja16.Range("A1:R1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = True
    End With


    Hoja16.Range("A2:R2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    
    Hoja16.Select
    
    
    Hoja16.Cells(Fila + 1, 1).Select
    Hoja16.Cells(Fila + 1, 1) = TOTAL
    

        
    Hoja16.Range(Cells(2, 1), Cells(Final, 18)).Select

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    
    
    Range("A2:R2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Range(Cells(2, 1), Cells(Final, 2)).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
    Range("A2:R2").Select
    Selection.Font.Bold = True
    Range("A1:R1").Select
    Selection.Font.Size = 10
    Selection.Font.Bold = True
    Range(Cells(Final, 1), Cells(Final, 18)).Select
    Selection.Font.Bold = True
    Range(Cells(2, 1), Cells(Final, 18)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A2:R2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
     Range(Cells(Final, 1), Cells(Final, 18)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    
    
'    Rows(Fila + 1).Select
'    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'
'     Rows(Fila + 3).Select
'     Selection.Copy
'     Rows(Fila + 1).Select
'    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'    SkipBlanks:=False, Transpose:=False
'    Application.CutCopyMode = False
'
'     Rows(Fila + 3).Select
'     Selection.Copy
'     Rows(Fila + 2).Select
'    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'    SkipBlanks:=False, Transpose:=False
'    Application.CutCopyMode = False
    
'
'
'   For xCol = 5 To 18
'
'   Hoja16.Cells(Fila + 1, xCol).Select
'
''   Hoja16.Cells(Fila + 1, xCol) = WorksheetFunction.SumIf(Range(Cells(3, 3), Cells(Fila, 3)), CK, Range(Cells(3, xCol), Cells(Fila, xCol)))
''   Hoja16.Cells(Fila + 2, xCol) = WorksheetFunction.SumIf(Range(Cells(3, 3), Cells(Fila, 3)), ACH, Range(Cells(3, xCol), Cells(Fila, xCol)))
'
'   Next xCol
    
'    Hoja16.Cells(Fila + 1, 4) = WorksheetFunction.CountIf(Range(Cells(3, 3), Cells(Fila, 3)), CK)
'    Hoja16.Cells(Fila + 2, 4) = WorksheetFunction.CountIf(Range(Cells(3, 3), Cells(Fila, 3)), ACH)

    Hoja16.Cells(Fila + 1, 4) = WorksheetFunction.CountA(Range(Cells(3, 3), Cells(Fila, 3)))
    
'    Hoja16.Cells(Fila + 1, 1) = "SUBTOTAL " & CK & ":"
'    Hoja16.Range(Cells(Fila + 1, 1), Cells(Fila + 1, 3)).Select
'        With Selection
'        .HorizontalAlignment = xlRight
'        .VerticalAlignment = xlCenter
'        .MergeCells = True
'        .InsertIndent 2
'        End With
'
'    Hoja16.Cells(Fila + 2, 1) = "SUBTOTAL " & ACH & ":"
'    Hoja16.Range(Cells(Fila + 2, 1), Cells(Fila + 2, 3)).Select
'        With Selection
'        .HorizontalAlignment = xlRight
'        .VerticalAlignment = xlCenter
'        .MergeCells = True
'        .InsertIndent 2
'        End With
   
 
Application.DisplayAlerts = False

    Hoja16.Cells(Fila + 1, 1) = "TOTAL PLANILLA " & CK & ":"
    Hoja16.Range(Cells(Fila + 1, 1), Cells(Fila + 1, 3)).Select
        With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .MergeCells = True
        .InsertIndent 2
        End With
Application.DisplayAlerts = True
'
'        Cheque
        CAJA
End Sub

Sub Cheque()
Dim X As String
Dim Y As String
Dim z As String
Dim a As String
Dim encontrado As Boolean

Hoja19.Activate
Hoja19.Cells.Select
    Selection.Clear

Hoja16.Activate
Hoja16.Cells.Select
Application.CutCopyMode = False
    Selection.Copy
    
Hoja19.Activate
Hoja19.Cells.Select
    ActiveSheet.Paste
    
Hoja19.Cells(1, 1).Select
With Selection
    .MergeCells = False
End With

Columns("E:Q").Select
    Selection.Delete

Hoja19.Range("A1:E1").Select
With Selection
    .MergeCells = True
End With

a = "SUBTOTAL ACH:"
X = Hoja81.Range("G2").Text 'CK
Y = "SUBTOTAL CK:"
z = "TOTAL PLANILLA CK & ACH:"

Hoja19.Activate
Hoja19.Select
Range("C2").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like X Then
            encontrado = True
              Selection.EntireRow.Delete
              ActiveCell.Offset(-1, 0).Select
        ElseIf ActiveCell.Value Like Y Then
            encontrado = True
              Selection.EntireRow.Delete
              ActiveCell.Offset(-1, 0).Select
        ElseIf ActiveCell.Value Like a Then
            encontrado = True
              ActiveCell.Value = "TOTAL PLANILLA ACH:"
              ActiveCell.Offset(-1, 0).Select
        ElseIf ActiveCell.Value Like z Then
            encontrado = True
              Selection.EntireRow.Delete
              ActiveCell.Offset(-1, 0).Select
        End If
    Loop

Hoja19.Cells(1, 1).Select
Hoja16.Activate
Hoja16.Cells(1, 1).Select
              
End Sub

Sub CAJA()
Dim X As String
Dim Y As String
Dim z As String
Dim a As String
Dim encontrado As Boolean
Dim Fila As Long
Dim Final As Long
Dim xCol As Long

Hoja22.Activate
Hoja22.Cells.Select
    Selection.Clear

Hoja16.Activate
Hoja16.Cells.Select
Application.CutCopyMode = False
    Selection.Copy

Hoja22.Activate
Hoja22.Cells.Select
    ActiveSheet.Paste

Hoja22.Cells(1, 1).Select
With Selection
    .MergeCells = False
End With

Columns("D:Q").Select
    Selection.Delete

Hoja22.Range("A1:L1").Select
With Selection
    .MergeCells = True
End With

'a = "SUBTOTAL ACH:"
'X = Hoja81.Range("G3").Text 'CK
'Y = "SUBTOTAL CK:"
'z = "TOTAL PLANILLA CK & ACH:"
'
'Hoja22.Activate
'Hoja22.Select
'Range("C2").Select
'
'    Do Until IsEmpty(ActiveCell)
'        ActiveCell.Offset(1, 0).Select
'        If ActiveCell.Value Like X Then
'            encontrado = True
'              Selection.EntireRow.Delete
'              ActiveCell.Offset(-1, 0).Select
'        ElseIf ActiveCell.Value Like a Then
'            encontrado = True
'              Selection.EntireRow.Delete
'              ActiveCell.Offset(-1, 0).Select
'        ElseIf ActiveCell.Value Like a Then
'            encontrado = True
'              ActiveCell.Value = "TOTAL PLANILLA CK:"
'              ActiveCell.Offset(-1, 0).Select
'        ElseIf ActiveCell.Value Like z Then
'            encontrado = True
'              Selection.EntireRow.Delete
'              ActiveCell.Offset(-1, 0).Select
'        End If
'    Loop

Hoja22.Cells(1, 1).Select
Hoja22.Range("E2") = "$20"
Hoja22.Range("F2") = "$10"
Hoja22.Range("G2") = "$1"
Hoja22.Range("H2") = "$0.5"
Hoja22.Range("I2") = "$0.25"
Hoja22.Range("J2") = "$0.1"
Hoja22.Range("K2") = "$0.05"
Hoja22.Range("L2") = "$0.01"




Final = GetUltimoR(Hoja22)

Final = Final - 1
    
For Fila = 3 To Final


Hoja22.Cells(Fila, 5) = Application.WorksheetFunction.RoundDown((Hoja22.Cells(Fila, 4)) / 20, 0)
Hoja22.Cells(Fila, 6) = Application.WorksheetFunction.RoundDown((Hoja22.Cells(Fila, 4) - (Hoja22.Cells(Fila, 5) * 20)) / 10, 0)
Hoja22.Cells(Fila, 7) = Application.WorksheetFunction.RoundDown((Hoja22.Cells(Fila, 4) - (Hoja22.Cells(Fila, 5) * 20 + Hoja22.Cells(Fila, 6) * 10)) / 1, 0)
Hoja22.Cells(Fila, 8) = Application.WorksheetFunction.RoundDown((Hoja22.Cells(Fila, 4) - (Hoja22.Cells(Fila, 5) * 20 + Hoja22.Cells(Fila, 6) * 10 + Hoja22.Cells(Fila, 7))) / 0.5, 0)
Hoja22.Cells(Fila, 9) = Application.WorksheetFunction.RoundDown((Hoja22.Cells(Fila, 4) - (Hoja22.Cells(Fila, 5) * 20 + Hoja22.Cells(Fila, 6) * 10 + Hoja22.Cells(Fila, 7) * 1 + Hoja22.Cells(Fila, 8) * 0.5)) / 0.25, 0)
Hoja22.Cells(Fila, 10) = Application.WorksheetFunction.RoundDown((Hoja22.Cells(Fila, 4) - (Hoja22.Cells(Fila, 5) * 20 + Hoja22.Cells(Fila, 6) * 10 + Hoja22.Cells(Fila, 7) * 1 + Hoja22.Cells(Fila, 8) * 0.5 + Hoja22.Cells(Fila, 9) * 0.25)) / 0.1, 0)
Hoja22.Cells(Fila, 11) = Application.WorksheetFunction.RoundDown((Hoja22.Cells(Fila, 4) - (Hoja22.Cells(Fila, 5) * 20 + Hoja22.Cells(Fila, 6) * 10 + Hoja22.Cells(Fila, 7) * 1 + Hoja22.Cells(Fila, 8) * 0.5 + Hoja22.Cells(Fila, 9) * 0.25 + Hoja22.Cells(Fila, 10) * 0.1)) / 0.05, 0)
Hoja22.Cells(Fila, 12) = Application.WorksheetFunction.Round((Hoja22.Cells(Fila, 4) - (Hoja22.Cells(Fila, 5) * 20 + Hoja22.Cells(Fila, 6) * 10 + Hoja22.Cells(Fila, 7) * 1 + Hoja22.Cells(Fila, 8) * 0.5 + Hoja22.Cells(Fila, 9) * 0.25 + Hoja22.Cells(Fila, 10) * 0.1 + Hoja22.Cells(Fila, 11) * 0.05)) / 0.01, 0)

Next

For xCol = 5 To 12
   
   Hoja22.Cells(Final + 1, xCol).Select
    
   Hoja22.Cells(Final + 1, xCol) = WorksheetFunction.Sum(Range(Cells(3, xCol), Cells(Final, xCol)))
  
   
Next xCol

    Range(Cells(2, 5), Cells(Final + 1, 12)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With


    Range("A2:L2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Font.Bold = True
    
    Range(Cells(Final + 1, 1), Cells(Final + 1, 12)).Select

    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Font.Bold = True
    
    
    For Fila = 3 To Final
    
    Range(Cells(Fila + 1, 1), Cells(Fila + 1, 12)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Range(Cells(Fila, 1), Cells(Fila, 12)).Select

    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Fila = Fila + 1
    
    Next
    
Hoja16.Activate
Hoja16.Cells(1, 1).Select
              
End Sub

