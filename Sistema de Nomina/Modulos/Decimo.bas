Attribute VB_Name = "Decimo"
Option Explicit
Sub Reporte_Decimo()
Dim Id As String
Dim Colaborador As String
Dim Fila As Long
Dim Final As Long
Dim TOTAL As String
Dim separador As String
Dim CK As String
Dim ACH As String
Dim xCol As Long
Dim Fecha As Date
Dim zFila As Long
Dim zFinal As Long

Dim QuincenaI As String
Dim MesII As String
Dim MesIII As String
Dim MesIV As String
Dim QuincenaII As String

QuincenaII = UCase("II " & Format(Hoja23.Cells(2, 11), "mmmm"))
MesII = UCase(Format(Hoja23.Cells(3, 12), "mmmm"))
MesIII = UCase(Format(Hoja23.Cells(3, 13), "mmmm"))
MesIV = UCase(Format(Hoja23.Cells(3, 14), "mmmm"))
QuincenaI = UCase("I " & Format(Hoja23.Cells(2, 7), "mmmm"))



Dim X As String
Dim Y As String
Dim z As String
Dim a As String
Dim W As String
Dim encontrado As Boolean

'a = "SUBTOTAL ACH:"
'X = Hoja81.Range("G2").Text 'CK
'W = Hoja81.Range("G3").Text 'SP
'Y = "SUBTOTAL CK:"
'z = "TOTAL PLANILLA CK & ACH:"

Fecha = Hoja23.Cells(2, 7)

CK = Hoja81.Range("G2").Text
ACH = Hoja81.Range("G3").Text

separador = Application.International(xlListSeparator)

Id = "ID"
Colaborador = "COLABORADOR"
TOTAL = "TOTAL"

    Hoja24.Activate
    Hoja24.Cells.Select
    Selection.Clear
    
    Hoja23.Activate
    Hoja23.Cells.Select
    Application.CutCopyMode = False
    Hoja23.Cells.Copy
    
    Hoja24.Activate
    Hoja24.Cells(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Hoja24.Columns("E:J").Select
    Selection.Delete

    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("1:2").Select
    Selection.RowHeight = 25
    
    Rows("3:500").Select
    Selection.RowHeight = 20
    
    Fila = Hoja24.Range("A" & Rows.Count).End(xlUp).Row
        Final = Fila + 1
        
    Hoja24.Cells(1, 1) = "PLANILLA DE PAGO DEL DECIMO TERCER MES DE " & UCase(Format(Fecha, "mmmm yyyy"))
    Hoja24.Cells(2, 5) = QuincenaII
    Hoja24.Cells(2, 6) = MesII
    Hoja24.Cells(2, 7) = MesIII
    Hoja24.Cells(2, 8) = MesIV
    Hoja24.Cells(2, 9) = QuincenaI
    
            
    Hoja24.Activate
    Hoja24.Cells.Select
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

    Hoja24.Activate
    
    Hoja24.Columns("A:A").Select
    Selection.ColumnWidth = 7
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Hoja24.Columns("C:D").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Hoja24.Range("A1:P1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = True
    End With


    Hoja24.Range("A2:P2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    
    Hoja24.Select
    
    
    Hoja24.Cells(Fila + 1, 1).Select
    Hoja24.Cells(Fila + 1, 1) = TOTAL
    

        
    Hoja24.Range(Cells(2, 1), Cells(Final, 16)).Select

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
    
    
    
    Range("A2:P2").Select
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
    
    Range("A2:P2").Select
    Selection.Font.Bold = True
    Range("A1:P1").Select
    Selection.Font.Size = 10
    Selection.Font.Bold = True
    Range(Cells(Final, 1), Cells(Final, 16)).Select
    Selection.Font.Bold = True
    Range(Cells(2, 1), Cells(Final, 16)).Select
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
    Range("A2:P2").Select
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
     Range(Cells(Final, 1), Cells(Final, 16)).Select
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

'   For xCol = 5 To 16
'
'   Hoja24.Cells(Fila + 1, xCol).Select
'
'   Hoja24.Cells(Fila + 1, xCol) = WorksheetFunction.SumIf(Range(Cells(3, 3), Cells(Fila, 3)), CK, Range(Cells(3, xCol), Cells(Fila, xCol)))
'   Hoja24.Cells(Fila + 2, xCol) = WorksheetFunction.SumIf(Range(Cells(3, 3), Cells(Fila, 3)), ACH, Range(Cells(3, xCol), Cells(Fila, xCol)))
'
'   Next xCol

'    Hoja24.Cells(Fila + 1, 4) = WorksheetFunction.CountIf(Range(Cells(3, 3), Cells(Fila, 3)), CK)
'    Hoja24.Cells(Fila + 2, 4) = WorksheetFunction.CountIf(Range(Cells(3, 3), Cells(Fila, 3)), ACH)
'    Hoja24.Cells(Fila + 3, 4) = WorksheetFunction.CountA(Range(Cells(3, 3), Cells(Fila, 3)))
'
'    Hoja24.Cells(Fila + 1, 1) = "SUBTOTAL " & CK & ":"
'    Hoja24.Range(Cells(Fila + 1, 1), Cells(Fila + 1, 3)).Select
'        With Selection
'        .HorizontalAlignment = xlRight
'        .VerticalAlignment = xlCenter
'        .MergeCells = True
'        .InsertIndent 2
'        End With
'
'    Hoja24.Cells(Fila + 2, 1) = "SUBTOTAL " & ACH & ":"
'    Hoja24.Range(Cells(Fila + 2, 1), Cells(Fila + 2, 3)).Select
'        With Selection
'        .HorizontalAlignment = xlRight
'        .VerticalAlignment = xlCenter
'        .MergeCells = True
'        .InsertIndent 2
'        End With


Application.DisplayAlerts = False

    Hoja24.Cells(Fila + 1, 1) = "TOTAL PLANILLA SPE:"
    Hoja24.Range(Cells(Fila + 1, 1), Cells(Fila + 1, 3)).Select
        With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .MergeCells = True
        .InsertIndent 2
        End With
Application.DisplayAlerts = True



'    Hoja24.Activate
'    Hoja24.Range(Cells(1, 1), Cells(Fila + 3, 16)).Select
'    Application.CutCopyMode = False
'    Selection.Copy
'
'    zFila = Fila + 5
'
'    Hoja24.Activate
'    Hoja24.Cells(zFila, 1).Select
'    ActiveSheet.Paste
'Application.CutCopyMode = False
'
'Hoja24.Select
'Hoja24.Cells(zFila + 1, 3).Select
'
'    Do Until IsEmpty(ActiveCell)
'        ActiveCell.Offset(1, 0).Select
'        If ActiveCell.Value Like W Then
'            encontrado = True
'              Selection.EntireRow.Delete
'              ActiveCell.Offset(-1, 0).Select
'        ElseIf ActiveCell.Value Like a Then
'            encontrado = True
'              Selection.EntireRow.Delete
'              ActiveCell.Offset(-1, 0).Select
'        ElseIf ActiveCell.Value Like Y Then
'            encontrado = True
'              ActiveCell.Value = "TOTAL DECIMO CK:"
'              ActiveCell.Offset(-1, 0).Select
'
'            ElseIf ActiveCell.Value Like z Then
'            encontrado = True
'              Selection.EntireRow.Cut
'              ActiveCell.Offset(1, 0).Select
'              ActiveSheet.Paste
'
'
'        End If
'    Loop
'
'
Hoja22.Activate
Hoja22.Cells.Select
    Selection.Clear

        zFinal = Hoja24.Range("A" & Rows.Count).End(xlUp).Row
        zFinal = zFinal

Hoja24.Activate
Hoja24.Select
Hoja24.Range(Cells(1, 1), Cells(zFinal, 16)).Select
Application.CutCopyMode = False
    Selection.Copy

Hoja22.Activate
Hoja22.Cells.Select
    ActiveSheet.Paste
'
'
'Hoja24.Select
'Range("C2").Select
'
'    Do Until IsEmpty(ActiveCell)
'        ActiveCell.Offset(1, 0).Select
'        If ActiveCell.Value Like X Then
'            encontrado = True
'              Selection.EntireRow.Delete
'              ActiveCell.Offset(-1, 0).Select
'        ElseIf ActiveCell.Value Like Y Then
'            encontrado = True
'              Selection.EntireRow.Delete
'              ActiveCell.Offset(-1, 0).Select
'        ElseIf ActiveCell.Value Like a Then
'            encontrado = True
'              ActiveCell.Value = "TOTAL DECIMO ACH:"
'              ActiveCell.Offset(-1, 0).Select
'        ElseIf ActiveCell.Value Like z Then
'            encontrado = True
'              Selection.EntireRow.Delete
'              ActiveCell.Offset(-1, 0).Select
'        End If
'    Loop
    
    Logaritmo
    

End Sub



Sub Logaritmo()
Dim X As String
Dim Y As String
Dim z As String
Dim a As String
Dim encontrado As Boolean
Dim Fila As Long
Dim Final As Long
Dim xCol As Long




Hoja22.Activate
Hoja22.Select
Hoja22.Cells(1, 1).Select
With Selection
    .MergeCells = False
End With

Columns("D:N").Select
    Selection.Delete

Hoja22.Range("A1:L1").Select
With Selection
    .MergeCells = True
End With

Hoja22.Cells(1, 1).Select
Hoja22.Range("E2") = "VEINTE"
Hoja22.Range("F2") = "DECENA"
Hoja22.Range("G2") = "UNIDAD"
Hoja22.Range("H2") = "MEDIO"
Hoja22.Range("I2") = "CUADRA"
Hoja22.Range("J2") = "DECIMO"
Hoja22.Range("K2") = "CINCO"
Hoja22.Range("L2") = "CENTAVO"


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
    
Hoja24.Activate
Hoja24.Cells(1, 1).Select
              
End Sub
Public Sub Grabar_Decimo()
Dim Fila As Long
Dim Final As Long
Dim xFila As Long
Dim Fecha As Date
Dim Colilla As Long
Dim Registro As Long
Dim Seguridad As String
Dim encontrado As Boolean
Dim Repetido As String
Dim X As Long
Dim Periodo As Long
Dim Dia As Date
Dim Mes As Date
Dim Ano As Date
Dim Titulo As String

Titulo = "Gestor de Recursos Humanos"
Seguridad = Hoja83.Range("L1").Text

X = 0

    Hoja7.Unprotect (Seguridad)
    Hoja23.Unprotect (Seguridad)

Hoja23.Activate
ActiveSheet.ListObjects("Tbl_Decimo").ShowTotals = False


Fila = 6

Do While Hoja23.Cells(Fila, 1) <> Empty
   Fila = Fila + 1
Loop
   Final = Fila - 1

Periodo = Hoja23.Cells(2, 7)
Fecha = Hoja23.Cells(2, 7)

For Fila = 6 To Final

Repetido = Periodo & "-" & Hoja23.Cells(Fila, 1) & "-" & "DTMS"

Hoja7.Select
Hoja7.Range("B1").Select
Do Until IsEmpty(ActiveCell)
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.Value Like Repetido Then
        encontrado = True
        Exit Do
    End If
Loop

If encontrado = True Then
    X = X + 1

Else

    Hoja7.Select

    Hoja7.Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

    Hoja7.Cells(2, 1) = Fecha
    Hoja7.Cells(2, 2) = Repetido
    Hoja7.Cells(2, 3) = Hoja23.Cells(Fila, 1)
    Hoja7.Cells(2, 8) = Hoja23.Cells(Fila, 17)
       

Mes = VBA.Month(Fecha)
Ano = VBA.Year(Fecha)

    Hoja7.Cells(2, 9) = DateSerial(Ano, Mes, 1)
    Hoja7.Cells(2, 10) = Hoja83.Range("G1")
    
End If

    
Next

Hoja23.Activate
ActiveSheet.ListObjects("Tbl_Decimo").ShowTotals = True

Hoja7.Select
Hoja7.Cells(1, 1).Select

    MsgBox "Se han encontrado " & X & " registros ya existentes en la hoja PAGOS"
    MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
    
    Hoja7.Protect (Seguridad)
    Hoja23.Protect (Seguridad)
    
End Sub



