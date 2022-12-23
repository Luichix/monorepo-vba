Attribute VB_Name = "Decimo_Recibo"
Option Explicit
Public Sub Recibo_DTM()
Dim Empresa As String
Dim Periodo As String
Dim Id As String
Dim Cedula As String
Dim Cuenta As String
Dim Detalle As String
Dim Ingreso As String
Dim Decimo As String
Dim Seguro As String
Dim ISR As String
Dim Adelanto As String
Dim Deduccion As String
Dim Neto As String
Dim Recibi As String

Dim Fecha As Date

Dim xID As String
Dim xColaborador As String
Dim xCedula As String
Dim xCuenta As String
Dim xIngresos As Currency
Dim xDecimo As Currency
Dim xSeguro As Currency
Dim xISR As Currency
Dim xDeduccion As Currency
Dim xAdelanto As Currency
Dim xNeto As Currency


Dim xFila As Long
Dim Fila As Long
Dim xFinal As Long
Dim encontrado As Boolean
Dim Referencia As String
Dim zContar As Long
Dim xContar As Long

Application.ScreenUpdating = False

Fecha = Hoja23.Cells(2, 7)

zContar = 0
xContar = 1

Empresa = "COMPROBANTE DE DECIMO"
Periodo = UCase("DECIMO TERCER MES - " & Format(Fecha, "MMMM YYYY"))
Id = "ID:"
Cedula = "CEDULA:"
Cuenta = "CUENTA:"
Detalle = "DETALLE DE DECIMO TERCER MES"
Ingreso = "INGRESOS ACUMULADOS:"
Decimo = "DECIMO TERCER MES:"
Seguro = "SEGURO SOCIAL:"
ISR = "ISR:"
Adelanto = "ADELANTO:"
Deduccion = "OTRAS DEDUCCIONES:"
Neto = "DECIMO NETO:"
Recibi = "RECIBI CONFORME"


Hoja25.Activate
Hoja25.Select
Hoja25.Cells.Select

Selection.Clear

    'FORMATO LETRA CALIBRI 9
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


Hoja23.Activate
Hoja23.Cells(5, 1).Select

xFila = 6
Do While Hoja23.Cells(xFila, 1) <> Empty
xFila = xFila + 1
Loop
xFinal = xFila - 1
Fila = 1

For xFila = 6 To xFinal

Hoja25.Activate
Hoja25.Cells(1, 1).Select

Hoja25.Cells(Fila, 1) = Empresa
Hoja25.Cells(Fila + 1, 1) = Periodo
Hoja25.Cells(Fila + 2, 1) = Id
Hoja25.Cells(Fila + 3, 1) = Cedula
Hoja25.Cells(Fila + 3, 4) = Cuenta
Hoja25.Cells(Fila + 4, 1) = Detalle
Hoja25.Cells(Fila + 5, 1) = Ingreso
Hoja25.Cells(Fila + 6, 1) = Decimo
Hoja25.Cells(Fila + 7, 1) = Seguro
Hoja25.Cells(Fila + 8, 1) = ISR
Hoja25.Cells(Fila + 9, 1) = Adelanto
Hoja25.Cells(Fila + 10, 1) = Deduccion
Hoja25.Cells(Fila + 11, 1) = Neto
Hoja25.Cells(Fila + 14, 2) = Recibi

Hoja25.Range(Cells(Fila, 1), Cells(Fila, 5)).Select
Diseño_A
Formato_B

Hoja25.Range(Cells(Fila + 1, 1), Cells(Fila + 1, 5)).Select
Diseño_A
Formato_A

Hoja25.Range(Cells(Fila + 4, 1), Cells(Fila + 4, 5)).Select
Diseño_A
Formato_A

Hoja25.Range(Cells(Fila + 11, 1), Cells(Fila + 11, 5)).Select
Formato_A

Hoja25.Range(Cells(Fila + 14, 1), Cells(Fila + 14, 5)).Select
Diseño_A

Hoja25.Range(Cells(Fila + 13, 2), Cells(Fila + 13, 4)).Select
Diseño_A
Borde_B

Hoja25.Range(Cells(Fila, 1), Cells(Fila + 14, 5)).Select
Borde_A
Borde_B
Borde_I
Borde_D



Referencia = Hoja23.Cells(xFila, 1)

Hoja23.Select
Range("A5").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like Referencia Then
            encontrado = True
            xID = ActiveCell.Offset(0, 0).Value
            xColaborador = ActiveCell.Offset(0, 1).Value
            xCedula = ActiveCell.Offset(0, 4).Value
            xCuenta = ActiveCell.Offset(0, 2).Value
            xIngresos = ActiveCell.Offset(0, 15).Value
            xDecimo = ActiveCell.Offset(0, 16).Value
            xSeguro = ActiveCell.Offset(0, 17).Value
            xISR = ActiveCell.Offset(0, 18).Value
            xAdelanto = ActiveCell.Offset(0, 19).Value
            xDeduccion = ActiveCell.Offset(0, 20).Value
            xNeto = ActiveCell.Offset(0, 21).Value
            Exit Do
        End If
    Loop


    Hoja25.Activate

    Hoja25.Cells(Fila + 2, 2) = xID & " - " & xColaborador
    Hoja25.Cells(Fila + 3, 2) = xCedula
    Hoja25.Cells(Fila + 3, 5) = xCuenta
    Hoja25.Cells(Fila + 5, 4) = xIngresos
    Hoja25.Cells(Fila + 6, 4) = xDecimo
    Hoja25.Cells(Fila + 7, 4) = xSeguro
    Hoja25.Cells(Fila + 8, 4) = xISR
    Hoja25.Cells(Fila + 9, 4) = xAdelanto
    Hoja25.Cells(Fila + 10, 4) = xDeduccion
    Hoja25.Cells(Fila + 11, 4) = xNeto
    
        Hoja25.Cells(Fila + 5, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Hoja25.Cells(Fila + 6, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Hoja25.Cells(Fila + 7, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Hoja25.Cells(Fila + 8, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Hoja25.Cells(Fila + 9, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Hoja25.Cells(Fila + 10, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Hoja25.Cells(Fila + 11, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"


zContar = zContar + 1
If zContar = 1 Then
    xContar = 1
    zContar = 1
ElseIf zContar = 2 Then
    xContar = 1
    zContar = 2
ElseIf zContar = 3 Then
    xContar = 0
    zContar = 0
    
End If
Fila = Fila + 15 + xContar

Next

    Hoja25.Activate
    Columns("A:E").Select
    Range("A3").Activate
    Selection.Copy
    Columns("H:H").Select
    Range("H3").Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False


Hoja25.Select
Hoja25.Cells(1, 1).Select
Application.ScreenUpdating = True



End Sub
Public Sub Diseño_A()
    With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .MergeCells = True
    End With
End Sub
Public Sub Diseño_B()
    With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
End Sub
Public Sub Diseño_C()
    With Selection
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlCenter
    .MergeCells = True
    End With
End Sub
Public Sub Diseño_D()
    With Selection
    .VerticalAlignment = xlCenter
    .MergeCells = True
    End With
End Sub
Public Sub Formato_A()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
End Sub
Public Sub Formato_B()
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = -0.249977111117893
    .PatternTintAndShade = 0
End With
End Sub
Public Sub Borde_A()
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
Public Sub Borde_B()
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
Public Sub Borde_I()
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
Public Sub Borde_D()
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub

Public Sub Borde_H()
With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
  

