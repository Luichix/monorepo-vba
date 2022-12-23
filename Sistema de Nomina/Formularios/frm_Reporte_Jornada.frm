VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Reporte_Jornada 
   Caption         =   "REPORTE DE JORNADA"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   8670.001
   OleObjectBlob   =   "frm_Reporte_Jornada.frx":0000
End
Attribute VB_Name = "frm_Reporte_Jornada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Reporte_Jornada()


Application.ScreenUpdating = False

Dim xFila As Long
Dim zFila As Long
Dim xConteo As Long
Dim zConteo As Long

'VARIABLES DE TEXTO
Dim nReporte As String
Dim nPeriodo As String
Dim nId As String
Dim nColaborador As String
Dim nIngreso As String
Dim nRegimen As String
Dim nJornada As String

Dim nFecha As String
Dim nLaborar As String
Dim nLaborado As String
Dim nFavor As String
Dim nPendiente As String
Dim nDiurna As String
Dim nVespertina5 As String
Dim nNocturna6 As String
Dim nNocturna8 As String
Dim nTotal As String

Dim nHPendiente As String
Dim nHLaborar As String
Dim nHFaltante As String
Dim nHReponer As String
Dim nTPendiente As String

Dim nHFavor As String
Dim nHLaborado As String
Dim nHSobrante As String
Dim nHExtra As String
Dim nTFavor As String

Dim nHoraPagar As String
Dim nSFavor As String
Dim nSPendiente As String
Dim nVPendiente As String

'VARIABLES DE VALORES
Dim xID As String
Dim xColaborador As String
Dim xIngreso As Date
Dim xRegimen As String
Dim xJornada As String

Dim xFecha As Date
Dim xLaborar As Date
Dim xLaborado As Date
Dim xFavor As Date
Dim xPendiente As Date
Dim xDiurnas As Date
Dim xVespertina5 As Date
Dim xNocturna6 As Date
Dim xNocturna8 As Date

Dim Referencia As String
Dim encontrado As Boolean
Dim Hora As Long
Dim Conteo As Long
Dim Dia As Long
Dim Recuento As Long
Dim Vacio As String
Dim Columna As Long

Dim zLaborar As Date
Dim zPendiente As Date
Dim zReponer As Date

Dim zLaboradas As Date
Dim zFavor As Date
Dim zExtra As Date

Dim zPagar As Date

Dim zSaldoFavor As Date
Dim zSaldoPendiente As Date
    
Dim zValorP As Currency


'VALOR DE VARIABLES DE TEXTO
nReporte = "REPORTE INDIVIDUAL DE JORNADA LABORAL"
nPeriodo = UCase(Hoja81.Cells(9, 26).Text)
nId = "ID:"
nColaborador = "COLABORADOR:"
nIngreso = "FECHA DE INGRESO:"
nRegimen = "RÉGIMEN:"
nJornada = "JORNADA:"

nFecha = "FECHA"
nLaborar = "HORAS A LABORAR"
nLaborado = "HORAS LABORADAS"
nFavor = "TIEMPO A FAVOR"
nPendiente = "TIEMPO PENDIENTE"
nDiurna = "EXTRAS DIURNAS"
nVespertina5 = "EXTRAS VESPERTINAS 5-6"
nNocturna6 = "EXTRAS NOCTURNAS 6-8"
nNocturna8 = "EXTRAS NOCTURNAS 8+"
nTotal = "TOTAL"

nHPendiente = "HORAS PENDIENTES"
nHLaborar = "HORAS A LABORAR:"
nHFaltante = "TIEMPO PENDIENTES:"
'nHReponer = "TIEMPO A REPONER:"
nTPendiente = "TOTAL PENDIENTE:"

nHFavor = "HORAS A FAVOR"
nHLaborado = "HORAS LABORADAS:"
nHSobrante = "TIEMPO A FAVOR:"
nHExtra = "TIEMPO EXTRA:"
nTFavor = "TOTAL A FAVOR:"

'nHoraPagar = "HORAS POR PAGAR"
nSFavor = "SALDO A FAVOR"
nSPendiente = "SALDO AUSENCIA"
nVPendiente = "VALOR AUSENCIA"


            

'PROCEDIMIENTO

Hoja21.Activate
Hoja21.Cells.Select
Selection.Clear

    
    Hoja21.Cells.Select
    
    'FORMATO LETRA CALIBRI 10
    With Selection.Font
        .Name = "Calibri"
        .Size = 10
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
    
    'INGRESO DE CAMPOS
    
Hoja21.Cells(1, 1) = nReporte
Hoja21.Cells(2, 1) = nPeriodo
Hoja21.Cells(3, 1) = nId
Hoja21.Cells(4, 1) = nColaborador
Hoja21.Cells(5, 1) = nIngreso
Hoja21.Cells(3, 6) = nRegimen
Hoja21.Cells(4, 6) = nJornada

Hoja21.Cells(6, 1) = nFecha
Hoja21.Cells(6, 2) = nLaborar
Hoja21.Cells(6, 3) = nLaborado
Hoja21.Cells(6, 4) = nFavor
Hoja21.Cells(6, 5) = nPendiente
Hoja21.Cells(6, 6) = nDiurna
Hoja21.Cells(6, 7) = nVespertina5
Hoja21.Cells(6, 8) = nNocturna6
Hoja21.Cells(6, 9) = nNocturna8
Hoja21.Cells(23, 1) = nTotal

Hoja21.Cells(25, 2) = nHLaborar
Hoja21.Cells(25, 5) = nHLaborado

Hoja21.Cells(27, 1) = nHPendiente
Hoja21.Cells(28, 2) = nHFaltante
'Hoja21.Cells(29, 2) = nHReponer
Hoja21.Cells(30, 2) = nTPendiente

Hoja21.Cells(27, 4) = nHFavor
Hoja21.Cells(28, 5) = nHSobrante
Hoja21.Cells(29, 5) = nHExtra
Hoja21.Cells(30, 5) = nTFavor

'Hoja21.Cells(25, 8) = nHoraPagar
Hoja21.Cells(27, 8) = nSFavor
Hoja21.Cells(27, 9) = nSPendiente
Hoja21.Cells(30, 8) = nVPendiente

    'FORMATO DE CAMPOS
    
Hoja21.Range(Cells(1, 1), Cells(1, 9)).Select
Combinado_Centrado
Fondo_Azul
Letra_Blanca
Fuente_Negrita

Hoja21.Range(Cells(2, 1), Cells(2, 9)).Select
Combinado_Centrado
Fondo_Celeste_Claro

xFila = 3

For xConteo = 1 To 10

Hoja21.Range(Cells(xFila, 1), Cells(xFila, 9)).Select
Fondo_Celeste_Claro

xFila = xFila + 2

Next

zFila = 4

For zConteo = 1 To 10

Hoja21.Range(Cells(zFila, 1), Cells(zFila, 9)).Select
Fondo_Celeste_Intenso

zFila = zFila + 2

Next

Hoja21.Range(Cells(3, 1), Cells(5, 2)).Select
Fondo_Azul
Letra_Blanca
Fuente_Negrita
Centrado_Superior

Hoja21.Range(Cells(3, 6), Cells(5, 6)).Select
Fondo_Azul
Letra_Blanca
Fuente_Negrita
Centrado_Superior

Hoja21.Range(Cells(6, 1), Cells(6, 9)).Select
Fondo_Azul
Letra_Blanca
Ajustar_Centrar
Fuente_Negrita


Hoja21.Range(Cells(23, 1), Cells(23, 9)).Select
Fuente_Negrita
Fondo_Azul


Hoja21.Range(Cells(7, 1), Cells(23, 9)).Select
Formato_Centrado
    
Hoja21.Range(Cells(25, 1), Cells(25, 6)).Select
Fondo_Celeste_Intenso

Hoja21.Range(Cells(25, 1), Cells(25, 2)).Select
Fondo_Azul
Fuente_Negrita
Centrado_Superior

Hoja21.Range(Cells(25, 4), Cells(25, 5)).Select
Fondo_Azul
Fuente_Negrita
Centrado_Superior

Hoja21.Range(Cells(27, 1), Cells(27, 3)).Select
Fondo_Azul
Fuente_Negrita
Combinado_Centrado

Hoja21.Range(Cells(27, 4), Cells(27, 6)).Select
Fondo_Azul
Fuente_Negrita
Combinado_Centrado

Hoja21.Range(Cells(28, 1), Cells(28, 6)).Select
Fondo_Celeste_Claro
Centrado_Superior

Hoja21.Range(Cells(29, 1), Cells(29, 6)).Select
Fondo_Celeste_Intenso
Centrado_Superior

Hoja21.Range(Cells(30, 1), Cells(30, 6)).Select
Fondo_Azul
Centrado_Superior
Fuente_Negrita

'Hoja21.Cells(25, 8).Select
'Fondo_Azul
'Formato_Centrado
'Fuente_Negrita
'
'Hoja21.Cells(25, 9).Select
'Fondo_Celeste_Intenso
'Formato_Centrado

Hoja21.Range(Cells(27, 8), Cells(27, 9)).Select
Fondo_Azul
Formato_Centrado
Fuente_Negrita

Hoja21.Range(Cells(28, 8), Cells(28, 9)).Select
Fondo_Celeste_Intenso
Formato_Centrado

Hoja21.Cells(30, 8).Select
Fondo_Azul
Formato_Centrado
Fuente_Negrita

Hoja21.Cells(30, 9).Select
Fondo_Celeste_Intenso
Formato_Centrado

Hoja21.Range(Cells(3, 3), Cells(6, 3)).Select
Formato_Izquierdo

Hoja21.Cells(25, 2).Select
Formato_Derecho

Hoja21.Cells(25, 5).Select
Formato_Derecho

Hoja21.Range(Cells(28, 2), Cells(30, 2)).Select
Formato_Derecho

Hoja21.Range(Cells(28, 5), Cells(30, 5)).Select
Formato_Derecho

Hoja21.Range(Cells(6, 1), Cells(23, 9)).Select
Borde_Blanco

Hoja21.Range(Cells(1, 1), Cells(2, 9)).Select
Borde_Grueso

Hoja21.Range(Cells(1, 1), Cells(23, 9)).Select
Borde_Grueso

Hoja21.Range(Cells(25, 1), Cells(25, 6)).Select
Borde_Grueso

Hoja21.Range(Cells(27, 1), Cells(30, 6)).Select
Borde_Grueso

'Hoja21.Range(Cells(25, 8), Cells(25, 9)).Select
'Borde_Grueso

Hoja21.Range(Cells(27, 8), Cells(28, 9)).Select
Borde_Grueso

Hoja21.Range(Cells(30, 8), Cells(30, 9)).Select
Borde_Grueso

'VALOR DE VARIABLES DE DATOS

Recuento = 0

            For Dia = 1 To 16
            
            If Hoja3.Cells(2, 9 + Recuento) = "-" Then
            
            
            Vacio = Hoja3.Cells(2, 9 + Recuento).Value
            Hoja21.Cells(6 + Dia, 1) = Vacio
            
            Else
            
            xFecha = Hoja3.Cells(2, 9 + Recuento).Value
            Hoja21.Cells(6 + Dia, 1) = xFecha
            
            
            End If
            
            Recuento = Recuento + 12
            
            Next


    

Referencia = Me.txt_id.Text

Hoja3.Select
Range("A4").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like Referencia Then
            encontrado = True
            xID = ActiveCell.Offset(0, 0).Value
            xColaborador = ActiveCell.Offset(0, 1).Value
            xIngreso = ActiveCell.Offset(0, 5).Value
            xRegimen = ActiveCell.Offset(0, 3).Value
            xJornada = ActiveCell.Offset(0, 4).Value
            
            
            
            zLaborar = ActiveCell.Offset(0, 200).Value
            zPendiente = ActiveCell.Offset(0, 204).Value
            'zReponer = ActiveCell.Offset(0, 233).Value
            zLaboradas = ActiveCell.Offset(0, 202).Value
            zFavor = ActiveCell.Offset(0, 209).Value
            zExtra = ActiveCell.Offset(0, 203).Value
            'zPagar = ActiveCell.Offset(0, 262).Value
            zSaldoFavor = ActiveCell.Offset(0, 211).Value
            zSaldoPendiente = ActiveCell.Offset(0, 212).Value
            
                         
            
            Hoja21.Cells(3, 3) = xID
            Hoja21.Cells(4, 3) = xColaborador
            Hoja21.Cells(5, 3) = xIngreso
            Hoja21.Cells(3, 7) = xRegimen
            Hoja21.Cells(4, 7) = xJornada
            
            Hoja21.Cells(25, 3) = zLaborar
            Hoja21.Cells(25, 3).NumberFormat = "[hh]:mm"
            Hoja21.Cells(25, 6) = zLaboradas
            Hoja21.Cells(25, 6).NumberFormat = "[hh]:mm"

            Hoja21.Cells(28, 3) = zPendiente
            Hoja21.Cells(28, 3).NumberFormat = "[hh]:mm"
'            Hoja21.Cells(29, 3) = zReponer
'            Hoja21.Cells(29, 3).NumberFormat = "[hh]:mm"

            Hoja21.Cells(28, 6) = zFavor
            Hoja21.Cells(28, 6).NumberFormat = "[hh]:mm"
            Hoja21.Cells(29, 6) = zExtra
            Hoja21.Cells(29, 6).NumberFormat = "[hh]:mm"


'            Hoja21.Cells(25, 9) = zPagar
'            Hoja21.Cells(25, 9).NumberFormat = "[hh]:mm"
            Hoja21.Cells(28, 8) = zSaldoFavor
            Hoja21.Cells(28, 8).NumberFormat = "[hh]:mm"
            Hoja21.Cells(28, 9) = zSaldoPendiente
            Hoja21.Cells(28, 9).NumberFormat = "[hh]:mm"
            
            
            
            
            
            
            
            
            Conteo = 0
            
            For Hora = 1 To 16
            
            xLaborar = ActiveCell.Offset(0, 17 + Conteo).Value
            xLaborado = ActiveCell.Offset(0, 10 + Conteo).Value
            xFavor = ActiveCell.Offset(0, 11 + Conteo).Value
            xPendiente = ActiveCell.Offset(0, 12 + Conteo).Value
            xDiurnas = ActiveCell.Offset(0, 13 + Conteo).Value
            xVespertina5 = ActiveCell.Offset(0, 14 + Conteo).Value
            xNocturna6 = ActiveCell.Offset(0, 15 + Conteo).Value
            xNocturna8 = ActiveCell.Offset(0, 16 + Conteo).Value
            
            Hoja21.Cells(6 + Hora, 2) = xLaborar
            Hoja21.Cells(6 + Hora, 2).NumberFormat = "[hh]:mm"
            Hoja21.Cells(6 + Hora, 3) = xLaborado
            Hoja21.Cells(6 + Hora, 3).NumberFormat = "[hh]:mm"
            Hoja21.Cells(6 + Hora, 4) = xFavor
            Hoja21.Cells(6 + Hora, 4).NumberFormat = "[hh]:mm"
            Hoja21.Cells(6 + Hora, 5) = xPendiente
            Hoja21.Cells(6 + Hora, 5).NumberFormat = "[hh]:mm"
            Hoja21.Cells(6 + Hora, 6) = xDiurnas
            Hoja21.Cells(6 + Hora, 6).NumberFormat = "[hh]:mm"
            Hoja21.Cells(6 + Hora, 7) = xVespertina5
            Hoja21.Cells(6 + Hora, 7).NumberFormat = "[hh]:mm"
            Hoja21.Cells(6 + Hora, 8) = xNocturna6
            Hoja21.Cells(6 + Hora, 8).NumberFormat = "[hh]:mm"
            Hoja21.Cells(6 + Hora, 9) = xNocturna8
            Hoja21.Cells(6 + Hora, 9).NumberFormat = "[hh]:mm"
            
            Conteo = Conteo + 12
            
            Next
            
            Exit Do
        End If
    Loop
    
    Hoja21.Activate
    Hoja21.Cells(1, 1).Select
    
    For Columna = 2 To 9
             Hoja21.Cells(23, Columna) = WorksheetFunction.Sum(Range(Cells(7, Columna), Cells(22, Columna)))
             Hoja21.Cells(23, Columna).NumberFormat = "[hh]:mm"
    Next
    
             Hoja21.Cells(30, 3) = WorksheetFunction.Sum(Range(Cells(28, 3), Cells(29, 3)))
             Hoja21.Cells(30, 3).NumberFormat = "[hh]:mm"
             
             Hoja21.Cells(30, 6) = WorksheetFunction.Sum(Range(Cells(28, 6), Cells(29, 6)))
             Hoja21.Cells(30, 6).NumberFormat = "[hh]:mm"
    

Hoja4.Select
Range("A4").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like Referencia Then
            encontrado = True
            zValorP = ActiveCell.Offset(0, 27).Value
            
            Hoja21.Cells(30, 9) = zValorP
            Hoja21.Cells(30, 9).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

            Exit Do
        End If
    Loop

    Hoja21.Activate
    Hoja21.Cells(1, 1).Select

Application.ScreenUpdating = True



End Sub
Private Sub Formato_Derecho()
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.InsertIndent 1
End Sub
Private Sub Combinado_Centrado()
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
      Selection.Merge
      
End Sub
Private Sub Formato_Centrado()
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Private Sub Centrado_Superior()
    With Selection
        .VerticalAlignment = xlCenter
    End With
End Sub
Private Sub Formato_Izquierdo()
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub
Private Sub Fondo_Azul()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Private Sub Letra_Blanca()
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
End Sub
Private Sub Fuente_Negrita()
    Selection.Font.Bold = True
End Sub

Private Sub Fondo_Celeste_Intenso()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
End Sub

Private Sub Fondo_Celeste_Claro()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub

Private Sub Ajustar_Centrar()
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
    
Private Sub Borde_Blanco()

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
Private Sub Borde_Grueso()

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
End Sub
    


Private Sub btn_Cargar_Click()
On Error GoTo Salir
Dim Titulo As String
Dim Seguridad As String


Seguridad = Hoja83.Range("L1").Text

Titulo = "Gestor de Recursos Humanos"

 'Application.Cursor = xlWait


    If Me.txt_id = Empty Then
            MsgBox "No se seleccionado ningún colaborador..!", vbInformation, Titulo
            Exit Sub
    End If
    
   
    Hoja3.Unprotect (Seguridad)
    Hoja4.Unprotect (Seguridad)
    Hoja5.Unprotect (Seguridad)
    Hoja21.Unprotect (Seguridad)
           
    Reporte_Jornada

    Hoja3.Protect (Seguridad)
    Hoja4.Protect (Seguridad)
    Hoja5.Protect (Seguridad)
    Hoja21.Unprotect (Seguridad)
    
        MsgBox "Reporte generado con exito..!!!", vbInformation, Titulo
        
 Application.Cursor = xlDefault
       Unload Me
     
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Ventas"
 End If
End Sub

Private Sub btn_individual_Click()
banderaPersonal = 13
Call LanzarListadoPersonal(Me, "Label14")
Me.btn_Cargar.SetFocus
End Sub
