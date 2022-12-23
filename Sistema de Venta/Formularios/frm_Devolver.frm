VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Devolver 
   Caption         =   "GESTOR DE CAJA"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   7220
   OleObjectBlob   =   "frm_Devolver.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Devolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long
Dim vPrecioVenta As Currency
Dim vImporte As Currency
Dim CostoUnitario As Currency

Private Sub CommandButton2_Click()
Unload Me
End Sub
Private Sub UserForm_Initialize()
Me.txt_Fecha = Date
Me.lbl_devolucion = "DEVOLUCIÓN N° " & Hoja93.Range("J2") + 1
Me.txt_monto = frm_Devolucion.txtTotal
End Sub
Private Sub ProcesarDevolver()
Dim Fila As Long
Dim Final As Long
Dim Existencia As Long
Dim TotalExistencia As Long
Dim Comprb As Long
Dim nDevolver As Long
Dim CostoTotal As Currency
Dim cUpromedio As Currency
Dim xCantidad As Currency
Dim xCodigo As String
Dim xDescrip As String
Dim xCosto As Currency

'Correlativo de la Devolver de venta
Comprb = Hoja93.Range("J2").Value + 1

''Envía los datos a la hoja de ventas

If Hoja25.Visible = xlSheetVisible Then

                For i = 0 To frm_Devolucion.ListBox1.ListCount - 1
                    xCodigo = frm_Devolucion.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Devolucion.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Devolucion.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Devolucion.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Devolucion.ListBox1.List(i, 4) 'Importe

                Hoja25.Select
                    Hoja25.Range("A2:N2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja25.Range("A3:N3").Select
                    Selection.Copy
                    Hoja25.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja25.Cells(2, 1) = Hoja25.Cells(3, 1) + 1
                    Hoja25.Cells(2, 2) = CDate(frm_Devolucion.txtFecha)
                    Hoja25.Cells(2, 4) = frm_Devolucion.txt_idcliente.Text
                    Hoja25.Cells(2, 5) = frm_Devolucion.txtCliente.Text
                    Hoja25.Cells(2, 6) = Format(Time)
                    Hoja25.Cells(2, 8) = xCodigo
                    Hoja25.Cells(2, 9) = xDescrip
                    Hoja25.Cells(2, 10) = Comprb
                    Hoja25.Cells(2, 11) = xCantidad
                    Hoja25.Cells(2, 12) = vPrecioVenta
                    Hoja25.Cells(2, 14) = Hoja92.Range("G1")

                    Final = Final + 1
                Next

ElseIf Hoja25.Visible = xlSheetVeryHidden Then
    Hoja25.Visible = xlSheetVisible

                For i = 0 To frm_Devolucion.ListBox1.ListCount - 1
                    xCodigo = frm_Devolucion.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Devolucion.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Devolucion.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Devolucion.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Devolucion.ListBox1.List(i, 4) 'Importe

                Hoja25.Select
                    Hoja25.Range("A2:N2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja25.Range("A3:N3").Select
                    Selection.Copy
                    Hoja25.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False



                    Application.CutCopyMode = False

                    Hoja25.Cells(2, 1) = Hoja25.Cells(3, 1) + 1
                    Hoja25.Cells(2, 2) = CDate(frm_Devolucion.txtFecha)
                    Hoja25.Cells(2, 4) = frm_Devolucion.txt_idcliente.Text
                    Hoja25.Cells(2, 5) = frm_Devolucion.txtCliente.Text
                    Hoja25.Cells(2, 6) = Format(Time)
                    Hoja25.Cells(2, 8) = xCodigo
                    Hoja25.Cells(2, 9) = xDescrip
                    Hoja25.Cells(2, 10) = Comprb
                    Hoja25.Cells(2, 11) = xCantidad
                    Hoja25.Cells(2, 12) = vPrecioVenta
                    Hoja25.Cells(2, 14) = Hoja92.Range("G1")

                    Final = Final + 1
                Next

   Hoja25.Visible = xlSheetVeryHidden

End If

End Sub
Private Sub Devolucion()
Dim Fila As Long
Dim Final As Long
Dim Existencia As Long
Dim TotalExistencia As Long
Dim Comprb As Long
Dim nDevolver As Long
Dim CostoTotal As Currency
Dim cUpromedio As Currency
Dim xCantidad As Long
Dim xCodigo As String
Dim xDescrip As String

Comprb = Hoja93.Range("J2").Value + 1

If Hoja23.Visible = xlSheetVisible Then

                Hoja23.Select
                    Hoja23.Range("A2:K2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja23.Range("A3:K3").Select
                    Selection.Copy
                    Hoja23.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja23.Cells(2, 1) = Hoja23.Cells(3, 1) + 1
                    Hoja23.Cells(2, 2) = CDate(frm_Devolver.txt_Fecha)
                    Hoja23.Cells(2, 4) = frm_Devolucion.txt_idcliente.Text
                    Hoja23.Cells(2, 5) = frm_Devolucion.txtCliente.Text
                    Hoja23.Cells(2, 6) = Format(Time)
                    Hoja23.Cells(2, 7) = Comprb
                    Hoja23.Cells(2, 8) = Me.txt_monto.Text
                    Hoja23.Cells(2, 9) = Me.txt_DevFactura.Text
                    Hoja23.Cells(2, 10) = Me.txt_DevObserva.Text
                    Hoja23.Cells(2, 11) = Hoja92.Range("G1")


ElseIf Hoja23.Visible = xlSheetVeryHidden Then
    Hoja23.Visible = xlSheetVisible


                Hoja23.Select
                    Hoja23.Range("A2:K2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja23.Range("A3:K3").Select
                    Selection.Copy
                    Hoja23.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja23.Cells(2, 1) = Hoja23.Cells(3, 1) + 1
                    Hoja23.Cells(2, 2) = CDate(frm_Devolver.txt_Fecha)
                    Hoja23.Cells(2, 4) = frm_Devolucion.txt_idcliente.Text
                    Hoja23.Cells(2, 5) = frm_Devolucion.txtCliente.Text
                    Hoja23.Cells(2, 6) = Format(Time)
                    Hoja23.Cells(2, 7) = Comprb
                    Hoja23.Cells(2, 8) = Me.txt_monto.Text
                    Hoja23.Cells(2, 9) = Me.txt_DevFactura.Text
                    Hoja23.Cells(2, 10) = Me.txt_DevObserva.Text
                    Hoja23.Cells(2, 11) = Hoja92.Range("G1")

   Hoja23.Visible = xlSheetVeryHidden
End If

End Sub
Private Sub xTemporal()
Dim Fila As Long
Dim Final As Long
Dim Detalle As String

'Correlativo de la Devolver de venta

Detalle = "DEVOLUCIÓN DE EFECTIVO"

''Envía los datos a la hoja de ventas

If Hoja26.Visible = xlSheetVisible Then

                Hoja26.Select
                    Hoja26.Range("A2:Q2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja26.Range("A3:Q3").Select
                    Selection.Copy
                    Hoja26.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja26.Cells(2, 1) = Hoja22.Cells(2, 1) + 1
                    Hoja26.Cells(2, 2) = CDate(frm_Devolver.txt_Fecha)
                    Hoja26.Cells(2, 4) = Format(Time)
                    Hoja26.Cells(2, 5) = frm_Devolver.lbl_devolucion.Caption
                    Hoja26.Cells(2, 6) = Detalle
                    Hoja26.Cells(2, 8) = frm_Devolver.txt_monto.Text
                    Hoja26.Cells(2, 14) = frm_Devolver.txt_monto.Text
                    Hoja26.Cells(2, 17) = Hoja92.Range("G1")


ElseIf Hoja26.Visible = xlSheetVeryHidden Then
    Hoja26.Visible = xlSheetVisible

                Hoja26.Select
                    Hoja26.Range("A2:Q2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja26.Range("A3:Q3").Select
                    Selection.Copy
                    Hoja26.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja26.Cells(2, 1) = Hoja22.Cells(2, 1) + 1
                    Hoja26.Cells(2, 2) = CDate(frm_Devolver.txt_Fecha)
                    Hoja26.Cells(2, 4) = Format(Time)
                    Hoja26.Cells(2, 5) = frm_Devolver.lbl_devolucion.Caption
                    Hoja26.Cells(2, 6) = Detalle
                    Hoja26.Cells(2, 8) = frm_Devolver.txt_monto.Text
                    Hoja26.Cells(2, 14) = frm_Devolver.txt_monto.Text
                    Hoja26.Cells(2, 17) = Hoja92.Range("G1")

   Hoja26.Visible = xlSheetVeryHidden
End If

End Sub
Private Sub zHistorico()
Dim Fila As Long
Dim Final As Long
Dim Detalle As String

'Correlativo de la Devolver de venta

Detalle = "DEVOLUCIÓN DE EFECTIVO"

''Envía los datos a la hoja de ventas

If Hoja22.Visible = xlSheetVisible Then

                Hoja22.Select
                    Hoja22.Range("A2:Q2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja22.Range("A3:Q3").Select
                    Selection.Copy
                    Hoja22.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja22.Cells(2, 1) = Hoja22.Cells(3, 1) + 1
                    Hoja22.Cells(2, 2) = CDate(frm_Devolver.txt_Fecha)
                    Hoja22.Cells(2, 4) = Format(Time)
                    Hoja22.Cells(2, 5) = frm_Devolver.lbl_devolucion.Caption
                    Hoja22.Cells(2, 6) = Detalle
                    Hoja22.Cells(2, 8) = frm_Devolver.txt_monto.Text
                    Hoja22.Cells(2, 14) = frm_Devolver.txt_monto.Text
                    Hoja22.Cells(2, 17) = Hoja92.Range("G1")


ElseIf Hoja22.Visible = xlSheetVeryHidden Then
    Hoja22.Visible = xlSheetVisible

                Hoja22.Select
                    Hoja22.Range("A2:Q2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja22.Range("A3:Q3").Select
                    Selection.Copy
                    Hoja22.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja22.Cells(2, 1) = Hoja22.Cells(3, 1) + 1
                    Hoja22.Cells(2, 2) = CDate(frm_Devolver.txt_Fecha)
                    Hoja22.Cells(2, 4) = Format(Time)
                    Hoja22.Cells(2, 5) = frm_Devolver.lbl_devolucion.Caption
                    Hoja22.Cells(2, 6) = Detalle
                    Hoja22.Cells(2, 8) = frm_Devolver.txt_monto.Text
                    Hoja22.Cells(2, 14) = frm_Devolver.txt_monto.Text
                    Hoja22.Cells(2, 17) = Hoja92.Range("G1")

   Hoja22.Visible = xlSheetVeryHidden
End If

End Sub
Private Sub btn_registrar_Click()
    If txt_monto = "" Then
        MsgBox "Debe registrar el efectivo", vbInformation, "GESTOR DE CAJA"
        txt_monto.SetFocus
        Exit Sub
    End If
    If Me.txt_DevFactura = "" Then
        MsgBox "Debe colocar el número de Devolver original de la venta", vbInformation, "GESTOR DE CAJA"
        Me.txt_DevFactura.SetFocus
        Exit Sub
    End If

    If Me.txt_DevObserva = "" Then
        MsgBox "Debe escribir las observaciones sobre la devolución realizada", vbInformation, "GESTOR DE CAJA"
        Me.txt_DevObserva.SetFocus
        Exit Sub
    End If

    If MsgBox("Son correctos los datos?", vbYesNo, "Gestor de Ventas") = vbNo Then
        Exit Sub
    Else
        Application.ScreenUpdating = False
        Hoja25.Unprotect ""
        ProcesarDevolver
        Devolucion
        EntradaInventario
        xTemporal
        zHistorico
'         Application.EnableEvents = False
'        Reporte
'         Application.EnableEvents = True

        MsgBox "Devolución de efectivo realizada con éxito!!!", , "Gestor de Caja"
        Unload Me
    End If
    
        Unload frm_Devolucion
        
        Hoja25.Protect ""
        Hoja93.Range("J2") = Hoja93.Range("J2") + 1
        
            Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True
    
    Application.ScreenUpdating = True
End Sub

Private Sub EntradaInventario()
Dim Devolucion As String
Dim nDevolver As Long
Dim Comprobante As String
Dim xCategoria As String
Dim xCantidad As Double
Dim xCodigo As String
Dim xDescrip As String
Dim xPrecioVenta As Currency

Devolucion = "CLIENTE N° "
nDevolver = Hoja93.Range("J2").Value + 1
Comprobante = "Devolución N° " & nDevolver


    Hoja3.Unprotect ""

    If Hoja3.Visible = xlSheetVisible Then

                For i = 0 To frm_Devolucion.ListBox1.ListCount - 1
                    xCodigo = frm_Devolucion.ListBox1.List(i, 0)
                    xCantidad = frm_Devolucion.ListBox1.List(i, 1)
                    xDescrip = frm_Devolucion.ListBox1.List(i, 2)
                    vPrecioVenta = 0
                    xCategoria = "PRODUCTO"

                Hoja3.Select
                    Hoja3.Range("A2:L2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja3.Range("A3:L3").Select
                    Selection.Copy
                    Hoja3.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja3.Cells(2, 1) = CDate(frm_Devolucion.txtFecha)
                    Hoja3.Cells(2, 3) = Comprobante
                    Hoja3.Cells(2, 4) = Devolucion & frm_Devolucion.txt_idcliente.Text
                    Hoja3.Cells(2, 6) = xDescrip
                    Hoja3.Cells(2, 7) = xCantidad
                    Hoja3.Cells(2, 9) = vPrecioVenta
                    Hoja3.Cells(2, 11) = xCategoria
                    Hoja3.Cells(2, 12) = Hoja92.Range("G1")

                    Final = Final + 1
                Next

    ElseIf Hoja3.Visible = xlSheetVeryHidden Then
        Hoja3.Visible = xlSheetVisible

                    For i = 0 To frm_Devolucion.ListBox1.ListCount - 1
                    xCodigo = frm_Devolucion.ListBox1.List(i, 0)
                    xCantidad = frm_Devolucion.ListBox1.List(i, 1)
                    xDescrip = frm_Devolucion.ListBox1.List(i, 2)
                    vPrecioVenta = 0
                    xCategoria = "PRODUCTO"

                Hoja3.Select
                    Hoja3.Range("A2:L2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja3.Range("A3:L3").Select
                    Selection.Copy
                    Hoja3.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja3.Cells(2, 1) = CDate(frm_Devolucion.txtFecha)
                    Hoja3.Cells(2, 3) = Comprobante
                    Hoja3.Cells(2, 4) = Devolucion & frm_Devolucion.txt_idcliente.Text
                    Hoja3.Cells(2, 6) = xDescrip
                    Hoja3.Cells(2, 7) = xCantidad
                    Hoja3.Cells(2, 9) = vPrecioVenta
                    Hoja3.Cells(2, 11) = xCategoria
                    Hoja3.Cells(2, 12) = Hoja92.Range("G1")

                    Final = Final + 1
                Next

    Hoja3.Visible = xlSheetVeryHidden

End If

    Hoja3.Protect ""

End Sub


Private Sub Reporte()
Dim Fila As Long
Dim Final As Long
Dim Existencia As Long
Dim TotalExistencia As Long
Dim Comprb As Long
Dim nDevolver As Long
Dim CostoTotal As Currency
Dim cUpromedio As Currency
Dim xCantidad As Currency
Dim xCodigo As String
Dim xDescrip As String
Dim xCosto As Currency
Dim FiladelTotal As Integer
Dim ValorSaldo As Double

Final = 14

If Hoja10.Visible = xlSheetVisible Then

                    'LIMPIAR HOJA
                    Hoja10.Select
                            Range("A14:D1000").Select
                                Selection.ClearContents
                                With Selection
                                    .HorizontalAlignment = xlGeneral
                                    .VerticalAlignment = xlBottom
                                    .WrapText = False
                                    .Orientation = 0
                                    .AddIndent = False
                                    .IndentLevel = 0
                                    .ShrinkToFit = False
                                    .ReadingOrder = xlContext
                                    .MergeCells = True
                                End With

                                Selection.UnMerge
                                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                                Selection.Borders(xlEdgeLeft).LineStyle = xlNone
                                Selection.Borders(xlEdgeTop).LineStyle = xlNone
                                Selection.Borders(xlEdgeBottom).LineStyle = xlNone
                                Selection.Borders(xlEdgeRight).LineStyle = xlNone
                                Selection.Borders(xlInsideVertical).LineStyle = xlNone
                                Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                                Selection.NumberFormat = "General"

                                With Selection.Font
                                    .Name = "Calibri"
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

    Hoja10.Rows("14:1000").Select
    Selection.RowHeight = 15

                    Hoja10.Select
                    Hoja10.Cells(8, 1) = "DEVOLUCIÓN SOBRE VENTA"
                    Hoja10.Cells(9, 1) = "CLIENTE: " & UCase(frm_Devolucion.txtCliente.Text)
                    Hoja10.Cells(10, 1) = "N° RUC: " & UCase(frm_Devolucion.txt_Ruc.Text)

 For i = 0 To frm_Devolucion.ListBox1.ListCount - 1
                    xCodigo = frm_Devolucion.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Devolucion.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Devolucion.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Devolucion.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Devolucion.ListBox1.List(i, 4) 'Importe


                Hoja10.Select

                    Hoja10.Cells(Final, 1) = xCantidad
                    Hoja10.Cells(Final, 1).NumberFormat = "0.00"
                    Hoja10.Cells(Final, 1).HorizontalAlignment = xlCenter
                    Hoja10.Cells(Final, 2) = xDescrip
                    Hoja10.Cells(Final, 4) = vPrecioVenta
                    Hoja10.Cells(Final, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    Hoja10.Cells(Final + 1, 2) = "SUB TOTAL"
                    Hoja10.Cells(Final + 1, 3) = vImporte
                    Hoja10.Cells(Final + 1, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

                    Final = Final + 2
                Next

'Determinar el final para colocar el saldo de Devolucion


        For FiladelTotal = 14 To 1000
            If Hoja10.Cells(FiladelTotal, 2) = "" Then
                saldototal = FiladelTotal
                Exit For
            End If
        Next

            Hoja10.Cells(saldototal, 1).Select


            Hoja10.Cells(saldototal + 2, 1) = "TOTAL DEVOLUCIÓN"
            Hoja10.Cells(saldototal + 2, 3) = frm_Devolucion.txtTotal.Text
            Hoja10.Cells(saldototal + 2, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

                Range(Cells(saldototal + 2, 3), Cells(saldototal + 2, 4)).Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
            Selection.Merge

           Hoja10.Cells(saldototal + 3, 1) = "N° FACTURA: " & frm_Devolver.txt_DevFactura.Text
           
           Hoja10.Cells(saldototal + 4, 1) = "OBSERVACIONES:"
           Hoja10.Cells(saldototal + 5, 1) = frm_Devolver.txt_DevObserva.Text
           
            Range(Cells(saldototal + 5, 1), Cells(saldototal + 5, 4)).Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
            Selection.Merge

            Hoja10.Cells(saldototal + 7, 1) = "CAJERO:"
            Hoja10.Cells(saldototal + 7, 2) = UCase(frm_Devolucion.txt_usuario.Text)
            Hoja10.Cells(saldototal + 8, 1) = "FECHA: " & Format(Date) & "  " & Format(Time)
            Hoja10.Cells(saldototal + 9, 1) = "REFERENCIA: " & UCase(Hoja93.Range("J2").Value + 1)

           Hoja10.Cells(saldototal + 12, 1) = "GRACIAS POR PREFERIRNOS"
           Hoja10.Cells(saldototal + 13, 1) = "¡DIOS LE BENDIGA!"

            Range(Cells(saldototal + 12, 1), Cells(saldototal + 12, 4)).Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
            Selection.Merge

             Range(Cells(saldototal + 12, 1), Cells(saldototal + 13, 4)).Select
                            With Selection.Font
                    .Name = "Maiandra GD"
                    .Size = 9
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                    .Underline = xlUnderlineStyleNone
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0
                    .ThemeFont = xlThemeFontNone
                End With
                Selection.Font.Bold = True

           Range(Cells(saldototal + 13, 1), Cells(saldototal + 13, 4)).Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
            Selection.Merge

            Range(Cells(saldototal + 2, 1), Cells(saldototal + 8, 1)).Select

                With Selection
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                End With
                    Columns("A:D").Select
                Range("A12").Activate
                With Selection.Font
                    .Size = 9
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                    .Underline = xlUnderlineStyleNone
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0
                End With

                Hoja10.Rows("11:12").Select
                Selection.Copy
                Rows(saldototal).Select
                ActiveSheet.Paste
                Rows(saldototal + 10).Select
                ActiveSheet.Paste

                Hoja10.Select
                Hoja10.Cells(1, 1).Select

                    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                IgnorePrintAreas:=False


ElseIf Hoja10.Visible = xlSheetVeryHidden Then
        Hoja10.Visible = xlSheetVisible


                    'LIMPIAR HOJA
                    Hoja10.Select
                            Range("A14:D1000").Select
                                Selection.ClearContents
                                With Selection
                                    .HorizontalAlignment = xlGeneral
                                    .VerticalAlignment = xlBottom
                                    .WrapText = False
                                    .Orientation = 0
                                    .AddIndent = False
                                    .IndentLevel = 0
                                    .ShrinkToFit = False
                                    .ReadingOrder = xlContext
                                    .MergeCells = True
                                End With

                                Selection.UnMerge
                                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                                Selection.Borders(xlEdgeLeft).LineStyle = xlNone
                                Selection.Borders(xlEdgeTop).LineStyle = xlNone
                                Selection.Borders(xlEdgeBottom).LineStyle = xlNone
                                Selection.Borders(xlEdgeRight).LineStyle = xlNone
                                Selection.Borders(xlInsideVertical).LineStyle = xlNone
                                Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                                Selection.NumberFormat = "General"

                                With Selection.Font
                                    .Name = "Calibri"
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

    Hoja10.Rows("14:1000").Select
    Selection.RowHeight = 15

                    Hoja10.Select
                    Hoja10.Cells(8, 1) = "DEVOLUCIÓN SOBRE VENTA"
                    Hoja10.Cells(9, 1) = "CLIENTE: " & UCase(frm_Devolucion.txtCliente.Text)
                    Hoja10.Cells(10, 1) = "N° RUC: " & UCase(frm_Devolucion.txt_Ruc.Text)

 For i = 0 To frm_Devolucion.ListBox1.ListCount - 1
                    xCodigo = frm_Devolucion.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Devolucion.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Devolucion.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Devolucion.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Devolucion.ListBox1.List(i, 4) 'Importe


                Hoja10.Select

                    Hoja10.Cells(Final, 1) = xCantidad
                    Hoja10.Cells(Final, 1).NumberFormat = "0.00"
                    Hoja10.Cells(Final, 1).HorizontalAlignment = xlCenter
                    Hoja10.Cells(Final, 2) = xDescrip
                    Hoja10.Cells(Final, 4) = vPrecioVenta
                    Hoja10.Cells(Final, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    Hoja10.Cells(Final + 1, 2) = "SUB TOTAL"
                    Hoja10.Cells(Final + 1, 3) = vImporte
                    Hoja10.Cells(Final + 1, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

                    Final = Final + 2
                Next

'Determinar el final para colocar el saldo de Devolucion


        For FiladelTotal = 14 To 1000
            If Hoja10.Cells(FiladelTotal, 2) = "" Then
                saldototal = FiladelTotal
                Exit For
            End If
        Next

            Hoja10.Cells(saldototal, 1).Select


            Hoja10.Cells(saldototal + 2, 1) = "TOTAL DEVOLUCIÓN"
            Hoja10.Cells(saldototal + 2, 3) = frm_Devolucion.txtTotal.Text
            Hoja10.Cells(saldototal + 2, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

                Range(Cells(saldototal + 2, 3), Cells(saldototal + 2, 4)).Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
            Selection.Merge

           Hoja10.Cells(saldototal + 3, 1) = "N° FACTURA: " & frm_Devolver.txt_DevFactura.Text

           Hoja10.Cells(saldototal + 4, 1) = "OBSERVACIONES:"
           Hoja10.Cells(saldototal + 5, 1) = frm_Devolver.txt_DevObserva.Text
           
            Range(Cells(saldototal + 5, 1), Cells(saldototal + 5, 4)).Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
            Selection.Merge

            Hoja10.Cells(saldototal + 7, 1) = "CAJERO:"
            Hoja10.Cells(saldototal + 7, 2) = UCase(frm_Devolucion.txt_usuario.Text)
            Hoja10.Cells(saldototal + 8, 1) = "FECHA: " & Format(Date) & "  " & Format(Time)
            Hoja10.Cells(saldototal + 9, 1) = "REFERENCIA: " & UCase(Hoja93.Range("J2").Value + 1)

           Hoja10.Cells(saldototal + 12, 1) = "GRACIAS POR PREFERIRNOS"
           Hoja10.Cells(saldototal + 13, 1) = "¡DIOS LE BENDIGA!"

            Range(Cells(saldototal + 12, 1), Cells(saldototal + 12, 4)).Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
            Selection.Merge

             Range(Cells(saldototal + 12, 1), Cells(saldototal + 13, 4)).Select
                            With Selection.Font
                    .Name = "Maiandra GD"
                    .Size = 9
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                    .Underline = xlUnderlineStyleNone
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0
                    .ThemeFont = xlThemeFontNone
                End With
                Selection.Font.Bold = True

           Range(Cells(saldototal + 13, 1), Cells(saldototal + 13, 4)).Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
            Selection.Merge

            Range(Cells(saldototal + 2, 1), Cells(saldototal + 8, 1)).Select

                With Selection
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                End With
                    Columns("A:D").Select
                Range("A12").Activate
                With Selection.Font
                    .Size = 9
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                    .Underline = xlUnderlineStyleNone
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0
                End With

                Hoja10.Rows("11:12").Select
                Selection.Copy
                Rows(saldototal).Select
                ActiveSheet.Paste
                Rows(saldototal + 10).Select
                ActiveSheet.Paste

                Hoja10.Select
                Hoja10.Cells(1, 1).Select

                    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                IgnorePrintAreas:=False

    Hoja10.Visible = xlSheetVeryHidden

End If

End Sub



