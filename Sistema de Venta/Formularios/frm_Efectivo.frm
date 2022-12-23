VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Efectivo 
   Caption         =   "FACTURACIÓN"
   ClientHeight    =   9540.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   6890
   OleObjectBlob   =   "frm_Efectivo.frx":0000
End
Attribute VB_Name = "frm_Efectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Long
Dim vPrecioVenta As Currency
Dim vImporte As Currency
Dim CostoUnitario As Currency
Private Sub Facturación()
Dim Fila As Long
Dim Final As Long
Dim Existencia As Long
Dim TotalExistencia As Long
Dim Comprb As Long
Dim nFactura As Long
Dim CostoTotal As Currency
Dim cUpromedio As Currency
Dim xCantidad As Long
Dim xCodigo As String
Dim xDescrip As String
Dim xTipoPago As String


'Correlativo de la factura de venta

Comprb = Hoja93.Range("C2").Value + 1
xTipoPago = "EFECTIVO"

''Envía los datos a la hoja de ventas

If Hoja21.Visible = xlSheetVisible Then

                Hoja21.Select
                    Hoja21.Range("A2:K2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja21.Range("A3:K3").Select
                    Selection.Copy
                    Hoja21.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja21.Cells(2, 1) = Hoja21.Cells(3, 1) + 1
                    Hoja21.Cells(2, 2) = CDate(frm_Factura.txtFecha)
                    Hoja21.Cells(2, 4) = frm_Factura.txt_idcliente.Text
                    Hoja21.Cells(2, 5) = frm_Factura.txtCliente.Text
                    Hoja21.Cells(2, 6) = Format(Time)
                    Hoja21.Cells(2, 7) = xTipoPago
                    Hoja21.Cells(2, 8) = Comprb
                    Hoja21.Cells(2, 9) = frm_Factura.txtSubtotal.Text
                    Hoja21.Cells(2, 10) = frm_Factura.txt_Abono.Text
                    Hoja21.Cells(2, 11) = Me.TextBox3.Text
                    Hoja21.Cells(2, 12) = Me.TextBox1.Text
                    Hoja21.Cells(2, 13) = Me.TextBox2.Text
                    Hoja21.Cells(2, 14) = Hoja92.Range("G1")


ElseIf Hoja21.Visible = xlSheetVeryHidden Then
    Hoja21.Visible = xlSheetVisible


                Hoja21.Select
                    Hoja21.Range("A2:K2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja21.Range("A3:K3").Select
                    Selection.Copy
                    Hoja21.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja21.Cells(2, 1) = Hoja21.Cells(3, 1) + 1
                    Hoja21.Cells(2, 2) = CDate(frm_Factura.txtFecha)
                    Hoja21.Cells(2, 4) = frm_Factura.txt_idcliente.Text
                    Hoja21.Cells(2, 5) = frm_Factura.txtCliente.Text
                    Hoja21.Cells(2, 6) = Format(Time)
                    Hoja21.Cells(2, 7) = xTipoPago
                    Hoja21.Cells(2, 8) = Comprb
                    Hoja21.Cells(2, 9) = frm_Factura.txtSubtotal.Text
                    Hoja21.Cells(2, 10) = frm_Factura.txt_Abono.Text
                    Hoja21.Cells(2, 11) = Me.TextBox3.Text
                    Hoja21.Cells(2, 12) = Me.TextBox1.Text
                    Hoja21.Cells(2, 13) = Me.TextBox2.Text
                    Hoja21.Cells(2, 14) = Hoja92.Range("G1")

   Hoja21.Visible = xlSheetVeryHidden
End If

End Sub
Private Sub xTemporal()
Dim Fila As Long
Dim Final As Long
Dim Detalle As String

'Correlativo de la factura de venta

Detalle = "INGRESO POR VENTAS"

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
                    Hoja26.Cells(2, 2) = CDate(frm_Factura.txtFecha)
                    Hoja26.Cells(2, 4) = Format(Time)
                    Hoja26.Cells(2, 5) = frm_Factura.lbl_nFactura.Caption
                    Hoja26.Cells(2, 6) = Detalle
                    Hoja26.Cells(2, 7) = frm_Efectivo.TextBox3.Text
                    Hoja26.Cells(2, 10) = frm_Efectivo.TextBox3.Text
                    Hoja26.Cells(2, 11) = frm_Efectivo.TextBox3.Text
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
                    Hoja26.Cells(2, 2) = CDate(frm_Factura.txtFecha)
                    Hoja26.Cells(2, 4) = Format(Time)
                    Hoja26.Cells(2, 5) = frm_Factura.lbl_nFactura.Caption
                    Hoja26.Cells(2, 6) = Detalle
                    Hoja26.Cells(2, 7) = frm_Efectivo.TextBox3.Text
                    Hoja26.Cells(2, 10) = frm_Efectivo.TextBox3.Text
                    Hoja26.Cells(2, 11) = frm_Efectivo.TextBox3.Text
                    Hoja26.Cells(2, 17) = Hoja92.Range("G1")

   Hoja26.Visible = xlSheetVeryHidden
End If

End Sub
Private Sub zHistorico()
Dim Fila As Long
Dim Final As Long
Dim Detalle As String

'Correlativo de la factura de venta

Detalle = "INGRESO POR VENTAS"

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
                    Hoja22.Cells(2, 2) = CDate(frm_Factura.txtFecha)
                    Hoja22.Cells(2, 4) = Format(Time)
                    Hoja22.Cells(2, 5) = frm_Factura.lbl_nFactura.Caption
                    Hoja22.Cells(2, 6) = Detalle
                    Hoja22.Cells(2, 7) = frm_Efectivo.TextBox3.Text
                    Hoja22.Cells(2, 10) = frm_Efectivo.TextBox3.Text
                    Hoja22.Cells(2, 11) = frm_Efectivo.TextBox3.Text
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
                    Hoja22.Cells(2, 2) = CDate(frm_Factura.txtFecha)
                    Hoja22.Cells(2, 4) = Format(Time)
                    Hoja22.Cells(2, 5) = frm_Factura.lbl_nFactura.Caption
                    Hoja22.Cells(2, 6) = Detalle
                    Hoja22.Cells(2, 7) = frm_Efectivo.TextBox3.Text
                    Hoja22.Cells(2, 10) = frm_Efectivo.TextBox3.Text
                    Hoja22.Cells(2, 11) = frm_Efectivo.TextBox3.Text
                    Hoja22.Cells(2, 17) = Hoja92.Range("G1")

   Hoja22.Visible = xlSheetVeryHidden
End If

End Sub
Private Sub SalidaInventario()
Dim Registro As String
Dim nFactura As Long
Dim Comprobante As String
Dim Categoria As String
Dim xCantidad As Double
Dim xCodigo As String
Dim xDescrip As String
Dim xPrecioVenta As Currency

Registro = "VENTA"
nFactura = Hoja93.Range("C2").Value + 1
Comprobante = "FACTURA N° " & nFactura
Categoria = "PRODUCTO"

    Hoja4.Unprotect ""

    If Hoja4.Visible = xlSheetVisible Then

                For i = 0 To frm_Factura.ListBox1.ListCount - 1
                    xCodigo = frm_Factura.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Factura.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Factura.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Factura.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Factura.ListBox1.List(i, 4) 'Importe

                Hoja4.Select
                    Hoja4.Range("A2:L2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja4.Range("A3:L3").Select
                    Selection.Copy
                    Hoja4.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja4.Cells(2, 1) = CDate(frm_Factura.txtFecha)
                    Hoja4.Cells(2, 3) = Registro
                    Hoja4.Cells(2, 5) = xDescrip
                    Hoja4.Cells(2, 6) = xCantidad
                    Hoja4.Cells(2, 8) = 0
                    Hoja4.Cells(2, 10) = Comprobante
                    Hoja4.Cells(2, 11) = Categoria
                    Hoja4.Cells(2, 12) = Hoja92.Range("G1")

                    Final = Final + 1
                Next

    ElseIf Hoja4.Visible = xlSheetVeryHidden Then
        Hoja4.Visible = xlSheetVisible

                    For i = 0 To frm_Factura.ListBox1.ListCount - 1
                    xCodigo = frm_Factura.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Factura.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Factura.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Factura.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Factura.ListBox1.List(i, 4) 'Importe

                Hoja4.Select
                    Hoja4.Range("A2:L2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja4.Range("A3:L3").Select
                    Selection.Copy
                    Hoja4.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja4.Cells(2, 1) = CDate(frm_Factura.txtFecha)
                    Hoja4.Cells(2, 3) = Registro
                    Hoja4.Cells(2, 5) = xDescrip
                    Hoja4.Cells(2, 6) = xCantidad
                    Hoja4.Cells(2, 8) = 0
                    Hoja4.Cells(2, 10) = Comprobante
                    Hoja4.Cells(2, 11) = Categoria
                    Hoja4.Cells(2, 12) = Hoja92.Range("G1")

                    Final = Final + 1
                Next

    Hoja4.Visible = xlSheetVeryHidden

End If

    Hoja4.Protect ""

End Sub
Private Sub Ticket()
Dim Fila As Long
Dim Final As Long
Dim Existencia As Long
Dim TotalExistencia As Long
Dim Comprb As Long
Dim nFactura As Long
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
                    Hoja10.Cells(8, 1) = "FACTURA ORIGINAL"
                    Hoja10.Cells(9, 1) = "CLIENTE: " & UCase(frm_Factura.txtCliente.Text)
                    Hoja10.Cells(10, 1) = "N° RUC: " & UCase(frm_Factura.txt_Ruc.Text)

 For i = 0 To frm_Factura.ListBox1.ListCount - 1
                    xCodigo = frm_Factura.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Factura.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Factura.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Factura.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Factura.ListBox1.List(i, 4) 'Importe

  
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
                
'Determinar el final para colocar el saldo de Factura


        For FiladelTotal = 14 To 1000
            If Hoja10.Cells(FiladelTotal, 2) = "" Then
                saldototal = FiladelTotal
                Exit For
            End If
        Next
        
            Hoja10.Cells(saldototal, 1).Select
            
            
            Hoja10.Cells(saldototal + 2, 1) = "TOTAL FACTURA"
            Hoja10.Cells(saldototal + 2, 3) = frm_Factura.txtTotal.Text
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
            
           Hoja10.Cells(saldototal + 3, 1) = "PAGO EN EFECTIVO:"
           Hoja10.Cells(saldototal + 3, 3) = frm_Efectivo.TextBox1.Text
           Hoja10.Cells(saldototal + 3, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
           

            Range(Cells(saldototal + 3, 3), Cells(saldototal + 3, 4)).Select
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
            
           Hoja10.Cells(saldototal + 4, 1) = "CAMBIO:"
           Hoja10.Cells(saldototal + 4, 3) = frm_Efectivo.TextBox2.Text
           Hoja10.Cells(saldototal + 4, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

            Range(Cells(saldototal + 4, 3), Cells(saldototal + 4, 4)).Select
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
            
            Hoja10.Cells(saldototal + 6, 1) = "CAJERO:"
            Hoja10.Cells(saldototal + 6, 2) = UCase(frm_Factura.txt_usuario.Text)
            Hoja10.Cells(saldototal + 7, 1) = "FECHA: " & Format(Date) & "  " & Format(Time)
            Hoja10.Cells(saldototal + 8, 1) = "REFERENCIA: " & UCase(Hoja93.Range("C2").Value + 1)
            
           Hoja10.Cells(saldototal + 11, 1) = "GRACIAS POR PREFERIRNOS"
           Hoja10.Cells(saldototal + 12, 1) = "¡DIOS LE BENDIGA!"

            Range(Cells(saldototal + 11, 1), Cells(saldototal + 11, 4)).Select
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
            
             Range(Cells(saldototal + 11, 1), Cells(saldototal + 12, 1)).Select
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
                Rows(saldototal + 9).Select
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
                    Hoja10.Cells(8, 1) = "FACTURA ORIGINAL"
                    Hoja10.Cells(9, 1) = "CLIENTE: " & UCase(frm_Factura.txtCliente.Text)
                    Hoja10.Cells(10, 1) = "N° RUC: " & UCase(frm_Factura.txt_Ruc.Text)

 For i = 0 To frm_Factura.ListBox1.ListCount - 1
                    xCodigo = frm_Factura.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Factura.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Factura.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Factura.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Factura.ListBox1.List(i, 4) 'Importe

  
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
                
'Determinar el final para colocar el saldo de Factura


        For FiladelTotal = 14 To 1000
            If Hoja10.Cells(FiladelTotal, 2) = "" Then
                saldototal = FiladelTotal
                Exit For
            End If
        Next
        
            Hoja10.Cells(saldototal, 1).Select
            
            
            Hoja10.Cells(saldototal + 2, 1) = "TOTAL FACTURA"
            Hoja10.Cells(saldototal + 2, 3) = frm_Factura.txtTotal.Text
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
            
           Hoja10.Cells(saldototal + 3, 1) = "PAGO EN EFECTIVO:"
           Hoja10.Cells(saldototal + 3, 3) = frm_Efectivo.TextBox1.Text
           Hoja10.Cells(saldototal + 3, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
           

            Range(Cells(saldototal + 3, 3), Cells(saldototal + 3, 4)).Select
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
            
           Hoja10.Cells(saldototal + 4, 1) = "CAMBIO:"
           Hoja10.Cells(saldototal + 4, 3) = frm_Efectivo.TextBox2.Text
           Hoja10.Cells(saldototal + 4, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

            Range(Cells(saldototal + 4, 3), Cells(saldototal + 4, 4)).Select
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
            
            Hoja10.Cells(saldototal + 6, 1) = "CAJERO:"
            Hoja10.Cells(saldototal + 6, 2) = UCase(frm_Factura.txt_usuario.Text)
            Hoja10.Cells(saldototal + 7, 1) = "FECHA: " & Format(Date) & "  " & Format(Time)
            Hoja10.Cells(saldototal + 8, 1) = "REFERENCIA: " & UCase(Hoja93.Range("C2").Value + 1)
            
           Hoja10.Cells(saldototal + 11, 1) = "GRACIAS POR PREFERIRNOS"
           Hoja10.Cells(saldototal + 12, 1) = "¡DIOS LE BENDIGA!"

            Range(Cells(saldototal + 11, 1), Cells(saldototal + 11, 4)).Select
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
            
             Range(Cells(saldototal + 11, 1), Cells(saldototal + 12, 1)).Select
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
                Rows(saldototal + 9).Select
                ActiveSheet.Paste

                Hoja10.Select
                Hoja10.Cells(1, 1).Select
                
                    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                IgnorePrintAreas:=False
                
    Hoja10.Visible = xlSheetVeryHidden

End If

End Sub
Private Sub ProcesarFactura()
Dim Fila As Long
Dim Final As Long
Dim Existencia As Long
Dim TotalExistencia As Long
Dim Comprb As Long
Dim nFactura As Long
Dim CostoTotal As Currency
Dim cUpromedio As Currency
Dim xCantidad As Currency
Dim xCodigo As String
Dim xDescrip As String
Dim xCosto As Currency
Dim Serie As String


'Correlativo de la factura de venta
Hoja93.Range("C2").Value = Hoja93.Range("C2").Value + 1
Comprb = Hoja93.Range("C2").Value
Serie = Hoja94.Range("C9").Text

''Envía los datos a la hoja de ventas

If Hoja2.Visible = xlSheetVisible Then

                For i = 0 To frm_Factura.ListBox1.ListCount - 1
                    xCodigo = frm_Factura.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Factura.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Factura.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Factura.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Factura.ListBox1.List(i, 4) 'Importe

                Hoja2.Select
                    Hoja2.Range("A2:N2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja2.Range("A3:N3").Select
                    Selection.Copy
                    Hoja2.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja2.Cells(2, 1) = Hoja2.Cells(3, 1) + 1
                    Hoja2.Cells(2, 2) = CDate(frm_Factura.txtFecha)
                    Hoja2.Cells(2, 4) = frm_Factura.txt_idcliente.Text
                    Hoja2.Cells(2, 5) = frm_Factura.txtCliente.Text
                    Hoja2.Cells(2, 6) = Format(Time)
                    Hoja2.Cells(2, 8) = xCodigo
                    Hoja2.Cells(2, 9) = xDescrip
                    Hoja2.Cells(2, 10) = Serie & " " & Comprb
                    Hoja2.Cells(2, 11) = xCantidad
                    Hoja2.Cells(2, 12) = vPrecioVenta
                    Hoja2.Cells(2, 14) = Hoja92.Range("G1")

                    Final = Final + 1
                Next

ElseIf Hoja2.Visible = xlSheetVeryHidden Then
    Hoja2.Visible = xlSheetVisible

                For i = 0 To frm_Factura.ListBox1.ListCount - 1
                    xCodigo = frm_Factura.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Factura.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Factura.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Factura.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Factura.ListBox1.List(i, 4) 'Importe

                Hoja2.Select
                    Hoja2.Range("A2:N2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja2.Range("A3:N3").Select
                    Selection.Copy
                    Hoja2.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False



                    Application.CutCopyMode = False

                    Hoja2.Cells(2, 1) = Hoja2.Cells(3, 1) + 1
                    Hoja2.Cells(2, 2) = CDate(frm_Factura.txtFecha)
                    Hoja2.Cells(2, 4) = frm_Factura.txt_idcliente.Text
                    Hoja2.Cells(2, 5) = frm_Factura.txtCliente.Text
                    Hoja2.Cells(2, 6) = Format(Time)
                    Hoja2.Cells(2, 8) = xCodigo
                    Hoja2.Cells(2, 9) = xDescrip
                    Hoja2.Cells(2, 10) = Serie & " " & Comprb
                    Hoja2.Cells(2, 11) = xCantidad
                    Hoja2.Cells(2, 12) = vPrecioVenta
                    Hoja2.Cells(2, 14) = Hoja92.Range("G1")

                    Final = Final + 1
                Next

   Hoja2.Visible = xlSheetVeryHidden

End If

'        LImpiarFactura
'        frm_Factura.lbl_nFactura.Caption = "Factura No. " & Hoja93.Range("C2").Value + 1 'Llamamos el número de la factura

End Sub


Private Sub btn_Facturar_Click()
Application.ScreenUpdating = False
    If TextBox1 = "" Then
        MsgBox "Debe registrar el efectivo", vbInformation, "GESTOR DE VENTAS"
        TextBox1.SetFocus
        Exit Sub
    End If
    If TextBox2 < 0 Then
        MsgBox "Debe registrar correctamente el efectivo", vbInformation, "GESTOR DE VENTAS"
        TextBox1 = ""
        TextBox1.SetFocus
        Exit Sub
    End If

    If MsgBox("Son correctos los datos?" + Chr(13) + "Desea procesar la factura?", vbYesNo, "Gestor de Ventas") = vbNo Then
        Exit Sub
    Else

        Hoja2.Unprotect ""
        Facturación
        xTemporal
        zHistorico
        SalidaInventario
        EstadoServicio
        EstadoPedido
'    Application.EnableEvents = False
'        Ticket
'    Application.EnableEvents = True
         
        ProcesarFactura
        MsgBox "Factura procesada con éxito!!!", , "Gestor de Venta"
        Unload Me
        Unload frm_Factura
    Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True
    
        frm_Factura.Reimprimir.Visible = True
        frm_Factura.btn_grabar.Visible = False
        frm_Factura.Show
End If
        Hoja2.Protect ""
        
Application.ScreenUpdating = True

End Sub

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub Frame_Efectivo()
    Frame1.BackColor = &HA99D36
    Frame1.SpecialEffect = fmSpecialEffectSunken
    Me.Label14.Caption = "RECIBO DE EFECTIVO"
        Frame2.BackColor = &HFFFFFF
    Frame2.SpecialEffect = fmSpecialEffectFlat
            Me.Label15.Visible = True
    Me.Label16.Visible = True
    Me.Label17.Visible = True
    Me.TextBox1.Visible = True
    Me.TextBox2.Visible = True
    Me.TextBox3.Visible = True
        Me.btn_tarjeta.Visible = False
    Me.txt_Tarjeta.Visible = False
    Me.lbl_Tarjeta.Visible = False
    Me.btn_Facturar.Visible = True
    Me.lbl_Referencia.Visible = False
    Me.txt_referencia.Visible = False
    Me.txt_referencia.Text = Empty
    Me.TextBox1.SetFocus
End Sub
Private Sub Frame1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Frame_Efectivo
End Sub
Private Sub Lbl1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Frame_Efectivo
End Sub
Private Sub cmd_materiales_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Frame_Efectivo
End Sub
Private Sub Frame_Tarjeta()
    Frame2.BackColor = &HA99D36
    Frame2.SpecialEffect = fmSpecialEffectSunken
    Me.Label14.Caption = "TARJETA DE DEBITO"
    Frame1.BackColor = &HFFFFFF
    Frame1.SpecialEffect = fmSpecialEffectFlat
    
    Me.Label14.Caption = "TARJETA DE DEBITO"
        Me.Label15.Visible = False
    Me.Label16.Visible = False
    Me.Label17.Visible = False
    Me.TextBox1.Visible = False
    Me.TextBox2.Visible = False
    Me.TextBox3.Visible = False
    Me.TextBox1.Value = ""
    Me.btn_tarjeta.Visible = True
    Me.btn_Facturar.Visible = False
    
    Me.txt_Tarjeta.Visible = True
    Me.lbl_Tarjeta.Visible = True
    Me.txt_Tarjeta = frm_Factura.txtTotal.Value
    Me.lbl_Referencia.Visible = True
    Me.txt_referencia.Visible = True
    Me.txt_referencia.SetFocus
End Sub

Private Sub Frame2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Frame_Tarjeta
End Sub
Private Sub Lbl2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Frame_Tarjeta
End Sub

Private Sub cmd_productos_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Frame_Tarjeta
End Sub
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

KeyAscii = ValidarDecimales(TextBox1, KeyAscii)

End Sub
Private Sub TextBox1_Change()
Dim antes As Currency
Dim ahora As Currency
Dim saldo As Currency
Me.TextBox1.BackColor = &H80000005

antes = frm_Factura.txtTotal.Value
If Me.TextBox1 = "" Then
    ahora = 0
Else

ahora = Me.TextBox1

End If


saldo = ahora - antes

Me.TextBox2 = saldo

If InStr(TextBox2, ",") > 0 Then
nuevo = Replace(TextBox2.Value, ",", ".")
TextBox2.Value = nuevo
End If

If Me.TextBox1 = "" Then
    Me.TextBox2 = ""
End If

End Sub

Private Sub UserForm_Initialize()

    Me.TextBox3 = frm_Factura.txtTotal.Value
    Me.lbl_nFactura.Caption = "Factura No. " & Hoja93.Range("C2").Value + 1
    Frame1.BackColor = &HA99D36
    Frame1.SpecialEffect = fmSpecialEffectSunken
    Me.Label14.Caption = "RECIBO DE EFECTIVO"

End Sub

Private Sub EstadoServicio()

Dim X As String
Dim encontrado As Boolean
Dim Titulo As String
Dim Estado As String

Application.ScreenUpdating = False
Titulo = "Gestor de Servicios"
Estado = "INACTIVO"

If frm_Factura.txt_nservicio = "" Then
    Exit Sub
End If

X = frm_Factura.txt_nservicio.Text
       
If Hoja32.Visible = xlSheetVisible Then

    Hoja32.Select
    Range("C1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like X Then
            encontrado = True
            Exit Do
                                 
        End If
    Loop
    If encontrado = False Then
        MsgBox "Pedido no registrado, informar a usuario administrativo", vbInformation, Titulo
        Exit Sub
    End If
    
       
           ActiveCell.Offset(0, 6) = Estado

ElseIf Hoja32.Visible = xlSheetVeryHidden Then
    Hoja32.Visible = xlSheetVisible
    
    Hoja32.Select
    Range("C1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like X Then
            encontrado = True
            Exit Do
                                 
        End If
    Loop
    If encontrado = False Then
        MsgBox "Pedido no registrado, informar a usuario administrativo", vbInformation, Titulo
        Exit Sub
    End If
    
       
           ActiveCell.Offset(0, 6) = Estado
        
    Hoja32.Visible = xlSheetVeryHidden

End If

           
           
Application.ScreenUpdating = True
End Sub

Private Sub EstadoPedido()

Dim X As String
Dim encontrado As Boolean
Dim Titulo As String
Dim Estado As String

Application.ScreenUpdating = False
Titulo = "Gestor de Pedidos"
Estado = "INACTIVO"

If frm_Factura.txt_nPedido.Text = "" Then
    Exit Sub
End If

X = frm_Factura.txt_nPedido.Text
       
If Hoja31.Visible = xlSheetVisible Then

    Hoja31.Select
    Range("C1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like X Then
            encontrado = True
            Exit Do
                                 
        End If
    Loop
    If encontrado = False Then
        MsgBox "Pedido no registrado, informar a usuario administrativo", vbInformation, Titulo
        Exit Sub
    End If
    
       
           ActiveCell.Offset(0, 8) = Estado

ElseIf Hoja31.Visible = xlSheetVeryHidden Then
    Hoja31.Visible = xlSheetVisible
    
    Hoja31.Select
    Range("C1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like X Then
            encontrado = True
            Exit Do
                                 
        End If
    Loop
    If encontrado = False Then
        MsgBox "Pedido no registrado, informar a usuario administrativo", vbInformation, Titulo
        Exit Sub
    End If
    
       
           ActiveCell.Offset(0, 8) = Estado
        
    Hoja31.Visible = xlSheetVeryHidden

End If

           
           
Application.ScreenUpdating = True
End Sub





