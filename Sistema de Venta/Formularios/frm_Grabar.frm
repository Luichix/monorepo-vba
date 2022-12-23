VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Grabar 
   Caption         =   "PEDIDOS"
   ClientHeight    =   8490.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   8330.001
   OleObjectBlob   =   "frm_Grabar.frx":0000
End
Attribute VB_Name = "frm_Grabar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Long
Dim vPrecioVenta As Currency
Dim vImporte As Currency
Dim CostoUnitario As Currency
Private Sub Solicitudes()
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
Dim Estado As String
Dim Adelanto As Currency
Dim Abono As Currency

'Correlativo de la factura de venta

Comprb = Hoja93.Range("K2").Value + 1
Estado = "ACTIVO"
Adelanto = Me.TextBox1.Value
If Me.TextBox5 <> "" Then
Abono = Me.TextBox5.Value
ElseIf Me.TextBox5 = "" Then
Abono = 0
End If



''Envía los datos a la hoja de ventas


If Hoja31.Visible = xlSheetVisible Then

                Hoja31.Select
                    Hoja31.Range("A2:K2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja31.Range("A3:K3").Select
                    Selection.Copy
                    Hoja31.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja31.Cells(2, 1) = CDate(frm_Factura.txtFecha)
                    Hoja31.Cells(2, 2) = Format(Time)
                    Hoja31.Cells(2, 3) = Comprb
                    Hoja31.Cells(2, 4) = frm_Factura.txt_idcliente.Text
                    Hoja31.Cells(2, 5) = frm_Factura.txtCliente.Text
                    Hoja31.Cells(2, 6) = CDate(frm_Grabar.txt_FechaEntrega)
                    Hoja31.Cells(2, 7) = Me.TextBox4.Text
                    Hoja31.Cells(2, 8) = Adelanto + Abono
                    Hoja31.Cells(2, 9) = Me.TextBox2.Text
                    Hoja31.Cells(2, 10) = Hoja92.Range("G1")
                    Hoja31.Cells(2, 11) = Estado
                    Hoja31.Cells(2, 12) = UCase(Me.txt_observacion.Text)


ElseIf Hoja31.Visible = xlSheetVeryHidden Then
    Hoja31.Visible = xlSheetVisible


                Hoja31.Select
                    Hoja31.Range("A2:K2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja31.Range("A3:K3").Select
                    Selection.Copy
                    Hoja31.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja31.Cells(2, 1) = CDate(frm_Factura.txtFecha)
                    Hoja31.Cells(2, 2) = Format(Time)
                    Hoja31.Cells(2, 3) = Comprb
                    Hoja31.Cells(2, 4) = frm_Factura.txt_idcliente.Text
                    Hoja31.Cells(2, 5) = frm_Factura.txtCliente.Text
                    Hoja31.Cells(2, 6) = CDate(frm_Grabar.txt_FechaEntrega)
                    Hoja31.Cells(2, 7) = Me.TextBox4.Text
                    Hoja31.Cells(2, 8) = Adelanto + Abono
                    Hoja31.Cells(2, 9) = Me.TextBox2.Text
                    Hoja31.Cells(2, 10) = Hoja92.Range("G1")
                    Hoja31.Cells(2, 11) = Estado
                    Hoja31.Cells(2, 12) = UCase(Me.txt_observacion.Text)


   Hoja31.Visible = xlSheetVeryHidden
End If

End Sub
Private Sub xTemporal()
Dim Fila As Long
Dim Final As Long
Dim Detalle As String

'Correlativo de la factura de venta

Detalle = "INGRESO POR PEDIDOS"

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
                    Hoja26.Cells(2, 5) = frm_Grabar.lbl_nFactura.Caption
                    Hoja26.Cells(2, 6) = Detalle
                    Hoja26.Cells(2, 7) = frm_Grabar.TextBox1.Text
                    Hoja26.Cells(2, 10) = frm_Grabar.TextBox1.Text
                    Hoja26.Cells(2, 13) = frm_Grabar.TextBox1.Text
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
                    Hoja26.Cells(2, 5) = frm_Grabar.lbl_nFactura.Caption
                    Hoja26.Cells(2, 6) = Detalle
                    Hoja26.Cells(2, 7) = frm_Grabar.TextBox1.Text
                    Hoja26.Cells(2, 10) = frm_Grabar.TextBox1.Text
                    Hoja26.Cells(2, 13) = frm_Grabar.TextBox1.Text
                    Hoja26.Cells(2, 17) = Hoja92.Range("G1")

   Hoja26.Visible = xlSheetVeryHidden
End If

End Sub
Private Sub zHistorico()
Dim Fila As Long
Dim Final As Long
Dim Detalle As String

'Correlativo de la factura de venta

Detalle = "INGRESO POR PEDIDOS"

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
                    Hoja22.Cells(2, 5) = frm_Grabar.lbl_nFactura.Caption
                    Hoja22.Cells(2, 6) = Detalle
                    Hoja22.Cells(2, 7) = frm_Grabar.TextBox1.Text
                    Hoja22.Cells(2, 10) = frm_Grabar.TextBox1.Text
                    Hoja22.Cells(2, 13) = frm_Grabar.TextBox1.Text
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
                    Hoja22.Cells(2, 5) = frm_Grabar.lbl_nFactura.Caption
                    Hoja22.Cells(2, 6) = Detalle
                    Hoja22.Cells(2, 7) = frm_Grabar.TextBox1.Text
                    Hoja22.Cells(2, 10) = frm_Grabar.TextBox1.Text
                    Hoja22.Cells(2, 13) = frm_Grabar.TextBox1.Text
                    Hoja22.Cells(2, 17) = Hoja92.Range("G1")

   Hoja22.Visible = xlSheetVeryHidden
End If

End Sub
Private Sub Recibo()
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

If Hoja13.Visible = xlSheetVisible Then

                    'LIMPIAR HOJA
                    Hoja13.Select
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
    
    Hoja13.Rows("14:1000").Select
    Selection.RowHeight = 15

                    Hoja13.Select
                    Hoja13.Cells(8, 1) = "PEDIDO DE CLIENTE"
                    Hoja13.Cells(9, 1) = "CLIENTE: " & UCase(frm_Factura.txtCliente.Text)
                    Hoja13.Cells(10, 1) = "N° RUC: " & UCase(frm_Factura.txt_Ruc.Text)

 For i = 0 To frm_Factura.ListBox1.ListCount - 1
                    xCodigo = frm_Factura.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Factura.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Factura.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Factura.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Factura.ListBox1.List(i, 4) 'Importe

  
                Hoja13.Select

                    Hoja13.Cells(Final, 1) = xCantidad
                    Hoja13.Cells(Final, 1).NumberFormat = "0.00"
                    Hoja13.Cells(Final, 1).HorizontalAlignment = xlCenter
                    Hoja13.Cells(Final, 2) = xDescrip
                    Hoja13.Cells(Final, 4) = vPrecioVenta
                    Hoja13.Cells(Final, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    Hoja13.Cells(Final + 1, 2) = "SUB TOTAL"
                    Hoja13.Cells(Final + 1, 3) = vImporte
                    Hoja13.Cells(Final + 1, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

                    Final = Final + 2
                Next
                
'Determinar el final para colocar el saldo de Factura


        For FiladelTotal = 14 To 1000
            If Hoja13.Cells(FiladelTotal, 2) = "" Then
                saldototal = FiladelTotal
                Exit For
            End If
        Next
        
            Hoja13.Cells(saldototal, 1).Select
            
            
            Hoja13.Cells(saldototal + 2, 1) = "RECIBO DE EFECTIVO"
            
            Range(Cells(saldototal + 2, 1), Cells(saldototal + 2, 4)).Select
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
                     
             Range(Cells(saldototal + 2, 1), Cells(saldototal + 2, 4)).Select
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
            
            Hoja13.Cells(saldototal + 3, 1) = "TOTAL FACTURA:"
            Hoja13.Cells(saldototal + 3, 3) = frm_Factura.txtSubtotal.Text
            Hoja13.Cells(saldototal + 3, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            
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
            
           Hoja13.Cells(saldototal + 4, 1) = "ABONO REALIZADO:"
           Hoja13.Cells(saldototal + 4, 3) = frm_Grabar.TextBox1.Text
           Hoja13.Cells(saldototal + 4, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
           

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
            
           Hoja13.Cells(saldototal + 5, 1) = "PAGO PENDIENTE:"
           Hoja13.Cells(saldototal + 5, 3) = frm_Grabar.TextBox2.Text
           Hoja13.Cells(saldototal + 5, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

            Range(Cells(saldototal + 5, 3), Cells(saldototal + 5, 4)).Select
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
            
            Hoja13.Cells(saldototal + 6, 1) = "FECHA DE ENTREGA: " & frm_Grabar.txt_FechaEntrega.Text
           
            
            Hoja13.Cells(saldototal + 9, 1) = "CAJERO:"
            Hoja13.Cells(saldototal + 9, 2) = UCase(frm_Factura.txt_usuario.Text)
            Hoja13.Cells(saldototal + 10, 1) = "FECHA: " & Format(Date) & "  " & Format(Time)
            Hoja13.Cells(saldototal + 11, 1) = "REFERENCIA: " & UCase(Hoja93.Range("K2").Value + 1)
            
            Hoja13.Cells(saldototal + 14, 1) = "FIRMA DEL VENDEDOR:"
                Range(Cells(saldototal + 14, 3), Cells(saldototal + 14, 4)).Select
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                Selection.Borders(xlEdgeLeft).LineStyle = xlNone
                Selection.Borders(xlEdgeTop).LineStyle = xlNone
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                Selection.Borders(xlEdgeRight).LineStyle = xlNone
                Selection.Borders(xlInsideVertical).LineStyle = xlNone
                Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
            
           Hoja13.Cells(saldototal + 17, 1) = "GRACIAS POR PREFERIRNOS"
           Hoja13.Cells(saldototal + 18, 1) = "¡DIOS LE BENDIGA!"

            Range(Cells(saldototal + 17, 1), Cells(saldototal + 17, 4)).Select
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
            
             Range(Cells(saldototal + 17, 1), Cells(saldototal + 18, 4)).Select
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
            
           Range(Cells(saldototal + 18, 1), Cells(saldototal + 18, 4)).Select
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
            
            Range(Cells(saldototal + 3, 1), Cells(saldototal + 14, 1)).Select

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
               
                Hoja13.Rows("11:12").Select
                Selection.Copy
                Rows(saldototal).Select
                ActiveSheet.Paste
                Rows(saldototal + 7).Select
                ActiveSheet.Paste
                Rows(saldototal + 15).Select
                ActiveSheet.Paste

                Hoja13.Select
                Hoja13.Cells(1, 1).Select
                
                    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                IgnorePrintAreas:=False

ElseIf Hoja13.Visible = xlSheetVeryHidden Then
        Hoja13.Visible = xlSheetVisible
        
        
                    'LIMPIAR HOJA
                    Hoja13.Select
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
    
    Hoja13.Rows("14:1000").Select
    Selection.RowHeight = 15

                    Hoja13.Select
                    Hoja13.Cells(8, 1) = "PEDIDO DE CLIENTE"
                    Hoja13.Cells(9, 1) = "CLIENTE: " & UCase(frm_Factura.txtCliente.Text)
                    Hoja13.Cells(10, 1) = "N° RUC: " & UCase(frm_Factura.txt_Ruc.Text)

 For i = 0 To frm_Factura.ListBox1.ListCount - 1
                    xCodigo = frm_Factura.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Factura.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Factura.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Factura.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Factura.ListBox1.List(i, 4) 'Importe

  
                Hoja13.Select

                    Hoja13.Cells(Final, 1) = xCantidad
                    Hoja13.Cells(Final, 1).NumberFormat = "0.00"
                    Hoja13.Cells(Final, 1).HorizontalAlignment = xlCenter
                    Hoja13.Cells(Final, 2) = xDescrip
                    Hoja13.Cells(Final, 4) = vPrecioVenta
                    Hoja13.Cells(Final, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    Hoja13.Cells(Final + 1, 2) = "SUB TOTAL"
                    Hoja13.Cells(Final + 1, 3) = vImporte
                    Hoja13.Cells(Final + 1, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

                    Final = Final + 2
                Next
                
'Determinar el final para colocar el saldo de Factura


        For FiladelTotal = 14 To 1000
            If Hoja13.Cells(FiladelTotal, 2) = "" Then
                saldototal = FiladelTotal
                Exit For
            End If
        Next
        
            Hoja13.Cells(saldototal, 1).Select
            
            
            Hoja13.Cells(saldototal + 2, 1) = "RECIBO DE EFECTIVO"
            
            Range(Cells(saldototal + 2, 1), Cells(saldototal + 2, 4)).Select
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
                     
             Range(Cells(saldototal + 2, 1), Cells(saldototal + 2, 4)).Select
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
            
            Hoja13.Cells(saldototal + 3, 1) = "TOTAL FACTURA:"
            Hoja13.Cells(saldototal + 3, 3) = frm_Factura.txtSubtotal.Text
            Hoja13.Cells(saldototal + 3, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            
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
            
           Hoja13.Cells(saldototal + 4, 1) = "ABONO REALIZADO:"
           Hoja13.Cells(saldototal + 4, 3) = frm_Grabar.TextBox1.Text
           Hoja13.Cells(saldototal + 4, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
           

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
            
           Hoja13.Cells(saldototal + 5, 1) = "PAGO PENDIENTE:"
           Hoja13.Cells(saldototal + 5, 3) = frm_Grabar.TextBox2.Text
           Hoja13.Cells(saldototal + 5, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

            Range(Cells(saldototal + 5, 3), Cells(saldototal + 5, 4)).Select
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
            
            Hoja13.Cells(saldototal + 6, 1) = "FECHA DE ENTREGA: " & frm_Grabar.txt_FechaEntrega.Text
           
            
            Hoja13.Cells(saldototal + 9, 1) = "CAJERO:"
            Hoja13.Cells(saldototal + 9, 2) = UCase(frm_Factura.txt_usuario.Text)
            Hoja13.Cells(saldototal + 10, 1) = "FECHA: " & Format(Date) & "  " & Format(Time)
            Hoja13.Cells(saldototal + 11, 1) = "REFERENCIA: " & UCase(Hoja93.Range("K2").Value + 1)
            
            Hoja13.Cells(saldototal + 14, 1) = "FIRMA DEL VENDEDOR:"
                Range(Cells(saldototal + 14, 3), Cells(saldototal + 14, 4)).Select
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                Selection.Borders(xlEdgeLeft).LineStyle = xlNone
                Selection.Borders(xlEdgeTop).LineStyle = xlNone
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                Selection.Borders(xlEdgeRight).LineStyle = xlNone
                Selection.Borders(xlInsideVertical).LineStyle = xlNone
                Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
            
           Hoja13.Cells(saldototal + 17, 1) = "GRACIAS POR PREFERIRNOS"
           Hoja13.Cells(saldototal + 18, 1) = "¡DIOS LE BENDIGA!"

            Range(Cells(saldototal + 17, 1), Cells(saldototal + 17, 4)).Select
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
            
             Range(Cells(saldototal + 17, 1), Cells(saldototal + 18, 4)).Select
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
            
           Range(Cells(saldototal + 18, 1), Cells(saldototal + 18, 4)).Select
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
            
            Range(Cells(saldototal + 3, 1), Cells(saldototal + 14, 1)).Select

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
               
                Hoja13.Rows("11:12").Select
                Selection.Copy
                Rows(saldototal).Select
                ActiveSheet.Paste
                Rows(saldototal + 7).Select
                ActiveSheet.Paste
                Rows(saldototal + 15).Select
                ActiveSheet.Paste

                Hoja13.Select
                Hoja13.Cells(1, 1).Select
                
                    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                IgnorePrintAreas:=False
                
    Hoja13.Visible = xlSheetVeryHidden

End If

End Sub
Private Sub GrabarPedido()
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

'Correlativo de la factura de venta
Hoja93.Range("K2").Value = Hoja93.Range("K2").Value + 1
Comprb = Hoja93.Range("K2").Value

''Envía los datos a la hoja de ventas

If Hoja29.Visible = xlSheetVisible Then

                For i = 0 To frm_Factura.ListBox1.ListCount - 1
                    xCodigo = frm_Factura.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Factura.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Factura.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Factura.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Factura.ListBox1.List(i, 4) 'Importe

                Hoja29.Select
                    Hoja29.Range("A2:N2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja29.Range("A3:N3").Select
                    Selection.Copy
                    Hoja29.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja29.Cells(2, 1) = Hoja29.Cells(3, 1) + 1
                    Hoja29.Cells(2, 2) = CDate(frm_Factura.txtFecha)
                    Hoja29.Cells(2, 3) = Format(Time)
                    Hoja29.Cells(2, 4) = Comprb
                    Hoja29.Cells(2, 5) = frm_Factura.txt_idcliente.Text
                    Hoja29.Cells(2, 6) = frm_Factura.txtCliente.Text
                    Hoja29.Cells(2, 8) = xCodigo
                    Hoja29.Cells(2, 9) = xDescrip
                    Hoja29.Cells(2, 10) = xCantidad
                    Hoja29.Cells(2, 11) = vPrecioVenta
                    Hoja29.Cells(2, 13) = Hoja92.Range("G1")

                    Final = Final + 1
                Next

ElseIf Hoja29.Visible = xlSheetVeryHidden Then
    Hoja29.Visible = xlSheetVisible

                For i = 0 To frm_Factura.ListBox1.ListCount - 1
                    xCodigo = frm_Factura.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Factura.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Factura.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Factura.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Factura.ListBox1.List(i, 4) 'Importe

                Hoja29.Select
                    Hoja29.Range("A2:N2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja29.Range("A3:N3").Select
                    Selection.Copy
                    Hoja29.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False



                    Application.CutCopyMode = False

                    Hoja29.Cells(2, 1) = Hoja29.Cells(3, 1) + 1
                    Hoja29.Cells(2, 2) = CDate(frm_Factura.txtFecha)
                    Hoja29.Cells(2, 3) = Format(Time)
                    Hoja29.Cells(2, 4) = Comprb
                    Hoja29.Cells(2, 5) = frm_Factura.txt_idcliente.Text
                    Hoja29.Cells(2, 6) = frm_Factura.txtCliente.Text
                    Hoja29.Cells(2, 8) = xCodigo
                    Hoja29.Cells(2, 9) = xDescrip
                    Hoja29.Cells(2, 10) = xCantidad
                    Hoja29.Cells(2, 11) = vPrecioVenta
                    Hoja29.Cells(2, 13) = Hoja92.Range("G1")

                    Final = Final + 1
                Next

   Hoja29.Visible = xlSheetVeryHidden

End If

End Sub


Private Sub btn_Facturar_Click()
Application.ScreenUpdating = False
    If Me.txt_FechaEntrega = "" Then
        MsgBox "Debe registrar la fecha de entrega del pedido", vbInformation, "GESTOR DE VENTAS"
        TextBox1.SetFocus
        Exit Sub
    End If
    If Me.txt_observacion = "" Then
        MsgBox "Agrege la observación del pedido solicitado", vbInformation, "GESTOR DE VENTAS"
        txt_observacion.SetFocus
        Exit Sub
    End If
    If TextBox1 = "" Then
        MsgBox "Debe registrar el efectivo", vbInformation, "GESTOR DE VENTAS"
        TextBox1.SetFocus
        Exit Sub
    End If
       

    If MsgBox("Son correctos los datos?" + Chr(13) + "Desea procesar la factura?", vbYesNo, "Gestor de Ventas") = vbNo Then
        Exit Sub
    Else

        Hoja29.Unprotect ""
        Solicitudes
        xTemporal
        zHistorico
    Application.EnableEvents = False
        Recibo
        Recibo
    Application.EnableEvents = True
        EstadoPedido
        GrabarPedido
        MsgBox "Pedido grabado con éxito!!!", , "Gestor de Pedidos"

        Unload Me
        Unload frm_Factura
    Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True

        frm_Factura.Reimprimir.Visible = True
        frm_Factura.btn_grabar.Visible = False
        frm_Factura.Show
End If
        Hoja29.Protect ""
        

Application.ScreenUpdating = True
End Sub
'
Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub CommandButton2_Click()
banderaCalendario = 3
    Call LanzarCalendario(Me, "lbl_fecha")
    
    Me.TextBox1.SetFocus
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


saldo = antes - ahora

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
EliminarTitulo Me.Caption
    Me.Height = Me.Height - 20
    Me.TextBox4 = frm_Factura.txtSubtotal.Value
    Me.TextBox5 = frm_Factura.txt_Abono.Value
    Me.TextBox3 = frm_Factura.txtTotal.Value
    Me.lbl_nFactura.Caption = "Pedido No. " & Hoja93.Range("K2").Value + 1
    Me.txt_pedido = frm_Factura.txt_nPedido.Text
    If frm_Factura.txt_nPedido <> Empty Then
        Me.txt_FechaEntrega.Text = frm_Factura.txt_FechaEntrega.Text
        Me.txt_observacion.Text = frm_Factura.txt_observacion.Text
    End If
End Sub
Private Sub EstadoPedido()

Dim X As String
Dim encontrado As Boolean
Dim Titulo As String
Dim Estado As String

Application.ScreenUpdating = False
Titulo = "Gestor de Pedidos"
Estado = "INACTIVO"

If Me.txt_pedido = "" Then
    Exit Sub
End If

X = Me.txt_pedido.Text
       
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


