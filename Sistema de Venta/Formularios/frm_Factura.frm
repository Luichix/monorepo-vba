VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Factura 
   Caption         =   "GESTOR DE VENTAS"
   ClientHeight    =   10095
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   11800
   OleObjectBlob   =   "frm_Factura.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Factura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long
Dim vPrecioVenta As Currency
Dim vImporte As Currency
Dim CostoUnitario As Currency
        
Private Sub btn_grabar_Click()
'On Error GoTo Salir

Application.ScreenUpdating = False
    If Me.txtTotal.Text = Empty Then
            MsgBox "No se ha registrado ningun producto", , "Gestor de Ventas"
            Exit Sub
    
    End If
    
    frm_Pedido_Select.Show
        
    If Hoja29.Visible = xlSheetVisible Then
        Hoja29.Select
        Hoja29.Cells(1, 1).Select
    End If
     Application.ScreenUpdating = True
                    
     
'Salir:
' If Err <> 0 Then
'    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
' End If

End Sub

Private Sub btn_Limpiar_Click()
    Me.ListBox1.Clear
    Me.txt_idcliente.Text = "0"
    Me.txtCliente.Text = "CLIENTE EVENTUAL"
    Me.lbl_npedido.Visible = False
    Me.lbl_BuscarCliente.Visible = True
    Me.txt_nPedido = Empty
    Me.txt_nservicio = Empty
    Me.txt_Abono = Empty
    Me.txt_FechaEntrega = Empty
    Me.txt_observacion = Empty
    sumarImporte
    Me.btn_Limpiar.Visible = False
    Me.btn_grabar.Visible = True
    Me.lbl_abono.Visible = False
    Me.txt_Abono.Visible = False
    Me.lbl_iva.Visible = True
    Me.txtIVA.Visible = True
    Me.lbl_nMesa.Visible = False
    Me.btn_Procesar.SetFocus
    
End Sub
Private Sub btn_Limpiar2_Click()
    Me.ListBox1.Clear
    Me.txt_idcliente.Text = "0"
    Me.txtCliente.Text = "CLIENTE EVENTUAL"
    Me.lbl_npedido.Visible = False
    Me.lbl_BuscarCliente.Visible = True
    Me.txt_nPedido = Empty
    Me.txt_nservicio = Empty
    Me.txt_FechaEntrega = Empty
        Me.txt_observacion = Empty
    Me.txt_Abono = Empty
    sumarImporte
    Me.btn_Limpiar.Visible = False
    Me.btn_grabar.Visible = True
    Me.btn_Limpiar2.Visible = False
    Me.btn_pedidos.Visible = True
        Me.lbl_abono.Visible = False
    Me.txt_Abono.Visible = False
    Me.lbl_iva.Visible = True
    Me.txtIVA.Visible = True
    Me.lbl_nMesa.Visible = False
    Me.btn_Procesar.SetFocus
       
End Sub
Private Sub btn_Limpiar3_Click()
    LimpiarFactura
End Sub
Public Sub LimpiarFactura()
    Me.ListBox1.Clear
    Me.txt_idcliente.Text = "0"
    Me.txtCliente.Text = "CLIENTE EVENTUAL"
    Me.lbl_npedido.Visible = False
    Me.lbl_BuscarCliente.Visible = True
    Me.txt_nPedido = Empty
    Me.txt_Abono = Empty
    Me.txt_FechaEntrega = Empty
        Me.txt_observacion = Empty
    Me.txt_nservicio = Empty
    sumarImporte
    Me.btn_Limpiar.Visible = False
    Me.btn_grabar.Visible = True
    Me.btn_Limpiar2.Visible = False
    Me.btn_pedidos.Visible = True
        Me.btn_Limpiar3.Visible = False
    Me.btn_servicio.Visible = True
        Me.lbl_abono.Visible = False
    Me.txt_Abono.Visible = False
    Me.lbl_iva.Visible = True
    Me.txtIVA.Visible = True
    Me.lbl_nMesa.Visible = False
    Me.btn_Procesar.SetFocus
End Sub

Private Sub btn_pedidos_Click()
    Me.Reimprimir.Visible = False
    Me.btn_grabar.Visible = True
        If Me.btn_Limpiar.Visible = True Then
        Me.btn_grabar.Visible = False
    End If
    frm_ListaPedido.Show

    Me.btn_Procesar.SetFocus
    
    
    
End Sub

Private Sub btn_Procesar_Click()
On Error GoTo Salir

Application.ScreenUpdating = False
    If Me.txtTotal.Text = Empty Then
            MsgBox "No se ha registrado ninguna venta", , "Gestor de Ventas"
            Exit Sub
    ElseIf Me.txtTotal.Text < 0 Then
            MsgBox "Reportar a un usuario administrativo,realizar la devolución en el modulo de devoluciones", vbInformation, "Gestor de Ventas"
            Exit Sub
    
    End If

        frm_Efectivo.Show
        
    If Hoja2.Visible = xlSheetVisible Then
        Hoja2.Select
        Hoja2.Cells(1, 1).Select
    End If
     Application.ScreenUpdating = True
                    
     
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Ventas"
 End If

End Sub
Private Sub CommandButton1_Click()
Unload Me
End Sub
Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.ListBox1.ListIndex = -1 ' Eliminar la "barra de selección"
End Sub


Private Sub Label22_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label22.SpecialEffect = fmSpecialEffectSunken
End Sub


Private Sub Label22_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label22.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Label21_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label21.SpecialEffect = fmSpecialEffectSunken
End Sub


Private Sub Label21_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Label21.SpecialEffect = fmSpecialEffectFlat
End Sub

Private Sub btn_servicio_Click()
     Me.Reimprimir.Visible = False
    Me.btn_grabar.Visible = True
          If Me.btn_Limpiar.Visible = True Then
        Me.btn_grabar.Visible = False
    End If
  frm_ListaServicio.Show

    Me.btn_Procesar.SetFocus
End Sub

Private Sub lbl_BuscarCliente_Click()
frm_Factura.btn_grabar.Visible = True
frm_Factura.Reimprimir.Visible = False
    banderaClientes = 1
    Call LanzarListadoClientes(Me, "lbl_LanzarListadoClientes")
End Sub
Private Sub lbl_BuscarCliente_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub



Private Sub lblEliminarItem_Click()
On Error Resume Next
Me.EliminarItem
Me.ctrls_FormatoMoneda
frm_Factura.btn_grabar.Visible = True
frm_Factura.Reimprimir.Visible = False
frm_Factura.btn_Limpiar.Visible = False

If Me.txt_nPedido <> "" Then
    Me.btn_Limpiar2.Visible = True
    Me.btn_pedidos.Visible = False
End If
If Me.lbl_nMesa.Visible = True Then
    Me.btn_Limpiar3.Visible = True
    Me.btn_servicio.Visible = False
End If
frm_Factura.btn_Procesar.SetFocus

End Sub
Private Sub lblEliminarItem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub lblProductos_Click()

frm_ProductoVenta.Show
frm_Factura.btn_grabar.Visible = True
frm_Factura.Reimprimir.Visible = False
frm_Factura.btn_Limpiar.Visible = False

If Me.txt_nPedido <> "" Then
    Me.btn_Limpiar2.Visible = True
    Me.btn_pedidos.Visible = False
End If
If Me.lbl_nMesa.Visible = True Then
    Me.btn_Limpiar3.Visible = True
    Me.btn_servicio.Visible = False
End If
frm_Factura.btn_Procesar.SetFocus

End Sub
Private Sub lblProductos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.ListBox1.ListIndex = -1 ' Eliminar la "barra de selección"
SetCursor LoadCursor(0, IDC_HAND)
End Sub




Private Sub Reimprimir_Click()
Application.ScreenUpdating = False
    Application.EnableEvents = False
     
  
    If Hoja10.Visible = xlSheetVisible Then
        Hoja10.Select
        Hoja10.Cells(1, 1).Select
        
            ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
        
    ElseIf Hoja10.Visible = xlSheetVeryHidden Then
        Hoja10.Visible = xlSheetVisible
                       Hoja10.Select
        Hoja10.Cells(1, 1).Select
        
            ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
        Hoja10.Visible = xlSheetVeryHidden
    End If
      Application.EnableEvents = True
Application.ScreenUpdating = True

                              Final = GetNuevoR(Hoja92)
                                  
                                  Hoja92.Cells(Final, 10) = "Factura No. " & Hoja93.Range("C2")
                              
                                  Hoja92.Cells(Final, 11) = "=NOW()"
                                  Hoja92.Cells(Final, 11).Copy
                                  Hoja92.Cells(Final, 11).PasteSpecial Paste:=xlPasteValues
                                  Application.CutCopyMode = False
                                  
                                  Hoja92.Cells(Final, 12) = Me.txt_usuario.Text
                         
                              Application.EnableEvents = False
                                  ThisWorkbook.Save
                              Application.EnableEvents = True
End Sub

Private Sub txt_idcliente_Change()
Dim Fila As Long
Dim xFila As Long
Dim Final As Long
Dim xFinal As Long
Dim Clase As String

 Fila = 2
    Do While Hoja7.Cells(Fila, 1) <> ""
        Fila = Fila + 1
    Loop

    Final = Fila - 1

    'Solicito la información de la hoja de materiales para que se reflejen en los controles
    For Fila = 2 To Final
        If txt_idcliente.Text = Hoja7.Cells(Fila, 1) Then
            Me.txt_Ruc = Hoja7.Cells(Fila, 3)
            Exit For
        End If
    Next
    
End Sub


Private Sub UserForm_Activate()
Me.txtFecha = Date
                                            'Me.txtHora = Format(Time)
Me.txt_usuario = Hoja92.Range("G1")

End Sub
Private Sub UserForm_Initialize()
                                            'Actualizar = True
                                            'Reloj = Format(Time)
                                            'Application.OnTime Now + TimeValue("00:00:01"), "Hora"
                                            '..... más código

Me.lbl_nFactura.Caption = "Factura No. " & Hoja93.Range("C2").Value + 1 'Llamamos el número de la factura
With ListBox1
    .ColumnCount = 5
    .ColumnWidths = "70 pt;85 pt;215 pt;100 pt;50 pt" ' Unidades de medida, 72 pt(puntos)=1 Pulgada
End With
End Sub

'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Actualizar = False
''..... más código
'End Sub

Public Sub sumarImporte()
'Suma la columna de los importes
Dim xImporte As Currency
Dim IvaPorcentaje As Byte
Dim sTotal As Currency
Dim xAbono As Currency
Dim xIVA As Currency

vSeparadorMiles = Hoja94.Range("C5")

If vSeparadorMiles = "." Then
    xDecimal = ","
        ElseIf vSeparadorMiles = "," Then
            xDecimal = "."
End If

sTotal = 0
        For i = 0 To Me.ListBox1.ListCount - 1
            Me.ListBox1.List(i, 4) = _
            Replace(Me.ListBox1.List(i, 4), Application.ThousandsSeparator, "")  'Aquí elimino el separador de miles en el Importe
            Me.ListBox1.List(i, 4) = _
            Replace(Me.ListBox1.List(i, 4), ",", ".") 'Ahora sustituyo la coma decimal por el punto decimal, para poder hacer la sumatoria con la variable sTotal, ya que con la coma decimal, no se suman los decimales
            
            sTotal = sTotal + Val(Me.ListBox1.List(i, 4)) 'Aquí hago la sumatoria del importe, utilizando el punto decimal

            Me.ListBox1.List(i, 4) = _
            Replace(Me.ListBox1.List(i, 4), ".", Application.DecimalSeparator)  'Aquí devuelvo el formato decimal para que no afecte al ListBox
            Me.ListBox1.List(i, 4) = FormatNumber(Me.ListBox1.List(i, 4), 2) 'Aqui doy formato de moneda para que aparezcan los separadores de miles y decimales
        Next

Me.txtSubtotal.Text = sTotal
IvaPorcentaje = Hoja94.Range("C6")

If Me.txt_Abono = Empty Then
    xAbono = 0
Else
    xAbono = Me.txt_Abono.Text
End If


            If sTotal > 0 Then ' aqui se hacen los calculos para el subtotal, iva y total
                    Me.txtIVA.Text = (sTotal / 100) * IvaPorcentaje
                    xIVA = Me.txtIVA.Text

                    Me.txtTotal.Text = sTotal + xIVA - xAbono
                    Me.txtLetras.Text = UCase(cMoneda(Me.txtTotal.Text))
            
                Else
                    Me.txtSubtotal.Text = Empty
                    Me.txtIVA.Text = Empty
                    Me.txtTotal.Text = Empty
                    Me.txtLetras.Text = Empty
            End If

End Sub



Public Sub AgregarProducto()
On Error Resume Next
Titulo = "Gestor de Ventas"

        If frm_ProductoVenta.ComboBox1.Text = "" Then
            frm_ProductoVenta.ComboBox1.BackColor = &HC0C0FF
            MsgBox "Ingrese un Código de Producto", , Titulo
            frm_ProductoVenta.ComboBox1.SetFocus
            Exit Sub
        End If

         If frm_ProductoVenta.txtCantidad.Text = "" Then
            frm_ProductoVenta.txtCantidad.BackColor = &HC0C0FF
            MsgBox "Ingrese una Cantidad", , Titulo
            frm_ProductoVenta.txtCantidad.SetFocus
            Exit Sub
        End If
             
             
        With frm_Factura
            .ListBox1.AddItem
            .ListBox1.List(i, 0) = frm_ProductoVenta.ComboBox1.Text
            .ListBox1.List(i, 1) = Val(frm_ProductoVenta.txtCantidad.Text) 'Código del producto
            .ListBox1.List(i, 1) = Replace(Me.ListBox1.List(i, 1), ",", ".")
            .ListBox1.List(i, 2) = frm_ProductoVenta.txt_Nombre.Text 'Nombre del producto
            .ListBox1.List(i, 3) = frm_ProductoVenta.txt_precioV.Text 'Precio Venta
            .ListBox1.List(i, 4) = frm_ProductoVenta.txtImporte.Text
            .ListBox1.List(i, 5) = frm_ProductoVenta.txt_categoria

            i = i + 1
        End With

        sumarImporte

        With frm_ProductoVenta
            .ComboBox1.ListIndex = -1
            .txt_Nombre = ""
            .txtCantidad = ""
            .txt_precioV = ""
            .txt_categoria = ""

        End With
        
        frm_ProductoVenta.txt_busqueda.SetFocus

End Sub
Public Sub EliminarItem()

' Elimina el item seleccionado y resta el importe de la columna de importes
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "Seleccionar un producto para eliminar", vbInformation
        Exit Sub
    End If

Me.ListBox1.RemoveItem (ListBox1.ListIndex)
Me.ListBox1.ListIndex = -1 ' Eliminar la "barra de selección"
Me.sumarImporte
            
End Sub
Public Sub ctrls_FormatoMoneda()
On Error Resume Next
    Me.txtSubtotal.Text = FormatNumber(Me.txtSubtotal.Text, 2)
    Me.txt_Abono.Text = FormatNumber(Me.txt_Abono.Text, 2)
    Me.txtTotal.Text = FormatNumber(Me.txtTotal.Text, 2)
End Sub
Private Sub btn_Procesar_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = ALT + 43 Then
    frm_Factura.btn_grabar.Visible = True
    frm_Factura.Reimprimir.Visible = False
    frm_Factura.btn_Limpiar.Visible = False
    frm_ProductoVenta.Show
    If Me.txt_nPedido <> "" Then
    Me.btn_Limpiar2.Visible = True
    Me.btn_pedidos.Visible = False
    If Me.lbl_nMesa.Visible = True Then
    Me.btn_Limpiar3.Visible = True
    Me.btn_servicio.Visible = False
    frm_Factura.btn_Procesar.SetFocus
End If

End If
    End If
    If KeyAscii = ALT + 42 Then
        frm_Factura.btn_grabar.Visible = True
    frm_Factura.Reimprimir.Visible = False
    frm_Factura.btn_Limpiar.Visible = False
    frm_ProductoAFacturar.Show
    If Me.txt_nPedido <> "" Then
    Me.btn_Limpiar2.Visible = True
    Me.btn_pedidos.Visible = False
    If Me.lbl_nMesa.Visible = True Then
    Me.btn_Limpiar3.Visible = True
    Me.btn_servicio.Visible = False
    frm_Factura.btn_Procesar.SetFocus
End If

End If
    End If
    
End Sub
Private Sub CommandButton1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = ALT + 43 Then
frm_Factura.btn_grabar.Visible = True
frm_Factura.Reimprimir.Visible = False
frm_Factura.btn_Limpiar.Visible = False
    frm_ProductoVenta.Show
    If Me.txt_nPedido <> "" Then
    Me.btn_Limpiar2.Visible = True
    Me.btn_pedidos.Visible = False
    If Me.lbl_nMesa.Visible = True Then
    Me.btn_Limpiar3.Visible = True
    Me.btn_servicio.Visible = False
    frm_Factura.btn_Procesar.SetFocus
End If

End If

    End If
        If KeyAscii = ALT + 42 Then
frm_Factura.btn_grabar.Visible = True
frm_Factura.Reimprimir.Visible = False
frm_Factura.btn_Limpiar.Visible = False
    frm_ProductoAFacturar.Show
    If Me.txt_nPedido <> "" Then
    Me.btn_Limpiar2.Visible = True
    Me.btn_pedidos.Visible = False
    If Me.lbl_nMesa.Visible = True Then
    Me.btn_Limpiar3.Visible = True
    Me.btn_servicio.Visible = False
    frm_Factura.btn_Procesar.SetFocus
End If

End If

    End If
End Sub

Private Sub LimpiarControles()
Me.ListBox1.Clear
Me.txt_idcliente = 0
Me.txtCliente = "CLIENTE EVENTUAL"

End Sub

Public Sub AgregarItems()
On Error Resume Next
Titulo = "Gestor de Ventas"

        If frm_ProductoAFacturar.ComboBox1.Text = "" Then
            frm_ProductoAFacturar.ComboBox1.BackColor = &HC0C0FF
            MsgBox "Ingrese un Código de Producto", , Titulo
            frm_ProductoAFacturar.ComboBox1.SetFocus
            Exit Sub
        End If

         If frm_ProductoAFacturar.txtCantidad.Text = "" Then
            frm_ProductoAFacturar.txtCantidad.BackColor = &HC0C0FF
            MsgBox "Ingrese una Cantidad", , Titulo
            frm_ProductoAFacturar.txtCantidad.SetFocus
            Exit Sub
        End If

        With frm_Factura
            .ListBox1.AddItem
            .ListBox1.List(i, 0) = frm_ProductoAFacturar.ComboBox1.Text
            .ListBox1.List(i, 1) = Val(frm_ProductoAFacturar.txtCantidad.Text) 'Código del producto
            .ListBox1.List(i, 1) = Replace(Me.ListBox1.List(i, 1), ",", ".")
            .ListBox1.List(i, 2) = frm_ProductoAFacturar.txt_Nombre.Text 'Nombre del producto
            .ListBox1.List(i, 3) = frm_ProductoAFacturar.txt_precioV.Text 'Precio Venta
            .ListBox1.List(i, 4) = frm_ProductoAFacturar.txtImporte.Text
            .ListBox1.List(i, 5) = frm_ProductoAFacturar.txt_categoria.Text

            i = i + 1
        End With

        sumarImporte

        With frm_ProductoAFacturar
            .ComboBox1.ListIndex = -1
            .txt_Nombre = ""
            .txtCantidad = ""
            .txt_precioV = ""
            .txt_categoria = ""

        End With

        frm_ProductoAFacturar.ComboBox1.SetFocus

End Sub
Public Sub CargarFacturacion()
On Error Resume Next

    Me.lbl_npedido.Caption = "Pedido No. " & frm_ListaPedido.txt_nPedido
    Me.txtCliente.Text = frm_ListaPedido.txt_cliente
    Me.txt_idcliente.Text = frm_ListaPedido.txt_idcliente
    Me.txt_nPedido.Text = frm_ListaPedido.txt_nPedido
    Me.txt_Abono.Text = frm_ListaPedido.txt_Abono
    Me.txt_FechaEntrega.Text = frm_ListaPedido.txt_FechaEntrega.Text
    Me.txt_observacion.Text = frm_ListaPedido.txt_Observa.Text

    For i = 0 To frm_ListaPedido.ListBox1.ListCount - 1
        frm_Factura.ListBox1.AddItem frm_ListaPedido.ListBox1.List(i, 0)
         frm_Factura.ListBox1.List(i, 1) = frm_ListaPedido.ListBox1.List(i, 1)
          frm_Factura.ListBox1.List(i, 2) = frm_ListaPedido.ListBox1.List(i, 2)
           frm_Factura.ListBox1.List(i, 3) = frm_ListaPedido.ListBox1.List(i, 3)
            frm_Factura.ListBox1.List(i, 4) = frm_ListaPedido.ListBox1.List(i, 4)
             frm_Factura.ListBox1.List(i, 5) = frm_ListaPedido.ListBox1.List(i, 5)
        
    Next
    
    sumarImporte
        
End Sub

Public Sub CargarFactuServi()
On Error Resume Next
    Me.lbl_nMesa.Caption = frm_ListaServicio.txt_Observa
    Me.txtCliente.Text = frm_ListaServicio.txt_cliente
    Me.txt_idcliente.Text = frm_ListaServicio.txt_idcliente
    Me.txt_nservicio.Text = frm_ListaServicio.txt_nServicios

    For i = 0 To frm_ListaServicio.ListBox1.ListCount - 1
        frm_Factura.ListBox1.AddItem frm_ListaServicio.ListBox1.List(i, 0)
         frm_Factura.ListBox1.List(i, 1) = frm_ListaServicio.ListBox1.List(i, 1)
          frm_Factura.ListBox1.List(i, 2) = frm_ListaServicio.ListBox1.List(i, 2)
           frm_Factura.ListBox1.List(i, 3) = frm_ListaServicio.ListBox1.List(i, 3)
            frm_Factura.ListBox1.List(i, 4) = frm_ListaServicio.ListBox1.List(i, 4)
             frm_Factura.ListBox1.List(i, 5) = frm_ListaServicio.ListBox1.List(i, 5)
        
    Next
    
    sumarImporte
        
End Sub

