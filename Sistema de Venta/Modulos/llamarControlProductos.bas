Attribute VB_Name = "llamarControlProductos"
Option Explicit
Public banderaListadoProductos As Long
Public banderaClientes As Long
Public banderaProductos As Long
Public banderaDevoluciones As Long
Public banderaProveedores As Long
Public banderaInsumos As Long
Public banderaProductosCompra As Long
Public banderaModificarPedido As Long
Public banderaCargarPedido As Long
Public banderaCargarFacturacion As Long
Public banderaModificarServicios As Long
Public banderaCargarServicios As Long
Public banderaCargarFactuServicio As Long

Public Function LanzarListadoProductos(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_ListadoProductos
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frmCalendario.StartUpPosition = 0
            frmCalendario.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frmCalendario.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_ListadoProductos.Show

End Function

Sub InsertarCuentadesdeListBox()

If frm_ListadoProductos.lbx_Cuentas.ListIndex = -1 Then
    MsgBox "Debe seleccionar un Código de Producto", vbInformation
    frm_ListadoProductos.lbx_Cuentas.SetFocus
    Exit Sub
End If

Select Case banderaListadoProductos
    Case 1
        With frm_ProductoAFacturar
            .ComboBox1 = frm_ListadoProductos.lbx_Cuentas.Column(0)
            .txt_Nombre = frm_ListadoProductos.lbx_Cuentas.Column(1)
            Unload frm_ListadoProductos
            frm_ProductoAFacturar.txtCantidad.SetFocus
        End With

    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Sub InsertarProductosdeListBox()

If frm_ProductoVenta.lbx_busqueda.ListIndex = -1 Then
    MsgBox "Debe seleccionar un Código de Producto", vbInformation
    frm_ProductoVenta.lbx_busqueda.SetFocus
    Exit Sub
End If

Select Case banderaProductos
    
    Case 1
        With frm_ProductoVenta
            .txt_Nombre = frm_ProductoVenta.lbx_busqueda.Column(1)
        End With

    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Sub InsertarDevolProductodeListBox()

If frm_DevolProducto.lbx_busqueda.ListIndex = -1 Then
    MsgBox "Debe seleccionar un Código de Producto", vbInformation
    frm_DevolProducto.lbx_busqueda.SetFocus
    Exit Sub
End If

Select Case banderaDevoluciones
    
    Case 1
        With frm_DevolProducto
            .txt_Nombre = frm_DevolProducto.lbx_busqueda.Column(1)
        End With
        
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Sub InsertarProductos()

If frm_ProductoAComprar.lbx_producto.ListIndex = -1 Then
    MsgBox "Debe seleccionar un Producto", vbInformation
    frm_ProductoAComprar.lbx_producto.SetFocus
    Exit Sub
End If

Select Case banderaProductosCompra
    
    Case 1
        With frm_ProductoAComprar
            .ComboBox1 = frm_ProductoAComprar.lbx_producto.Column(1)
        End With

    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Public Function LanzarListadoClientes(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_Clientes
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frmCalendario.StartUpPosition = 0
            frmCalendario.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frmCalendario.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_Clientes.Show

End Function
Sub InsertarClientes()

If frm_Clientes.lbx_clientes.ListIndex = -1 Then
    MsgBox "Debe seleccionar un Cliente", vbInformation
    frm_Clientes.lbx_clientes.SetFocus
    Exit Sub
End If

Select Case banderaClientes
    Case 1
        With frm_Factura
            .txtCliente = frm_Clientes.lbx_clientes.Column(1)
            .txt_idcliente = frm_Clientes.lbx_clientes.Column(0)
            .txt_Ruc = frm_Clientes.lbx_clientes.Column(2)
            Unload frm_Clientes
            
        End With
    Case 2
        With frm_Devolucion
            .txtCliente = frm_Clientes.lbx_clientes.Column(1)
            .txt_idcliente = frm_Clientes.lbx_clientes.Column(0)
            .txt_Ruc = frm_Clientes.lbx_clientes.Column(2)
            Unload frm_Clientes
            
        End With
'
'    Case 2
'        With frm_LibroDiario
'            .cbo_CodCuenta = frm_ListadoCuentas.lbx_Cuentas.Column(0)
'            .txt_NombreCuenta = frm_ListadoCuentas.lbx_Cuentas.Column(1)
'            Unload frm_ListadoCuentas
'        End With
'
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
'
'Sub BuscarItemEnListBox()
'Dim i As Long
'
'Select Case banderaProductos
'
'    Case 1
'        For i = 0 To frm_ProductoVenta.lbx_busqueda.ListCount - 1
'            If frm_ProductoVenta.lbx_busqueda.List(i, 0) = frm_ProductoVenta.lbx_busqueda Then
'                frm_ProductoVenta.lbx_busqueda.ListIndex = i
'                Exit For
'            End If
'        Next
'
'    End Select
'End Sub


Public Function LanzarListadoProveedores(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_Proveedores
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frmCalendario.StartUpPosition = 0
            frmCalendario.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frmCalendario.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_Proveedores.Show

End Function
Sub InsertarProveedores()

If frm_Proveedores.lbx_proveedor.ListIndex = -1 Then
    MsgBox "Debe seleccionar un Proveedor", vbInformation
    frm_Proveedores.lbx_proveedor.SetFocus
    Exit Sub
End If

Select Case banderaProveedores
    Case 1
        With frm_fCompras
            .txtProveedor = frm_Proveedores.lbx_proveedor.Column(1)
            .txtNRF = frm_Proveedores.lbx_proveedor.Column(2)
            .txtTELF = frm_Proveedores.lbx_proveedor.Column(3)
            .txtUBIC = frm_Proveedores.lbx_proveedor.Column(4)
            Unload frm_Proveedores
            
        End With
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub

Public Function LanzarListadoInsumos(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_ProductoAComprar
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frmCalendario.StartUpPosition = 0
            frmCalendario.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frmCalendario.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_ProductoAComprar.Show

End Function
Sub InsertarInsumos()

If frm_ProductoAComprar.lbx_insumo.ListIndex = -1 Then
    MsgBox "Debe seleccionar un insumo", vbInformation
    frm_ProductoAComprar.lbx_insumo.SetFocus
    Exit Sub
End If

Select Case banderaInsumos
    Case 1
        With frm_ProductoAComprar
            .ComboBox1 = frm_ProductoAComprar.lbx_insumo.Column(1)
           
        End With
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Sub InsertarModificarPerdido()

Select Case banderaModificarPedido
    Case 1
        With frm_pedidos
            .txt_nPedido = frm_ListaPedido.lbx_Pedidos.Column(2)
            frm_pedidos.lbl_npedido.Caption = "Pedido No. " & frm_pedidos.txt_nPedido.Text
                Unload frm_ListaPedido
            frm_pedidos.txt_producto.SetFocus
        End With

    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Sub InsertarCargarPedido()

If frm_ListaPedido.lbx_Pedidos.ListIndex = -1 Then
   
    frm_ListaPedido.lbx_Pedidos.SetFocus
    Exit Sub
End If

Select Case banderaCargarPedido
    Case 1
        With frm_ListaPedido
            .txt_nPedido = frm_ListaPedido.lbx_Pedidos.Column(2)
            .txt_idcliente = frm_ListaPedido.lbx_Pedidos.Column(3)
            .txt_cliente = frm_ListaPedido.lbx_Pedidos.Column(4)
            .lbl_npedido.Caption = "No. Pedido " & frm_ListaPedido.lbx_Pedidos.Column(2)
            .txt_FechaEntrega.Text = frm_ListaPedido.lbx_Pedidos.Column(5)
            .txt_Abono = frm_ListaPedido.lbx_Pedidos.Column(7)
            
        End With

    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub


Sub InsertarCargarServicios()

If frm_ListaServicio.lbx_Servicios.ListIndex = -1 Then
   
    frm_ListaServicio.lbx_Servicios.SetFocus
    Exit Sub
End If

Select Case banderaCargarServicios
    Case 1
        With frm_ListaServicio
            .txt_nServicios = frm_ListaServicio.lbx_Servicios.Column(2)
            .txt_idcliente = frm_ListaServicio.lbx_Servicios.Column(4)
            .txt_cliente = frm_ListaServicio.lbx_Servicios.Column(5)
            .lbl_nServicio.Caption = "No. Servicio " & frm_ListaServicio.lbx_Servicios.Column(2)
            .txt_Observa = frm_ListaServicio.lbx_Servicios.Column(3)
            
        End With

    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub

