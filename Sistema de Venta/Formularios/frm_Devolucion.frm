VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Devolucion 
   Caption         =   "GESTOR DE DEVOLUCIÓNES"
   ClientHeight    =   9960.001
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   11800
   OleObjectBlob   =   "frm_Devolucion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Devolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long
Dim vPrecioVenta As Currency
Dim vImporte As Currency
Dim CostoUnitario As Currency
Private Sub btn_Procesar_Click()
On Error GoTo Salir

Application.ScreenUpdating = False
    If txtCliente = "" Then
        MsgBox "Debe ingresar los datos del cliente.", vbInformation, "GESTOR DE CAJA"
        txtCliente.SetFocus
        Exit Sub
    End If

    If Me.txtTotal.Text = Empty Then
            MsgBox "No se ha registrado ninguna venta", , "Gestor de Ventas"
            Exit Sub
    
    End If

        frm_Devolver.Show
        
    If Hoja25.Visible = xlSheetVisible Then
        Hoja25.Select
        Hoja25.Cells(1, 1).Select
    End If
     Application.ScreenUpdating = True
     
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If
End Sub
Private Sub CommandButton1_Click()
Unload Me
End Sub
Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.ListBox1.ListIndex = -1 ' Eliminar la "barra de selección"
End Sub
Private Sub lbl_BuscarCliente_Click()
     frm_Clientes.Caption = "GESTOR DE DEVOLUCIÓNES"
     frm_Clientes.BackColor = &H404000
     frm_Clientes.cmdCerrar.Caption = "Salir"
     frm_RegistrarClientes.BackColor = &H404000
     banderaClientes = 2
    Call LanzarListadoClientes(Me, "lbl_LanzarListadoClientes")
End Sub
Private Sub lbl_BuscarCliente_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub lblEliminarItem_Click()
Me.EliminarItem
Me.ctrls_FormatoMoneda
End Sub
Private Sub lblEliminarItem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub lblProductos_Click()
frm_DevolProducto.Show
End Sub
Private Sub lblProductos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.ListBox1.ListIndex = -1
SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub UserForm_Activate()
Me.txtFecha = Date
Me.txt_usuario = Hoja92.Range("G1")
End Sub
Private Sub UserForm_Initialize()
Me.lbl_nDevolución.Caption = "Devolución No. " & Hoja93.Range("J2").Value + 1 'Llamamos el número de la factura
With ListBox1
    .ColumnCount = 5
    .ColumnWidths = "70 pt;85 pt;215 pt;100 pt;50 pt" ' Unidades de medida, 72 pt(puntos)=1 Pulgada
End With
End Sub
Public Sub sumarImporte()
Dim xImporte As Currency
Dim IvaPorcentaje As Byte
Dim sTotal As Currency
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

            If sTotal > 0 Then
                    Me.txtIVA.Text = (sTotal / 100) * IvaPorcentaje
                    xIVA = Me.txtIVA.Text
                    Me.txtTotal.Text = sTotal + xIVA
                    Me.txtLetras.Text = UCase(cMoneda(Me.txtTotal.Text))
                Else
                    Me.txtSubtotal.Text = Empty
                    Me.txtIVA.Text = Empty
                    Me.txtTotal.Text = Empty
                    Me.txtLetras.Text = Empty
            End If

End Sub
Public Sub DevolucionAgregar()
Titulo = "Gestor de Devoluciónes"

        If frm_DevolProducto.ComboBox1.Text = "" Then
            frm_DevolProducto.ComboBox1.BackColor = &HC0C0FF
            MsgBox "Ingrese un Código de Producto", , Titulo
            frm_DevolProducto.ComboBox1.SetFocus
            Exit Sub
        End If

         If frm_DevolProducto.txtCantidad.Text = "" Then
            frm_DevolProducto.txtCantidad.BackColor = &HC0C0FF
            MsgBox "Ingrese una Cantidad", , Titulo
            frm_DevolProducto.txtCantidad.SetFocus
            Exit Sub
        End If
             
        With frm_Devolucion
            .ListBox1.AddItem frm_DevolProducto.ComboBox1.Text
            .ListBox1.List(i, 1) = Val(frm_DevolProducto.txtCantidad.Text) 'Código del producto
            .ListBox1.List(i, 1) = Replace(Me.ListBox1.List(i, 1), ",", ".")
            .ListBox1.List(i, 2) = frm_DevolProducto.txt_Nombre.Text 'Nombre del producto
            .ListBox1.List(i, 3) = frm_DevolProducto.txt_precioV.Text 'Precio Venta
            .ListBox1.List(i, 4) = frm_DevolProducto.txtImporte.Text
            .ListBox1.List(i, 5) = frm_DevolProducto.txt_categoria.Text
            

            i = i + 1
        End With

        sumarImporte

        With frm_DevolProducto
            .ComboBox1.ListIndex = -1
            .txt_Nombre = ""
            .txtCantidad = ""
            .txt_precioV = ""
            .txt_categoria = ""

        End With
        
        frm_DevolProducto.txt_busqueda.SetFocus

End Sub
Public Sub EliminarItem()
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "Seleccionar un producto para eliminar", vbInformation
        Exit Sub
    End If
Me.ListBox1.RemoveItem (ListBox1.ListIndex)
Me.ListBox1.ListIndex = -1
Me.sumarImporte
End Sub
Public Sub ctrls_FormatoMoneda()
On Error Resume Next
    Me.txtSubtotal.Text = FormatNumber(Me.txtSubtotal.Text, 2)
    Me.txtTotal.Text = FormatNumber(Me.txtTotal.Text, 2)
End Sub
Private Sub btn_Procesar_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = ALT + 43 Then
    frm_DevolProducto.Show
    End If
End Sub
Private Sub LimpiarControles()
Me.ListBox1.Clear
Me.txt_idcliente = 0
Me.txtCliente = "CLIENTE EVENTUAL"
End Sub
