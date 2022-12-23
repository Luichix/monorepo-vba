VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ListaPedido 
   Caption         =   "GESTOR DE PEDIDOS"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   13940
   OleObjectBlob   =   "frm_ListaPedido.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ListaPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btn_Facturar_Click()
    With frm_Factura
        .LimpiarFactura
    End With
   Call InsertarCargarPedido
   
    If frm_ListaPedido.lbx_Pedidos.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Pedido del Cliente", vbInformation
        frm_ListaPedido.lbx_Pedidos.SetFocus
        Exit Sub
        
    End If
    With frm_ListaPedido
        .Cargar_Pedido
    End With
    
    With frm_Factura
        .ListBox1.Clear
        .CargarFacturacion
        .ctrls_FormatoMoneda
        .btn_Limpiar.Visible = True
        .lbl_npedido.Visible = True
        .lbl_BuscarCliente.Visible = False
        .lbl_abono.Visible = True
        .txt_Abono.Visible = True
        .txtIVA.Visible = False
        .lbl_iva.Visible = False
        .btn_grabar.Visible = False
        
    End With
    Unload Me
End Sub

Private Sub btn_Cargar_Click()
   Call InsertarCargarPedido
   
    If frm_ListaPedido.lbx_Pedidos.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Pedido del Cliente", vbInformation
        frm_ListaPedido.lbx_Pedidos.SetFocus
        Exit Sub
    End If

    With frm_ListaPedido
        .Cargar_Pedido
    End With
    
Me.Frame2.Visible = False
Me.Frame1.Visible = True
Me.btn_Cargar.Visible = False
Me.btn_regresar.Visible = True

End Sub

Private Sub btn_Nuevo_Click()
    frm_pedidos.Show
    frm_pedidos.lbl_npedido.Caption = "Pedido No. " & Hoja93.Range("K2").Value + 1
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub


Private Sub btn_regresar_Click()
Me.btn_Cargar.Visible = True

Me.Frame2.Visible = True
Me.Frame1.Visible = False

Me.btn_regresar.Visible = False

End Sub

Private Sub btn_Salir_Click()
Unload Me
End Sub



Private Sub btn_Salir2_Click()
Unload Me
End Sub




Private Sub Frame1_Click()

End Sub

Private Sub txt_nPedido_Change()
Dim Fila As Long
Dim Final As Long

If Me.txt_nPedido.Text = "" Then
    Me.txt_Observa = ""
End If


Fila = 2

    Do While Hoja31.Cells(Fila, 3) <> ""
        Fila = Fila + 1
    Loop

    Final = Fila - 1



    'Solicito la información de la hoja de materiales para que se reflejen en los controles
    For Fila = 2 To Final
        If Me.txt_nPedido.Text = Hoja31.Cells(Fila, 3) Then
            Me.txt_Observa = Hoja31.Cells(Fila, 12)
            Exit For
        End If
    Next
  
End Sub

Private Sub UserForm_Initialize()
Dim Fila As Long
Dim Final As Long
Dim Estado As String
On Error Resume Next

banderaModificarPedido = 1
banderaCargarPedido = 1


Estado = "ACTIVO"

Me.lbx_Pedidos.ColumnCount = 10
Me.lbx_Pedidos.ColumnWidths = "0 pt;0 pt;40 pt;20 pt;190 pt;90 pt;80 pt"
Me.lbx_Pedidos.RowSource = "tbl_Solicitudes"

uf = Hoja31.Range("A" & Rows.Count).End(xlUp).Row

Hoja31.AutoFilterMode = False
Me.lbx_Pedidos = Clear
Me.lbx_Pedidos.RowSource = Clear

For Fila = 2 To uf
    strg = Hoja31.Cells(Fila, 11).Value 'Variable para descripción

    If UCase(strg) Like Estado Then
        Me.lbx_Pedidos.AddItem
        Me.lbx_Pedidos.List(X, 0) = Hoja31.Cells(Fila, 1).Value
        Me.lbx_Pedidos.List(X, 1) = Hoja31.Cells(Fila, 2).Value
        Me.lbx_Pedidos.List(X, 2) = Hoja31.Cells(Fila, 3).Value
        Me.lbx_Pedidos.List(X, 3) = Hoja31.Cells(Fila, 4).Value
        Me.lbx_Pedidos.List(X, 4) = Hoja31.Cells(Fila, 5).Value
        Me.lbx_Pedidos.List(X, 5) = Hoja31.Cells(Fila, 6).Value
        Me.lbx_Pedidos.List(X, 6) = Hoja31.Cells(Fila, 7).Value
        Me.lbx_Pedidos.List(X, 6) = Replace(Me.lbx_Pedidos.List(X, 6), ",", ".")
        Me.lbx_Pedidos.List(X, 7) = Hoja31.Cells(Fila, 8).Value
        Me.lbx_Pedidos.List(X, 7) = Replace(Me.lbx_Pedidos.List(X, 7), ",", ".")
        Me.lbx_Pedidos.List(X, 8) = Hoja31.Cells(Fila, 9).Value
        Me.lbx_Pedidos.List(X, 8) = Replace(Me.lbx_Pedidos.List(X, 8), ",", ".")
        Me.lbx_Pedidos.List(X, 9) = Hoja31.Cells(Fila, 10).Value
        
        X = X + 1
        
   End If
Next

Me.lbx_Pedidos.ColumnCount = 10
Me.lbx_Pedidos.ColumnWidths = "0 pt;0 pt;40 pt;20 pt;190 pt;110 pt;80 pt"

Me.txt_busqueda.SetFocus
End Sub

Public Sub Cargar_Pedido()
Dim Fila As Long
Dim Final As Long
Dim Estado As String
On Error Resume Next

Estado = txt_nPedido.Text

frm_ListaPedido.ListBox1.ColumnCount = 5
frm_ListaPedido.ListBox1.ColumnWidths = "70 pt;85 pt;215 pt;100 pt;50 pt"
frm_ListaPedido.ListBox1.RowSource = "tbl_encargos"

uf = Hoja29.Range("A" & Rows.Count).End(xlUp).Row

Hoja29.AutoFilterMode = False
frm_ListaPedido.ListBox1 = Clear
frm_ListaPedido.ListBox1.RowSource = Clear

For Fila = 2 To uf
    strg = Hoja29.Cells(Fila, 4).Value 'Variable para descripción

    If UCase(strg) Like Estado Then
        frm_ListaPedido.ListBox1.AddItem
        frm_ListaPedido.ListBox1.List(X, 0) = Hoja29.Cells(Fila, 8).Value
        frm_ListaPedido.ListBox1.List(X, 1) = Hoja29.Cells(Fila, 10).Value
        frm_ListaPedido.ListBox1.List(X, 1) = Replace(frm_ListaPedido.ListBox1.List(X, 1), ",", ".")
        frm_ListaPedido.ListBox1.List(X, 2) = Hoja29.Cells(Fila, 9).Value
        frm_ListaPedido.ListBox1.List(X, 3) = Hoja29.Cells(Fila, 11).Value
        frm_ListaPedido.ListBox1.List(X, 3) = Replace(frm_ListaPedido.ListBox1.List(X, 3), ",", ".")
        frm_ListaPedido.ListBox1.List(X, 4) = Hoja29.Cells(Fila, 12).Value
        frm_ListaPedido.ListBox1.List(X, 4) = Replace(frm_ListaPedido.ListBox1.List(X, 4), ",", ".")
        frm_ListaPedido.ListBox1.List(X, 5) = Hoja29.Cells(Fila, 7).Value

     
        X = X + 1
   End If
Next

frm_ListaPedido.ListBox1.ColumnCount = 5
frm_ListaPedido.ListBox1.ColumnWidths = "70 pt;85 pt;215 pt;100 pt;50 pt"

'Me.txt_busqueda.SetFocus
End Sub
