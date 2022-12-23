VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ListaServicio 
   Caption         =   "GESTOR DE SERVICIOS"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   12150
   OleObjectBlob   =   "frm_ListaServicio.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ListaServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_FacturarServicio_Click()
     With frm_Factura
        .LimpiarFactura
    End With
   Call InsertarCargarServicios

    If frm_ListaServicio.lbx_Servicios.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Pedido del Cliente", vbInformation
        frm_ListaServicio.lbx_Servicios.SetFocus
        Exit Sub

    End If
    With frm_ListaServicio
        .Cargar_Servicio
    End With

    With frm_Factura
        .ListBox1.Clear
        .CargarFactuServi
        .ctrls_FormatoMoneda
        .btn_Limpiar.Visible = True
        .btn_grabar.Visible = False
        .lbl_nMesa.Visible = True
        .lbl_BuscarCliente.Visible = True

    End With
    Unload Me

End Sub

Private Sub btn_Cargar_Click()
   Call InsertarCargarServicios
   
    If frm_ListaServicio.lbx_Servicios.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Servicio del Cliente", vbInformation
        frm_ListaServicio.lbx_Servicios.SetFocus
        Exit Sub
    End If

    With frm_ListaServicio
        .Cargar_Servicio
    End With
    
Me.Frame2.Visible = False
Me.Frame1.Visible = True
Me.btn_Cargar.Visible = False
Me.btn_regresar.Visible = True

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



Private Sub Frame2_Click()

End Sub

Private Sub UserForm_Initialize()
Dim Fila As Long
Dim Final As Long
Dim Estado As String
On Error Resume Next


banderaCargarServicios = 1


Estado = "ACTIVO"

Me.lbx_Servicios.ColumnCount = 8
Me.lbx_Servicios.ColumnWidths = "0 pt;0 pt;40 pt;70 pt;20 pt;260 pt;70 pt;60 pt"
Me.lbx_Servicios.RowSource = "tbl_atencion"

uf = Hoja32.Range("A" & Rows.Count).End(xlUp).Row

Hoja32.AutoFilterMode = False
Me.lbx_Servicios = Clear
Me.lbx_Servicios.RowSource = Clear

For Fila = 2 To uf
    strg = Hoja32.Cells(Fila, 9).Value 'Variable para descripción

    If UCase(strg) Like Estado Then
        Me.lbx_Servicios.AddItem
        Me.lbx_Servicios.List(X, 0) = Hoja32.Cells(Fila, 1).Value
        Me.lbx_Servicios.List(X, 1) = Hoja32.Cells(Fila, 2).Value
        Me.lbx_Servicios.List(X, 2) = Hoja32.Cells(Fila, 3).Value
        Me.lbx_Servicios.List(X, 3) = Hoja32.Cells(Fila, 4).Value
        Me.lbx_Servicios.List(X, 4) = Hoja32.Cells(Fila, 5).Value
        Me.lbx_Servicios.List(X, 5) = Hoja32.Cells(Fila, 6).Value
        Me.lbx_Servicios.List(X, 6) = Hoja32.Cells(Fila, 7).Value
        Me.lbx_Servicios.List(X, 6) = Replace(Me.lbx_Servicios.List(X, 7), ",", ".")
        Me.lbx_Servicios.List(X, 7) = Hoja32.Cells(Fila, 8).Value
        Me.lbx_Servicios.List(X, 8) = Hoja32.Cells(Fila, 9).Value
     
        X = X + 1
   End If
Next

Me.lbx_Servicios.ColumnCount = 8
Me.lbx_Servicios.ColumnWidths = "0 pt;0 pt;40 pt;70 pt;20 pt;260 pt;70 pt;60 pt"

Me.txt_busqueda.SetFocus
End Sub

Public Sub Cargar_Servicio()
Dim Fila As Long
Dim Final As Long
Dim Estado As String
On Error Resume Next

Estado = txt_nServicios.Text

frm_ListaServicio.ListBox1.ColumnCount = 5
frm_ListaServicio.ListBox1.ColumnWidths = "70 pt;85 pt;215 pt;100 pt;50 pt"
frm_ListaServicio.ListBox1.RowSource = "tbl_encargos"

uf = Hoja30.Range("A" & Rows.Count).End(xlUp).Row

Hoja30.AutoFilterMode = False
frm_ListaServicio.ListBox1 = Clear
frm_ListaServicio.ListBox1.RowSource = Clear

For Fila = 2 To uf
    strg = Hoja30.Cells(Fila, 4).Value 'Variable para descripción

    If UCase(strg) Like Estado Then
        frm_ListaServicio.ListBox1.AddItem
        frm_ListaServicio.ListBox1.List(X, 0) = Hoja30.Cells(Fila, 9).Value
        frm_ListaServicio.ListBox1.List(X, 1) = Hoja30.Cells(Fila, 11).Value
        frm_ListaServicio.ListBox1.List(X, 1) = Replace(frm_ListaServicio.ListBox1.List(X, 1), ",", ".")
        frm_ListaServicio.ListBox1.List(X, 2) = Hoja30.Cells(Fila, 10).Value
        frm_ListaServicio.ListBox1.List(X, 3) = Hoja30.Cells(Fila, 12).Value
        frm_ListaServicio.ListBox1.List(X, 3) = Replace(frm_ListaServicio.ListBox1.List(X, 3), ",", ".")
        frm_ListaServicio.ListBox1.List(X, 4) = Hoja30.Cells(Fila, 13).Value
        frm_ListaServicio.ListBox1.List(X, 4) = Replace(frm_ListaServicio.ListBox1.List(X, 4), ",", ".")
        frm_ListaServicio.ListBox1.List(X, 5) = Hoja30.Cells(Fila, 8).Value

     
        X = X + 1
   End If
Next

frm_ListaServicio.ListBox1.ColumnCount = 5
frm_ListaServicio.ListBox1.ColumnWidths = "70 pt;85 pt;215 pt;100 pt;50 pt"

End Sub


