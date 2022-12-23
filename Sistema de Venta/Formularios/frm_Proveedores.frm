VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Proveedores 
   Caption         =   "AGENDA DE PROVEEDORES"
   ClientHeight    =   6525
   ClientLeft      =   15050
   ClientTop       =   3390
   ClientWidth     =   5730
   OleObjectBlob   =   "frm_Proveedores.frx":0000
End
Attribute VB_Name = "frm_Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Nuevo_Click()
frm_RegistrarProveedor.Show
Unload Me
frm_Proveedores.Show
End Sub

Private Sub cmdAceptar_Click()
      Call InsertarProveedores
End Sub


Private Sub lbx_Proveedor_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call InsertarProveedores
End Sub
 
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub TextBox1_Change()
On Error Resume Next
uf = Hoja8.Range("A" & Rows.Count).End(xlUp).Row

If TextBox1 = "" Then
    Me.lbx_proveedor.RowSource = "tbl_Proveedores"
    Exit Sub
End If

Hoja8.AutoFilterMode = False
Me.lbx_proveedor = Clear
Me.lbx_proveedor.RowSource = Clear

For Fila = 2 To uf
    strg = Hoja8.Cells(Fila, 2).Value 'Variable para descripción
    Codigo = Hoja8.Cells(Fila, 1).Value 'Variable para codigo
    
    If UCase(strg) Like "*" & UCase(TextBox1.Value) & "*" Then
        Me.lbx_proveedor.AddItem
        Me.lbx_proveedor.List(X, 0) = Hoja8.Cells(Fila, 1).Value
        Me.lbx_proveedor.List(X, 1) = Hoja8.Cells(Fila, 2).Value
        Me.lbx_proveedor.List(X, 2) = Hoja8.Cells(Fila, 3).Value
        Me.lbx_proveedor.List(X, 3) = Hoja8.Cells(Fila, 4).Value
        Me.lbx_proveedor.List(X, 4) = Hoja8.Cells(Fila, 5).Value
        X = X + 1
   '----------------------------------------------------------------------------------
    'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
    ElseIf Codigo Like "*" & UCase(TextBox1.Value) & "*" Then
        Me.lbx_proveedor.AddItem
        Me.lbx_proveedor.List(X, 0) = Hoja8.Cells(Fila, 1).Value
        Me.lbx_proveedor.List(X, 1) = Hoja8.Cells(Fila, 2).Value
        Me.lbx_proveedor.List(X, 2) = Hoja8.Cells(Fila, 3).Value
        Me.lbx_proveedor.List(X, 3) = Hoja8.Cells(Fila, 4).Value
        Me.lbx_proveedor.List(X, 4) = Hoja8.Cells(Fila, 5).Value
        X = X + 1
    
    End If
    '----------------------------------------------------------------------------------
Next
Me.lbx_proveedor.ColumnCount = 5
Me.lbx_proveedor.ColumnWidths = "45 pt;150 pt;0 pt;0 pt;0 pt"

End Sub
Private Sub UserForm_Initialize()

Me.lbx_proveedor.ColumnCount = 5
Me.lbx_proveedor.ColumnWidths = "45 pt;150 pt;0 pt;0 pt;0 pt"
Me.lbx_proveedor.RowSource = "tbl_Proveedores"

Me.TextBox1.SetFocus

End Sub


