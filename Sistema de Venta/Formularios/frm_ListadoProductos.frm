VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ListadoProductos 
   Caption         =   "GESTOR DE VENTAS"
   ClientHeight    =   9480.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5060
   OleObjectBlob   =   "frm_ListadoProductos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ListadoProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_InsertarItem_Click()
    Call InsertarCuentadesdeListBox
End Sub

Private Sub btn_Salir_Click()
    Unload Me
End Sub

Private Sub lbx_Cuentas_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call InsertarCuentadesdeListBox
End Sub

Private Sub TextBox1_Change()
On Error Resume Next
uf = Hoja1.Range("A" & Rows.Count).End(xlUp).Row

If TextBox1 = "" Then
    Me.lbx_Cuentas.RowSource = "Código_Venta"
    Exit Sub
End If


Hoja1.AutoFilterMode = False
Me.lbx_Cuentas = Clear
Me.lbx_Cuentas.RowSource = Clear

For Fila = 2 To uf
    strg = Hoja1.Cells(Fila, 2).Value 'Variable para descripción
    Codigo = Hoja1.Cells(Fila, 1).Value 'Variable para codigo
    
    If UCase(strg) Like "*" & UCase(TextBox1.Value) & "*" Then
        Me.lbx_Cuentas.AddItem
        Me.lbx_Cuentas.List(X, 0) = Hoja1.Cells(Fila, 1).Value
        Me.lbx_Cuentas.List(X, 1) = Hoja1.Cells(Fila, 2).Value
        X = X + 1
   '----------------------------------------------------------------------------------
    'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
    ElseIf Codigo Like "*" & UCase(TextBox1.Value) & "*" Then
        Me.lbx_Cuentas.AddItem
        Me.lbx_Cuentas.List(X, 0) = Hoja1.Cells(Fila, 1).Value
        Me.lbx_Cuentas.List(X, 1) = Hoja1.Cells(Fila, 2).Value
        X = X + 1
    
    End If
    '----------------------------------------------------------------------------------
Next

Me.lbx_Cuentas.ColumnWidths = "45 pt;150 pt"

End Sub
Private Sub UserForm_Initialize()

Me.lbx_Cuentas.ColumnCount = 2
Me.lbx_Cuentas.ColumnWidths = "45 pt;150 pt"
Me.lbx_Cuentas.RowSource = "Código_Venta"

Me.TextBox1.SetFocus

End Sub

