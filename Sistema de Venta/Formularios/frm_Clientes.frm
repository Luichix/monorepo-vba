VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Clientes 
   Caption         =   "GESTOR DE VENTAS"
   ClientHeight    =   6450
   ClientLeft      =   14930
   ClientTop       =   3390
   ClientWidth     =   5470
   OleObjectBlob   =   "frm_Clientes.frx":0000
End
Attribute VB_Name = "frm_Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_Nuevo_Click()
    
    If frm_Clientes.cmdCerrar.Caption = "Salir" Then
    frm_RegistrarClientes.Caption = "GESTOR DE DEVOLUCIÓNES"
    frm_RegistrarClientes.BackColor = &H404000
    End If
    frm_RegistrarClientes.Show
    
    
    If frm_Clientes.Caption = "GESTOR DE DEVOLUCIÓNES" Then
     Unload Me
    frm_Clientes.BackColor = &H404000
    frm_Clientes.Caption = "GESTOR DE DEVOLUCIÓNES"
    frm_Clientes.cmdCerrar.Caption = "Salir"
    Else
    Unload Me
    End If
    
    frm_Clientes.Show
    
    
End Sub

Private Sub cmdAceptar_Click()
      Call InsertarClientes
End Sub


Private Sub lbx_clientes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call InsertarClientes
End Sub
 
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub TextBox1_Change()
On Error Resume Next
uf = Hoja7.Range("A" & Rows.Count).End(xlUp).Row

If TextBox1 = "" Then
    Me.lbx_clientes.RowSource = "ID_Clientes"
    Exit Sub
End If

Hoja7.AutoFilterMode = False
Me.lbx_clientes = Clear
Me.lbx_clientes.RowSource = Clear

For Fila = 2 To uf
    strg = Hoja7.Cells(Fila, 2).Value 'Variable para descripción
    Codigo = Hoja7.Cells(Fila, 1).Value 'Variable para codigo
    
    If UCase(strg) Like "*" & UCase(TextBox1.Value) & "*" Then
        Me.lbx_clientes.AddItem
        Me.lbx_clientes.List(X, 0) = Hoja7.Cells(Fila, 1).Value
        Me.lbx_clientes.List(X, 1) = Hoja7.Cells(Fila, 2).Value
        X = X + 1
   '----------------------------------------------------------------------------------
    'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
    ElseIf Codigo Like "*" & UCase(TextBox1.Value) & "*" Then
        Me.lbx_clientes.AddItem
        Me.lbx_clientes.List(X, 0) = Hoja7.Cells(Fila, 1).Value
        Me.lbx_clientes.List(X, 1) = Hoja7.Cells(Fila, 2).Value
        X = X + 1
    
    End If
    '----------------------------------------------------------------------------------
Next

Me.lbx_clientes.ColumnWidths = "45 pt;150 pt"

End Sub
Private Sub UserForm_Initialize()

Me.lbx_clientes.ColumnCount = 2
Me.lbx_clientes.ColumnWidths = "45 pt;150 pt"
Me.lbx_clientes.RowSource = "ID_Clientes"

Me.TextBox1.SetFocus
     

End Sub

