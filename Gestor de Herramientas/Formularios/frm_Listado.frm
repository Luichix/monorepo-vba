VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Listado 
   Caption         =   "Listado de Herramientas"
   ClientHeight    =   7908
   ClientLeft      =   15020
   ClientTop       =   3300
   ClientWidth     =   13950
   OleObjectBlob   =   "frm_Listado.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frm_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_item_Click()
frm_registro.Show
Unload Me
frm_Listado.Show
End Sub

Private Sub cmdAceptar_Click()
      Call InsertarHerramienta
End Sub

Private Sub lbx_Herramienta_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call InsertarHerramienta
End Sub
 
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub TextBox1_Change()
On Error Resume Next
uf = Hoja1.Range("A" & Rows.Count).End(xlUp).Row

If TextBox1 = "" Then
    Me.lbx_herramienta.RowSource = "tbl_Herramienta"
    Exit Sub
End If

Hoja1.AutoFilterMode = False
Me.lbx_herramienta = Clear
Me.lbx_herramienta.RowSource = Clear

For Fila = 2 To uf
    STRG = Hoja1.Cells(Fila, 2).Value 'Variable para descripción
    Codigo = Hoja1.Cells(Fila, 3).Value 'Variable para codigo
    
    If UCase(STRG) Like "*" & UCase(TextBox1.Value) & "*" Then
        Me.lbx_herramienta.AddItem
        Me.lbx_herramienta.List(x, 0) = Hoja1.Cells(Fila, 1).Value
        Me.lbx_herramienta.List(x, 1) = Hoja1.Cells(Fila, 2).Value
        Me.lbx_herramienta.List(x, 2) = Hoja1.Cells(Fila, 3).Value
       
        x = x + 1
   '----------------------------------------------------------------------------------
    'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
    ElseIf Codigo Like "*" & UCase(TextBox1.Value) & "*" Then
        Me.lbx_herramienta.AddItem
        Me.lbx_herramienta.List(x, 0) = Hoja1.Cells(Fila, 1).Value
        Me.lbx_herramienta.List(x, 1) = Hoja1.Cells(Fila, 2).Value
        Me.lbx_herramienta.List(x, 2) = Hoja1.Cells(Fila, 3).Value
       
        x = x + 1
    
    End If
    '----------------------------------------------------------------------------------
Next
Me.lbx_herramienta.ColumnCount = 4
Me.lbx_herramienta.ColumnWidths = "45 pt;130 pt;800 pt;0 pt"

End Sub
Private Sub UserForm_Initialize()

Me.lbx_herramienta.ColumnCount = 4
Me.lbx_herramienta.ColumnWidths = "45 pt;130 pt;800 pt;0 pt"
Me.lbx_herramienta.RowSource = "tbl_Herramienta"

Me.TextBox1.SetFocus

End Sub


