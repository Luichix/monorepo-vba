VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_caja 
   Caption         =   "Cajas de Herramienta"
   ClientHeight    =   7824
   ClientLeft      =   20
   ClientTop       =   300
   ClientWidth     =   17040
   OleObjectBlob   =   "frm_caja.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_caja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_Cargar_Click()
    Call InsertarCaja
End Sub
Private Sub btn_salir_Click()
Unload Me
End Sub

Private Sub lbx_cuenta_Click()
    Call InsertarCaja
End Sub

Private Sub TextBox1_Change()
On Error Resume Next
uf = Hoja2.Range("A" & Rows.Count).End(xlUp).Row

If TextBox1 = "" Then
    Me.lbx_cuenta.RowSource = "Box"
    Exit Sub
End If

Hoja2.AutoFilterMode = False
Me.lbx_cuenta = Clear
Me.lbx_cuenta.RowSource = Clear

For Fila = 2 To uf
    STRG = Hoja2.Cells(Fila, 2).Value 'Variable para descripción
    Codigo = Hoja2.Cells(Fila, 3).Value 'Variable para codigo
    
    If UCase(STRG) Like "*" & UCase(TextBox1.Value) & "*" Then
        Me.lbx_cuenta.AddItem
        Me.lbx_cuenta.List(x, 0) = Hoja2.Cells(Fila, 1).Value
        Me.lbx_cuenta.List(x, 1) = Hoja2.Cells(Fila, 2).Value
        Me.lbx_cuenta.List(x, 2) = Hoja2.Cells(Fila, 3).Value
        Me.lbx_cuenta.List(x, 3) = Hoja2.Cells(Fila, 4).Value
        Me.lbx_cuenta.List(x, 4) = Hoja2.Cells(Fila, 5).Value
        Me.lbx_cuenta.List(x, 5) = Hoja2.Cells(Fila, 6).Value
        Me.lbx_cuenta.List(x, 6) = Hoja2.Cells(Fila, 7).Value
        Me.lbx_cuenta.List(x, 7) = Hoja2.Cells(Fila, 8).Value
        Me.lbx_cuenta.List(x, 8) = Hoja2.Cells(Fila, 9).Value
       
        x = x + 1
   '----------------------------------------------------------------------------------
    'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
    ElseIf Codigo Like "*" & UCase(TextBox1.Value) & "*" Then
        Me.lbx_cuenta.AddItem
        Me.lbx_cuenta.List(x, 0) = Hoja2.Cells(Fila, 1).Value
        Me.lbx_cuenta.List(x, 1) = Hoja2.Cells(Fila, 2).Value
        Me.lbx_cuenta.List(x, 2) = Hoja2.Cells(Fila, 3).Value
        Me.lbx_cuenta.List(x, 3) = Hoja2.Cells(Fila, 4).Value
        Me.lbx_cuenta.List(x, 4) = Hoja2.Cells(Fila, 5).Value
        Me.lbx_cuenta.List(x, 5) = Hoja2.Cells(Fila, 6).Value
        Me.lbx_cuenta.List(x, 6) = Hoja2.Cells(Fila, 7).Value
        Me.lbx_cuenta.List(x, 7) = Hoja2.Cells(Fila, 8).Value
        Me.lbx_cuenta.List(x, 8) = Hoja2.Cells(Fila, 9).Value
       
        x = x + 1
    
    End If
    '----------------------------------------------------------------------------------
Next
    Me.lbx_cuenta.ColumnCount = 9
    Me.lbx_cuenta.ColumnWidths = "30 pt; 100 pt; 50 pt; 250 pt"

End Sub

Private Sub UserForm_Initialize()
On Error Resume Next
    Me.lbx_cuenta.ColumnCount = 9
    Me.lbx_cuenta.ColumnWidths = "30 pt; 200 pt; 50 pt; 250 pt"
    Me.lbx_cuenta.RowSource = "Box"
End Sub

