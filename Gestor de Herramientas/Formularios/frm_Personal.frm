VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Personal 
   Caption         =   "Colaboradores"
   ClientHeight    =   5964
   ClientLeft      =   15000
   ClientTop       =   3200
   ClientWidth     =   10690
   OleObjectBlob   =   "frm_Personal.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frm_personal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
      Call InsertarPersonal
End Sub


Private Sub lbx_Personal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call InsertarPersonal
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub


Private Sub TextBox1_Change()
On Error Resume Next
uf = Hoja4.Range("A" & Rows.Count).End(xlUp).Row

If TextBox1 = "" Then
    Me.lbx_personal.RowSource = "tbl_personal"
    Exit Sub
End If

Hoja4.AutoFilterMode = False
Me.lbx_personal = Clear
Me.lbx_personal.RowSource = Clear

For Fila = 2 To uf
    STRG = Hoja4.Cells(Fila, 2).Value 'Variable para descripción
    Codigo = Hoja4.Cells(Fila, 3).Value 'Variable para codigo
    
    If UCase(STRG) Like "*" & UCase(TextBox1.Value) & "*" Then
        Me.lbx_personal.AddItem
        Me.lbx_personal.List(x, 0) = Hoja4.Cells(Fila, 1).Value
        Me.lbx_personal.List(x, 1) = Hoja4.Cells(Fila, 2).Value
        Me.lbx_personal.List(x, 2) = Hoja4.Cells(Fila, 3).Value
        Me.lbx_personal.List(x, 3) = Hoja4.Cells(Fila, 5).Value
        x = x + 1
   '----------------------------------------------------------------------------------
    'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
    ElseIf Codigo Like "*" & UCase(TextBox1.Value) & "*" Then
        Me.lbx_personal.AddItem
        Me.lbx_personal.List(x, 0) = Hoja4.Cells(Fila, 1).Value
        Me.lbx_personal.List(x, 1) = Hoja4.Cells(Fila, 2).Value
        Me.lbx_personal.List(x, 2) = Hoja4.Cells(Fila, 3).Value
        Me.lbx_personal.List(x, 3) = Hoja4.Cells(Fila, 5).Value
        x = x + 1
    
    End If
    '----------------------------------------------------------------------------------
Next
Me.lbx_personal.ColumnCount = 4
Me.lbx_personal.ColumnWidths = "45 pt;70 pt;250 pt;100 pt"

End Sub
Private Sub UserForm_Initialize()

Me.lbx_personal.ColumnCount = 5
Me.lbx_personal.ColumnWidths = "45 pt;70 pt;250 pt;0 pt;100 pt"
Me.lbx_personal.RowSource = "tbl_personal"

Me.TextBox1.SetFocus

End Sub


