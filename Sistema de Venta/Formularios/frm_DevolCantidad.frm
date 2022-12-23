VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_DevolCantidad 
   Caption         =   "GESTOR DE VENTAS"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5310
   OleObjectBlob   =   "frm_DevolCantidad.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_DevolCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_Agregar_Click()
    frm_DevolProducto.txtCantidad.Text = frm_DevolCantidad.TextBox1.Text
    Unload Me
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

KeyAscii = ValidarDecimales(TextBox1, KeyAscii)

End Sub

Private Sub btn_Salir_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.TextBox2.Text = frm_DevolProducto.txt_medida.Text
    
End Sub
