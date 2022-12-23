VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Pago 
   Caption         =   "GESTIÓN DE PERSONAL"
   ClientHeight    =   2130
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   2830
   OleObjectBlob   =   "frm_Pago.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub btn_Cargar_Click()
    Call InsertarPago
      
End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub

Private Sub Label37_Click()

End Sub

Private Sub lbx_cuenta_Click()
    Call InsertarPago
End Sub

Private Sub UserForm_Initialize()
On Error Resume Next


Me.lbx_cuenta.ColumnCount = 1
Me.lbx_cuenta.ColumnWidths = "60 pt"
Me.lbx_cuenta.RowSource = "Tbl_Cheque"

End Sub

