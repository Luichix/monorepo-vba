VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Regimen 
   Caption         =   "GESTIÓN DE PERSONAL"
   ClientHeight    =   3195
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   5170
   OleObjectBlob   =   "frm_Regimen.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Regimen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btn_Cargar_Click()
    Call InsertarRegimen
      
End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub

Private Sub lbx_cuenta_Click()
    Call InsertarRegimen
End Sub

Private Sub UserForm_Initialize()
On Error Resume Next


Me.lbx_cuenta.ColumnCount = 3
Me.lbx_cuenta.ColumnWidths = "110 pt;70 pt;50 pt"
Me.lbx_cuenta.RowSource = "Tbl_regimen"

End Sub

