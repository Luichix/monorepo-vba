VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Jornada 
   Caption         =   "GESTIÓN DE PERSONAL"
   ClientHeight    =   3705
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   12360
   OleObjectBlob   =   "frm_Jornada.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Jornada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_Cargar_Click()
    Call InsertarJornada
      
End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub

Private Sub lbx_cuenta_Click()
    Call InsertarJornada
End Sub

Private Sub UserForm_Initialize()
On Error Resume Next
Dim Fila As Long
Dim Final As Long


Me.lbx_cuenta.ColumnCount = 12
Me.lbx_cuenta.ColumnWidths = "150 pt;0 pt;0 pt;0 pt;0 pt;0 pt;0 pt;0 pt;0 pt;70 pt;0 pt"
Me.lbx_cuenta.RowSource = "Tbl_Jornada"

End Sub

