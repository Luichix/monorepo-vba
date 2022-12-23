VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Estado 
   Caption         =   "Estado de Inventario"
   ClientHeight    =   3204
   ClientLeft      =   20
   ClientTop       =   300
   ClientWidth     =   4360
   OleObjectBlob   =   "frm_Estado.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Estado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub btn_Cargar_Click()
    Call InsertaEstado
      
End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub

Private Sub lbx_cuenta_Click()
    Call InsertaEstado
End Sub

Private Sub UserForm_Initialize()
On Error Resume Next


Me.lbx_cuenta.ColumnCount = 1
Me.lbx_cuenta.ColumnWidths = "60 pt"
Me.lbx_cuenta.RowSource = "Tbl_Estado"

End Sub

