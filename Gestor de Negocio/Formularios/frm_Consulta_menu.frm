VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Consulta_menu 
   Caption         =   "Consulta de Movimientos"
   ClientHeight    =   1800
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frm_Consulta_menu.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Consulta_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
frm_ConsultaProducto.Show
End Sub

Private Sub CommandButton2_Click()
frm_ConsultaFecha.Show
End Sub
