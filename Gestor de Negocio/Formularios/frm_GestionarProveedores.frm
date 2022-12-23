VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_GestionarProveedores 
   Caption         =   "Proveedores"
   ClientHeight    =   1340
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frm_GestionarProveedores.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_GestionarProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
Unload Me
frm_RegistrarProveedor.Show
End Sub

Private Sub CommandButton2_Click()
Unload Me
frm_EliminarProveedor.Show
End Sub
