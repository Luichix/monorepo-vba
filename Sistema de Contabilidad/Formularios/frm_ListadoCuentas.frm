VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ListadoCuentas 
   Caption         =   "Catálogo de Cuentas"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5040
   OleObjectBlob   =   "frm_ListadoCuentas.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ListadoCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btn_InsertarItem_Click()
    Call InsertarCuentadesdeListBox
End Sub

Private Sub btn_Salir_Click()
    Unload Me
End Sub

Private Sub lbx_Cuentas_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call InsertarCuentadesdeListBox
End Sub

Private Sub UserForm_Activate()
Dim Fila As Long
Dim Final As Long


    Final = nReg(Hoja2, 2, 1) - 1
   
    With frm_ListadoCuentas
        For Fila = 2 To Final
            .lbx_Cuentas.AddItem Hoja2.Cells(Fila, 1)
            .lbx_Cuentas.List(.lbx_Cuentas.ListCount - 1, 1) = Hoja2.Cells(Fila, 2)
        Next
    End With
    
    Call BuscarItemEnListBox
End Sub

Private Sub UserForm_Initialize()
    
    Call CambiarTamanoListboxCuentas

    Me.lbx_Cuentas.ColumnCount = 2
    Me.lbx_Cuentas.ColumnWidths = "45 pt;150 pt"
End Sub
