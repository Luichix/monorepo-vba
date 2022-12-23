VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Proveedores 
   Caption         =   "Base de datos de Proveedores"
   ClientHeight    =   2520
   ClientLeft      =   15050
   ClientTop       =   3390
   ClientWidth     =   5190
   OleObjectBlob   =   "frm_Proveedores.frx":0000
End
Attribute VB_Name = "frm_Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
On Error GoTo Salir
    Final = GetUltimoR(Hoja23)

With frm_fCompras

    For Fila = 2 To Final
        If ComboBox1.Text = Hoja23.Cells(Fila, 1) Then
            .txtProveedor.Text = Hoja23.Cells(Fila, 1)
            .txtNRF.Text = Hoja23.Cells(Fila, 2)
            .txtTELF.Text = Hoja23.Cells(Fila, 3)
            .txtUBIC.Text = Hoja23.Cells(Fila, 4)
            Exit For
        End If
    Next
    
End With

Unload Me

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor Administrativo"
 End If
 
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long
Dim Registro As Integer

    
Final = GetUltimoR(Hoja23)

    For Fila = 2 To Final
        If ComboBox1.Text = Hoja23.Cells(Fila, 1) Then
            Me.ComboBox1.Text = Hoja23.Cells(Fila, 1)
            Exit For
        
        End If
    Next


End Sub
Private Sub ComboBox1_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox1.ListCount
    ComboBox1.RemoveItem 0
Next Fila

Final = GetUltimoR(Hoja23)
    
    For Fila = 2 To Final
        Lista = Hoja23.Cells(Fila, 1)
        ComboBox1.AddItem (Lista)
    Next
End Sub
