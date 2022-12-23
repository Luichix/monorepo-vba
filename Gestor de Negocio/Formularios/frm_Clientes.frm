VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Clientes 
   Caption         =   "Base de datos de clientes"
   ClientHeight    =   2520
   ClientLeft      =   14930
   ClientTop       =   3390
   ClientWidth     =   5190
   OleObjectBlob   =   "frm_Clientes.frx":0000
End
Attribute VB_Name = "frm_Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAceptar_Click()
 
On Error GoTo Salir
    Final = GetUltimoR(Hoja4)

With frm_Factura

    For Fila = 2 To Final
        If ComboBox1.Text = Hoja4.Cells(Fila, 1) Then
            .txtCliente.Text = Hoja4.Cells(Fila, 1)
            .txt_idcliente.Text = Hoja4.Cells(Fila, 2)
            .txtNIT.Text = Hoja4.Cells(Fila, 3)
            .txtMail.Text = Hoja4.Cells(Fila, 4)
            Exit For
        End If
    Next
    
End With

Unload Me

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If


End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long
Dim Registro As Integer

    
Final = GetUltimoR(Hoja4)

    For Fila = 2 To Final
        If ComboBox1.Text = Hoja4.Cells(Fila, 1) Then
            Me.ComboBox1.Text = Hoja4.Cells(Fila, 1)
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

Final = GetUltimoR(Hoja4)
    
    For Fila = 2 To Final
        Lista = Hoja4.Cells(Fila, 1)
        ComboBox1.AddItem (Lista)
    Next
End Sub
