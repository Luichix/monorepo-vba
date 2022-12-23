VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_RegistrarClientes 
   Caption         =   "Registro de Clientes"
   ClientHeight    =   2490
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   5620
   OleObjectBlob   =   "frm_RegistrarClientes.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_RegistrarClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    Dim Fila As Long
    Dim Final As Long
    Dim Registro As Integer
    Dim Titulo As String
    Dim xTextBox As Control

On Error GoTo Salir

    Titulo = "Gestor de Inventarios"

        For Each xTextBox In Controls
            If xTextBox.Name Like "txt*" And xTextBox = Empty Then
                MsgBox "Debe completar todos los campos", , Titulo
                xTextBox.SetFocus
                Exit Sub
            End If
        Next
 
    
        'Determina el final del listado de Clientes
        Final = GetNuevoR(Hoja4)
        
        'Validación para impedir Clientes repetidos
        For Fila = 2 To Final
            If Hoja4.Cells(Fila, 1) = UCase(Me.txt_Cliente.Text) Then
                MsgBox ("Cliente ya existe en la Base de Datos"), , Titulo
                LimpiarControles
                Me.txt_Cliente.SetFocus
                Exit Sub
                Exit For
            End If
        Next
        
      If MsgBox("Son correctos los datos?" + Chr(13) + "Desea proceder?", vbOKCancel, Titulo) = vbOK Then
          
                'Envía los datos a la hoja de Clientes
                Hoja4.Cells(Final, 1) = UCase(Me.txt_Cliente.Text)
                Hoja4.Cells(Final, 3) = "'" & Me.txt_NIT.Text
                Hoja4.Cells(Final, 4) = Me.txt_email.Text
                '-----------------------------------------------
                'Limpia los controles
                LimpiarControles
            Else
                Exit Sub
    End If
    

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, Titulo
 End If
 
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub LimpiarControles()
    Dim xTextBox As Control
        
        For Each xTextBox In Controls
            If xTextBox.Name Like "txt*" Then
                xTextBox = Empty
                Me.txt_Cliente.SetFocus
            End If
        Next

End Sub
