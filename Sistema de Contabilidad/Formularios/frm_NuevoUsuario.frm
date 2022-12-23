VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_NuevoUsuario 
   Caption         =   "Registro de Usuarios"
   ClientHeight    =   6465
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   7170
   OleObjectBlob   =   "frm_NuevoUsuario.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_NuevoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmd_Registrar_Click()
    Dim Fila As Long
    Dim Final As Long
    Dim Registro As Integer
    Dim Seguridad As String

On Error GoTo Salir

Application.ScreenUpdating = False


Hoja9.Select

Final = GetNuevoR(Hoja9)

        For Registro = 2 To Final
            If Hoja9.Cells(Registro, 1) = Me.txt_nUser.Text Then
                Me.txt_nUser.BackColor = &H8080FF
                MsgBox ("El usuario ya existe" + Chr(13) + "Ingrese un usuario diferente")
                Me.txt_nUser.SetFocus
                Exit Sub
                Exit For
            End If
        Next

      If Me.txt_pass1.Text = Me.txt_pass2.Text Then
                
                Me.txt_nUser.BackColor = &HFFFFFF
                Hoja9.Cells(Final, 1) = Me.txt_nUser.Text
                Hoja9.Cells(Final, 2) = Me.txt_pass1.Text
                        
                 MsgBox "Espere un Momento, Click para continuar...!", vbInformation, "Configuración"

        'VALORES PARA HOJAS Y BOTONES
            'GRUPO ADMINISTRATIVO
            Application.Cursor = xlWait
                'Hojas

                Hoja9.Cells(Final, 3).Value = False
                Hoja9.Cells(Final, 4).Value = True
                Hoja9.Cells(Final, 5).Value = True
                
                Me.txt_nUser.Text = ""
                Me.txt_pass1.Text = ""
                Me.txt_pass2.Text = ""

                Me.txt_nUser.SetFocus
                
                Hoja9.Protect (Seguridad)
                
                
    Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True
    
Application.Cursor = xlDefault
            
                 MsgBox "Usuario registrado satisfactoriamente", vbInformation, "Configuración"
                 
                 Unload Me
            Else
                MsgBox "Las contraseñas deben coincidir..!"
                Me.txt_pass1 = Empty
                Me.txt_pass2 = Empty
                Me.txt_pass1.SetFocus
            
    End If

Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Usuarios"
 End If

End Sub


Private Sub UserForm_Initialize()
EliminarTitulo Me.Caption
    Me.Height = Me.Height - 20
End Sub

Private Sub cmd_Finalizar_Click()
    Unload Me
End Sub
