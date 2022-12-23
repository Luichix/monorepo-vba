VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_NuevoUsuario 
   Caption         =   "Registro de Usuarios"
   ClientHeight    =   5580
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

Seguridad = Hoja83.Range("L1").Text



Hoja82.Select

Final = GetNuevoR(Hoja82)

        For Registro = 2 To Final
            If Hoja82.Cells(Registro, 1) = Me.txt_nUser.Text Then
                Me.txt_nUser.BackColor = &H8080FF
                MsgBox ("El usuario ya existe" + Chr(13) + "Ingrese un usuario diferente")
                Me.txt_nUser.SetFocus
                Exit Sub
                Exit For
            End If
        Next

      If Me.txt_pass1.Text = Me.txt_pass2.Text Then
                Hoja82.Unprotect (Seguridad)
                
                Me.txt_nUser.BackColor = &HFFFFFF
                Hoja82.Cells(Final, 1) = Me.txt_nUser.Text
                Hoja82.Cells(Final, 2) = Me.txt_pass1.Text
                        If Me.OptionButton1.Value = True Then
                            Hoja82.Cells(Final, 3) = "USUARIO"
                                Else
                            Hoja82.Cells(Final, 3) = "ADMINISTRADOR"
                        End If
                 MsgBox "Espere un Momento, Click para continuar...!", vbInformation, "Configuración"

        'VALORES PARA HOJAS Y BOTONES
            'GRUPO ADMINISTRATIVO
            Application.Cursor = xlWait
                'Hojas
                If Me.OptionButton1.Value = True Then
                Hoja82.Cells(Final, 4).Value = False
                Hoja82.Cells(Final, 5).Value = False
                Hoja82.Cells(Final, 6).Value = False
                Hoja82.Cells(Final, 7).Value = False
                Hoja82.Cells(Final, 8).Value = False
                Hoja82.Cells(Final, 9).Value = False
                Hoja82.Cells(Final, 10) = False
                Hoja82.Cells(Final, 11) = False
                Hoja82.Cells(Final, 12) = False
                Hoja82.Cells(Final, 13) = False
                Hoja82.Cells(Final, 14) = False
                Hoja82.Cells(Final, 15) = False
                Hoja82.Cells(Final, 16) = False
                Hoja82.Cells(Final, 17) = False
                Hoja82.Cells(Final, 18) = False
                Hoja82.Cells(Final, 19) = False
                Hoja82.Cells(Final, 20) = False
                Hoja82.Cells(Final, 21) = False
                Hoja82.Cells(Final, 22) = False
                Hoja82.Cells(Final, 23) = False
                Hoja82.Cells(Final, 24) = False
                Hoja82.Cells(Final, 25) = False
                Hoja82.Cells(Final, 26) = False
                Hoja82.Cells(Final, 27) = False
                Hoja82.Cells(Final, 28) = False
                Hoja82.Cells(Final, 29) = False
                Hoja82.Cells(Final, 30) = False
                Hoja82.Cells(Final, 31) = False
                Hoja82.Cells(Final, 32) = False
                Hoja82.Cells(Final, 33) = False
                Hoja82.Cells(Final, 34) = False
                
                
                ElseIf Me.OptionButton2.Value = True Then
                Hoja82.Cells(Final, 4).Value = True
                Hoja82.Cells(Final, 5).Value = True
                Hoja82.Cells(Final, 6).Value = True
                Hoja82.Cells(Final, 7).Value = True
                Hoja82.Cells(Final, 8).Value = True
                Hoja82.Cells(Final, 9).Value = False
                Hoja82.Cells(Final, 10) = True
                Hoja82.Cells(Final, 11) = True
                Hoja82.Cells(Final, 12) = True
                Hoja82.Cells(Final, 13) = False
                Hoja82.Cells(Final, 14) = False
                Hoja82.Cells(Final, 15) = False
                Hoja82.Cells(Final, 16) = True
                Hoja82.Cells(Final, 17) = True
                Hoja82.Cells(Final, 18) = False
                Hoja82.Cells(Final, 19) = True
                Hoja82.Cells(Final, 20) = True
                Hoja82.Cells(Final, 21) = True
                Hoja82.Cells(Final, 22) = True
                Hoja82.Cells(Final, 23) = True
                Hoja82.Cells(Final, 24) = True
                Hoja82.Cells(Final, 25) = True
                Hoja82.Cells(Final, 26) = True
                Hoja82.Cells(Final, 27) = False
                Hoja82.Cells(Final, 28) = False
                Hoja82.Cells(Final, 29) = False
                Hoja82.Cells(Final, 30) = False
                Hoja82.Cells(Final, 31) = False
                Hoja82.Cells(Final, 32) = False
                Hoja82.Cells(Final, 33) = False
                Hoja82.Cells(Final, 34) = False
                
                End If
                
                'Botones

                Hoja82.Cells(Final, 35) = True
                Hoja82.Cells(Final, 36) = True
                Hoja82.Cells(Final, 37) = True
                Hoja82.Cells(Final, 38) = True
                Hoja82.Cells(Final, 39) = True
                Hoja82.Cells(Final, 40) = True
                Hoja82.Cells(Final, 41) = True
                Hoja82.Cells(Final, 42) = True
                Hoja82.Cells(Final, 43) = True
                Hoja82.Cells(Final, 44) = True
                Hoja82.Cells(Final, 45) = True
                Hoja82.Cells(Final, 46) = True
                Hoja82.Cells(Final, 47) = True
                Hoja82.Cells(Final, 48) = True
                Hoja82.Cells(Final, 49) = True
                Hoja82.Cells(Final, 50) = True
                Hoja82.Cells(Final, 51) = True
                Hoja82.Cells(Final, 52) = True
                Hoja82.Cells(Final, 53) = True
                Hoja82.Cells(Final, 54) = True
            

                Me.txt_nUser.Text = ""
                Me.txt_pass1.Text = ""
                Me.txt_pass2.Text = ""

                Me.txt_nUser.SetFocus
                
                Hoja82.Protect (Seguridad)
                
                
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
