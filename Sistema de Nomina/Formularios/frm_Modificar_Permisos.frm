VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Modificar_Permisos 
   Caption         =   "Modificar Permisos"
   ClientHeight    =   5880
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   4990
   OleObjectBlob   =   "frm_Modificar_Permisos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Modificar_Permisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub ComboBox1_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String

For Fila = 1 To ComboBox1.ListCount
    ComboBox1.RemoveItem 0
Next Fila

Final = GetUltimoR(Hoja82)

    For Fila = 3 To Final
        Lista = Hoja82.Cells(Fila, 1)
        ComboBox1.AddItem (Lista)
    Next
End Sub

Private Sub cmd_Guardar_Click()
Dim Fila As Long
Dim Final As Long
Dim Seguridad As String
On Error GoTo Salir
Application.ScreenUpdating = False


Hoja82.Select

Seguridad = Hoja83.Range("L1").Text
If Me.ComboBox1.Text = Empty Then
    MsgBox "Debe seleccionar un usuario..!", vbInformation
    Exit Sub
End If

If Me.OptionButton1.Value = False And Me.OptionButton2.Value = False Then
    MsgBox ("Debe seleccionar un nivel de privilegio..!"), vbInformation, "Gestor de Recursos Humanos"
    Exit Sub
End If
                
                    If Me.ComboBox1.Text = Hoja83.Range("G1").Text Then
                        MsgBox ("El usuario actual no puede ser modificado..!"), vbCritical, "Gestor de Recursos Humanos"
                        Exit Sub
                    End If

    Final = GetUltimoR(Hoja82)
   


    For Fila = 3 To Final
        If Me.ComboBox1.Text = Hoja82.Cells(Fila, 1) Then
          'VALORES PARA HOJAS Y BOTONES
            'GRUPO ADMINISTRATIVO
                Hoja82.Unprotect (Seguridad)
                'Hojas
                
                Application.Cursor = xlWait
                If Me.OptionButton1.Value = True Then
                Hoja82.Cells(Final, 3).Value = "USUARIO"
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
                Hoja82.Cells(Final, 3).Value = "ADMINISTRADOR"
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
                Hoja82.Cells(Final, 27) = True
                Hoja82.Cells(Final, 28) = True
                Hoja82.Cells(Final, 29) = True
                Hoja82.Cells(Final, 30) = True
                Hoja82.Cells(Final, 31) = True
                Hoja82.Cells(Final, 32) = True
                Hoja82.Cells(Final, 33) = True
                Hoja82.Cells(Final, 34) = True

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
                
                Hoja83.Protect (Seguridad)
            Exit For
        End If
    Next

    Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True
    
    Application.Cursor = xlDefault

MsgBox "Cambios guardados satisfactoriamente..!", vbInformation, "Configuración"

    Unload Me
Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Configuración"
 End If

End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub


Private Sub UserForm_Initialize()
EliminarTitulo Me.Caption
    Me.Height = Me.Height - 20
End Sub

