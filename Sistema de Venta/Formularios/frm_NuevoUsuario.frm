VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_NuevoUsuario 
   Caption         =   "Registro de Usuarios"
   ClientHeight    =   3795
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   5130
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
    
On Error GoTo Salir

Application.ScreenUpdating = False

If Hoja91.Visible = xlSheetVisible Then

Hoja91.Select
        
Final = GetNuevoR(Hoja91)
        
        For Registro = 2 To Final
            If Hoja91.Cells(Registro, 1) = Me.txt_nUser.Text Then
                Me.txt_nUser.BackColor = &H8080FF
                MsgBox ("El usuario ya existe" + Chr(13) + "Ingrese un usuario diferente")
                Me.txt_nUser.SetFocus
                Exit Sub
                Exit For
            End If
        Next
        
      If Me.txt_pass1.Text = Me.txt_pass2.Text Then
                Me.txt_nUser.BackColor = &HFFFFFF
                Hoja91.Cells(Final, 1) = Me.txt_nUser.Text
                Hoja91.Cells(Final, 2) = Me.txt_pass1.Text
                        If Me.OptionButton1.Value = True Then
                            Hoja91.Cells(Final, 3) = "USUARIO"
                                Else
                            Hoja91.Cells(Final, 3) = "ADMINISTRADOR"
                        End If
                 MsgBox "Espere un Momento, Cargando Permisos y Restricciones...!", vbInformation, "Configuración"
        
        'VALORES PARA HOJAS Y BOTONES
            'GRUPO ADMINISTRATIVO
                'Hojas
                'N/A
                'Botones
                 Hoja91.Cells(Final, 61).Value = True
                 Hoja91.Cells(Final, 78).Value = True
                 Hoja91.Cells(Final, 86).Value = True
                
                
            'GRUPO GESTOR DE INVENTARIO 1
                'Hojas
                Hoja91.Cells(Final, 13) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 14) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 15) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 7) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 25) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 28) = Me.CheckBox7.Value
                'Botones
                Hoja91.Cells(Final, 62) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 63) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 64) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 96) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 65) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 66) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 67) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 113) = Me.CheckBox1.Value
        
            'GRUPO GESTOR DE RECURSOS HUMANOS 2
                'Hojas
                Hoja91.Cells(Final, 8) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 9) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 19) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 20) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 21) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 26) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 27) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 35) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 36) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 37) = Me.CheckBox7.Value
                'Botones
                Hoja91.Cells(Final, 59) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 60) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 68) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 69) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 70) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 71) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 72) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 114) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 115) = Me.CheckBox1.Value
               
        
            'GRUPO EGRESOS VARIOS 3
                'Hojas
                Hoja91.Cells(Final, 10) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 11) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 16) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 17) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 18) = Me.CheckBox7.Value
                'Botones
                Hoja91.Cells(Final, 73) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 74) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 75) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 76) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 77) = Me.CheckBox1.Value
                
                
            'GRUPO GANADERIA 4
                'Hojas
                Hoja91.Cells(Final, 31) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 32) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 33) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 34) = Me.CheckBox7.Value
                'Botones
                Hoja91.Cells(Final, 79) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 80) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 81) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 82) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 83) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 84) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 85) = Me.CheckBox1.Value
            
            'GRUPO FINANCIERO 5
                'Hojas
                Hoja91.Cells(Final, 5) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 6) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 12) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 30) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 36) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 37) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 38) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 39) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 40) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 41) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 42) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 43) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 44) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 45) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 46) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 47) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 48) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 49) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 50) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 51) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 52) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 53) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 54) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 55) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 56) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 57) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 58) = Me.CheckBox7.Value    ''''''
                'Botones
                Hoja91.Cells(Final, 87) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 88) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 89) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 90) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 91) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 92) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 93) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 94) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 95) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 105) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 106) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 107) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 108) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 109) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 110) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 111) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 112) = Me.CheckBox1.Value
        
            'GRUPO CONFIGURACIÓN 6
                'Hojas
                Hoja91.Cells(Final, 4) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 22) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 23) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 24) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 29) = Me.CheckBox7.Value
                'Botones
                Hoja91.Cells(Final, 97) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 98) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 99) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 100) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 101) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 102) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 103) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 104) = Me.CheckBox6.Value
                
                '----------------------------------------------
                
                Me.txt_nUser.Text = ""
                Me.txt_pass1.Text = ""
                Me.txt_pass2.Text = ""
                Me.CheckBox1.Value = False
               Me.CheckBox7.Value = False
                Me.CheckBox6.Value = False

                
                Me.txt_nUser.SetFocus
                
    Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True

    
                 MsgBox "Usuario registrado satisfactoriamente", vbInformation, "Configuración"
            Else
                MsgBox "Las contraseñas deben coincidir"
    End If
   
  ElseIf Hoja91.Visible = xlSheetVeryHidden Then
    Hoja91.Visible = xlSheetVisible
     
        Hoja91.Select
        
        Final = GetNuevoR(Hoja91)
        
        For Registro = 2 To Final
            If Hoja91.Cells(Registro, 1) = Me.txt_nUser.Text Then
                Me.txt_nUser.BackColor = &H8080FF
                MsgBox ("El usuario ya existe" + Chr(13) + "Ingrese un usuario diferente")
                Me.txt_nUser.SetFocus
                Exit Sub
                Exit For
            End If
        Next
        
        If Me.txt_pass1.Text = Me.txt_pass2.Text Then
                Me.txt_nUser.BackColor = &HFFFFFF
                Hoja91.Cells(Final, 1) = Me.txt_nUser.Text
                Hoja91.Cells(Final, 2) = Me.txt_pass1.Text
                        If Me.OptionButton1.Value = True Then
                            Hoja91.Cells(Final, 3) = "USUARIO"
                                Else
                            Hoja91.Cells(Final, 3) = "ADMINISTRADOR"
                        End If
                 MsgBox "Espere un Momento, Cargando Permisos y Restricciones...!", vbInformation, "Configuración"
        
        'VALORES PARA HOJAS Y BOTONES
            'GRUPO ADMINISTRATIVO
                'Hojas
                'N/A
                'Botones
                 Hoja91.Cells(Final, 61).Value = True
                 Hoja91.Cells(Final, 78).Value = True
                 Hoja91.Cells(Final, 86).Value = True
                
                
            'GRUPO GESTOR DE INVENTARIO 1
                'Hojas
                Hoja91.Cells(Final, 13) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 14) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 15) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 7) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 25) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 28) = Me.CheckBox7.Value
                'Botones
                Hoja91.Cells(Final, 62) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 63) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 64) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 96) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 65) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 66) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 67) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 113) = Me.CheckBox1.Value
        
            'GRUPO GESTOR DE RECURSOS HUMANOS 2
                'Hojas
                Hoja91.Cells(Final, 8) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 9) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 19) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 20) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 21) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 26) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 27) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 35) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 36) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 37) = Me.CheckBox7.Value
                'Botones
                Hoja91.Cells(Final, 59) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 60) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 68) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 69) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 70) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 71) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 72) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 114) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 115) = Me.CheckBox1.Value
               
        
            'GRUPO EGRESOS VARIOS 3
                'Hojas
                Hoja91.Cells(Final, 10) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 11) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 16) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 17) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 18) = Me.CheckBox7.Value
                'Botones
                Hoja91.Cells(Final, 73) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 74) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 75) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 76) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 77) = Me.CheckBox1.Value
                
                
            'GRUPO GANADERIA 4
                'Hojas
                Hoja91.Cells(Final, 31) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 32) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 33) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 34) = Me.CheckBox7.Value
                'Botones
                Hoja91.Cells(Final, 79) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 80) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 81) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 82) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 83) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 84) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 85) = Me.CheckBox1.Value
            
            'GRUPO FINANCIERO 5
                'Hojas
                Hoja91.Cells(Final, 5) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 6) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 12) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 30) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 36) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 37) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 38) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 39) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 40) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 41) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 42) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 43) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 44) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 45) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 46) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 47) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 48) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 49) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 50) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 51) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 52) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 53) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 54) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 55) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 56) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 57) = Me.CheckBox7.Value    ''''''
                Hoja91.Cells(Final, 58) = Me.CheckBox7.Value    ''''''
                'Botones
                Hoja91.Cells(Final, 87) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 88) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 89) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 90) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 91) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 92) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 93) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 94) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 95) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 105) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 106) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 107) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 108) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 109) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 110) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 111) = Me.CheckBox1.Value
                Hoja91.Cells(Final, 112) = Me.CheckBox1.Value
        
            'GRUPO CONFIGURACIÓN 6
                'Hojas
                Hoja91.Cells(Final, 4) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 22) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 23) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 24) = Me.CheckBox7.Value
                Hoja91.Cells(Final, 29) = Me.CheckBox7.Value
                'Botones
                Hoja91.Cells(Final, 97) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 98) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 99) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 100) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 101) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 102) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 103) = Me.CheckBox6.Value
                Hoja91.Cells(Final, 104) = Me.CheckBox6.Value
                
                '----------------------------------------------
                
                Me.txt_nUser.Text = ""
                Me.txt_pass1.Text = ""
                Me.txt_pass2.Text = ""
                Me.CheckBox1.Value = False
               Me.CheckBox7.Value = False
                Me.CheckBox6.Value = False

                
                Me.txt_nUser.SetFocus
                
       
    
                 MsgBox "Usuario registrado satisfactoriamente", vbInformation, "Configuración"
            Else
                MsgBox "Las contraseñas deben coincidir"
        End If
   
    Hoja91.Visible = xlSheetVeryHidden
  End If
  
   Application.EnableEvents = False
              ThisWorkbook.Save
        Application.EnableEvents = True
   
   Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Usuarios"
 End If
 
End Sub

Private Sub cmd_Finalizar_Click()
    Unload Me
End Sub


Private Sub OptionButton1_Click()
    If OptionButton1.Value = True Then
        Me.CheckBox1.Value = True
        Me.CheckBox6.Value = False
    End If
    
End Sub
Private Sub OptionButton2_Click()
    If OptionButton2.Value = True Then
        Me.CheckBox1.Value = True
        Me.CheckBox6.Value = True
    End If
    
End Sub

Private Sub UserForm_Initialize()
    Me.CheckBox1.Value = True
    If Hoja92.Range("G1") = "LUICHIX" And Hoja92.Range("H1") = "ADMINISTRADOR" Then
        
        Me.CheckBox7.Visible = True
        Me.CheckBox7.Enabled = True
        Me.CheckBox7.Value = False
    End If
End Sub
