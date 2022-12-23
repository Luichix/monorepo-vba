VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_iniciosesion 
   Caption         =   "GESTOR ADMINISTRATIVO"
   ClientHeight    =   3470
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   9340.001
   OleObjectBlob   =   "form_iniciosesion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "form_iniciosesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub btn_Ingresar_Click()
Dim Usuario As String
Dim Fila, Final As Long
Dim password As String, UsuarioEncontrado As String, yaExiste As Byte, Status As String
Dim Rango As Range
Dim Titulo As String
Dim Hoja As Worksheet
Dim vHoja(58) As String
Dim vBoton(42) As String
Dim i As Byte
Dim ii As Byte
Dim x As Byte

Application.ScreenUpdating = False

Titulo = "Gestor de Inventarios"


yaExiste = Application.WorksheetFunction.CountIf(Hoja19.Range("tbl_Usuario[Usuario]"), Me.txt_Usuario.Text)
Set Rango = Hoja19.Range("tbl_Usuario[Usuario]")

If Me.txt_Usuario.Text = "" Or Me.txt_Contraseña.Text = "" Then
    MsgBox "Introduce usuario y contraseña", vbExclamation, Titulo
    Me.txt_Usuario.SetFocus

            ElseIf yaExiste = 0 Then
                MsgBox "El usuario '" & Me.txt_Usuario.Text & "' no existe", vbExclamation, Titulo
            
            ElseIf yaExiste = 1 Then
                UsuarioEncontrado = Rango.Find(What:=Me.txt_Usuario.Text, after:=Rango.Range("A1"), _
                                                LookAt:=xlWhole, MatchCase:=False).Address
                
                password = Hoja19.Range(UsuarioEncontrado).Offset(0, 1).Value
                Status = Hoja19.Range(UsuarioEncontrado).Offset(0, 2).Value
                
                'Permisos y restricciones
                vHoja(1) = Hoja19.Range(UsuarioEncontrado).Offset(0, 3).Value
                vHoja(2) = Hoja19.Range(UsuarioEncontrado).Offset(0, 4).Value
                vHoja(3) = Hoja19.Range(UsuarioEncontrado).Offset(0, 5).Value
                vHoja(4) = Hoja19.Range(UsuarioEncontrado).Offset(0, 6).Value
                vHoja(5) = Hoja19.Range(UsuarioEncontrado).Offset(0, 7).Value
                vHoja(6) = Hoja19.Range(UsuarioEncontrado).Offset(0, 8).Value
                vHoja(7) = Hoja19.Range(UsuarioEncontrado).Offset(0, 9).Value
                vHoja(8) = Hoja19.Range(UsuarioEncontrado).Offset(0, 10).Value
                vHoja(9) = Hoja19.Range(UsuarioEncontrado).Offset(0, 11).Value
                vHoja(10) = Hoja19.Range(UsuarioEncontrado).Offset(0, 12).Value
                vHoja(11) = Hoja19.Range(UsuarioEncontrado).Offset(0, 13).Value
                vHoja(12) = Hoja19.Range(UsuarioEncontrado).Offset(0, 14).Value
                vHoja(13) = Hoja19.Range(UsuarioEncontrado).Offset(0, 15).Value
                vHoja(14) = Hoja19.Range(UsuarioEncontrado).Offset(0, 16).Value
                vHoja(15) = Hoja19.Range(UsuarioEncontrado).Offset(0, 17).Value
                vHoja(16) = Hoja19.Range(UsuarioEncontrado).Offset(0, 18).Value
                vHoja(17) = Hoja19.Range(UsuarioEncontrado).Offset(0, 19).Value
                vHoja(18) = Hoja19.Range(UsuarioEncontrado).Offset(0, 20).Value
                vHoja(19) = Hoja19.Range(UsuarioEncontrado).Offset(0, 21).Value
                vHoja(21) = Hoja19.Range(UsuarioEncontrado).Offset(0, 22).Value
                vHoja(22) = Hoja19.Range(UsuarioEncontrado).Offset(0, 23).Value
                vHoja(23) = Hoja19.Range(UsuarioEncontrado).Offset(0, 24).Value
                vHoja(24) = Hoja19.Range(UsuarioEncontrado).Offset(0, 25).Value
                vHoja(25) = Hoja19.Range(UsuarioEncontrado).Offset(0, 26).Value
                vHoja(26) = Hoja19.Range(UsuarioEncontrado).Offset(0, 27).Value
                vHoja(27) = Hoja19.Range(UsuarioEncontrado).Offset(0, 28).Value
                vHoja(28) = Hoja19.Range(UsuarioEncontrado).Offset(0, 29).Value
                vHoja(29) = Hoja19.Range(UsuarioEncontrado).Offset(0, 30).Value
                vHoja(30) = Hoja19.Range(UsuarioEncontrado).Offset(0, 31).Value
                vHoja(31) = Hoja19.Range(UsuarioEncontrado).Offset(0, 32).Value
                vHoja(32) = Hoja19.Range(UsuarioEncontrado).Offset(0, 33).Value
                vHoja(33) = Hoja19.Range(UsuarioEncontrado).Offset(0, 34).Value
                vHoja(34) = Hoja19.Range(UsuarioEncontrado).Offset(0, 35).Value
                vHoja(35) = Hoja19.Range(UsuarioEncontrado).Offset(0, 36).Value
                vHoja(36) = Hoja19.Range(UsuarioEncontrado).Offset(0, 37).Value
                vHoja(37) = Hoja19.Range(UsuarioEncontrado).Offset(0, 38).Value
                vHoja(38) = Hoja19.Range(UsuarioEncontrado).Offset(0, 39).Value
                vHoja(39) = Hoja19.Range(UsuarioEncontrado).Offset(0, 40).Value
                vHoja(40) = Hoja19.Range(UsuarioEncontrado).Offset(0, 41).Value
                vHoja(41) = Hoja19.Range(UsuarioEncontrado).Offset(0, 42).Value
                vHoja(42) = Hoja19.Range(UsuarioEncontrado).Offset(0, 43).Value
                vHoja(43) = Hoja19.Range(UsuarioEncontrado).Offset(0, 44).Value
                vHoja(44) = Hoja19.Range(UsuarioEncontrado).Offset(0, 45).Value
                vHoja(45) = Hoja19.Range(UsuarioEncontrado).Offset(0, 46).Value
                vHoja(46) = Hoja19.Range(UsuarioEncontrado).Offset(0, 47).Value
                vHoja(47) = Hoja19.Range(UsuarioEncontrado).Offset(0, 48).Value
                vHoja(48) = Hoja19.Range(UsuarioEncontrado).Offset(0, 49).Value
                vHoja(49) = Hoja19.Range(UsuarioEncontrado).Offset(0, 50).Value
                vHoja(50) = Hoja19.Range(UsuarioEncontrado).Offset(0, 51).Value
                vHoja(51) = Hoja19.Range(UsuarioEncontrado).Offset(0, 52).Value
                vHoja(52) = Hoja19.Range(UsuarioEncontrado).Offset(0, 53).Value
                vHoja(53) = Hoja19.Range(UsuarioEncontrado).Offset(0, 54).Value
                vHoja(54) = Hoja19.Range(UsuarioEncontrado).Offset(0, 55).Value
                vHoja(55) = Hoja19.Range(UsuarioEncontrado).Offset(0, 56).Value
                vHoja(56) = Hoja19.Range(UsuarioEncontrado).Offset(0, 57).Value
                vHoja(55) = Hoja19.Range(UsuarioEncontrado).Offset(0, 56).Value
                vHoja(56) = Hoja19.Range(UsuarioEncontrado).Offset(0, 57).Value
                vHoja(57) = Hoja19.Range(UsuarioEncontrado).Offset(0, 58).Value
                vHoja(58) = Hoja19.Range(UsuarioEncontrado).Offset(0, 59).Value
                 
                 
                vBoton(1) = Hoja19.Range(UsuarioEncontrado).Offset(0, 60).Value
                vBoton(2) = Hoja19.Range(UsuarioEncontrado).Offset(0, 61).Value
                vBoton(3) = Hoja19.Range(UsuarioEncontrado).Offset(0, 62).Value
                vBoton(4) = Hoja19.Range(UsuarioEncontrado).Offset(0, 63).Value
                vBoton(5) = Hoja19.Range(UsuarioEncontrado).Offset(0, 64).Value
                vBoton(6) = Hoja19.Range(UsuarioEncontrado).Offset(0, 65).Value
                vBoton(7) = Hoja19.Range(UsuarioEncontrado).Offset(0, 66).Value
                vBoton(8) = Hoja19.Range(UsuarioEncontrado).Offset(0, 67).Value
                vBoton(9) = Hoja19.Range(UsuarioEncontrado).Offset(0, 68).Value
                vBoton(10) = Hoja19.Range(UsuarioEncontrado).Offset(0, 69).Value
                vBoton(11) = Hoja19.Range(UsuarioEncontrado).Offset(0, 70).Value
                vBoton(12) = Hoja19.Range(UsuarioEncontrado).Offset(0, 71).Value
                vBoton(13) = Hoja19.Range(UsuarioEncontrado).Offset(0, 72).Value
                vBoton(14) = Hoja19.Range(UsuarioEncontrado).Offset(0, 73).Value
                vBoton(15) = Hoja19.Range(UsuarioEncontrado).Offset(0, 74).Value
                vBoton(16) = Hoja19.Range(UsuarioEncontrado).Offset(0, 75).Value
                vBoton(17) = Hoja19.Range(UsuarioEncontrado).Offset(0, 76).Value
                vBoton(18) = Hoja19.Range(UsuarioEncontrado).Offset(0, 77).Value
                vBoton(19) = Hoja19.Range(UsuarioEncontrado).Offset(0, 78).Value
                vBoton(20) = Hoja19.Range(UsuarioEncontrado).Offset(0, 79).Value
                vBoton(21) = Hoja19.Range(UsuarioEncontrado).Offset(0, 80).Value
                vBoton(22) = Hoja19.Range(UsuarioEncontrado).Offset(0, 81).Value
                vBoton(23) = Hoja19.Range(UsuarioEncontrado).Offset(0, 82).Value
                vBoton(24) = Hoja19.Range(UsuarioEncontrado).Offset(0, 83).Value
                vBoton(25) = Hoja19.Range(UsuarioEncontrado).Offset(0, 84).Value
                vBoton(26) = Hoja19.Range(UsuarioEncontrado).Offset(0, 85).Value
                vBoton(27) = Hoja19.Range(UsuarioEncontrado).Offset(0, 86).Value
                vBoton(28) = Hoja19.Range(UsuarioEncontrado).Offset(0, 87).Value
                vBoton(29) = Hoja19.Range(UsuarioEncontrado).Offset(0, 88).Value
                vBoton(30) = Hoja19.Range(UsuarioEncontrado).Offset(0, 89).Value
                vBoton(31) = Hoja19.Range(UsuarioEncontrado).Offset(0, 90).Value
                vBoton(32) = Hoja19.Range(UsuarioEncontrado).Offset(0, 91).Value
                vBoton(33) = Hoja19.Range(UsuarioEncontrado).Offset(0, 92).Value
                vBoton(34) = Hoja19.Range(UsuarioEncontrado).Offset(0, 93).Value
                vBoton(35) = Hoja19.Range(UsuarioEncontrado).Offset(0, 94).Value
                vBoton(36) = Hoja19.Range(UsuarioEncontrado).Offset(0, 95).Value
                vBoton(37) = Hoja19.Range(UsuarioEncontrado).Offset(0, 96).Value
                vBoton(38) = Hoja19.Range(UsuarioEncontrado).Offset(0, 97).Value
                vBoton(39) = Hoja19.Range(UsuarioEncontrado).Offset(0, 98).Value
                vBoton(40) = Hoja19.Range(UsuarioEncontrado).Offset(0, 99).Value
                vBoton(41) = Hoja19.Range(UsuarioEncontrado).Offset(0, 100).Value
                vBoton(42) = Hoja19.Range(UsuarioEncontrado).Offset(0, 101).Value
                'vBoton(43) = Hoja19.Range(UsuarioEncontrado).Offset(0, 102).Value
                'vBoton(44) = Hoja19.Range(UsuarioEncontrado).Offset(0, 103).Value
    
            If Hoja19.Range(UsuarioEncontrado).Value = Me.txt_Usuario.Text And password = Me.txt_Contraseña.Text Then
            
                        'Validando los permisos y restricciones en las hojas de cálculo
                        For i = 1 To 19
                            For Each Hoja In Worksheets
                            If Hoja.CodeName = "Hoja" & i Then
                                If vHoja(i) = False Then
                                    Hoja.Visible = xlSheetVeryHidden
                                Else
                                    Hoja.Visible = xlSheetVisible
                                End If
                            End If
                            Next Hoja
                        Next i
                                                                     
                           For ii = 21 To 58
                            For Each Hoja In Worksheets
                            If Hoja.CodeName = "Hoja" & ii Then
                                If vHoja(ii) = False Then
                                    Hoja.Visible = xlSheetVeryHidden
                                Else
                                    Hoja.Visible = xlSheetVisible
                                End If
                            End If
                            Next Hoja
                        Next ii
                        
                         'Validando los permisos y restricciones de los botones
                     
   
                        For x = 1 To 42
                             If vBoton(x) = True Then
                                RetVal(x) = True
                                CintaDeRibbon.InvalidateControl ("Button" & (x))
                               
                            Else
                                RetVal(x) = False
                                CintaDeRibbon.InvalidateControl ("Button" & (x))
                              
                            End If
                        Next x
                        
     
                        ' Registrar al usuario en la hoja Logs
                              
                              Final = GetNuevoR(Hoja21)
                                  Hoja21.Cells(Final, 1) = "=NOW()"
                                  Hoja21.Cells(Final, 1).Copy
                                  Hoja21.Cells(Final, 1).PasteSpecial Paste:=xlPasteValues
                                  Application.CutCopyMode = False
                                  
                                  Hoja21.Cells(Final, 2) = Me.txt_Usuario.Text
                                  
                                  Hoja20.txt_UsuarioActual.Caption = "Usuario actual: " & UCase(Me.txt_Usuario.Text)
                                  
                                  Hoja21.Cells(Final, 3) = Status
                                  
                    
                                 
                                  Hoja21.Range("G1") = Me.txt_Usuario.Text
                                  Hoja21.Range("H1") = Status
                                  
                                  
                                  
                                  ThisWorkbook.Save
                              
                              
                                  Unload Me
                                  Hoja20.Activate
                        Else
                     MsgBox "La contraseña es incorrecta", vbExclamation, Titulo
            End If
End If

Application.ScreenUpdating = True


End Sub
Private Sub btn_Salir_Click()
    Unload Me
    ThisWorkbook.Save
    ThisWorkbook.Application.DisplayAlerts = False
    Application.ActiveWorkbook.Close
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Unload Me
        ThisWorkbook.Save
        ThisWorkbook.Application.DisplayAlerts = False
        Application.ActiveWorkbook.Close
    End If
End Sub

