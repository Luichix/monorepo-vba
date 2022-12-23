VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_iniciosesion 
   Caption         =   "GESTOR ADMINISTRATIVO"
   ClientHeight    =   3435
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   11020
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
Dim Seguridad As String
Dim Usuario As String
Dim Fila, Final As Long
Dim password As String, UsuarioEncontrado As String, yaExiste As Byte, Status As String
Dim Rango As Range
Dim Titulo As String
Dim Hoja As Worksheet
Dim vHoja(99) As String
Dim vBoton(99) As String
Dim i As Byte
Dim X As Byte

Application.ScreenUpdating = False

Titulo = "Gestor de Recursos Humanos"


yaExiste = Application.WorksheetFunction.CountIf(Hoja82.Range("Tbl_usuario[USUARIO]"), Me.txt_Usuario.Text)
Set Rango = Hoja82.Range("Tbl_usuario[USUARIO]")

If Me.txt_Usuario.Text = "" Or Me.txt_Contraseña.Text = "" Then
    MsgBox "Introduce usuario y contraseña", vbExclamation, Titulo
    Me.txt_Usuario.SetFocus

            ElseIf yaExiste = 0 Then
                MsgBox "El usuario '" & Me.txt_Usuario.Text & "' no existe", vbExclamation, Titulo
            
            ElseIf yaExiste = 1 Then
                UsuarioEncontrado = Rango.Find(What:=Me.txt_Usuario.Text, After:=Rango.Range("A1"), _
                                                LookAt:=xlWhole, MatchCase:=False).Address
                
                password = Hoja82.Range(UsuarioEncontrado).Offset(0, 1).Value
                Status = Hoja82.Range(UsuarioEncontrado).Offset(0, 2).Value
                
                'Permisos y restricciones
                vHoja(1) = Hoja82.Range(UsuarioEncontrado).Offset(0, 3).Value
                vHoja(2) = Hoja82.Range(UsuarioEncontrado).Offset(0, 4).Value
                vHoja(3) = Hoja82.Range(UsuarioEncontrado).Offset(0, 5).Value
                vHoja(4) = Hoja82.Range(UsuarioEncontrado).Offset(0, 6).Value
                vHoja(5) = Hoja82.Range(UsuarioEncontrado).Offset(0, 7).Value
                vHoja(58) = Hoja82.Range(UsuarioEncontrado).Offset(0, 8).Value
                vHoja(6) = Hoja82.Range(UsuarioEncontrado).Offset(0, 9).Value
                vHoja(7) = Hoja82.Range(UsuarioEncontrado).Offset(0, 10).Value
                vHoja(8) = Hoja82.Range(UsuarioEncontrado).Offset(0, 11).Value
                vHoja(81) = Hoja82.Range(UsuarioEncontrado).Offset(0, 12).Value
                vHoja(82) = Hoja82.Range(UsuarioEncontrado).Offset(0, 13).Value
                vHoja(83) = Hoja82.Range(UsuarioEncontrado).Offset(0, 14).Value
                vHoja(9) = Hoja82.Range(UsuarioEncontrado).Offset(0, 15).Value
                vHoja(10) = Hoja82.Range(UsuarioEncontrado).Offset(0, 16).Value
                vHoja(11) = Hoja82.Range(UsuarioEncontrado).Offset(0, 17).Value
                vHoja(12) = Hoja82.Range(UsuarioEncontrado).Offset(0, 18).Value
                vHoja(13) = Hoja82.Range(UsuarioEncontrado).Offset(0, 19).Value
                vHoja(14) = Hoja82.Range(UsuarioEncontrado).Offset(0, 20).Value
                vHoja(15) = Hoja82.Range(UsuarioEncontrado).Offset(0, 21).Value
                vHoja(16) = Hoja82.Range(UsuarioEncontrado).Offset(0, 22).Value
                vHoja(17) = Hoja82.Range(UsuarioEncontrado).Offset(0, 23).Value
                vHoja(18) = Hoja82.Range(UsuarioEncontrado).Offset(0, 24).Value
                vHoja(19) = Hoja82.Range(UsuarioEncontrado).Offset(0, 25).Value
                vHoja(20) = Hoja82.Range(UsuarioEncontrado).Offset(0, 26).Value
                vHoja(21) = Hoja82.Range(UsuarioEncontrado).Offset(0, 27).Value
                vHoja(22) = Hoja82.Range(UsuarioEncontrado).Offset(0, 28).Value
                vHoja(23) = Hoja82.Range(UsuarioEncontrado).Offset(0, 29).Value
                vHoja(24) = Hoja82.Range(UsuarioEncontrado).Offset(0, 30).Value
                vHoja(25) = Hoja82.Range(UsuarioEncontrado).Offset(0, 31).Value
                vHoja(26) = Hoja82.Range(UsuarioEncontrado).Offset(0, 32).Value
                vHoja(27) = Hoja82.Range(UsuarioEncontrado).Offset(0, 33).Value
                 
                 
                vBoton(1) = Hoja82.Range(UsuarioEncontrado).Offset(0, 34).Value
                vBoton(2) = Hoja82.Range(UsuarioEncontrado).Offset(0, 35).Value
                vBoton(3) = Hoja82.Range(UsuarioEncontrado).Offset(0, 36).Value
                vBoton(4) = Hoja82.Range(UsuarioEncontrado).Offset(0, 37).Value
                vBoton(5) = Hoja82.Range(UsuarioEncontrado).Offset(0, 38).Value
                vBoton(6) = Hoja82.Range(UsuarioEncontrado).Offset(0, 39).Value
                vBoton(7) = Hoja82.Range(UsuarioEncontrado).Offset(0, 40).Value
                vBoton(8) = Hoja82.Range(UsuarioEncontrado).Offset(0, 41).Value
                vBoton(9) = Hoja82.Range(UsuarioEncontrado).Offset(0, 42).Value
                vBoton(10) = Hoja82.Range(UsuarioEncontrado).Offset(0, 43).Value
                vBoton(11) = Hoja82.Range(UsuarioEncontrado).Offset(0, 44).Value
                vBoton(12) = Hoja82.Range(UsuarioEncontrado).Offset(0, 45).Value
                vBoton(13) = Hoja82.Range(UsuarioEncontrado).Offset(0, 46).Value
                vBoton(14) = Hoja82.Range(UsuarioEncontrado).Offset(0, 47).Value
                vBoton(15) = Hoja82.Range(UsuarioEncontrado).Offset(0, 48).Value
                vBoton(16) = Hoja82.Range(UsuarioEncontrado).Offset(0, 49).Value
                vBoton(17) = Hoja82.Range(UsuarioEncontrado).Offset(0, 50).Value
                vBoton(18) = Hoja82.Range(UsuarioEncontrado).Offset(0, 51).Value
                vBoton(19) = Hoja82.Range(UsuarioEncontrado).Offset(0, 52).Value
                vBoton(20) = Hoja82.Range(UsuarioEncontrado).Offset(0, 53).Value
                

             
                                
            If Hoja82.Range(UsuarioEncontrado).Value = Me.txt_Usuario.Text And password = Me.txt_Contraseña.Text Then
            
                        'Validando los permisos y restricciones en las hojas de cálculo
                        For i = 1 To 99
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
                                                                     
                      
                         'Validando los permisos y restricciones de los botones
                     
   
                        For X = 1 To 20
                             If vBoton(X) = True Then
                                RetVal(X) = True
                                If Not CintaDeRibbon Is Nothing Then
                                    CintaDeRibbon.InvalidateControl ("Button" & (X))
                                    Else
                                        MsgBox "Requiere reiniciar la aplicacion de excel", vbInformation, "GESTOR"
                                        Exit For
                                End If
                            Else
                                RetVal(X) = False
                                If Not CintaDeRibbon Is Nothing Then
                                    CintaDeRibbon.InvalidateControl ("Button" & (X))
                                    Else
                                        MsgBox "Requiere reiniciar la aplicacion de excel", vbInformation, "GESTOR"
                                        Exit For
                                End If
                            End If
                        Next X
                        
     Seguridad = Hoja83.Range("L1").Text
                        ' Registrar al usuario en la hoja Logs
                                Hoja83.Unprotect (Seguridad)
                              Final = GetNuevoR(Hoja83)
                                  Hoja83.Cells(Final, 1) = "=NOW()"
                                  Hoja83.Cells(Final, 1).Copy
                                  Hoja83.Cells(Final, 1).PasteSpecial Paste:=xlPasteValues
                                  Application.CutCopyMode = False
                                  
                                  Hoja83.Cells(Final, 2) = Me.txt_Usuario.Text
                                  
                                  
                                  Hoja0.lbl_usuario.Caption = "Usuario actual: " & UCase(Me.txt_Usuario.Text)
                                  
                                  Hoja83.Cells(Final, 3) = Status
                                  
                    
                                 
                                  Hoja83.Range("G1") = Me.txt_Usuario.Text
                                  Hoja83.Range("H1") = Status
                                  
                                  Hoja83.Protect (Seguridad)
                                  
                                  'ThisWorkbook.Save
                              
                              
                                  Unload Me
                                  Hoja0.Activate
                        Else
                     MsgBox "La contraseña es incorrecta", vbExclamation, Titulo
            End If
End If

Application.ScreenUpdating = True


End Sub
Private Sub btn_salir_Click()
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

