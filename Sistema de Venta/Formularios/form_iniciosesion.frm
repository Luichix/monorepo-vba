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
Dim vHoja(94) As String
Dim vBoton(54) As String
Dim i As Byte
Dim ii As Byte
Dim X As Byte

Application.ScreenUpdating = False

Titulo = "Gestor de Inventarios"


yaExiste = Application.WorksheetFunction.CountIf(Hoja91.Range("tbl_Usuario[Usuario]"), Me.txt_usuario.Text)
Set Rango = Hoja91.Range("tbl_Usuario[Usuario]")

If Me.txt_usuario.Text = "" Or Me.txt_Contraseña.Text = "" Then
    MsgBox "Introduce usuario y contraseña", vbExclamation, Titulo
    Me.txt_usuario.SetFocus

            ElseIf yaExiste = 0 Then
                MsgBox "El usuario '" & Me.txt_usuario.Text & "' no existe", vbExclamation, Titulo
            
            ElseIf yaExiste = 1 Then
                UsuarioEncontrado = Rango.Find(What:=Me.txt_usuario.Text, After:=Rango.Range("A1"), _
                                                LookAt:=xlWhole, MatchCase:=False).Address
                
                password = Hoja91.Range(UsuarioEncontrado).Offset(0, 1).Value
                Status = Hoja91.Range(UsuarioEncontrado).Offset(0, 2).Value
                
                'Permisos y restricciones
                vHoja(1) = Hoja91.Range(UsuarioEncontrado).Offset(0, 3).Value
                vHoja(10) = Hoja91.Range(UsuarioEncontrado).Offset(0, 4).Value
                vHoja(11) = Hoja91.Range(UsuarioEncontrado).Offset(0, 5).Value
                vHoja(12) = Hoja91.Range(UsuarioEncontrado).Offset(0, 6).Value
                vHoja(13) = Hoja91.Range(UsuarioEncontrado).Offset(0, 7).Value
                vHoja(14) = Hoja91.Range(UsuarioEncontrado).Offset(0, 8).Value
                vHoja(2) = Hoja91.Range(UsuarioEncontrado).Offset(0, 9).Value
                vHoja(21) = Hoja91.Range(UsuarioEncontrado).Offset(0, 10).Value
                vHoja(22) = Hoja91.Range(UsuarioEncontrado).Offset(0, 11).Value
                vHoja(23) = Hoja91.Range(UsuarioEncontrado).Offset(0, 12).Value
                vHoja(24) = Hoja91.Range(UsuarioEncontrado).Offset(0, 13).Value
                vHoja(25) = Hoja91.Range(UsuarioEncontrado).Offset(0, 14).Value
                vHoja(26) = Hoja91.Range(UsuarioEncontrado).Offset(0, 15).Value
                vHoja(27) = Hoja91.Range(UsuarioEncontrado).Offset(0, 16).Value
                vHoja(28) = Hoja91.Range(UsuarioEncontrado).Offset(0, 17).Value
                vHoja(29) = Hoja91.Range(UsuarioEncontrado).Offset(0, 18).Value
                vHoja(3) = Hoja91.Range(UsuarioEncontrado).Offset(0, 19).Value
                vHoja(30) = Hoja91.Range(UsuarioEncontrado).Offset(0, 20).Value
                vHoja(31) = Hoja91.Range(UsuarioEncontrado).Offset(0, 21).Value
                vHoja(32) = Hoja91.Range(UsuarioEncontrado).Offset(0, 22).Value
                vHoja(4) = Hoja91.Range(UsuarioEncontrado).Offset(0, 23).Value
                vHoja(5) = Hoja91.Range(UsuarioEncontrado).Offset(0, 24).Value
                vHoja(6) = Hoja91.Range(UsuarioEncontrado).Offset(0, 25).Value
                vHoja(7) = Hoja91.Range(UsuarioEncontrado).Offset(0, 26).Value
                vHoja(8) = Hoja91.Range(UsuarioEncontrado).Offset(0, 27).Value
                vHoja(9) = Hoja91.Range(UsuarioEncontrado).Offset(0, 28).Value
                vHoja(91) = Hoja91.Range(UsuarioEncontrado).Offset(0, 29).Value
                vHoja(92) = Hoja91.Range(UsuarioEncontrado).Offset(0, 30).Value
                vHoja(93) = Hoja91.Range(UsuarioEncontrado).Offset(0, 31).Value
                vHoja(94) = Hoja91.Range(UsuarioEncontrado).Offset(0, 32).Value
               
                 
                 
                vBoton(1) = Hoja91.Range(UsuarioEncontrado).Offset(0, 60).Value
                vBoton(2) = Hoja91.Range(UsuarioEncontrado).Offset(0, 61).Value
                vBoton(3) = Hoja91.Range(UsuarioEncontrado).Offset(0, 62).Value
                vBoton(4) = Hoja91.Range(UsuarioEncontrado).Offset(0, 63).Value
                vBoton(5) = Hoja91.Range(UsuarioEncontrado).Offset(0, 64).Value
                vBoton(6) = Hoja91.Range(UsuarioEncontrado).Offset(0, 65).Value
                vBoton(7) = Hoja91.Range(UsuarioEncontrado).Offset(0, 66).Value
                vBoton(8) = Hoja91.Range(UsuarioEncontrado).Offset(0, 67).Value
                vBoton(9) = Hoja91.Range(UsuarioEncontrado).Offset(0, 68).Value
                vBoton(10) = Hoja91.Range(UsuarioEncontrado).Offset(0, 69).Value
                vBoton(11) = Hoja91.Range(UsuarioEncontrado).Offset(0, 70).Value
                vBoton(12) = Hoja91.Range(UsuarioEncontrado).Offset(0, 71).Value
                vBoton(13) = Hoja91.Range(UsuarioEncontrado).Offset(0, 72).Value
                vBoton(14) = Hoja91.Range(UsuarioEncontrado).Offset(0, 73).Value
                vBoton(15) = Hoja91.Range(UsuarioEncontrado).Offset(0, 74).Value
                vBoton(16) = Hoja91.Range(UsuarioEncontrado).Offset(0, 75).Value
                vBoton(17) = Hoja91.Range(UsuarioEncontrado).Offset(0, 76).Value
                vBoton(18) = Hoja91.Range(UsuarioEncontrado).Offset(0, 77).Value
                vBoton(19) = Hoja91.Range(UsuarioEncontrado).Offset(0, 78).Value
                vBoton(20) = Hoja91.Range(UsuarioEncontrado).Offset(0, 79).Value
                vBoton(21) = Hoja91.Range(UsuarioEncontrado).Offset(0, 80).Value
                vBoton(22) = Hoja91.Range(UsuarioEncontrado).Offset(0, 81).Value
                vBoton(23) = Hoja91.Range(UsuarioEncontrado).Offset(0, 82).Value
                vBoton(24) = Hoja91.Range(UsuarioEncontrado).Offset(0, 83).Value
                vBoton(25) = Hoja91.Range(UsuarioEncontrado).Offset(0, 84).Value
                vBoton(26) = Hoja91.Range(UsuarioEncontrado).Offset(0, 85).Value
                vBoton(27) = Hoja91.Range(UsuarioEncontrado).Offset(0, 86).Value
                vBoton(28) = Hoja91.Range(UsuarioEncontrado).Offset(0, 87).Value
                vBoton(29) = Hoja91.Range(UsuarioEncontrado).Offset(0, 88).Value
                vBoton(30) = Hoja91.Range(UsuarioEncontrado).Offset(0, 89).Value
                vBoton(31) = Hoja91.Range(UsuarioEncontrado).Offset(0, 90).Value
                vBoton(32) = Hoja91.Range(UsuarioEncontrado).Offset(0, 91).Value
                vBoton(33) = Hoja91.Range(UsuarioEncontrado).Offset(0, 92).Value
                vBoton(34) = Hoja91.Range(UsuarioEncontrado).Offset(0, 93).Value
                vBoton(35) = Hoja91.Range(UsuarioEncontrado).Offset(0, 94).Value
                vBoton(36) = Hoja91.Range(UsuarioEncontrado).Offset(0, 95).Value
                vBoton(37) = Hoja91.Range(UsuarioEncontrado).Offset(0, 96).Value
                vBoton(38) = Hoja91.Range(UsuarioEncontrado).Offset(0, 97).Value
                vBoton(39) = Hoja91.Range(UsuarioEncontrado).Offset(0, 98).Value
                vBoton(40) = Hoja91.Range(UsuarioEncontrado).Offset(0, 99).Value
                vBoton(41) = Hoja91.Range(UsuarioEncontrado).Offset(0, 100).Value
                vBoton(42) = Hoja91.Range(UsuarioEncontrado).Offset(0, 101).Value
                vBoton(43) = Hoja91.Range(UsuarioEncontrado).Offset(0, 102).Value
                vBoton(44) = Hoja91.Range(UsuarioEncontrado).Offset(0, 103).Value
                vBoton(45) = Hoja91.Range(UsuarioEncontrado).Offset(0, 104).Value
                vBoton(46) = Hoja91.Range(UsuarioEncontrado).Offset(0, 105).Value
                vBoton(47) = Hoja91.Range(UsuarioEncontrado).Offset(0, 106).Value
                vBoton(48) = Hoja91.Range(UsuarioEncontrado).Offset(0, 107).Value
                vBoton(49) = Hoja91.Range(UsuarioEncontrado).Offset(0, 108).Value
                vBoton(50) = Hoja91.Range(UsuarioEncontrado).Offset(0, 109).Value
                vBoton(51) = Hoja91.Range(UsuarioEncontrado).Offset(0, 110).Value
                vBoton(52) = Hoja91.Range(UsuarioEncontrado).Offset(0, 111).Value
                vBoton(53) = Hoja91.Range(UsuarioEncontrado).Offset(0, 113).Value
                vBoton(54) = Hoja91.Range(UsuarioEncontrado).Offset(0, 114).Value

                
                
            If Hoja91.Range(UsuarioEncontrado).Value = Me.txt_usuario.Text And password = Me.txt_Contraseña.Text Then
            
                        'Validando los permisos y restricciones en las hojas de cálculo
                        For i = 1 To 94
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
                     
   
                        For X = 1 To 54
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
                        
     
                        ' Registrar al usuario en la hoja Logs
                              
                              Final = GetNuevoR(Hoja92)
                                  Hoja92.Cells(Final, 1) = "=NOW()"
                                  Hoja92.Cells(Final, 1).Copy
                                  Hoja92.Cells(Final, 1).PasteSpecial Paste:=xlPasteValues
                                  Application.CutCopyMode = False
                                  
                                  Hoja92.Cells(Final, 2) = Me.txt_usuario.Text
                                  
                                  Hoja0.txt_UsuarioActual.Caption = "Usuario actual: " & UCase(Me.txt_usuario.Text)
                                  
                                  Hoja92.Cells(Final, 3) = Status
                                  
                    
                                 
                                  Hoja92.Range("G1") = Me.txt_usuario.Text
                                  Hoja92.Range("H1") = Status
                                  
                                    If Hoja92.Range("H1") = "USUARIO" Then
                                        ActiveWindow.DisplayWorkbookTabs = False
                                    ElseIf Hoja92.Range("H1") = "ADMINISTRADOR" Then
                                        ActiveWindow.DisplayWorkbookTabs = True
                                    End If
                                  
                              Application.EnableEvents = False
                                  ThisWorkbook.Save
                              Application.EnableEvents = True
                              
                                  Unload Me
                                  Hoja0.Activate
                        Else
                     MsgBox "La contraseña es incorrecta", vbExclamation, Titulo
            End If
End If

Application.ScreenUpdating = True




End Sub
Private Sub btn_Salir_Click()
    Unload Me
    Application.EnableEvents = False
     ThisWorkbook.Save
     Application.EnableEvents = True
     ThisWorkbook.Application.DisplayAlerts = False
    Application.ActiveWorkbook.Close
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Unload Me
        
        ThisWorkbook.Application.DisplayAlerts = False
        ThisWorkbook.Save
 
        Application.ActiveWorkbook.Close
    End If
End Sub

