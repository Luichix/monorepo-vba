VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Iniciosesion 
   Caption         =   "GESTOR ADMINISTRATIVO"
   ClientHeight    =   3465
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   9340.001
   OleObjectBlob   =   "frm_Iniciosesion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Iniciosesion"
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
Dim vHoja(91) As String
Dim vBoton(12) As String
Dim i As Byte
Dim X As Byte

Application.ScreenUpdating = False

Titulo = "Gestor de Inventarios"


yaExiste = Application.WorksheetFunction.CountIf(Hoja9.Range("tbl_Usuario[Usuario]"), Me.txt_Usuario.Text)
Set Rango = Hoja9.Range("tbl_Usuario[Usuario]")

If Me.txt_Usuario.Text = "" Or Me.txt_Contraseña.Text = "" Then
    MsgBox "Introduce usuario y contraseña", vbExclamation, Titulo
    Me.txt_Usuario.SetFocus

            ElseIf yaExiste = 0 Then
                MsgBox "El usuario '" & Me.txt_Usuario.Text & "' no existe", vbExclamation, Titulo
            
            ElseIf yaExiste = 1 Then
                UsuarioEncontrado = Rango.Find(What:=Me.txt_Usuario.Text, after:=Rango.Range("A1"), _
                                                LookAt:=xlWhole, MatchCase:=False).Address
                
                password = Hoja9.Range(UsuarioEncontrado).Offset(0, 1).Value
               
                
                'Permisos y restricciones
                vHoja(1) = Hoja9.Range(UsuarioEncontrado).Offset(0, 2).Value
                vHoja(2) = Hoja9.Range(UsuarioEncontrado).Offset(0, 3).Value
                vHoja(3) = Hoja9.Range(UsuarioEncontrado).Offset(0, 3).Value
                vHoja(4) = Hoja9.Range(UsuarioEncontrado).Offset(0, 3).Value
                vHoja(5) = Hoja9.Range(UsuarioEncontrado).Offset(0, 3).Value
                vHoja(6) = Hoja9.Range(UsuarioEncontrado).Offset(0, 3).Value
                vHoja(7) = Hoja9.Range(UsuarioEncontrado).Offset(0, 3).Value
                vHoja(8) = Hoja9.Range(UsuarioEncontrado).Offset(0, 3).Value
                vHoja(9) = Hoja9.Range(UsuarioEncontrado).Offset(0, 2).Value
                vHoja(91) = Hoja9.Range(UsuarioEncontrado).Offset(0, 2).Value
                 
                vBoton(1) = Hoja9.Range(UsuarioEncontrado).Offset(0, 4).Value
                vBoton(2) = Hoja9.Range(UsuarioEncontrado).Offset(0, 4).Value
                vBoton(3) = Hoja9.Range(UsuarioEncontrado).Offset(0, 4).Value
                vBoton(4) = Hoja9.Range(UsuarioEncontrado).Offset(0, 4).Value
                vBoton(5) = Hoja9.Range(UsuarioEncontrado).Offset(0, 4).Value
                vBoton(6) = Hoja9.Range(UsuarioEncontrado).Offset(0, 4).Value
                vBoton(7) = Hoja9.Range(UsuarioEncontrado).Offset(0, 4).Value
                vBoton(8) = Hoja9.Range(UsuarioEncontrado).Offset(0, 4).Value
                vBoton(9) = Hoja9.Range(UsuarioEncontrado).Offset(0, 4).Value
                vBoton(10) = Hoja9.Range(UsuarioEncontrado).Offset(0, 4).Value
                vBoton(11) = Hoja9.Range(UsuarioEncontrado).Offset(0, 4).Value
                vBoton(12) = Hoja9.Range(UsuarioEncontrado).Offset(0, 4).Value
                

                
                
            If Hoja9.Range(UsuarioEncontrado).Value = Me.txt_Usuario.Text And password = Me.txt_Contraseña.Text Then
            
                        'Validando los permisos y restricciones en las hojas de cálculo
                        For i = 1 To 91
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
                     
   
                        For X = 1 To 12
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

                              Final = GetNuevoR(Hoja91)
                                  Hoja91.Cells(Final, 1) = "=NOW()"
                                  Hoja91.Cells(Final, 1).Copy
                                  Hoja91.Cells(Final, 1).PasteSpecial Paste:=xlPasteValues
                                  Application.CutCopyMode = False

                                  Hoja91.Cells(Final, 2) = Me.txt_Usuario.Text

                                  Hoja0.txt_UsuarioActual.Caption = "Usuario actual: " & UCase(Me.txt_Usuario.Text)

                                 



                                  Hoja91.Range("G1") = Me.txt_Usuario.Text
                                 

                                  
                                  
'                                  ThisWorkbook.Save
                              
                              
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



