VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_hora 
   Caption         =   "CONTROL DE ENTRADAS Y SALIDAS"
   ClientHeight    =   9735.001
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   7600
   OleObjectBlob   =   "frm_hora.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_hora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Private Sub btn_Calendario_Click()
Dim Seguridad As String

    If Me.txt_id.Text = "" Then
        MsgBox "Seleccione un Codigo de Personal", vbInformation, "Gestor de Personal"
        Exit Sub
    Else
    

Seguridad = Hoja83.Range("L1").Text

Hoja58.Unprotect (Seguridad)
Hoja58.Cells(6, 11) = Me.txt_id.Text
        frm_Calendario_Asistencia.Show
Hoja58.Protect (Seguridad)
    End If
End Sub
Private Sub btn_personal_Click()
Me.txt_id.BackColor = &H80000005
Me.txt_xEntrada.SetFocus
banderaPersonal = 1
Call LanzarListadoPersonal(Me, "btn_Fecha_Horas")
End Sub
Private Sub btn_Fecha_Horas_Click()
Me.txt_Fecha.BackColor = &H80000005
Me.txt_xEntrada.SetFocus
banderaCalendario = 3
  Call LanzarCalendario(Me, "txt_fecha")
End Sub

Private Sub btn_Registrar_Click()
Dim Titulo As String
Dim Formato As String
Dim LEntrada As Date
Dim LSalida As Date
Dim NEntrada As Date
Dim NSalida As Date

On Error GoTo Salir


Titulo = "Gestion del Personal"
Formato = "00:00"


If Me.txt_Fecha.Text = "" Then
    Me.txt_Fecha.BackColor = &HC0C0FF
    MsgBox "Ingrese la fecha de labores..!", vbInformation, Titulo
    Me.txt_Fecha.SetFocus
    Exit Sub
End If

        If Me.txt_id.Text = "" Then
            Me.txt_id.BackColor = &HC0C0FF
            Me.txt_nombre.BackColor = &HC0C0FF
            MsgBox "Seleccione un colaborador del listado..!", vbInformation, Titulo
            Me.txt_id.SetFocus
            Exit Sub
        End If
                        
                        
                        If ckx_inicio.Value = True Then
                                If Me.txt_xEntrada.Text = "" Or Me.txt_xEntrada.Text = Formato Then
                                    Me.txt_xEntrada.BackColor = &HC0C0FF
                                    MsgBox "Ingrese las horas correctamente", vbInformation, Titulo
                                    Me.txt_xEntrada.SetFocus
                                    Exit Sub
                                End If
                        End If
                                
                        If ckx_inicio.Value = True Then
                                If Me.txt_xSalida.Text = "" Or Me.txt_xSalida.Text = Formato Then
                                    Me.txt_xSalida.BackColor = &HC0C0FF
                                    MsgBox "Ingrese las horas correctamente...!", vbInformation, Titulo
                                    Me.txt_xSalida.SetFocus
                                    Exit Sub
                                End If
                         End If
                                
                         If ckx_final.Value = True Then
                                If Me.txt_yEntrada.Text = "" Or Me.txt_yEntrada.Text = Formato Then
                                    Me.txt_yEntrada.BackColor = &HC0C0FF
                                    MsgBox "Ingrese la horas correctamente..!", vbInformation, Titulo
                                    Me.txt_yEntrada.SetFocus
                                    Exit Sub
                                End If
                        End If
                        
                        If ckx_final.Value = True Then
                                If Me.txt_ySalida.Text = "" Or Me.txt_ySalida.Text = Formato Then
                                    Me.txt_ySalida.BackColor = &HC0C0FF
                                    MsgBox "Ingrese las horas correctamente", vbInformation, Titulo
                                    Me.txt_ySalida.SetFocus
                                    Exit Sub
                                End If
                        End If
                                
    
LEntrada = Me.txt_xEntrada.Value
LSalida = Me.txt_xSalida.Value
NEntrada = Me.txt_yEntrada.Value
NSalida = Me.txt_ySalida.Value

                        If ckx_inicio.Value = True Then
                        If LEntrada >= LSalida Then
                            Me.txt_xEntrada.BackColor = &HC0C0FF
                            Me.txt_xSalida.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada.SetFocus
                            Exit Sub
                        End If
                        End If
                        
                        If ckx_final.Value = True And ckx_inicio.Value = True Then
                        If LSalida >= NEntrada Then
                            Me.txt_xEntrada.BackColor = &HC0C0FF
                            Me.txt_xSalida.BackColor = &HC0C0FF
                            Me.txt_yEntrada.BackColor = &HC0C0FF
                            Me.txt_ySalida.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada.SetFocus
                            Exit Sub
                        End If
                        End If
                        
                        If ckx_final.Value = True Then
                        If NEntrada >= NSalida Then
                            Me.txt_yEntrada.BackColor = &HC0C0FF
                            Me.txt_ySalida.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_yEntrada.SetFocus
                            Exit Sub
                        End If
                        End If

    
If MsgBox("¿Son correctos los datos?", vbYesNo, Titulo) = vbNo Then
        Exit Sub
    Else

        Registrar_Hora
        Limpiar_Hora

End If

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Recursos Humanos"
 End If

End Sub
Private Sub Limpiar_Hora()
Dim Formato As String
Formato = "00:00"
    Me.txt_xEntrada.Text = Formato
    Me.txt_xSalida.Text = Formato
    Me.txt_yEntrada.Text = Formato
    Me.txt_ySalida.Text = Formato
End Sub
Private Sub Registrar_Hora()
Dim Fecha As Date
Dim Titulo As String
Dim xEnter As Date
Dim xExit As Date
Dim yEnter As Date
Dim yExit As Date
Dim Seguridad As String

Seguridad = Hoja83.Range("L1").Text

Hoja2.Unprotect (Seguridad)

Titulo = "Gestor de Recursos Humanos"
    
Fecha = Me.txt_Fecha.Text
xEnter = Me.txt_xEntrada.Value
xExit = Me.txt_xSalida.Value
yEnter = Me.txt_yEntrada.Value
yExit = Me.txt_ySalida.Value
            
            If Me.ckx_final.Value = False Then
            
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora.txt_Fecha)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter
                Hoja2.Cells(3, 6) = xExit
            Else
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora.txt_Fecha)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter
                Hoja2.Cells(3, 6) = xExit
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora.txt_Fecha)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter
                Hoja2.Cells(3, 6) = yExit
                
            End If
 Hoja2.Protect (Seguridad)
 
         MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
             
End Sub
Private Sub Sombra()
        Me.txt_xEntrada.BackColor = &H80000005
        Me.txt_xSalida.BackColor = &H80000005
        Me.txt_yEntrada.BackColor = &H80000005
        Me.txt_ySalida.BackColor = &H80000005
End Sub
Private Sub btn_salir_Click()
Unload Me
End Sub

Private Sub ckx_final_Click()
    If ckx_final.Value = True Then
        Me.txt_yEntrada.Enabled = True
        Me.txt_ySalida.Enabled = True
        ckx_inicio.Enabled = True
        Me.txt_yEntrada.Text = "00:00"
        Me.txt_ySalida.Text = "00:00"
        Me.txt_yEntrada.BackColor = &H80000005
        Me.txt_ySalida.BackColor = &H80000005
    ElseIf ckx_final.Value = False Then
        Me.txt_yEntrada.Enabled = False
        Me.txt_ySalida.Enabled = False
        ckx_inicio.Enabled = False
        Me.txt_yEntrada.Text = "00:00"
        Me.txt_ySalida.Text = "00:00"
        Me.txt_yEntrada.BackColor = &H80000005
        Me.txt_ySalida.BackColor = &H80000005
    End If
    
End Sub

Private Sub ckx_inicio_Click()
    
    If ckx_inicio.Value = True Then
        Me.txt_xEntrada.Enabled = True
        Me.txt_xSalida.Enabled = True
        Me.txt_xEntrada.Text = "00:00"
        Me.txt_xSalida.Text = "00:00"
        ckx_final.Enabled = True
        Me.txt_xEntrada.BackColor = &H80000005
        Me.txt_xSalida.BackColor = &H80000005
        
    ElseIf ckx_inicio.Value = False Then
        Me.txt_xEntrada.Enabled = False
        Me.txt_xSalida.Enabled = False
        ckx_final.Enabled = False
        Me.txt_xEntrada.Text = "00:00"
        Me.txt_xSalida.Text = "00:00"
        Me.txt_xEntrada.BackColor = &H80000005
        Me.txt_xSalida.BackColor = &H80000005
    End If
    
End Sub

Private Sub txt_xEntrada_Change()

Sombra

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub txt_xEntrada_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    
If txt_xEntrada.Text <> "1" And txt_xEntrada.Text <> "2" And txt_xEntrada.Text <> "3" And txt_xEntrada.Text <> "4" And txt_xEntrada.Text <> "0" Then
    
    Select Case Len(txt_xEntrada.Value)
        Case 1
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 4
          End Select
        
    End If

If txt_xEntrada.Value = 10 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If
If txt_xEntrada.Value = 11 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If
If txt_xEntrada.Value = 12 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If
If txt_xEntrada.Value = 13 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If
If txt_xEntrada.Value = 14 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If
If txt_xEntrada.Value = 15 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If
If txt_xEntrada.Value = 16 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If
If txt_xEntrada.Value = 17 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If
If txt_xEntrada.Value = 18 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If
If txt_xEntrada.Value = 19 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If
If txt_xEntrada.Value = 20 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If
If txt_xEntrada.Value = 21 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If
If txt_xEntrada.Value = 22 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If
If txt_xEntrada.Value = 23 Then
    Select Case Len(txt_xEntrada.Value)
        Case 2
        txt_xEntrada.Value = txt_xEntrada.Value & ":"
        Me.txt_xEntrada.MaxLength = 5
        End Select

End If


If txt_xEntrada.Value = 25 Or txt_xEntrada.Value = 24 Or txt_xEntrada.Value = 0 Or txt_xEntrada.Value = 26 Or txt_xEntrada.Value = 27 Or txt_xEntrada.Value = 28 Or txt_xEntrada.Value = 29 Or txt_xEntrada.Value = 3 Or txt_xEntrada.Value = 4 Then
    txt_xEntrada = "00:00"
     Me.txt_xEntrada.MaxLength = 4
End If

End Sub

Private Sub txt_xEntrada_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 47 And KeyAscii < 59 Then
 
KeyAscii = KeyAscii
 
Else
KeyAscii = 0
 
End If
End Sub

Private Sub txt_xEntrada_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyBack Then
Me.txt_xEntrada = Empty
End If
If KeyCode = vbKeyDelete Then
Me.txt_xEntrada = Empty
End If
End Sub

Private Sub txt_xSalida_Change()

Sombra
End Sub

Private Sub txt_xSalida_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    
If txt_xSalida.Text <> "1" And txt_xSalida.Text <> "2" And txt_xSalida.Text <> "3" And txt_xSalida.Text <> "4" And txt_xSalida.Text <> "0" Then
    
    Select Case Len(txt_xSalida.Value)
        Case 1
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 4
          End Select
        
    End If

If txt_xSalida.Value = 10 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If
If txt_xSalida.Value = 11 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If
If txt_xSalida.Value = 12 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If
If txt_xSalida.Value = 13 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If
If txt_xSalida.Value = 14 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If
If txt_xSalida.Value = 15 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If
If txt_xSalida.Value = 16 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If
If txt_xSalida.Value = 17 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If
If txt_xSalida.Value = 18 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If
If txt_xSalida.Value = 19 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If
If txt_xSalida.Value = 20 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If
If txt_xSalida.Value = 21 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If
If txt_xSalida.Value = 22 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If
If txt_xSalida.Value = 23 Then
    Select Case Len(txt_xSalida.Value)
        Case 2
        txt_xSalida.Value = txt_xSalida.Value & ":"
        Me.txt_xSalida.MaxLength = 5
        End Select

End If


If txt_xSalida.Value = 25 Or txt_xSalida.Value = 24 Or txt_xSalida.Value = 0 Or txt_xSalida.Value = 26 Or txt_xSalida.Value = 27 Or txt_xSalida.Value = 28 Or txt_xSalida.Value = 29 Or txt_xSalida.Value = 3 Or txt_xSalida.Value = 4 Then
    txt_xSalida = "00:00"
     Me.txt_xSalida.MaxLength = 4
End If

End Sub




Private Sub txt_xSalida_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 47 And KeyAscii < 59 Then
 
KeyAscii = KeyAscii
 
Else
KeyAscii = 0
 
End If
End Sub

Private Sub txt_xSalida_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyBack Then
Me.txt_xSalida = Empty
End If
If KeyCode = vbKeyDelete Then
Me.txt_xSalida = Empty
End If
End Sub

Private Sub txt_yEntrada_Change()
Sombra
End Sub

Private Sub txt_yEntrada_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    
If txt_yEntrada.Text <> "1" And txt_yEntrada.Text <> "2" And txt_yEntrada.Text <> "3" And txt_yEntrada.Text <> "4" And txt_yEntrada.Text <> "0" Then
    
    Select Case Len(txt_yEntrada.Value)
        Case 1
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 4
          End Select
        
    End If

If txt_yEntrada.Value = 10 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If
If txt_yEntrada.Value = 11 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If
If txt_yEntrada.Value = 12 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If
If txt_yEntrada.Value = 13 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If
If txt_yEntrada.Value = 14 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If
If txt_yEntrada.Value = 15 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If
If txt_yEntrada.Value = 16 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If
If txt_yEntrada.Value = 17 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If
If txt_yEntrada.Value = 18 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If
If txt_yEntrada.Value = 19 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If
If txt_yEntrada.Value = 20 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If
If txt_yEntrada.Value = 21 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If
If txt_yEntrada.Value = 22 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If
If txt_yEntrada.Value = 23 Then
    Select Case Len(txt_yEntrada.Value)
        Case 2
        txt_yEntrada.Value = txt_yEntrada.Value & ":"
        Me.txt_yEntrada.MaxLength = 5
        End Select

End If


If txt_yEntrada.Value = 25 Or txt_yEntrada.Value = 24 Or txt_yEntrada.Value = 0 Or txt_yEntrada.Value = 26 Or txt_yEntrada.Value = 27 Or txt_yEntrada.Value = 28 Or txt_yEntrada.Value = 29 Or txt_yEntrada.Value = 3 Or txt_yEntrada.Value = 4 Then
    txt_yEntrada = "00:00"
     Me.txt_yEntrada.MaxLength = 4
End If

End Sub




Private Sub txt_yEntrada_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 47 And KeyAscii < 59 Then
 
KeyAscii = KeyAscii
 
Else
KeyAscii = 0
 
End If

End Sub

Private Sub txt_yEntrada_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyBack Then
Me.txt_yEntrada = Empty
End If
If KeyCode = vbKeyDelete Then
Me.txt_yEntrada = Empty
End If
End Sub

Private Sub txt_ySalida_Change()

Sombra
End Sub

Private Sub txt_ySalida_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    
If txt_ySalida.Text <> "1" And txt_ySalida.Text <> "2" And txt_ySalida.Text <> "3" And txt_ySalida.Text <> "4" And txt_ySalida.Text <> "0" Then
    
    Select Case Len(txt_ySalida.Value)
        Case 1
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 4
          End Select
        
    End If

If txt_ySalida.Value = 10 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If
If txt_ySalida.Value = 11 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If
If txt_ySalida.Value = 12 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If
If txt_ySalida.Value = 13 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If
If txt_ySalida.Value = 14 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If
If txt_ySalida.Value = 15 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If
If txt_ySalida.Value = 16 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If
If txt_ySalida.Value = 17 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If
If txt_ySalida.Value = 18 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If
If txt_ySalida.Value = 19 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If
If txt_ySalida.Value = 20 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If
If txt_ySalida.Value = 21 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If
If txt_ySalida.Value = 22 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If
If txt_ySalida.Value = 23 Then
    Select Case Len(txt_ySalida.Value)
        Case 2
        txt_ySalida.Value = txt_ySalida.Value & ":"
        Me.txt_ySalida.MaxLength = 5
        End Select

End If


If txt_ySalida.Value = 25 Or txt_ySalida.Value = 24 Or txt_ySalida.Value = 0 Or txt_ySalida.Value = 26 Or txt_ySalida.Value = 27 Or txt_ySalida.Value = 28 Or txt_ySalida.Value = 29 Or txt_ySalida.Value = 3 Or txt_ySalida.Value = 4 Then
    txt_ySalida = "00:00"
     Me.txt_ySalida.MaxLength = 4
End If

End Sub




Private Sub txt_ySalida_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 47 And KeyAscii < 59 Then
 
KeyAscii = KeyAscii
 
Else
KeyAscii = 0
 
End If

End Sub

Private Sub txt_ySalida_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyBack Then
Me.txt_ySalida = Empty
End If
If KeyCode = vbKeyDelete Then
Me.txt_ySalida = Empty
End If
End Sub



Private Sub UserForm_Click()

End Sub
