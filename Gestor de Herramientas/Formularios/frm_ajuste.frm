VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ajuste 
   Caption         =   "Gestor"
   ClientHeight    =   6780
   ClientLeft      =   70
   ClientTop       =   300
   ClientWidth     =   6250
   OleObjectBlob   =   "frm_ajuste.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ajuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Fecha_Click()
 banderaCalendario = 5
    Call LanzarCalendario(Me, "txt_Fecha")
End Sub
Private Sub btn_modificar_Click()
On Error GoTo Salir

Application.ScreenUpdating = False

    If Me.txt_Fecha.Text = Empty Or _
        Me.txt_item.Text = Empty Or _
        Me.txt_pieza.Text = Empty Or _
        Me.txt_estado.Text = Empty Or _
        Me.txt_detalle.Text = Empty Or _
        Me.txt_cantidad.Text = Empty Then
                If Me.txt_numero.Text = Empty Then
                    MsgBox "Notifique al programador, error en la estructura de datos", vbCritical, "luisreynaldo.pch@gmail.com"
                    Exit Sub
                End If
            MsgBox "Hay campos vacíos en el registro", , "Gestor de Inventario de Herramientas"
            Exit Sub
        
    End If
    
If MsgBox("Son correctos los datos?" + Chr(13) + "Desea procesar el registro?", vbYesNo, "Gestor de Inventarios") = vbNo Then
        Exit Sub
    Else

        Modificar
End If

Hoja0.Activate
Hoja0.Select
     Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventario de Herramientas"
 End If

End Sub
Private Sub Modificar()
Dim x As String

x = Me.txt_numero.Text

Hoja11.Select
Range("A1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like x And ActiveCell.Offset(0, 2) Like Me.txt_item Then
            
            If ActiveCell.Offset(0, 9) = "" And Me.txt_detalle.Text <> "Bueno" Then
                    If Me.txt_estado.Text = "Inactivo" Then
                        ActiveCell.Offset(0, 7) = Me.txt_estado.Text
                        ActiveCell.Offset(0, 8) = Me.txt_detalle.Text
                        ActiveCell.Offset(0, 10) = CDate(Me.txt_Fecha)
                        MsgBox "Registro ha sido inhabilitado correctamente..!", vbInformation, "Gestor"
                    Else
                        ActiveCell.Offset(0, 7) = Me.txt_estado.Text
                        ActiveCell.Offset(0, 8) = Me.txt_detalle.Text
                        ActiveCell.Offset(0, 9) = CDate(Me.txt_Fecha)
                        MsgBox "Registro ha sido modificado Correctamente..!", vbInformation, "Gestor"
                    End If
                    Unload Me
                    ThisWorkbook.Save
                    Exit Do
                    Exit Sub
                    
            ElseIf ActiveCell.Offset(0, 9) <> "" Then
                ActiveCell.Offset(0, 7) = Me.txt_estado.Text
                If Me.txt_estado.Text = "Inactivo" Then
                    ActiveCell.Offset(0, 7) = Me.txt_estado.Text
                    ActiveCell.Offset(0, 10) = CDate(Me.txt_Fecha)
                    MsgBox "Registro ha sido inhabilitado correctamente..!", vbInformation, "Gestor"
                    Unload Me
                    ThisWorkbook.Save
                Else
                    Me.txt_detalle.BackColor = &H8080FF
                    MsgBox "No se puede modificar el detalle del registro.!", vbExclamation, "Gestor"
                End If
                    Exit Do
                    Exit Sub
            ElseIf Me.txt_estado.Text = "Inactivo" Then
                        ActiveCell.Offset(0, 7) = Me.txt_estado.Text
                        ActiveCell.Offset(0, 10) = CDate(Me.txt_Fecha)
                        MsgBox "Registro ha sido inhabilitado correctamente..!", vbInformation, "Gestor"
                        Unload Me
                        ThisWorkbook.Save
                        Exit Do
                        Exit Sub
            Else
                MsgBox "No se ha modificado el registro.!", vbExclamation, "Gestor"
                Exit Do
                Exit Sub
             
            End If
            
        End If
    Loop
    
End Sub
Private Sub txt_cancelar_Click()
Unload Me
End Sub
Private Sub txt_cantidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 48 And KeyAscii < 58 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub
