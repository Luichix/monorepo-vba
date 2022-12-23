VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_cajon 
   Caption         =   "Cajas de Herramientas"
   ClientHeight    =   7488
   ClientLeft      =   70
   ClientTop       =   300
   ClientWidth     =   8550.001
   OleObjectBlob   =   "frm_cajon.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_cajon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_caja_Click()
 banderaCaja = 2
    Call LanzarListadoCaja(Me, "lbl_Fecha")
End Sub

Private Sub btn_cancelar_Click()
Unload Me
End Sub

Private Sub btn_Fecha_Click()
 banderaCalendario = 3
    Call LanzarCalendario(Me, "txt_Fecha")
    Me.txt_caja.SetFocus
End Sub

Private Sub btn_modificar_Click()
On Error GoTo Salir

Application.ScreenUpdating = False

    If Me.txt_Fecha.Text = Empty Or _
        Me.txt_caja.Text = Empty Or _
        Me.txt_id.Text = Empty Or _
        Me.txt_personal.Text = Empty Or _
        Me.txt_puesto.Text = Empty Or _
        Me.txt_area.Text = Empty Or _
        Me.txt_estado.Text = Empty Or _
        Me.txt_activo.Text = Empty Or _
        Me.txt_observacion.Text = Empty Then

            MsgBox "Hay campos vacíos en el registro", , "Gestor de Inventario de Herramientas"
            Exit Sub
    
    End If
    
If MsgBox("Son correctos los datos a modificar?" + Chr(13) + "Desea procesar el registro?", vbYesNo, "Gestor de Inventarios") = vbNo Then
        Exit Sub
    Else

        Modificador

End If

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventario de Herramientas"
 End If
End Sub

Private Sub btn_personal_Click()
 banderaPersonal = 1
    Call LanzarListadoPersonal(Me, "lbl_caja")
End Sub


Private Sub btn_Registrar_Click()
On Error GoTo Salir

Application.ScreenUpdating = False

    If Me.txt_Fecha.Text = Empty Or _
        Me.txt_caja.Text = Empty Or _
        Me.txt_id.Text = Empty Or _
        Me.txt_personal.Text = Empty Or _
        Me.txt_puesto.Text = Empty Or _
        Me.txt_area.Text = Empty Or _
        Me.txt_estado.Text = Empty Or _
        Me.txt_activo.Text = Empty Or _
        Me.txt_observacion.Text = Empty Then

            MsgBox "Hay campos vacíos en el registro", , "Gestor de Inventario de Herramientas"
            Exit Sub
    
    End If
    
If MsgBox("Son correctos los datos?" + Chr(13) + "Desea procesar el registro?", vbYesNo, "Gestor de Inventarios") = vbNo Then
        Exit Sub
    Else

        Verificador

End If

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventario de Herramientas"
 End If
 
End Sub

Private Sub ProcesarHerramienta()
Dim Indice As Long
        
        Indice = Hoja5.Range("S2").Value
        
            Hoja2.Activate
            Hoja2.Select


                    Hoja2.Rows("2:2").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                    
                    Hoja2.Cells(2, 1) = Indice + 1
                    Hoja2.Cells(2, 2) = Me.txt_area & "-" & Me.txt_caja.Text
                    Hoja2.Cells(2, 3) = Me.txt_id.Text
                    Hoja2.Cells(2, 4) = Me.txt_personal.Text
                    Hoja2.Cells(2, 5) = Me.txt_puesto.Text
                    Hoja2.Cells(2, 6) = Me.txt_area.Text
                    Hoja2.Cells(2, 7) = Me.txt_estado.Text
                    Hoja2.Cells(2, 8) = Me.txt_activo.Text
                    Hoja2.Cells(2, 9) = Me.txt_observacion.Text
                    Hoja2.Cells(2, 10) = CDate(Me.txt_Fecha)
                    
                     
                    Call Carpetas
                    Call Ruta_Caja
                    
                    Hoja5.Range("S2") = Indice + 1
                    

                
End Sub
Private Sub ModificarHerramienta()

        'Modificar los datos de la caja de herramienta en la hoja CAJA
                    
                    ActiveCell.Offset(0, 1) = Me.txt_id.Text
                    ActiveCell.Offset(0, 2) = Me.txt_personal.Text
                    ActiveCell.Offset(0, 3) = Me.txt_puesto.Text
                    ActiveCell.Offset(0, 4) = Me.txt_area.Text
                    ActiveCell.Offset(0, 5) = Me.txt_estado.Text
                    ActiveCell.Offset(0, 6) = Me.txt_activo.Text
                    ActiveCell.Offset(0, 7) = Me.txt_observacion.Text
                    
                    If Me.txt_activo.Text = "Activo" Then
                    ActiveCell.Offset(0, 9) = CDate(Me.txt_Fecha)
                    Else
                    ActiveCell.Offset(0, 10) = CDate(Me.txt_Fecha)
                    End If
                    

End Sub

Private Sub Carpetas()
On Error Resume Next

Dim Ruta As String
Dim carpeta As String

Ruta = "C:\Users\usuario\Desktop\Inventario de Herramientas\Fotos"
carpeta = Me.txt_area.Text & "-" & Me.txt_caja.Text

MkDir (Ruta & "/" & carpeta)

End Sub
Private Sub Ruta_Caja()
Dim xFoto As String
Dim xRuta As String

xFoto = Me.txt_area.Text & "-" & Me.txt_caja.Text

xRuta = "Fotos\" & Me.txt_area.Text & "-" & Me.txt_caja.Text & "\" & Me.txt_caja.Text & ".jpeg"

    Hoja2.Cells(2, 13).Select
    Hoja2.Cells(2, 12).Hyperlinks.Add Anchor:=Selection, Address:= _
        xRuta, TextToDisplay:=xFoto
        
End Sub
Private Sub Verificador()
Dim x As String
Dim Z As String
Dim encontrado As Boolean

x = Me.txt_caja.Text
Z = Me.txt_area.Text & "-" & Me.txt_caja.Text

Hoja2.Select
Range("B1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like x Or ActiveCell.Value Like Z Then
            encontrado = True
            Exit Do
                                 
        End If
    Loop
    If encontrado = True Then
        MsgBox "La caja de herramienta ya existe", vbInformation, Titulo
    
    End If
    
     If encontrado = False Then
        ProcesarHerramienta
        Unload Me
        MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
            
        Hoja0.Activate
        Hoja0.Select
         Application.ScreenUpdating = True
             Application.EnableEvents = False
            ThisWorkbook.Save
        Application.EnableEvents = True
        frm_cajon.Show
            
        
     End If
              
End Sub
Private Sub Modificador()
Dim x As String
Dim encontrado As Boolean

x = Me.txt_caja.Text

Hoja2.Select
Range("B1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like x Then
            encontrado = True
            Exit Do
                                 
        End If
    Loop
    If encontrado = False Then
        MsgBox "La caja de herramienta no existe", vbInformation, Titulo
    
    End If
    
     If encontrado = True Then
        ModificarHerramienta
        Unload Me
        MsgBox "Registro modificado con éxito!!!", vbInformation, Titulo
            
        Hoja0.Activate
        Hoja0.Select
         Application.ScreenUpdating = True
             Application.EnableEvents = False
            ThisWorkbook.Save
        Application.EnableEvents = True
        frm_cajon.Show
            
        
     End If
              
End Sub


Private Sub txt_caja_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 47 And KeyAscii < 58 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub

Private Sub txt_estado_Change()
If Me.txt_estado.Text = "Dañado" Then
    Me.txt_estado.BackColor = &H8080FF
ElseIf Me.txt_estado.Text = "Faltante" Then
    Me.txt_estado.BackColor = &H80FFFF
Else
    Me.txt_estado.BackColor = &H80FF80

End If

End Sub
