VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Incapacidad 
   Caption         =   "Gestor de Recursos Humanos"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   9000.001
   OleObjectBlob   =   "frm_Incapacidad.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Incapacidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_Fin_Click()
Me.txt_Fin.BackColor = &H80000005
banderaCalendario = 6
  Call LanzarCalendario(Me, "txt_fin")
  Me.txt_tiempo.SetFocus
End Sub
Private Sub btn_Inicio_Click()
Me.txt_Inicio.BackColor = &H80000005
banderaCalendario = 5
  Call LanzarCalendario(Me, "txt_inicio")
  Me.txt_tiempo.SetFocus
End Sub
Private Sub btn_personal_Click()
banderaPersonal = 16
Call LanzarListadoPersonal(Me, "btn_personal")
Me.txt_tiempo.SetFocus
End Sub
Private Sub btn_salir_Click()
Unload Me
End Sub
Private Sub txt_tiempo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_tiempo, KeyAscii)
End Sub
Private Sub txt_tiempo_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_tiempo, KeyCode)
End Sub

Private Sub btn_guardar_Click()
Dim Titulo As String
Dim Seguridad As String

On Error GoTo Salir

Seguridad = Hoja83.Range("L1").Text
Titulo = "Gestion del Personal"
  
If Me.txt_id.Text = "" Or Me.txt_colaborador.Text = "" Then
    Me.txt_id.BackColor = &HC0C0FF
    Me.txt_colaborador.BackColor = &HC0C0FF
    MsgBox "Seleccione un personal del listado", vbInformation, Titulo
    Me.txt_id.BackColor = &HFFFFFF
    Me.txt_colaborador.BackColor = &HFFFFFF
    Exit Sub
End If
If Me.txt_Inicio.Text = "" Then
    Me.txt_Inicio.BackColor = &HC0C0FF
    MsgBox "Ingrese la fecha de inicio", vbInformation, Titulo
    Me.txt_Inicio.BackColor = &HFFFFFF
    Exit Sub
End If
If Me.txt_Fin.Text = "" Then
    Me.txt_Fin.BackColor = &HC0C0FF
    MsgBox "Ingrese la fecha de fin", vbInformation, Titulo
    Me.txt_Fin.BackColor = &HFFFFFF
    Exit Sub
End If
If Me.txt_tiempo.Text = Empty Then
    Me.txt_tiempo.BackColor = &HC0C0FF
    MsgBox "Ingrese el tiempo de incapacidad", vbInformation, Titulo
    Me.txt_tiempo.BackColor = &HFFFFFF
    Me.txt_tiempo.SetFocus
    Exit Sub
End If
If Me.txt_detalle.Text = Empty Then
    Me.txt_detalle.BackColor = &HC0C0FF
    MsgBox "Detalle alguna observacion", vbInformation, Titulo
    Me.txt_detalle.BackColor = &HFFFFFF
   Me.txt_detalle.SetFocus
    Exit Sub
End If
                
    Hoja27.Unprotect (Seguridad)
       Registrar_Incapacidad
       LimpiarControles
    Hoja27.Protect (Seguridad)

   
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Recursos Humanos"
 End If
End Sub
Private Sub Registrar_Incapacidad()
Dim Inicio As Date
Dim Fin As Date
Dim Titulo As String

Titulo = "Gestor de Recursos Humanos"
    
Inicio = Me.txt_Inicio.Text
Fin = Me.txt_Fin.Text

            
                     
                Hoja27.Select
                Hoja27.Rows("2:2").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja27.Cells(2, 1) = Date
                Hoja27.Cells(2, 2) = Me.txt_id.Text
                Hoja27.Cells(2, 3) = Me.txt_colaborador.Text
                Hoja27.Cells(2, 4) = Inicio
                Hoja27.Cells(2, 5) = Fin
                Hoja27.Cells(2, 6) = Me.txt_tiempo.Value
                Hoja27.Cells(2, 7) = Me.txt_detalle
                Hoja27.Cells(2, 8) = Hoja83.Range("G1")

                    
         MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
             


End Sub

Private Sub LimpiarControles()
Me.txt_colaborador.Text = Empty
Me.txt_id.Text = Empty
Me.txt_Inicio = Empty
Me.txt_Fin = Empty
Me.txt_detalle = Empty
Me.txt_tiempo = "00:00"
End Sub


