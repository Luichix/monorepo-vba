VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Nota 
   Caption         =   "GESTOR DE RECURSOS HUMANOS"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   9230.001
   OleObjectBlob   =   "frm_Nota.frx":0000
   StartUpPosition =   3  'Predeterminado de Widnows
End
Attribute VB_Name = "frm_Nota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btn_Cargar_Click()
On Error GoTo Salir
Dim Titulo As String

Titulo = "Gestor de Recursos Humanos"

Application.ScreenUpdating = False
     
     If Me.txt_Fecha = Empty Then
            Me.txt_Fecha.BackColor = &HC0C0FF
            MsgBox "Debe seleccionar el dia a configurar libre..!", vbInformation, "Gestor de Recursos Humanos"
            Me.txt_Fecha.BackColor = &HFFFFFF
            Exit Sub
    End If
    
        If Me.txt_Aid = Empty Then
            Me.txt_Aid.BackColor = &HC0C0FF
            MsgBox "Debe seleccionar un colaborador del listado..!", vbInformation, "Gestor de Recursos Humanos"
            Me.txt_Aid.BackColor = &HFFFFFF
            Exit Sub
    End If

    If Me.txt_motivo = Empty Then
            Me.txt_motivo.BackColor = &HC0C0FF
            MsgBox "Detalle una observación sobre la fecha libre..!", vbInformation, "Gestor de Recursos Humanos"
            Me.txt_motivo.BackColor = &HFFFFFF
            Me.txt_motivo.SetFocus
            Exit Sub
    End If

    Registrar_Recordatorio
    Limpiar_Controles
    


     MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo

Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Recursos Humanos"
 End If
End Sub

Private Sub Registrar_Recordatorio()
Dim Titulo As String
Dim Seguridad As String
Dim Estado As String

Seguridad = Hoja83.Range("L1").Text
Estado = "ACTIVO"
Titulo = "Gestor de Personal"
Hoja9.Unprotect (Seguridad)

    Hoja9.Select
    Hoja9.Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja9.Cells(2, 1) = Date
                Hoja9.Cells(2, 2) = Me.txt_Aid.Text
                Hoja9.Cells(2, 3) = Me.txt_Anombre.Text
                Hoja9.Cells(2, 4) = UCase(Me.txt_motivo.Text)
                Hoja9.Cells(2, 5) = CDate(Me.txt_Fecha)
                Hoja9.Cells(2, 6) = Estado
                Hoja9.Cells(2, 7) = Hoja83.Range("G1")

                Hoja9.Cells(1, 1).Select
                
Hoja9.Protect (Seguridad)
                
End Sub
Private Sub Limpiar_Controles()
    Me.txt_Aid = Empty
    Me.txt_Anombre = Empty
    Me.txt_motivo = Empty
    Me.txt_Fecha = Empty
End Sub

Private Sub btn_Fecha_Click()
    banderaPeriodo = 3
  Call LanzarPeriodo(Me, "txt_Fecha")
    Me.txt_motivo.SetFocus
End Sub
Private Sub btn_listadopersonal_Click()
banderaPersonal = 7
Call LanzarListadoPersonal(Me, "btn_listadopersonal")
Me.txt_motivo.SetFocus
End Sub
Private Sub btn_salir_Click()
Unload Me
End Sub




