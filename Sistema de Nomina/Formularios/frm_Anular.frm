VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Anular 
   Caption         =   "GESTOR DE RECURSOS HUMANOS"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   8370.001
   OleObjectBlob   =   "frm_Anular.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Anular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub btn_Cargar_Click()
On Error GoTo Salir
Dim Titulo As String
Dim Seguridad As String

Titulo = "Gestor de Recursos Humanos"

Seguridad = Hoja83.Range("L1").Text

     If Me.cbx_nombre = Empty Or Me.cbx_personal = Empty Then
            Me.cbx_nombre.BackColor = &HC0C0FF
            Me.cbx_personal.BackColor = &HC0C0FF
            MsgBox "Debe seleccionarun colaborador del listado..!", vbInformation, "Gestor de Recursos Humanos"
            Me.cbx_nombre.BackColor = &HFFFFFF
            Me.cbx_personal.BackColor = &HFFFFFF
            Exit Sub
    End If
    If Me.txt_motivo = Empty Then
            Me.txt_motivo.BackColor = &HC0C0FF
            MsgBox "Detalle una observación sobre la fecha libre..!", vbInformation, "Gestor de Recursos Humanos"
            Me.txt_motivo.BackColor = &HFFFFFF
            Me.txt_motivo.SetFocus
            Exit Sub
    End If




Hoja15.Unprotect (Seguridad)
Hoja11.Unprotect (Seguridad)

    Registrar_Recargo
    Limpiar_Controles
    Unload Me

Hoja15.Protect (Seguridad)
Hoja11.Protect (Seguridad)



Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Recursos Humanos"
 End If
End Sub

Private Sub Registrar_Recargo()
Dim Comprb As Long

Dim Titulo As String
Dim Registro As Date

Titulo = "Gestor de Personal"

Hoja11.Range("F2").Value = Hoja11.Range("F2").Value + 1
Comprb = Hoja11.Range("F2").Value

Registro = Date

    Hoja15.Select
    Hoja15.Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja15.Cells(2, 1) = Comprb
                Hoja15.Cells(2, 2) = Format(Registro, "MM/DD/YYYY")
                Hoja15.Cells(2, 3) = Me.cbx_personal.Text
                Hoja15.Cells(2, 4) = Me.cbx_nombre.Text
                Hoja15.Cells(2, 5) = CDate(txt_Fecha)
                Hoja15.Cells(2, 6) = Me.txt_motivo.Text
                Hoja15.Cells(2, 7) = Hoja83.Range("G1")

         MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
             
                
End Sub
Private Sub Limpiar_Controles()

    Me.txt_motivo = Empty
    Me.cbx_personal = Empty
    Me.cbx_nombre = Empty
End Sub

Private Sub btn_Fecha_Click()
 banderaPeriodo = 8
    Call LanzarPeriodo(Me, "txt_Fecha")
    Me.txt_motivo.SetFocus
End Sub

Private Sub btn_listadopersonal_Click()
banderaPersonal = 10
Call LanzarListadoPersonal(Me, "btn_ListadoPersonal")
Me.txt_motivo.SetFocus
End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub

