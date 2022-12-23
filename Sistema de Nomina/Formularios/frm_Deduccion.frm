VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Deduccion 
   Caption         =   "Gestor de Recursos Humanos"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   8760.001
   OleObjectBlob   =   "frm_Deduccion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Deduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit



Private Sub btn_personal_Click()
banderaPersonal = 14
Call LanzarListadoPersonal(Me, "btn_personal")
Me.txt_Isr.SetFocus
End Sub

Private Sub txt_isr_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_Isr, KeyAscii)
End Sub
Private Sub txt_deduccion_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_deduccion, KeyAscii)
End Sub
Private Sub adelanto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_adelanto, KeyAscii)
End Sub
Private Sub btn_salir_Click()
Unload Me
End Sub
Private Sub CommandButton3_Click()
Dim Titulo As String
Dim Seguridad As String

On Error GoTo Salir

Seguridad = Hoja83.Range("L1").Text
Titulo = "Gestion del Personal"
  
If Me.cboMes.Text = "" Then
    Me.cboMes.BackColor = &HC0C0FF
    MsgBox "Seleccione el periodo del decimo", vbInformation, Titulo
    Me.cboMes.BackColor = &HFFFFFF
    Me.cboMes.SetFocus
    Exit Sub
End If

        If Me.ComboBox1.Text = "" Then
            Me.ComboBox1.BackColor = &HC0C0FF
            Me.ComboBox2.BackColor = &HC0C0FF
            MsgBox "Seleccione un personal del listado", vbInformation, Titulo
            Me.ComboBox2.BackColor = &HFFFFFF
            Me.ComboBox1.BackColor = &HFFFFFF
            Me.btn_personal.SetFocus
            Exit Sub
        End If
        
                          If Me.txt_Isr = Empty And Me.txt_adelanto = Empty And Me.txt_deduccion = Empty Then
                            Me.txt_Isr.BackColor = &HC0C0FF
                            Me.txt_adelanto.BackColor = &HC0C0FF
                            Me.txt_deduccion.BackColor = &HC0C0FF
                            MsgBox "Ingrese el monto de comisión", vbInformation, Titulo
                            Me.txt_Isr.BackColor = &HFFFFFF
                            Me.txt_adelanto.BackColor = &HFFFFFF
                            Me.txt_deduccion.BackColor = &HFFFFFF
                            Me.txt_Isr.SetFocus
                            Exit Sub
                        End If
                        
                                If Me.txt_detalle = "" Then
                                    Me.txt_detalle.BackColor = &HC0C0FF
                                    MsgBox "Registre las observaciones sobre la deduccion", vbInformation, Titulo
                                    Me.txt_detalle.BackColor = &HFFFFFF
                                    Me.txt_detalle.SetFocus
                                    Exit Sub
                                End If
                                
  

  Hoja26.Unprotect (Seguridad)
  
       Registrar_Deduccion
       LimpiarControles

    Hoja26.Protect (Seguridad)
   
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Recursos Humanos"
 End If
End Sub
Private Sub Registrar_Deduccion()
Dim Comprb As Long
Dim Fecha As Date
Dim Titulo As String
Dim Mes As Date
Dim Ano As Date
Dim Dia As Date

Titulo = "Gestor de Recursos Humanos"


Dia = 15

If Me.cboMes.ListIndex = 0 Then
Mes = 4
ElseIf Me.cboMes.ListIndex = 1 Then
Mes = 8
Else
Mes = 12
End If

Ano = Me.label_año2.Value

Fecha = DateSerial(Ano, Mes, Dia)

                Hoja26.Select
                Hoja26.Rows("2:2").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja26.Cells(2, 1) = Date
                Hoja26.Cells(2, 2) = Me.ComboBox1.Text
                Hoja26.Cells(2, 3) = Me.ComboBox2.Text
                Hoja26.Cells(2, 4) = Fecha
                Hoja26.Cells(2, 5) = Me.txt_Isr.Value
                Hoja26.Cells(2, 6) = Me.txt_adelanto.Value
                Hoja26.Cells(2, 7) = Me.txt_deduccion.Value
                Hoja26.Cells(2, 8) = UCase(Me.txt_detalle.Text)
                Hoja26.Cells(2, 9) = Hoja83.Range("G1")

                    
         MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
             


End Sub

Private Sub LimpiarControles()
Me.ComboBox1.Text = Empty
Me.ComboBox2.Text = Empty
Me.txt_detalle.Text = Empty
Me.txt_Isr.Value = Empty
Me.txt_adelanto.Value = Empty
Me.txt_deduccion.Value = Empty
End Sub

Private Sub SpinButton2_Change()
frm_Deduccion.label_año2.Value = frm_Deduccion.SpinButton2.Value
End Sub

Private Sub UserForm_Initialize()
 With frm_Deduccion.cboMes
        .AddItem 1
        .List(0, 1) = "Abril"
        .AddItem 2
        .List(1, 1) = "Agosto"
        .AddItem 3
        .List(2, 1) = "Diciembre"
    End With
    
    frm_Deduccion.cboMes.ListIndex = 0
       
    frm_Deduccion.SpinButton2.Value = VBA.Year(VBA.Date)
    
    frm_Deduccion.label_año2.Value = VBA.Year(VBA.Date)
    
End Sub
