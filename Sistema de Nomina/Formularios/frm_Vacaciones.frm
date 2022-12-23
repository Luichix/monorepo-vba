VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Vacaciones 
   Caption         =   "GESTOR DE RECURSOS HUMANOS"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   8790.001
   OleObjectBlob   =   "frm_Vacaciones.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Vacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btn_Fecha_Click()
Me.txt_Fecha.BackColor = &H80000005
banderaPeriodo = 11
  Call LanzarPeriodo(Me, "txt_Fecha")
  Me.txt_ingresos.SetFocus
End Sub

Private Sub btn_personal_Click()
banderaPersonal = 3
Call LanzarListadoPersonal(Me, "btn_personal")
Me.txt_Comision.SetFocus
End Sub


Private Sub txt_Comision_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_Comision, KeyAscii)
End Sub
Private Sub txt_Decimo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_Decimo, KeyAscii)
End Sub
Private Sub txt_ingresos_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_ingresos, KeyAscii)
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
  
If Me.txt_Fecha.Text = "" Then
    Me.txt_Fecha.BackColor = &HC0C0FF
    MsgBox "Ingrese la fecha de cargo del ajuste", vbInformation, Titulo
    Me.btn_Fecha.SetFocus
    Exit Sub
End If

        If Me.ComboBox1.Text = "" Then
            Me.ComboBox1.BackColor = &HC0C0FF
            MsgBox "Seleccione un personal del listado", vbInformation, Titulo
            Me.btn_personal.SetFocus
            Exit Sub
        End If
        
                          If Me.txt_Comision.Text = Empty And Me.txt_Decimo.Text = Empty And Me.txt_ingresos.Text = Empty Then
                            Me.txt_ingresos.BackColor = &HC0C0FF
                            Me.txt_Comision.BackColor = &HC0C0FF
                            Me.txt_Decimo.BackColor = &HC0C0FF
                            MsgBox "Ingrese el monto del ajuste", vbInformation, Titulo
                            Me.txt_ingresos.BackColor = &HFFFFFF
                            Me.txt_Comision.BackColor = &HFFFFFF
                            Me.txt_Decimo.BackColor = &HFFFFFF
                            Me.txt_ingresos.SetFocus
                            Exit Sub
                        End If
                        
  

  Hoja7.Unprotect (Seguridad)
  
  
  
       Registrar_Comision
       LimpiarControles
    Hoja7.Protect (Seguridad)

   
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Recursos Humanos"
 End If
End Sub
Private Sub Registrar_Comision()
Dim Fecha As Date
Dim Titulo As String
Dim Dia As Date
Dim Mes As Date
Dim Ano As Date
Dim Codigo As Long
Dim Valor As String

Titulo = "Gestor de Recursos Humanos"
    
Fecha = Me.txt_Fecha.Text
            
Dia = Fecha + 10
Mes = VBA.Month(Dia)
Ano = VBA.Year(Dia)

If Me.txt_Comision <> "" And Me.txt_Decimo <> "" Then
    Valor = "PTS"

ElseIf Me.txt_Comision <> "" Then
    Valor = "VAC"

ElseIf Me.txt_Decimo <> "" Then
    Valor = "DTM"
    
Else
    Valor = "ING"

End If
            
            
            
            Codigo = Fecha
                     
                Hoja7.Select
                Hoja7.Rows("2:2").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja7.Cells(2, 1) = Fecha
                Hoja7.Cells(2, 2) = Codigo & "-" & Me.ComboBox1.Text & "-" & Valor
                Hoja7.Cells(2, 3) = Me.ComboBox1.Value
                Hoja7.Cells(2, 5) = Me.txt_ingresos.Value
                Hoja7.Cells(2, 7) = Me.txt_Comision.Value
                Hoja7.Cells(2, 8) = Me.txt_Decimo.Value
                Hoja7.Cells(2, 9) = DateSerial(Ano, Mes, 1)
                Hoja7.Cells(2, 10) = Hoja83.Range("G1")

                    
         MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
             


End Sub

Private Sub LimpiarControles()

Me.ComboBox1.Text = Empty
Me.ComboBox2.Text = Empty
Me.txt_Comision.Value = Empty
Me.txt_ingresos.Value = Empty
Me.txt_Decimo.Value = Empty

End Sub


