VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Comisiones 
   Caption         =   "Gestor de Recursos Humanos"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   8660.001
   OleObjectBlob   =   "frm_Comisiones.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Comisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub btn_Fecha_Click()
Me.txt_Fecha.BackColor = &H80000005
banderaPeriodo = 10
  Call LanzarPeriodo(Me, "txt_Fecha")
  Me.txt_Comision.SetFocus
End Sub

Private Sub btn_personal_Click()
banderaPersonal = 6
Call LanzarListadoPersonal(Me, "btn_personal")
Me.txt_Comision.SetFocus
End Sub

Private Sub Comision_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.Comision, KeyAscii)
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
    MsgBox "Ingrese la fecha de cargo de la comisión", vbInformation, Titulo
    Me.btn_Fecha.SetFocus
    Exit Sub
End If

        If Me.ComboBox1.Text = "" Then
            Me.ComboBox1.BackColor = &HC0C0FF
            MsgBox "Seleccione un personal del listado", vbInformation, Titulo
            Me.btn_personal.SetFocus
            Exit Sub
        End If
        
                          If Me.txt_Comision.Text = "" Then
                            Me.txt_Comision.BackColor = &HC0C0FF
                            MsgBox "Ingrese el monto de comisión", vbInformation, Titulo
                            Me.txt_Comision.SetFocus
                            Exit Sub
                        End If
                        
                                If Me.txt_detalle = "" Then
                                    Me.txt_detalle.BackColor = &HC0C0FF
                                    MsgBox "Registre las observaciones sobre la comision", vbInformation, Titulo
                                    Me.txt_detalle.SetFocus
                                    Exit Sub
                                End If
                                
  
  Hoja11.Unprotect (Seguridad)
  Hoja13.Unprotect (Seguridad)
  
       Registrar_Comision
       LimpiarControles

    Hoja11.Protect (Seguridad)
    Hoja13.Protect (Seguridad)
   
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Recursos Humanos"
 End If
End Sub
Private Sub Registrar_Comision()
Dim Comprb As Long
Dim Fecha As Date
Dim Titulo As String
Dim Mes As Date
Dim Ano As Date
Dim Dia As Date

Titulo = "Gestor de Recursos Humanos"
    
Hoja11.Range("C2").Value = Hoja11.Range("C2").Value + 1
Comprb = Hoja11.Range("C2").Value
Fecha = Me.txt_Fecha.Text

Dia = Fecha + 10
Mes = VBA.Month(Dia)
Ano = VBA.Year(Dia)


                Hoja13.Select
                Hoja13.Rows("2:2").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja13.Cells(2, 1) = Date
                Hoja13.Cells(2, 2) = Me.ComboBox1.Text
                Hoja13.Cells(2, 3) = Me.ComboBox2.Text
                Hoja13.Cells(2, 4) = Format(Fecha, "MM/DD/YYYY")
                Hoja13.Cells(2, 5) = DateSerial(Ano, Mes, 1)
                Hoja13.Cells(2, 6) = Me.txt_Comision.Text
                Hoja13.Cells(2, 7) = UCase(Me.txt_detalle.Text)
                Hoja13.Cells(2, 8) = Hoja83.Range("G1")

                    
         MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
             


End Sub

Private Sub UserForm_Initialize()
Me.Label16.Caption = "No. " & Hoja11.Range("C2").Value + 1 'Llamamos el número de la factura
End Sub

Private Sub LimpiarControles()
Me.ComboBox1.Text = Empty
Me.ComboBox2.Text = Empty
Me.txt_detalle.Text = Empty
Me.txt_Comision.Value = Empty


End Sub

