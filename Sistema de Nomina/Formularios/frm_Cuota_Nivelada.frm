VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Cuota_Nivelada 
   Caption         =   "GESTOR DE RECURSOS HUMANOS"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   7940
   OleObjectBlob   =   "frm_Cuota_Nivelada.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Cuota_Nivelada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Registro_Cuenta()
Dim Comprb As Long
Dim AñoActual As String
Dim Referencia As Long
Dim Estado As String


Comprb = Hoja11.Range("B2") + 1
Referencia = Comprb & frm_Cuenta.txt_ccuenta.Text
Estado = "ACTIVO"

   Hoja8.Select
    Limpiar_Filtro
    Orden_Filtro
    
    Hoja8.Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
            Hoja8.Cells(2, 1) = Comprb
            Hoja8.Cells(2, 2) = Date
            Hoja8.Cells(2, 3) = frm_Cuenta.cbx_personal.Text
            Hoja8.Cells(2, 4) = frm_Cuenta.cbx_nombre.Text
            Hoja8.Cells(2, 5) = frm_Cuenta.txt_ccuenta.Text
            Hoja8.Cells(2, 6) = frm_Cuenta.txt_cuenta.Text
            Hoja8.Cells(2, 7) = UCase(frm_Cuenta.txt_detalle.Text)
            Hoja8.Cells(2, 8) = frm_Cuenta.txt_Principal.Value
            Hoja8.Cells(2, 9) = frm_Cuenta.txt_tasa.Value
            Hoja8.Cells(2, 10) = frm_Cuenta.txt_interes.Value
            Hoja8.Cells(2, 11) = frm_Cuenta.txt_monto.Value
            Hoja8.Cells(2, 12) = Me.txt_Deposito.Value
            Hoja8.Cells(2, 17) = Referencia
            Hoja8.Cells(2, 18) = Hoja83.Range("G1")
            Hoja8.Cells(2, 19) = Estado
            
    
End Sub
Private Sub Registro_Deposito()
Dim Comprb As Long
Dim Titulo As String
Dim X As Long
Dim Deposito As Long
Dim Estado As String
Dim Referencia As Long
Dim Fecha_Deposito As String

Titulo = "Gestor de Personal"
Estado = "SIN ABONAR"

Hoja11.Range("B2").Value = Hoja11.Range("B2").Value + 1
Comprb = Hoja11.Range("B2").Value
Referencia = Comprb & frm_Cuenta.txt_ccuenta.Text
Fecha_Deposito = "SIN ASIGNAR"
            
Deposito = Me.txt_Deposito.Value
                   
Hoja12.Select
Limpiar_Filtro
Orden_Filtro

                   
                 For X = 1 To Deposito
                 
                    Hoja12.Select
                    Hoja12.Rows("2:2").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                            Hoja12.Cells(2, 1) = Hoja12.Cells(3, 1) + 1
                            Hoja12.Cells(2, 2) = Date
                            Hoja12.Cells(2, 3) = frm_Cuenta.cbx_personal.Text
                            Hoja12.Cells(2, 4) = frm_Cuenta.cbx_nombre.Text
                            Hoja12.Cells(2, 5) = frm_Cuenta.txt_ccuenta.Text
                            Hoja12.Cells(2, 6) = frm_Cuenta.txt_cuenta.Text
                            Hoja12.Cells(2, 7) = Me.txt_Cuota_Nivelada.Value
                            Hoja12.Cells(2, 8) = Fecha_Deposito
                            Hoja12.Cells(2, 9) = Referencia
                            Hoja12.Cells(2, 10) = Referencia + X
                            Hoja12.Cells(2, 12) = Hoja83.Range("G1")
            
                            
                    Next X
                    
         MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
             
   
                    
            
End Sub

Private Sub btn_grabar_Click()
On Error GoTo Salir
Dim Seguridad As String

Seguridad = Hoja83.Range("L1").Text


    If Me.txt_Deposito.Value = 0 Or Me.txt_Deposito.Value = "" Then
            MsgBox "Debe detallar las cuotas de pago correctamente", vbInformation, "Gestor de Recursos Humanos"
            Exit Sub

    End If


Hoja8.Unprotect (Seguridad)
Hoja11.Unprotect (Seguridad)
Hoja12.Unprotect (Seguridad)

Registro_Cuenta
Registro_Deposito
LimpiarControles
Unload Me

Hoja8.Protect (Seguridad)
Hoja11.Protect (Seguridad)
Hoja12.Protect (Seguridad)



                    
     
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Ventas"
 End If
End Sub
Private Sub LimpiarControles()
frm_Cuenta.txt_ccuenta = Empty
frm_Cuenta.txt_cuenta = Empty
frm_Cuenta.cbx_personal = Empty
frm_Cuenta.cbx_nombre = Empty
frm_Cuenta.txt_Principal = Empty
frm_Cuenta.txt_tasa = Empty
frm_Cuenta.txt_detalle = Empty

End Sub


Private Sub txt_deposito_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 47 And KeyAscii < 58 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub



Private Sub txt_deposito_Change()
Dim Final As Currency
Dim Deposito As Currency
Dim Cuota As Currency

Me.txt_Deposito.BackColor = &H80000005

Final = frm_Cuenta.txt_monto.Value

If Me.txt_Deposito.Value = "" Or Me.txt_Deposito.Value = 0 Then
    Cuota = 0
ElseIf Me.txt_Deposito.Value > 0 Then
Deposito = Me.txt_Deposito.Value
Cuota = Final / Deposito
Me.txt_Cuota_Nivelada = Cuota


End If

If InStr(Me.txt_Cuota_Nivelada, ",") > 0 Then
nuevo = Replace(Me.txt_Cuota_Nivelada.Value, ",", ".")
Me.txt_Cuota_Nivelada.Value = Format(CDbl(nuevo), "0.00")
End If

If Me.txt_Deposito = "" Then
    Me.txt_Cuota_Nivelada = ""
End If

End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub

Private Sub UserForm_Initialize()
EliminarTitulo Me.Caption
    Me.Height = Me.Height - 20
    
    Me.txt_Final = frm_Cuenta.txt_monto.Value
End Sub
Private Sub Limpiar_Filtro()

Range("A1").Select

          If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

           If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData


End Sub
Private Sub Orden_Filtro()

Range("A1").Sort Key1:=Range("A1"), Order1:=xlDescending, Header:=xlYes

End Sub


