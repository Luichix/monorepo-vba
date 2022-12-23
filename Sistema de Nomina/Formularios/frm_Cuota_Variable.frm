VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Cuota_Variable 
   Caption         =   "GESTOR DE RECURSOS HUMANOS"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   12690
   OleObjectBlob   =   "frm_Cuota_Variable.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Cuota_Variable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_grabar_Click()
On Error GoTo Salir
Dim Titulo As String
Dim Seguridad As String

Seguridad = Hoja83.Range("L1").Text

Titulo = "Gestor de Recursos Humanos"


    If Me.txt_Restante.Value <> 0 Then
            MsgBox "No se a detallado las cuotas de pago correctamente..!", vbInformation, "Gestor de Recursos Humanos"
            Exit Sub
    End If
    
    If Me.txt_Valor1 = "" Or Me.txt_Valor1 = 0 Then
        If Me.txt_Monto1 <> "" Or Me.txt_Cuota1 <> "" Then
            MsgBox "Limpie los montos o cuotas ingresados incorrectamente: Registros 01..!", vbInformation, Titulo
            Exit Sub
        End If
    End If
    If Me.txt_Valor2 = "" Or Me.txt_Valor2 = 0 Then
        If Me.txt_Monto2 <> "" Or Me.txt_Cuota2 <> "" Then
            MsgBox "Limpie los montos o cuotas ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Exit Sub
        End If
    End If
    If Me.txt_Valor3 = "" Or Me.txt_Valor3 = 0 Then
        If Me.txt_Monto3 <> "" Or Me.txt_Cuota3 <> "" Then
            MsgBox "Limpie los montos o cuotas ingresados incorrectamente: Registros 03..!", vbInformation, Titulo
            Exit Sub
        End If
    End If
    If Me.txt_Valor4 = "" Or Me.txt_Valor4 = 0 Then
        If Me.txt_Monto4 <> "" Or Me.txt_Cuota4 <> "" Then
            MsgBox "Limpie los montos o cuotas ingresados incorrectamente: Registros 04..!", vbInformation, Titulo
            Exit Sub
        End If
    End If
    If Me.txt_Valor5 = "" Or Me.txt_Valor5 = 0 Then
        If Me.txt_Monto5 <> "" Or Me.txt_Cuota5 <> "" Then
            MsgBox "Limpie los montos o cuotas ingresados incorrectamente: Registros 05..!", vbInformation, Titulo
            Exit Sub
        End If
    End If
    

Hoja8.Unprotect (Seguridad)
Hoja11.Unprotect (Seguridad)
Hoja12.Unprotect (Seguridad)
Application.Cursor = xlWait

Registro_Cuenta
Registro_Deposito

Application.Cursor = xlDefault
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
Private Sub Registro_Deposito()
Dim Comprb As Long
Dim Titulo As String
Dim V As Long
Dim W As Long
Dim X As Long
Dim Y As Long
Dim z As Long
Dim vCuota As Integer
Dim wCuota As Integer
Dim xCuota As Integer
Dim yCuota As Integer
Dim zCuota As Integer
Dim Estado As String
Dim Referencia As Long
Dim Fecha_Deposito As String

Titulo = "Gestor de Personal"
Estado = "SIN ABONAR"
Fecha_Deposito = "SIN ASIGNAR"

Hoja11.Range("B2").Value = Hoja11.Range("B2").Value + 1
Comprb = Hoja11.Range("B2").Value
Referencia = Comprb & frm_Cuenta.txt_ccuenta.Text


Hoja12.Select
    Limpiar_Filtro
    Orden_Filtro
                               
If Me.txt_Cuota1.Value = "" Then
    vCuota = 0
Else
    vCuota = Me.txt_Cuota1.Value
End If

If Me.txt_Cuota2.Value = "" Then
    wCuota = 0
Else
    wCuota = Me.txt_Cuota2.Value
End If

If Me.txt_Cuota3.Value = "" Then
    xCuota = 0
Else
    xCuota = Me.txt_Cuota3.Value
End If

If Me.txt_Cuota4.Value = "" Then
    yCuota = 0
Else
    yCuota = Me.txt_Cuota4.Value
End If
                 
If Me.txt_Cuota5.Value = "" Then
    zCuota = 0
Else
    zCuota = Me.txt_Cuota5.Value
End If

            If vCuota = 0 Then
                
            Else
            
                 For V = 1 To vCuota
                 
                    Hoja12.Select
                    Hoja12.Rows("2:2").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                            Hoja12.Cells(2, 1) = Hoja12.Cells(3, 1) + 1
                            Hoja12.Cells(2, 2) = Date
                            Hoja12.Cells(2, 3) = frm_Cuenta.cbx_personal.Text
                            Hoja12.Cells(2, 4) = frm_Cuenta.cbx_nombre.Text
                            Hoja12.Cells(2, 5) = frm_Cuenta.txt_ccuenta.Text
                            Hoja12.Cells(2, 6) = frm_Cuenta.txt_cuenta.Text
                            Hoja12.Cells(2, 7) = Me.txt_Monto1.Value
                            Hoja12.Cells(2, 8) = Fecha_Deposito
                            Hoja12.Cells(2, 9) = Referencia
                            Hoja12.Cells(2, 10) = Referencia + V
                            Hoja12.Cells(2, 12) = Hoja83.Range("G1")
                            
                            
                         
                    Next V
            End If
    
    
            If wCuota = 0 Then
            Else
                    For W = 1 To wCuota

                    Hoja12.Select
                    Hoja12.Rows("2:2").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                            Hoja12.Cells(2, 1) = Hoja12.Cells(3, 1) + 1
                            Hoja12.Cells(2, 2) = Date
                            Hoja12.Cells(2, 3) = frm_Cuenta.cbx_personal.Text
                            Hoja12.Cells(2, 4) = frm_Cuenta.cbx_nombre.Text
                            Hoja12.Cells(2, 5) = frm_Cuenta.txt_ccuenta.Text
                            Hoja12.Cells(2, 6) = frm_Cuenta.txt_cuenta.Text
                            Hoja12.Cells(2, 7) = Me.txt_Monto2.Value
                            Hoja12.Cells(2, 8) = Fecha_Deposito
                            Hoja12.Cells(2, 9) = Referencia
                            Hoja12.Cells(2, 10) = Referencia + vCuota + W
                            Hoja12.Cells(2, 12) = Hoja83.Range("G1")
                            


                    Next W
            End If
            
            If xCuota = 0 Then
            
            Else
                    For X = 1 To xCuota

                    Hoja12.Select
                    Hoja12.Rows("2:2").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                            Hoja12.Cells(2, 1) = Hoja12.Cells(3, 1) + 1
                            Hoja12.Cells(2, 2) = Date
                            Hoja12.Cells(2, 3) = frm_Cuenta.cbx_personal.Text
                            Hoja12.Cells(2, 4) = frm_Cuenta.cbx_nombre.Text
                            Hoja12.Cells(2, 5) = frm_Cuenta.txt_ccuenta.Text
                            Hoja12.Cells(2, 6) = frm_Cuenta.txt_cuenta.Text
                            Hoja12.Cells(2, 7) = Me.txt_Monto3.Value
                            Hoja12.Cells(2, 8) = Fecha_Deposito
                            Hoja12.Cells(2, 9) = Referencia
                            Hoja12.Cells(2, 10) = Referencia + vCuota + wCuota + X
                            Hoja12.Cells(2, 12) = Hoja83.Range("G1")
                            


                    Next X
            End If


            If yCuota = 0 Then
            Else
                    For Y = 1 To yCuota

                    Hoja12.Select
                    Hoja12.Rows("2:2").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                            Hoja12.Cells(2, 1) = Hoja12.Cells(3, 1) + 1
                            Hoja12.Cells(2, 2) = Date
                            Hoja12.Cells(2, 3) = frm_Cuenta.cbx_personal.Text
                            Hoja12.Cells(2, 4) = frm_Cuenta.cbx_nombre.Text
                            Hoja12.Cells(2, 5) = frm_Cuenta.txt_ccuenta.Text
                            Hoja12.Cells(2, 6) = frm_Cuenta.txt_cuenta.Text
                            Hoja12.Cells(2, 7) = Me.txt_Monto4.Value
                            Hoja12.Cells(2, 8) = Fecha_Deposito
                            Hoja12.Cells(2, 9) = Referencia
                            Hoja12.Cells(2, 10) = Referencia + vCuota + wCuota + xCuota + Y
                            Hoja12.Cells(2, 12) = Hoja83.Range("G1")
                           


                    Next Y
            End If

            
            If zCuota = 0 Then
            Else
                    For z = 1 To zCuota

                    Hoja12.Select
                    Hoja12.Rows("2:2").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                            Hoja12.Cells(2, 1) = Hoja12.Cells(3, 1) + 1
                            Hoja12.Cells(2, 2) = Date
                            Hoja12.Cells(2, 3) = frm_Cuenta.cbx_personal.Text
                            Hoja12.Cells(2, 4) = frm_Cuenta.cbx_nombre.Text
                            Hoja12.Cells(2, 5) = frm_Cuenta.txt_ccuenta.Text
                            Hoja12.Cells(2, 6) = frm_Cuenta.txt_cuenta.Text
                            Hoja12.Cells(2, 7) = Me.txt_Monto5.Value
                            Hoja12.Cells(2, 8) = Fecha_Deposito
                            Hoja12.Cells(2, 9) = Referencia
                            Hoja12.Cells(2, 10) = Referencia + vCuota + wCuota + xCuota + yCuota + z
                            Hoja12.Cells(2, 12) = Hoja83.Range("G1")
                            


                    Next z
              End If
            
                    
         MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
             
   
                    
            
End Sub
Private Sub Registro_Cuenta()
Dim Comprb As Long
Dim AñoActual As String
Dim Referencia As Long
Dim Suma As Integer
Dim vCuota As Integer
Dim wCuota As Integer
Dim xCuota As Integer
Dim yCuota As Integer
Dim zCuota As Integer
Dim Estado As String

Estado = "ACTIVO"
Comprb = Hoja11.Range("B2") + 1
Referencia = Comprb & frm_Cuenta.txt_ccuenta.Text

'''''''''''''''
Hoja8.Select
    Limpiar_Filtro
    Orden_Filtro
                               
If Me.txt_Cuota1.Value = "" Then
    vCuota = 0
Else
    vCuota = Me.txt_Cuota1.Value
End If

If Me.txt_Cuota2.Value = "" Then
    wCuota = 0
Else
    wCuota = Me.txt_Cuota2.Value
End If

If Me.txt_Cuota3.Value = "" Then
    xCuota = 0
Else
    xCuota = Me.txt_Cuota3.Value
End If

If Me.txt_Cuota4.Value = "" Then
    yCuota = 0
Else
    yCuota = Me.txt_Cuota4.Value
End If
                 
If Me.txt_Cuota5.Value = "" Then
    zCuota = 0
Else
    zCuota = Me.txt_Cuota5.Value
End If

'''''''''''''''


Suma = vCuota + wCuota + xCuota + yCuota + zCuota

   Hoja8.Select

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
            Hoja8.Cells(2, 12) = Suma
            Hoja8.Cells(2, 17) = Referencia
            Hoja8.Cells(2, 18) = Hoja83.Range("G1")
            Hoja8.Cells(2, 19) = Estado
            
    
End Sub

Private Sub btn_limpiar1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.btn_limpiar1.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub btn_limpiar1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.btn_limpiar1.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub btn_limpiar2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.btn_limpiar2.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub btn_limpiar2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.btn_limpiar2.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub btn_limpiar3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.btn_limpiar3.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub btn_limpiar3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.btn_limpiar3.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub btn_limpiar4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.btn_limpiar4.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub btn_limpiar4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.btn_limpiar4.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub btn_limpiar5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.btn_limpiar5.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub btn_limpiar5_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.btn_limpiar5.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub btn_limpiar1_Click()
Me.txt_Monto1 = Empty
Me.txt_Cuota1 = Empty
Me.txt_Cuota1.SetFocus
End Sub
Private Sub btn_limpiar2_Click()
Me.txt_Monto2 = Empty
Me.txt_Cuota2 = Empty
Me.txt_Cuota2.SetFocus
End Sub

Private Sub btn_limpiar3_Click()
Me.txt_Monto3 = Empty
Me.txt_Cuota3 = Empty
Me.txt_Cuota3.SetFocus
End Sub

Private Sub btn_limpiar4_Click()
Me.txt_Monto4 = Empty
Me.txt_Cuota4 = Empty
Me.txt_Cuota4.SetFocus
End Sub

Private Sub btn_limpiar5_Click()
Me.txt_Monto5 = Empty
Me.txt_Cuota5 = Empty
Me.txt_Cuota5.SetFocus
End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub
Private Sub txt_Cuota1_Change()
Dim Monto As Currency
Dim Cuota As Integer
Dim Valor As Currency

If Me.txt_Monto1.Value = "" Or Me.txt_Monto1.Value = 0 Then
    Valor = 0
    Me.txt_Valor1 = Valor
    Calculo
ElseIf Me.txt_Cuota1.Value = "" Or Me.txt_Cuota1.Value = 0 Then
    Valor = 0
    Me.txt_Valor1 = Valor
    Calculo
Else
    Monto = Me.txt_Monto1.Value
    Cuota = Me.txt_Cuota1.Value
    Valor = Monto * Cuota
    Me.txt_Valor1 = Valor
    Calculo
End If

End Sub

Private Sub txt_Monto1_Change()
Dim Monto As Currency
Dim Cuota As Integer
Dim Valor As Currency

If Me.txt_Monto1.Value = "" Or Me.txt_Monto1.Value = 0 Then
    Valor = 0
    Me.txt_Valor1.Value = Valor
    Calculo
ElseIf Me.txt_Cuota1.Value = "" Or Me.txt_Cuota1.Value = 0 Then
    Valor = 0
    Me.txt_Valor1.Value = Valor
    Calculo
Else
    Monto = Me.txt_Monto1.Value
    Cuota = Me.txt_Cuota1.Value
    Valor = Monto * Cuota
    Me.txt_Valor1 = Format(CDbl(Valor), "0.00")
    Calculo
End If

End Sub
Private Sub txt_Cuota2_Change()
Dim Monto As Currency
Dim Cuota As Integer
Dim Valor As Currency

If Me.txt_Monto2.Value = "" Or Me.txt_Monto2.Value = 0 Then
    Valor = 0
    Me.txt_Valor2.Value = Valor
    Calculo
ElseIf Me.txt_Cuota2.Value = "" Or Me.txt_Cuota2.Value = 0 Then
    Valor = 0
    Me.txt_Valor2.Value = Valor
    Calculo
Else
    Monto = Me.txt_Monto2.Value
    Cuota = Me.txt_Cuota2.Value
    Valor = Monto * Cuota
    Me.txt_Valor2 = Format(CDbl(Valor), "0.00")
    Calculo
End If

End Sub



Private Sub txt_Monto2_Change()
Dim Monto As Currency
Dim Cuota As Integer
Dim Valor As Currency

If Me.txt_Monto2.Value = "" Or Me.txt_Monto2.Value = 0 Then
    Valor = 0
    Me.txt_Valor2.Value = Valor
    Calculo
ElseIf Me.txt_Cuota2.Value = "" Or Me.txt_Cuota2.Value = 0 Then
    Valor = 0
    Me.txt_Valor2.Value = Valor
    Calculo
Else
    Monto = Me.txt_Monto2.Value
    Cuota = Me.txt_Cuota2.Value
    Valor = Monto * Cuota
    Me.txt_Valor2 = Format(CDbl(Valor), "0.00")
    Calculo
End If

End Sub
Private Sub txt_Cuota3_Change()
Dim Monto As Currency
Dim Cuota As Integer
Dim Valor As Currency

If Me.txt_Monto3.Value = "" Or Me.txt_Monto3.Value = 0 Then
    Valor = 0
    Me.txt_Valor3 = Valor
    Calculo
ElseIf Me.txt_Cuota3.Value = "" Or Me.txt_Cuota3.Value = 0 Then
    Valor = 0
    Me.txt_Valor3 = Valor
    Calculo
Else
    Monto = Me.txt_Monto3.Value
    Cuota = Me.txt_Cuota3.Value
    Valor = Monto * Cuota
    Me.txt_Valor3 = Format(CDbl(Valor), "0.00")
    Calculo
End If

End Sub

Private Sub txt_Monto3_Change()
Dim Monto As Currency
Dim Cuota As Integer
Dim Valor As Currency

If Me.txt_Monto3.Value = "" Or Me.txt_Monto3.Value = 0 Then
    Valor = 0
    Me.txt_Valor3 = Valor
    Calculo
ElseIf Me.txt_Cuota3.Value = "" Or Me.txt_Cuota3.Value = 0 Then
    Valor = 0
    Me.txt_Valor3 = Valor
    Calculo
Else
    Monto = Me.txt_Monto3.Value
    Cuota = Me.txt_Cuota3.Value
    Valor = Monto * Cuota
    Me.txt_Valor3 = Format(CDbl(Valor), "0.00")
    Calculo
End If

End Sub
Private Sub txt_Cuota4_Change()
Dim Monto As Currency
Dim Cuota As Integer
Dim Valor As Currency

If Me.txt_Monto4.Value = "" Or Me.txt_Monto4.Value = 0 Then
    Valor = 0
    Me.txt_Valor4 = Valor
    Calculo
ElseIf Me.txt_Cuota4.Value = "" Or Me.txt_Cuota4.Value = 0 Then
    Valor = 0
    Me.txt_Valor4 = Valor
    Calculo
Else
    Monto = Me.txt_Monto4.Value
    Cuota = Me.txt_Cuota4.Value
    Valor = Monto * Cuota
    Me.txt_Valor4 = Format(CDbl(Valor), "0.00")
    Calculo
End If

End Sub

Private Sub txt_Monto4_Change()
Dim Monto As Currency
Dim Cuota As Integer
Dim Valor As Currency

If Me.txt_Monto4.Value = "" Or Me.txt_Monto4.Value = 0 Then
    Valor = 0
    Me.txt_Valor4 = Valor
    Calculo
ElseIf Me.txt_Cuota4.Value = "" Or Me.txt_Cuota4.Value = 0 Then
    Valor = 0
    Me.txt_Valor4 = Valor
    Calculo
Else
    Monto = Me.txt_Monto4.Value
    Cuota = Me.txt_Cuota4.Value
    Valor = Monto * Cuota
    Me.txt_Valor4 = Format(CDbl(Valor), "0.00")
    Calculo
End If

End Sub
Private Sub txt_Cuota5_Change()
Dim Monto As Currency
Dim Cuota As Integer
Dim Valor As Currency

If Me.txt_Monto5.Value = "" Or Me.txt_Monto5.Value = 0 Then
    Valor = 0
    Me.txt_Valor5 = Valor
    Calculo
ElseIf Me.txt_Cuota5.Value = "" Or Me.txt_Cuota5.Value = 0 Then
    Valor = 0
    Me.txt_Valor5 = Valor
    Calculo
Else
    Monto = Me.txt_Monto5.Value
    Cuota = Me.txt_Cuota5.Value
    Valor = Monto * Cuota
    Me.txt_Valor5 = Format(CDbl(Valor), "0.00")
    Calculo
End If

End Sub

Private Sub txt_Monto5_Change()
Dim Monto As Currency
Dim Cuota As Integer
Dim Valor As Currency

If Me.txt_Monto5.Value = "" Or Me.txt_Monto5.Value = 0 Then
    Valor = 0
    Me.txt_Valor5 = Valor
    Calculo
ElseIf Me.txt_Cuota5.Value = "" Or Me.txt_Cuota5.Value = 0 Then
    Valor = 0
    Me.txt_Valor5 = Valor
    Calculo
Else
    Monto = Me.txt_Monto5.Value
    Cuota = Me.txt_Cuota5.Value
    Valor = Monto * Cuota
    Me.txt_Valor5 = Format(CDbl(Valor), "0.00")
    Calculo
End If

End Sub

Private Sub Calculo()
Dim Suma As Currency
Dim Restante As Currency
Dim Final As Currency
Dim VA1 As Currency
Dim VA2 As Currency
Dim VA3 As Currency
Dim VA4 As Currency
Dim VA5 As Currency



Final = Me.txt_Final.Value

If Me.txt_Valor1.Value = "" Then
VA1 = 0
Else
VA1 = Me.txt_Valor1.Value
End If

If Me.txt_Valor2.Value = "" Then
VA2 = 0
Else
VA2 = Me.txt_Valor2.Value
End If

If Me.txt_Valor3.Value = "" Then
VA3 = 0
Else
VA3 = Me.txt_Valor3.Value
End If

If Me.txt_Valor4.Value = "" Then
VA4 = 0
Else
VA4 = Me.txt_Valor4.Value
End If

If Me.txt_Valor5.Value = "" Then
VA5 = 0
Else
VA5 = Me.txt_Valor5.Value
End If

Suma = VA1 + VA2 + VA3 + VA4 + VA5

Me.txt_Suma.Value = Suma

Restante = Final - Suma

Me.txt_Restante.Value = Restante

End Sub

Private Sub txt_Monto1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_Monto1, KeyAscii)
End Sub
Private Sub txt_Monto2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_Monto2, KeyAscii)
End Sub
Private Sub txt_Monto3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_Monto3, KeyAscii)
End Sub
Private Sub txt_Monto4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_Monto4, KeyAscii)
End Sub
Private Sub txt_Monto5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_Monto5, KeyAscii)
End Sub
Private Sub txt_Cuota1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 47 And KeyAscii < 58 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub
Private Sub txt_Cuota2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 47 And KeyAscii < 58 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub
Private Sub txt_Cuota3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 47 And KeyAscii < 58 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub
Private Sub txt_Cuota4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 47 And KeyAscii < 58 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub
Private Sub txt_Cuota5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 47 And KeyAscii < 58 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub

Private Sub UserForm_Initialize()
EliminarTitulo Me.Caption
    Me.Height = Me.Height - 20
    
    Me.txt_Final = frm_Cuenta.txt_monto.Value
    Me.txt_Restante = frm_Cuenta.txt_monto.Value
End Sub
Private Sub Limpiar_Filtro()

Range("A1").Select

          If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

           If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData


End Sub
Private Sub Orden_Filtro()

Range("A1").Sort Key1:=Range("A1"), Order1:=xlDescending, Header:=xlYes

End Sub


