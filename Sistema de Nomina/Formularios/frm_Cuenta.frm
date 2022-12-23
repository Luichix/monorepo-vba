VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Cuenta 
   Caption         =   "GESTOR DE RECURSOS HUMANOS"
   ClientHeight    =   8730.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   10550
   OleObjectBlob   =   "frm_Cuenta.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Cuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_Cargar_Click()
 Dim Titulo As String

Titulo = "Gestión de Personal"




    If Me.cbx_personal = "" Or Me.cbx_nombre = "" Then
        MsgBox "Debe seleccionar un personal del listado", vbInformation, Titulo
        Me.cbx_personal.BackColor = &HC0C0FF
        Me.cbx_nombre.BackColor = &HC0C0FF
        Exit Sub
    End If
    If Me.txt_ccuenta = "" Or Me.txt_cuenta = "" Then
        MsgBox "Debe seleccionar la cuenta a cargar", vbInformation, Titulo
        Me.txt_ccuenta.BackColor = &HC0C0FF
        Me.txt_cuenta.BackColor = &HC0C0FF
        Exit Sub
    End If
        If Me.txt_Principal.Value = "" Or Me.txt_Principal.Value = 0 Then
        MsgBox "Debe introducir el monto inicial de la cuenta", vbInformation, Titulo
        Me.txt_Principal.BackColor = &HC0C0FF
        Exit Sub
    End If
        If Me.txt_detalle = "" Then
        MsgBox "Ingrese las observaciones de la cuenta", vbInformation, Titulo
        Me.txt_detalle.BackColor = &HC0C0FF
        Exit Sub
    End If
       

    frm_Cuotas_Select.Show
    
        


End Sub

Private Sub btn_cuenta_Click()
banderaCuenta = 1
    Call LanzarCuenta(Me, "btn_Cuenta")
    Me.txt_Principal.SetFocus
End Sub


Private Sub btn_personal_Click()
banderaPersonal = 2
Call LanzarListadoPersonal(Me, "btn_Fecha")
End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub
Private Sub txt_cuota_Change()
Dim MontoFinal As Double
Dim Cuota As Double
Dim Abono As Double

If Me.txt_cuota = "" Or Me.txt_cuota = 0 Then
Me.txt_abono = ""
Exit Sub
Else
MontoFinal = Me.txt_monto.Value
Cuota = Me.txt_cuota.Value
Abono = MontoFinal / Cuota

Me.txt_abono.Value = Abono
End If
Me.txt_abono.BackColor = &HFFFFFF
End Sub



Private Sub cbx_personal_Change()
cbx_personal.BackColor = &H80000005
cbx_nombre.BackColor = &H80000005
End Sub

Private Sub txt_ccuenta_Change()
txt_ccuenta.BackColor = &H80000005
txt_cuenta.BackColor = &H80000005
End Sub

Private Sub txt_fecha_Change()
txt_Fecha.BackColor = &H80000005
End Sub

Private Sub txt_interes_Change()

End Sub

Private Sub txt_Principal_Change()
Dim Principal As Double
Dim Tasa As Double
Dim Interes As Double
Dim Final As Double

If Me.txt_Principal.Value = "" Or Me.txt_Principal.Value = 0 Then
Me.txt_monto = ""
Me.txt_interes = 0
Exit Sub
ElseIf Me.txt_tasa = "" Or Me.txt_tasa = 0 Then
Me.txt_interes = 0
Me.txt_monto.Value = Me.txt_Principal.Value
Else
Principal = Me.txt_Principal.Value
Tasa = Me.txt_tasa.Value

Interes = Principal * (Tasa / 100)
Final = Principal + Interes

Me.txt_interes = Format(CDbl(Interes), "0.00")
Me.txt_monto = Format(CDbl(Final), "0.00")


End If
Me.txt_Principal.BackColor = &HFFFFFF
End Sub

Private Sub txt_Principal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(txt_Principal, KeyAscii)
End Sub
Private Sub txt_tasa_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(txt_tasa, KeyAscii)
End Sub
Private Sub txt_cuota_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(txt_cuota, KeyAscii)
End Sub

Private Sub txt_tasa_Change()
Dim Principal As Double
Dim Tasa As Double
Dim Interes As Double
Dim Final As Double

If Me.txt_Principal.Value = "" Or Me.txt_Principal.Value = 0 Then
Me.txt_monto = ""
Me.txt_interes = 0
Exit Sub
ElseIf Me.txt_tasa = "" Or Me.txt_tasa = 0 Then
Me.txt_interes = 0
Me.txt_monto.Value = Me.txt_Principal.Value
Else
Principal = Me.txt_Principal.Value
Tasa = Me.txt_tasa.Value

Interes = Principal * (Tasa / 100)
Final = Principal + Interes

Me.txt_interes = Format(CDbl(Interes), "0.00")
Me.txt_monto = Format(CDbl(Final), "0.00")
End If
Me.txt_tasa.BackColor = &HFFFFFF

End Sub
