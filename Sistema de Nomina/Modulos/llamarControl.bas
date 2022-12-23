Attribute VB_Name = "llamarControl"
Option Explicit
Public banderaPersonal As Long
Public banderaCuenta As Long
Public banderaListadoAbono As Long
Public banderaEliminarAbono As Long
Public banderaContrato As Long
Public banderaRegimen As Long
Public banderaJornada As Long
Public banderaPago As Long
Public banderaColillaPago As Long


Public Function LanzarListadoPersonal(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_ListadoPersonal
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_ListadoPersonal.StartUpPosition = 0
            frm_ListadoPersonal.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_ListadoPersonal.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_ListadoPersonal.Show

End Function
Sub InsertarPersonal()

If frm_ListadoPersonal.lbx_Personal.ListIndex = -1 Then
    MsgBox "Debe seleccionar un Colaborador", vbInformation
    frm_ListadoPersonal.lbx_Personal.SetFocus
    Exit Sub
End If

Select Case banderaPersonal
    Case 1
        With frm_hora
            .txt_nombre = frm_ListadoPersonal.lbx_Personal.Column(1)
            .txt_id = frm_ListadoPersonal.lbx_Personal.Column(0)
            Unload frm_ListadoPersonal
            
        End With
        
    Case 2
        frm_Cuenta.cbx_personal = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Cuenta.cbx_nombre = frm_ListadoPersonal.lbx_Personal.Column(1)

            Unload frm_ListadoPersonal

        Case 3
        frm_Vacaciones.ComboBox1 = frm_ListadoPersonal.lbx_Personal.Column(0)
         frm_Vacaciones.ComboBox2 = frm_ListadoPersonal.lbx_Personal.Column(1)
        
            Unload frm_ListadoPersonal
            
        Case 4
        frm_Personal.txt_Aid = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Personal.txt_Anombre = frm_ListadoPersonal.lbx_Personal.Column(1)
             Unload frm_ListadoPersonal
            
        Case 5
        Hoja58.Range("K6") = frm_ListadoPersonal.lbx_Personal.Column(0)
            Unload frm_ListadoPersonal
        
        Case 6
        frm_Comisiones.ComboBox1 = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Comisiones.ComboBox2 = frm_ListadoPersonal.lbx_Personal.Column(1)
        
            Unload frm_ListadoPersonal
        
        Case 7
        frm_Nota.txt_Aid = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Nota.txt_Anombre = frm_ListadoPersonal.lbx_Personal.Column(1)
        
            Unload frm_ListadoPersonal
                        
        Case 8
        With frm_hora_Multiple
            .txt_nombre = frm_ListadoPersonal.lbx_Personal.Column(1)
            .txt_id = frm_ListadoPersonal.lbx_Personal.Column(0)
            Unload frm_ListadoPersonal
        End With
    
            Case 9
        frm_Exonera.cbx_personal = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Exonera.cbx_nombre = frm_ListadoPersonal.lbx_Personal.Column(1)
        
            Unload frm_ListadoPersonal
            
                    Case 10
        frm_Anular.cbx_personal = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Anular.cbx_nombre = frm_ListadoPersonal.lbx_Personal.Column(1)
        
            Unload frm_ListadoPersonal
            
             Case 11
        frm_Ajuste.ComboBox1 = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Ajuste.ComboBox2 = frm_ListadoPersonal.lbx_Personal.Column(1)
        
            Unload frm_ListadoPersonal
    
            Case 12
        frm_Viatico.ComboBox1 = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Viatico.ComboBox2 = frm_ListadoPersonal.lbx_Personal.Column(1)
        
            Unload frm_ListadoPersonal
            
                    Case 13
        frm_Reporte_Jornada.txt_id = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Reporte_Jornada.txt_nombre = frm_ListadoPersonal.lbx_Personal.Column(1)
        
            Unload frm_ListadoPersonal
            
                 Case 14
        frm_Deduccion.ComboBox1 = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Deduccion.ComboBox2 = frm_ListadoPersonal.lbx_Personal.Column(1)
        
            Unload frm_ListadoPersonal
            
            Case 15
        frm_ISR.ComboBox1 = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_ISR.ComboBox2 = frm_ListadoPersonal.lbx_Personal.Column(1)
        
            Unload frm_ListadoPersonal
    
    
    
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub

Public Function LanzarCuenta(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_Cuentapersonal
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frmCalendario.StartUpPosition = 0
            frmCalendario.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frmCalendario.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_Cuentapersonal.Show

End Function
Sub Insertarcuenta()

If frm_Cuentapersonal.lbx_cuenta.ListIndex = -1 Then
    MsgBox "Debe seleccionar una cuenta", vbInformation
    frm_Cuentapersonal.lbx_cuenta.SetFocus
    Exit Sub
End If

Select Case banderaCuenta
    Case 1
       frm_Cuenta.txt_ccuenta = frm_Cuentapersonal.lbx_cuenta.Column(0)
       frm_Cuenta.txt_cuenta = frm_Cuentapersonal.lbx_cuenta.Column(1)
       
        Unload frm_Cuentapersonal
       
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Public Function LanzarContrato(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_Contrato
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_Contrato.StartUpPosition = 0
            frm_Contrato.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_Contrato.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_Contrato.Show

End Function
Sub Insertarcontrato()

If frm_Contrato.lbx_cuenta.ListIndex = -1 Then
    MsgBox "Debe seleccionar una categoria", vbInformation
    frm_Contrato.lbx_cuenta.SetFocus
    Exit Sub
End If

Select Case banderaContrato
    Case 1
       frm_Personal.txt_Contrato = frm_Contrato.lbx_cuenta.Column(0)
      
        Unload frm_Contrato
    Case 2
       frm_Personal.txt_Acontrato = frm_Contrato.lbx_cuenta.Column(0)
      
        Unload frm_Contrato
       
       
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Public Function LanzarRegimen(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_Regimen
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_Regimen.StartUpPosition = 0
            frm_Regimen.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_Regimen.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_Regimen.Show

End Function
Sub InsertarRegimen()

If frm_Regimen.lbx_cuenta.ListIndex = -1 Then
    MsgBox "Debe seleccionar una categoria", vbInformation
    frm_Regimen.lbx_cuenta.SetFocus
    Exit Sub
End If

Select Case banderaRegimen
    Case 1
       frm_Personal.txt_Regimen = frm_Regimen.lbx_cuenta.Column(0)
      
        Unload frm_Regimen
    Case 2
       frm_Personal.txt_Aregimen = frm_Regimen.lbx_cuenta.Column(0)
      
        Unload frm_Regimen
              
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Public Function LanzarJornada(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_Jornada
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_Jornada.StartUpPosition = 0
            frm_Jornada.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_Jornada.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_Jornada.Show

End Function
Sub InsertarJornada()

If frm_Jornada.lbx_cuenta.ListIndex = -1 Then
    MsgBox "Debe seleccionar una categoria", vbInformation
    frm_Jornada.lbx_cuenta.SetFocus
    Exit Sub
End If

Select Case banderaJornada
    Case 1
       frm_Personal.txt_Jornada = frm_Jornada.lbx_cuenta.Column(0)
      
        Unload frm_Jornada
    Case 2
       frm_Personal.txt_Ajornada = frm_Jornada.lbx_cuenta.Column(0)
      
        Unload frm_Jornada
       
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Public Function LanzarPago(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_Pago
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_Pago.StartUpPosition = 0
            frm_Pago.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_Pago.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_Pago.Show

End Function
Sub InsertarPago()

If frm_Pago.lbx_cuenta.ListIndex = -1 Then
    MsgBox "Debe seleccionar una categoria", vbInformation
    frm_Pago.lbx_cuenta.SetFocus
    Exit Sub
End If

Select Case banderaPago
    Case 1
       frm_Personal.txt_Pago = frm_Pago.lbx_cuenta.Column(0)
      
        Unload frm_Pago
    Case 2
       frm_Personal.txt_APago = frm_Pago.lbx_cuenta.Column(0)
      
        Unload frm_Pago
       
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Public Function LanzarColillaPago(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_ColillaPago
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_ColillaPago.StartUpPosition = 0
            frm_ColillaPago.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_ColillaPago.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_ColillaPago.Show

End Function
Sub InsertarColillaPago()

If frm_ColillaPago.lbx_cuenta.ListIndex = -1 Then
    MsgBox "Debe seleccionar una categoria..!", vbInformation
    frm_ColillaPago.lbx_cuenta.SetFocus
    Exit Sub
End If

Select Case banderaColillaPago
    Case 1
       frm_General.txt_ColillaPago = frm_ColillaPago.lbx_cuenta.Column(0)

        Unload frm_ColillaPago


    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Sub InsertarListadoAbono()

If frm_ListadoAbono.lbx_ListadoAbono.ListIndex = -1 Then
    MsgBox "Debe seleccionar un registro", vbInformation
    frm_ListadoAbono.lbx_ListadoAbono.SetFocus
    Exit Sub
End If

Select Case banderaListadoAbono
    
    Case 1
        With frm_ListadoAbono
            .txt_idpersonal = frm_ListadoAbono.lbx_ListadoAbono.Column(0)
            .txt_nombre = frm_ListadoAbono.lbx_ListadoAbono.Column(1)
            .txt_referencia = frm_ListadoAbono.lbx_ListadoAbono.Column(9)
            
        End With

           
    
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Sub InsertarEliminarAbono()

If frm_EliminarAbono.lbx_ListadoAbono.ListIndex = -1 Then
    MsgBox "Debe seleccionar un registro", vbInformation
    frm_EliminarAbono.lbx_ListadoAbono.SetFocus
    Exit Sub
End If

Select Case banderaEliminarAbono
    
    Case 1
        With frm_EliminarAbono
            .txt_idpersonal = frm_EliminarAbono.lbx_ListadoAbono.Column(0)
            .txt_nombre = frm_EliminarAbono.lbx_ListadoAbono.Column(1)
            .txt_referencia = frm_EliminarAbono.lbx_ListadoAbono.Column(9)
            .txt_Valor_actual = frm_EliminarAbono.lbx_ListadoAbono.Column(8)
        End With

           
    
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
