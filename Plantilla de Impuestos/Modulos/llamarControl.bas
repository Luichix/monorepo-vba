Attribute VB_Name = "llamarControl"
Option Explicit
Public banderaPersonal As Long
Public banderaCategoria As Long
Public banderaListadoAbono As Long


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
        frm_Personal.ComboBox7 = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Personal.ComboBox8 = frm_ListadoPersonal.lbx_Personal.Column(1)
             Unload frm_ListadoPersonal
            
        Case 5
        Hoja58.Range("K6") = frm_ListadoPersonal.lbx_Personal.Column(0)
            Unload frm_ListadoPersonal
        
        Case 6
        frm_Comisiones.ComboBox1 = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Comisiones.ComboBox2 = frm_ListadoPersonal.lbx_Personal.Column(1)
        
            Unload frm_ListadoPersonal
        
        Case 7
        frm_Colilla.cbx_id = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Colilla.cbx_nombre = frm_ListadoPersonal.lbx_Personal.Column(1)
        
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
    
    
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub

Public Function LanzarCuenta(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_cuentapersonal
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frmCalendario.StartUpPosition = 0
            frmCalendario.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frmCalendario.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_cuentapersonal.Show

End Function
Sub Insertarcuenta()

If frm_Categoria.lbx_cuenta.ListIndex = -1 Then
    MsgBox "Debe seleccionar un concepto", vbInformation
    frm_Categoria.lbx_cuenta.SetFocus
    Exit Sub
End If

Select Case banderaCategoria
    Case 1
       frm_Factura.txt_Concepto = frm_Categoria.lbx_cuenta.Column(0)
       
       Unload frm_Categoria
       
    Case 2
        frm_Nota_Credito.txt_Concepto = frm_Categoria.lbx_cuenta.Column(0)

        Unload frm_Categoria
       
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
