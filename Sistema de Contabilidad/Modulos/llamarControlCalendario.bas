Attribute VB_Name = "llamarControlCalendario"
Option Explicit
Public banderaCalendario As Long

Public Function LanzarCalendario(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frmCalendario
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frmCalendario.StartUpPosition = 0
            frmCalendario.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frmCalendario.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frmCalendario.Show

End Function

Public Function InsertarFecha(Fecha As Date)
    Select Case banderaCalendario
'        Case 1
'            frm_Factura.txtFecha.Text = Fecha
'
        Case 2
            frm_Compra.txt_Fecha.Text = Fecha
            
        'Case 3
            'frm_Consulta2.txtFecha1.Text = Fecha
        
        'Case 4
            'frm_Consulta2.txtFecha2.Text = Fecha
            
'        Case 5
'            frm_Transferencias.txt_FechaSal.Text = Fecha
'        Case 6
'            form_registropagos.Text_fecha.Text = Fecha
'        Case 7
'            frm_Ganado.Text_fecha2 = Fecha
'        Case 8
'            frm_Ganado.Text_fecha1 = Fecha
'        Case 9
'            frm_enfermeria.Text_fecha = Fecha
'        Case 10
'            frm_Personal.Text_fecha = Fecha
'        Case 11
'            frm_egresosganaderia.Text_fecha = Fecha
'        Case 12
'            frm_egresosmadera.Text_fecha = Fecha
'        Case 13
'            frm_egresosgastos.Text_fecha = Fecha
'        Case 14
'            frm_egresosfamiliares.Text_fecha = Fecha
'        Case 15
'            frm_egresoslegales.Text_fecha = Fecha
'        Case 16
'            frm_Personal.Text_Fecha_Acc = Fecha
'        Case 17
'            frm_ausencias.TextBox1 = Fecha
'        Case 18
'            frm_ausencias.TextBox2 = Fecha
'        Case 19
'            frm_horas.TextBox1 = Fecha
'        Case 20
            frm_Compra.txt_Fecha = Fecha
                    
        Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
    End Select
End Function
