Attribute VB_Name = "LlamarPeriodo"
Option Explicit
Public banderaPeriodo As Long

Public Function LanzarPeriodo(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_Periodo
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_Periodo.StartUpPosition = 0
            frm_Periodo.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_Periodo.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_Periodo.Show

End Function

Public Function InsertarPeriodo(Fecha As Date)

Dim i As Byte
Dim txt_Fecha As textbox

    Select Case banderaPeriodo
        Case 1
            frm_Colilla.txt_Fecha.Text = Fecha
        
        Case 2
            frm_General.txt_Fecha.Text = Fecha
            
        
        Case 3
            frm_Nota.txt_Fecha.Text = Fecha
            
        Case 4
            frm_Ajuste.txt_Fecha.Text = Fecha
        
        Case 5
            frm_Cuota_Abono.txt_Fecha.Text = Fecha
        
         Case 6
            frm_Viatico.txt_Fecha.Text = Fecha
          
          Case 7
            frm_Exonera.txt_Fecha.Text = Fecha
       
          Case 8
            frm_Anular.txt_Fecha.Text = Fecha
            
          Case 9
            frm_Hora_Marca.txt_fecha1.Text = Fecha
            frm_Hora_Marca.txt_fecha2.Text = Fecha + 1
            frm_Hora_Marca.txt_fecha3.Text = Fecha + 2
            frm_Hora_Marca.txt_fecha4.Text = Fecha + 3
            frm_Hora_Marca.txt_fecha5.Text = Fecha + 4
            frm_Hora_Marca.txt_fecha6.Text = Fecha + 5
            frm_Hora_Marca.txt_fecha7.Text = Fecha + 6
            frm_Hora_Marca.txt_fecha8.Text = Fecha + 7
            frm_Hora_Marca.txt_fecha9.Text = Fecha + 8
            frm_Hora_Marca.txt_fecha10.Text = Fecha + 9
            frm_Hora_Marca.txt_fecha11.Text = Fecha + 10
            frm_Hora_Marca.txt_fecha12.Text = Fecha + 11
            frm_Hora_Marca.txt_fecha13.Text = Fecha + 12
            frm_Hora_Marca.txt_fecha14.Text = Fecha + 13
            frm_Hora_Marca.txt_fecha15.Text = Fecha + 14
            frm_Hora_Marca.txt_fecha16.Text = Fecha + 15
            
            Case 10
               frm_Comisiones.txt_Fecha.Text = Fecha
               
            Case 11
               frm_Vacaciones.txt_Fecha.Text = Fecha
            
            Case 12
                frm_ISR.txt_Fecha.Text = Fecha
                

        
        Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
    End Select
End Function
