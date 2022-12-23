Attribute VB_Name = "Reloj"
Public Actualizar As Boolean
Sub Reloj()
'Carga el userform
frm_Factura.Show
End Sub
Sub Hora()
'Actualiza la etiqueta llamada Reloj cada segundo
If Actualizar Then
frm_Factura.Reloj = Format(Time)
Application.OnTime Now + TimeValue("00:00:01"), "Hora"
End If
End Sub
