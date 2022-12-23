Attribute VB_Name = "Modulo_Correo"
Sub EnviaPDF()
Dim mi_hoja As Worksheet
Dim mi_ventana As FileDialog
Dim mi_carpeta As String
Dim mi_archivo As Integer
Dim mi_outlook As Object
Dim mi_correo As Object
Dim mi_rango As Range
Set mi_hoja = ActiveSheet
Set mi_ventana = Application.FileDialog(msoFileDialogFolderPicker)
If mi_ventana.Show = True Then
mi_carpeta = mi_ventana.SelectedItems(1)
Else
MsgBox "No se indicó la carpeta donde guardar el PDF…" _
& vbCrLf & vbCrLf & "Operación cancelada.", _
vbCritical, "Carpeta de almacenamiento pdf"
Exit Sub
End If
mi_carpeta = mi_carpeta + "\" + mi_hoja.Name + ".pdf"
If Len(Dir(mi_carpeta)) > 0 Then
mi_archivo = MsgBox(mi_carpeta & " existente." _
& vbCrLf & vbCrLf & "¿Desea reemplazarlo?", _
vbYesNo + vbQuestion, "Archivo existente")
On Error Resume Next
If mi_archivo = vbYes Then
Kill mi_carpeta
Else
MsgBox "Reemplazar el archivo PDF existente para continuar…" _
& vbCrLf & vbCrLf & "Operación cancelada.", _
vbCritical, "Confirmar guardar como"
Exit Sub
End If
If Err.Number <> 0 Then

MsgBox "El archivo pdf se encuentra abierto o protegido como sólo lectura." _
& vbCrLf & vbCrLf, vbCritical, _
"Error al guardar el archivo"

Exit Sub
End If
End If
Set mi_rango = mi_hoja.UsedRange
If Application.WorksheetFunction.CountA(mi_rango.Cells) <> 0 Then
mi_hoja.ExportAsFixedFormat Type:=xlTypePDF, _
FileName:=mi_carpeta, Quality:=xlQualityStandard
Set mi_outlook = CreateObject("Outlook.Application")
Set mi_correo = mi_outlook.CreateItem(0)
With mi_correo
.Display
.To = ""
.CC = ""
.Subject = mi_hoja.Name + ".pdf"
.Attachments.Add mi_carpeta
If DisplayEmail = False Then
End If
End With
Else
MsgBox "La hoja activa no puede estar vacía…"
Exit Sub
End If
End Sub

