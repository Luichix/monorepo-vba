Attribute VB_Name = "Modulo_PDF"

Option Explicit

Sub Generar_PDF()

Dim NombreArchivo, RutaArchivo As String

NombreArchivo = ActiveSheet.Name
RutaArchivo = ActiveWorkbook.Path & "\" & NombreArchivo & ".pdf"

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=RutaArchivo, _
Quality:=xlQualityStandard, IncludeDocProperties:=True, _
IgnorePrintAreas:=False, OpenAfterPublish:=True

End Sub

