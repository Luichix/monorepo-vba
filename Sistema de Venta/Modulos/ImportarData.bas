Attribute VB_Name = "ImportarData"
Dim i As Long

Sub Archivo_Abrir()
On Error Resume Next

Dim Estado As String
Dim cCarpeta As String
Dim ArchivoNombre As String
Dim LibroNombre As String
Dim xFilaData As Long
Dim xFinalData As Long

Estado = "Espere un momento... Procesando la información"
Application.StatusBar = texto
LibroNombre = ActiveWorkbook.Name

Application.EnableEvents = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False

     cCarpeta = Application.GetOpenFilename("Reporte de Ventas,*.xl*", 0, "Seleccionar el reporte a importar", , False)
         
        If cCarpeta = "Falso" Then
            Exit Sub
        ElseIf IsFileOpen(cCarpeta) Then
                MsgBox "El archivo se encuentra abierto actualmente...!", vbInformation
                Exit Sub
        Else
            Workbooks.Open (cCarpeta)
            ArchivoNombre = ActiveWorkbook.Name
            Hoja2.Cells(1, 1) = ArchivoNombre
            
            With Workbooks(ArchivoNombre)
                .Worksheets("Hoja1").Select
                
                xFilaData = 2

                Do While .Worksheets("Hoja1").Cells(xFilaData, 1) <> Empty
                    xFilaData = xFilaData + 1
                Loop
                xFinalData = xFilaData - 1
                
                .Worksheets("Hoja1").Range(Cells(2, 1), Cells(xFinalData, 15)).Select
                    Application.CutCopyMode = False
                    Selection.Copy
                End With
            
                    Windows(LibroNombre).Activate
                    Hoja1.Select
                    Rows("2:2").Select
                    Selection.Insert Shift:=xlDown
                    
            With Workbooks(ArchivoNombre)
                .Close SaveChanges:=True
            End With

        End If
Application.EnableEvents = True
Application.DisplayAlerts = True

    Call LiberarBarra
    MsgBox "Información procesada con exito...!"
    
End Sub

Sub LiberarBarra()
Application.StatusBar = False
End Sub

