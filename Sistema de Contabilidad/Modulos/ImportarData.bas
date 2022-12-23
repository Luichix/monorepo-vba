Attribute VB_Name = "ImportarData"
Dim i As Long

Public Sub Importar_Data()
On Error Resume Next

Dim Estado As String
Dim cCarpeta As String
Dim xLibroPrincipal As String
Dim xLibroSecundario As String

Dim xFilaResumen As Long
Dim xFinalResumen As Long

Dim xFilaData As Long
Dim xFinalData As Long

Dim yFilaData As Long
Dim yFinalData As Long

Dim zFilaData As Long
Dim zFinalData As Long

Dim vFilaData As Long
Dim vFinalData As Long

Dim wFilaData As Long
Dim wFinalData As Long

Dim nombreHoja As String
Dim BuscarHoja As Boolean
Dim Hoja As Worksheet

Estado = "Espere un momento... Procesando la información"
Application.StatusBar = texto


Application.EnableEvents = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False

        xLibroPrincipal = ThisWorkbook.Name
        Workbooks(xLibroPrincipal).Activate
            Hoja27.Range("B10").Text = xLibroPrincipal
        
     cCarpeta = Application.GetOpenFilename("Reporte de Ventas,*.xl*", 0, "Seleccionar el reporte a importar", , False)
         
    If cCarpeta = "Falso" Then
            Exit Sub
    ElseIf IsFileOpen(cCarpeta) Then
                MsgBox "El archivo se encuentra abierto actualmente...!", vbInformation
                Exit Sub
    Else
            Workbooks.Open (cCarpeta)
            xLibroSecundario = ActiveWorkbook.Name
            
            Workbooks(xLibroPrincipal).Activate
                Hoja27.Range("B11").Text = xLibroSecundario
                
                Workbooks(xLibroSecundario).Activate
            Workbooks(xLibroSecundario).Worksheets("Resumen").Activate
            
nombreHoja = "Resumen"

For Each Hoja In Workbooks(xLibroSecundario).Worksheets
    If nombreHoja = Hoja.Name Then
         BuscarHoja = True
        Exit For
    Else
         BuscarHoja = False
    End If
Next

If BuscarHoja = True Then
MsgBox "Analizando los datos de importacion...!"


 If Workbooks(xLibroSecundario).Worksheets("Resumen").Range("AI1") = "REPORTE ACTIVO" Then

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Workbooks(xLibroSecundario).Activate
            Workbooks(xLibroSecundario).Worksheets("Venta").Activate

    If Workbooks(xLibroSecundario).Worksheets("Venta").Cells(2, 1) = Empty Then
         MsgBox "El reporte de venta no poseia ningun registro"


    ElseIf Workbooks(xLibroSecundario).Worksheets("Venta").Cells(2, 1) <> Empty Then

            With Workbooks(xLibroSecundario)
                .Worksheets("Venta").Activate
                    xFilaData = 2

                    Do While .Worksheets("Venta").Cells(xFilaData, 1) <> Empty
                        xFilaData = xFilaData + 1
                    Loop
                    xFinalData = xFilaData - 1

                    .Worksheets("Venta").Range(Cells(2, 1), Cells(xFinalData, 15)).Select
                        Application.CutCopyMode = False
                        Selection.Copy


            End With

            Workbooks(xLibroPrincipal).Activate
                Hoja2.Select
                    Rows("2:2").Select
                    Selection.Insert Shift:=xlDown

    End If

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Workbooks(xLibroSecundario).Activate
             Workbooks(xLibroSecundario).Worksheets("Facturacion").Activate

    If Workbooks(xLibroSecundario).Worksheets("Facturacion").Cells(2, 1) = Empty Then
         MsgBox "El reporte de facturación no poseia ningun registro"


    ElseIf Workbooks(xLibroSecundario).Worksheets("Facturacion").Cells(2, 1) <> Empty Then

            With Workbooks(xLibroSecundario)
                .Worksheets("Facturacion").Activate
                    yFilaData = 2

                    Do While .Worksheets("Facturacion").Cells(yFilaData, 1) <> Empty
                        yFilaData = yFilaData + 1
                    Loop
                    yFinalData = yFilaData - 1

                    .Worksheets("Facturacion").Range(Cells(2, 1), Cells(yFinalData, 15)).Select
                        Application.CutCopyMode = False
                        Selection.Copy

            End With

            Workbooks(xLibroPrincipal).Activate
                Hoja9.Select
                    Rows("2:2").Select
                    Selection.Insert Shift:=xlDown

    End If

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

             Workbooks(xLibroSecundario).Activate
             Workbooks(xLibroSecundario).Worksheets("Devoluciones").Activate

    If Workbooks(xLibroSecundario).Worksheets("Devoluciones").Cells(2, 1) = Empty Then
         MsgBox "El reporte de devoluciones no poseia ningun registro"


    ElseIf Workbooks(xLibroSecundario).Worksheets("Devoluciones").Cells(2, 1) <> Empty Then

            With Workbooks(xLibroSecundario)
                .Worksheets("Devoluciones").Activate
                    zFilaData = 2

                    Do While .Worksheets("Devoluciones").Cells(zFilaData, 1) <> Empty
                        zFilaData = zFilaData + 1
                    Loop
                    zFinalData = zFilaData - 1

                    .Worksheets("Devoluciones").Range(Cells(2, 1), Cells(zFinalData, 15)).Select
                        Application.CutCopyMode = False
                        Selection.Copy

            End With


            Workbooks(xLibroPrincipal).Activate
                Hoja3.Select
                    Rows("2:2").Select
                    Selection.Insert Shift:=xlDown

    End If

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              Workbooks(xLibroSecundario).Activate
             Workbooks(xLibroSecundario).Worksheets("Devolucion").Activate

    If Workbooks(xLibroSecundario).Worksheets("Devolucion").Cells(2, 1) = Empty Then
         MsgBox "El reporte de devolucion no poseia ningun registro"


    ElseIf Workbooks(xLibroSecundario).Worksheets("Devolucion").Cells(2, 1) <> Empty Then

            With Workbooks(xLibroSecundario)
                .Worksheets("Devolucion").Activate
                    vFilaData = 2

                    Do While .Worksheets("Devolucion").Cells(vFilaData, 1) <> Empty
                        vFilaData = vFilaData + 1
                    Loop
                    vFinalData = vFilaData - 1

                    .Worksheets("Devolucion").Range(Cells(2, 1), Cells(vFinalData, 12)).Select
                        Application.CutCopyMode = False
                        Selection.Copy

            End With


            Workbooks(xLibroPrincipal).Activate
                Hoja26.Select
                    Rows("2:2").Select
                    Selection.Insert Shift:=xlDown

    End If

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

             Workbooks(xLibroPrincipal).Activate
                Hoja28.Activate
                ActiveSheet.ListObjects("tbl_Resumen").ShowTotals = False

            Workbooks(xLibroSecundario).Activate
            Workbooks(xLibroSecundario).Worksheets("Resumen").Activate

    If Workbooks(xLibroSecundario).Worksheets("Resumen").Cells(2, 1) = Empty Then
         MsgBox "El reporte cierre no poseia ningun registro"


    ElseIf Workbooks(xLibroSecundario).Worksheets("Resumen").Cells(2, 1) <> Empty Then

            With Workbooks(xLibroSecundario)
                .Worksheets("Resumen").Activate
                    wFilaData = 2

                    Do While .Worksheets("Resumen").Cells(wFilaData, 1) <> Empty
                        wFilaData = wFilaData + 1
                    Loop
                    wFinalData = wFilaData - 1

                    .Worksheets("Resumen").Range(Cells(2, 1), Cells(wFinalData, 32)).Select
                        Application.CutCopyMode = False
                        Selection.Copy

            End With

            Workbooks(xLibroPrincipal).Activate
                Hoja28.Select
                Hoja28.Cells(1, 1).Select

                xFilaResumen = 1

                Do While Hoja28.Cells(xFilaResumen, 1) <> Empty
                        xFilaResumen = xFilaResumen + 1
                    Loop
                    xFinalResumen = xFilaResumen

                    Hoja28.Cells(xFilaResumen, 1).Select


                    ActiveSheet.Paste

                 ActiveSheet.ListObjects("tbl_Resumen").ShowTotals = True
                 
            With Workbooks(xLibroSecundario)
                .Worksheets("Resumen").Range("AI1") = "REPORTE INACTIVO"
                .Close SaveChanges:=True
            End With

    End If

      ElseIf Workbooks(xLibroSecundario).Worksheets("Resumen").Range("AI1") = "REPORTE INACTIVO" Then
            MsgBox "Este reporte ya ha sido importado", vbInformation, "Gestor Administrativo"
            With Workbooks(xLibroSecundario)
                .Worksheets("Resumen").Range("AI1") = "REPORTE INACTIVO"
                .Close SaveChanges:=True
            End With


     End If
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
    ElseIf BuscarHoja = False Then
    
    MsgBox "Este archivo no corresponde a los reportes a importar...!", vbExclamation, "Gestor Administrativo"
    
    End If
    
            With Workbooks(xLibroSecundario)
                .Close SaveChanges:=True
            End With

            Workbooks(xLibroPrincipal).Activate

            MsgBox "Datos de importación analizados exitosamente...!", vbInformation, "Gestor Administrativo"

    
    End If

Application.EnableEvents = True
Application.DisplayAlerts = True
Application.ScreenUpdating = True

    Call LiberarBarra

    
End Sub

Sub LiberarBarra()
Application.StatusBar = False
End Sub

