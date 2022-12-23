Attribute VB_Name = "Importar_Hora"
Dim i As Long
Dim l As Long

Public Sub Importar_Data()
On Error Resume Next

Dim Seguridad As String

Dim Estado As String
Dim cCarpeta As String
Dim xLibroPrincipal As String
Dim xLibroSecundario As String

Dim xFilaData As Long
Dim xFinalData As Long

Dim xFilaHora As Long
Dim xFinalHora As Long

Dim NombreHoja As String
Dim BuscarHoja As Boolean
Dim Hoja As Worksheet
Dim Fila As Long
Dim Lista As Long
Dim encontrado As Boolean

Application.EnableEvents = False
Application.DisplayAlerts = False



Estado = "Espere un momento... Procesando la información"
Application.StatusBar = texto

Seguridad = Hoja83.Range("L1").Text


                                Hoja2.Unprotect (Seguridad)
                                Hoja83.Unprotect (Seguridad)

                        
                                Fila = 2
    
                                Do While Hoja83.Cells(Fila, 14) <> ""
                                Fila = Fila + 1
                                Loop
                                    
                                    Hoja83.Cells(Fila, 14) = "=NOW()"
                                    Hoja83.Cells(Fila, 14).Copy
                                    Hoja83.Cells(Fila, 14).PasteSpecial Paste:=xlPasteValues
                                    Application.CutCopyMode = False
                                  
                                    Hoja83.Cells(Fila, 15) = Hoja83.Range("G1").Text
                                    Hoja83.Cells(Fila, 16) = Hoja83.Range("H1").Text
                                  
                                Hoja83.Protect (Seguridad)


    

        xLibroPrincipal = ThisWorkbook.Name
        Workbooks(xLibroPrincipal).Activate
            
        
     cCarpeta = Application.GetOpenFilename("Reporte de Horas,*.xl*", 0, "Seleccionar el reporte a importar", , False)
         
    If cCarpeta = "Falso" Then
            Exit Sub
    ElseIf IsFileOpen(cCarpeta) Then
                MsgBox "El archivo se encuentra abierto actualmente...!", vbInformation
                Exit Sub
    Else
            Workbooks.Open (cCarpeta)
            xLibroSecundario = ActiveWorkbook.Name
            
            Workbooks(xLibroPrincipal).Activate
        
                
                Workbooks(xLibroSecundario).Activate
            Workbooks(xLibroSecundario).Worksheets("EmployeeData").Activate
            
NombreHoja = "EmployeeData"

For Each Hoja In Workbooks(xLibroSecundario).Worksheets
    If NombreHoja = Hoja.Name Then
         BuscarHoja = True
        Exit For
    Else
         BuscarHoja = False
    End If
Next

If BuscarHoja = True Then
MsgBox "Analizando los datos de importacion...!"



    If Workbooks(xLibroSecundario).Worksheets("EmployeeData").Range("P1") = Empty Then

            Workbooks(xLibroSecundario).Activate
            Workbooks(xLibroSecundario).Worksheets("EmployeeData").Activate

       
encontrado = True

If encontrado = True Then
        
            Workbooks(xLibroSecundario).SaveCopyAs ("Copia_" & Workbooks(xLibroSecundario).Name)
            
            With Workbooks(xLibroSecundario)
                .Worksheets("EmployeeData").Activate
                Modificar_Reporte
                Borrar_Filas_Vacias
            End With
                
        
             Workbooks(xLibroPrincipal).Activate
                Hoja2.Activate
                ActiveSheet.ListObjects("Tbl_tiempo").ShowTotals = False
                
             With Workbooks(xLibroSecundario)
                .Worksheets("EmployeeData").Activate
                    wFilaData = 2

                    Do While .Worksheets("EmployeeData").Cells(wFilaData, 1) <> Empty
                        wFilaData = wFilaData + 1
                    Loop
                    wFinalData = wFilaData - 1

                    .Worksheets("EmployeeData").Range(Cells(2, 1), Cells(wFinalData, 6)).Select
                        Application.CutCopyMode = False
                        Selection.Copy

            End With

            Workbooks(xLibroPrincipal).Activate
                Hoja2.Select
                Hoja2.Cells(1, 1).Select

                xFilaHora = 1

                Do While Hoja2.Cells(xFilaHora, 1) <> Empty
                        xFilaHora = xFilaHora + 1
                    Loop
                    xFinalHora = xFilaHora

                    Hoja2.Cells(xFinalHora, 1).Select

                    ActiveSheet.Paste

                 ActiveSheet.ListObjects("Tbl_tiempo").ShowTotals = True
                             
          

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
                 
            With Workbooks(xLibroSecundario)
                .Worksheets("EmployeeData").Range("P1") = "ARCHIVO IMPORTADO"
                .Close SaveChanges:=True
            End With

        
          End If
          
      ElseIf Workbooks(xLibroSecundario).Worksheets("EmployeeData").Range("P1") = "ARCHIVO IMPORTADO" Then
            MsgBox "Este reporte ya ha sido importado", vbInformation, "Gestor Administrativo"
            With Workbooks(xLibroSecundario)
                .Worksheets("EmployeeData").Range("P1") = "ARCHIVO IMPORTADO"
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
            Hoja2.Protect (Seguridad)
     

            MsgBox "Datos de importación analizados exitosamente...!", vbInformation, "Gestor Administrativo"

    
    End If

Application.EnableEvents = True
Application.DisplayAlerts = True


    Call LiberarBarra

    
End Sub

Sub LiberarBarra()
Application.StatusBar = False
End Sub


Sub Modificar_Reporte()
    Columns("D:D").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("C:C").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("C:D").Select
    Selection.ClearContents
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:L").Select
    Selection.ClearContents
    Range("A1").Select
End Sub
Sub Borrar_Filas_Vacias()

On Error Resume Next
    
        Range("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
        Range("A1").Select

        
End Sub


Sub CambiarUsuario()
form_iniciosesion.Show
End Sub


