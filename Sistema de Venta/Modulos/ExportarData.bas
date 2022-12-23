Attribute VB_Name = "ExportarData"
Private Sub Exportar()

Dim xFila As Long
Dim xFinal As Long
Dim xDato As Long

Dim yFila As Long
Dim yFinal As Long
Dim yDato As Long

Dim zFila As Long
Dim zFinal As Long
Dim zDato As Long

Dim vFila As Long
Dim vFinal As Long
Dim vDato As Long

Dim wFila As Long
Dim wFinal As Long
Dim wDato As Long

Dim sRuta As String
Dim sNombreFolder As String
Dim sSeparador As String
Dim sRutaDestino As String
Dim sBackUp As String
Dim xLibroSecundario As String
Dim xLibroPrincipal As String

On Error GoTo Salir
   

    xLibroPrincipal = Application.ThisWorkbook.Name
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    sRuta = Application.ActiveWorkbook.Path
    sSeparador = Application.PathSeparator
    
    sBackUp = "Reporte_" & CStr(Format(Date, "yyyymmdd")) _
            & "_" & CStr(Format(Time, "hh-mm-ss")) & ".xlsx"
            
    sNombreFolder = "Reporte_" & CStr(Format(Date, "yyyymmdd"))
            
    sRutaDestino = sRuta & sSeparador & sNombreFolder
            
            If Dir(sRutaDestino, vbDirectory) = Empty Then
                MkDir (sRutaDestino)
            End If
           
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          If Hoja2.Visible = xlSheetVisible Then
          
                Hoja2.Select
                
                xDato = Hoja93.Range("B6")
                xFila = 1
                
                Do While Hoja2.Cells(xFila, 1) <> xDato
                    xFila = xFila + 1
                Loop
                xFinal = xFila - 1
                
                Hoja2.Range(Cells(1, 1), Cells(xFinal, 15)).Select
                    Application.CutCopyMode = False
                    Selection.Copy
            
            ElseIf Hoja2.Visible = xlSheetVeryHidden Then
                Hoja2.Visible = xlSheetVisible
          
                Hoja2.Select
                
                xDato = Hoja93.Range("B6")
                xFila = 1
                
                Do While Hoja2.Cells(xFila, 1) <> xDato
                    xFila = xFila + 1
                Loop
                xFinal = xFila - 1
                
                Hoja2.Range(Cells(1, 1), Cells(xFinal, 15)).Select
                    Application.CutCopyMode = False
                    Selection.Copy
            
                Hoja2.Visible = xlSheetVeryHidden
                    
            End If
            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
            
            Workbooks.Add
            ActiveSheet.Select
        
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
          
                        
            Application.ActiveWorkbook.SaveAs FileName:=sRutaDestino & sSeparador & sBackUp
            xLibroSecundario = ActiveWorkbook.Name
           
            With Workbooks(xLibroSecundario)
                .Worksheets("Hoja1").Name = "Venta"
                .Sheets.Add After:=ActiveSheet
                .Worksheets("Hoja2").Name = "Facturacion"
                .Sheets.Add After:=ActiveSheet
                .Worksheets("Hoja3").Name = "Devoluciones"
                .Sheets.Add After:=ActiveSheet
                .Worksheets("Hoja4").Name = "Devolucion"
                .Sheets.Add After:=ActiveSheet
                .Worksheets("Hoja5").Name = "Resumen"
            End With
            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            Workbooks(xLibroPrincipal).Activate
            
            If Hoja21.Visible = xlSheetVisible Then
                
                Hoja21.Select
                
                yDato = Hoja93.Range("C6")
                yFila = 1
                
                Do While Hoja21.Cells(yFila, 1) <> yDato
                    yFila = yFila + 1
                Loop
                yFinal = yFila - 1
                
                Hoja21.Range(Cells(1, 1), Cells(yFinal, 15)).Select
                    Application.CutCopyMode = False
                    Selection.Copy
                    
            ElseIf Hoja21.Visible = xlSheetVeryHidden Then
                Hoja21.Visible = xlSheetVisible
                
                Hoja21.Select
                
                yDato = Hoja93.Range("C6")
                yFila = 1
                
                Do While Hoja21.Cells(yFila, 1) <> yDato
                    yFila = yFila + 1
                Loop
                yFinal = yFila - 1
                
                Hoja21.Range(Cells(1, 1), Cells(yFinal, 15)).Select
                    Application.CutCopyMode = False
                    Selection.Copy
                
                Hoja21.Visible = xlSheetVeryHidden
                    
            End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                              
            With Workbooks(xLibroSecundario)
                .Worksheets("Facturacion").Activate
                .Sheets("Facturacion").Select
        
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
          
            End With

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            Workbooks(xLibroPrincipal).Activate
            
            If Hoja25.Visible = xlSheetVisible Then
                
                Hoja25.Select
                
                zDato = Hoja93.Range("D6")
                zFila = 1
                
                Do While Hoja25.Cells(zFila, 1) <> zDato
                    zFila = zFila + 1
                Loop
                zFinal = zFila - 1
                
                Hoja25.Range(Cells(1, 1), Cells(zFinal, 15)).Select
                    Application.CutCopyMode = False
                    Selection.Copy
                    
            ElseIf Hoja25.Visible = xlSheetVeryHidden Then
                Hoja25.Visible = xlSheetVisible
                
                Hoja25.Select
                
                zDato = Hoja93.Range("D6")
                zFila = 1
                
                Do While Hoja25.Cells(zFila, 1) <> zDato
                    zFila = zFila + 1
                Loop
                zFinal = zFila - 1
                
                Hoja25.Range(Cells(1, 1), Cells(zFinal, 15)).Select
                    Application.CutCopyMode = False
                    Selection.Copy
                
                Hoja25.Visible = xlSheetVeryHidden
                    
            End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                              
            With Workbooks(xLibroSecundario)
                .Worksheets("Devoluciones").Activate
                .Sheets("Devoluciones").Select
                
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
            End With

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            Workbooks(xLibroPrincipal).Activate
            
            If Hoja23.Visible = xlSheetVisible Then
                
                Hoja23.Select
                
                vDato = Hoja93.Range("E6")
                vFila = 1
                
                Do While Hoja23.Cells(vFila, 1) <> vDato
                    vFila = vFila + 1
                Loop
                vFinal = vFila - 1
                
                Hoja23.Range(Cells(1, 1), Cells(vFinal, 12)).Select
                    Application.CutCopyMode = False
                    Selection.Copy
                    
            ElseIf Hoja23.Visible = xlSheetVeryHidden Then
                Hoja23.Visible = xlSheetVisible
                
                Hoja23.Select
                
                vDato = Hoja93.Range("E6")
                vFila = 1
                
                Do While Hoja23.Cells(vFila, 1) <> vDato
                    vFila = vFila + 1
                Loop
                vFinal = vFila - 1
                
                Hoja23.Range(Cells(1, 1), Cells(vFinal, 12)).Select
                    Application.CutCopyMode = False
                    Selection.Copy
                
                Hoja23.Visible = xlSheetVeryHidden
                    
            End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                              
            With Workbooks(xLibroSecundario)
                .Worksheets("Devolucion").Activate
                .Sheets("Devolucion").Select
                
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
            End With


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            Workbooks(xLibroPrincipal).Activate
            
            If Hoja9.Visible = xlSheetVisible Then
                
                Hoja9.Select
                
                wDato = Hoja93.Range("F6")
                wFila = 1

                Do While Hoja9.Cells(wFila, 1) <> wDato
                    wFila = wFila + 1
                Loop
                wFinal = wFila - 1

                Hoja9.Range(Cells(1, 1), Cells(wFinal, 30)).Select
                    Application.CutCopyMode = False
                    Selection.Copy
                    
            ElseIf Hoja9.Visible = xlSheetVeryHidden Then
                Hoja9.Visible = xlSheetVisible
                
                Hoja9.Select
                
                wDato = Hoja93.Range("F6")
                wFila = 1

                Do While Hoja9.Cells(wFila, 1) <> wDato
                    wFila = wFila + 1
                Loop
                wFinal = wFila - 1

                Hoja9.Range(Cells(1, 1), Cells(wFinal, 30)).Select
                    Application.CutCopyMode = False
                    Selection.Copy

                Hoja9.Visible = xlSheetVeryHidden
                    
            End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                              
            With Workbooks(xLibroSecundario)
                .Worksheets("Resumen").Activate
                .Sheets("Resumen").Select

    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
            End With


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            With Workbooks(xLibroSecundario)
                .Close SaveChanges:=True
            End With

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Programación"
 End If
            
End Sub


