Attribute VB_Name = "Historico"
Option Explicit
Public Sub Reporte_Historico()
Dim Fila As Long
Dim Final As Long
Dim xFila As Long
Dim Fecha As Date
Dim Colilla As Long
Dim Registro As Long
Dim Ultimo As Long
Dim Seguridad As String

Seguridad = Hoja83.Range("L1").Text

    Hoja11.Unprotect (Seguridad)
    Hoja20.Unprotect (Seguridad)
    Hoja4.Unprotect (Seguridad)
    Hoja3.Unprotect (Seguridad)

Colilla = Hoja3.Range("C2").Value
Fecha = Hoja3.Range("C2").Value

Hoja4.Activate
Hoja4.Cells(5, 1).Select

Fila = 5

Do While Hoja4.Cells(Fila, 1) <> Empty
   Fila = Fila + 1
Loop
   Final = Fila - 1
 

Ultimo = GetUltimoR(Hoja20)

If Ultimo = 1 Then
ElseIf Ultimo > 1 Then
Hoja20.Activate
Hoja20.Select
Hoja20.Rows("2:" & Ultimo).Select
Selection.Delete Shift:=xlUp
End If
    

For xFila = 5 To Final

    Hoja20.Select

    Hoja20.Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

    Hoja20.Cells(2, 1) = Fecha
    Hoja20.Cells(2, 2) = Colilla & "-" & Hoja4.Cells(xFila, 1).Text & "-" & Hoja4.Cells(xFila, 9).Text
    Hoja20.Cells(2, 3) = Hoja4.Cells(xFila, 1).Text
    Hoja20.Cells(2, 4) = Hoja4.Cells(2, 1).Text
    Hoja20.Cells(2, 5) = Hoja4.Cells(xFila, 27).Value
    Hoja20.Cells(2, 5).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 6) = Hoja4.Cells(xFila, 28).Value
    Hoja20.Cells(2, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 7) = Hoja4.Cells(xFila, 29).Value
    Hoja20.Cells(2, 7).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 8) = Hoja4.Cells(xFila, 30).Value
    Hoja20.Cells(2, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 9) = Hoja4.Cells(xFila, 31).Value
    Hoja20.Cells(2, 9).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 10) = Hoja4.Cells(xFila, 32).Value
    Hoja20.Cells(2, 10).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 11) = Hoja4.Cells(xFila, 33).Value
    Hoja20.Cells(2, 11).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 12) = Hoja4.Cells(xFila, 34).Value
    Hoja20.Cells(2, 12).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 13) = Hoja4.Cells(xFila, 35).Value
    Hoja20.Cells(2, 13).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 14) = Hoja11.Range("K2").Text


Next
    Hoja20.Select
    ActiveSheet.ListObjects("tbl_Dato").ShowTotals = False
    
    Mensualidad
    
    
    Hoja11.Protect (Seguridad)
    Hoja20.Protect (Seguridad)
    Hoja4.Protect (Seguridad)
    Hoja3.Protect (Seguridad)

End Sub
Sub Guardar_Historico()
Dim xLibroPrincipal As String
Dim xLibroSecundario As String
Dim xLibroTerceario As String
Dim xNombre As String
Dim Confirmacion As String
Dim GuardarComo As Variant
Dim mi_ventana As FileDialog
Dim mi_carpeta As String
Dim mi_referencia As String
Dim mi_archivo As String
Dim RutaArchivo As String


Application.ScreenUpdating = False

    xLibroPrincipal = ActiveWorkbook.Name

    xNombre = Hoja11.Range("J2").Text
    
    
Set mi_ventana = Application.FileDialog(msoFileDialogFolderPicker)

If mi_ventana.Show = True Then
    mi_carpeta = mi_ventana.SelectedItems(1)
    mi_referencia = mi_ventana.SelectedItems(1)
Else
    MsgBox "No se indicó la carpeta donde guardar el archivo…" _
    & vbCrLf & vbCrLf & "Operación cancelada.", _
    vbCritical, "Carpeta de almacenamiento"
    Exit Sub
End If

    mi_carpeta = mi_carpeta + "\" + xNombre + ".xlsx"
    
    If Len(Dir(mi_carpeta)) > 0 Then
    mi_archivo = MsgBox(mi_carpeta & " existente." _
    & vbCrLf & vbCrLf & "¿Desea reemplazarlo?", _
    vbYesNo + vbQuestion, "Archivo existente")
    
    On Error Resume Next
    
    If mi_archivo = vbYes Then
        Kill mi_carpeta
    Else
        MsgBox "Reemplazar el archivo existente para continuar…" _
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

RutaArchivo = mi_referencia & "\" & xNombre + ".xlsx"

        Workbooks(xLibroPrincipal).Activate
        Hoja20.Activate
        Hoja20.Select
        Hoja20.Copy
        
        xLibroSecundario = ActiveWorkbook.Name
        
        Workbooks(xLibroSecundario).SaveAs FileName:=RutaArchivo
        
        xLibroTerceario = ActiveWorkbook.Name
        
        Workbooks(xLibroTerceario).Close SaveChanges:=True
       
       Workbooks(xLibroPrincipal).Activate
       
       MsgBox ("El archivo ha sido guardado Exitosamente"), vbInformation, "Sistema de Planilla"
       
      
End Sub

Public Sub Mensualidad()
Dim Fila As Long
Dim Final As Long
Dim xFila As Long
Dim Fecha As Date
Dim Colilla As Long
Dim Registro As Long
Dim Seguridad As String
Dim encontrado As Boolean
Dim Repetido As String
Dim X As Long
Dim Y As Long
Dim Dia As Date
Dim Mes As Date
Dim Ano As Date
Dim Titulo As String

Titulo = "Gestor de Recursos Humanos"
Seguridad = Hoja83.Range("L1").Text

X = 0

    Hoja7.Unprotect (Seguridad)
    Hoja20.Unprotect (Seguridad)

Hoja20.Activate
ActiveSheet.ListObjects("tbl_Dato").ShowTotals = False


Fila = 2

Do While Hoja20.Cells(Fila, 2) <> Empty
   Fila = Fila + 1
Loop
   Final = Fila - 1


For Fila = 2 To Final

Repetido = Hoja20.Cells(Fila, 2)

Hoja7.Select
Hoja7.Range("B1").Select
Do Until IsEmpty(ActiveCell)
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.Value Like Repetido Then
        encontrado = True
        Exit Do
    End If
Loop

If encontrado = True Then
    X = X + 1

Else

    Hoja7.Select

    Hoja7.Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

    Hoja7.Cells(2, 1) = Hoja20.Cells(Fila, 1)
    Hoja7.Cells(2, 2) = Hoja20.Cells(Fila, 2)
    Hoja7.Cells(2, 3) = Hoja20.Cells(Fila, 3)
    Hoja7.Cells(2, 5) = Hoja20.Cells(Fila, 7)
    Hoja7.Cells(2, 6) = Hoja20.Cells(Fila, 8) + Hoja20.Cells(Fila, 9) + Hoja20.Cells(Fila, 10) + Hoja20.Cells(Fila, 11)
    
    Fecha = Hoja20.Cells(Fila, 1)
    
Dia = Fecha + 10
Mes = VBA.Month(Dia)
Ano = VBA.Year(Dia)

    Hoja7.Cells(2, 9) = DateSerial(Ano, Mes, 1)
    Hoja7.Cells(2, 10) = Hoja83.Range("G1")
    
End If

    
Next

    MsgBox "Se han encontrado " & X & " registros ya existentes en la hoja PAGOS"
    MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
    
    Hoja7.Protect (Seguridad)
    Hoja20.Protect (Seguridad)

    
End Sub
Public Sub Exportar_Excel()
Dim Fila As Long
Dim Final As Long
Dim xFila As Long
Dim Fecha As Date
Dim Colilla As Long
Dim Registro As Long
Dim Ultimo As Long
Dim Seguridad As String

Seguridad = Hoja83.Range("L1").Text

    Hoja11.Unprotect (Seguridad)
    Hoja20.Unprotect (Seguridad)
    Hoja4.Unprotect (Seguridad)
    Hoja3.Unprotect (Seguridad)

Colilla = Hoja3.Range("C2").Value
Fecha = Hoja3.Range("C2").Value

Hoja4.Activate
Hoja4.Cells(5, 1).Select

Fila = 5

Do While Hoja4.Cells(Fila, 1) <> Empty
   Fila = Fila + 1
Loop
   Final = Fila - 1
 

Ultimo = GetUltimoR(Hoja20)

If Ultimo = 1 Then
ElseIf Ultimo > 1 Then
Hoja20.Activate
Hoja20.Select
Hoja20.Rows("2:" & Ultimo).Select
Selection.Delete Shift:=xlUp
End If
    

For xFila = 5 To Final

    Hoja20.Select

    Hoja20.Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

    Hoja20.Cells(2, 1) = Fecha
    Hoja20.Cells(2, 2) = Colilla & "-" & Hoja4.Cells(xFila, 1).Text & "-" & Hoja4.Cells(xFila, 9).Text
    Hoja20.Cells(2, 3) = Hoja4.Cells(xFila, 1).Text
    Hoja20.Cells(2, 4) = Hoja4.Cells(2, 1).Text
    Hoja20.Cells(2, 5) = Hoja4.Cells(xFila, 27).Value
    Hoja20.Cells(2, 5).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 6) = Hoja4.Cells(xFila, 28).Value
    Hoja20.Cells(2, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 7) = Hoja4.Cells(xFila, 29).Value
    Hoja20.Cells(2, 7).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 8) = Hoja4.Cells(xFila, 30).Value
    Hoja20.Cells(2, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 9) = Hoja4.Cells(xFila, 31).Value
    Hoja20.Cells(2, 9).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 10) = Hoja4.Cells(xFila, 32).Value
    Hoja20.Cells(2, 10).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 11) = Hoja4.Cells(xFila, 33).Value
    Hoja20.Cells(2, 11).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 12) = Hoja4.Cells(xFila, 34).Value
    Hoja20.Cells(2, 12).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 13) = Hoja4.Cells(xFila, 35).Value
    Hoja20.Cells(2, 13).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja20.Cells(2, 14) = Hoja11.Range("K2").Text


Next
    Hoja20.Select
    ActiveSheet.ListObjects("tbl_Dato").ShowTotals = False
    
    Guardar_Historico
    
    
    Hoja11.Protect (Seguridad)
    Hoja20.Protect (Seguridad)
    Hoja4.Protect (Seguridad)
    Hoja3.Protect (Seguridad)

End Sub
