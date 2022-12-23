Attribute VB_Name = "Cierre"
Option Explicit
Public Sub CierreZ()
Dim Lista As Long
Dim Factura As String
Dim Nota As String
Dim Fila As Long
Dim Limpiar As Long
Dim Fecha As Date
Dim Bucle As Long
Dim xFecha As Date


Factura = Hoja2.Range("D2").Text
Nota = Hoja2.Range("D3").Text
    
xFecha = frm_Cierre.txt_Fecha.Text

MsgBox ("Clik para continuar, espere un momento..!"), vbOKOnly, "ITBMS"

     
    For Bucle = 0 To 30
    
    Fecha = xFecha + Bucle
    
    Lista = 2
    
    Do While Hoja2.Cells(Lista, 1) <> ""
        Lista = Lista + 1
    Loop
    
    Lista = Lista - 1
    
Hoja7.Activate
Hoja7.ListObjects("Tbl_CierreX").ShowTotals = False
                
'FACTURA
              

              
   For Fila = 2 To Lista
   
                Hoja7.Activate
                Hoja7.Select
                Hoja7.Rows("2:2").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                
                Hoja7.Cells(2, 1) = Format(Fecha, "MM/DD/YYYY")
                Hoja7.Cells(2, 2) = Factura
                Hoja7.Cells(2, 3) = Hoja2.Cells(Fila, 1).Value
             
    Next Fila
    
'NOTA DE CREDITO
    
    For Fila = 2 To Lista
   
                Hoja7.Activate
                Hoja7.Select
                Hoja7.Rows("2:2").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja7.Cells(2, 1) = Format(Fecha, "MM/DD/YYYY")
                Hoja7.Cells(2, 2) = Nota
                Hoja7.Cells(2, 3) = Hoja2.Cells(Fila, 1).Value
                
    Next Fila
    
    Limpiar = (Lista) * 2
    
    Hoja7.Rows(Limpiar).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    Hoja7.ListObjects("Tbl_CierreX").ShowTotals = True
    
    Hoja7.Cells(1, 1).Select
    
    
    Hoja8.Activate
    Hoja8.Select
    Hoja8.Cells(2, 1) = Format(Fecha, "MM/DD/YYYY")
    
    Hoja8.Cells(1, 1).Select
    
     Range("Tabla16").Select
    Selection.ListObject.ListRows.Add (1)
    Range("Tabla1").Select
    Selection.Copy
    Range("D8").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
     Application.CutCopyMode = False
    
    Hoja8.Select
    Hoja8.Activate
    
    Hoja8.Cells(2, 1).Select
    Hoja8.Cells(2, 1) = Hoja8.Cells(2, 1) + 1
    
    Hoja8.Range("A1").Select
    
    
    Next Bucle

    
End Sub
Sub ClearTable()
On Error Resume Next

Range("Tabla16").Select
If Not ActiveCell.ListObject Is Nothing Then
    ActiveCell.ListObject.DataBodyRange.Rows.ClearContents
End If
    
Range("Tabla16").Select
If Not ActiveCell.ListObject Is Nothing Then
    ActiveCell.ListObject.DataBodyRange.Delete
End If
    
    Hoja8.Cells(1, 1).Select
    
  On Error GoTo 0
    
End Sub

Public Sub CierreX()
Dim Lista As Long
Dim Factura As String
Dim Nota As String
Dim Fila As Long
Dim Limpiar As Long
Dim Fecha As Date



Factura = Hoja2.Range("D2").Text
Nota = Hoja2.Range("D3").Text
    
Fecha = frm_Cierre.txt_Fecha.Text

    Lista = 2
    
    Do While Hoja2.Cells(Lista, 1) <> ""
        Lista = Lista + 1
    Loop
    
    Lista = Lista - 1
    
Hoja7.Activate
Hoja7.ListObjects("Tbl_CierreX").ShowTotals = False
                
'FACTURA
              

              
   For Fila = 2 To Lista
   
                Hoja7.Activate
                Hoja7.Select
                Hoja7.Rows("2:2").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                
                Hoja7.Cells(2, 1) = Format(Fecha, "MM/DD/YYYY")
                Hoja7.Cells(2, 2) = Factura
                Hoja7.Cells(2, 3) = Hoja2.Cells(Fila, 1).Value
             
    Next Fila
    
'NOTA DE CREDITO
    
    For Fila = 2 To Lista
   
                Hoja7.Activate
                Hoja7.Select
                Hoja7.Rows("2:2").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja7.Cells(2, 1) = Format(Fecha, "MM/DD/YYYY")
                Hoja7.Cells(2, 2) = Nota
                Hoja7.Cells(2, 3) = Hoja2.Cells(Fila, 1).Value
                
    Next Fila
    
    Limpiar = (Lista) * 2
    
    Hoja7.Rows(Limpiar).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    Hoja7.ListObjects("Tbl_CierreX").ShowTotals = True
    
    Hoja7.Cells(1, 1).Select
    
    
    Hoja8.Activate
    Hoja8.Select
    Hoja8.Cells(2, 1) = Format(Fecha, "MM/DD/YYYY")
    
    Hoja8.Cells(1, 1).Select
    
    
End Sub
