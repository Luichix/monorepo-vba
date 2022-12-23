VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ingreso 
   Caption         =   "GESTOR DE CAJA"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   7490
   OleObjectBlob   =   "frm_ingreso.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ingreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub UserForm_Initialize()
Me.txt_Fecha = Date
Me.lbl_ingreso = "INGRESO N° " & Hoja93.Range("H2") + 1
End Sub
Private Sub txt_monto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

KeyAscii = ValidarDecimales(txt_monto, KeyAscii)

End Sub


Private Sub Ingreso()
Dim Fila As Long
Dim Final As Long
Dim Detalle As String

'Correlativo de la factura de venta

Detalle = "FONDO DE EFECTIVO"

''Envía los datos a la hoja de ventas

If Hoja27.Visible = xlSheetVisible Then

                Hoja27.Select
                    Hoja27.Range("A2:H2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja27.Range("A3:H3").Select
                    Selection.Copy
                    Hoja27.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja27.Cells(2, 1) = Hoja27.Cells(3, 1) + 1
                    Hoja27.Cells(2, 2) = CDate(txt_Fecha)
                    Hoja27.Cells(2, 4) = Format(Time)
                    Hoja27.Cells(2, 5) = Me.lbl_ingreso.Caption
                    Hoja27.Cells(2, 6) = Detalle
                    Hoja27.Cells(2, 7) = Me.txt_monto.Text
                    Hoja27.Cells(2, 8) = Me.txt_detalle.Text
                    Hoja27.Cells(2, 9) = Hoja92.Range("G1")


ElseIf Hoja27.Visible = xlSheetVeryHidden Then
    Hoja27.Visible = xlSheetVisible

                Hoja27.Select
                    Hoja27.Range("A2:H2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja27.Range("A3:H3").Select
                    Selection.Copy
                    Hoja27.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja27.Cells(2, 1) = Hoja27.Cells(3, 1) + 1
                    Hoja27.Cells(2, 2) = CDate(txt_Fecha)
                    Hoja27.Cells(2, 4) = Format(Time)
                    Hoja27.Cells(2, 5) = Me.lbl_ingreso.Caption
                    Hoja27.Cells(2, 6) = Detalle
                    Hoja27.Cells(2, 7) = Me.txt_monto.Text
                    Hoja27.Cells(2, 8) = Me.txt_detalle.Text
                    Hoja27.Cells(2, 9) = Hoja92.Range("G1")

   Hoja27.Visible = xlSheetVeryHidden
End If


End Sub
Private Sub Reporte()
Dim Fila As Long
Dim Final As Long
Dim Existencia As Long
Dim TotalExistencia As Long
Dim Comprb As Long
Dim nFactura As Long
Dim CostoTotal As Currency
Dim cUpromedio As Currency
Dim xCantidad As Currency
Dim xCodigo As String
Dim xDescrip As String
Dim xCosto As Currency
Dim FiladelTotal As Integer
Dim ValorSaldo As Double

Final = 14

If Hoja11.Visible = xlSheetVisible Then

                    Hoja11.Select

                    Hoja11.Select
                    Hoja11.Cells(11, 1) = "ENTRADA EFECTIVO:"
                    Hoja11.Cells(12, 1) = "DETALLE DE ENTRADA:"

                    Hoja11.Cells(11, 3) = frm_ingreso.txt_monto.Text
                    Hoja11.Cells(11, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    Hoja11.Cells(13, 1) = UCase(frm_ingreso.txt_detalle.Text)
                    Hoja11.Cells(16, 2) = Hoja92.Range("G1")
                    Hoja11.Cells(17, 1) = "FECHA: " & Format(Date) & "   " & Format(Time)
                    Hoja11.Cells(18, 1) = "REFERENCIA: " & UCase(Hoja93.Range("H2").Value + 1)
                    
                Hoja11.Select
                Hoja11.Cells(1, 1).Select
                
                    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                IgnorePrintAreas:=False
                

ElseIf Hoja11.Visible = xlSheetVeryHidden Then
        Hoja11.Visible = xlSheetVisible
        
                            Hoja11.Select

                    Hoja11.Select
                    Hoja11.Cells(11, 1) = "ENTRADA EFECTIVO:"
                    Hoja11.Cells(12, 1) = "DETALLE DE ENTRADA:"

                    Hoja11.Cells(11, 3) = frm_ingreso.txt_monto.Text
                    Hoja11.Cells(11, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    Hoja11.Cells(13, 1) = UCase(frm_ingreso.txt_detalle.Text)
                    Hoja11.Cells(16, 2) = Hoja92.Range("G1")
                    Hoja11.Cells(17, 1) = "FECHA: " & Format(Date) & "   " & Format(Time)
                    Hoja11.Cells(18, 1) = "REFERENCIA: " & UCase(Hoja93.Range("H2").Value + 1)
                    
                Hoja11.Select
                Hoja11.Cells(1, 1).Select
                
                    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                IgnorePrintAreas:=False
                
                
    Hoja11.Visible = xlSheetVeryHidden

End If

End Sub
Private Sub xTemporal()
Dim Fila As Long
Dim Final As Long
Dim Detalle As String

'Correlativo de la factura de venta

Detalle = "FONDO DE EFECTIVO"

''Envía los datos a la hoja de ventas

If Hoja26.Visible = xlSheetVisible Then

                Hoja26.Select
                    Hoja26.Range("A2:Q2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja26.Range("A3:Q3").Select
                    Selection.Copy
                    Hoja26.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja26.Cells(2, 1) = Hoja22.Cells(2, 1) + 1
                    Hoja26.Cells(2, 2) = CDate(frm_ingreso.txt_Fecha)
                    Hoja26.Cells(2, 4) = Format(Time)
                    Hoja26.Cells(2, 5) = frm_ingreso.lbl_ingreso.Caption
                    Hoja26.Cells(2, 6) = Detalle
                    Hoja26.Cells(2, 7) = frm_ingreso.txt_monto.Text
                    Hoja26.Cells(2, 15) = frm_ingreso.txt_monto.Text
                    Hoja26.Cells(2, 17) = Hoja92.Range("G1")


ElseIf Hoja26.Visible = xlSheetVeryHidden Then
    Hoja26.Visible = xlSheetVisible


                Hoja26.Select
                    Hoja26.Range("A2:Q2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja26.Range("A3:Q3").Select
                    Selection.Copy
                    Hoja26.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja26.Cells(2, 1) = Hoja22.Cells(2, 1) + 1
                    Hoja26.Cells(2, 2) = CDate(frm_ingreso.txt_Fecha)
                    Hoja26.Cells(2, 4) = Format(Time)
                    Hoja26.Cells(2, 5) = frm_ingreso.lbl_ingreso.Caption
                    Hoja26.Cells(2, 6) = Detalle
                    Hoja26.Cells(2, 7) = frm_ingreso.txt_monto.Text
                    Hoja26.Cells(2, 15) = frm_ingreso.txt_monto.Text
                    Hoja26.Cells(2, 17) = Hoja92.Range("G1")

   Hoja26.Visible = xlSheetVeryHidden
End If

End Sub
Private Sub zHistorico()
Dim Fila As Long
Dim Final As Long
Dim Detalle As String

'Correlativo de la factura de venta

Detalle = "FONDO DE EFECTIVO"

''Envía los datos a la hoja de ventas

If Hoja22.Visible = xlSheetVisible Then

                Hoja22.Select
                    Hoja22.Range("A2:Q2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja22.Range("A3:Q3").Select
                    Selection.Copy
                    Hoja22.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja22.Cells(2, 1) = Hoja22.Cells(3, 1) + 1
                    Hoja22.Cells(2, 2) = CDate(frm_ingreso.txt_Fecha)
                    Hoja22.Cells(2, 4) = Format(Time)
                    Hoja22.Cells(2, 5) = frm_ingreso.lbl_ingreso.Caption
                    Hoja22.Cells(2, 6) = Detalle
                    Hoja22.Cells(2, 7) = frm_ingreso.txt_monto.Text
                    Hoja22.Cells(2, 15) = frm_ingreso.txt_monto.Text
                    Hoja22.Cells(2, 17) = Hoja92.Range("G1")


ElseIf Hoja22.Visible = xlSheetVeryHidden Then
    Hoja22.Visible = xlSheetVisible


                Hoja22.Select
                    Hoja22.Range("A2:Q2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja22.Range("A3:Q3").Select
                    Selection.Copy
                    Hoja22.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja22.Cells(2, 1) = Hoja22.Cells(3, 1) + 1
                    Hoja22.Cells(2, 2) = CDate(frm_ingreso.txt_Fecha)
                    Hoja22.Cells(2, 4) = Format(Time)
                    Hoja22.Cells(2, 5) = frm_ingreso.lbl_ingreso.Caption
                    Hoja22.Cells(2, 6) = Detalle
                    Hoja22.Cells(2, 7) = frm_ingreso.txt_monto.Text
                    Hoja22.Cells(2, 15) = frm_ingreso.txt_monto.Text
                    Hoja22.Cells(2, 17) = Hoja92.Range("G1")

   Hoja22.Visible = xlSheetVeryHidden
End If

End Sub
Private Sub btn_registrar_Click()
    If txt_monto = "" Then
        MsgBox "Debe registrar el efectivo", vbInformation, "GESTOR DE CAJA"
        txt_monto.SetFocus
        Exit Sub
    End If
        If txt_detalle = "" Then
        MsgBox "Debe registrar el detalle de ingreso", vbInformation, "GESTOR DE CAJA"
        txt_detalle.SetFocus
        Exit Sub
    End If

    If MsgBox("Son correctos los datos?", vbYesNo, "Gestor de Ventas") = vbNo Then
        Exit Sub
    Else
        Application.ScreenUpdating = False
        Hoja27.Unprotect ""
        Ingreso
        xTemporal
        zHistorico
'    Application.EnableEvents = False
'        Reporte
'    Application.EnableEvents = True
        MsgBox "Fondo de efectivo ingresado con éxito!!!", , "Gestor de Caja"
        Unload Me
End If
        Hoja27.Protect ""
        Hoja93.Range("h2") = Hoja93.Range("H2") + 1
        Application.ScreenUpdating = True
        
    Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True

End Sub
