Attribute VB_Name = "Compra"

Dim i As Long
Private Sub Compra()
Dim Comprob As Long
Dim Pago As String

xComprobante = Hoja22.Range("U2").Value + 1
xPago = "EFECTIVO EN CAJA"
xEstado = "INACTIVO"

''Envía los datos a la hoja de ventas

If Hoja61.Visible = xlSheetVisible Then

                Hoja61.Select
                    Hoja61.Range("A2:L2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja61.Range("A3:L3").Select
                    Selection.Copy
                    Hoja61.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja61.Cells(2, 1) = Hoja61.Cells(3, 1) + 1
                    Hoja61.Cells(2, 2) = CDate(frm_Compra.txt_Fecha)
                    Hoja21.Cells(2, 3) = xComprobante
                    Hoja21.Cells(2, 4) = frm_Compra.txt_idproveedor.Text
                    Hoja21.Cells(2, 5) = frm_Compra.txt_proveedor.Text
                    Hoja21.Cells(2, 6) = frm_Compra.txt_documento.Text
                    Hoja21.Cells(2, 7) = frm_Compra.txt_referencia.Text
                    Hoja21.Cells(2, 8) = frm_Compra.txt_Concepto.Text
                    
                If frm_Compra.txt_Total.Value <> 0 Then
                    Hoja21.Cells(2, 9) = frm_Compra.txt_Total.Text
                ElseIf frm_Compra.txt_TotalPapel.Value <> 0 Then
                    Hoja21.Cells(2, 9) = frm_Compra.txt_TotalPapel.Text
                ElseIf frm_Compra.txt_TotalActivo.Value <> 0 Then
                    Hoja21.Cells(2, 9) = frm_Compra.txt_TotalActivo.Text
                End If
                                                     
                    Hoja21.Cells(2, 10) = xPago
                    Hoja21.Cells(2, 14) = Hoja21.Range("G1")


ElseIf Hoja61.Visible = xlSheetVeryHidden Then
    Hoja61.Visible = xlSheetVisible

                Hoja61.Select
                    Hoja61.Range("A2:L2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja61.Range("A3:L3").Select
                    Selection.Copy
                    Hoja61.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja61.Cells(2, 1) = Hoja61.Cells(3, 1) + 1
                    Hoja61.Cells(2, 2) = CDate(frm_Compra.txt_Fecha)
                    Hoja21.Cells(2, 3) = xComprobante
                    Hoja21.Cells(2, 4) = frm_Compra.txt_idproveedor.Text
                    Hoja21.Cells(2, 5) = frm_Compra.txt_proveedor.Text
                    Hoja21.Cells(2, 6) = frm_Compra.txt_documento.Text
                    Hoja21.Cells(2, 7) = frm_Compra.txt_referencia.Text
                    Hoja21.Cells(2, 8) = frm_Compra.txt_Concepto.Text
                    
                If frm_Compra.txt_Total.Value <> 0 Then
                    Hoja21.Cells(2, 9) = frm_Compra.txt_Total.Text
                ElseIf frm_Compra.txt_TotalPapel.Value <> 0 Then
                    Hoja21.Cells(2, 9) = frm_Compra.txt_TotalPapel.Text
                ElseIf frm_Compra.txt_TotalActivo.Value <> 0 Then
                    Hoja21.Cells(2, 9) = frm_Compra.txt_TotalActivo.Text
                End If
                                                     
                    Hoja21.Cells(2, 10) = xPago
                    Hoja21.Cells(2, 14) = Hoja21.Range("G1")


   Hoja61.Visible = xlSheetVeryHidden
End If

End Sub


