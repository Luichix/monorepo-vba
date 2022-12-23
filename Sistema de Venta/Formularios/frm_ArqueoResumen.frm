VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ArqueoResumen 
   Caption         =   "RESUMEN DE CAJA"
   ClientHeight    =   10905
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   6990
   OleObjectBlob   =   "frm_ArqueoResumen.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ArqueoResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Unload Me
End Sub





Private Sub UserForm_Initialize()
Dim xHora As Date
Dim zHora As Date



xHora = Hoja9.Cells(2, 9)
zHora = Hoja9.Cells(2, 10)


Me.txtFecha = Date
Me.lbl_resumen = "No. " & Hoja9.Cells(2, 1)
Me.lbl_cierre = Hoja9.Cells(2, 21)

'If Me.lbl_cierre.Caption = "CIERRE X" Then
'    frm_ArqueoResumen.Height = 380
' Else
'  frm_ArqueoResumen.Height = 567
'
'End If


Me.txt_venta.Text = Hoja9.Cells(2, 18)
Me.txt_arqueo.Text = Hoja9.Cells(2, 19)
Me.txt_cuadre.Text = Hoja9.Cells(2, 20)
Me.txt_VentaTotal.Text = Hoja9.Cells(2, 22)
Me.txt_efectivo.Text = Hoja9.Cells(2, 23)
Me.txt_Tarjeta.Text = Hoja9.Cells(2, 24)
Me.txt_anticipo.Text = Hoja9.Cells(2, 25)
Me.txt_devolucion.Text = Hoja9.Cells(2, 26)
Me.txt_ingreso.Text = Hoja9.Cells(2, 27)
Me.txt_egreso.Text = Hoja9.Cells(2, 28)

ctrls_FormatoMoneda

Me.TextBox1 = Hoja9.Cells(2, 11) & "  -  " & Hoja9.Cells(2, 12)
Me.TextBox2 = xHora & "  -  " & zHora
Me.TextBox3 = "No. " & Hoja9.Cells(2, 7) & "  -  " & "No. " & Hoja9.Cells(2, 8)

Me.TextBox4 = Hoja92.Range("G1")

'     Application.EnableEvents = False
'    Reporte_Arqueo
'    Application.EnableEvents = True


End Sub
Private Sub Reporte_Arqueo()

If Hoja12.Visible = xlSheetVisible Then

                    Hoja12.Select
                    
                    Hoja12.Cells(11, 3) = frm_ArqueoResumen.txt_venta.Text
                    Hoja12.Cells(11, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    Hoja12.Cells(12, 3) = frm_ArqueoResumen.txt_arqueo.Text
                    Hoja12.Cells(12, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    Hoja12.Cells(13, 3) = frm_ArqueoResumen.txt_cuadre.Text
                    Hoja12.Cells(13, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    
                    Hoja12.Cells(16, 3) = frm_ArqueoResumen.TextBox1.Text
                    Hoja12.Cells(17, 3) = frm_ArqueoResumen.TextBox2.Text
                    Hoja12.Cells(18, 3) = frm_ArqueoResumen.TextBox3.Text
                    Hoja12.Cells(19, 3) = frm_ArqueoResumen.TextBox4.Text
                    
                    Hoja12.Cells(22, 2) = Hoja92.Range("G1")
                    Hoja12.Cells(23, 2) = Format(Date) & "    " & Format(Time)
                    Hoja12.Cells(24, 1) = "RESUMEN NO. " & UCase(Hoja93.Range("F2").Value + 1)
                    
                Hoja12.Select
                Hoja12.Cells(1, 1).Select
                
                    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                IgnorePrintAreas:=False
                

ElseIf Hoja12.Visible = xlSheetVeryHidden Then
        Hoja12.Visible = xlSheetVisible
        
                            Hoja12.Select

                    Hoja12.Select
                    
                    Hoja12.Cells(11, 3) = frm_ArqueoResumen.txt_venta.Text
                    Hoja12.Cells(11, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    Hoja12.Cells(12, 3) = frm_ArqueoResumen.txt_arqueo.Text
                    Hoja12.Cells(12, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    Hoja12.Cells(13, 3) = frm_ArqueoResumen.txt_cuadre.Text
                    Hoja12.Cells(13, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    
                    Hoja12.Cells(16, 3) = frm_ArqueoResumen.TextBox1.Text
                    Hoja12.Cells(17, 3) = frm_ArqueoResumen.TextBox2.Text
                    Hoja12.Cells(18, 3) = frm_ArqueoResumen.TextBox3.Text
                    Hoja12.Cells(19, 3) = frm_ArqueoResumen.TextBox4.Text
                    
                    Hoja12.Cells(22, 2) = Hoja92.Range("G1")
                    Hoja12.Cells(23, 2) = Format(Date) & "    " & Format(Time)
                    Hoja12.Cells(24, 1) = "RESUMEN NO. " & UCase(Hoja93.Range("F2").Value + 1)
                    
                Hoja12.Select
                Hoja12.Cells(1, 1).Select
                
                    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                IgnorePrintAreas:=False
                
                
    Hoja12.Visible = xlSheetVeryHidden

End If

End Sub
Public Sub ctrls_FormatoMoneda()
On Error Resume Next

Me.txt_venta.Text = FormatNumber(Me.txt_venta.Text, 2)
Me.txt_arqueo.Text = FormatNumber(Me.txt_arqueo.Text, 2)
Me.txt_cuadre.Text = FormatNumber(Me.txt_cuadre.Text, 2)
Me.txt_VentaTotal.Text = FormatNumber(Me.txt_VentaTotal.Text, 2)
Me.txt_efectivo.Text = FormatNumber(Me.txt_efectivo.Text, 2)
Me.txt_Tarjeta.Text = FormatNumber(Me.txt_Tarjeta.Text, 2)
Me.txt_anticipo.Text = FormatNumber(Me.txt_anticipo.Text, 2)
Me.txt_devolucion.Text = FormatNumber(Me.txt_devolucion.Text, 2)
Me.txt_ingreso.Text = FormatNumber(Me.txt_ingreso.Text, 2)
Me.txt_egreso.Text = FormatNumber(Me.txt_egreso.Text, 2)

End Sub
