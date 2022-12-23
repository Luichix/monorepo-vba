VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ArqueoCaja 
   Caption         =   "GESTOR DE VENTAS"
   ClientHeight    =   9030.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   17130
   OleObjectBlob   =   "frm_ArqueoCaja.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ArqueoCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Me.TextBox2.Value = Me.TextBox2.Value & "1"
End Sub
Private Sub CommandButton2_Click()
    Me.TextBox2.Value = Me.TextBox2.Value & "2"
End Sub
Private Sub CommandButton3_Click()
    Me.TextBox2.Value = Me.TextBox2.Value & "3"
End Sub
Private Sub CommandButton4_Click()
    Me.TextBox2.Value = Me.TextBox2.Value & "4"
End Sub
Private Sub CommandButton5_Click()
    Me.TextBox2.Value = Me.TextBox2.Value & "5"
End Sub
Private Sub CommandButton6_Click()
    Me.TextBox2.Value = Me.TextBox2.Value & "6"
End Sub
Private Sub CommandButton7_Click()
    Me.TextBox2.Value = Me.TextBox2.Value & "7"
End Sub
Private Sub CommandButton8_Click()
    Me.TextBox2.Value = Me.TextBox2.Value & "8"
End Sub
Private Sub CommandButton9_Click()
    Me.TextBox2.Value = Me.TextBox2.Value & "9"
End Sub
Private Sub CommandButton10_Click()
    Me.TextBox2.Value = Me.TextBox2.Value & "0"
End Sub
Private Sub CommandButton11_Click()
    Me.TextBox2.Value = Me.TextBox2.Value & "00"
End Sub
Private Sub CommandButton12_Click()
    Me.TextBox2.Value = Empty
End Sub

Private Sub Image1_Click()
Dim Valor1 As Integer
    Me.TextBox3.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image2_Click()
    Me.TextBox4.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image3_Click()
    Me.TextBox5.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image4_Click()
    Me.TextBox6.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image5_Click()
    Me.TextBox7.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image6_Click()
    Me.TextBox8.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image7_Click()
    Me.TextBox9.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image8_Click()
    Me.TextBox10.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image9_Click()
    Me.TextBox11.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image10_Click()
    Me.TextBox12.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image11_Click()
    Me.TextBox13.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image12_Click()
    Me.TextBox14.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image13_Click()
    Me.TextBox15.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image14_Click()
    Me.TextBox16.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image15_Click()
    Me.TextBox17.Value = Me.TextBox2.Value
    Me.TextBox2 = Empty
    Valor
End Sub
Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image1.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image1.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image2.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image2.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image3.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image3.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image4.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image4.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image5.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image5_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image5.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image6.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image6_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image6.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image7.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image7_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image7.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image8.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image8_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image8.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image9_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image9.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image9_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image9.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image10_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image10.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image10_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image10.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image11_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image11.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image11_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image11.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image12_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image12.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image12_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image12.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image13_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image13.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image13_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image13.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image14_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image14.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image14_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image14.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub Image15_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image15.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image15_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image15.SpecialEffect = fmSpecialEffectFlat
End Sub

Sub Valor()
Dim Moneda25 As Double
Dim Moneda50 As Double
Dim Moneda1 As Integer
Dim Moneda5 As Integer
Dim Billete10 As Integer
Dim Billete20 As Integer
Dim Billete50 As Integer
Dim Billete100 As Integer
Dim Billete200 As Integer
Dim Billete500 As Integer
Dim Billete1000 As Integer
Dim Dolar1 As Double
Dim Dolar5 As Double
Dim Dolar10 As Double
Dim Dolar20 As Double

    If Me.TextBox14.Value = "" Then
        Moneda25 = 0
        Else
        Moneda25 = Me.TextBox14.Value * 0.25
    End If
    If Me.TextBox15.Value = "" Then
        Moneda50 = 0
        Else
        Moneda50 = Me.TextBox15.Value * 0.5
    End If
    If Me.TextBox16.Value = "" Then
        Moneda1 = 0
        Else
        Moneda1 = Me.TextBox16.Value * 1
    End If
    If Me.TextBox17.Value = "" Then
        Moneda5 = 0
        Else
        Moneda5 = Me.TextBox17.Value * 5
    End If
    If Me.TextBox3.Value = "" Then
        Billete10 = 0
        Else
        Billete10 = Me.TextBox3.Value * 10
    End If
    If Me.TextBox4.Value = "" Then
        Billete20 = 0
        Else
        Billete20 = Me.TextBox4.Value * 20
    End If
    If Me.TextBox5.Value = "" Then
        Billete50 = 0
        Else
        Billete50 = Me.TextBox5.Value * 50
    End If
    If Me.TextBox6.Value = "" Then
        Billete100 = 0
        Else
        Billete100 = Me.TextBox6.Value * 100
    End If
        If Me.TextBox7.Value = "" Then
        Billete200 = 0
        Else
        Billete200 = Me.TextBox7.Value * 200
    End If
    If Me.TextBox8.Value = "" Then
        Billete500 = 0
        Else
        Billete500 = Me.TextBox8.Value * 500
    End If
    If Me.TextBox9.Value = "" Then
        Billete1000 = 0
        Else
        Billete1000 = Me.TextBox9.Value * 1000
    End If
    If Me.TextBox10.Value = "" Then
        Dolar1 = 0
        Else
        Dolar1 = Me.TextBox10.Value * Me.txt_Cambio.Value * 1
    End If
    If Me.TextBox11.Value = "" Then
        Dolar5 = 0
        Else
        Dolar5 = Me.TextBox11.Value * Me.txt_Cambio.Value * 5
    End If
    If Me.TextBox12.Value = "" Then
        Dolar10 = 0
        Else
        Dolar10 = Me.TextBox12.Value * Me.txt_Cambio.Value * 10
    End If
    If Me.TextBox13.Value = "" Then
        Dolar20 = 0
        Else
        Dolar20 = Me.TextBox13.Value * Me.txt_Cambio.Value * 20
    End If
    

Me.TextBox1 = Moneda25 + Moneda50 + Moneda1 + Moneda5 + Billete10 + Billete20 + Billete50 + Billete100 + Billete200 + Billete500 + Billete1000 + Dolar1 + Dolar5 + Dolar10 + Dolar20

Me.TextBox1.Value = Replace(Me.TextBox1.Text, ",", ".")

End Sub

Private Sub UserForm_Initialize()
   Application.ScreenUpdating = False
    Me.lbl_arqueo.Caption = "Arqueo de Caja  No. " & Hoja93.Range("F2").Value + 1
    Me.txt_Cambio = Hoja94.Range("c8")
    Me.txt_Cambio.Value = Replace(Me.txt_Cambio.Text, ",", ".")
    Me.txtFecha = Date
    
     Application.EnableEvents = False
     If Hoja12.Visible = xlSheetVisible Then
        Hoja12.Select
        Hoja12.Range("A1", "D1").Select
         Selection.PrintOut Copies:=1, Collate:=True
     ElseIf Hoja12.Visible = xlSheetVeryHidden Then
        Hoja12.Visible = xlSheetVisible
        Hoja12.Select
        Hoja12.Range("A1", "D1").Select
        Selection.PrintOut Copies:=1, Collate:=True
        Hoja12.Visible = xlSheetVeryHidden
     End If
    Application.EnableEvents = True
    Application.ScreenUpdating = True
   
   
End Sub
Private Sub btn_Procesar_Click()
Dim Titulo As String

On Error GoTo Salir
    
Titulo = "GESTOR DE VENTAS"

Application.ScreenUpdating = False
    If Me.TextBox1.Text = Empty Then
            MsgBox "No se ha registrado ninguna monto", vbInformation, "Gestor de Ventas"
            Exit Sub
    
    End If
    

        
    If Hoja24.Visible = xlSheetVisible Then
        Hoja24.Select
        Hoja24.Cells(1, 1).Select
    End If
    
    If MsgBox("¿Son correctos los datos?" + Chr(13) + "¿Desea cargar el arqueo de caja", vbYesNo, "GESTOR DE VENTAS") = vbNo Then
        Exit Sub
    Else
    Hoja24.Unprotect ""
    CargarArqueo
    Resumen
    
        Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True
    
    
    MsgBox "Recuento procesado con éxito!!!", , "GESTOR DE VENTAS"
    Unload Me
    frm_ArqueoResumen.Show
    
    
    End If
    
        Hoja24.Protect ""
    
     Application.ScreenUpdating = True
                    
     
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If

    
End Sub

Private Sub CargarArqueo()
Dim Arqueo As Long
Dim Detalle As String

Hoja93.Range("F2").Value = Hoja93.Range("F2").Value + 1
Arqueo = Hoja93.Range("F2").Value
Detalle = "ARQUEO DE CAJA"

If Hoja24.Visible = xlSheetVisible Then
                             
                Hoja24.Select
                    Hoja24.Range("A2:AM2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja24.Range("A3:AM3").Select
                    Selection.Copy
                    Hoja24.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                    
                    Hoja24.Cells(2, 1) = Hoja24.Cells(3, 1) + 1
                    Hoja24.Cells(2, 2) = CDate(Me.txtFecha)
                    Hoja24.Cells(2, 4) = Format(Time)
                    Hoja24.Cells(2, 5) = Detalle
                    Hoja24.Cells(2, 6) = Arqueo
                    Hoja24.Cells(2, 7) = Me.txt_Cambio.Value
                    Hoja24.Cells(2, 8) = Me.TextBox14.Text
                    Hoja24.Cells(2, 10) = Me.TextBox15.Text
                    Hoja24.Cells(2, 12) = Me.TextBox16.Text
                    Hoja24.Cells(2, 14) = Me.TextBox17.Text
                    Hoja24.Cells(2, 16) = Me.TextBox3.Text
                    Hoja24.Cells(2, 18) = Me.TextBox4.Text
                    Hoja24.Cells(2, 20) = Me.TextBox5.Text
                    Hoja24.Cells(2, 22) = Me.TextBox6.Text
                    Hoja24.Cells(2, 24) = Me.TextBox7.Text
                    Hoja24.Cells(2, 26) = Me.TextBox8.Text
                    Hoja24.Cells(2, 28) = Me.TextBox9.Text
                    Hoja24.Cells(2, 30) = Me.TextBox10.Text
                    Hoja24.Cells(2, 32) = Me.TextBox11.Text
                    Hoja24.Cells(2, 34) = Me.TextBox12.Text
                    Hoja24.Cells(2, 36) = Me.TextBox13.Text
                    Hoja24.Cells(2, 38) = Me.TextBox1.Text
                    Hoja24.Cells(2, 40) = Me.lbl_cierre.Caption
                    Hoja24.Cells(2, 41) = Hoja92.Range("G1")

                                     

ElseIf Hoja24.Visible = xlSheetVeryHidden Then
        Hoja24.Visible = xlSheetVisible
        
Hoja24.Select
                    Hoja24.Range("A2:AM2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja24.Range("A3:AM3").Select
                    Selection.Copy
                    Hoja24.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                    
                    Hoja24.Cells(2, 1) = Hoja24.Cells(3, 1) + 1
                    Hoja24.Cells(2, 2) = CDate(Me.txtFecha)
                    Hoja24.Cells(2, 4) = Format(Time)
                    Hoja24.Cells(2, 5) = Detalle
                    Hoja24.Cells(2, 6) = Arqueo
                    Hoja24.Cells(2, 7) = Me.txt_Cambio.Value
                    Hoja24.Cells(2, 8) = Me.TextBox14.Text
                    Hoja24.Cells(2, 10) = Me.TextBox15.Text
                    Hoja24.Cells(2, 12) = Me.TextBox16.Text
                    Hoja24.Cells(2, 14) = Me.TextBox17.Text
                    Hoja24.Cells(2, 16) = Me.TextBox3.Text
                    Hoja24.Cells(2, 18) = Me.TextBox4.Text
                    Hoja24.Cells(2, 20) = Me.TextBox5.Text
                    Hoja24.Cells(2, 22) = Me.TextBox6.Text
                    Hoja24.Cells(2, 24) = Me.TextBox7.Text
                    Hoja24.Cells(2, 26) = Me.TextBox8.Text
                    Hoja24.Cells(2, 28) = Me.TextBox9.Text
                    Hoja24.Cells(2, 30) = Me.TextBox10.Text
                    Hoja24.Cells(2, 32) = Me.TextBox11.Text
                    Hoja24.Cells(2, 34) = Me.TextBox12.Text
                    Hoja24.Cells(2, 36) = Me.TextBox13.Text
                    Hoja24.Cells(2, 38) = Me.TextBox1.Text
                    Hoja24.Cells(2, 40) = Me.lbl_cierre.Caption
                    Hoja24.Cells(2, 41) = Hoja92.Range("G1")

     Hoja24.Visible = xlSheetVeryHidden
End If
                     
End Sub
Private Sub Resumen()

Dim nResumen As Long

Hoja93.Range("G2").Value = Hoja93.Range("G2").Value + 1
nResumen = Hoja93.Range("G2").Value

If Hoja26.Visible = xlSheetVisible Then
               
      
                Hoja9.Select
                    Hoja9.Range("A2:V2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja9.Range("A3:V3").Select
                    Selection.Copy
                    Hoja9.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                    
                    Hoja9.Cells(2, 1) = Hoja9.Cells(3, 1) + 1
                    Hoja9.Cells(2, 2) = CDate(Me.txtFecha)
                    Hoja9.Cells(2, 4) = Format(Time)
                    Hoja9.Cells(2, 5) = nResumen
                    Hoja9.Cells(2, 6) = Me.lbl_arqueo
                    Hoja9.Cells(2, 19) = Me.TextBox1.Text
                    Hoja9.Cells(2, 21) = Me.lbl_cierre.Caption
                    Hoja9.Cells(2, 29) = Hoja92.Range("G1")
                
                    Numero_resumen
                    Hora_resumen
                    Fecha_resumen
                    Comprobante_resumen
                    Ingreso_resumen
                    Egreso_resumen
                    Saldo_resumen
                    xVentaTotal_resumen
                    xEfectivo_resumen
                    xTarjeta_resumen
                    xAnticipo_resumen
                    xDevolucion_resumen
                    xIngresos_resumen
                    xEgresos_resumen
                    
                    If frm_ArqueoCaja.lbl_cierre.Caption = "CIERRE Z" Then
                        CierreZ
                    End If

                                     
    ElseIf Hoja26.Visible = xlSheetVeryHidden Then
        Hoja26.Visible = xlSheetVisible
         Hoja9.Visible = xlSheetVisible
        
                    Hoja9.Select
                    Hoja9.Range("A2:V2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja9.Range("A3:V3").Select
                    Selection.Copy
                    Hoja9.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                    
                    Hoja9.Cells(2, 1) = Hoja9.Cells(3, 1) + 1
                    Hoja9.Cells(2, 2) = CDate(Me.txtFecha)
                    Hoja9.Cells(2, 4) = Format(Time)
                    Hoja9.Cells(2, 5) = nResumen
                    Hoja9.Cells(2, 6) = Me.lbl_arqueo
                    Hoja9.Cells(2, 19) = Me.TextBox1.Text
                    Hoja9.Cells(2, 21) = Me.lbl_cierre.Caption
                    Hoja9.Cells(2, 29) = Hoja92.Range("G1")
                
                    Numero_resumen
                    Hora_resumen
                    Fecha_resumen
                    Comprobante_resumen
                    Ingreso_resumen
                    Egreso_resumen
                    Saldo_resumen
                    xVentaTotal_resumen
                    xEfectivo_resumen
                    xTarjeta_resumen
                    xAnticipo_resumen
                    xDevolucion_resumen
                    xIngresos_resumen
                    xEgresos_resumen
                    
                    
                    If frm_ArqueoCaja.lbl_cierre.Caption = "CIERRE Z" Then
                        CierreZ
                    End If

     Hoja9.Visible = xlSheetVeryHidden
     Hoja26.Visible = xlSheetVeryHidden
End If


End Sub
Private Sub Numero_resumen()
Dim Fila As Long
Dim Final As Long
Dim inumero As String
Dim fnumero As String

Hoja26.Select

Fila = 2
    
    Do While Hoja26.Cells(Fila, 1) <> Empty
        Fila = Fila + 1
    Loop
    Final = Fila - 1
       
    inumero = Hoja26.Cells(Final, 1)
    fnumero = Hoja26.Cells(2, 1)
    
    
    Hoja9.Cells(2, 7) = inumero
    Hoja9.Cells(2, 8) = fnumero
    
End Sub
Private Sub Hora_resumen()
Dim Fila As Long
Dim Final As Long
Dim ihora As Date
Dim fhora As Date

Hoja26.Select

Fila = 2
    
    Do While Hoja26.Cells(Fila, 4) <> Empty
        Fila = Fila + 1
    Loop
    Final = Fila - 1
       
    ihora = Hoja26.Cells(Final, 4)
    fhora = Hoja26.Cells(2, 4)
    
    
    Hoja9.Cells(2, 9) = ihora
    Hoja9.Cells(2, 10) = fhora
    
End Sub
Private Sub Fecha_resumen()
Dim Fila As Long
Dim Final As Long
Dim FechaI As Date
Dim FechaF As Date

Hoja26.Select

Fila = 2
    
    Do While Hoja26.Cells(Fila, 2) <> Empty
        Fila = Fila + 1
    Loop
    Final = Fila - 1
       
    FechaI = Hoja26.Cells(Final, 2)
    FechaF = Hoja26.Cells(2, 2)
    
    
    Hoja9.Cells(2, 11) = FechaI
    Hoja9.Cells(2, 12) = FechaF
    
End Sub

Private Sub Comprobante_resumen()
Dim Fila As Long
Dim Final As Long
Dim iComprobante As String
Dim fComprobante As String
Hoja26.Select

Fila = 2
    
    Do While Hoja26.Cells(Fila, 5) <> Empty
        Fila = Fila + 1
    Loop
    Final = Fila - 1
       
    iComprobante = Hoja26.Cells(Final, 5)
    fComprobante = Hoja26.Cells(2, 5)
    
    
    Hoja9.Cells(2, 13) = iComprobante
    Hoja9.Cells(2, 14) = fComprobante
    
End Sub
Private Sub Ingreso_resumen()
Dim Ingreso As Currency
Dim Egreso As Currency

Hoja26.Select

Hoja26.Range("G1").Select

Ingreso = Hoja26.Cells(Rows.Count, "G").End(xlUp)
Hoja9.Cells(2, 16) = Ingreso

End Sub
Private Sub Egreso_resumen()
Dim Egreso As Currency

Hoja26.Select
Hoja26.Range("H1").Select

Egreso = Hoja26.Cells(Rows.Count, "H").End(xlUp)
Hoja9.Cells(2, 17) = Egreso

End Sub
Private Sub Saldo_resumen()
Dim saldo As Currency

Hoja26.Select
Hoja26.Range("I1").Select

saldo = Hoja26.Cells(Rows.Count, "I").End(xlUp)
Hoja9.Cells(2, 18) = saldo

End Sub
Private Sub xVentaTotal_resumen()
Dim saldo As Currency

Hoja26.Select
Hoja26.Range("J1").Select

saldo = Hoja26.Cells(Rows.Count, "J").End(xlUp)
Hoja9.Cells(2, 22) = saldo

End Sub
Private Sub xEfectivo_resumen()
Dim saldo As Currency

Hoja26.Select
Hoja26.Range("K1").Select

saldo = Hoja26.Cells(Rows.Count, "K").End(xlUp)
Hoja9.Cells(2, 23) = saldo

End Sub
Private Sub xTarjeta_resumen()
Dim saldo As Currency

Hoja26.Select
Hoja26.Range("L1").Select

saldo = Hoja26.Cells(Rows.Count, "L").End(xlUp)
Hoja9.Cells(2, 24) = saldo

End Sub
Private Sub xAnticipo_resumen()
Dim saldo As Currency

Hoja26.Select
Hoja26.Range("M1").Select

saldo = Hoja26.Cells(Rows.Count, "M").End(xlUp)
Hoja9.Cells(2, 25) = saldo

End Sub
Private Sub xDevolucion_resumen()
Dim saldo As Currency

Hoja26.Select
Hoja26.Range("N1").Select

saldo = Hoja26.Cells(Rows.Count, "N").End(xlUp)
Hoja9.Cells(2, 26) = saldo

End Sub
Private Sub xIngresos_resumen()
Dim saldo As Currency

Hoja26.Select
Hoja26.Range("O1").Select

saldo = Hoja26.Cells(Rows.Count, "O").End(xlUp)
Hoja9.Cells(2, 27) = saldo

End Sub
Private Sub xEgresos_resumen()
Dim saldo As Currency

Hoja26.Select
Hoja26.Range("P1").Select

saldo = Hoja26.Cells(Rows.Count, "P").End(xlUp)
Hoja9.Cells(2, 28) = saldo

End Sub

Private Sub CierreZ()
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
Dim Titulo As String

Titulo = "GESTOR DE CAJA"

        Hoja26.Select
        
        If Hoja26.Cells(2, 1) = "" Then
            MsgBox "No se ha registrado ninguna transacción", vbInformation, Titulo
            Exit Sub
        End If
        
        For FiladelTotal = 2 To 10000
            If Hoja26.Cells(FiladelTotal, 1) = "" Then
                saldototal = FiladelTotal
                Exit For
            End If
        Next
        
        MsgBox "Datos Temporales limpiado con exito...!"
             
        Range(Cells(2, 1), Cells(saldototal - 1, 10)).Select
        Selection.Delete Shift:=xlUp
        
        zHistorico
End Sub
Private Sub zHistorico()
Dim Fila As Long
Dim Final As Long
Dim Detalle As String
Dim nResumen As String



Detalle = "CIERRE Z"
nResumen = "RESUMEN N° " & Hoja93.Range("G2").Value


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
                    Hoja22.Cells(2, 2) = CDate(frm_egreso.txt_Fecha)
                    Hoja22.Cells(2, 4) = Format(Time)
                    Hoja22.Cells(2, 5) = nResumen
                    Hoja22.Cells(2, 6) = Detalle
                    Hoja22.Cells(2, 8) = Hoja9.Cells(2, 18)
                    Hoja22.Cells(2, 16) = Hoja9.Cells(2, 18)
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
                    Hoja22.Cells(2, 2) = CDate(frm_egreso.txt_Fecha)
                    Hoja22.Cells(2, 4) = Format(Time)
                    Hoja22.Cells(2, 5) = nResumen
                    Hoja22.Cells(2, 6) = Detalle
                    Hoja22.Cells(2, 8) = Hoja9.Cells(2, 18)
                    Hoja22.Cells(2, 16) = Hoja9.Cells(2, 18)
                    Hoja22.Cells(2, 17) = Hoja92.Range("G1")


   Hoja22.Visible = xlSheetVeryHidden
End If

End Sub

