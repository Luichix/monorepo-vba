VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Servicio 
   Caption         =   "PEDIDOS"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   6630
   OleObjectBlob   =   "frm_Servicio.frx":0000
End
Attribute VB_Name = "frm_Servicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long
Dim vPrecioVenta As Currency
Dim vImporte As Currency
Dim CostoUnitario As Currency
Private Sub Atencion()
Dim Fila As Long
Dim Final As Long
Dim Existencia As Long
Dim TotalExistencia As Long
Dim Comprb As Long
Dim nFactura As Long
Dim CostoTotal As Currency
Dim cUpromedio As Currency
Dim xCantidad As Long
Dim xCodigo As String
Dim xDescrip As String
Dim Estado As String


'Correlativo de la factura de venta

Comprb = Hoja93.Range("L2").Value + 1
Estado = "ACTIVO"

''Envía los datos a la hoja de ventas

If Hoja32.Visible = xlSheetVisible Then

                Hoja32.Select
                    Hoja32.Range("A2:K2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja32.Range("A3:K3").Select
                    Selection.Copy
                    Hoja32.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja32.Cells(2, 1) = CDate(frm_Factura.txtFecha)
                    Hoja32.Cells(2, 2) = Format(Time)
                    Hoja32.Cells(2, 3) = Comprb
                    Hoja32.Cells(2, 4) = Me.txt_mesa.Text
                    Hoja32.Cells(2, 5) = frm_Factura.txt_idcliente.Text
                    Hoja32.Cells(2, 6) = frm_Factura.txtCliente.Text
                    Hoja32.Cells(2, 7) = Me.TextBox4.Text
                    Hoja32.Cells(2, 8) = Hoja92.Range("G1")
                    Hoja32.Cells(2, 9) = Estado


ElseIf Hoja32.Visible = xlSheetVeryHidden Then
    Hoja32.Visible = xlSheetVisible


                Hoja32.Select
                    Hoja32.Range("A2:K2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja32.Range("A3:K3").Select
                    Selection.Copy
                    Hoja32.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja32.Cells(2, 1) = CDate(frm_Factura.txtFecha)
                    Hoja32.Cells(2, 2) = Format(Time)
                    Hoja32.Cells(2, 3) = Comprb
                    Hoja32.Cells(2, 4) = Me.txt_mesa.Text
                    Hoja32.Cells(2, 5) = frm_Factura.txt_idcliente.Text
                    Hoja32.Cells(2, 6) = frm_Factura.txtCliente.Text
                    Hoja32.Cells(2, 7) = Me.TextBox4.Text
                    Hoja32.Cells(2, 8) = Hoja92.Range("G1")
                    Hoja32.Cells(2, 9) = Estado

   Hoja32.Visible = xlSheetVeryHidden
End If

End Sub

Private Sub GrabarServicio()
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

'Correlativo de la factura de venta
Hoja93.Range("L2").Value = Hoja93.Range("L2").Value + 1
Comprb = Hoja93.Range("L2").Value

''Envía los datos a la hoja de ventas

If Hoja30.Visible = xlSheetVisible Then

                For i = 0 To frm_Factura.ListBox1.ListCount - 1
                    xCodigo = frm_Factura.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Factura.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Factura.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Factura.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Factura.ListBox1.List(i, 4) 'Importe

                Hoja30.Select
                    Hoja30.Range("A2:N2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja30.Range("A3:N3").Select
                    Selection.Copy
                    Hoja30.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja30.Cells(2, 1) = Hoja30.Cells(3, 1) + 1
                    Hoja30.Cells(2, 2) = CDate(frm_Factura.txtFecha)
                    Hoja30.Cells(2, 3) = Format(Time)
                    Hoja30.Cells(2, 4) = Comprb
                    Hoja30.Cells(2, 5) = txt_mesa.Text
                    Hoja30.Cells(2, 6) = frm_Factura.txt_idcliente.Text
                    Hoja30.Cells(2, 7) = frm_Factura.txtCliente.Text
                    Hoja30.Cells(2, 9) = xCodigo
                    Hoja30.Cells(2, 10) = xDescrip
                    Hoja30.Cells(2, 11) = xCantidad
                    Hoja30.Cells(2, 12) = vPrecioVenta
                    Hoja30.Cells(2, 14) = Hoja92.Range("G1")

                    Final = Final + 1
                Next

ElseIf Hoja30.Visible = xlSheetVeryHidden Then
    Hoja30.Visible = xlSheetVisible

                For i = 0 To frm_Factura.ListBox1.ListCount - 1
                    xCodigo = frm_Factura.ListBox1.List(i, 0) 'Codigo
                    xCantidad = frm_Factura.ListBox1.List(i, 1) 'Cantidad de Producto
                    xDescrip = frm_Factura.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = frm_Factura.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = frm_Factura.ListBox1.List(i, 4) 'Importe

                Hoja30.Select
                    Hoja30.Range("A2:N2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja30.Range("A3:N3").Select
                    Selection.Copy
                    Hoja30.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                    Hoja30.Cells(2, 1) = Hoja30.Cells(3, 1) + 1
                    Hoja30.Cells(2, 2) = CDate(frm_Factura.txtFecha)
                    Hoja30.Cells(2, 3) = Format(Time)
                    Hoja30.Cells(2, 4) = Comprb
                    Hoja30.Cells(2, 5) = txt_mesa.Text
                    Hoja30.Cells(2, 6) = frm_Factura.txt_idcliente.Text
                    Hoja30.Cells(2, 7) = frm_Factura.txtCliente.Text
                    Hoja30.Cells(2, 9) = xCodigo
                    Hoja30.Cells(2, 10) = xDescrip
                    Hoja30.Cells(2, 11) = xCantidad
                    Hoja30.Cells(2, 12) = vPrecioVenta
                    Hoja30.Cells(2, 14) = Hoja92.Range("G1")

                    Final = Final + 1
                Next

   Hoja30.Visible = xlSheetVeryHidden

End If

End Sub


Private Sub MesaServicio()

Application.ScreenUpdating = False
        Hoja30.Unprotect ""
        
        Atencion
        EstadoServicio
        GrabarServicio
        MsgBox "Servicio grabado con éxito!!!", , "Gestor de Servicios"

        Unload Me
        Unload frm_Factura
    Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True
        frm_Factura.Show
        
        Hoja30.Protect ""
Application.ScreenUpdating = True
        
End Sub

Private Sub btn_Mesa1_Click()
    Me.txt_mesa.Text = "Mesa 1"
    MesaServicio
End Sub

Private Sub btn_Mesa2_Click()
    Me.txt_mesa.Text = "Mesa 2"
    MesaServicio
End Sub

Private Sub btn_Mesa3_Click()
    Me.txt_mesa.Text = "Mesa 3"
    MesaServicio
End Sub

Private Sub btn_Mesa4_Click()
    Me.txt_mesa.Text = "Mesa 4"
    MesaServicio
End Sub

Private Sub btn_Mesa5_Click()
    Me.txt_mesa.Text = "Mesa 5"
    MesaServicio
End Sub

Private Sub btn_Mesa6_Click()
    Me.txt_mesa.Text = "Mesa 6"
    MesaServicio
End Sub

Private Sub btn_Mesa7_Click()
    Me.txt_mesa.Text = "Mesa 7"
    MesaServicio
End Sub

Private Sub btn_Mesa8_Click()
    Me.txt_mesa.Text = "Mesa 8"
    MesaServicio
End Sub

Private Sub btn_Mesa9_Click()
    Me.txt_mesa.Text = "Mesa 9"
    MesaServicio
End Sub

Private Sub btn_Otros_Click()
    Me.txt_mesa.Text = "Otros"
    MesaServicio
End Sub

'
Private Sub CommandButton1_Click()
    Unload Me
End Sub


Private Sub UserForm_Initialize()
    Me.TextBox4 = frm_Factura.txtSubtotal.Value
    Me.lbl_nFactura.Caption = "Servicio No. " & Hoja93.Range("L2").Value + 1

End Sub

Private Sub EstadoServicio()

Dim X As String
Dim encontrado As Boolean
Dim Titulo As String
Dim Estado As String

Application.ScreenUpdating = False
Titulo = "Gestor de Servicios"
Estado = "INACTIVO"

If frm_Factura.txt_nservicio = "" Then
    Exit Sub
End If

X = frm_Factura.txt_nservicio.Text
       
If Hoja32.Visible = xlSheetVisible Then

    Hoja32.Select
    Range("C1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like X Then
            encontrado = True
            Exit Do
                                 
        End If
    Loop
    If encontrado = False Then
        MsgBox "Pedido no registrado, informar a usuario administrativo", vbInformation, Titulo
        Exit Sub
    End If
    
       
           ActiveCell.Offset(0, 6) = Estado

ElseIf Hoja32.Visible = xlSheetVeryHidden Then
    Hoja32.Visible = xlSheetVisible
    
    Hoja32.Select
    Range("C1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like X Then
            encontrado = True
            Exit Do
                                 
        End If
    Loop
    If encontrado = False Then
        MsgBox "Pedido no registrado, informar a usuario administrativo", vbInformation, Titulo
        Exit Sub
    End If
    
       
           ActiveCell.Offset(0, 6) = Estado
        
    Hoja32.Visible = xlSheetVeryHidden

End If

           
           
Application.ScreenUpdating = True
End Sub
