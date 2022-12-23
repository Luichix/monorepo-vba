VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Factura 
   Caption         =   "REGISTRO DE ENCARGOS"
   ClientHeight    =   7160
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   9490.001
   OleObjectBlob   =   "frm_Factura.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Factura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Long
Dim vPrecioVenta As Currency
Dim vImporte As Currency
Dim CostoUnitario As Currency

Private Sub ProcesarFactura()
Dim Fila As Long
Dim Final As Long
Dim Existencia As Long
Dim TotalExistencia As Long
Dim Comprb As Long
Dim nFactura As Long
Dim CostoTotal As Currency
Dim cUpromedio As Currency
Dim xCantidad As Long
Dim xCodigo
Dim xDescrip As String
Dim xCosto As Currency


''Aquí manejo el correlativo del ID del comprobante según sistema
'Hoja22.Range("E2").Value = Hoja7.Range("E2").Value + 1
'Comprb = Hoja7.Range("E2").Value

'Correlativo de la factura de venta
Hoja22.Range("C2").Value = Hoja22.Range("C2").Value + 1
Comprb = Hoja22.Range("C2").Value


    'Determina el final del listado de salidas
    Final = GetNuevoR(Hoja26)


        'Envía los datos a la hoja de salidas


                For i = 0 To Me.ListBox1.ListCount - 1
                    xCantidad = Me.ListBox1.List(i, 0) 'Cantidad
                    xCodigo = Me.ListBox1.List(i, 1) 'Código de Producto
                    xDescrip = Me.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    vPrecioVenta = Me.ListBox1.List(i, 3) 'Precio Venta
                    vImporte = Me.ListBox1.List(i, 4) 'Importe


                    CostoTotal = Val("-" & Me.ListBox1.List(i, 0)) * CostoUnitario 'Obtengo el costo total

                    Hoja26.Cells(Final, 5) = "Comprb-" & Comprb
                    Hoja26.Cells(Final, 6) = xCodigo
                    Hoja26.Cells(Final, 1) = CDate(Me.txtFecha)
                    Hoja26.Cells(Final, 8) = xCantidad 'Cantidad
                    Hoja26.Cells(Final, 9) = vPrecioVenta 'Precio Venta

                    Hoja26.Cells(Final, 4) = Me.txtCliente

'''''                    Hoja27.Cells(Final, 3) = xDescrip
'''''                    Hoja27.Cells(Final, 5) = nFactura
'''''                    Hoja4.Cells(Final, 8) = "Facturación"
'''''                    Hoja4.Cells(Final, 10) = CostoUnitario 'Costo Unitario
'''''                    Hoja4.Cells(Final, 11) = CostoTotal
'''''
'''''                    xCosto = Replace(Hoja4.Cells(Final, 11), "-", "")
'''''
'''''                    Hoja4.Cells(Final, 13) = Hoja8.Range("G1") 'Usuario responsable de la operación
'''''
'''''                    Hoja4.Cells(Final, 16) = vImporte 'Importe
'''''                    Hoja4.Cells(Final, 17) = Hoja8.Range("G1") 'Usuario responsable de la operación
'''''
'''''                    descontarCosto xCodigo, xCosto

                    Final = Final + 1


                Next



End Sub

Private Sub btn_FechaFact_Click()
banderaCalendario = 1
    Call LanzarCalendario(Me, "txtFecha")
End Sub


Private Sub btn_Procesar_Click()

On Error GoTo Salir

   
    If Me.txtCliente.Text = Empty Or _
        Me.txtMail.Text = Empty Or _
        Me.txtNIT.Text = Empty Then

            MsgBox "Hay campos vacíos en la compra", , "Gestor Administrativo"
            Exit Sub
    
    End If

'        Me.txtNRF.Text = Empty Or _


If MsgBox("Son correctos los datos?" + Chr(13) + "Desea procesar la factura?", vbYesNo, "Gestor de Inventarios") = vbNo Then
        Exit Sub
    Else
        RegistrarCliente
        ProcesarFactura
        MsgBox "Factura procesada con éxito!!!", , "Gestor de Inventarios"
        Unload Me
End If

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If

End Sub

Private Sub RegistrarCliente()
    Dim Fila As Long
    Dim Final As Long
    

        'Determina el final del listado de Clientes
        Final = GetNuevoR(Hoja4)
        
        'Validación para impedir Clientes repetidos
        For Fila = 2 To Final
            If Hoja4.Cells(Fila, 1) = UCase(Me.txtCliente.Text) Then
                Exit Sub
                Exit For
            End If
        Next
        
         
                'Envía los datos a la hoja de Clientes
                Hoja4.Cells(Final, 1) = UCase(Me.txtCliente.Text)

                Hoja4.Cells(Final, 3) = Me.txtNIT.Text
                Hoja4.Cells(Final, 4) = Me.txtMail.Text
 
End Sub




Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Me.ListBox1.ListIndex = -1 ' Eliminar la "barra de selección"
End Sub
Private Sub lbl_BuscarCliente_Click()
frm_Clientes.Show
End Sub
Private Sub lbl_BuscarCliente_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub



Private Sub lblEliminarItem_Click()
Me.EliminarItem
Me.ctrls_FormatoMoneda
End Sub
Private Sub lblEliminarItem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub lblProductos_Click()
frm_ProductoAFacturar.Show
End Sub
Private Sub lblProductos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Me.ListBox1.ListIndex = -1 ' Eliminar la "barra de selección"
SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub UserForm_Activate()
Me.txtFecha = Date
End Sub
Private Sub UserForm_Initialize()

Me.lbl_nFactura.Caption = "No. " & Hoja22.Range("C2").Value + 1 'Llamamos el número de la factura


With ListBox1
    .ColumnCount = 5
    .ColumnWidths = "50 pt;75 pt;165 pt;70 pt;70 pt" ' Unidades de medida, 72 pt(puntos)=1 Pulgada

End With

End Sub

Public Sub sumarImporte()

'Suma la columna de los importes
Dim xImporte As Currency
Dim IvaPorcentaje As Byte
Dim sTotal As Currency
Dim xIVA As Currency

vSeparadorMiles = Hoja27.Range("C5")

If vSeparadorMiles = "." Then
    xDecimal = ","
        ElseIf vSeparadorMiles = "," Then
            xDecimal = "."
End If



sTotal = 0
        For i = 0 To Me.ListBox1.ListCount - 1
            Me.ListBox1.List(i, 4) = _
            Replace(Me.ListBox1.List(i, 4), Application.ThousandsSeparator, "")  'Aquí elimino el separador de miles en el Importe
            Me.ListBox1.List(i, 4) = _
            Replace(Me.ListBox1.List(i, 4), ",", ".") 'Ahora sustituyo la coma decimal por el punto decimal, para poder hacer la sumatoria con la variable sTotal, ya que con la coma decimal, no se suman los decimales

            sTotal = sTotal + Val(Me.ListBox1.List(i, 4)) 'Aquí hago la sumatoria del importe, utilizando el punto decimal

            Me.ListBox1.List(i, 4) = _
            Replace(Me.ListBox1.List(i, 4), ".", Application.DecimalSeparator)  'Aquí devuelvo el formato decimal para que no afecte al ListBox
            Me.ListBox1.List(i, 4) = FormatNumber(Me.ListBox1.List(i, 4), 2) 'Aqui doy formato de moneda para que aparezcan los separadores de miles y decimales
        Next

Me.txtSubtotal.Text = sTotal

IvaPorcentaje = Hoja27.Range("C6")


            If sTotal > 0 Then ' aqui se hacen los calculos para el subtotal, iva y total

                    Me.txtIVA.Text = (sTotal / 100) * IvaPorcentaje
                    xIVA = Me.txtIVA.Text
                    Me.txtTotal.Text = sTotal + xIVA
                    Me.txtLetras.Text = UCase(cMoneda(Me.txtTotal.Text))
                    
                Else
                    Me.txtSubtotal.Text = Empty
                    Me.txtIVA.Text = Empty
                    Me.txtTotal.Text = Empty
                    Me.txtLetras.Text = Empty
            End If

End Sub

Public Sub AgregarItems()
'Agrega los items al listbox

        If frm_ProductoAFacturar.ComboBox1.Text = "" Then MsgBox ("Elija un código de producto"): Exit Sub
        
    

        If Trim(frm_ProductoAFacturar.txtCantidad.Text) = "" Then MsgBox ("Debe ingresar la cantidad"): Exit Sub
       
        With frm_Factura
            .ListBox1.AddItem Val(frm_ProductoAFacturar.txtCantidad.Text)
            .ListBox1.List(i, 1) = frm_ProductoAFacturar.ComboBox1.Text 'Código del producto
            .ListBox1.List(i, 2) = frm_ProductoAFacturar.txt_nombre.Text 'Nombre del producto
            .ListBox1.List(i, 3) = frm_ProductoAFacturar.txt_PrecioV.Text 'Precio Venta
            .ListBox1.List(i, 4) = frm_ProductoAFacturar.txtImporte.Text
            
            
            
            
            i = i + 1
        End With
    
        sumarImporte
    

        With frm_ProductoAFacturar
            .ComboBox1.ListIndex = -1
            .txt_nombre = ""
            .txtCantidad = ""
            .txt_PrecioV = ""
           
        End With

End Sub
Public Sub EliminarItem()
' Elimina el item seleccionado y resta el importe de la columna de importes

    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "Seleccionar un producto para eliminar", vbInformation
        Exit Sub
    End If

Me.ListBox1.RemoveItem (ListBox1.ListIndex)
Me.ListBox1.ListIndex = -1 ' Eliminar la "barra de selección"

Me.sumarImporte
            
End Sub

Public Sub ctrls_FormatoMoneda()
On Error Resume Next
    Me.txtSubtotal.Text = FormatNumber(Me.txtSubtotal.Text, 2)
    Me.txtTotal.Text = FormatNumber(Me.txtTotal.Text, 2)
End Sub
