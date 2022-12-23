VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_fCompras 
   Caption         =   "ENTRADAS DE INVENTARIO"
   ClientHeight    =   8790.001
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   11890
   OleObjectBlob   =   "frm_fCompras.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_fCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Long
Dim vCostoU As Currency
Dim vCostoT As Currency
Dim vPrecioVenta As Currency
Private Sub ProcesarCompra()
Dim Fila As Long
Dim Final As Long
Dim nFactura As Long
Dim xCantidad As Double
Dim xDescrip As String
Dim xCategoria As String

'Correlativo de la factura de venta
Hoja93.Range("A2").Value = Hoja93.Range("A2").Value + 1
nFactura = Hoja93.Range("A2").Value

        'Envía los datos a la hoja de ENTRADAS

 If Hoja3.Visible = xlSheetVisible Then
                For i = 0 To Me.ListBox1.ListCount - 1
                    xCantidad = Me.ListBox1.List(i, 0) 'Cantidad
                    xDescrip = Me.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    xCategoria = Me.ListBox1.List(i, 3) 'Clase de Producto
                    vCostoU = Me.ListBox1.List(i, 4) 'Costo Unitario

                Hoja3.Select
                    Hoja3.Range("A2:L2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja3.Range("A3:L3").Select
                    Selection.Copy
                    Hoja3.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                    
                    Hoja3.Cells(2, 1) = CDate(Me.txtFecha) 'Fecha
                    Hoja3.Cells(2, 3) = nFactura   'Número de Factura
                    Hoja3.Cells(2, 4) = Me.txtProveedor 'Proveedor
                    Hoja3.Cells(2, 6) = xDescrip 'Descripción
                    Hoja3.Cells(2, 7) = xCantidad 'Cantidad
                    Hoja3.Cells(2, 9) = vCostoU 'Costo Unitario
                    Hoja3.Cells(2, 11) = xCategoria 'Clase de Producto
                    Hoja3.Cells(2, 12) = Hoja92.Range("G1")
                                     
                    Final = Final + 1
                    
                Next

ElseIf Hoja3.Visible = xlSheetVeryHidden Then
        Hoja3.Visible = xlSheetVisible
        
          For i = 0 To Me.ListBox1.ListCount - 1
                    xCantidad = Me.ListBox1.List(i, 0) 'Cantidad
                    xDescrip = Me.ListBox1.List(i, 2) 'Nombre del Producto o Descripción
                    xCategoria = Me.ListBox1.List(i, 3) 'Clase de Producto
                    vCostoU = Me.ListBox1.List(i, 4) 'Costo Unitario
                    
                Hoja3.Select
                    Hoja3.Range("A2:L2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja3.Range("A3:L3").Select
                    Selection.Copy
                    Hoja3.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                    
                    Hoja3.Cells(2, 1) = CDate(Me.txtFecha) 'Fecha
                    Hoja3.Cells(2, 3) = nFactura   'Número de Factura
                    Hoja3.Cells(2, 4) = Me.txtProveedor 'Proveedor
                    Hoja3.Cells(2, 6) = xDescrip 'Descripción
                    Hoja3.Cells(2, 7) = xCantidad 'Cantidad
                    Hoja3.Cells(2, 9) = vCostoU 'Costo Unitario
                    Hoja3.Cells(2, 11) = xCategoria 'Clase de Producto
                    Hoja3.Cells(2, 12) = Hoja92.Range("G1")
                                     
                    Final = Final + 1
                    
                Next

     Hoja3.Visible = xlSheetVeryHidden
End If
                     
End Sub
Private Sub btn_Cancelar_Click()
Unload Me
End Sub

Private Sub btn_FechaFact_Click()
banderaCalendario = 2
    Call LanzarCalendario(Me, "txtFecha")
End Sub


Private Sub btn_Procesar_Click()
    
On Error GoTo Salir

Application.ScreenUpdating = False

    If Me.txtProveedor.Text = Empty Or _
        Me.txtUBIC.Text = Empty Or _
        Me.txtNRF.Text = Empty Or _
        Me.txtTELF.Text = Empty Then

            MsgBox "Hay campos vacíos en la compra", , "Gestor Administrativo"
            Exit Sub
    
    End If
    

If MsgBox("Son correctos los datos?" + Chr(13) + "Desea procesar la compra?", vbYesNo, "Gestor de Inventarios") = vbNo Then
        Exit Sub
    Else
     Hoja3.Unprotect ""
     Hoja8.Unprotect ""
        RegistrarProveedor
        ProcesarCompra
        MsgBox "Compra procesada con éxito!!!", , "Gestor de Inventarios"
        Unload Me
End If
     Hoja3.Protect ""
     Hoja8.Protect ""
     
    If Hoja3.Visible = xlSheetVisible Then
     Hoja3.Select
     Hoja3.Cells(1, 1).Select
    End If
    
     Application.ScreenUpdating = True
         Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If
 

End Sub

Private Sub RegistrarProveedor()
    Dim Fila As Long
    Dim Final As Long
    

        'Determina el final del listado de Proveedores
        Final = GetNuevoR(Hoja8)
        
        'Validación para impedir Proveedores repetidos
        For Fila = 2 To Final
            If Hoja8.Cells(Fila, 1) = UCase(Me.txtProveedor.Text) Then
                Exit Sub
                Exit For
            End If
        Next
        
                'Envía los datos a la hoja de Proveedores
                Hoja8.Cells(Final, 1) = UCase(Me.txtProveedor.Text)
                Hoja8.Cells(Final, 2) = "'" & Me.txtNRF.Text
                Hoja8.Cells(Final, 3) = "'" & Me.txtTELF.Text
                Hoja8.Cells(Final, 4) = Me.txtUBIC.Text

  
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.ListBox1.ListIndex = -1 ' Eliminar la "barra de selección"
End Sub

Private Sub btn_Cancelar_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = ALT + 43 Then
    frm_ProductoAComprar.Show
    End If
End Sub
Private Sub btn_Procesar_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = ALT + 43 Then
    frm_ProductoAComprar.Show
    End If
End Sub

Private Sub lbl_BuscarProveedor_Click()
    banderaProveedores = 1
    Call LanzarListadoProveedores(Me, "lbl_LanzarListadoProveedores")
End Sub

Private Sub lbl_BuscarProveedor_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub lblEliminarItem_Click()
Me.EliminarItem
End Sub

Private Sub lblEliminarItem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub lblProductos_Click()
frm_ProductoAComprar.Show
End Sub

Private Sub lblProductos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.ListBox1.ListIndex = -1 ' Eliminar la "barra de selección"
SetCursor LoadCursor(0, IDC_HAND)
End Sub


Private Sub UserForm_Activate()
Me.txtFecha = Date
End Sub

Private Sub UserForm_Initialize()

Me.lbl_nFactura.Caption = "No. " & Hoja93.Range("A2").Value + 1 'Llamamos el número de la compra

With ListBox1
    .ColumnCount = 6
    .ColumnWidths = "70 pt;80 pt;150 pt;110 pt;75 pt;40 pt" ' Unidades de medida, 72 pt(puntos)=1 Pulgada

End With

End Sub

Public Sub sumarImporte()
'Suma la columna de los importes
Dim xImporte As Currency
Dim sTotal As Currency
Dim xIVA As Currency

vSeparadorMiles = Hoja5.Range("C5")

If vSeparadorMiles = "." Then
    xDecimal = ","
        ElseIf vSeparadorMiles = "," Then
            xDecimal = "."
End If
        
sTotal = 0
        For i = 0 To Me.ListBox1.ListCount - 1
            Me.ListBox1.List(i, 5) = _
            Replace(Me.ListBox1.List(i, 5), Application.ThousandsSeparator, "")  'Aquí elimino el separador de miles en el Importe
            Me.ListBox1.List(i, 5) = _
            Replace(Me.ListBox1.List(i, 5), ",", ".") 'Ahora sustituyo la coma decimal por el punt decimal, para poder hacer la sumatoria con la variable sTotal, ya que con la coma decimal, no se suman los decimales
            
            sTotal = sTotal + Val(Me.ListBox1.List(i, 5)) 'Aquí hago la sumatoria del importe, utilizando el punto decimal
            
            Me.ListBox1.List(i, 5) = _
            Replace(Me.ListBox1.List(i, 5), ".", Application.DecimalSeparator)  'Aquí devuelvo el formato decimal para que no afecte al ListBox
            Me.ListBox1.List(i, 5) = FormatNumber(Me.ListBox1.List(i, 5), 2) 'Aqui doy formato de moneda para que aparezcan los separadores de miles y decimales
        Next
        
Me.txt_Total.Text = sTotal
            
End Sub
Public Sub AgregarItems()
'Agrega los items al listbox

       
       
        With frm_fCompras
            .ListBox1.AddItem frm_ProductoAComprar.txtCantidad.Text
            .ListBox1.List(i, 1) = frm_ProductoAComprar.txt_codigo.Text 'Nombre del producto
            .ListBox1.List(i, 2) = frm_ProductoAComprar.ComboBox1.Text 'Código del producto
            .ListBox1.List(i, 3) = frm_ProductoAComprar.txt_categoria.Text 'Categoria del Producto
            .ListBox1.List(i, 4) = frm_ProductoAComprar.txt_CostoU.Text 'Costo Unitario
            .ListBox1.List(i, 5) = frm_ProductoAComprar.txtCostoTot.Text
                                  
        i = i + 1
          
        End With
        
        sumarImporte
    
        With frm_ProductoAComprar
            .ComboBox1.ListIndex = -1
            .txt_codigo = ""
            .txtCantidad = ""
            .txt_categoria = ""
            .txt_CostoU = ""
            .txt_Existencia = ""
            .txt_medida = ""
        End With

End Sub
Public Sub EliminarItem()
' Elimina el item seleccionado y resta el importe de la columna de importes

Dim sTotal As Currency
On Error GoTo Errores

Me.ListBox1.RemoveItem (ListBox1.ListIndex)
Me.ListBox1.ListIndex = -1 ' Eliminar la "barra de selección"

Me.sumarImporte
Exit Sub

Errores:
MsgBox "Debe seleccionar un producto"

            
End Sub

Public Sub ctrls_FormatoMoneda()
On Error Resume Next
    Me.txt_Total.Text = FormatNumber(Me.txt_Total.Text, 2)
End Sub
Public Sub SumarExistencia(ByVal sCodigo As String, ByVal nCantidad As Long)
Dim Fila As Long
Dim Final As Long
Dim Existencia As Long
Dim TotalExistencia As Long
    
    Final = GetUltimoR(Hoja5)
    

    'Actualiza las existencias
    For Fila = 2 To Final
        If Hoja5.Cells(Fila, 1) = sCodigo Then
            Existencia = Hoja5.Cells(Fila, 3) 'Existencia
            vPrecioVenta = Hoja5.Cells(Fila, 4) 'Precio venta
            TotalExistencia = Existencia + nCantidad 'Suma las existencias
            Hoja5.Cells(Fila, 3) = TotalExistencia
            Exit For
        End If
    Next

End Sub

Public Sub SumarCosto(ByVal sCodigo As String, ByVal nCosto As Currency)
Dim Fila As Long
Dim Final As Long
Dim Existencia As Long
Dim tCosto As Currency
Dim TotalCosto As Currency
Dim cUpromedio As Currency


    Final = GetUltimoR(Hoja5)

    'Actualiza las costos
    For Fila = 2 To Final
        If Hoja5.Cells(Fila, 1) = sCodigo Then
            Existencia = Hoja5.Cells(Fila, 3) 'Existencia
            tCosto = Hoja5.Cells(Fila, 6) 'Costo Total
            TotalCosto = tCosto + nCosto 'Suma los costos
            
            
        If Existencia = 0 Then
                
            Hoja5.Cells(Fila, 3) = Existencia
            Hoja5.Cells(Fila, 5) = vCostoU
            Hoja5.Cells(Fila, 6) = TotalCosto
                    
        Else
            
            cUpromedio = TotalCosto / Existencia
            
            Hoja5.Cells(Fila, 5) = cUpromedio
            Hoja5.Cells(Fila, 6) = TotalCosto
            Exit For
        End If
    End If
Next

End Sub


Public Sub RestarExistenciayCosto(ByVal sCodigo As String, ByVal nCantidad As Long, ByVal nCosto As Currency)
Dim Fila As Long
Dim Final As Long
Dim Existencia As Long
Dim TotalExistencia As Long
Dim tCosto As Currency
Dim TotalCosto As Currency
Dim cUpromedio As Currency
    
    Final = GetUltimoR(Hoja5)
    
    'Actualiza las existencias y los costos
    For Fila = 2 To Final
        If Hoja5.Cells(Fila, 1) = sCodigo Then
            Existencia = Hoja5.Cells(Fila, 3) 'Existencia
            TotalExistencia = Existencia - nCantidad 'Resta las existencias
             Hoja5.Cells(Fila, 3) = TotalExistencia

            
            tCosto = Hoja5.Cells(Fila, 6) 'Costo Total
            TotalCosto = tCosto - nCosto 'Resta los costos
            cUpromedio = TotalCosto / TotalExistencia
            
            Hoja5.Cells(Fila, 5) = cUpromedio
            Hoja5.Cells(Fila, 6) = TotalCosto
            
            Exit For
        End If
    Next

End Sub


