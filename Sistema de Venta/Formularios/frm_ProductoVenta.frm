VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ProductoVenta 
   Caption         =   "PRODUCTO A FACTURAR"
   ClientHeight    =   7305
   ClientLeft      =   14930
   ClientTop       =   1970
   ClientWidth     =   15970
   OleObjectBlob   =   "frm_ProductoVenta.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ProductoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_encargos_Click()
    frm_ListaPedido.Show
End Sub

Private Sub UserForm_Initialize()
    Me.lbx_busqueda.ColumnCount = 7
    Me.lbx_busqueda.ColumnWidths = "45 pt;150 pt;0 pt"
    Me.lbx_busqueda.RowSource = "Código_Venta"
    Me.txt_busqueda.SetFocus
    banderaProductos = 1
End Sub
Private Sub btn_Salir_Click()
    Unload Me
End Sub
Private Sub btnAgregar_Click()
  If ComboBox1 = "" Then
   Call InsertarProductosdeListBox
    Exit Sub
    
  Else
        With frm_Factura
            .AgregarProducto
            .ctrls_FormatoMoneda
        End With
        Me.ComboBox1 = ""
    End If
If Me.CheckBox1.Value = True Then
    btn_Limpiar_Click
End If
    txt_busqueda.SetFocus
End Sub
Private Sub btnAgregar_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = ALT + 43 Then
    frm_ProductoCantidad.Show
    End If
End Sub
Private Sub btn_Limpiar_Click()
    LimpiarControles
    txt_busqueda.Text = ""
    txt_busqueda.SetFocus
End Sub
Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long
Dim encontrado As Boolean
Dim Default As Integer

    Default = 1

    If ComboBox1.Text = "+" Then
        ComboBox1 = ""
    End If
    If ComboBox1.Text = "" Then
        LimpiarControles
    End If

    Fila = 2
        Do While Hoja1.Cells(Fila, 1) <> ""
            Fila = Fila + 1
        Loop
    Final = Fila - 1
        
        For Fila = 2 To Final
            If ComboBox1.Text = Hoja1.Cells(Fila, 1) Then
                Me.txt_Nombre = Hoja1.Cells(Fila, 2)
                Me.txt_categoria = Hoja1.Cells(Fila, 4)
                Me.txtCantidad = Default
                Me.txt_precioV = FormatNumber(Hoja1.Cells(Fila, 5))
                Me.txt_medida.Text = Hoja1.Cells(Fila, 6)
                Exit For
            End If
        Next
End Sub

Private Sub LimpiarControles()
        Me.ComboBox1.Text = ""
        Me.txt_Nombre = ""
        Me.txtCantidad = ""
        Me.txt_precioV = ""
        Me.txtImporte = ""
        Me.txt_categoria = ""
        Me.txt_medida = ""
End Sub

Private Sub txt_busqueda_Change()
Dim Fila As Long
Dim Final As Long

On Error Resume Next


uf = Hoja1.Range("A" & Rows.Count).End(xlUp).Row

If txt_busqueda = "" Then
    Me.lbx_busqueda.RowSource = "Código_Venta"
'    Me.txt_Nombre.Text = ""
    Exit Sub
End If
    
        'Determino el final de la hoja de existencias
'    Fila = 2
'
'        Do While Hoja1.Cells(Fila, 1) <> ""
'            Fila = Fila + 1
'        Loop
'
'        Final = Fila - 1
'
'        'Solicito la información de la hoja de existencias para que se reflejen en los controles
'        For Fila = 2 To Final
'            If txt_busqueda.Text = Hoja1.Cells(Fila, 2) Then
'                    Me.txt_Nombre.Text = txt_busqueda.Text
'                Exit For
'            End If
'        Next
'
'        For Fila = 2 To Final
'            If txt_busqueda.Text = Hoja1.Cells(Fila, 1) Then
'                    Me.Combobox1.Text = txt_busqueda.Text
'                Exit For
'            End If
'        Next


Hoja1.AutoFilterMode = False
Me.lbx_busqueda = Clear
Me.lbx_busqueda.RowSource = Clear

For Fila = 2 To uf
    strg = Hoja1.Cells(Fila, 2).Value 'Variable para descripción
    Codigo = Hoja1.Cells(Fila, 1).Value 'Variable para codigo
    Categoria = Hoja1.Cells(Fila, 7).Value

    If UCase(strg) Like "*" & UCase(txt_busqueda.Value) & "*" Then
        Me.lbx_busqueda.AddItem
        Me.lbx_busqueda.List(X, 0) = Hoja1.Cells(Fila, 1).Value
        Me.lbx_busqueda.List(X, 1) = Hoja1.Cells(Fila, 2).Value
        Me.lbx_busqueda.List(X, 2) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_busqueda.List(X, 3) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_busqueda.List(X, 4) = Hoja1.Cells(Fila, 5).Value
        Me.lbx_busqueda.List(X, 5) = Hoja1.Cells(Fila, 6).Value
        Me.lbx_busqueda.List(X, 6) = Hoja1.Cells(Fila, 7).Value

        X = X + 1
   '----------------------------------------------------------------------------------
    'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
    ElseIf Codigo Like "*" & UCase(txt_busqueda.Value) & "*" Then
        Me.lbx_busqueda.AddItem
        Me.lbx_busqueda.List(X, 0) = Hoja1.Cells(Fila, 1).Value
        Me.lbx_busqueda.List(X, 1) = Hoja1.Cells(Fila, 2).Value
        Me.lbx_busqueda.List(X, 2) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_busqueda.List(X, 3) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_busqueda.List(X, 4) = Hoja1.Cells(Fila, 5).Value
        Me.lbx_busqueda.List(X, 5) = Hoja1.Cells(Fila, 6).Value
        Me.lbx_busqueda.List(X, 6) = Hoja1.Cells(Fila, 7).Value
        X = X + 1

    '----------------------------------------------------------------------------------
        'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
    ElseIf Categoria Like "*" & UCase(txt_busqueda.Value) & "*" Then
        Me.lbx_busqueda.AddItem
        Me.lbx_busqueda.List(X, 0) = Hoja1.Cells(Fila, 1).Value
        Me.lbx_busqueda.List(X, 1) = Hoja1.Cells(Fila, 2).Value
        Me.lbx_busqueda.List(X, 2) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_busqueda.List(X, 3) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_busqueda.List(X, 4) = Hoja1.Cells(Fila, 5).Value
        Me.lbx_busqueda.List(X, 5) = Hoja1.Cells(Fila, 6).Value
        Me.lbx_busqueda.List(X, 6) = Hoja1.Cells(Fila, 7).Value
        X = X + 1

    End If
    '----------------------------------------------------------------------------------
Next

Me.lbx_busqueda.ColumnCount = 7
Me.lbx_busqueda.ColumnWidths = "45 pt;150 pt;0 pt"

End Sub

Private Sub lbx_busqueda_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call InsertarProductosdeListBox
    Me.btnAgregar.SetFocus
End Sub
Private Sub txt_Nombre_Change()
Dim Fila As Long
Dim Final As Long
Dim encontrado As Boolean

    If txt_Nombre.Text = "" Then
        LimpiarControles
    End If

        'Determino el final de la hoja de existencias
    Fila = 2

        Do While Hoja1.Cells(Fila, 1) <> ""
            Fila = Fila + 1
        Loop

        Final = Fila - 1

        'Solicito la información de la hoja de existencias para que se reflejen en los controles
        For Fila = 2 To Final
            If txt_Nombre.Text = Hoja1.Cells(Fila, 2) Then
                Me.ComboBox1 = Hoja1.Cells(Fila, 1)
                Me.txt_categoria = Hoja1.Cells(Fila, 4)
                Me.txt_precioV = FormatNumber(Hoja1.Cells(Fila, 5))
                Me.txt_medida.Text = Hoja1.Cells(Fila, 6)
                Exit For
            End If
        Next
End Sub

Private Sub txtCantidad_Change()
Dim totImporte As Currency
Dim Fila As Long
Dim Final As Long
Dim Registro As Long
Dim antes As Long
Dim ahora As Long
Dim saldo As Long

    If Me.txtCantidad.Text = "+" Then
        Me.txtCantidad = ""
    End If

Me.txtCantidad.BackColor = &H80000005

    totImporte = Val(Me.txtCantidad) * Val(Me.txt_precioV)
    Me.txtImporte.Value = FormatNumber(totImporte, 2)

End Sub
Private Sub txt_precioV_Change()
Dim totImporte As Currency

    totImporte = Val(Me.txtCantidad) * Val(Me.txt_precioV)
    Me.txtImporte.Value = FormatNumber(totImporte, 2)

End Sub


Public Sub ctrls_FormatoMoneda()
On Error Resume Next
    Me.txtCantidad.Text = FormatNumber(Me.txtSubtotal.Text, 2)
    Me.txtSubtotal.Text = FormatNumber(Me.txtSubtotal.Text, 2)
    Me.txtTotal.Text = FormatNumber(Me.txtTotal.Text, 2)
End Sub
