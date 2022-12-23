VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ProductoAComprar 
   Caption         =   "GESTOR DE INVENTARIOS"
   ClientHeight    =   6210
   ClientLeft      =   14930
   ClientTop       =   2460
   ClientWidth     =   12880
   OleObjectBlob   =   "frm_ProductoAComprar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ProductoAComprar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    Me.lbx_insumo.ColumnCount = 5
    Me.lbx_insumo.ColumnWidths = "45 pt;130 pt;0 pt"
    Me.lbx_insumo.RowSource = "tbl_insumos"
    banderaInsumos = 1

    Me.lbx_producto.ColumnCount = 7
    Me.lbx_producto.ColumnWidths = "45 pt;130 pt;0 pt;0;0 pt;"
    Me.lbx_producto.RowSource = "Código_Venta"
    banderaProductosCompra = 1
End Sub

Private Sub btn_Salir_Click()
    Unload Me
End Sub

Private Sub txt_producto_Change()
Dim Fila As Long
Dim Final As Long

On Error Resume Next


uf = Hoja1.Range("A" & Rows.Count).End(xlUp).Row

If txt_producto = "" Then
    Me.lbx_producto.RowSource = "Código_Venta"
    Exit Sub
End If

Hoja1.AutoFilterMode = False
Me.lbx_producto = Clear
Me.lbx_producto.RowSource = Clear

For Fila = 2 To uf
    strg = Hoja1.Cells(Fila, 2).Value 'Variable para descripción
    Codigo = Hoja1.Cells(Fila, 1).Value 'Variable para codigo
    Categoria = Hoja1.Cells(Fila, 7).Value

    If UCase(strg) Like "*" & UCase(txt_producto.Value) & "*" Then
        Me.lbx_producto.AddItem
        Me.lbx_producto.List(X, 0) = Hoja1.Cells(Fila, 1).Value
        Me.lbx_producto.List(X, 1) = Hoja1.Cells(Fila, 2).Value
        Me.lbx_producto.List(X, 2) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_producto.List(X, 3) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_producto.List(X, 4) = Hoja1.Cells(Fila, 5).Value
        Me.lbx_producto.List(X, 5) = Hoja1.Cells(Fila, 6).Value
        Me.lbx_producto.List(X, 6) = Hoja1.Cells(Fila, 7).Value

        X = X + 1
   '----------------------------------------------------------------------------------
    'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
    ElseIf Codigo Like "*" & UCase(txt_producto.Value) & "*" Then
        Me.lbx_producto.AddItem
        Me.lbx_producto.List(X, 0) = Hoja1.Cells(Fila, 1).Value
        Me.lbx_producto.List(X, 1) = Hoja1.Cells(Fila, 2).Value
        Me.lbx_producto.List(X, 2) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_producto.List(X, 3) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_producto.List(X, 4) = Hoja1.Cells(Fila, 5).Value
        Me.lbx_producto.List(X, 5) = Hoja1.Cells(Fila, 6).Value
        Me.lbx_producto.List(X, 6) = Hoja1.Cells(Fila, 7).Value
        X = X + 1

    '----------------------------------------------------------------------------------
        'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
    ElseIf Categoria Like "*" & UCase(txt_producto.Value) & "*" Then
        Me.lbx_producto.AddItem
        Me.lbx_producto.List(X, 0) = Hoja1.Cells(Fila, 1).Value
        Me.lbx_producto.List(X, 1) = Hoja1.Cells(Fila, 2).Value
        Me.lbx_producto.List(X, 2) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_producto.List(X, 3) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_producto.List(X, 4) = Hoja1.Cells(Fila, 5).Value
        Me.lbx_producto.List(X, 5) = Hoja1.Cells(Fila, 6).Value
        Me.lbx_producto.List(X, 6) = Hoja1.Cells(Fila, 7).Value
        X = X + 1

    End If
    '----------------------------------------------------------------------------------
Next

    Me.lbx_producto.ColumnCount = 7
    Me.lbx_producto.ColumnWidths = "45 pt;130 pt;0 pt;0;0 pt;"

End Sub


Private Sub txt_insumo_Change()
Dim Fila As Long
Dim Final As Long

On Error Resume Next


uf = Hoja1.Range("J" & Rows.Count).End(xlUp).Row

If txt_insumo = "" Then
    Me.lbx_insumo.RowSource = "tbl_insumos"
    Exit Sub
End If

Hoja1.AutoFilterMode = False
Me.lbx_insumo = Clear
Me.lbx_insumo.RowSource = Clear

For Fila = 2 To uf
    strg = Hoja1.Cells(Fila, 11).Value 'Variable para descripción
    Codigo = Hoja1.Cells(Fila, 10).Value 'Variable para codigo
    Categoria = Hoja1.Cells(Fila, 14).Value

    If UCase(strg) Like "*" & UCase(txt_insumo.Value) & "*" Then
        Me.lbx_insumo.AddItem
        Me.lbx_insumo.List(X, 0) = Hoja1.Cells(Fila, 10).Value
        Me.lbx_insumo.List(X, 1) = Hoja1.Cells(Fila, 11).Value
        Me.lbx_insumo.List(X, 2) = Hoja1.Cells(Fila, 12).Value
        Me.lbx_insumo.List(X, 3) = Hoja1.Cells(Fila, 13).Value
        Me.lbx_insumo.List(X, 4) = Hoja1.Cells(Fila, 14).Value
        Me.lbx_insumo.List(X, 5) = Hoja1.Cells(Fila, 15).Value

        X = X + 1
   '----------------------------------------------------------------------------------
    'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
    ElseIf Codigo Like "*" & UCase(txt_insumo.Value) & "*" Then
        Me.lbx_insumo.AddItem
        Me.lbx_insumo.List(X, 0) = Hoja1.Cells(Fila, 10).Value
        Me.lbx_insumo.List(X, 1) = Hoja1.Cells(Fila, 11).Value
        Me.lbx_insumo.List(X, 2) = Hoja1.Cells(Fila, 12).Value
        Me.lbx_insumo.List(X, 3) = Hoja1.Cells(Fila, 13).Value
        Me.lbx_insumo.List(X, 4) = Hoja1.Cells(Fila, 14).Value
        Me.lbx_insumo.List(X, 5) = Hoja1.Cells(Fila, 15).Value
        X = X + 1

    '----------------------------------------------------------------------------------
        'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
    ElseIf Categoria Like "*" & UCase(txt_insumo.Value) & "*" Then
        Me.lbx_insumo.AddItem
        Me.lbx_insumo.List(X, 0) = Hoja1.Cells(Fila, 10).Value
        Me.lbx_insumo.List(X, 1) = Hoja1.Cells(Fila, 11).Value
        Me.lbx_insumo.List(X, 2) = Hoja1.Cells(Fila, 12).Value
        Me.lbx_insumo.List(X, 3) = Hoja1.Cells(Fila, 13).Value
        Me.lbx_insumo.List(X, 4) = Hoja1.Cells(Fila, 14).Value
        Me.lbx_insumo.List(X, 5) = Hoja1.Cells(Fila, 15).Value
        X = X + 1

    End If
    '----------------------------------------------------------------------------------
Next

    Me.lbx_insumo.ColumnCount = 5
    Me.lbx_insumo.ColumnWidths = "45 pt;130 pt;0 pt"

End Sub





Private Sub txtCantidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

        If KeyAscii = ALT + 43 Then
                    If txtCantidad.Text = "+" Then
                        txtCantidad = ""
                    End If
                    frm_ProductocOMPRAS.Show
        End If
    
End Sub

Private Sub lbx_Insumo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call InsertarInsumos
    Me.txtCantidad.SetFocus
End Sub
Private Sub lbx_producto_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call InsertarProductos
    Me.txtCantidad.SetFocus
End Sub



Private Sub btn_Limpiar_Click()
    ComboBox1 = ""
    txt_insumo.Text = ""
    txt_insumo.SetFocus
End Sub
Private Sub btn_Limpiar2_Click()
    ComboBox1 = ""
    txt_producto.Text = ""
    txt_producto.SetFocus
End Sub


Private Sub btnAgregar_Click()
Dim Titulo As String
On Error GoTo Salir

Application.ScreenUpdating = False

Titulo = "Gestor de Inventario"

If Me.ComboBox1.Text = "" Then
    Me.ComboBox1.BackColor = &HC0C0FF
    MsgBox "Ingrese el nombre del producto", vbInformation, Titulo
    Me.ComboBox1.SetFocus
    Exit Sub
End If
    If Me.txtCantidad.Text = "" Then
          Me.txtCantidad.BackColor = &HC0C0FF
          MsgBox "Ingrese la cantidad", vbInformation, Titulo
          Me.txtCantidad.SetFocus
          Exit Sub
      End If
                If Me.txt_CostoU.Text = "" Then
                    Me.txt_CostoU.BackColor = &HC0C0FF
                    MsgBox "Ingrese el costo unitario", vbInformation, Titulo
                    Me.txt_CostoU.SetFocus
                    Exit Sub
                End If


    buscar_producto
    

  If Hoja3.Visible = xlSheetVisible Then
    Hoja3.Select
  End If

   Application.ScreenUpdating = True
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor Administrativo"


 End If


 ComboBox1.SetFocus

End Sub
Private Sub buscar_producto()


X = Me.ComboBox1.Text
If Hoja1.Visible = xlSheetVisible Then

                Hoja1.Select
                Range("K1").Select

                    Do Until IsEmpty(ActiveCell)
                          ActiveCell.Offset(1, 0).Select
                          If ActiveCell.Value Like X Then
                              encontrado = True
                              Exit Do


                          End If

                    Loop

                Hoja1.Select
                Range("B1").Select
                    Do Until IsEmpty(ActiveCell)
                          ActiveCell.Offset(1, 0).Select
                          If ActiveCell.Value Like X Then
                              encontrado = True
                              Exit Do


                          End If

                    Loop

                  If encontrado = True Then
                         With frm_fCompras
                                 .AgregarItems
                            End With

                        Else: encontrado = False
                            MsgBox "Producto no Existente", vbInformation, Titulo
                  End If

Else: Hoja1.Visible = xlSheetVeryHidden

                Hoja1.Visible = xlSheetVisible

                Hoja1.Select
                Range("K1").Select
                  Do Until IsEmpty(ActiveCell)
                        ActiveCell.Offset(1, 0).Select
                        If ActiveCell.Value Like X Then
                            encontrado = True
                            Exit Do

                        End If

                Loop

                  Hoja1.Select
                Range("B1").Select
                    Do Until IsEmpty(ActiveCell)
                          ActiveCell.Offset(1, 0).Select
                          If ActiveCell.Value Like X Then
                              encontrado = True
                              Exit Do


                          End If

                    Loop

                If encontrado = True Then
                 With frm_fCompras
                         .AgregarItems
                    End With

                Else: encontrado = False
                    MsgBox "Producto no Existente", vbInformation, Titulo
                End If

                Hoja1.Visible = xlSheetVeryHidden

End If
End Sub
Private Sub ComboBox1_Change()
Dim Fila As Long
Dim xFila As Long
Dim Final As Long
Dim xFinal As Long
Dim Clase As String
Me.ComboBox1.BackColor = &HFFFFFF

If ComboBox1.Text = "" Then
    LimpiarControles
End If

 Fila = 2
    Do While Hoja1.Cells(Fila, 11) <> ""
        Fila = Fila + 1
    Loop

    Final = Fila - 1

    'Solicito la información de la hoja de materiales para que se reflejen en los controles
    For Fila = 2 To Final
        If ComboBox1.Text = Hoja5.Cells(Fila, 1) Then
            Me.txt_codigo = Hoja5.Cells(Fila, 2)
            Me.txt_medida = Hoja5.Cells(Fila, 3)
            Me.txt_Existencia = Hoja5.Cells(Fila, 10)
            Me.txt_categoria = Hoja5.Cells(Fila, 4)
            Exit For
        End If
    Next
    
xFila = 2
    Do While Hoja1.Cells(xFila, 1) <> ""
        xFila = xFila + 1
    Loop

    xFinal = xFila - 1
    
    'Solicito la información de la hoja de productos para que se reflejen en los controles
    For xFila = 2 To xFinal
        If ComboBox1.Text = Hoja6.Cells(xFila, 1) Then
            Me.txt_codigo = Hoja6.Cells(xFila, 2)
            Me.txt_medida = Hoja6.Cells(xFila, 3)
            Me.txt_Existencia = Hoja6.Cells(xFila, 10)
            Me.txt_categoria = Hoja6.Cells(xFila, 4)

            Exit For
        End If
    Next
End Sub
Private Sub ComboBox1_Enter()
Dim Fila As Long
Dim Final As Long
Dim xFinal As Long
Dim Lista As String

'Toda esta rutina agrega los items al ComboBox

For Fila = 1 To ComboBox1.ListCount
    ComboBox1.RemoveItem 0
Next Fila


    'Inspecciono la hoja de productos para determinar el final del listado
Final = GetUltimoR(Hoja5)
xFinal = GetUltimoR(Hoja6)


    'Agrego el listado de códigos de productos al ComboBox desde la hoja de productos
    For Fila = 2 To Final
        Lista = Hoja5.Cells(Fila, 1)
        ComboBox1.AddItem (Lista)
    Next
    For Fila = 2 To xFinal
        Lista = Hoja6.Cells(Fila, 1)
        ComboBox1.AddItem (Lista)
    Next
End Sub

Private Sub txtCantidad_Change()
Dim totImporte As Currency

txtCantidad.BackColor = &H80000005

    totImporte = Val(Me.txtCantidad) * Val(Me.txt_CostoU)
    Me.txtCostoTot.Value = FormatNumber(totImporte, 2)

End Sub
Private Sub txt_CostoU_Change()
Dim totImporte As Currency

Me.txt_CostoU.BackColor = &HFFFFFF

    totImporte = Val(Me.txtCantidad) * Val(Me.txt_CostoU)
    Me.txtCostoTot.Value = FormatNumber(totImporte, 2)

End Sub

Private Sub LimpiarControles()
        'Limpia los controles
        Me.ComboBox1.Text = ""
        Me.txt_medida = ""
        Me.txt_codigo = ""
        Me.txtCantidad = ""
        Me.txt_CostoU = ""
        Me.txt_Existencia = ""
        Me.txtCostoTot = ""
        Me.txt_categoria = ""
        Me.txtCantidad.BackColor = &HFFFFFF
        Me.txt_CostoU.BackColor = &HFFFFFF
        Me.ComboBox1.BackColor = &HFFFFFF
      
End Sub

