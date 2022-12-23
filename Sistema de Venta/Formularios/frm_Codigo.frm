VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Codigo 
   Caption         =   "GESTOR DE CODIGOS"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   6180
   OleObjectBlob   =   "frm_Codigo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Codigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_Crear1_Click()
Dim Titulo As String
On Error GoTo Salir

Application.ScreenUpdating = False

Titulo = "Gestor de Códigos"

If Me.txt_areap1.Text = "" Then
    Me.txt_areap1.BackColor = &HC0C0FF
    MsgBox "Ingrese el nombre del Área de Producción", , Titulo
    Me.txt_areap1.SetFocus
    Exit Sub
         ElseIf Me.txt_descrip1 = "" Then
            Me.txt_descrip1.BackColor = &HC0C0FF
            MsgBox "Detalle una descripción del Area de producción", , Titulo
            Me.txt_descrip1.SetFocus
            Exit Sub
End If

Buscar_xProduccion
Unload Me
    Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True

Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, Titulo
 End If

End Sub

Private Sub btn_Crear2_Click()
Dim Titulo As String
On Error GoTo Salir

Application.ScreenUpdating = False

Titulo = "Gestor de Códigos"

If Me.txt_areat1.Text = "" Then
    Me.txt_areat1.BackColor = &HC0C0FF
    MsgBox "Ingrese el nombre de la nueva Área de transferencia", , Titulo
    Me.txt_areat1.SetFocus
    Exit Sub
         ElseIf Me.txt_Descrip2 = "" Then
            Me.txt_Descrip2.BackColor = &HC0C0FF
            MsgBox "Detalle una descripción del Área", , Titulo
            Me.txt_Descrip2.SetFocus
            Exit Sub
End If

Buscar_xTransferencia
Unload Me

    Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True

Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, Titulo
 End If
End Sub

Private Sub btn_Crear3_Click()
Dim Titulo As String
On Error GoTo Salir

Application.ScreenUpdating = False

Titulo = "Gestor de Códigos"

If Me.cbx_pro1.Text = "" Then
    Me.cbx_pro1.BackColor = &HC0C0FF
    MsgBox "Seleccione el área del producto", , Titulo
    Me.cbx_pro1.SetFocus
    Exit Sub
         ElseIf Me.txt_npro = "" Then
            Me.txt_npro.BackColor = &HC0C0FF
            MsgBox "Registre el nombre del nuevo producto", , Titulo
            Me.txt_npro.SetFocus
            Exit Sub
                                    ElseIf Me.cbx_Categoria = "" Then
                            Me.cbx_Categoria.BackColor = &HC0C0FF
                            MsgBox "Seleccione la categoria del Producto", , Titulo
                            Me.cbx_Categoria.SetFocus
                            Exit Sub
                    
                ElseIf Me.txt_medida = "" Then
                    Me.txt_medida.BackColor = &HC0C0FF
                    MsgBox "Detalle la Unidad de Medida del Producto", , Titulo
                    Me.txt_medida.SetFocus
                    Exit Sub
                        ElseIf Me.txt_precioV = "" Then
                            Me.txt_precioV.BackColor = &HC0C0FF
                            MsgBox "Detalle el precio de Venta del Producto", , Titulo
                            Me.txt_precioV.SetFocus
                            Exit Sub
                    
End If

Buscar_xProducto
Unload Me
    Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True

Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, Titulo
 End If
End Sub

Private Sub btn_Crear4_Click()
Dim Titulo As String
On Error GoTo Salir

Application.ScreenUpdating = False

Titulo = "Gestor de Códigos"

If Me.txt_nInsumo.Text = "" Then
    Me.txt_nInsumo.BackColor = &HC0C0FF
    MsgBox "Escriba el nombre del nuevo insumo", , Titulo
    Me.txt_nInsumo.SetFocus
    Exit Sub
         ElseIf Me.txt_Medida2 = "" Then
            Me.txt_Medida2.BackColor = &HC0C0FF
            MsgBox "Detalle la Unidad de Medida del Insumo", , Titulo
            Me.txt_Medida2.SetFocus
            Exit Sub
                ElseIf Me.txt_descrip3 = "" Then
                    Me.txt_descrip3.BackColor = &HC0C0FF
                    MsgBox "Detalle una descripción del insumo", , Titulo
                    Me.txt_descrip3.SetFocus
                    Exit Sub
                                          
End If

Buscar_xInsumo
Unload Me
    Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True

Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, Titulo
 End If
End Sub

Private Sub btn_Insumo_Click()
Me.Frame1.Visible = False
Me.Frame2.Visible = False
Me.Frame3.Visible = False
Me.Frame4.Visible = True
Limpiar
Me.txt_nInsumo.SetFocus
End Sub

Private Sub btn_Produccion_Click()
Me.Frame1.Visible = True
Me.Frame2.Visible = False
Me.Frame3.Visible = False
Me.Frame4.Visible = False
Limpiar
Me.txt_areap1.SetFocus
End Sub

Private Sub btn_Producto_Click()
Me.Frame1.Visible = False
Me.Frame2.Visible = False
Me.Frame3.Visible = True
Me.Frame4.Visible = False
Limpiar
Me.cbx_pro1.SetFocus
End Sub

Private Sub btn_Salir1_Click()
Unload Me
End Sub

Private Sub btn_Salir2_Click()
Unload Me
End Sub

Private Sub btn_Salir3_Click()
Unload Me
End Sub

Private Sub btn_Salir4_Click()
Unload Me
End Sub

Private Sub btn_SalirT_Click()
Unload Me
End Sub

Private Sub btn_Transferencia_Click()
Me.Frame1.Visible = False
Me.Frame2.Visible = True
Me.Frame3.Visible = False
Me.Frame4.Visible = False
Limpiar
Me.txt_areat1.SetFocus
End Sub

Private Sub cbx_Categoria_Enter()
Dim Fila As Long
Dim xFila As Long
Dim Final As Long
Dim Lista As String

For Fila = 1 To cbx_Categoria.ListCount
    cbx_Categoria.RemoveItem 0
Next Fila

    xFila = 2
    
    Do While Hoja1.Cells(xFila, 27) <> ""
        xFila = xFila + 1
    Loop
    
    Final = xFila - 1

    For Fila = 2 To Final
        Lista = Hoja1.Cells(Fila, 27)
        cbx_Categoria.AddItem (Lista)
    Next
End Sub

Private Sub cbx_pro1_Change()
Me.cbx_pro1.BackColor = &HFFFFFF
End Sub


Private Sub txt_areap1_Change()
Me.txt_areap1.BackColor = &HFFFFFF
End Sub

Private Sub txt_areat1_Change()
Me.txt_areat1.BackColor = &HFFFFFF
End Sub

Private Sub txt_descrip1_Change()
Me.txt_descrip1.BackColor = &HFFFFFF
End Sub

Private Sub txt_descrip2_Change()
Me.txt_Descrip2.BackColor = &HFFFFFF
End Sub

Private Sub txt_descrip3_Change()
Me.txt_descrip3.BackColor = &HFFFFFF
End Sub

Private Sub txt_medida_Change()
Me.txt_medida.BackColor = &HFFFFFF
End Sub

Private Sub txt_Medida2_Change()
Me.txt_Medida2.BackColor = &HFFFFFF
End Sub

Private Sub txt_nInsumo_Change()
Me.txt_nInsumo.BackColor = &HFFFFFF
End Sub

Private Sub txt_npro_Change()
Me.txt_npro.BackColor = &HFFFFFF
End Sub

Private Sub txt_precioV_Change()
Me.txt_precioV.BackColor = &HFFFFFF
End Sub
Private Sub txt_precioV_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

KeyAscii = ValidarDecimales(txt_precioV, KeyAscii)

End Sub

Private Sub UserForm_Initialize()
EliminarTitulo Me.Caption
    Me.Height = Me.Height - 20
End Sub

Private Sub cbx_pro1_Enter()
Dim Fila As Long
Dim xFila As Long
Dim Final As Long
Dim Lista As String

For Fila = 1 To cbx_pro1.ListCount
    cbx_pro1.RemoveItem 0
Next Fila

    xFila = 2
    
    Do While Hoja1.Cells(xFila, 26) <> ""
        xFila = xFila + 1
    Loop
    
    Final = xFila - 1

    For Fila = 2 To Final
        Lista = Hoja1.Cells(Fila, 26)
        cbx_pro1.AddItem (Lista)
    Next
End Sub

Private Sub Limpiar()
Me.txt_areap1 = Empty
Me.txt_areat1 = Empty
Me.txt_descrip1 = Empty
Me.txt_Descrip2 = Empty
Me.txt_descrip3 = Empty
Me.txt_medida = Empty
Me.txt_Medida2 = Empty
Me.txt_nInsumo = Empty
Me.txt_npro = Empty
Me.txt_precioV = Empty
Me.cbx_pro1 = Empty
Me.cbx_Categoria = Empty
End Sub

Private Sub Buscar_xProduccion()
Dim Titulo As String

Titulo = "GESTIÓN DE CODIGOS"
X = UCase(Me.txt_areap1.Text)
If Hoja1.Visible = xlSheetVisible Then

                Hoja1.Select
                Range("Z1").Select
                    Do Until IsEmpty(ActiveCell)
                          ActiveCell.Offset(1, 0).Select
                          If ActiveCell.Value Like X Then
                              encontrado = True
                              Exit Do
                                                                  
                          End If
                    Loop
                    
                  If encontrado = True Then
                        MsgBox "Area de Producción ya existente", vbInformation, Titulo
                        Exit Sub
                          
                        Else: encontrado = False
                            Agregar_xProduccion
                  End If

Else: Hoja1.Visible = xlSheetVeryHidden
    
                Hoja1.Visible = xlSheetVisible
                
                Hoja1.Select
                Range("Z1").Select
                    Do Until IsEmpty(ActiveCell)
                          ActiveCell.Offset(1, 0).Select
                          If ActiveCell.Value Like X Then
                              encontrado = True
                              Exit Do
                                                                  
                          End If
                    Loop
                    
                  If encontrado = True Then
                        MsgBox "Area de Producción ya existente", vbInformation, Titulo
                        Exit Sub
                          
                        Else: encontrado = False
                            Agregar_xProduccion
                  End If

                Hoja1.Visible = xlSheetVeryHidden
                
End If
''''''''''''''''''''''''''''''''
End Sub
Private Sub Agregar_xProduccion()

'Envía los datos a la hoja de salidas
Hoja1.Unprotect ""

Hoja1.Select
    Hoja1.Range("Z2:AA2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja1.Range("Z3:AA3").Select
    Selection.Copy
    Hoja1.Range("Z2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja1.Cells(2, 26) = UCase(Me.txt_areap1.Text)
        Hoja1.Cells(2, 27) = UCase(Me.txt_descrip1.Text)

Hoja1.Select
    Hoja1.Range("AC2:AD2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja1.Range("AC3:AD3").Select
    Selection.Copy
    Hoja1.Range("AC2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja1.Cells(2, 29) = UCase(Me.txt_areap1.Text)
        Hoja1.Cells(2, 30) = UCase(Me.txt_descrip1.Text)

Hoja1.Protect ""

    MsgBox "Area de Producción agregada con exito...!"
    Limpiar
End Sub
Private Sub Buscar_xTransferencia()
Dim Titulo As String

Titulo = "GESTIÓN DE CODIGOS"
X = UCase(Me.txt_areat1.Text)
If Hoja1.Visible = xlSheetVisible Then

                Hoja1.Select
                Range("AC1").Select
                    Do Until IsEmpty(ActiveCell)
                          ActiveCell.Offset(1, 0).Select
                          If ActiveCell.Value Like X Then
                              encontrado = True
                              Exit Do
                                                                  
                          End If
                    Loop
                    
                  If encontrado = True Then
                        MsgBox "Area de Transferencia ya existente", vbInformation, Titulo
                        Exit Sub
                          
                        Else: encontrado = False
                            Agregar_xTransferencia
                  End If

Else: Hoja1.Visible = xlSheetVeryHidden
    
                Hoja1.Visible = xlSheetVisible
                
               Hoja1.Select
                Range("AC1").Select
                    Do Until IsEmpty(ActiveCell)
                          ActiveCell.Offset(1, 0).Select
                          If ActiveCell.Value Like X Then
                              encontrado = True
                              Exit Do
                                                                  
                          End If
                    Loop
                    
                  If encontrado = True Then
                        MsgBox "Area de Transferencia ya existente", vbInformation, Titulo
                        Exit Sub
                          
                        Else: encontrado = False
                            Agregar_xTransferencia
                  End If

                Hoja1.Visible = xlSheetVeryHidden
                
End If
''''''''''''''''''''''''''''''''
End Sub
Private Sub Agregar_xTransferencia()

'Envía los datos a la hoja de salidas
Hoja1.Unprotect ""

Hoja1.Select
    Hoja1.Range("AC2:AD2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja1.Range("AC3:AD3").Select
    Selection.Copy
    Hoja1.Range("AC2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja1.Cells(2, 29) = UCase(Me.txt_areat1.Text)
        Hoja1.Cells(2, 30) = UCase(Me.txt_Descrip2.Text)

Hoja1.Protect ""

    MsgBox "Area de Transferencia agregada con exito...!"
    Limpiar
End Sub

Private Sub Buscar_xProducto()
Dim Titulo As String

Titulo = "GESTIÓN DE CODIGOS"
X = UCase(Me.txt_npro.Text)
If Hoja1.Visible = xlSheetVisible Then

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
                        MsgBox "Producto ya existente", vbInformation, Titulo
                        Exit Sub
                          
                        Else: encontrado = False
                            Agregar_xProducto
                  End If

Else: Hoja1.Visible = xlSheetVeryHidden
                Hoja6.Visible = xlSheetVisible
                Hoja1.Visible = xlSheetVisible
                
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
                        MsgBox "Producto ya existente", vbInformation, Titulo
                        Exit Sub
                          
                        Else: encontrado = False
                            Agregar_xProducto
                  End If

                Hoja6.Visible = xlSheetVeryHidden
                Hoja1.Visible = xlSheetVeryHidden
                
End If
''''''''''''''''''''''''''''''''
End Sub
Private Sub Agregar_xProducto()
Dim Numero As Integer

Hoja93.Range("M2").Value = Hoja93.Range("M2").Value + 1
Numero = Hoja93.Range("M2").Value

'Envía los datos a la hoja de salidas
Hoja1.Unprotect ""
Hoja6.Unprotect ""

Hoja1.Select
    Hoja1.Range("A2:G2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja1.Range("A3:G3").Select
    Selection.Copy
    Hoja1.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja1.Cells(2, 1) = Numero
        Hoja1.Cells(2, 2) = UCase(Me.txt_npro.Text)
        Hoja1.Cells(2, 4) = UCase(Me.cbx_pro1.Text)
        Hoja1.Cells(2, 5) = txt_precioV.Value
        Hoja1.Cells(2, 6) = UCase(Me.txt_medida.Text)
        Hoja1.Cells(2, 7) = UCase(Me.cbx_Categoria.Text)
        
    Hoja1.ListObjects("Código_Venta").Sort.SortFields.Clear
    Hoja1.ListObjects("Código_Venta").Sort.SortFields.Add _
        Key:=Range("Código_Venta[[#All],[N DEPT]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With Hoja1.ListObjects("Código_Venta").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
         
    Hoja6.Select
    Hoja6.Range("A2:K2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja6.Range("A3:K3").Select
    Selection.Copy
    Hoja6.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja6.Cells(2, 1) = UCase(Me.txt_npro.Text)
        
            Hoja6.ListObjects("PRODUCTOS").Sort.SortFields.Clear
    Hoja6.ListObjects("PRODUCTOS").Sort.SortFields.Add _
        Key:=Range("PRODUCTOS[[#All],[ID ITEM]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With Hoja6.ListObjects("PRODUCTOS").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
Hoja1.Protect ""
Hoja6.Protect ""

    MsgBox "Producto agregado con exito...!"
    
    Limpiar
    
End Sub


Private Sub Buscar_xInsumo()
Dim Titulo As String

Titulo = "GESTIÓN DE CODIGOS"
X = UCase(Me.txt_nInsumo.Text)
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
                    
                  If encontrado = True Then
                        MsgBox "Insumo ya existente", vbInformation, Titulo
                        Exit Sub
                          
                        Else: encontrado = False
                            Agregar_xInsumo
                  End If

Else: Hoja1.Visible = xlSheetVeryHidden
                Hoja5.Visible = xlSheetVisible
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
                    
                  If encontrado = True Then
                        MsgBox "Insumo ya existente", vbInformation, Titulo
                        Exit Sub
                          
                        Else: encontrado = False
                            Agregar_xInsumo
                  End If

                Hoja5.Visible = xlSheetVeryHidden
                Hoja1.Visible = xlSheetVeryHidden
                
End If
''''''''''''''''''''''''''''''''
End Sub
Private Sub Agregar_xInsumo()
Dim Numero As Integer
Dim Categoria As String

Hoja93.Range("N2").Value = Hoja93.Range("N2").Value + 1
Numero = Hoja93.Range("N2").Value
Categoria = "INSUMO"

'Envía los datos a la hoja de salidas
Hoja1.Unprotect ""
Hoja5.Unprotect ""

Hoja1.Select
    Hoja1.Range("J2:O2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja1.Range("J3:O3").Select
    Selection.Copy
    Hoja1.Range("J2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja1.Cells(2, 10) = Numero
        Hoja1.Cells(2, 11) = UCase(Me.txt_nInsumo.Text)
        Hoja1.Cells(2, 13) = UCase(Me.txt_Medida2.Text)
        Hoja1.Cells(2, 14) = Categoria
        Hoja1.Cells(2, 15) = UCase(txt_descrip3.Text)
        
    Hoja1.ListObjects("tbl_insumos").Sort.SortFields.Clear
    Hoja1.ListObjects("tbl_insumos").Sort.SortFields.Add _
        Key:=Range("tbl_insumos[[#All],[N DEPT]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With Hoja1.ListObjects("tbl_insumos").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
         
    Hoja5.Select
    Hoja5.Range("A2:K2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja5.Range("A3:K3").Select
    Selection.Copy
    Hoja5.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja5.Cells(2, 1) = UCase(Me.txt_nInsumo.Text)
        
            Hoja5.ListObjects("MATERIALES").Sort.SortFields.Clear
    Hoja5.ListObjects("MATERIALES").Sort.SortFields.Add _
        Key:=Range("MATERIALES[[#All],[ID ITEM]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With Hoja5.ListObjects("MATERIALES").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
Hoja1.Protect ""
Hoja5.Protect ""

    MsgBox "Insumo agregado con exito...!"
    
    Limpiar
    
End Sub




