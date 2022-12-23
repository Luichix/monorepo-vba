VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ModificarProducto 
   Caption         =   "Modificar Productos"
   ClientHeight    =   4260
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   8500.001
   OleObjectBlob   =   "frm_ModificarProducto.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ModificarProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long
Dim imgEncontrada As Boolean
Me.ComboBox1.BackColor = &H80000005

On Error GoTo SinFoto



If ComboBox1.Text = "" Then
    LimpiarControles
End If

   
Final = GetUltimoR(Hoja2)

    For Fila = 2 To Final
        If ComboBox1.Text = Hoja2.Cells(Fila, 1) Then
            Me.txt_nombre = Hoja2.Cells(Fila, 2)
            Me.txt_Descrip = Hoja2.Cells(Fila, 3)
            Me.txt_CostoUnitario.Value = FormatNumber(Hoja2.Cells(Fila, 4).Value, 2)
            Me.txt_PrecioVenta.Value = FormatNumber(Hoja2.Cells(Fila, 5).Value, 2)
            Exit For
        
        End If
    Next

    img_ModificarProducto.Picture = LoadPicture(ActiveWorkbook.Path & "\imágenes\" & Me.ComboBox1 & ".jpg")

SinFoto:
If Err = 53 Then
    img_ModificarProducto.Picture = LoadPicture(ActiveWorkbook.Path & "\imágenes\" & Hoja12.Range("C7") & ".jpg")
End If


End Sub
Private Sub ComboBox1_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String





For Fila = 1 To ComboBox1.ListCount
    ComboBox1.RemoveItem 0
Next Fila

Final = GetUltimoR(Hoja2)
    
    For Fila = 2 To Final
        Lista = Hoja2.Cells(Fila, 1)
        ComboBox1.AddItem (Lista)
    Next
End Sub

Private Sub ComboBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
' Validación para que el control solo acepte números
If Hoja12.Range("C2") = True Then
    If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
    End If
End If
End Sub

Private Sub CommandButton1_Click()
    Dim Fila As Long
    Dim Final As Long
    Dim Titulo As String
    Dim vPrecioVenta As Currency
    Dim vCostoUnitario As Currency
    
On Error GoTo Salir

    Titulo = "Gestor de Inventarios"
    
'Validando los controles sin datos
If Me.ComboBox1.Text = "" Then
    Me.ComboBox1.BackColor = &HC0C0FF
    MsgBox "Debe ingresar un Código", , Titulo
    Me.ComboBox1.SetFocus
    Exit Sub
        ElseIf Me.txt_nombre = "" Then
            Me.txt_nombre.BackColor = &HC0C0FF
            MsgBox "Debe ingresar un Nombre de Producto", , Titulo
            Me.txt_nombre.SetFocus
            Exit Sub
                ElseIf Me.txt_Descrip = "" Then
                    Me.txt_Descrip.BackColor = &HC0C0FF
                    MsgBox "Debe ingresar una Descripción", , Titulo
                    Me.txt_Descrip.SetFocus
                    Exit Sub
                        ElseIf Me.txt_CostoUnitario = "" Then
                            Me.txt_CostoUnitario.BackColor = &HC0C0FF
                            MsgBox "Debe ingresar el Costo Unitario", , Titulo
                            Me.txt_CostoUnitario.SetFocus
                            Exit Sub
                                ElseIf Me.txt_PrecioVenta = "" Then
                                    Me.txt_PrecioVenta.BackColor = &HC0C0FF
                                    MsgBox "Debe ingresar el Precio de Venta", , Titulo
                                    Me.txt_PrecioVenta.SetFocus
                                    Exit Sub
End If
    
    
    
    'Inspecciona el listado de la hoja de productos y existencias
    Fila = 2
    Do While Hoja2.Cells(Fila, 1) <> "" And Hoja5.Cells(Fila, 1) <> ""
        Fila = Fila + 1
    Loop
    Final = Fila - 1
    
    vPrecioVenta = Me.txt_PrecioVenta.Value
    vCostoUnitario = Me.txt_CostoUnitario.Value
    
    
    'Modifica datos en la hoja de productos
    For Fila = 2 To Final
        If Me.ComboBox1.Text = Hoja2.Cells(Fila, 1) Then
            Hoja2.Cells(Fila, 2) = Me.txt_nombre
            Hoja2.Cells(Fila, 3) = Me.txt_Descrip
            Hoja2.Cells(Fila, 4) = vCostoUnitario
            Hoja2.Cells(Fila, 5) = vPrecioVenta
            
            Exit For

        End If
    Next
    
    'Modifica datos en la hoja de existencias
    For Fila = 2 To Final
        If Me.ComboBox1.Text = Hoja5.Cells(Fila, 1) Then
            Hoja5.Cells(Fila, 2) = Me.txt_nombre
            Hoja5.Cells(Fila, 4) = vPrecioVenta
            Exit For
        End If
    Next
    
    '-------------------------------------------------


    'Limpia los controles
    LimpiarControles

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, Titulo
 End If
 
End Sub
Private Sub LimpiarControles()
    Me.ComboBox1.Text = ""
    Me.txt_nombre = ""
    Me.txt_Descrip = ""
    Me.txt_CostoUnitario = ""
    Me.txt_PrecioVenta = ""
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub txt_CostoUnitario_Change()
Me.txt_CostoUnitario.BackColor = &H80000005
End Sub

Private Sub txt_CostoUnitario_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
Me.txt_CostoUnitario.Value = FormatNumber(Me.txt_CostoUnitario.Value, 2)
End Sub

Private Sub txt_descrip_Change()
Me.txt_Descrip.BackColor = &H80000005
End Sub

Private Sub txt_Nombre_Change()
Me.txt_nombre.BackColor = &H80000005
End Sub

Private Sub txt_PrecioVenta_Change()
Me.txt_PrecioVenta.BackColor = &H80000005
End Sub

Private Sub txt_PrecioVenta_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
Me.txt_PrecioVenta.Value = FormatNumber(Me.txt_PrecioVenta.Value, 2)
End Sub

Private Sub UserForm_Initialize()
    ComboBox1.MaxLength = Hoja12.Range("C3").Value
    img_ModificarProducto.Picture = LoadPicture(ActiveWorkbook.Path & "\imágenes\" & Hoja12.Range("C7") & ".jpg")

End Sub
