VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ProductoAFacturar 
   Caption         =   "PRODUCTO A FACTURAR"
   ClientHeight    =   6330
   ClientLeft      =   14930
   ClientTop       =   1970
   ClientWidth     =   5970
   OleObjectBlob   =   "frm_ProductoAFacturar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ProductoAFacturar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_Productos_Click()

banderaListadoProductos = 1
    Call LanzarListadoProductos(Me, "lbl_LanzarListadoCuentas")

End Sub


Private Sub btnAgregar_Click()
'On Error GoTo Salir
    With frm_Factura
        .AgregarItems
        .ctrls_FormatoMoneda
    End With
    
'Salir:
' If Err <> 0 Then
'    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
' End If
End Sub


Private Sub btnCerrar_Click()
Unload Me
End Sub
Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long
Dim encontrado As Boolean
'Rutina que permite reflejar el resto de la información en los demás controles
'después de haber realizado una selección en el ComboBox
    Me.ComboBox1.BackColor = &HFFFFFF

    If ComboBox1.Text = "+" Then
        ComboBox1 = ""
    End If
    

    If ComboBox1.Text = "" Then
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
            If ComboBox1.Text = Hoja1.Cells(Fila, 1) Then
                Me.txt_Nombre = Hoja1.Cells(Fila, 2)
                Me.txt_categoria = Hoja1.Cells(Fila, 4)
                Me.txt_precioV = FormatNumber(Hoja1.Cells(Fila, 5))
                Me.TextBox2.Text = Hoja1.Cells(Fila, 6)
                Me.txtCantidad = 1
                Exit For
            End If
        Next
        


End Sub




Private Sub txt_Nombre_Change()
Dim Fila As Long
Dim Final As Long
Dim encontrado As Boolean
'Rutina que permite reflejar el resto de la información en los demás controles
'después de haber realizado una selección en el ComboBox

    If txt_Nombre.Text = "+" Then
        txt_Nombre = ""
    End If
    

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
                Me.TextBox2.Text = Hoja1.Cells(Fila, 6)
                Me.txtCantidad = 1
                Exit For
            End If
        Next
  

End Sub
Private Sub ComboBox1_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String

'Toda esta rutina agrega los items al ComboBox

For Fila = 1 To ComboBox1.ListCount
    ComboBox1.RemoveItem 0
Next Fila

    'Inspecciono la hoja de Existencias para determinar el final del listado
Final = GetUltimoR(Hoja1)
    
    'Agrego el listado de códigos de productos al ComboBox desde la hoja de existencias
    For Fila = 2 To Final
        If Hoja1.Cells(Fila, 1) > 0 Then
            Lista = Hoja1.Cells(Fila, 1)
            ComboBox1.AddItem (Lista)
        End If
    Next
End Sub

Private Sub txt_Nombre_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String

'Toda esta rutina agrega los items al ComboBox

For Fila = 1 To Me.txt_Nombre.ListCount
    txt_Nombre.RemoveItem 0
Next Fila

    'Inspecciono la hoja de Existencias para determinar el final del listado
Final = GetUltimoR(Hoja1)
    
    'Agrego el listado de códigos de productos al ComboBox desde la hoja de existencias
    For Fila = 2 To Final
        If Hoja1.Cells(Fila, 2) > 0 Then
            Lista = Hoja1.Cells(Fila, 2)
            txt_Nombre.AddItem (Lista)
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
    
  
    'Determino el final del listdo de existencias
Fila = 1
    Do While Hoja1.Cells(Fila, 1) <> Empty
        Fila = Fila + 1
    Loop
Final = Fila

    'Compruebo que el código ingresado en el ComboBox, coincida en hoja de existencias
    ' para realizar la respectiva operación aritmética
    

End Sub
Private Sub txt_precioV_Change()
Dim totImporte As Currency

    totImporte = Val(Me.txtCantidad) * Val(Me.txt_precioV)
    Me.txtImporte.Value = FormatNumber(totImporte, 2)

End Sub
Private Sub LimpiarControles()
        'Limpia los controles
        Me.ComboBox1.Text = ""
        Me.txt_Nombre = ""
        Me.txtCantidad = 1
        Me.txt_precioV = ""
        Me.txtImporte = ""

        Me.txt_categoria = ""
        Me.TextBox2 = ""

End Sub

Public Sub ctrls_FormatoMoneda()
On Error Resume Next
    Me.txtSubtotal.Text = FormatNumber(Me.txtSubtotal.Text, 2)
    Me.txtTotal.Text = FormatNumber(Me.txtTotal.Text, 2)
End Sub



Private Sub ComboBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If KeyAscii = ALT + 43 Then

btn_Productos_Click

End If

End Sub



Private Sub txtCantidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If KeyAscii = ALT + 43 Then

btn_Productos_Click

End If
End Sub

Private Sub UserForm_Click()

End Sub
