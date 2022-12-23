VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ProductoAFacturar 
   Caption         =   "Producto a Facturar"
   ClientHeight    =   4890
   ClientLeft      =   14930
   ClientTop       =   1970
   ClientWidth     =   4990
   OleObjectBlob   =   "frm_ProductoAFacturar.frx":0000
End
Attribute VB_Name = "frm_ProductoAFacturar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnAgregar_Click()
On Error GoTo Salir
    With frm_Factura
        .AgregarItems
        .ctrls_FormatoMoneda
    End With
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If
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



    If ComboBox1.Text = "" Then
        LimpiarControles
    End If
    
        'Determino el final de la hoja de existencias
    Fila = 2
        
        Do While Hoja1.Cells(Fila, 21) <> ""
            Fila = Fila + 1
        Loop
        
        Final = Fila - 1
        
        
        
        'Solicito la información de la hoja de existencias para que se reflejen en los controles
        For Fila = 2 To Final
            If ComboBox1.Text = Hoja1.Cells(Fila, 21) Then
                encontrado = True
                Me.txt_nombre = Hoja1.Cells(Fila, 22)
    
                Exit For
            End If
        Next
        
        
        If encontrado = False Then
                Me.txt_nombre = Empty
                Me.txt_PrecioV.Value = Empty
               
        End If

    
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
        If Hoja1.Cells(Fila, 21) > 0 Then
            Lista = Hoja1.Cells(Fila, 21)
            ComboBox1.AddItem (Lista)
        End If
    Next
End Sub

'Private Sub ComboBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'' Validación para que el control solo acepte números
'If Hoja12.Range("C2") = True Then
'    If KeyAscii < 48 Or KeyAscii > 57 Then
'    KeyAscii = 0
'    End If
'End If
'End Sub

Private Sub txtCantidad_Change()
Dim totImporte As Currency


    totImporte = Val(Me.txtCantidad) * Val(Me.txt_PrecioV)
    Me.txtImporte.Value = FormatNumber(totImporte, 2)

End Sub
Private Sub txt_PrecioV_Change()
Dim totImporte As Currency


    totImporte = Val(Me.txtCantidad) * Val(Me.txt_PrecioV)
    Me.txtImporte.Value = FormatNumber(totImporte, 2)

End Sub





Private Sub LimpiarControles()
        'Limpia los controles
        Me.ComboBox1.Text = ""
        Me.txt_nombre = ""
        Me.txtCantidad = ""
        Me.txt_PrecioV = ""
        Me.txtImporte = ""
End Sub

'Private Sub txtCantidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'' Validación para que el control solo acepte números
'    If KeyAscii < 48 Or KeyAscii > 57 Then
'        KeyAscii = 0
'    End If
'End Sub


Public Sub ctrls_FormatoMoneda()
On Error Resume Next
    Me.txtSubtotal.Text = FormatNumber(Me.txtSubtotal.Text, 2)
    Me.txtTotal.Text = FormatNumber(Me.txtTotal.Text, 2)
End Sub

