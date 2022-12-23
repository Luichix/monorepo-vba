VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ProductoAComprar 
   Caption         =   "COMPRA DE PRODUCTOS"
   ClientHeight    =   3740
   ClientLeft      =   14930
   ClientTop       =   2460
   ClientWidth     =   8160
   OleObjectBlob   =   "frm_ProductoAComprar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ProductoAComprar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btnAgregar_Click()
On Error GoTo Salir
    With frm_fCompras
        .AgregarItems
    End With

               
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor Administrativo"
    
    
 End If
    
 
 ComboBox1.SetFocus
 
End Sub

Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long
'Rutina que permite reflejar el resto de la información en los demás controles
'después de haber realizado una selección en el ComboBox

If ComboBox1.Text = "" Then
    LimpiarControles
End If

    'Determino el final de la hoja de productos y de existencias
Fila = 2
    
    Do While Hoja1.Cells(Fila, 11) <> "" And Hoja12.Cells(Fila, 1) <> ""
        Fila = Fila + 1
    Loop
    
    Final = Fila - 1
    
    
    
    'Solicito la información de la hoja de productos para que se reflejen en los controles
    For Fila = 2 To Final
        If ComboBox1.Text = Hoja12.Cells(Fila, 1) Then
            Me.txt_nombre = Hoja12.Cells(Fila, 2)
            Exit For
        End If
    Next
    
    'Solicito información de la hoja de existencias para reflejarlas en los respectivos controles
    For Fila = 2 To Final
        If ComboBox1.Text = Hoja12.Cells(Fila, 1) Then
            Me.txt_Existencia = Hoja12.Cells(Fila, 13)
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


    'Inspecciono la hoja de productos para determinar el final del listado
Final = GetUltimoR(Hoja12)
    
    
    'Agrego el listado de códigos de productos al ComboBox desde la hoja de productos
    For Fila = 2 To Final
        Lista = Hoja12.Cells(Fila, 1)
        ComboBox1.AddItem (Lista)
    Next
End Sub



Private Sub txtCantidad_Change()
Dim totImporte As Currency


    totImporte = Val(Me.txtCantidad) * Val(Me.txt_CostoU)
    Me.txtCostoTot.Value = FormatNumber(totImporte, 2)

End Sub
Private Sub LimpiarControles()
        'Limpia los controles
        Me.ComboBox1.Text = ""
        Me.txt_nombre = ""
        Me.txtCantidad = ""
        Me.txt_CostoU = ""
        Me.txt_Existencia = ""
        Me.txtCostoTot = ""
End Sub



Private Sub btnCerrar_Click()
Unload Me
End Sub


