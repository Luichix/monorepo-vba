VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ConsultaProducto 
   Caption         =   "CONSULTAR MOVIMIENTOS POR PRODUCTO"
   ClientHeight    =   6860
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   11700
   OleObjectBlob   =   "frm_ConsultaProducto.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ConsultaProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Long, j As Long, items As Long, items2 As Long



Private Sub btn_Limpiar_Click()
    ComboBox1 = Empty
End Sub

Private Sub CommandButton1_Click()
Unload Me

End Sub

Private Sub UserForm_Initialize()
'Me.txt_Buscar.MaxLength = Hoja12.Range("C3").Value
'Le digo cuántas columnas
    ListBox1.ColumnCount = 5
    ListBox2.ColumnCount = 5

    'Asigno el ancho a cada columna
    Me.ListBox1.ColumnWidths = "60 pt;40 pt;70 pt;60 pt;8 pt"
    Me.ListBox2.ColumnWidths = "60 pt;40 pt;70 pt;60 pt;8 pt"

        'El origen de los datos es la Tabla1
         '   ListBox1.RowSource = "Tabla1"
End Sub
Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long
'Rutina que permite reflejar el resto de la información en los demás controles
'después de haber realizado una selección en el ComboBox

Me.ComboBox1.BackColor = &H80000005

If ComboBox1.Text = "" Then
            Me.ListBox1.Clear
            Me.ListBox2.Clear
            Me.txt_item = Empty
            Me.txt_Descrip = Empty
            Me.txt_Medida = Empty
            Me.txt_clase = Empty
            Me.txt_Saldo = Empty
            Me.txt_cantNeta = Empty
            Me.txt_costoNeto = Empty
            Me.txt_CostoFinal = Empty
            Me.lbl_TotSalidas = Empty
            Me.txt_CantVenta = Empty
            Me.txt_CostoVentas = emtpy
            Me.ComboBox1.SetFocus
            Exit Sub
                
End If

  
    Final = GetUltimoR(Hoja12)
    
    
    'Solicito la información de la hoja de productos para que se reflejen en los controles
    For Fila = 2 To Final
        If ComboBox1.Text = Hoja12.Cells(Fila, 1) Then

        End If
    Next
    
       
    Final = GetUltimoR(Hoja12)
    
    'Solicito información de la hoja de existencias para reflejarlas en los respectivos controles
    For Fila = 2 To Final
        If ComboBox1.Text = Hoja12.Cells(Fila, 1) Then
          
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

Private Sub btn_Buscar_Click()
Dim Fila As Long, Final As Long
        
On Error GoTo Salir
Application.ScreenUpdating = False
Hoja10.Unprotect "355365847"
Hoja11.Unprotect "355365847"
Hoja12.Unprotect "355365847"
'Buscamos la última fila en la hoja de existencias
Fila = 2
    Do While Hoja12.Cells(Fila, 1) <> Empty
        Fila = Fila + 1
    Loop
    Final = Fila - 1
    

'Solicito datos desde la hoja de productos
    For Fila = 2 To Final
        If Me.ComboBox1.Text = Hoja12.Cells(Fila, 1) Then
            Me.txt_item = Hoja12.Cells(Fila, 2)
            Me.txt_Descrip = Hoja12.Cells(Fila, 1)
            Me.txt_Medida = Hoja12.Cells(Fila, 3)
            Me.txt_clase = Hoja12.Cells(Fila, 4)
            Exit For
        
        End If
    Next

'Solicitamos datos de la hoja de existencias.
    For Fila = 2 To Final
        If Me.ComboBox1.Text = Hoja12.Cells(Fila, 1) Then
            Me.txt_Saldo = Hoja12.Cells(Fila, 13)
            Me.txt_CostoFinal = "C$" & "      " & FormatNumber(Hoja12.Cells(Fila, 15), 2)
            
            Exit For
        
        End If
    Next
'--------------------------------------

        
        If Me.ComboBox1.Text = Empty Then
            Me.ListBox1.Clear
            Me.ListBox2.Clear
            Me.txt_item = Empty
            Me.txt_Descrip = Empty
            Me.txt_Medida = Empty
            Me.txt_clase = Empty
            Me.txt_Saldo = Empty
            Me.txt_cantNeta = Empty
            Me.txt_costoNeto = Empty
            Me.txt_CostoFinal = Empty
            Me.lbl_TotSalidas = Empty
            Me.txt_CantVenta = Empty
            Me.txt_CostoVentas = emtpy
            MsgBox "Escriba un código para buscar", vbExclamation
            Me.ComboBox1.SetFocus
            Exit Sub
                
End If

Me.ListBox1.Clear
Me.ListBox2.Clear

'Mostrar ENTRADAS Registro_Entradas

items = Hoja10.Range("Registro_Entradas").CurrentRegion.Rows.Count
        For i = 2 To items
            If Hoja10.Cells(i, 6).Value Like Me.ComboBox1.Text Then
                Me.ListBox1.AddItem Hoja10.Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Hoja10.Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Hoja10.Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Hoja10.Cells(i, 9)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Hoja10.Cells(i, 7)
            End If
        Next i


'Mostrar SALIDAS Registro_Salidas

items2 = Hoja11.Range("Registro_Salidas").CurrentRegion.Rows.Count
        For j = 2 To items2
            If Hoja11.Cells(j, 5).Value Like Me.ComboBox1.Text Then
                Me.ListBox2.AddItem Hoja11.Cells(j, 1)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 1) = Hoja11.Cells(j, 10)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 2) = Hoja11.Cells(j, 3)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 3) = Hoja11.Cells(j, 8)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 4) = Hoja11.Cells(j, 6)
            End If
        Next j

        Me.ComboBox1.SetFocus
        Me.ComboBox1.SelStart = 0
        Me.ComboBox1.SelLength = Len(Me.ComboBox1.Text)

Call ComprasNetas
Call CostoVentas
Hoja10.Protect "355365847"
Hoja11.Protect "355365847"
Hoja12.Protect "355365847"
Application.ScreenUpdating = True
Exit Sub

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If


End Sub

'Calculo de Compras Netas Registro_Entradas

Public Sub ComprasNetas()
Dim cantNeta As Integer, costoNeto As Currency
items = Hoja10.Range("Registro_Entradas").CurrentRegion.Rows.Count
cantNeta = 0
costoNeto = 0
        For i = 2 To items
            If Hoja10.Cells(i, 6).Value Like Me.ComboBox1.Text Then
             cantNeta = cantNeta + Val(Hoja10.Cells(i, 7))
             costoNeto = costoNeto + Val(Hoja10.Cells(i, 10))
            End If
        Next i
Me.txt_cantNeta = cantNeta
Me.txt_costoNeto = FormatNumber(costoNeto, 2)

End Sub

'Calcular Costo de Venta Registro_Salidas

Public Sub CostoVentas()
Dim cVenta As Currency, cantVenta As Integer
items = Hoja11.Range("Registro_Salidas").CurrentRegion.Rows.Count
cVenta = 0
cantVenta = 0
        For i = 2 To items
            If Hoja11.Cells(i, 5).Value Like Me.ComboBox1.Text Then
                cantVenta = cantVenta + Val(Hoja11.Cells(i, 6))
                cVenta = cVenta + Val(Hoja11.Cells(i, 9))
            End If
        Next i

Me.txt_CantVenta.Value = cantVenta
Me.txt_CostoVentas = FormatNumber(cVenta, 2)

End Sub


