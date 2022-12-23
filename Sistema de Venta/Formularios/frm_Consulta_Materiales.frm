VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Consulta_Materiales 
   Caption         =   "GESTOR DE INVENTARIOS"
   ClientHeight    =   9780.001
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   13750
   OleObjectBlob   =   "frm_Consulta_Materiales.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Consulta_Materiales"
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

Private Sub CommandButton5_Click()
If Hoja6.Visible = xlSheetVisible Then
    Hoja6.Select
    Hoja6.Cells(1, 1).Select
    Unload Me
    Reportes_Inventario
ElseIf Hoja6.Visible = xlSheetVeryHidden Then
    Hoja14.Visible = xlSheetVisible
    Hoja6.Visible = xlSheetVisible
          Hoja6.Select
    Hoja6.Cells(1, 1).Select
    Unload Me
    Reportes_Inventario
    Hoja14.Visible = xlSheetVeryHidden
    Hoja6.Visible = xlSheetVeryHidden
End If
End Sub

Private Sub CommandButton6_Click()
If Hoja5.Visible = xlSheetVisible Then
    
    Hoja5.Select
    Hoja5.Cells(1, 1).Select
    Unload Me
    Reportes_Inventario
ElseIf Hoja5.Visible = xlSheetVeryHidden Then
    Hoja14.Visible = xlSheetVisible
    Hoja5.Visible = xlSheetVisible
    Hoja5.Select
    Hoja5.Cells(1, 1).Select
    
    Unload Me
    Reportes_Inventario
    Hoja14.Visible = xlSheetVeryHidden
    Hoja5.Visible = xlSheetVeryHidden
End If
End Sub

Private Sub UserForm_Initialize()
    ListBox1.ColumnCount = 5
    ListBox2.ColumnCount = 5
    ListBox3.ColumnCount = 5
    ListBox4.ColumnCount = 5
    
    Me.ListBox1.ColumnWidths = "60 pt;55 pt;70 pt;85 pt;8 pt"
    Me.ListBox2.ColumnWidths = "60 pt;90 pt;65 pt;60 pt;8 pt"
    Me.ListBox3.ColumnWidths = "60 pt;55 pt;70 pt;85 pt;8 pt"
    Me.ListBox4.ColumnWidths = "60 pt;90 pt;65 pt;60 pt;8 pt"
End Sub
Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long

Me.ComboBox1.BackColor = &H80000005

If ComboBox1.Text = "" Then
            Me.ListBox1.Clear
            Me.ListBox2.Clear
            Me.txt_item = Empty
            Me.txt_Descrip = Empty
            Me.txt_medida = Empty
            Me.txt_clase = Empty
            Me.txt_saldo = Empty
            Me.txt_cantNeta = Empty
            Me.txt_CostoNeto = Empty
            Me.txt_CostoFinal = Empty
            Me.txt_CantVenta = Empty
            Me.txt_CostoVentas = emtpy
            Me.txt_CostoPromedio = Empty
            Me.ComboBox1.SetFocus
            Exit Sub
                
End If

    Final = GetUltimoR(Hoja5)
    
    
    'Solicito la información de la hoja de productos para que se reflejen en los controles
    For Fila = 2 To Final
        If ComboBox1.Text = Hoja5.Cells(Fila, 1) Then

        End If
    Next
    
       
    Final = GetUltimoR(Hoja5)
    
    'Solicito información de la hoja de existencias para reflejarlas en los respectivos controles
    For Fila = 2 To Final
        If ComboBox1.Text = Hoja5.Cells(Fila, 1) Then
          
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
    Final = GetUltimoR(Hoja5)
    
    'Agrego el listado de códigos de productos al ComboBox desde la hoja de productos
    For Fila = 2 To Final
        Lista = Hoja5.Cells(Fila, 1)
        ComboBox1.AddItem (Lista)
    Next
End Sub

Private Sub btn_Buscar_Click()
Dim Fila As Long, Final As Long
        
On Error GoTo Salir
Application.ScreenUpdating = False
Hoja3.Unprotect ""
Hoja4.Unprotect ""
Hoja5.Unprotect ""
'Buscamos la última fila en la hoja de existencias
Fila = 2
    Do While Hoja5.Cells(Fila, 1) <> Empty
        Fila = Fila + 1
    Loop
    Final = Fila - 1
    

'Solicito datos desde la hoja de productos
    For Fila = 2 To Final
        If Me.ComboBox1.Text = Hoja5.Cells(Fila, 1) Then
            Me.txt_item = Hoja5.Cells(Fila, 2)
            Me.txt_Descrip = Hoja5.Cells(Fila, 1)
            Me.txt_medida = Hoja5.Cells(Fila, 3)
            Me.txt_clase = Hoja5.Cells(Fila, 4)
            Exit For
        
        End If
    Next

'Solicitamos datos de la hoja de existencias.
    For Fila = 2 To Final
        If Me.ComboBox1.Text = Hoja5.Cells(Fila, 1) Then
            Me.txt_saldo = Hoja5.Cells(Fila, 10)
            Me.txt_CostoFinal = "C$" & "      " & FormatNumber(Hoja5.Cells(Fila, 7), 2)
            Me.txt_CostoPromedio = "C$" & "      " & FormatNumber(Hoja5.Cells(Fila, 11), 2)
              Me.txt_cantNeta = FormatNumber(Hoja5.Cells(Fila, 8), 0)
              Me.txt_CostoNeto = "C$" & "      " & FormatNumber(Hoja5.Cells(Fila, 5), 2)
              Me.txt_CantVenta = FormatNumber(Hoja5.Cells(Fila, 9), 0)
              Me.txt_CostoVentas = "C$" & "      " & FormatNumber(Hoja5.Cells(Fila, 6), 2)
            
            Exit For
        
        End If
    Next
'--------------------------------------

        
        If Me.ComboBox1.Text = Empty Then
            Me.ListBox1.Clear
            Me.ListBox2.Clear
            Me.txt_item = Empty
            Me.txt_Descrip = Empty
            Me.txt_medida = Empty
            Me.txt_clase = Empty
            Me.txt_saldo = Empty
            Me.txt_cantNeta = Empty
            Me.txt_CostoNeto = Empty
            Me.txt_CostoFinal = Empty
            Me.txt_CantVenta = Empty
            Me.txt_CostoVentas = emtpy
            Me.txt_CostoPromedio = Empty
            MsgBox "Escriba un código para buscar", vbExclamation
            Me.ComboBox1.SetFocus
            Exit Sub
                
End If

Me.ListBox1.Clear
Me.ListBox2.Clear

'Mostrar ENTRADAS Registro_Entradas

items = Hoja3.Range("Registro_Entradas").CurrentRegion.Rows.Count
        For i = 2 To items
            If Hoja3.Cells(i, 6).Value Like Me.ComboBox1.Text Then
                Me.ListBox1.AddItem Hoja3.Cells(i, 1)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Hoja3.Cells(i, 3)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Hoja3.Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Hoja3.Cells(i, 9)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Hoja3.Cells(i, 7)
            End If
        Next i


'Mostrar SALIDAS Registro_Salidas

items2 = Hoja4.Range("Registro_Salidas").CurrentRegion.Rows.Count
        For j = 2 To items2
            If Hoja4.Cells(j, 5).Value Like Me.ComboBox1.Text Then
                Me.ListBox2.AddItem Hoja4.Cells(j, 1)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 1) = Hoja4.Cells(j, 10)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 2) = Hoja4.Cells(j, 3)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 3) = Hoja4.Cells(j, 8)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 4) = Hoja4.Cells(j, 6)
            End If
        Next j

        Me.ComboBox1.SetFocus
        Me.ComboBox1.SelStart = 0
        Me.ComboBox1.SelLength = Len(Me.ComboBox1.Text)

Hoja3.Protect ""
Hoja4.Protect ""
Hoja5.Protect ""
Application.ScreenUpdating = True
Exit Sub

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If
End Sub

Private Sub CommandButton3_Click()
    ComboBox2 = Empty
End Sub

Private Sub CommandButton4_Click()
    Unload Me
End Sub

Private Sub ComboBox2_Change()
Dim Fila As Long
Dim Final As Long
Me.ComboBox2.BackColor = &H80000005

If ComboBox2.Text = "" Then
            Me.ListBox3.Clear
            Me.ListBox4.Clear
            Me.txt_item2 = Empty
            Me.txt_Descrip2 = Empty
            Me.txt_Medida2 = Empty
            Me.txt_clase2 = Empty
            Me.txt_Saldo2 = Empty
            Me.txt_cantNeta2 = Empty
            Me.txt_costoNeto2 = Empty
            Me.txt_CostoFinal2 = Empty
            Me.txt_CantVenta2 = Empty
            Me.txt_CostoVentas2 = emtpy
            Me.txt_CostoPromedio2 = Empty
            Me.ComboBox2.SetFocus
            Exit Sub

End If


    Final = GetUltimoR(Hoja6)


    'Solicito la información de la hoja de productos para que se reflejen en los controles
    For Fila = 2 To Final
        If ComboBox2.Text = Hoja6.Cells(Fila, 1) Then

        End If
    Next

    Final = GetUltimoR(Hoja6)

    'Solicito información de la hoja de existencias para reflejarlas en los respectivos controles
    For Fila = 2 To Final
        If ComboBox2.Text = Hoja6.Cells(Fila, 1) Then

            Exit For
        End If
    Next

End Sub
Private Sub ComboBox2_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String

'Toda esta rutina agrega los items al ComboBox

For Fila = 1 To ComboBox2.ListCount
    ComboBox2.RemoveItem 0
Next Fila

    'Inspecciono la hoja de productos para determinar el final del listado
    Final = GetUltimoR(Hoja6)

    'Agrego el listado de códigos de productos al ComboBox desde la hoja de productos
    For Fila = 2 To Final
        Lista = Hoja6.Cells(Fila, 1)
        ComboBox2.AddItem (Lista)
    Next
End Sub

Private Sub CommandButton2_Click()

Dim Fila As Long, Final As Long

On Error GoTo Salir
Application.ScreenUpdating = False
Hoja3.Unprotect ""
Hoja4.Unprotect ""
Hoja6.Unprotect ""
'Buscamos la última fila en la hoja de existencias
Fila = 2
    Do While Hoja6.Cells(Fila, 1) <> Empty
        Fila = Fila + 1
    Loop
    Final = Fila - 1


'Solicito datos desde la hoja de productos
    For Fila = 2 To Final
        If Me.ComboBox2.Text = Hoja6.Cells(Fila, 1) Then
            Me.txt_item2 = Hoja6.Cells(Fila, 2)
            Me.txt_Descrip2 = Hoja6.Cells(Fila, 1)
            Me.txt_Medida2 = Hoja6.Cells(Fila, 3)
            Me.txt_clase2 = Hoja6.Cells(Fila, 4)
            Exit For

        End If
    Next

'Solicitamos datos de la hoja de existencias.
    For Fila = 2 To Final
        If Me.ComboBox2.Text = Hoja6.Cells(Fila, 1) Then
            Me.txt_Saldo2 = Hoja6.Cells(Fila, 10)
            Me.txt_CostoFinal2 = "C$" & "      " & FormatNumber(Hoja6.Cells(Fila, 7), 2)
            Me.txt_CostoPromedio2 = "C$" & "      " & FormatNumber(Hoja6.Cells(Fila, 11), 2)
              Me.txt_cantNeta2 = FormatNumber(Hoja6.Cells(Fila, 8), 0)
              Me.txt_costoNeto2 = "C$" & "      " & FormatNumber(Hoja6.Cells(Fila, 5), 2)
              Me.txt_CantVenta2 = FormatNumber(Hoja6.Cells(Fila, 9), 0)
              Me.txt_CostoVentas2 = "C$" & "      " & FormatNumber(Hoja6.Cells(Fila, 6), 2)

            Exit For

        End If
    Next
'--------------------------------------

        If Me.ComboBox2.Text = Empty Then
            Me.ListBox3.Clear
            Me.ListBox4.Clear
            Me.txt_item2 = Empty
            Me.txt_Descrip2 = Empty
            Me.txt_Medida2 = Empty
            Me.txt_clase2 = Empty
            Me.txt_Saldo2 = Empty
            Me.txt_cantNeta2 = Empty
            Me.txt_costoNeto2 = Empty
            Me.txt_CostoFinal2 = Empty
            Me.txt_CantVenta2 = Empty
            Me.txt_CostoVentas2 = emtpy
            Me.txt_CostoPromedio2 = Empty
            MsgBox "Escriba un código para buscar", vbExclamation
            Me.ComboBox2.SetFocus
            Exit Sub

        End If

Me.ListBox3.Clear
Me.ListBox4.Clear

'Mostrar ENTRADAS Registro_Entradas

items = Hoja3.Range("Registro_Entradas").CurrentRegion.Rows.Count
        For i = 2 To items
            If Hoja3.Cells(i, 6).Value Like Me.ComboBox2.Text Then
                Me.ListBox3.AddItem Hoja3.Cells(i, 1)
                Me.ListBox3.List(Me.ListBox3.ListCount - 1, 1) = Hoja3.Cells(i, 3)
                Me.ListBox3.List(Me.ListBox3.ListCount - 1, 2) = Hoja3.Cells(i, 4)
                Me.ListBox3.List(Me.ListBox3.ListCount - 1, 3) = Hoja3.Cells(i, 9)
                Me.ListBox3.List(Me.ListBox3.ListCount - 1, 4) = Hoja3.Cells(i, 7)
            End If
        Next i

'Mostrar SALIDAS Registro_Salidas

items2 = Hoja4.Range("Registro_Salidas").CurrentRegion.Rows.Count
        For j = 2 To items2
            If Hoja4.Cells(j, 5).Value Like Me.ComboBox2.Text Then
                Me.ListBox4.AddItem Hoja4.Cells(j, 1)
                Me.ListBox4.List(Me.ListBox4.ListCount - 1, 1) = Hoja4.Cells(j, 10)
                Me.ListBox4.List(Me.ListBox4.ListCount - 1, 2) = Hoja4.Cells(j, 3)
                Me.ListBox4.List(Me.ListBox4.ListCount - 1, 3) = Hoja4.Cells(j, 8)
                Me.ListBox4.List(Me.ListBox4.ListCount - 1, 4) = Hoja4.Cells(j, 6)
            End If
        Next j

        Me.ComboBox2.SetFocus
        Me.ComboBox2.SelStart = 0
        Me.ComboBox2.SelLength = Len(Me.ComboBox2.Text)

Hoja3.Protect ""
Hoja4.Protect ""
Hoja6.Protect ""
Application.ScreenUpdating = True
Exit Sub

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If
End Sub
