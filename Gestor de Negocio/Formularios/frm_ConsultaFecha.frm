VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ConsultaFecha 
   Caption         =   "CONSULTAR MOVIMIENTOS POR RANGO DE FECHA"
   ClientHeight    =   5970
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   11820
   OleObjectBlob   =   "frm_ConsultaFecha.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ConsultaFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim i As Long, j As Long, items As Long, items2 As Long

Private Sub btn_Buscar_Click()
Dim Fila As Long, Final As Long
        
On Error GoTo Salir
        
'Buscamos la última fila en la hoja de existencias
Fila = 2
    Do While Hoja5.Cells(Fila, 1) <> Empty
        Fila = Fila + 1
    Loop
    Final = Fila - 1
    

'Solicito datos desde la hoja de productos
    For Fila = 2 To Final
        If Me.txt_Buscar.Text = Hoja2.Cells(Fila, 1) Then
            Me.txt_nombre = Hoja2.Cells(Fila, 2)
            Me.txt_Descrip = Hoja2.Cells(Fila, 3)
            Exit For
        
        End If
    Next

'Solicitamos datos de la hoja de existencias.
    For Fila = 2 To Final
        If Me.txt_Buscar.Text = Hoja5.Cells(Fila, 1) Then
            Me.txt_Saldo = Hoja5.Cells(Fila, 3)
            Me.txt_CostoFinal = FormatNumber(Hoja5.Cells(Fila, 6), 2)
            
            Exit For
        
        End If
    Next
'--------------------------------------

        
        If Me.txt_Buscar.Text = Empty Then
            Me.ListBox1.Clear
            Me.ListBox2.Clear
            Me.txt_nombre = Empty
            Me.txt_Descrip = Empty
            Me.txt_Saldo = Empty
            Me.txt_CostoFinal = Empty
            Me.txtFecha1 = Empty
            Me.txtFecha2 = Empty
            MsgBox "Escriba un código para buscar", vbExclamation
            Me.txt_Buscar.SetFocus
            Exit Sub
                ElseIf Me.txtFecha1 = Empty Then
                    Me.txtFecha1.SetFocus
                    MsgBox "Ingrese la fecha inicial", vbExclamation
                    Exit Sub
                        ElseIf Me.txtFecha2 = Empty Then
                            Me.txtFecha2.SetFocus
                            MsgBox "Ingrese la fecha final", vbExclamation
                            Exit Sub
End If

Me.ListBox1.Clear
Me.ListBox2.Clear


'Mostrar ENTRADAS

items = Hoja3.Range("tbl_Entradas").CurrentRegion.Rows.Count
        For i = 2 To items
            If Hoja3.Cells(i, 2).Value Like Me.txt_Buscar.Text _
            And CDate(Hoja3.Cells(i, 4).Value) >= CDate(Me.txtFecha1) _
            And CDate(Hoja3.Cells(i, 4).Value) <= CDate(Me.txtFecha2) Then

                Me.ListBox1.AddItem Hoja3.Cells(i, 4)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Hoja3.Cells(i, 5)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Hoja3.Cells(i, 6)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Hoja3.Cells(i, 8)
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Hoja3.Cells(i, 7)
            End If
        Next i
        
'Mostrar SALIDAS

items2 = Hoja4.Range("tbl_Salidas").CurrentRegion.Rows.Count
        For j = 2 To items2
            If Hoja4.Cells(j, 2).Value Like Me.txt_Buscar.Text _
            And CDate(Hoja4.Cells(j, 4).Value) >= CDate(Me.txtFecha1) _
            And CDate(Hoja4.Cells(j, 4).Value) <= CDate(Me.txtFecha2) Then
            
                Me.ListBox2.AddItem Hoja4.Cells(j, 4)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 1) = Hoja4.Cells(j, 5)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 2) = Hoja4.Cells(j, 6)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 3) = Hoja4.Cells(j, 8)
                Me.ListBox2.List(Me.ListBox2.ListCount - 1, 4) = Hoja4.Cells(j, 7)
            End If
        Next j

        
        
        Me.txt_Buscar.SetFocus
        Me.txt_Buscar.SelStart = 0
        Me.txt_Buscar.SelLength = Len(Me.txt_Buscar.Text)
        


Exit Sub
   
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If

End Sub


Private Sub btn_Fecha1_Click()
banderaCalendario = 3
    Call LanzarCalendario(Me, "txtFecha1")
End Sub

Private Sub btn_Fecha2_Click()
banderaCalendario = 4
    Call LanzarCalendario(Me, "txtFecha2")
End Sub

Private Sub txt_Buscar_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
' Validación para que el control solo acepte números
If Hoja12.Range("C2") = True Then
    If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
    End If
End If
End Sub


Private Sub txtFecha1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
' Validación para que el TextBox, solo acepte números
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If

End Sub

Private Sub txtFecha2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
' Validación para que el TextBox, solo acepte números
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If

End Sub

Private Sub UserForm_Initialize()
Me.txt_Buscar.MaxLength = Hoja12.Range("C3").Value
'Le digo cuántas columnas
    ListBox1.ColumnCount = 5
    ListBox2.ColumnCount = 5
    
    'Asigno el ancho a cada columna
    Me.ListBox1.ColumnWidths = "60 pt;40 pt;70 pt;60 pt;8 pt"
    Me.ListBox2.ColumnWidths = "60 pt;40 pt;70 pt;60 pt;8 pt"
            
        'El origen de los datos es la Tabla1
         '   ListBox1.RowSource = "Tabla1"
End Sub

