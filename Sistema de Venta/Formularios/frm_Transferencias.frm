VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Transferencias 
   Caption         =   "GESTOR DE INVENTARIO"
   ClientHeight    =   5535
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   12960
   OleObjectBlob   =   "frm_Transferencias.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Transferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_FechaFact_Click()
banderaCalendario = 5
    Call LanzarCalendario(Me, "lbl_FechaSal")
End Sub
Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long
'Rutina que permite reflejar el resto de la información en los demás controles
'después de haber realizado una selección en el ComboBox

Me.ComboBox1.BackColor = &H80000005
Me.txt_Nombre.Text = ""
Me.txt_Salida.Text = ""
Me.txt_saldo.Text = ""
Me.txt_CostoU.Text = ""
Me.TextBox1.Text = ""
Me.TextBox2.Text = ""
Me.txt_Existencia.Text = ""

If ComboBox1.Text = "" Then
    LimpiarControles
End If

  
  '''''''''''''''''''''''

'    'Determino el final de la hoja de productos y de existencias

Fila = 2

    Do While Hoja1.Cells(Fila, 11) <> ""
        Fila = Fila + 1
    Loop

    Final = Fila - 1



    'Solicito la información de la hoja de materiales para que se reflejen en los controles
    For Fila = 2 To Final
        If ComboBox1.Text = Hoja5.Cells(Fila, 1) Then
            Me.txt_Nombre = Hoja5.Cells(Fila, 2)
            Me.TextBox1 = Hoja5.Cells(Fila, 4)
             Me.txt_Existencia.Text = Hoja5.Cells(Fila, 10)
             Me.TextBox2.Text = Hoja5.Cells(Fila, 3)
             Me.txt_CostoU.Text = Hoja5.Cells(Fila, 11)
            Exit For
        End If
    Next

  
      'Solicito la información de la hoja de PRODUCTOS para que se reflejen en los controles
    For Fila = 2 To Final
        If ComboBox1.Text = Hoja6.Cells(Fila, 1) Then
            Me.txt_Nombre = Hoja6.Cells(Fila, 2)
            Me.TextBox1 = Hoja6.Cells(Fila, 4)
             Me.txt_Existencia.Text = Hoja6.Cells(Fila, 10)
              Me.TextBox2.Text = Hoja6.Cells(Fila, 3)
             Me.txt_CostoU.Text = Hoja6.Cells(Fila, 11)
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


Private Sub CommandButton1_Click()
Dim Fila As Long
Dim Final As Long
Dim Final2 As Long
Dim Existencia As Integer
Dim TotalExistencia As Integer
Dim Comprb As Long
Dim vPrecioVenta As Currency
Dim CostoTotal As Currency
Dim cUpromedio As Currency
Dim Titulo As String

On Error GoTo Salir

Application.ScreenUpdating = False

Titulo = "Gestor de Inventario"
'Validación para evitar los controles vacíos
If Me.ComboBox1.Text = "" Then
    Me.ComboBox1.BackColor = &HC0C0FF
    MsgBox "Ingrese una Descripción", , Titulo
    Me.ComboBox1.SetFocus
    Exit Sub
         ElseIf Me.txt_FechaSal = "" Then
            Me.txt_FechaSal.BackColor = &HC0C0FF
            MsgBox "Introduzca la fecha de salida", , Titulo
            Me.txt_FechaSal.SetFocus
            Exit Sub
                
                        ElseIf Me.txt_Destino = "" Then
                            Me.txt_Destino.BackColor = &HC0C0FF
                            MsgBox "Ingrese el destino", , Titulo
                            Me.txt_Destino.SetFocus
                            Exit Sub
                                ElseIf Me.txt_Salida = "" Then
                                    Me.txt_Salida.BackColor = &HC0C0FF
                                    MsgBox "Ingrese una cantidad", , Titulo
                                    Me.txt_Salida.SetFocus
                                    Exit Sub
End If


Buscar_x
If Hoja4.Visible = xlSheetVisible Then
    Hoja4.Select
    Hoja4.Cells(1, 1).Select
End If
Application.ScreenUpdating = True
    Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, Titulo
 End If

End Sub
Private Sub Buscar_x()
'''''''''''''''''''''''''''''
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
                Range("b1").Select
                    Do Until IsEmpty(ActiveCell)
                          ActiveCell.Offset(1, 0).Select
                          If ActiveCell.Value Like X Then
                              encontrado = True
                              Exit Do
                                                                  
                          End If
                          
                    Loop
                  
                  If encontrado = True Then
                        Agregar_Registro
                            
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
                Range("b1").Select
                    Do Until IsEmpty(ActiveCell)
                          ActiveCell.Offset(1, 0).Select
                          If ActiveCell.Value Like X Then
                              encontrado = True
                              Exit Do
                                                                  
                          End If
                          
                    Loop
                
                If encontrado = True Then
                    Agregar_Registro
                    
                Else: encontrado = False
                    MsgBox "Producto no Existente", vbInformation, Titulo
                End If
                
                Hoja1.Visible = xlSheetVeryHidden
                
End If
''''''''''''''''''''''''''''''''
End Sub
Private Sub Agregar_Registro()


'Aquí manejo el correlativo del comprobante
Hoja93.Range("B2").Value = Hoja93.Range("B2").Value + 1
Comprb = "TRANSFERENCIA N° " & Hoja93.Range("B2").Value

'Envía los datos a la hoja de salidas
Hoja4.Unprotect ""
Hoja5.Unprotect ""

If Hoja4.Visible = xlSheetVisible Then

Hoja4.Select
    Hoja4.Range("A2:I2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja4.Range("A3:I3").Select
    Selection.Copy
    Hoja4.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja4.Cells(2, 1) = CDate(Me.txt_FechaSal)
        Hoja4.Cells(2, 3) = Me.txt_Destino
        Hoja4.Cells(2, 5) = Me.ComboBox1.Text
        Hoja4.Cells(2, 6) = Me.txt_Salida.Text
        Hoja4.Cells(2, 8) = Me.txt_CostoU.Value
        Hoja4.Cells(2, 10) = Comprb
        Hoja4.Cells(2, 11) = Me.TextBox1.Text
        Hoja4.Cells(2, 12) = Hoja92.Range("G1")
        
ElseIf Hoja4.Visible = xlSheetVeryHidden Then
    Hoja4.Visible = xlSheetVisible
    
    Hoja4.Select
    Hoja4.Range("A2:I2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja4.Range("A3:I3").Select
    Selection.Copy
    Hoja4.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja4.Cells(2, 1) = CDate(Me.txt_FechaSal)
        Hoja4.Cells(2, 3) = Me.txt_Destino
        Hoja4.Cells(2, 5) = Me.ComboBox1.Text
        Hoja4.Cells(2, 6) = Me.txt_Salida.Text
        Hoja4.Cells(2, 8) = Me.txt_CostoU.Value
        Hoja4.Cells(2, 10) = Comprb
        Hoja4.Cells(2, 11) = Me.TextBox1.Text
        Hoja4.Cells(2, 12) = Hoja92.Range("G1")
    
    Hoja4.Visible = xlSheetVeryHidden

End If
        
        LimpiarControles
        
        ComboBox1.SetFocus
        
 Me.Label17.Caption = "No. " & Hoja93.Range("B2").Value + 1 'Llamamos el número de la factura
 Me.txt_FechaSal.BackColor = &H80000018

Hoja4.Protect ""
Hoja5.Protect ""

End Sub
Private Sub CommandButton2_Click()
Unload Me
End Sub





Private Sub txt_Destino_Change()
Me.txt_Destino.BackColor = &H80000005
End Sub

Private Sub txt_FechaSal_Change()
Me.txt_FechaSal.BackColor = &H80000005
End Sub

Private Sub txt_Saldo_Change()
'Validación para que se borre la información de txt_saldo, si el control txt_Salida se encuentra vacío o en cero
If Me.txt_Salida = "" Or Me.txt_Salida = 0 Then
                ahora = 0
                Me.txt_saldo = ""
            End If
End Sub
Private Sub txt_Salida_Change()
Dim Fila As Long
Dim Final As Long
Dim Registro As Long
Dim antes As Long
Dim ahora As Long
Dim saldo As Long

Me.txt_Salida.BackColor = &H80000005

'Con esta rutina actualizo el saldo existencia reflejado en el control txt_Saldo


    'Determino el final del listdo de existencias
Fila = 1
    Do While Hoja1.Cells(Fila, 11) <> Empty
        Fila = Fila + 1
    Loop
Final = Fila

    'Compruebo que el código ingresado en el ComboBox, coincida en hoja de existencias
    ' para realizar la respectiva operación aritmética
    For Registro = 1 To Final
        If ComboBox1.Text = Hoja5.Cells(Registro, 1) Then
            antes = Hoja5.Cells(Registro, 10)
            ahora = Val(Me.txt_Salida)
                
                saldo = antes - ahora
                Me.txt_saldo = saldo
            Exit For

        End If
    Next
For Registro = 1 To Final
        If ComboBox1.Text = Hoja6.Cells(Registro, 1) Then
            antes = Hoja6.Cells(Registro, 10)
            ahora = Val(Me.txt_Salida)
                
                saldo = antes - ahora
                Me.txt_saldo = saldo
            Exit For

        End If
    Next


End Sub
Private Sub LimpiarControles()
        Me.ComboBox1 = ""
        Me.txt_Nombre = ""
'        Me.txt_FechaSal = ""
        Me.txt_Destino = ""
        Me.txt_Salida = ""
        Me.txt_CostoU = ""
        Me.txt_Existencia = ""
        Me.TextBox1 = ""
        Me.TextBox2 = ""
       
End Sub
Private Sub UserForm_Initialize()
Me.txt_FechaSal = Date
 Me.Label17.Caption = "No. " & Hoja93.Range("B2").Value + 1 'Llamamos el número de la factura

End Sub

Private Sub txt_Destino_Enter()
Dim Fila As Long
Dim xFila As Long
Dim Final As Long
Dim Lista As String

'Toda esta rutina agrega los items al ComboBox

For Fila = 1 To txt_Destino.ListCount
    txt_Destino.RemoveItem 0
Next Fila

    xFila = 2
    
    Do While Hoja1.Cells(xFila, 24) <> ""
        xFila = xFila + 1
    Loop
    
    Final = xFila - 1

    For Fila = 2 To Final
        Lista = Hoja1.Cells(Fila, 24)
        txt_Destino.AddItem (Lista)
    Next
End Sub
