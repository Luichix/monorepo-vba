VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Transferencias 
   Caption         =   "SALIDAS DE INVENTARIO"
   ClientHeight    =   3870
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   10180
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

If ComboBox1.Text = "" Then
    LimpiarControles
End If

  
    Final = GetUltimoR(Hoja12)
    
    
    'Solicito la información de la hoja de productos para que se reflejen en los controles
    For Fila = 2 To Final
        If ComboBox1.Text = Hoja12.Cells(Fila, 1) Then
            Me.txt_FechaSal = Date
            Me.txt_nombre = Hoja12.Cells(Fila, 2)
            Exit For
        End If
    Next
    
       
    Final = GetUltimoR(Hoja12)
    
    'Solicito información de la hoja de existencias para reflejarlas en los respectivos controles
    For Fila = 2 To Final
        If ComboBox1.Text = Hoja12.Cells(Fila, 1) Then
            Me.txt_Existencia.Text = Hoja12.Cells(Fila, 13)
            Me.txt_CostoU.Text = Hoja12.Cells(Fila, 14)
                                    
            
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

Titulo = "Gestor Administrativo"
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

'vPrecioVenta = Me.txt_PrecioV.Value

'Aquí manejo el correlativo del comprobante
Hoja22.Range("B2").Value = Hoja22.Range("B2").Value + 1
Comprb = Hoja22.Range("B2").Value

    'Determina el final del listado de salidas
    'Final = GetNuevoR(Hoja4)
    
    
        'Envía los datos a la hoja de salidas
Hoja11.Unprotect "355365847"
Hoja12.Unprotect "355365847"

Hoja11.Select
    Hoja11.Range("A2:I2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja11.Range("A3:I3").Select
    Selection.Copy
    Hoja11.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja11.Cells(2, 1) = CDate(Me.txt_FechaSal)
        Hoja11.Cells(2, 3) = Me.txt_Destino
        Hoja11.Cells(2, 5) = Me.ComboBox1.Text
        Hoja11.Cells(2, 6) = Me.txt_Salida.Text
        Hoja11.Cells(2, 8) = Me.txt_CostoU.Value
        Hoja11.Cells(2, 10) = Comprb
        Hoja11.Cells(2, 11) = Hoja21.Range("G1")
        
        
        'Hoja11.Cells(Final, 1) = "T-" & Comprb
        'Hoja11.Cells(Final, 3) = Me.txt_Nombre
        'Hoja11.Cells(Final, 5) = 0
        'Hoja11.Cells(Final, 8) = "Transf."
        'Hoja11.Cells(Final, 9) = vPrecioVenta
        'Hoja11.Cells(Final, 11) = "-" & Me.txt_Salida.Text * CostoUnitario 'Obtengo el costo total
        'Hoja11.Cells(Final, 17) = Hoja8.Range("G1") 'Usuario responsable de la operación

 
           
        'Limpia los controles
        LimpiarControles
        
        ComboBox1.SetFocus
        
 Me.Label17.Caption = "No. " & Hoja22.Range("B2").Value + 1 'Llamamos el número de la factura
Hoja11.Protect "355365847"
Hoja12.Protect "355365847"
Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, Titulo
 End If

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
                Me.txt_Saldo = ""
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
    Do While Hoja12.Cells(Fila, 1) <> Empty
        Fila = Fila + 1
    Loop
Final = Fila

    'Compruebo que el código ingresado en el ComboBox, coincida en hoja de existencias
    ' para realizar la respectiva operación aritmética
    For Registro = 1 To Final
        If ComboBox1.Text = Hoja12.Cells(Registro, 1) Then
            antes = Hoja12.Cells(Registro, 13)
            ahora = Val(Me.txt_Salida)
                
                saldo = antes - ahora
                Me.txt_Saldo = saldo
            Exit For

        End If
    Next

End Sub
Private Sub LimpiarControles()
        Me.ComboBox1.Text = ""
        Me.txt_nombre = ""
        Me.txt_FechaSal = ""
        Me.txt_Destino = ""
        Me.txt_Salida = ""
        Me.txt_CostoU = ""
        Me.txt_Existencia = ""
       
End Sub
Private Sub UserForm_Initialize()
Me.txt_FechaSal = Date
 Me.Label17.Caption = "No. " & Hoja22.Range("B2").Value + 1 'Llamamos el número de la factura

End Sub

Private Sub txt_Destino_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String

'Toda esta rutina agrega los items al ComboBox

For Fila = 1 To txt_Destino.ListCount
    txt_Destino.RemoveItem 0
Next Fila


    
    For Fila = 2 To 9
        Lista = Hoja1.Cells(Fila, 19)
        txt_Destino.AddItem (Lista)
    Next
End Sub
