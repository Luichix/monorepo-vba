VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_registropagos 
   Caption         =   "PAGOS AL PERSONAL"
   ClientHeight    =   6120
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   4780
   OleObjectBlob   =   "form_registropagos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "form_registropagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox2_Change()
ComboBox2.BackColor = &H80000005
End Sub

Private Sub txt_cantidad_Change()
txt_cantidad.BackColor = &H80000005
End Sub
Private Sub UserForm_Initialize()
Me.Text_fecha = Date
Me.Label16.Caption = "No. " & Hoja22.Range("D2").Value + 1 'Llamamos el número de la factura

End Sub

Private Sub btn_FechaFact_Click()
banderaCalendario = 6
    Call LanzarCalendario(Me, "lbl_FechaSal")
End Sub
Private Sub LimpiarControles()
        Me.ComboBox1.Text = ""
        Me.txt_nombre = ""
        Me.Text_fecha = ""
        Me.ComboBox2 = ""
        Me.txt_area.Text = ""
        Me.txt_cantidad = ""
        Me.ComboBox3 = ""
        
End Sub
Private Sub btn_Registrar_Click()
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
Dim xControl As Control

'Aquí manejo el correlativo del comprobante
Hoja22.Range("D2").Value = Hoja22.Range("D2").Value + 1
Comprb = Hoja22.Range("D2").Value

On Error GoTo Salir
Titulo = "Gestor Administrativo"

If Me.ComboBox1.Text = "" Then
    Me.ComboBox1.BackColor = &HC0C0FF
    MsgBox "Ingrese el código del personal", vbInformation, Titulo
    Me.ComboBox1.SetFocus
    Exit Sub
End If

If Me.ComboBox2.Text = "" Then
    Me.ComboBox2.BackColor = &HC0C0FF
    MsgBox "Ingrese el código del personal", vbInformation, Titulo
    Me.ComboBox2.SetFocus
    Exit Sub
End If
    If Me.ComboBox3.Text = "" Then
        Me.ComboBox3.BackColor = &HC0C0FF
        MsgBox "Ingrese el periodo", vbInformation, Titulo
        Me.ComboBox3.SetFocus
        Exit Sub
    End If

        If Me.Text_fecha = "" Then
            Me.Text_fecha.BackColor = &HC0C0FF
            MsgBox "Introduzca la fecha de salida", , Titulo
            Me.Text_fecha.SetFocus
            Exit Sub
        End If
                           If Me.txt_nombre = "" Then
                                Me.txt_nombre.BackColor = &HC0C0FF
                                MsgBox "Digite el código correctamente", , Titulo
                                Me.txt_nombre.SetFocus
                                Exit Sub
                            End If
                                   If Me.txt_cantidad = "" Then
                                        Me.txt_cantidad.BackColor = &HC0C0FF
                                        MsgBox "Ingrese una cantidad", , Titulo
                                        Me.txt_cantidad.SetFocus
                                        Exit Sub
                                    End If


    Hoja6.Select
    Hoja6.Range("A2:H2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja6.Range("A3:H3").Select
    Selection.Copy
    Hoja6.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja6.Cells(2, 1) = CDate(Me.Text_fecha)
        Hoja6.Cells(2, 3) = Comprb
        Hoja6.Cells(2, 4) = Me.ComboBox1.Text
        Hoja6.Cells(2, 6) = Me.txt_area.Text
        Hoja6.Cells(2, 7) = Me.ComboBox2.Text
        Hoja6.Cells(2, 8) = Me.txt_cantidad.Value
        Hoja6.Cells(2, 10) = Me.ComboBox3.Text
        Hoja6.Cells(2, 12) = Hoja21.Range("G1").Text
     
        LimpiarControles
        
        ComboBox1.SetFocus
        
        
        
Salir:
     If Err <> 0 Then
        MsgBox Err.Description, vbExclamation, Titulo
     End If
   Me.Text_fecha = Date
   Me.Label16.Caption = "No. " & Hoja22.Range("D2").Value + 1 'Llamamos el número de la factura
    
End Sub
Private Sub btn_Salir_Click()
Unload Me
End Sub
Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long
Dim Registro As Integer

ComboBox1.BackColor = &H80000005
If ComboBox1.Text = "" Then
    LimpiarControles
End If

    
Final = GetUltimoR(Hoja5)

    For Fila = 2 To Final
        If ComboBox1.Text = Hoja5.Cells(Fila, 1) Then
            Me.txt_nombre.Text = Hoja5.Cells(Fila, 2)
            Me.txt_area.Text = Hoja5.Cells(Fila, 5)
            Exit For
        
        End If
    Next
End Sub
Private Sub ComboBox1_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox1.ListCount
    ComboBox1.RemoveItem 0
Next Fila


Final = GetUltimoR(Hoja5)

    
        For Fila = 2 To Final
            
            If Hoja5.Cells(Fila, 9) = "ACTIVO" Then
                Lista = Hoja5.Cells(Fila, 1)
                ComboBox1.AddItem (Lista)
            End If
        Next
End Sub

Private Sub ComboBox2_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox2.ListCount
    ComboBox2.RemoveItem 0
Next Fila

   
        For Fila = 2 To 6
                Lista = Hoja1.Cells(Fila, 55)
                ComboBox2.AddItem (Lista)
        Next
End Sub

Private Sub ComboBox3_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox3.ListCount
    ComboBox3.RemoveItem 0
Next Fila

   
        For Fila = 2 To 3
                Lista = Hoja1.Cells(Fila, 66)
                ComboBox3.AddItem (Lista)
        Next
End Sub
         
                        
