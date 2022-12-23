VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_egresosmadera 
   Caption         =   "EGRESOS DE MADERA"
   ClientHeight    =   4350
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   4840
   OleObjectBlob   =   "frm_egresosmadera.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_egresosmadera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_cantidad_Change()
txt_cantidad.BackColor = &H80000005
End Sub

Private Sub txt_costo_Change()
txt_costo.BackColor = &H80000005
End Sub

Private Sub UserForm_Initialize()
Me.Text_fecha = Date
Me.Label16.Caption = "No. " & Hoja22.Range("I2").Value + 1 'Llamamos el número de la factura

End Sub

Private Sub btn_FechaFact_Click()
banderaCalendario = 12
    Call LanzarCalendario(Me, "lbl_FechaSal")
End Sub
Private Sub LimpiarControles()
        Me.ComboBox1.Text = ""
        Me.Text_fecha = ""
        Me.txt_cantidad = ""
        Me.txt_costo.Text = ""
       
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
Hoja22.Range("I2").Value = Hoja22.Range("I2").Value + 1
Comprb = Hoja22.Range("I2").Value

On Error GoTo Salir
Titulo = "Gestor Administrativo"

'Validación para evitar los controles vacíos
If Me.ComboBox1.Text = "" Then
    Me.ComboBox1.BackColor = &HC0C0FF
    MsgBox "Seleccione el DETALLE", vbInformation, Titulo
    Me.ComboBox1.SetFocus
    Exit Sub
End If
    If Me.txt_cantidad.Text = "" Then
        Me.txt_cantidad.BackColor = &HC0C0FF
        MsgBox "Seleccione el DETALLE", vbInformation, Titulo
        Me.txt_cantidad.SetFocus
        Exit Sub
    End If
        If Me.txt_costo.Text = "" Then
        Me.txt_costo.BackColor = &HC0C0FF
        MsgBox "Seleccione el DETALLE", vbInformation, Titulo
        Me.txt_costo.SetFocus
        Exit Sub
    End If

     Hoja8.Select
    Hoja8.Range("A2:G2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja8.Range("A3:G3").Select
    Selection.Copy
    Hoja8.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja8.Cells(2, 1) = CDate(Me.Text_fecha)
        Hoja8.Cells(2, 2) = Comprb
        Hoja8.Cells(2, 3) = Me.ComboBox1.Text
        Hoja8.Cells(2, 4) = Me.txt_cantidad.Text
        Hoja8.Cells(2, 5) = Me.txt_costo.Value
        Hoja8.Cells(2, 7) = Hoja21.Cells(1, 7)
     
        LimpiarControles
        
        ComboBox1.SetFocus
        
        
       
Salir:
     If Err <> 0 Then
        MsgBox Err.Description, vbExclamation, Titulo
     End If
   Me.Text_fecha = Date
   Me.Label16.Caption = "No. " & Hoja22.Range("I2").Value + 1 'Llamamos el número de la factura
    
End Sub
Private Sub btn_Salir_Click()
Unload Me
End Sub
Private Sub ComboBox1_Change()

ComboBox1.BackColor = &H80000005

End Sub

'Private Sub Combobox1_Enter()
'Dim Fila As Long
'Dim Final As Long
'Dim Lista As String
'
'
'For Fila = 1 To ComboBox1.ListCount
'    ComboBox1.RemoveItem 0
'Next Fila
'
'            For Fila = 2 To 9
'                Lista = Hoja1.Cells(Fila, 48)
'                ComboBox1.AddItem (Lista)
'        Next
'End Sub

                        




