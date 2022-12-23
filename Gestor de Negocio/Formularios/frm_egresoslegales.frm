VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_egresoslegales 
   Caption         =   "UserForm1"
   ClientHeight    =   5220
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   5430
   OleObjectBlob   =   "frm_egresoslegales.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_egresoslegales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_Monto_Change()
txt_monto.BackColor = &H80000005
End Sub

Private Sub txt_detalle_Change()
txt_detalle.BackColor = &H80000005
End Sub

Private Sub UserForm_Initialize()
Me.Text_fecha = Date
Me.Label16.Caption = "No. " & Hoja22.Range("L2").Value + 1 'Llamamos el número de la factura

End Sub

Private Sub btn_FechaFact_Click()
banderaCalendario = 15
    Call LanzarCalendario(Me, "lbl_FechaSal")
End Sub
Private Sub LimpiarControles()
        Me.ComboBox1.Text = ""
        Me.Text_fecha = ""
        Me.txt_monto = ""
        Me.txt_detalle.Text = ""
       
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
Hoja22.Range("L2").Value = Hoja22.Range("L2").Value + 1
Comprb = Hoja22.Range("L2").Value

On Error GoTo Salir
Titulo = "Gestor Administrativo"

'Validación para evitar los controles vacíos
If Me.ComboBox1.Text = "" Then
    Me.ComboBox1.BackColor = &HC0C0FF
    MsgBox "Seleccione el ÁREA", vbInformation, Titulo
    Me.ComboBox1.SetFocus
    Exit Sub
End If

    If Me.txt_monto.Text = "" Then
        Me.txt_monto.BackColor = &HC0C0FF
        MsgBox "Seleccione el DETALLE", vbInformation, Titulo
        Me.txt_monto.SetFocus
        Exit Sub
    End If
        If Me.txt_detalle.Text = "" Then
        Me.txt_detalle.BackColor = &HC0C0FF
        MsgBox "Seleccione el DETALLE", vbInformation, Titulo
        Me.txt_detalle.SetFocus
        Exit Sub
    End If

     Hoja14.Select
    Hoja14.Range("A2:G2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja14.Range("A3:G3").Select
    Selection.Copy
    Hoja14.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja14.Cells(2, 1) = CDate(Me.Text_fecha)
        Hoja14.Cells(2, 3) = Comprb
        Hoja14.Cells(2, 4) = Me.ComboBox1.Text
        Hoja14.Cells(2, 5) = Me.txt_monto.Text
        Hoja14.Cells(2, 6) = Me.txt_detalle.Value
        Hoja14.Cells(2, 7) = Hoja21.Cells(1, 7)
     
        LimpiarControles
        
        ComboBox1.SetFocus
              
     
Salir:
     If Err <> 0 Then
        MsgBox Err.Description, vbExclamation, Titulo
     End If
   Me.Text_fecha = Date
   Me.Label16.Caption = "No. " & Hoja22.Range("L2").Value + 1 'Llamamos el número de la factura
    
End Sub
Private Sub btn_Salir_Click()
Unload Me
End Sub
Private Sub ComboBox1_Change()

ComboBox1.BackColor = &H80000005

End Sub

Private Sub ComboBox1_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox1.ListCount
    ComboBox1.RemoveItem 0
Next Fila

            For Fila = 2 To 6
                Lista = Hoja1.Cells(Fila, 54)
                ComboBox1.AddItem (Lista)
        Next
End Sub


