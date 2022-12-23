VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_egresosgastos 
   Caption         =   "GASTOS OPERATIVOS"
   ClientHeight    =   5900
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   5560
   OleObjectBlob   =   "frm_egresosgastos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_egresosgastos"
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
Me.Label16.Caption = "No. " & Hoja22.Range("J2").Value + 1 'Llamamos el número de la factura

End Sub

Private Sub btn_FechaFact_Click()
banderaCalendario = 13
    Call LanzarCalendario(Me, "lbl_FechaSal")
End Sub
Private Sub LimpiarControles()
        Me.ComboBox1.Text = ""
        Me.ComboBox2.Text = ""
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
Hoja22.Range("J2").Value = Hoja22.Range("J2").Value + 1
Comprb = Hoja22.Range("J2").Value

On Error GoTo Salir
Titulo = "Gestor Administrativo"

'Validación para evitar los controles vacíos
If Me.ComboBox1.Text = "" Then
    Me.ComboBox1.BackColor = &HC0C0FF
    MsgBox "Seleccione el ÁREA", vbInformation, Titulo
    Me.ComboBox1.SetFocus
    Exit Sub
End If
If Me.ComboBox2.Text = "" Then
    Me.ComboBox2.BackColor = &HC0C0FF
    MsgBox "Seleccione una DESCRIPCIÓN", vbInformation, Titulo
    Me.ComboBox2.SetFocus
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

     Hoja13.Select
    Hoja13.Range("A2:H2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja13.Range("A3:H3").Select
    Selection.Copy
    Hoja13.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja13.Cells(2, 1) = CDate(Me.Text_fecha)
        Hoja13.Cells(2, 3) = Comprb
        Hoja13.Cells(2, 4) = Me.ComboBox1.Text
        Hoja13.Cells(2, 5) = Me.ComboBox2.Text
        Hoja13.Cells(2, 6) = Me.txt_monto.Text
        Hoja13.Cells(2, 7) = Me.txt_detalle.Value
        Hoja13.Cells(2, 8) = Hoja21.Cells(1, 7)
     
        LimpiarControles
        
        ComboBox1.SetFocus
        
        
       
Salir:
     If Err <> 0 Then
        MsgBox Err.Description, vbExclamation, Titulo
     End If
   Me.Text_fecha = Date
   Me.Label16.Caption = "No. " & Hoja22.Range("J2").Value + 1 'Llamamos el número de la factura
    
End Sub
Private Sub btn_Salir_Click()
Unload Me
End Sub
Private Sub ComboBox1_Change()

ComboBox1.BackColor = &H80000005

End Sub
Private Sub ComboBox2_Change()

ComboBox1.BackColor = &H80000005

End Sub

Private Sub ComboBox1_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox1.ListCount
    ComboBox1.RemoveItem 0
Next Fila

            For Fila = 2 To 3
                Lista = Hoja1.Cells(Fila, 51)
                ComboBox1.AddItem (Lista)
        Next
End Sub

                        
Private Sub ComboBox2_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox2.ListCount
    ComboBox2.RemoveItem 0
Next Fila

            For Fila = 2 To 9
                Lista = Hoja1.Cells(Fila, 50)
                ComboBox2.AddItem (Lista)
        Next
End Sub







