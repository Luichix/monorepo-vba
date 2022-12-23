VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_horas 
   Caption         =   "CONTROL DE ENTRADAS Y SALIDAS"
   ClientHeight    =   5580
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   6090
   OleObjectBlob   =   "frm_horas.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_horas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CommandButton1_Click()
Dim Titulo As String

On Error GoTo Salir

Application.ScreenUpdating = False
Titulo = "Gestion del Personal"

If Me.ComboBox1.Text = "" Then
    Me.ComboBox1.BackColor = &HC0C0FF
    MsgBox "Ingrese el código del personal", vbInformation, Titulo
    Me.ComboBox1.SetFocus
    Exit Sub
End If

        If Me.ComboBox2.Text = "" Then
            Me.ComboBox2.BackColor = &HC0C0FF
            MsgBox "Ingrese el nombre del personal", vbInformation, Titulo
            Me.ComboBox2.SetFocus
            Exit Sub
        End If

                        If Me.TextBox1.Text = "" Then
                            Me.TextBox1.BackColor = &HC0C0FF
                            MsgBox "Ingrese la fecha de asistencia", vbInformation, Titulo
                            Me.TextBox1.SetFocus
                            Exit Sub
                        End If

                                If Me.TextBox2.Text = "00:00" Or Me.TextBox2.Text = "" Then
                                    Me.TextBox2.BackColor = &HC0C0FF
                                    MsgBox "Ingrese la Hora de Ingreso", vbInformation, Titulo
                                    Me.TextBox2 = ""
                                    Me.TextBox2.SetFocus
                                    Exit Sub
                                End If
                                                                
                                        If Me.OptionButton1.Value = False Then
                                            If Me.OptionButton2.Value = False Then
                                            Me.OptionButton1.BackColor = &HC0C0FF
                                            Me.OptionButton2.BackColor = &HC0C0FF
                                            MsgBox "Seleccione el detalle horario", vbInformation, Titulo
                                            Exit Sub
                                            End If
                                        End If

If MsgBox("¿Son correctos los datos?", vbYesNo, Titulo) = vbNo Then
        Exit Sub
    Else

       Registrar_Asistencias
       LimpiarControles

       ComboBox1.SetFocus
    Me.Label16.Caption = "No. " & Hoja22.Range("N2").Value + 1 'Llamamos el número de la factura
End If
     Application.ScreenUpdating = True
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Recursos Humanos"
 End If

End Sub

Private Sub Registrar_Asistencias()
Dim Fecha As Date
Dim Comprb As Long

    
'Aquí manejo el correlativo del comprobante
Hoja22.Range("N2").Value = Hoja22.Range("N2").Value + 1
Comprb = Hoja22.Range("N2").Value

Fecha = Me.TextBox1.Text

Hoja33.Select
    Hoja33.Range("A2:F2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja33.Range("A3:F3").Select
    Selection.Copy
    Hoja33.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

        Hoja33.Cells(2, 1) = Format(Fecha, "MM/DD/YYYY")
        Hoja33.Cells(2, 2) = Me.ComboBox1.Text
        Hoja33.Cells(2, 5) = Me.TextBox2.Text
        Hoja33.Cells(2, 6) = Comprb
        Hoja33.Cells(2, 7) = Hoja21.Range("G1")

        If Me.OptionButton1.Value = True Then
            Hoja33.Cells(2, 4) = "ENTRADA"
        End If
         If Me.OptionButton2.Value = True Then
            Hoja33.Cells(2, 4) = "SALIDA"
        End If
        
Hoja33.Select
Hoja33.Cells(1, 1).Select

End Sub

Private Sub btn_Fecha_Horas_Click()
Me.TextBox1.BackColor = &H80000005
banderaCalendario = 19
  Call LanzarCalendario(Me, "TextBox1")
End Sub

Private Sub UserForm_Initialize()
Me.Label16.Caption = "No. " & Hoja22.Range("N2").Value + 1 'Llamamos el número de la factura
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

Final = GetUltimoR(Hoja5)

        For Fila = 2 To Final
            If Hoja5.Cells(Fila, 9) = "ACTIVO" Then
                Lista = Hoja5.Cells(Fila, 2)
                ComboBox2.AddItem (Lista)
            End If
        Next
End Sub

Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long
Dim Actividad As String


ComboBox1.BackColor = &H80000005
If ComboBox1.Text = "" Then
    ComboBox2.Text = ""
    
End If

Final = GetUltimoR(Hoja5)

    For Fila = 2 To Final
        If ComboBox1.Text = Hoja5.Cells(Fila, 1) Then
            Me.ComboBox2.Text = Hoja5.Cells(Fila, 2)
            Exit For
        End If
    Next


End Sub


Private Sub ComboBox2_Change()
Dim Fila As Long
Dim Final As Long

ComboBox2.BackColor = &H80000005
If ComboBox2.Text = "" Then
    ComboBox1.Text = ""
    
 End If

Final = GetUltimoR(Hoja5)
    For Fila = 2 To Final
        If ComboBox2.Text = Hoja5.Cells(Fila, 2) Then
            Me.ComboBox1.Text = Hoja5.Cells(Fila, 1)
           Exit For

        End If
    Next
End Sub

Private Sub textbox2_enter()
Me.TextBox2 = ""
End Sub

Private Sub OptionButton1_Click()
Me.OptionButton1.BackColor = &H80000005
Me.OptionButton2.BackColor = &H80000005

End Sub

Private Sub OptionButton2_Click()
Me.OptionButton1.BackColor = &H80000005
Me.OptionButton2.BackColor = &H80000005
End Sub

Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
If TextBox2.Text <> "1" And TextBox2.Text <> "2" And TextBox2.Text <> "3" And TextBox2.Text <> "4" And TextBox2.Text <> "0" Then
    
    Select Case Len(TextBox2.Value)
        Case 1
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 4
          End Select
        
    End If

If TextBox2.Value = 10 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If
If TextBox2.Value = 11 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If
If TextBox2.Value = 12 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If
If TextBox2.Value = 13 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If
If TextBox2.Value = 14 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If
If TextBox2.Value = 15 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If
If TextBox2.Value = 16 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If
If TextBox2.Value = 17 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If
If TextBox2.Value = 18 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If
If TextBox2.Value = 19 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If
If TextBox2.Value = 20 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If
If TextBox2.Value = 21 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If
If TextBox2.Value = 22 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If
If TextBox2.Value = 23 Then
    Select Case Len(TextBox2.Value)
        Case 2
        TextBox2.Value = TextBox2.Value & ":"
        Me.TextBox2.MaxLength = 5
        End Select

End If


If TextBox2.Value = 25 Or TextBox2.Value = 24 Or TextBox2.Value = 0 Or TextBox2.Value = 26 Or TextBox2.Value = 27 Or TextBox2.Value = 28 Or TextBox2.Value = 29 Or TextBox2.Value = 3 Or TextBox2.Value = 4 Then
    TextBox2 = "00:00"
     Me.TextBox2.MaxLength = 4
End If

End Sub
Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub LimpiarControles()
Me.ComboBox1 = ""
Me.ComboBox2 = ""
Me.TextBox1.BackColor = &H80000018
Me.TextBox2 = "00:00"
Me.OptionButton1.BackColor = &H80000018
Me.OptionButton2.BackColor = &H80000018
Me.TextBox2.BackColor = &H80000005

End Sub

