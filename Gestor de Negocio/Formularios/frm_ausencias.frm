VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ausencias 
   Caption         =   "REGISTROS DE AUSENCIAS"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5660
   OleObjectBlob   =   "frm_ausencias.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_ausencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub btn_Fecha_Final_Click()
Me.TextBox2.BackColor = &H80000005
banderaCalendario = 18
  Call LanzarCalendario(Me, "TextBox2")
End Sub

Private Sub btn_Fecha_Inicio_Click()
Me.TextBox1.BackColor = &H80000005

banderaCalendario = 17
  Call LanzarCalendario(Me, "TextBox1")
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
    ComboBox3.Text = ""
    
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
    ComboBox3.Text = ""
 End If

Final = GetUltimoR(Hoja5)
    For Fila = 2 To Final
        If ComboBox2.Text = Hoja5.Cells(Fila, 2) Then
            Me.ComboBox1.Text = Hoja5.Cells(Fila, 1)
           Exit For

        End If
    Next
End Sub


Private Sub ComboBox3_Change()
ComboBox3.BackColor = &H80000005
End Sub

Private Sub ComboBox3_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String

For Fila = 1 To ComboBox3.ListCount
    ComboBox3.RemoveItem 0
Next Fila

        For Fila = 2 To 6
                Lista = Hoja1.Cells(Fila, 67)
                ComboBox3.AddItem (Lista)
           
        Next
End Sub


Private Sub ComboBox4_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String

For Fila = 1 To ComboBox4.ListCount
    ComboBox4.RemoveItem 0
Next Fila

        For Fila = 2 To 3
                Lista = Hoja1.Cells(Fila, 66)
                ComboBox4.AddItem (Lista)
           
        Next
End Sub


Private Sub CommandButton2_Click()
Unload Me
End Sub
Private Sub CommandButton3_Click()
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
        
                If Me.ComboBox3.Text = "" Then
                    Me.ComboBox3.BackColor = &HC0C0FF
                    MsgBox "Ingrese el motivo", vbInformation, Titulo
                    Me.ComboBox3.SetFocus
                    Exit Sub
                End If
                        If Me.TextBox1.Text = "" Then
                            Me.TextBox1.BackColor = &HC0C0FF
                            MsgBox "Ingrese la Fecha Inicial", vbInformation, Titulo
                            Me.TextBox1.SetFocus
                            Exit Sub
                        End If
                        
                                If Me.TextBox2.Text = "" Then
                                    Me.TextBox2.BackColor = &HC0C0FF
                                    MsgBox "Ingrese la Fecha Final", vbInformation, Titulo
                                    Me.TextBox2.SetFocus
                                    Exit Sub
                                End If
                                
                                        If Me.OptionButton1.Value = False Then
                                            If Me.OptionButton2.Value = False Then
                                            Me.OptionButton1.BackColor = &HC0C0FF
                                            Me.OptionButton2.BackColor = &HC0C0FF
                                            MsgBox "Seleccione un Periodo de Quincena", vbInformation, Titulo
                                            Exit Sub
                                            End If
                                        End If

  
       Registrar_Ausencias
       LimpiarControles
        
       ComboBox1.SetFocus
        
 Me.Label16.Caption = "No. " & Hoja22.Range("M2").Value + 1 'Llamamos el número de la factura

     Application.ScreenUpdating = True
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Recursos Humanos"
 End If


End Sub

Private Sub Registrar_Ausencias()
Dim Comprb As Long
Dim Fecha1 As Date
Dim Fecha2 As Date
    
'Aquí manejo el correlativo del comprobante
Hoja22.Range("M2").Value = Hoja22.Range("M2").Value + 1
Comprb = Hoja22.Range("M2").Value

Fecha1 = Me.TextBox1.Text
Fecha2 = Me.TextBox2.Text

Hoja17.Select
    Hoja17.Range("A2:k2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja17.Range("A3:k3").Select
    Selection.Copy
    Hoja17.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
        Hoja17.Cells(2, 1) = Date
        Hoja17.Cells(2, 2) = Me.ComboBox1.Text
        Hoja17.Cells(2, 6) = Me.ComboBox3.Text
        Hoja17.Cells(2, 10) = Comprb
        Hoja17.Cells(2, 11) = Hoja21.Range("G1")
        
 
        Hoja17.Cells(2, 4) = Format(Fecha1, "MM/DD/YYYY")
   
        Hoja17.Cells(2, 5) = Format(Fecha2, "MM/DD/YYYY")
        
        If Me.OptionButton1.Value = True Then
            Hoja17.Cells(2, 8) = "PRIMERA"
        End If
         If Me.OptionButton2.Value = True Then
            Hoja17.Cells(2, 8) = "SEGUNDA"
        End If
        
        
      
End Sub

Private Sub OptionButton1_Click()
Me.OptionButton1.BackColor = &H80000005
Me.OptionButton2.BackColor = &H80000005

End Sub

Private Sub OptionButton2_Click()
Me.OptionButton1.BackColor = &H80000005
Me.OptionButton2.BackColor = &H80000005
End Sub


Private Sub UserForm_Initialize()
Me.Label16.Caption = "No. " & Hoja22.Range("M2").Value + 1 'Llamamos el número de la factura
End Sub

Private Sub LimpiarControles()
Me.ComboBox1 = ""
Me.ComboBox2 = ""
Me.ComboBox3 = ""
Me.TextBox1 = ""
Me.TextBox2 = ""
Me.OptionButton1.Value = False
Me.OptionButton2.Value = False

End Sub

