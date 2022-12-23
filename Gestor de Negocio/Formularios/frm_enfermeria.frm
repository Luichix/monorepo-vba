VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_enfermeria 
   Caption         =   "ENFERMERIA"
   ClientHeight    =   7550
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   5290
   OleObjectBlob   =   "frm_enfermeria.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_enfermeria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Registrar_Click()
 
 'Validación para evitar los controles vacíos
If Me.Text_fecha1.Text = Empty Or _
    Me.ComboBox1 = Empty Or _
    Me.ComboBox2 = Empty Or _
    Me.ComboBox3 = Empty Or _
    txt_medicamento.Text = Empty Or _
    txt_responsable.Text = Empty Or _
    txt_observaciones = Empty Then

            MsgBox "Hay campos vacíos en el registro", , "Gestor de Ganaderia"
            Exit Sub

End If
    
If MsgBox("¿Son Correctos los Datos?" + Chr(13) + "¿Desea Continuar?", vbYesNo, "Gestor de Ganaderia") = vbNo Then
        Exit Sub
    Else
        EnfermeriaAnimal
        
        MsgBox "Registro procesado con éxito!!!", , "Gestor de Ganaderia"
        
End If

End Sub

Private Sub btn_Salir_Click()
Unload Me
End Sub
Private Sub EnfermeriaAnimal()
Dim Comprb As Long

'Aquí manejo el correlativo del comprobante
Hoja22.Range("F2").Value = Hoja22.Range("F2").Value + 1
Comprb = Hoja22.Range("F2").Value


    Hoja32.Select
    Hoja32.Range("A2:H2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja32.Range("A3:H2").Select
    Selection.Copy
    Hoja32.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
            Hoja32.Cells(2, 1) = Comprb
            Hoja32.Cells(2, 2) = CDate(Me.Text_fecha)
            Hoja32.Cells(2, 3) = Me.ComboBox1
            Hoja32.Cells(2, 4) = Me.ComboBox2
            Hoja32.Cells(2, 5) = Me.ComboBox3
            Hoja32.Cells(2, 6) = Me.txt_medicamento
            Hoja32.Cells(2, 7) = Me.txt_responsable
            Hoja32.Cells(2, 8) = Me.txt_observaciones
      
        LimpiarControles
        txt_nombre.SetFocus
     
     
   Me.Text_fecha = Date
   Me.Label16.Caption = "No. " & Hoja22.Range("F2").Value + 1 'Llamamos el número de la factura
    
End Sub

Private Sub UserForm_Initialize()
 Me.Text_fecha = Date
 Me.Label16.Caption = "No  " & Hoja22.Range("F2").Value + 1 'Llamamos el número de la factura
End Sub
Private Sub btn_Fecha_Medic_Click()
banderaCalendario = 9
    Call LanzarCalendario(Me, "lbl_FechaSal")
End Sub
Private Sub LimpiarControles()
    Me.ComboBox1 = Empty
    Me.ComboBox2 = Empty
    Me.ComboBox3 = Empty
    txt_medicamento.Text = Empty
    txt_responsable.Text = Empty
    txt_observaciones = Empty
 End Sub
Private Sub ComboBox1_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox1.ListCount
    ComboBox1.RemoveItem 0
Next Fila
  
Final = GetUltimoR(Hoja29)

        For Fila = 2 To Final
                Lista = Hoja29.Cells(Fila, 4)
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
  
Final = GetUltimoR(Hoja29)

        For Fila = 2 To Final
           Lista = Hoja29.Cells(Fila, 5)
           ComboBox2.AddItem (Lista)
        Next
End Sub

Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long
Dim Registro As Integer

Me.ComboBox1.BackColor = &H80000005

If ComboBox1.Text = "" Then
    ComboBox2 = Empty
End If

    
Final = GetUltimoR(Hoja29)

    For Fila = 2 To Final
        If ComboBox1.Text = Hoja29.Cells(Fila, 4) Then
            Me.ComboBox2.Text = Hoja29.Cells(Fila, 5)
            Exit For
        
        End If
    Next
End Sub
Private Sub ComboBox2_Change()
Dim Fila As Long
Dim Final As Long
Dim Registro As Integer

Me.ComboBox2.BackColor = &H80000005

If ComboBox2.Text = "" Then
    ComboBox1 = Empty
End If

    
Final = GetUltimoR(Hoja29)

    For Fila = 2 To Final
        If ComboBox2.Text = Hoja29.Cells(Fila, 5) Then
            Me.ComboBox1.Text = Hoja29.Cells(Fila, 4)
            Exit For
        
        End If
    Next
End Sub
Private Sub ComboBox3_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox3.ListCount
    ComboBox3.RemoveItem 0
Next Fila

        For Fila = 2 To 5
           Lista = Hoja1.Cells(Fila, 38)
           ComboBox3.AddItem (Lista)
        Next
End Sub
