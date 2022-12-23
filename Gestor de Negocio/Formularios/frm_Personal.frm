VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Personal 
   Caption         =   "GESTIÓN DE PERSONAL"
   ClientHeight    =   7520
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   5710
   OleObjectBlob   =   "frm_Personal.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Personal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btn_Registrar_Click()
 Dim Titulo As String
 
If Me.Text_fecha.Text = Empty Or _
    Me.txt_nombre = Empty Or _
    Me.txt_cedula = Empty Or _
    Me.txt_telefono = Empty Or _
    Me.ComboBox1 = Empty Or _
    Me.ComboBox2 = Empty Or _
    Me.ComboBox3 = Empty Then
    
           
            MsgBox "Hay campos vacíos en el registro", , "Gestor Administrativo"
            Exit Sub
    
End If
    

Titulo = "Gestión de Personal"
Application.ScreenUpdating = False

If MsgBox("¿Son Correctos los Datos?" + Chr(13) + "¿Desea Continuar?", vbYesNo, "Gestor Administrativo") = vbNo Then
        Exit Sub
    Else
        RegistrarPersonal
        
        MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
        
End If
Application.ScreenUpdating = True
End Sub
Private Sub btn_Salir_Click()
Unload Me
End Sub
Private Sub RegistrarPersonal()
Dim Comprb As Long




'Aquí manejo el correlativo del comprobante
Hoja22.Range("G2").Value = Hoja22.Range("G2").Value + 1
Comprb = Hoja22.Range("G2").Value

   Hoja5.Select
   ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort. _
        SortFields.Add Key:=Range("IDPERSONAL[CODIGO DE EMPLEADO]"), SortOn:= _
        xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
           
    Hoja5.Select
    Hoja5.Range("A2:I2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja5.Range("A3:I2").Select
    Selection.Copy
    Hoja5.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
            Hoja5.Cells(2, 1) = "F0" & Comprb
            Hoja5.Cells(2, 8) = CDate(Me.Text_fecha)
            Hoja5.Cells(2, 2) = txt_apellido & " " & txt_nombre
            Hoja5.Cells(2, 3) = txt_cedula
            Hoja5.Cells(2, 4) = txt_telefono
            Hoja5.Cells(2, 5) = ComboBox1
            Hoja5.Cells(2, 6) = ComboBox2
            Hoja5.Cells(2, 7) = ComboBox3.Value
            Hoja5.Cells(2, 9) = "ACTIVO"
         
   ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort. _
        SortFields.Add Key:=Range("IDPERSONAL[CODIGO DE EMPLEADO]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Ajustar Planilla
    
       Hoja16.Select
       Range("A3").Select

       Selection.ListObject.ListRows.Add (1)
                    Hoja16.Range("A5:V5").Select
                    Selection.Copy
                    Hoja16.Range("A4").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                Hoja16.Cells(4, 1) = "F0" & Comprb
                
    ActiveWorkbook.Worksheets("PLANILLA").ListObjects("Tabla912").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("PLANILLA").ListObjects("Tabla912").Sort.SortFields. _
        Add Key:=Range("A4"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("PLANILLA").ListObjects("Tabla912").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
                
     '''''''''''''''''''''''''''''''''
         'Ajustar Horas
    
       Hoja18.Select
       Range("A4").Select

       Selection.ListObject.ListRows.Add (1)
                    Hoja18.Range("A6:DH6").Select
                    Selection.Copy
                    Hoja18.Range("A5").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                Hoja18.Cells(5, 1) = "F0" & Comprb
                
        ActiveWorkbook.Worksheets("HORAS").ListObjects("tbl_Horarios154").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("HORAS").ListObjects("tbl_Horarios154").Sort. _
        SortFields.Add Key:=Range("A5"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("HORAS").ListObjects("tbl_Horarios154").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
     '''''''''''''''''''''''''''''''''
                
        Hoja5.Select
        LimpiarControles
        txt_nombre.SetFocus
     
     
   Me.Text_fecha = Date
   Me.Label16.Caption = "No. " & "F0" & Hoja22.Range("G2").Value + 1 'Llamamos el número de la factura
    
End Sub

Private Sub CheckBox4_Click()
    If CheckBox4 = True Then
        CheckBox2 = False
        CheckBox3 = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
    End If
    If CheckBox4 = False Then
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
    End If
       
End Sub

Private Sub CheckBox5_Click()
   If CheckBox5 = False Then
        CheckBox2 = False
        CheckBox3 = False
        CheckBox3.Enabled = False
        CheckBox2.Enabled = False
        ComboBox7_Change
    End If
    If CheckBox5 = True Then
        CheckBox2 = False
        CheckBox3 = False
        CheckBox3.Enabled = True
        CheckBox2.Enabled = True
        Me.Text_Fecha_Acc = Empty
        
    End If
    
    

End Sub

Private Sub ComboBox4_Change()
Dim Fila As Long
Dim Final As Long
Dim Columna As String
Dim Lista As String
Dim IndiceCat As Integer
Dim j As Integer

IndiceCat = ComboBox4.ListIndex + 57

    Fila = Hoja1.Cells(Rows.Count, IndiceCat).End(xlUp).Row
ComboBox5 = Empty
ComboBox5.Clear
    For j = 2 To Fila
        ComboBox5.AddItem (Hoja1.Cells(j, IndiceCat))
    Next j

End Sub

Private Sub UserForm_Initialize()
Dim Fila As Long
Dim Final As Long
Dim Columna As String
Dim Lista As String

 Me.Text_fecha = Date
 Me.Label16.Caption = "No  " & "F0" & Hoja22.Range("G2").Value + 1 'Llamamos el número de la factura
 Me.ComboBox4.Enabled = False
 Me.ComboBox5.Enabled = False
 Me.ComboBox6.Enabled = False
 Me.CheckBox5.Enabled = False
 

For Final = 1 To ComboBox1.ListCount
   ComboBox1.RemoveItem 0
Next Final
  
        For Final = 57 To 64
                Columna = Hoja1.Cells(1, Final)
                ComboBox1.AddItem (Columna)
        Next
      
End Sub
Private Sub btn_Fecha_Contra_Click()
banderaCalendario = 10
    Call LanzarCalendario(Me, "lbl_FechaSal")
End Sub
Private Sub btn_Fecha_Acc_Click()
banderaCalendario = 16
    Call LanzarCalendario(Me, "lbl_FechaSal")
End Sub
Private Sub LimpiarControles()
    Text_fecha = Empty
    txt_nombre = Empty
    txt_cedula = Empty
    Text_fecha = Empty
    ComboBox1 = Empty
    ComboBox2 = Empty
    ComboBox3 = Empty
End Sub

Private Sub ComboBox1_Enter()

End Sub

Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long
Dim Columna As String
Dim Lista As String
Dim IndiceCat As Integer
Dim j As Integer

IndiceCat = ComboBox1.ListIndex + 57

    Fila = Hoja1.Cells(Rows.Count, IndiceCat).End(xlUp).Row
ComboBox2.Clear
    For j = 2 To Fila
        ComboBox2.AddItem (Hoja1.Cells(j, IndiceCat))
    Next j

End Sub
Private Sub ComboBox3_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox3.ListCount
    ComboBox3.RemoveItem 0
Next Fila

        For Fila = 2 To 5
           Lista = Hoja1.Cells(Fila, 46)
           ComboBox3.AddItem (Lista)
        Next
End Sub
Private Sub txt_cedula_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

Select Case Len(txt_cedula.Text)
Case 3
txt_cedula.Text = txt_cedula.Text & "-"
Case 10
txt_cedula.Text = txt_cedula.Text & "-"

End Select

End Sub
Private Sub txt_telefono_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

Select Case Len(txt_telefono.Text)
Case 4
txt_telefono.Text = txt_telefono.Text & "-"

End Select

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Salir_Accion_Click()
Unload Me
End Sub
Private Sub ComboBox7_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String

For Fila = 1 To ComboBox7.ListCount
    ComboBox7.RemoveItem 0
Next Fila

Final = GetUltimoR(Hoja5)

        For Fila = 2 To Final
            If Hoja5.Cells(Fila, 9) = "ACTIVO" Then
                Lista = Hoja5.Cells(Fila, 1)
                ComboBox7.AddItem (Lista)
            End If
        Next
End Sub

Private Sub ComboBox7_Change()
Dim Fila As Long
Dim Final As Long
Dim Actividad As String


ComboBox7.BackColor = &H80000005
If ComboBox7.Text = "" Then
    ComboBox8.Text = ""
    ComboBox9.Text = ""
    ComboBox4.Text = ""
    ComboBox5.Text = ""
    ComboBox6.Text = ""
    CheckBox2 = False
    CheckBox3 = False
    CheckBox4 = False
    CheckBox5 = False
    Me.Text_Fecha_Acc = ""
    Me.Text_Fecha_Acc.BackColor = &H80000005
    
End If

Final = GetUltimoR(Hoja5)

    For Fila = 2 To Final
        If ComboBox7.Text = Hoja5.Cells(Fila, 1) Then
            Me.ComboBox8.Text = Hoja5.Cells(Fila, 2)
            Me.ComboBox9.Text = Hoja5.Cells(Fila, 9)
            Me.ComboBox4.Text = Hoja5.Cells(Fila, 5)
            Me.ComboBox5.Text = Hoja5.Cells(Fila, 6)
            Me.ComboBox6.Text = Hoja5.Cells(Fila, 7)
            Me.Text_Fecha_Acc = Hoja5.Cells(Fila, 8)
            Actividad = Hoja5.Cells(Fila, 9)
            Exit For
        End If
    Next
    
    If Actividad = "INACTIVO" Then
        CheckBox5.Enabled = True
        CheckBox2 = False
        CheckBox2.Enabled = False
        CheckBox3 = False
        CheckBox3.Enabled = False
        CheckBox4 = False
        CheckBox4.Enabled = False
        Me.Text_Fecha_Acc.BackColor = &H80000018
    End If
    If Actividad = "ACTIVO" Then
        CheckBox5.Enabled = False
        CheckBox4.Enabled = True
        CheckBox3.Enabled = True
        CheckBox2.Enabled = True
        Me.Text_Fecha_Acc.BackColor = &H80000005
        
    End If
    
    
End Sub

Private Sub ComboBox8_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox8.ListCount
    ComboBox8.RemoveItem 0
Next Fila

Final = GetUltimoR(Hoja5)

        For Fila = 2 To Final
'            If Hoja5.Cells(Fila, 9) = "ACTIVO" Then
                Lista = Hoja5.Cells(Fila, 2)
                ComboBox8.AddItem (Lista)
'            End If
        Next
End Sub

Private Sub ComboBox8_Change()
Dim Fila As Long
Dim Final As Long

ComboBox8.BackColor = &H80000005
If ComboBox8.Text = "" Then
    ComboBox7.Text = ""
    ComboBox9.Text = ""
End If

    For Fila = 2 To 1000
        If ComboBox8.Text = Hoja5.Cells(Fila, 2) Then
            Me.ComboBox7.Text = Hoja5.Cells(Fila, 1)
            Me.ComboBox9.Text = Hoja5.Cells(Fila, 9)
            Exit For

        End If
    Next
End Sub

Private Sub ComboBox6_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox6.ListCount
    ComboBox6.RemoveItem 0
Next Fila

        For Fila = 2 To 5
           Lista = Hoja1.Cells(Fila, 46)
           ComboBox6.AddItem (Lista)
        Next
End Sub

Private Sub CheckBox2_Click()
    If Me.CheckBox2.Value = True Then
        Me.ComboBox6.Enabled = True
        Me.ComboBox6.SetFocus
    End If
    If Me.CheckBox2.Value = False Then
        ComboBox7_Change
        Me.ComboBox6.Enabled = False
        Me.CheckBox5 = False
    End If
End Sub
Private Sub CheckBox3_Click()
 Dim Final As Long
 Dim Fila As Long
 Dim Columna As String
 Dim Lista As String
 
     
For Final = 1 To ComboBox4.ListCount
   ComboBox4.RemoveItem 0
Next Final
  
        For Final = 57 To 64
                Columna = Hoja1.Cells(1, Final)
                ComboBox4.AddItem (Columna)
        Next
      
    If Me.CheckBox3.Value = True Then
        Me.ComboBox4.Enabled = True
        Me.ComboBox5.Enabled = True
        Me.ComboBox4.SetFocus
    End If
     If Me.CheckBox3.Value = False Then
        ComboBox7_Change
        Me.ComboBox4.Enabled = False
        Me.ComboBox5.Enabled = False
        Me.CheckBox5 = False
    End If
End Sub

Private Sub Registrar_Accion_Click()
 If Me.OptionButton1.Value = True Then
            Hoja17.Cells(2, 8) = "PRIMERA"
        End If
        Dim x As String
Dim encontrado As Boolean
Dim Titulo As String
Dim Fecha As String

Application.ScreenUpdating = False
Titulo = "Gestion del Personal"

If Me.ComboBox7.Text = "" Then
    Me.ComboBox7.BackColor = &HC0C0FF
    MsgBox "Ingrese el código del personal", vbInformation, Titulo
    Me.ComboBox7.SetFocus
    Exit Sub
End If

If Me.ComboBox8.Text = "" Then
    Me.ComboBox8.BackColor = &HC0C0FF
    MsgBox "Ingrese el nombre del personal", vbInformation, Titulo
    Me.ComboBox8.SetFocus
    Exit Sub
End If
If Me.Text_Fecha_Acc.Text = "" Then
    Me.Text_Fecha_Acc.BackColor = &HC0C0FF
    MsgBox "Ingrese la fecha de ingreso del personal", vbInformation, Titulo
    Me.Text_Fecha_Acc.SetFocus
    Exit Sub
End If
       
x = ComboBox7.Text

Hoja5.Select
Range("A1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like x Then
            encontrado = True
            Exit Do
                                 
        End If
    Loop
    If encontrado = False Then
        MsgBox "Personal no Existente", vbInformation, Titulo
    End If
    
    Fecha = Me.Text_Fecha_Acc.Text
    
           ActiveCell.Offset(0, 4) = ComboBox4.Text
           ActiveCell.Offset(0, 5) = ComboBox5.Text
           ActiveCell.Offset(0, 6) = ComboBox6.Text
           ActiveCell.Offset(0, 7) = Fecha
           ActiveCell.Offset(0, 7) = Format(Fecha, "MM/DD/YYYY")
           
           
           If CheckBox4 = True Then
                ActiveCell.Offset(0, 8) = "INACTIVO"
           End If
           
           If Me.CheckBox5 = True Then
                ActiveCell.Offset(0, 8) = "ACTIVO"
            End If
        
  Ajustar_planilla
  Ajustar_horas
        Hoja5.Select
        Range("A1").Select
        MsgBox "Ajustes Realizados Correctamente..!", vbInformation, Titulo
        Limpiar_Acciones
        ComboBox7.SetFocus
    Application.ScreenUpdating = True
      
End Sub
Sub Ajustar_planilla()
Dim Final As String
Dim x As String
Dim encontrado As Boolean
Dim Titulo As String

x = ComboBox7.Text

Hoja16.Select
Range("A3").Select
Titulo = "Gestion del Personal"

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like x Then
            encontrado = True
            Exit Do

        End If
    Loop
    If encontrado = False Then
        MsgBox "Personal Agregado en Planilla", vbInformation, Titulo

    End If
           
           If CheckBox4 = True Then
                ActiveCell.EntireRow.Delete
           End If
            If Me.CheckBox5 = True Then
                Selection.ListObject.ListRows.Add (1)
                    Hoja16.Range("A5:V5").Select
                    Selection.Copy
                    Hoja16.Range("A4").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                Hoja16.Cells(4, 1) = x
            End If

        

End Sub
Sub Ajustar_horas()
Dim Final As String
Dim x As String
Dim encontrado As Boolean
Dim Titulo As String

x = ComboBox7.Text

Hoja18.Select
Range("A4").Select
Titulo = "Gestion del Personal"

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like x Then
            encontrado = True
            Exit Do

        End If
    Loop
    If encontrado = False Then
        MsgBox "Personal Agregado en Control de Horas", vbInformation, Titulo

    End If
           
           If CheckBox4 = True Then
                ActiveCell.EntireRow.Delete
           End If
            If Me.CheckBox5 = True Then
                Selection.ListObject.ListRows.Add (1)
                    Hoja18.Range("A6:DH6").Select
                    Selection.Copy
                    Hoja18.Range("A5").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                Hoja18.Cells(5, 1) = x
            End If

        

End Sub

Private Sub Limpiar_Acciones()

    Me.ComboBox7 = Empty
    Me.CheckBox2 = False
    Me.CheckBox3 = False
    Me.CheckBox4 = False
    Me.CheckBox5 = False
    Me.Text_Fecha_Acc.BackColor = &H80000005
    

End Sub

