VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Personal 
   Caption         =   "GESTIÓN DE PERSONAL"
   ClientHeight    =   10665
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   11050
   OleObjectBlob   =   "frm_Personal.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Personal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_Apersonal_Click()

frm_Base_Personal.Show
End Sub

Private Sub btn_personal_Click()

frm_Base_Personal.Show
End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub

Private Sub btn_Asalir_Click()
Unload Me
End Sub



Private Sub txt_Asalario_Change()

End Sub

Private Sub txt_Fin_Change()

End Sub

Private Sub txt_Id_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_id, KeyAscii)
End Sub

Private Sub txt_Salario_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_Salario, KeyAscii)
End Sub

Private Sub UserForm_Initialize()
On Error Resume Next

 Me.txt_Acontrato.Enabled = False
 Me.txt_Ainicio.Enabled = False
 Me.txt_Afin.Enabled = False
 Me.txt_Aarea.Enabled = False
 Me.txt_Apuesto.Enabled = False
 Me.txt_Asalario.Enabled = False
 Me.txt_Aregimen.Enabled = False
 Me.txt_Ajornada.Enabled = False
 Me.txt_Aregimen.Enabled = False
 Me.txt_Ajornada.Enabled = False
 Me.opt_Asi.Enabled = False
 Me.opt_Ano.Enabled = False
 Me.txt_APago.Enabled = False
 Me.txt_Abancaria.Enabled = False

 Me.ck_Contrato.Enabled = False
 Me.ck_Traslado.Enabled = False
 Me.ck_Aumento.Enabled = False
 Me.ck_Baja.Enabled = False
 Me.ck_Reintegro.Enabled = False
 Me.ck_Otros.Enabled = False
 
 Me.btn_Acontrato.Enabled = False
 Me.btn_Afin.Enabled = False
 Me.btn_Ainicio.Enabled = False
 Me.btn_Aregimen.Enabled = False
 Me.btn_Ajornada.Enabled = False
 Me.btn_Apago.Enabled = False
 
End Sub
Private Sub btn_Contrato_Click()
banderaContrato = 1
    Call LanzarContrato(Me, "lbl_Fecha3")
End Sub
Private Sub btn_Inicio_Click()
banderaCalendario = 1
    Call LanzarCalendario(Me, "lbl_Fecha2")
End Sub
Private Sub btn_Fin_Click()
banderaCalendario = 2
    Call LanzarCalendario(Me, "lbl_Fecha3")
End Sub
Private Sub btn_Regimen_Click()
banderaRegimen = 1
    Call LanzarRegimen(Me, "lbl_fecha3")
End Sub

Private Sub btn_Jornada_Click()
banderaJornada = 1
    Call LanzarJornada(Me, "lbl_fecha3")
End Sub
Private Sub btn_Pago_Click()
banderaPago = 1
    Call LanzarPago(Me, "lbl_fecha3")
End Sub
Private Sub btn_listadopersonal_Click()
banderaPersonal = 4
Call LanzarListadoPersonal(Me, "btn_listadopersonal")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub txt_Aid_Change()
Dim Fila As Long
Dim Final As Long
Dim Actividad As String


If Me.txt_Aid.Text = Empty Then
    Me.txt_Anombre.Text = Empty
    Me.txt_Aestado.Text = Empty
    Me.txt_Acedula.Text = Empty
    Me.txt_Atelefono.Text = Empty
    Me.txt_Acontrato.Text = Empty
    Me.txt_Ainicio.Text = Empty
    Me.txt_Afin.Text = Empty
    Me.txt_Aarea.Text = Empty
    Me.txt_Apuesto.Text = Empty
    Me.txt_Asalario.Text = Empty
    Me.txt_Aregimen.Text = Empty
    Me.txt_Ajornada.Text = Empty
    Me.txt_APago.Text = Empty
    Me.txt_Abancaria.Text = Empty
    Me.opt_Ano.Value = False
    Me.opt_Asi.Value = False
    Me.ck_Contrato.Value = False
    Me.ck_Traslado.Value = False
    Me.ck_Aumento.Value = False
    Me.ck_Baja.Value = False
    Me.ck_Reintegro.Value = False
    Me.ck_Otros.Value = False
End If

Final = GetUltimoR(Hoja1)

    For Fila = 2 To Final
        If Me.txt_Aid.Text = Hoja1.Cells(Fila, 1) Then
            Me.txt_Anombre.Text = Hoja1.Cells(Fila, 2)
            Me.txt_Acedula.Text = Hoja1.Cells(Fila, 3)
            Me.txt_Atelefono.Text = Hoja1.Cells(Fila, 4)
            Me.txt_Aarea.Text = Hoja1.Cells(Fila, 5)
            Me.txt_Apuesto.Text = Hoja1.Cells(Fila, 6)
            Me.txt_Aregimen.Text = Hoja1.Cells(Fila, 7)
            Me.txt_Ajornada.Text = Hoja1.Cells(Fila, 8)
            Me.txt_Asalario.Value = Hoja1.Cells(Fila, 10)
            Me.txt_Acontrato.Text = Hoja1.Cells(Fila, 11)
            Me.txt_Ainicio.Text = Hoja1.Cells(Fila, 12)
            Me.txt_Afin.Text = Hoja1.Cells(Fila, 13)
            Me.txt_APago.Text = Hoja1.Cells(Fila, 14)
            Me.txt_Abancaria.Text = Hoja1.Cells(Fila, 15)
            Me.txt_Aestado.Text = Hoja1.Cells(Fila, 16)
            
            If Hoja1.Cells(Fila, 9) = "SI" Then
                Me.opt_Asi.Value = True
                Me.opt_Ano.Value = False
            ElseIf Hoja1.Cells(Fila, 9) = "NO" Then
                Me.opt_Asi.Value = False
                Me.opt_Ano.Value = True
            End If
            
            Exit For
        End If
    Next
    
    Actividad = Me.txt_Aestado.Text
    
    If Actividad = "INACTIVO" Then
        
        Me.ck_Baja.Enabled = False
        Me.ck_Reintegro.Enabled = True
        
        If Me.ck_Reintegro.Value = False Then
        Me.ck_Contrato.Enabled = False
        Me.ck_Traslado.Enabled = False
        Me.ck_Aumento.Enabled = False
        Me.ck_Otros.Enabled = False
        End If
        
    End If

    If Actividad = "ACTIVO" Then
        Me.ck_Contrato.Enabled = True
        Me.ck_Traslado.Enabled = True
        Me.ck_Aumento.Enabled = True
        Me.ck_Baja.Enabled = True
        Me.ck_Reintegro.Enabled = False
        Me.ck_Otros.Enabled = True
    End If
End Sub
Private Sub ck_Contrato_Click()
    If Me.ck_Contrato.Value = True Then
        Me.txt_Acontrato.Enabled = True
        Me.txt_Ainicio.Enabled = True
        Me.txt_Afin.Enabled = True
        Me.btn_Acontrato.Enabled = True
        Me.btn_Ainicio.Enabled = True
        Me.btn_Afin.Enabled = True
        Me.txt_Afin.Enabled = True
        Me.txt_Afin.Locked = False
        Me.txt_Ainicio.SetFocus
    End If
    If Me.ck_Contrato.Value = False Then
        txt_Aid_Change
        Me.txt_Acontrato.Enabled = False
        Me.txt_Ainicio.Enabled = False
        Me.txt_Afin.Enabled = False
        Me.btn_Acontrato.Enabled = True
        Me.btn_Ainicio.Enabled = True
        Me.btn_Afin.Enabled = True
        Me.txt_Afin.Locked = True
    End If
    
End Sub
Private Sub ck_Traslado_Click()
    If Me.ck_Traslado.Value = True Then
        Me.txt_Aarea.Enabled = True
        Me.txt_Apuesto.Enabled = True
        Me.txt_Aarea.SetFocus
    End If
    If Me.ck_Traslado.Value = False Then
        txt_Aid_Change
        Me.txt_Aarea.Enabled = False
        Me.txt_Apuesto.Enabled = False
    End If
    
End Sub

Private Sub ck_Aumento_Click()
    If Me.ck_Aumento.Value = True Then
        Me.txt_Asalario.Enabled = True
        Me.txt_Asalario.SetFocus
    End If
    If Me.ck_Aumento.Value = False Then
        txt_Aid_Change
        Me.txt_Asalario.Enabled = False
    End If
End Sub
Private Sub ck_Otros_Click()
    If Me.ck_Otros.Value = True Then
        Me.txt_Aregimen.Enabled = True
        Me.txt_Ajornada.Enabled = True
        Me.txt_APago.Enabled = True
        Me.txt_Abancaria.Enabled = True
        Me.opt_Ano.Enabled = True
        Me.opt_Asi.Enabled = True
        Me.btn_Aregimen.Enabled = True
        Me.btn_Ajornada.Enabled = True
        Me.btn_Apago.Enabled = True
    End If
    If Me.ck_Otros.Value = False Then
        txt_Aid_Change
        Me.txt_Aregimen.Enabled = False
        Me.txt_Ajornada.Enabled = False
        Me.txt_APago.Enabled = False
        Me.txt_Abancaria.Enabled = False
        Me.opt_Ano.Enabled = False
        Me.opt_Asi.Enabled = False
        Me.btn_Aregimen.Enabled = False
        Me.btn_Ajornada.Enabled = False
        Me.btn_Apago.Enabled = False
    End If
    
End Sub
Private Sub ck_Baja_Click()
    If Me.ck_Baja.Value = True Then
        Me.ck_Contrato.Enabled = False
        Me.ck_Aumento.Enabled = False
        Me.ck_Traslado.Enabled = False
        Me.ck_Otros.Enabled = False
        
        Me.btn_Afin.Enabled = True
        Me.txt_Afin.Enabled = True
        Me.txt_Afin.SetFocus
    End If
    If Me.ck_Baja.Value = False Then
        txt_Aid_Change
        Me.ck_Contrato.Enabled = True
        Me.ck_Aumento.Enabled = True
        Me.ck_Traslado.Enabled = True
        Me.ck_Otros.Enabled = True
        
        Me.btn_Afin.Enabled = False
        Me.txt_Afin.Enabled = False
    End If
    
End Sub
Private Sub ck_Reintegro_Click()
    If Me.ck_Reintegro.Value = True Then
        Me.ck_Contrato.Enabled = True
        Me.ck_Aumento.Enabled = True
        Me.ck_Traslado.Enabled = True
        Me.ck_Otros.Enabled = True
        
        Me.btn_Ainicio.Enabled = True
        Me.txt_Ainicio.Enabled = True
       
        Me.btn_Afin.Enabled = True
        Me.txt_Fin.Enabled = True
        
        Me.txt_Ainicio.SetFocus
        
        
    End If
    If Me.ck_Reintegro.Value = False Then
        txt_Aid_Change
        Me.ck_Contrato.Enabled = False
        Me.ck_Aumento.Enabled = False
        Me.ck_Traslado.Enabled = False
        Me.ck_Otros.Enabled = False
        
        Me.btn_Ainicio.Enabled = False
        Me.txt_Ainicio.Enabled = False
        Me.btn_Afin.Enabled = False
        Me.txt_Afin.Enabled = False
    End If
    
End Sub

Private Sub btn_Acontrato_Click()
banderaContrato = 2
    Call LanzarContrato(Me, "lbl_Fecha3")
End Sub
Private Sub btn_Ainicio_Click()
banderaCalendario = 28
    Call LanzarCalendario(Me, "lbl_Fecha2")
End Sub
Private Sub btn_Afin_Click()
banderaCalendario = 29
    Call LanzarCalendario(Me, "lbl_Fecha3")
End Sub
Private Sub btn_Aregimen_Click()
banderaRegimen = 2
    Call LanzarRegimen(Me, "lbl_fecha3")
End Sub

Private Sub btn_Ajornada_Click()
banderaJornada = 2
    Call LanzarJornada(Me, "lbl_fecha3")
End Sub
Private Sub btn_Apago_Click()
banderaPago = 2
    Call LanzarPago(Me, "lbl_fecha3")
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LimpiarControles()
    Me.txt_id = Empty
    Me.txt_Apellido = Empty
    Me.txt_nombre = Empty
    Me.txt_Cedula = Empty
    Me.txt_Telefono = Empty
    Me.txt_Contrato = Empty
    Me.txt_Inicio = Empty
    Me.txt_Fin = Empty
    Me.txt_Area = Empty
    Me.txt_Puesto = Empty
    Me.txt_Salario = Empty
    Me.txt_Regimen = Empty
    Me.txt_Jornada = Empty
    Me.txt_Pago = Empty
    
    Me.opt_Si = False
    Me.opt_No = False
    
End Sub
Private Sub btn_Registrar_Click()
Dim Titulo As String

Titulo = "Gestión de Personal"

If Me.txt_id.Text = Empty Or _
    Me.txt_nombre = Empty Or _
    Me.txt_Apellido = Empty Or _
    Me.txt_Cedula = Empty Or _
    Me.txt_Telefono = Empty Or _
    Me.txt_Contrato = Empty Or _
    Me.txt_Inicio = Empty Or _
    Me.txt_Area = Empty Or _
    Me.txt_Puesto = Empty Or _
    Me.txt_Salario = Empty Or _
    Me.txt_Regimen = Empty Or _
    Me.txt_Jornada = Empty Or _
    Me.opt_Si = False And Me.opt_No = False Or _
    Me.txt_Pago = Empty Then
        
            MsgBox "Hay campos vacíos en el registro..!", vbInformation, Titulo
            Exit Sub
    
End If

If MsgBox("¿Son Correctos los Datos?" + Chr(13) + "¿Desea Continuar?", vbYesNo, Titulo) = vbNo Then
        Exit Sub
    Else
    
           Verificador
      
End If

End Sub
Private Sub Verificador()
Dim X As String
Dim encontrado As Boolean
Dim Seguridad As String

Seguridad = Hoja83.Range("L1").Text

X = Me.txt_id

Hoja1.Select
Range("A1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like X Then
            encontrado = True
            Exit Do
                                 
        End If
    Loop
    If encontrado = True Then
        MsgBox "ID de Personal ya Existente", vbInformation, Titulo
    
    End If
    
     If encontrado = False Then
     
     
          Hoja1.Unprotect (Seguridad)
          Hoja10.Unprotect (Seguridad)
          Hoja3.Unprotect (Seguridad)
          Hoja4.Unprotect (Seguridad)
        
        AccionPersonalContrato
        RegistrarPersonal

          Hoja1.Protect (Seguridad)
          Hoja10.Protect (Seguridad)
          Hoja3.Protect (Seguridad)
          Hoja4.Protect (Seguridad)
        
        MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
        
     End If
              
End Sub

Private Sub RegistrarPersonal()
Dim Comprb As Long
Dim FechaActual As Date

FechaActual = Date

    Hoja1.Select

    Hoja1.Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
       
            Hoja1.Cells(2, 1) = Me.txt_id.Text
            Hoja1.Cells(2, 2) = UCase(txt_Apellido) & " " & UCase(txt_nombre)
            Hoja1.Cells(2, 3) = txt_Cedula.Text
            Hoja1.Cells(2, 4) = txt_Telefono.Text
            Hoja1.Cells(2, 5) = UCase(txt_Area.Text)
            Hoja1.Cells(2, 6) = UCase(txt_Puesto.Text)
            Hoja1.Cells(2, 7) = txt_Regimen.Text
            Hoja1.Cells(2, 8) = txt_Jornada.Text
            Hoja1.Cells(2, 10) = txt_Salario.Value
            Hoja1.Cells(2, 11) = txt_Contrato.Text
            Hoja1.Cells(2, 15) = "-"
            Hoja1.Cells(2, 12) = CDate(Me.txt_Inicio)
            If Me.txt_Fin.Text = Empty Then
            Else
            Hoja1.Cells(2, 13) = CDate(Me.txt_Fin)
            End If
            Hoja1.Cells(2, 14) = txt_Pago.Text
            Hoja1.Cells(2, 16) = "ACTIVO"
            Hoja1.Cells(2, 17) = FechaActual
                       
            If Me.opt_Si.Value = True Then
            Hoja1.Cells(2, 9) = "SI"
            End If
            If Me.opt_No.Value = True Then
            Hoja1.Cells(2, 9) = "NO"
            End If
         
   ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("Tbl_personal").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("Tbl_personal").Sort. _
        SortFields.Add Key:=Range("Tbl_personal[ID PERSONAL]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("Tbl_personal").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Ajustar Planilla
    
       Hoja4.Select
       

    Hoja4.Rows("5:5").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    
                Hoja4.Cells(5, 1) = Me.txt_id.Text
                
    ActiveWorkbook.Worksheets("PLANILLA").ListObjects("Tbl_planilla").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("PLANILLA").ListObjects("Tbl_planilla").Sort.SortFields. _
        Add Key:=Range("A4"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("PLANILLA").ListObjects("Tbl_planilla").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
                

         'Ajustar CONSOLIDADO
    
       Hoja3.Select
       Range("A5").Select

       Selection.ListObject.ListRows.Add (1)
                    Hoja3.Range("A6:GQ6").Select
                    Selection.Copy
                    Hoja3.Range("A5").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                    
                Hoja3.Cells(5, 1) = Me.txt_id.Text
                
        ActiveWorkbook.Worksheets("CONSOLIDADO").ListObjects("Tbl_horario").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("CONSOLIDADO").ListObjects("Tbl_horario").Sort. _
        SortFields.Add Key:=Range("A5"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("CONSOLIDADO").ListObjects("Tbl_horario").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
     '''''''''''''''''''''''''''''''''
                
        Hoja1.Select
        Hoja1.Cells(1, 1).Select
        LimpiarControles
        txt_Apellido.SetFocus
       
End Sub
Private Sub AccionPersonalContrato()
Dim FechaActual As Date
Dim Fila As Long
Dim Final As Long
Dim Accion As String

FechaActual = Date
Accion = "CONTRATACIÓN"
    Hoja10.Select
    Hoja10.Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
        
   
            Hoja10.Cells(2, 1) = FechaActual
            Hoja10.Cells(2, 2) = Accion
            Hoja10.Cells(2, 3) = Me.txt_id.Text
            Hoja10.Cells(2, 4) = UCase(txt_Apellido) & " " & UCase(txt_nombre)
            Hoja10.Cells(2, 5) = txt_Cedula.Text
            Hoja10.Cells(2, 6) = txt_Telefono.Text
            Hoja10.Cells(2, 7) = UCase(txt_Area.Text)
            Hoja10.Cells(2, 8) = UCase(txt_Puesto.Text)
            Hoja10.Cells(2, 9) = txt_Regimen.Text
            Hoja10.Cells(2, 10) = txt_Jornada.Text
            Hoja10.Cells(2, 12) = txt_Salario.Value
            Hoja10.Cells(2, 13) = txt_Contrato.Text
            Hoja10.Cells(2, 14) = CDate(Me.txt_Inicio)
            If Me.txt_Fin.Text = Empty Then
            Else
            Hoja10.Cells(2, 15) = CDate(Me.txt_Fin)
            End If
            Hoja10.Cells(2, 16) = txt_Pago.Text
            Hoja10.Cells(2, 18) = "ACTIVO"
            Hoja10.Cells(2, 19) = Hoja83.Range("G1").Text
                       
            If Me.opt_Si.Value = True Then
            Hoja10.Cells(2, 11) = "SI"
            End If
            If Me.opt_No.Value = True Then
            Hoja10.Cells(2, 11) = "NO"
            End If

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub btn_Accion_Click()
Dim X As String
Dim encontrado As Boolean
Dim Titulo As String
Dim Fecha As String
Dim FechaActual As Date
Dim Seguridad As String
Seguridad = Hoja83.Range("L1").Text


Titulo = "Gestion del Personal"

FechaActual = Date

If Me.txt_Aid.Text = "" Then
    Me.txt_Aid.BackColor = &HC0C0FF
    Me.txt_Anombre.BackColor = &HC0C0FF
    MsgBox "Seleccione el código del personal..!", vbInformation, Titulo
    Me.txt_Aid.BackColor = &HFFFFFF
    Me.txt_Anombre.BackColor = &HFFFFFF
    Me.btn_ListadoPersonal.SetFocus
    Exit Sub
End If
      
X = Me.txt_Aid.Text

Hoja1.Select
Range("A1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like X Then
            encontrado = True
            Exit Do
                                 
        End If
    Loop
    If encontrado = False Then
        MsgBox "Personal no Existente", vbInformation, Titulo
    End If
    
    If encontrado = True Then
      
           
          Hoja1.Unprotect (Seguridad)
          Hoja10.Unprotect (Seguridad)
          Hoja3.Unprotect (Seguridad)
          Hoja4.Unprotect (Seguridad)
          
           
           ActiveCell.Offset(0, 2) = Me.txt_Acedula.Text
           ActiveCell.Offset(0, 3) = Me.txt_Atelefono.Text
           ActiveCell.Offset(0, 4) = Me.txt_Aarea.Text
           ActiveCell.Offset(0, 5) = Me.txt_Apuesto.Text
           ActiveCell.Offset(0, 6) = Me.txt_Aregimen.Text
           ActiveCell.Offset(0, 7) = Me.txt_Ajornada.Text
           ActiveCell.Offset(0, 9) = Me.txt_Asalario.Value
           ActiveCell.Offset(0, 10) = Me.txt_Acontrato.Text
           ActiveCell.Offset(0, 11) = CDate(txt_Ainicio.Text)


           If Me.txt_Afin.Text = Empty Then
           Else
           ActiveCell.Offset(0, 12) = CDate(txt_Afin.Text)
           End If
           
           ActiveCell.Offset(0, 13) = Me.txt_APago.Text
           ActiveCell.Offset(0, 14) = Me.txt_Abancaria.Text
           

           If Me.ck_Baja = True Then
                ActiveCell.Offset(0, 15) = "INACTIVO"
           End If
           
           If Me.ck_Reintegro = True Then
                ActiveCell.Offset(0, 15) = "ACTIVO"
            End If
            
            If Me.opt_Asi.Value = True Then
                ActiveCell.Offset(0, 8) = "SI"
             ElseIf Me.opt_Ano.Value = True Then
                ActiveCell.Offset(0, 8) = "NO"
            End If
            
            ActiveCell.Offset(0, 16) = FechaActual
           
        
    
  AccionPersonal
  Ajustar_planilla
  Ajustar_CONSOLIDADO
  
        Hoja1.Protect (Seguridad)
          Hoja10.Protect (Seguridad)
          Hoja3.Protect (Seguridad)
          Hoja4.Protect (Seguridad)
  
        Hoja1.Select
        Range("A1").Select
        MsgBox "Ajustes Realizados Correctamente..!", vbInformation, Titulo
        Limpiar_Acciones
        Me.txt_Aid.Text = ""
        Me.btn_ListadoPersonal.SetFocus
    End If
 
End Sub
Sub Ajustar_planilla()
Dim Final As String
Dim X As String
Dim encontrado As Boolean
Dim Titulo As String

X = Me.txt_Aid.Text

Hoja4.Select
Range("A4").Select
Titulo = "Gestion del Personal"

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like X Then
            encontrado = True
          Exit Do
        End If
    Loop
    If encontrado = True Then
      
            
            If ck_Baja.Value = True Then
                ActiveCell.EntireRow.Delete
            End If
                        
    End If

            If Me.ck_Reintegro.Value = True Then
            


    Hoja4.Rows("5:5").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    
                Hoja4.Cells(5, 1) = Me.txt_Aid.Text
                
    ActiveWorkbook.Worksheets("PLANILLA").ListObjects("Tbl_planilla").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("PLANILLA").ListObjects("Tbl_planilla").Sort.SortFields. _
        Add Key:=Range("A4"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("PLANILLA").ListObjects("Tbl_planilla").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
                MsgBox "Planilla Modificada Correctamente..!", vbInformation, Titulo
                
        End If
    
    If encontrado = False Then


    End If
           
                  

End Sub
Sub Ajustar_CONSOLIDADO()
Dim Final As String
Dim X As String
Dim encontrado As Boolean
Dim Titulo As String

X = Me.txt_Aid.Text

Hoja3.Select
Range("A4").Select
Titulo = "Gestion del Personal"

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like X Then
            encontrado = True
             Exit Do
        End If
    Loop
        
        If encontrado = True Then
    
          If ck_Baja = True Then
                ActiveCell.EntireRow.Delete
          End If
        End If
        
        If Me.ck_Reintegro = True Then
                 'Ajustar CONSOLIDADO
    
       Hoja3.Select
       Range("A5").Select

       Selection.ListObject.ListRows.Add (1)
                    Hoja3.Range("A6:GQ6").Select
                    Selection.Copy
                    Hoja3.Range("A5").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False
                    
                Hoja3.Cells(5, 1) = Me.txt_Aid.Text
                
        ActiveWorkbook.Worksheets("CONSOLIDADO").ListObjects("Tbl_horario").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("CONSOLIDADO").ListObjects("Tbl_horario").Sort. _
        SortFields.Add Key:=Range("A5"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("CONSOLIDADO").ListObjects("Tbl_horario").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
     '''''''''''''''''''''''''''''''''
     
      MsgBox "Consolidador de Horas Modificado Correctamente..!", vbInformation, Titu
      
          End If
    
    
    If encontrado = False Then
       

    End If
           
     

        

End Sub

Private Sub Limpiar_Acciones()

    Me.txt_Aid.Text = Empty
    Me.ck_Contrato.Value = False
    Me.ck_Traslado.Value = False
    Me.ck_Aumento.Value = False
    Me.ck_Otros.Value = False
    Me.ck_Baja.Value = False
    Me.ck_Reintegro.Value = False
    

End Sub

Private Sub AccionPersonal()
Dim FechaActual As Date
Dim Fila As Long
Dim Final As Long
Dim Accion As String

FechaActual = Date

    Hoja10.Select
    Hoja10.Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
        
   
    If Me.ck_Baja.Value = True Then
        Accion = "BAJA"
        
    ElseIf Me.ck_Reintegro.Value = True Then
        Accion = "REINTEGRO"
    ElseIf Me.ck_Traslado.Value = True Or Me.ck_Aumento.Value = True Or Me.ck_Contrato.Value = True Then
        Accion = "VARIOS"
    Else: Accion = "OTROS"
    End If
             
            
            Hoja10.Cells(2, 1) = FechaActual
            Hoja10.Cells(2, 2) = Accion
            Hoja10.Cells(2, 3) = Me.txt_Aid.Text
            Hoja10.Cells(2, 4) = Me.txt_Anombre.Text
            Hoja10.Cells(2, 5) = txt_Acedula.Text
            Hoja10.Cells(2, 6) = txt_Atelefono.Text
            Hoja10.Cells(2, 7) = UCase(txt_Aarea.Text)
            Hoja10.Cells(2, 8) = UCase(txt_Apuesto.Text)
            Hoja10.Cells(2, 9) = txt_Aregimen.Text
            Hoja10.Cells(2, 10) = txt_Ajornada.Text
            Hoja10.Cells(2, 12) = txt_Asalario.Value
            Hoja10.Cells(2, 13) = txt_Acontrato.Text
            Hoja10.Cells(2, 14) = CDate(Me.txt_Ainicio)
            If Me.txt_Afin.Text = Empty Then
            Else
            Hoja10.Cells(2, 15) = CDate(Me.txt_Afin)
            End If
            Hoja10.Cells(2, 16) = txt_APago.Text
            Hoja10.Cells(2, 17) = txt_Abancaria.Text
            Hoja10.Cells(2, 18) = "ACTIVO"
            Hoja10.Cells(2, 19) = Hoja83.Range("G1").Text
                       
            If Me.opt_Asi.Value = True Then
            Hoja10.Cells(2, 11) = "SI"
            End If
            If Me.opt_Ano.Value = True Then
            Hoja10.Cells(2, 11) = "NO"
            End If
            

End Sub
