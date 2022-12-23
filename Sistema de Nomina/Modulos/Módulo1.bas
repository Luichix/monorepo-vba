Attribute VB_Name = "Módulo1"

Option Explicit
Private Sub Registrar_Hora()
Dim Fecha As Date
Dim Titulo As String
Dim Seguridad As String
Dim Formato As String

Dim xEnter1 As Date
Dim xExit1 As Date
Dim xEnter2 As Date
Dim xExit2 As Date
Dim xEnter3 As Date
Dim xExit3 As Date
Dim xEnter4 As Date
Dim xExit4 As Date
Dim xEnter5 As Date
Dim xExit5 As Date
Dim xEnter6 As Date
Dim xExit6 As Date
Dim xEnter7 As Date
Dim xExit7 As Date
Dim xEnter8 As Date
Dim xExit8 As Date
Dim xEnter9 As Date
Dim xExit9 As Date
Dim xEnter10 As Date
Dim xExit10 As Date
Dim xEnter11 As Date
Dim xExit11 As Date
Dim xEnter12 As Date
Dim xExit12 As Date
Dim xEnter13 As Date
Dim xExit13 As Date
Dim xEnter14 As Date
Dim xExit14 As Date
Dim xEnter15 As Date
Dim xExit15 As Date
Dim xEnter16 As Date
Dim xExit16 As Date
Dim xEnter17 As Date
Dim xExit17 As Date

Seguridad = Hoja83.Range("L1").Text

Hoja2.Unprotect (Seguridad)

Titulo = "Gestor de Recursos Humanos"
Formato = "00:00"
    

xEnter1 = Me.txt_xEntrada1.Value
xExit1 = Me.txt_xSalida1.Value
          
           
         If Me.txt_xEntrada1 <> Formato And Me.txt_xSalida1 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha1)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter1
                Hoja2.Cells(3, 6) = xExit1
        End If

    
xEnter2 = Me.txt_xEntrada2.Value
xExit2 = Me.txt_xSalida2.Value
              
        If Me.txt_xEntrada2 <> Formato And Me.txt_xSalida2 <> Formato Then
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha2)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter2
                Hoja2.Cells(3, 6) = xExit2
        End If
           
xEnter3 = Me.txt_xEntrada3.Value
xExit3 = Me.txt_xSalida3.Value

           
           If Me.txt_xEntrada3 <> Formato And Me.txt_xSalida3 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha3)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter3
                Hoja2.Cells(3, 6) = xExit3
          End If
        
xEnter4 = Me.txt_xEntrada4.Value
xExit4 = Me.txt_xSalida4.Value

           
 
        If Me.txt_xEntrada4 <> Formato And Me.txt_xSalida4 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha4)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter4
                Hoja2.Cells(3, 6) = xExit4
        End If

    
xEnter5 = Me.txt_xEntrada5.Value
xExit5 = Me.txt_xSalida5.Value

           
 
        If Me.txt_xEntrada5 <> Formato And Me.txt_xSalida5 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha5)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter5
                Hoja2.Cells(3, 6) = xExit5
        End If
    
xEnter6 = Me.txt_xEntrada6.Value
xExit6 = Me.txt_xSalida6.Value

           
           
 
        If Me.txt_xEntrada6 <> Formato And Me.txt_xSalida6 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha6)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter6
                Hoja2.Cells(3, 6) = xExit6
        End If
        
xEnter7 = Me.txt_xEntrada7.Value
xExit7 = Me.txt_xSalida7.Value

           
        If Me.txt_xEntrada7 <> Formato And Me.txt_xSalida7 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha7)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter7
                Hoja2.Cells(3, 6) = xExit7
        End If
   
xEnter8 = Me.txt_xEntrada8.Value
xExit8 = Me.txt_xSalida8.Value

        If Me.txt_xEntrada8 <> Formato And Me.txt_xSalida8 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha8)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter8
                Hoja2.Cells(3, 6) = xExit8
        End If
     
    
xEnter9 = Me.txt_xEntrada9.Value
xExit9 = Me.txt_xSalida9.Value
           

        If Me.txt_xEntrada9 <> Formato And Me.txt_xSalida9 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha9)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter9
                Hoja2.Cells(3, 6) = xExit9
        End If

    
xEnter10 = Me.txt_xEntrada10.Value
xExit10 = Me.txt_xSalida10.Value
           
        If Me.txt_xEntrada10 <> Formato And Me.txt_xSalida10 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha10)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter10
                Hoja2.Cells(3, 6) = xExit10
        End If
      
    
xEnter11 = Me.txt_xEntrada11.Value
xExit11 = Me.txt_xSalida11.Value
           
        If Me.txt_xEntrada11 <> Formato And Me.txt_xSalida11 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha11)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter11
                Hoja2.Cells(3, 6) = xExit11
        End If
      
xEnter12 = Me.txt_xEntrada12.Value
xExit12 = Me.txt_xSalida12.Value

        If Me.txt_xEntrada12 <> Formato And Me.txt_xSalida12 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha12)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter12
                Hoja2.Cells(3, 6) = xExit12
        End If
   
    
xEnter13 = Me.txt_xEntrada13.Value
xExit13 = Me.txt_xSalida13.Value

  
        If Me.txt_xEntrada13 <> Formato And Me.txt_xSalida13 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha13)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter13
                Hoja2.Cells(3, 6) = xExit13
        End If
    
xEnter14 = Me.txt_xEntrada14.Value
xExit14 = Me.txt_xSalida14.Value
           
        If Me.txt_xEntrada14 <> Formato And Me.txt_xSalida14 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha14)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter14
                Hoja2.Cells(3, 6) = xExit14
        End If
    
xEnter15 = Me.txt_xEntrada15.Value
xExit15 = Me.txt_xSalida15.Value
           
        If Me.txt_xEntrada15 <> Formato And Me.txt_xSalida15 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha15)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter15
                Hoja2.Cells(3, 6) = xExit15
        End If
    
xEnter16 = Me.txt_xEntrada16.Value
xExit16 = Me.txt_xSalida16.Value
           
        If Me.txt_xEntrada16 <> Formato And Me.txt_xSalida16 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_Hora_Marca.txt_fecha16)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter16
                Hoja2.Cells(3, 6) = xExit16
        End If

 LimpiarHora
 Hoja2.Protect (Seguridad)
 
 
         MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
             
End Sub
Private Sub LimpiarHora()

Dim Ctrl As Control
    Me.txt_id = Empty
    Me.txt_nombre = Empty
    Me.txt_Fecha = Empty
    
    For Each Ctrl In Me.Controls
        If Ctrl.Name Like "txt_fecha" & "*" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name Like "txt_xEntrada" & "*" Or Ctrl.Name Like "txt_xSalida" & "*" Or Ctrl.Name Like "txt_yEntrada" & "*" Or Ctrl.Name Like "txt_ySalida" & "*" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub

Private Sub btn_Fecha_Click()
banderaCalendario = 24
  Call LanzarCalendario(Me, "btn_fecha")
End Sub

Private Sub btn_Registrar_Click()
On Error GoTo Salir
Dim Titulo As String
Dim Formato As String

Dim LEntrada1 As Date
Dim LSalida1 As Date
Dim LEntrada2 As Date
Dim LSalida2 As Date
Dim LEntrada3 As Date
Dim LSalida3 As Date
Dim LEntrada4 As Date
Dim LSalida4 As Date
Dim LEntrada5 As Date
Dim LSalida5 As Date
Dim LEntrada6 As Date
Dim LSalida6 As Date
Dim LEntrada7 As Date
Dim LSalida7 As Date
Dim LEntrada8 As Date
Dim LSalida8 As Date
Dim NEntrada8 As Date
Dim NSalida8 As Date
Dim LEntrada9 As Date
Dim LSalida9 As Date
Dim LEntrada10 As Date
Dim LSalida10 As Date
Dim LEntrada11 As Date
Dim LSalida11 As Date
Dim LEntrada12 As Date
Dim LSalida12 As Date
Dim LEntrada13 As Date
Dim LSalida13 As Date
Dim LEntrada14 As Date
Dim LSalida14 As Date
Dim LEntrada15 As Date
Dim LSalida15 As Date
Dim LEntrada16 As Date
Dim LSalida16 As Date

Formato = "00:00"
Titulo = "Gestor de Recursos Humanos"

    If Me.txt_id.Text = "" Or Me.txt_nombre.Text = "" Then
            MsgBox "Debe seleccionar un colaborador del Listado..!", vbInformation, "Gestor de Recursos Humanos"
            Exit Sub
    End If
    
    If Me.txt_fecha1 = "" Then
        If Me.txt_xEntrada1 <> Formato Or Me.txt_xSalida1 <> Formato Then
            Me.txt_fecha1.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 01..!", vbInformation, Titulo
            Me.txt_fecha1.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    If Me.txt_fecha1 <> Empty Then
        If Me.txt_fecha1 = Me.txt_fecha2 Or Me.txt_fecha1 = Me.txt_fecha3 Or Me.txt_fecha1 = Me.txt_fecha4 Or _
        Me.txt_fecha1 = Me.txt_fecha5 Or Me.txt_fecha1 = Me.txt_fecha6 Or Me.txt_fecha1 = Me.txt_fecha7 Or _
        Me.txt_fecha1 = Me.txt_fecha8 Or Me.txt_fecha1 = Me.txt_fecha9 Or Me.txt_fecha1 = Me.txt_fecha10 Or _
        Me.txt_fecha1 = Me.txt_fecha11 Or Me.txt_fecha1 = Me.txt_fecha12 Or Me.txt_fecha1 = Me.txt_fecha13 Or _
        Me.txt_fecha1 = Me.txt_fecha14 Or Me.txt_fecha1 = Me.txt_fecha15 Or Me.txt_fecha1 = Me.txt_fecha16 Then
            Me.txt_fecha1.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha1.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If
    
LEntrada1 = Me.txt_xEntrada1.Value
LSalida1 = Me.txt_xSalida1.Value

                        
                        If LEntrada1 >= LSalida1 Then
                            Me.txt_xEntrada1.BackColor = &HC0C0FF
                            Me.txt_xSalida1.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada1.BackColor = &HFFFFFF
                            Me.txt_xSalida1.BackColor = &HFFFFFF
                            Me.txt_xEntrada1.SetFocus
                            Exit Sub
                        End If
                        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha2 = "" Then
        If Me.txt_xEntrada2 <> Formato Or Me.txt_xSalida2 <> Formato Then
            Me.txt_fecha2.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha2.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha2 <> Empty Then
        If Me.txt_fecha2 = Me.txt_fecha1 Or Me.txt_fecha2 = Me.txt_fecha3 Or Me.txt_fecha2 = Me.txt_fecha4 Or _
        Me.txt_fecha2 = Me.txt_fecha5 Or Me.txt_fecha2 = Me.txt_fecha6 Or Me.txt_fecha2 = Me.txt_fecha7 Or _
        Me.txt_fecha2 = Me.txt_fecha8 Or Me.txt_fecha2 = Me.txt_fecha9 Or Me.txt_fecha2 = Me.txt_fecha10 Or _
        Me.txt_fecha2 = Me.txt_fecha11 Or Me.txt_fecha2 = Me.txt_fecha12 Or Me.txt_fecha2 = Me.txt_fecha13 Or _
        Me.txt_fecha2 = Me.txt_fecha14 Or Me.txt_fecha2 = Me.txt_fecha15 Or Me.txt_fecha2 = Me.txt_fecha16 Then
            Me.txt_fecha2.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha2.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If
    
LEntrada2 = Me.txt_xEntrada2.Value
LSalida2 = Me.txt_xSalida2.Value
                        
                If Me.txt_xEntrada2 <> Formato Or Me.txt_xSalida2 <> Formato Then
                        If LEntrada2 >= LSalida2 Then
                            Me.txt_xEntrada2.BackColor = &HC0C0FF
                            Me.txt_xSalida2.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada2.BackColor = &HFFFFFF
                            Me.txt_xSalida2.BackColor = &HFFFFFF
                            Me.txt_xEntrada2.SetFocus
                            Exit Sub
                        End If
                End If
                        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha3 = "" Then
        If Me.txt_xEntrada3 <> Formato Or Me.txt_xSalida3 <> Formato Then
            Me.txt_fecha3.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha3.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha3 <> Empty Then
        If Me.txt_fecha3 = Me.txt_fecha1 Or Me.txt_fecha3 = Me.txt_fecha2 Or Me.txt_fecha3 = Me.txt_fecha4 Or _
        Me.txt_fecha3 = Me.txt_fecha5 Or Me.txt_fecha3 = Me.txt_fecha6 Or Me.txt_fecha3 = Me.txt_fecha7 Or _
        Me.txt_fecha3 = Me.txt_fecha8 Or Me.txt_fecha3 = Me.txt_fecha9 Or Me.txt_fecha3 = Me.txt_fecha10 Or _
        Me.txt_fecha3 = Me.txt_fecha11 Or Me.txt_fecha3 = Me.txt_fecha12 Or Me.txt_fecha3 = Me.txt_fecha13 Or _
        Me.txt_fecha3 = Me.txt_fecha14 Or Me.txt_fecha3 = Me.txt_fecha15 Or Me.txt_fecha3 = Me.txt_fecha16 Then
            Me.txt_fecha3.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha3.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If

LEntrada3 = Me.txt_xEntrada3.Value
LSalida3 = Me.txt_xSalida3.Value


                        
                If Me.txt_xEntrada3 <> Formato Or Me.txt_xSalida3 <> Formato Then
                        If LEntrada3 >= LSalida3 Then
                            Me.txt_xEntrada3.BackColor = &HC0C0FF
                            Me.txt_xSalida3.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada3.BackColor = &HFFFFFF
                            Me.txt_xSalida3.BackColor = &HFFFFFF
                            Me.txt_xEntrada3.SetFocus
                            Exit Sub
                        End If
                End If
            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha4 = "" Then
        If Me.txt_xEntrada4 <> Formato Or Me.txt_xSalida4 <> Formato Then
            Me.txt_fecha4.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha4.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha4 <> Empty Then
        If Me.txt_fecha4 = Me.txt_fecha1 Or Me.txt_fecha4 = Me.txt_fecha3 Or Me.txt_fecha4 = Me.txt_fecha2 Or _
        Me.txt_fecha4 = Me.txt_fecha5 Or Me.txt_fecha4 = Me.txt_fecha6 Or Me.txt_fecha4 = Me.txt_fecha7 Or _
        Me.txt_fecha4 = Me.txt_fecha8 Or Me.txt_fecha4 = Me.txt_fecha9 Or Me.txt_fecha4 = Me.txt_fecha10 Or _
        Me.txt_fecha4 = Me.txt_fecha11 Or Me.txt_fecha4 = Me.txt_fecha12 Or Me.txt_fecha4 = Me.txt_fecha13 Or _
        Me.txt_fecha4 = Me.txt_fecha14 Or Me.txt_fecha4 = Me.txt_fecha15 Or Me.txt_fecha4 = Me.txt_fecha16 Then
            Me.txt_fecha4.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha4.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If

    
LEntrada4 = Me.txt_xEntrada4.Value
LSalida4 = Me.txt_xSalida4.Value

                        
                If Me.txt_xEntrada4 <> Formato Or Me.txt_xSalida4 <> Formato Then
                        If LEntrada4 >= LSalida4 Then
                            Me.txt_xEntrada4.BackColor = &HC0C0FF
                            Me.txt_xSalida4.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada4.BackColor = &HFFFFFF
                            Me.txt_xSalida4.BackColor = &HFFFFFF
                            Me.txt_xEntrada4.SetFocus
                            Exit Sub
                        End If
                End If
                        
           
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha5 = "" Then
        If Me.txt_xEntrada5 <> Formato Or Me.txt_xSalida5 <> Formato Then
            Me.txt_fecha5.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha5.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha5 <> Empty Then
        If Me.txt_fecha5 = Me.txt_fecha1 Or Me.txt_fecha5 = Me.txt_fecha3 Or Me.txt_fecha5 = Me.txt_fecha4 Or _
        Me.txt_fecha5 = Me.txt_fecha2 Or Me.txt_fecha5 = Me.txt_fecha6 Or Me.txt_fecha5 = Me.txt_fecha7 Or _
        Me.txt_fecha5 = Me.txt_fecha8 Or Me.txt_fecha5 = Me.txt_fecha9 Or Me.txt_fecha5 = Me.txt_fecha10 Or _
        Me.txt_fecha5 = Me.txt_fecha11 Or Me.txt_fecha5 = Me.txt_fecha12 Or Me.txt_fecha5 = Me.txt_fecha13 Or _
        Me.txt_fecha5 = Me.txt_fecha14 Or Me.txt_fecha5 = Me.txt_fecha15 Or Me.txt_fecha5 = Me.txt_fecha16 Then
            Me.txt_fecha5.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha5.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If

LEntrada5 = Me.txt_xEntrada5.Value
LSalida5 = Me.txt_xSalida5.Value

                        
                If Me.txt_xEntrada5 <> Formato Or Me.txt_xSalida5 <> Formato Then
                        If LEntrada5 >= LSalida5 Then
                            Me.txt_xEntrada5.BackColor = &HC0C0FF
                            Me.txt_xSalida5.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada5.BackColor = &HFFFFFF
                            Me.txt_xSalida5.BackColor = &HFFFFFF
                            Me.txt_xEntrada5.SetFocus
                            Exit Sub
                        End If
                End If
                        
      
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha6 = "" Then
        If Me.txt_xEntrada6 <> Formato Or Me.txt_xSalida6 <> Formato Then
            Me.txt_fecha6.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha6.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha6 <> Empty Then
        If Me.txt_fecha6 = Me.txt_fecha1 Or Me.txt_fecha6 = Me.txt_fecha3 Or Me.txt_fecha6 = Me.txt_fecha4 Or _
        Me.txt_fecha6 = Me.txt_fecha5 Or Me.txt_fecha6 = Me.txt_fecha2 Or Me.txt_fecha6 = Me.txt_fecha7 Or _
        Me.txt_fecha6 = Me.txt_fecha8 Or Me.txt_fecha6 = Me.txt_fecha9 Or Me.txt_fecha6 = Me.txt_fecha10 Or _
        Me.txt_fecha6 = Me.txt_fecha11 Or Me.txt_fecha6 = Me.txt_fecha12 Or Me.txt_fecha6 = Me.txt_fecha13 Or _
        Me.txt_fecha6 = Me.txt_fecha14 Or Me.txt_fecha6 = Me.txt_fecha15 Or Me.txt_fecha6 = Me.txt_fecha16 Then
            Me.txt_fecha6.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha6.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If
  
LEntrada6 = Me.txt_xEntrada6.Value
LSalida6 = Me.txt_xSalida6.Value


                        
                If Me.txt_xEntrada6 <> Formato Or Me.txt_xSalida6 <> Formato Then
                        If LEntrada6 >= LSalida6 Then
                            Me.txt_xEntrada6.BackColor = &HC0C0FF
                            Me.txt_xSalida6.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada6.BackColor = &HFFFFFF
                            Me.txt_xSalida6.BackColor = &HFFFFFF
                            Me.txt_xEntrada6.SetFocus
                            Exit Sub
                        End If
                End If
                        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha7 = "" Then
        If Me.txt_xEntrada7 <> Formato Or Me.txt_xSalida7 <> Formato Then
            Me.txt_fecha7.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha7.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha7 <> Empty Then
        If Me.txt_fecha7 = Me.txt_fecha1 Or Me.txt_fecha7 = Me.txt_fecha3 Or Me.txt_fecha7 = Me.txt_fecha4 Or _
        Me.txt_fecha7 = Me.txt_fecha5 Or Me.txt_fecha7 = Me.txt_fecha6 Or Me.txt_fecha7 = Me.txt_fecha2 Or _
        Me.txt_fecha7 = Me.txt_fecha8 Or Me.txt_fecha7 = Me.txt_fecha9 Or Me.txt_fecha7 = Me.txt_fecha10 Or _
        Me.txt_fecha7 = Me.txt_fecha11 Or Me.txt_fecha7 = Me.txt_fecha12 Or Me.txt_fecha7 = Me.txt_fecha13 Or _
        Me.txt_fecha7 = Me.txt_fecha14 Or Me.txt_fecha7 = Me.txt_fecha15 Or Me.txt_fecha7 = Me.txt_fecha16 Then
            Me.txt_fecha7.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha7.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If

    
LEntrada7 = Me.txt_xEntrada7.Value
LSalida7 = Me.txt_xSalida7.Value


                        
                If Me.txt_xEntrada7 <> Formato Or Me.txt_xSalida7 <> Formato Then
                        If LEntrada7 >= LSalida7 Then
                            Me.txt_xEntrada7.BackColor = &HC0C0FF
                            Me.txt_xSalida7.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada7.BackColor = &HFFFFFF
                            Me.txt_xSalida7.BackColor = &HFFFFFF
                            Me.txt_xEntrada7.SetFocus
                            Exit Sub
                        End If
                End If
                        

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha8 = "" Then
        If Me.txt_xEntrada8 <> Formato Or Me.txt_xSalida8 <> Formato Then
            Me.txt_fecha8.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha8.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha8 <> Empty Then
        If Me.txt_fecha8 = Me.txt_fecha1 Or Me.txt_fecha8 = Me.txt_fecha3 Or Me.txt_fecha8 = Me.txt_fecha4 Or _
        Me.txt_fecha8 = Me.txt_fecha5 Or Me.txt_fecha8 = Me.txt_fecha6 Or Me.txt_fecha8 = Me.txt_fecha7 Or _
        Me.txt_fecha8 = Me.txt_fecha2 Or Me.txt_fecha8 = Me.txt_fecha9 Or Me.txt_fecha8 = Me.txt_fecha10 Or _
        Me.txt_fecha8 = Me.txt_fecha11 Or Me.txt_fecha8 = Me.txt_fecha12 Or Me.txt_fecha8 = Me.txt_fecha13 Or _
        Me.txt_fecha8 = Me.txt_fecha14 Or Me.txt_fecha8 = Me.txt_fecha15 Or Me.txt_fecha8 = Me.txt_fecha16 Then
            Me.txt_fecha8.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha8.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If
    
LEntrada8 = Me.txt_xEntrada8.Value
LSalida8 = Me.txt_xSalida8.Value


                        
                If Me.txt_xEntrada8 <> Formato Or Me.txt_xSalida8 <> Formato Then
                        If LEntrada8 >= LSalida8 Then
                            Me.txt_xEntrada8.BackColor = &HC0C0FF
                            Me.txt_xSalida8.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada8.BackColor = &HFFFFFF
                            Me.txt_xSalida8.BackColor = &HFFFFFF
                            Me.txt_xEntrada8.SetFocus
                            Exit Sub
                        End If
                End If
                        


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha9 = "" Then
        If Me.txt_xEntrada9 <> Formato Or Me.txt_xSalida9 <> Formato Then
            Me.txt_fecha9.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha9.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha9 <> Empty Then
        If Me.txt_fecha9 = Me.txt_fecha1 Or Me.txt_fecha9 = Me.txt_fecha3 Or Me.txt_fecha9 = Me.txt_fecha4 Or _
        Me.txt_fecha9 = Me.txt_fecha5 Or Me.txt_fecha9 = Me.txt_fecha6 Or Me.txt_fecha9 = Me.txt_fecha7 Or _
        Me.txt_fecha9 = Me.txt_fecha8 Or Me.txt_fecha9 = Me.txt_fecha2 Or Me.txt_fecha9 = Me.txt_fecha10 Or _
        Me.txt_fecha9 = Me.txt_fecha11 Or Me.txt_fecha9 = Me.txt_fecha12 Or Me.txt_fecha9 = Me.txt_fecha13 Or _
        Me.txt_fecha9 = Me.txt_fecha14 Or Me.txt_fecha9 = Me.txt_fecha15 Or Me.txt_fecha9 = Me.txt_fecha16 Then
            Me.txt_fecha9.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha9.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If
 
LEntrada9 = Me.txt_xEntrada9.Value
LSalida9 = Me.txt_xSalida9.Value

                        
                If Me.txt_xEntrada9 <> Formato Or Me.txt_xSalida9 <> Formato Then
                        If LEntrada9 >= LSalida9 Then
                            Me.txt_xEntrada9.BackColor = &HC0C0FF
                            Me.txt_xSalida9.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada9.BackColor = &HFFFFFF
                            Me.txt_xSalida9.BackColor = &HFFFFFF
                            Me.txt_xEntrada9.SetFocus
                            Exit Sub
                        End If
                End If
                        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha10 = "" Then
        If Me.txt_xEntrada10 <> Formato Or Me.txt_xSalida10 <> Formato Then
            Me.txt_fecha10.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha10.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha10 <> Empty Then
        If Me.txt_fecha10 = Me.txt_fecha1 Or Me.txt_fecha10 = Me.txt_fecha3 Or Me.txt_fecha10 = Me.txt_fecha4 Or _
        Me.txt_fecha10 = Me.txt_fecha5 Or Me.txt_fecha10 = Me.txt_fecha6 Or Me.txt_fecha10 = Me.txt_fecha7 Or _
        Me.txt_fecha10 = Me.txt_fecha8 Or Me.txt_fecha10 = Me.txt_fecha9 Or Me.txt_fecha10 = Me.txt_fecha2 Or _
        Me.txt_fecha10 = Me.txt_fecha11 Or Me.txt_fecha10 = Me.txt_fecha12 Or Me.txt_fecha10 = Me.txt_fecha13 Or _
        Me.txt_fecha10 = Me.txt_fecha14 Or Me.txt_fecha10 = Me.txt_fecha15 Or Me.txt_fecha10 = Me.txt_fecha16 Then
            Me.txt_fecha10.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha10.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If
      
LEntrada10 = Me.txt_xEntrada10.Value
LSalida10 = Me.txt_xSalida10.Value
                        
                If Me.txt_xEntrada10 <> Formato Or Me.txt_xSalida10 <> Formato Then
                        If LEntrada10 >= LSalida10 Then
                            Me.txt_xEntrada10.BackColor = &HC0C0FF
                            Me.txt_xSalida10.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada10.BackColor = &HFFFFFF
                            Me.txt_xSalida10.BackColor = &HFFFFFF
                            Me.txt_xEntrada10.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha11 = "" Then
        If Me.txt_xEntrada11 <> Formato Or Me.txt_xSalida11 <> Formato Then
            Me.txt_fecha11.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha11.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha11 <> Empty Then
        If Me.txt_fecha11 = Me.txt_fecha1 Or Me.txt_fecha11 = Me.txt_fecha3 Or Me.txt_fecha11 = Me.txt_fecha4 Or _
        Me.txt_fecha11 = Me.txt_fecha5 Or Me.txt_fecha11 = Me.txt_fecha6 Or Me.txt_fecha11 = Me.txt_fecha7 Or _
        Me.txt_fecha11 = Me.txt_fecha8 Or Me.txt_fecha11 = Me.txt_fecha9 Or Me.txt_fecha11 = Me.txt_fecha10 Or _
        Me.txt_fecha11 = Me.txt_fecha2 Or Me.txt_fecha11 = Me.txt_fecha12 Or Me.txt_fecha11 = Me.txt_fecha13 Or _
        Me.txt_fecha11 = Me.txt_fecha14 Or Me.txt_fecha11 = Me.txt_fecha15 Or Me.txt_fecha11 = Me.txt_fecha16 Then
            Me.txt_fecha11.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha11.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If
   
    
LEntrada11 = Me.txt_xEntrada11.Value
LSalida11 = Me.txt_xSalida11.Value

                        
                If Me.txt_xEntrada11 <> Formato Or Me.txt_xSalida11 <> Formato Then
                        If LEntrada11 >= LSalida11 Then
                            Me.txt_xEntrada11.BackColor = &HC0C0FF
                            Me.txt_xSalida11.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada11.BackColor = &HFFFFFF
                            Me.txt_xSalida11.BackColor = &HFFFFFF
                            Me.txt_xEntrada11.SetFocus
                            Exit Sub
                        End If
                End If
                        
           
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha12 = "" Then
        If Me.txt_xEntrada12 <> Formato Or Me.txt_xSalida12 <> Formato Then
            Me.txt_fecha12.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha12.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha12 <> Empty Then
        If Me.txt_fecha12 = Me.txt_fecha1 Or Me.txt_fecha12 = Me.txt_fecha3 Or Me.txt_fecha12 = Me.txt_fecha4 Or _
        Me.txt_fecha12 = Me.txt_fecha5 Or Me.txt_fecha12 = Me.txt_fecha6 Or Me.txt_fecha12 = Me.txt_fecha7 Or _
        Me.txt_fecha12 = Me.txt_fecha8 Or Me.txt_fecha12 = Me.txt_fecha9 Or Me.txt_fecha12 = Me.txt_fecha10 Or _
        Me.txt_fecha12 = Me.txt_fecha11 Or Me.txt_fecha12 = Me.txt_fecha2 Or Me.txt_fecha12 = Me.txt_fecha13 Or _
        Me.txt_fecha12 = Me.txt_fecha14 Or Me.txt_fecha12 = Me.txt_fecha15 Or Me.txt_fecha12 = Me.txt_fecha16 Then
            Me.txt_fecha12.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha12.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If

    
LEntrada12 = Me.txt_xEntrada12.Value
LSalida12 = Me.txt_xSalida12.Value


                        
                If Me.txt_xEntrada12 <> Formato Or Me.txt_xSalida12 <> Formato Then
                        If LEntrada12 >= LSalida12 Then
                            Me.txt_xEntrada12.BackColor = &HC0C0FF
                            Me.txt_xSalida12.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada12.BackColor = &HFFFFFF
                            Me.txt_xSalida12.BackColor = &HFFFFFF
                            Me.txt_xEntrada12.SetFocus
                            Exit Sub
                        End If
                End If
                        
    

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha13 = "" Then
        If Me.txt_xEntrada13 <> Formato Or Me.txt_xSalida13 <> Formato Then
            Me.txt_fecha13.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha13.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha13 <> Empty Then
        If Me.txt_fecha13 = Me.txt_fecha1 Or Me.txt_fecha13 = Me.txt_fecha3 Or Me.txt_fecha13 = Me.txt_fecha4 Or _
        Me.txt_fecha13 = Me.txt_fecha5 Or Me.txt_fecha13 = Me.txt_fecha6 Or Me.txt_fecha13 = Me.txt_fecha7 Or _
        Me.txt_fecha13 = Me.txt_fecha8 Or Me.txt_fecha13 = Me.txt_fecha9 Or Me.txt_fecha13 = Me.txt_fecha10 Or _
        Me.txt_fecha13 = Me.txt_fecha11 Or Me.txt_fecha13 = Me.txt_fecha12 Or Me.txt_fecha13 = Me.txt_fecha2 Or _
        Me.txt_fecha13 = Me.txt_fecha14 Or Me.txt_fecha13 = Me.txt_fecha15 Or Me.txt_fecha13 = Me.txt_fecha16 Then
            Me.txt_fecha13.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha13.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If

LEntrada13 = Me.txt_xEntrada13.Value
LSalida13 = Me.txt_xSalida13.Value


                        
                If Me.txt_xEntrada13 <> Formato Or Me.txt_xSalida13 <> Formato Then
                        If LEntrada13 >= LSalida13 Then
                            Me.txt_xEntrada13.BackColor = &HC0C0FF
                            Me.txt_xSalida13.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada13.BackColor = &HFFFFFF
                            Me.txt_xSalida13.BackColor = &HFFFFFF
                            Me.txt_xEntrada13.SetFocus
                            Exit Sub
                        End If
                End If
 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha14 = "" Then
        If Me.txt_xEntrada14 <> Formato Or Me.txt_xSalida14 <> Formato Then
            Me.txt_fecha14.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha14.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha14 <> Empty Then
        If Me.txt_fecha14 = Me.txt_fecha1 Or Me.txt_fecha14 = Me.txt_fecha3 Or Me.txt_fecha14 = Me.txt_fecha4 Or _
        Me.txt_fecha14 = Me.txt_fecha5 Or Me.txt_fecha14 = Me.txt_fecha6 Or Me.txt_fecha14 = Me.txt_fecha7 Or _
        Me.txt_fecha14 = Me.txt_fecha8 Or Me.txt_fecha14 = Me.txt_fecha9 Or Me.txt_fecha14 = Me.txt_fecha10 Or _
        Me.txt_fecha14 = Me.txt_fecha11 Or Me.txt_fecha14 = Me.txt_fecha12 Or Me.txt_fecha14 = Me.txt_fecha13 Or _
        Me.txt_fecha14 = Me.txt_fecha2 Or Me.txt_fecha14 = Me.txt_fecha15 Or Me.txt_fecha14 = Me.txt_fecha16 Then
            Me.txt_fecha14.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha14.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If
    
LEntrada14 = Me.txt_xEntrada14.Value
LSalida14 = Me.txt_xSalida14.Value


                        
                If Me.txt_xEntrada14 <> Formato Or Me.txt_xSalida14 <> Formato Then
                        If LEntrada14 >= LSalida14 Then
                            Me.txt_xEntrada14.BackColor = &HC0C0FF
                            Me.txt_xSalida14.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada14.BackColor = &HFFFFFF
                            Me.txt_xSalida14.BackColor = &HFFFFFF
                            Me.txt_xEntrada14.SetFocus
                            Exit Sub
                        End If
                End If
                        
      

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha15 = "" Then
        If Me.txt_xEntrada15 <> Formato Or Me.txt_xSalida15 <> Formato Then
            Me.txt_fecha15.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha15.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha15 <> Empty Then
        If Me.txt_fecha15 = Me.txt_fecha1 Or Me.txt_fecha15 = Me.txt_fecha3 Or Me.txt_fecha15 = Me.txt_fecha4 Or _
        Me.txt_fecha15 = Me.txt_fecha5 Or Me.txt_fecha15 = Me.txt_fecha6 Or Me.txt_fecha15 = Me.txt_fecha7 Or _
        Me.txt_fecha15 = Me.txt_fecha8 Or Me.txt_fecha15 = Me.txt_fecha9 Or Me.txt_fecha15 = Me.txt_fecha10 Or _
        Me.txt_fecha15 = Me.txt_fecha11 Or Me.txt_fecha15 = Me.txt_fecha12 Or Me.txt_fecha15 = Me.txt_fecha13 Or _
        Me.txt_fecha15 = Me.txt_fecha14 Or Me.txt_fecha15 = Me.txt_fecha2 Or Me.txt_fecha15 = Me.txt_fecha16 Then
            Me.txt_fecha15.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha15.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If

LEntrada15 = Me.txt_xEntrada15.Value
LSalida15 = Me.txt_xSalida15.Value


                        
                If Me.txt_xEntrada15 <> Formato Or Me.txt_xSalida15 <> Formato Then
                        If LEntrada15 >= LSalida15 Then
                            Me.txt_xEntrada15.BackColor = &HC0C0FF
                            Me.txt_xSalida15.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada15.BackColor = &HFFFFFF
                            Me.txt_xSalida15.BackColor = &HFFFFFF
                            Me.txt_xEntrada15.SetFocus
                            Exit Sub
                        End If
                End If
                        
      

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha16 = "" Then
        If Me.txt_xEntrada16 <> Formato Or Me.txt_xSalida16 <> Formato Then
            Me.txt_fecha16.BackColor = &HC0C0FF
            MsgBox "Seleccione la fecha de registro: Registros 02..!", vbInformation, Titulo
            Me.txt_fecha16.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
    If Me.txt_fecha16 <> Empty Then
        If Me.txt_fecha16 = Me.txt_fecha1 Or Me.txt_fecha16 = Me.txt_fecha3 Or Me.txt_fecha16 = Me.txt_fecha4 Or _
        Me.txt_fecha16 = Me.txt_fecha5 Or Me.txt_fecha16 = Me.txt_fecha6 Or Me.txt_fecha16 = Me.txt_fecha7 Or _
        Me.txt_fecha16 = Me.txt_fecha8 Or Me.txt_fecha16 = Me.txt_fecha9 Or Me.txt_fecha16 = Me.txt_fecha10 Or _
        Me.txt_fecha16 = Me.txt_fecha11 Or Me.txt_fecha16 = Me.txt_fecha12 Or Me.txt_fecha16 = Me.txt_fecha13 Or _
        Me.txt_fecha16 = Me.txt_fecha14 Or Me.txt_fecha16 = Me.txt_fecha15 Or Me.txt_fecha16 = Me.txt_fecha2 Then
            Me.txt_fecha16.BackColor = &HC0C0FF
            MsgBox "La fecha de registro seleccionada se encuentra repetida..!", vbInformation, Titulo
            Me.txt_fecha16.BackColor = &HFFFFFF
        Exit Sub
        End If
    End If
     
LEntrada16 = Me.txt_xEntrada16.Value
LSalida16 = Me.txt_xSalida16.Value


                        
                If Me.txt_xEntrada16 <> Formato Or Me.txt_xSalida16 <> Formato Then
                        If LEntrada16 >= LSalida16 Then
                            Me.txt_xEntrada16.BackColor = &HC0C0FF
                            Me.txt_xSalida16.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada16.BackColor = &HFFFFFF
                            Me.txt_xSalida16.BackColor = &HFFFFFF
                            Me.txt_xEntrada16.SetFocus
                            Exit Sub
                        End If
                End If
                        

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Registrar_Hora

                    
     
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Ventas"
 End If

End Sub
Private Sub btn_fecha1_Click()
banderaCalendario = 8
  Call LanzarCalendario(Me, "btn_fecha1")
End Sub
Private Sub btn_fecha2_Click()
banderaCalendario = 9
  Call LanzarCalendario(Me, "txt_fecha2")
End Sub
Private Sub btn_fecha3_Click()
banderaCalendario = 10
  Call LanzarCalendario(Me, "txt_fecha3")
End Sub
Private Sub btn_fecha4_Click()
banderaCalendario = 11
  Call LanzarCalendario(Me, "txt_fecha4")
End Sub
Private Sub btn_fecha5_Click()
banderaCalendario = 12
  Call LanzarCalendario(Me, "txt_fecha5")
End Sub
Private Sub btn_fecha6_Click()
banderaCalendario = 13
  Call LanzarCalendario(Me, "txt_fecha6")
End Sub
Private Sub btn_fecha7_Click()
banderaCalendario = 14
  Call LanzarCalendario(Me, "txt_fecha7")
End Sub
Private Sub btn_fecha8_Click()
banderaCalendario = 15
  Call LanzarCalendario(Me, "txt_fecha8")
End Sub
Private Sub btn_fecha9_Click()
banderaCalendario = 16
  Call LanzarCalendario(Me, "txt_fecha9")
End Sub
Private Sub btn_fecha10_Click()
banderaCalendario = 17
  Call LanzarCalendario(Me, "txt_fecha10")
End Sub
Private Sub btn_fecha11_Click()
banderaCalendario = 18
  Call LanzarCalendario(Me, "txt_fecha11")
End Sub
Private Sub btn_fecha12_Click()
banderaCalendario = 19
  Call LanzarCalendario(Me, "txt_fecha12")
End Sub
Private Sub btn_fecha13_Click()
banderaCalendario = 20
  Call LanzarCalendario(Me, "txt_fecha13")
End Sub
Private Sub btn_fecha14_Click()
banderaCalendario = 21
  Call LanzarCalendario(Me, "txt_fecha14")
End Sub
Private Sub btn_fecha15_Click()
banderaCalendario = 22
  Call LanzarCalendario(Me, "txt_fecha15")
End Sub
Private Sub btn_fecha16_Click()
banderaCalendario = 23
  Call LanzarCalendario(Me, "txt_fecha16")
End Sub

Private Sub btn_limpiar1_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha1" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada1" Or Ctrl.Name = "txt_xSalida1" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar2_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha2" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada2" Or Ctrl.Name = "txt_xSalida2" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar3_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha3" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada3" Or Ctrl.Name = "txt_xSalida3" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar4_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha4" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada4" Or Ctrl.Name = "txt_xSalida4" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar5_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha5" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada5" Or Ctrl.Name = "txt_xSalida5" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar6_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha6" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada6" Or Ctrl.Name = "txt_xSalida6" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar7_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha7" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada7" Or Ctrl.Name = "txt_xSalida7" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar8_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha8" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada8" Or Ctrl.Name = "txt_xSalida8" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar9_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha9" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada9" Or Ctrl.Name = "txt_xSalida9" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar10_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha10" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada10" Or Ctrl.Name = "txt_xSalida10" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar11_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha11" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada11" Or Ctrl.Name = "txt_xSalida11" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar12_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha12" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada12" Or Ctrl.Name = "txt_xSalida12" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar13_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha13" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada13" Or Ctrl.Name = "txt_xSalida13" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar14_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha14" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada14" Or Ctrl.Name = "txt_xSalida14" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar15_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha15" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada15" Or Ctrl.Name = "txt_xSalida15" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub btn_limpiar16_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha16" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada16" Or Ctrl.Name = "txt_xSalida16" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub
Private Sub txt_xEntrada1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada1.Text <> "1" And txt_xEntrada1.Text <> "2" And txt_xEntrada1.Text <> "3" And txt_xEntrada1.Text <> "4" And txt_xEntrada1.Text <> "0" Then
    Select Case Len(txt_xEntrada1.Value)
        Case 1
        txt_xEntrada1.Value = txt_xEntrada1.Value & ":"
        Me.txt_xEntrada1.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada1.Value > 9 And txt_xEntrada1.Value < 24 Then
    Select Case Len(txt_xEntrada1.Value)
        Case 2
        txt_xEntrada1.Value = txt_xEntrada1.Value & ":"
        Me.txt_xEntrada1.MaxLength = 5
        End Select
End If
If txt_xEntrada1.Value > 23 And txt_xEntrada1.Value < 30 Or txt_xEntrada1.Value = 0 Or txt_xEntrada1.Value = 3 Or txt_xEntrada1.Value = 4 Then
    txt_xEntrada1 = "00:00"
     Me.txt_xEntrada1.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida1.Text <> "1" And txt_xSalida1.Text <> "2" And txt_xSalida1.Text <> "3" And txt_xSalida1.Text <> "4" And txt_xSalida1.Text <> "0" Then
    Select Case Len(txt_xSalida1.Value)
        Case 1
        txt_xSalida1.Value = txt_xSalida1.Value & ":"
        Me.txt_xSalida1.MaxLength = 4
          End Select
        
    End If
If txt_xSalida1.Value > 9 And txt_xSalida1.Value < 24 Then
    Select Case Len(txt_xSalida1.Value)
        Case 2
        txt_xSalida1.Value = txt_xSalida1.Value & ":"
        Me.txt_xSalida1.MaxLength = 5
        End Select
End If
If txt_xSalida1.Value > 23 And txt_xSalida1.Value < 30 Or txt_xSalida1.Value = 0 Or txt_xSalida1.Value = 3 Or txt_xSalida1.Value = 4 Then
    txt_xSalida1 = "00:00"
     Me.txt_xSalida1.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada2.Text <> "1" And txt_xEntrada2.Text <> "2" And txt_xEntrada2.Text <> "3" And txt_xEntrada2.Text <> "4" And txt_xEntrada2.Text <> "0" Then
    Select Case Len(txt_xEntrada2.Value)
        Case 1
        txt_xEntrada2.Value = txt_xEntrada2.Value & ":"
        Me.txt_xEntrada2.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada2.Value > 9 And txt_xEntrada2.Value < 24 Then
    Select Case Len(txt_xEntrada2.Value)
        Case 2
        txt_xEntrada2.Value = txt_xEntrada2.Value & ":"
        Me.txt_xEntrada2.MaxLength = 5
        End Select
End If
If txt_xEntrada2.Value > 23 And txt_xEntrada2.Value < 30 Or txt_xEntrada2.Value = 0 Or txt_xEntrada2.Value = 3 Or txt_xEntrada2.Value = 4 Then
    txt_xEntrada2 = "00:00"
     Me.txt_xEntrada2.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida2.Text <> "1" And txt_xSalida2.Text <> "2" And txt_xSalida2.Text <> "3" And txt_xSalida2.Text <> "4" And txt_xSalida2.Text <> "0" Then
    Select Case Len(txt_xSalida2.Value)
        Case 1
        txt_xSalida2.Value = txt_xSalida2.Value & ":"
        Me.txt_xSalida2.MaxLength = 4
          End Select
        
    End If
If txt_xSalida2.Value > 9 And txt_xSalida2.Value < 24 Then
    Select Case Len(txt_xSalida2.Value)
        Case 2
        txt_xSalida2.Value = txt_xSalida2.Value & ":"
        Me.txt_xSalida2.MaxLength = 5
        End Select
End If
If txt_xSalida2.Value > 23 And txt_xSalida2.Value < 30 Or txt_xSalida2.Value = 0 Or txt_xSalida2.Value = 3 Or txt_xSalida2.Value = 4 Then
    txt_xSalida2 = "00:00"
     Me.txt_xSalida2.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada3.Text <> "1" And txt_xEntrada3.Text <> "2" And txt_xEntrada3.Text <> "3" And txt_xEntrada3.Text <> "4" And txt_xEntrada3.Text <> "0" Then
    Select Case Len(txt_xEntrada3.Value)
        Case 1
        txt_xEntrada3.Value = txt_xEntrada3.Value & ":"
        Me.txt_xEntrada3.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada3.Value > 9 And txt_xEntrada3.Value < 24 Then
    Select Case Len(txt_xEntrada3.Value)
        Case 2
        txt_xEntrada3.Value = txt_xEntrada3.Value & ":"
        Me.txt_xEntrada3.MaxLength = 5
        End Select
End If
If txt_xEntrada3.Value > 23 And txt_xEntrada3.Value < 30 Or txt_xEntrada3.Value = 0 Or txt_xEntrada3.Value = 3 Or txt_xEntrada3.Value = 4 Then
    txt_xEntrada3 = "00:00"
     Me.txt_xEntrada3.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida3.Text <> "1" And txt_xSalida3.Text <> "2" And txt_xSalida3.Text <> "3" And txt_xSalida3.Text <> "4" And txt_xSalida3.Text <> "0" Then
    Select Case Len(txt_xSalida3.Value)
        Case 1
        txt_xSalida3.Value = txt_xSalida3.Value & ":"
        Me.txt_xSalida3.MaxLength = 4
          End Select
        
    End If
If txt_xSalida3.Value > 9 And txt_xSalida3.Value < 24 Then
    Select Case Len(txt_xSalida3.Value)
        Case 2
        txt_xSalida3.Value = txt_xSalida3.Value & ":"
        Me.txt_xSalida3.MaxLength = 5
        End Select
End If
If txt_xSalida3.Value > 23 And txt_xSalida3.Value < 30 Or txt_xSalida3.Value = 0 Or txt_xSalida3.Value = 3 Or txt_xSalida3.Value = 4 Then
    txt_xSalida3 = "00:00"
     Me.txt_xSalida3.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada4.Text <> "1" And txt_xEntrada4.Text <> "2" And txt_xEntrada4.Text <> "3" And txt_xEntrada4.Text <> "4" And txt_xEntrada4.Text <> "0" Then
    Select Case Len(txt_xEntrada4.Value)
        Case 1
        txt_xEntrada4.Value = txt_xEntrada4.Value & ":"
        Me.txt_xEntrada4.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada4.Value > 9 And txt_xEntrada4.Value < 24 Then
    Select Case Len(txt_xEntrada4.Value)
        Case 2
        txt_xEntrada4.Value = txt_xEntrada4.Value & ":"
        Me.txt_xEntrada4.MaxLength = 5
        End Select
End If
If txt_xEntrada4.Value > 23 And txt_xEntrada4.Value < 30 Or txt_xEntrada4.Value = 0 Or txt_xEntrada4.Value = 3 Or txt_xEntrada4.Value = 4 Then
    txt_xEntrada4 = "00:00"
     Me.txt_xEntrada4.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida4.Text <> "1" And txt_xSalida4.Text <> "2" And txt_xSalida4.Text <> "3" And txt_xSalida4.Text <> "4" And txt_xSalida4.Text <> "0" Then
    Select Case Len(txt_xSalida4.Value)
        Case 1
        txt_xSalida4.Value = txt_xSalida4.Value & ":"
        Me.txt_xSalida4.MaxLength = 4
          End Select
        
    End If
If txt_xSalida4.Value > 9 And txt_xSalida4.Value < 24 Then
    Select Case Len(txt_xSalida4.Value)
        Case 2
        txt_xSalida4.Value = txt_xSalida4.Value & ":"
        Me.txt_xSalida4.MaxLength = 5
        End Select
End If
If txt_xSalida4.Value > 23 And txt_xSalida4.Value < 30 Or txt_xSalida4.Value = 0 Or txt_xSalida4.Value = 3 Or txt_xSalida4.Value = 4 Then
    txt_xSalida4 = "00:00"
     Me.txt_xSalida4.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada5.Text <> "1" And txt_xEntrada5.Text <> "2" And txt_xEntrada5.Text <> "3" And txt_xEntrada5.Text <> "4" And txt_xEntrada5.Text <> "0" Then
    Select Case Len(txt_xEntrada5.Value)
        Case 1
        txt_xEntrada5.Value = txt_xEntrada5.Value & ":"
        Me.txt_xEntrada5.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada5.Value > 9 And txt_xEntrada5.Value < 24 Then
    Select Case Len(txt_xEntrada5.Value)
        Case 2
        txt_xEntrada5.Value = txt_xEntrada5.Value & ":"
        Me.txt_xEntrada5.MaxLength = 5
        End Select
End If
If txt_xEntrada5.Value > 23 And txt_xEntrada5.Value < 30 Or txt_xEntrada5.Value = 0 Or txt_xEntrada5.Value = 3 Or txt_xEntrada5.Value = 4 Then
    txt_xEntrada5 = "00:00"
     Me.txt_xEntrada5.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida5.Text <> "1" And txt_xSalida5.Text <> "2" And txt_xSalida5.Text <> "3" And txt_xSalida5.Text <> "4" And txt_xSalida5.Text <> "0" Then
    Select Case Len(txt_xSalida5.Value)
        Case 1
        txt_xSalida5.Value = txt_xSalida5.Value & ":"
        Me.txt_xSalida5.MaxLength = 4
          End Select
        
    End If
If txt_xSalida5.Value > 9 And txt_xSalida5.Value < 24 Then
    Select Case Len(txt_xSalida5.Value)
        Case 2
        txt_xSalida5.Value = txt_xSalida5.Value & ":"
        Me.txt_xSalida5.MaxLength = 5
        End Select
End If
If txt_xSalida5.Value > 23 And txt_xSalida5.Value < 30 Or txt_xSalida5.Value = 0 Or txt_xSalida5.Value = 3 Or txt_xSalida5.Value = 4 Then
    txt_xSalida5 = "00:00"
     Me.txt_xSalida5.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada6.Text <> "1" And txt_xEntrada6.Text <> "2" And txt_xEntrada6.Text <> "3" And txt_xEntrada6.Text <> "4" And txt_xEntrada6.Text <> "0" Then
    Select Case Len(txt_xEntrada6.Value)
        Case 1
        txt_xEntrada6.Value = txt_xEntrada6.Value & ":"
        Me.txt_xEntrada6.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada6.Value > 9 And txt_xEntrada6.Value < 24 Then
    Select Case Len(txt_xEntrada6.Value)
        Case 2
        txt_xEntrada6.Value = txt_xEntrada6.Value & ":"
        Me.txt_xEntrada6.MaxLength = 5
        End Select
End If
If txt_xEntrada6.Value > 23 And txt_xEntrada6.Value < 30 Or txt_xEntrada6.Value = 0 Or txt_xEntrada6.Value = 3 Or txt_xEntrada6.Value = 4 Then
    txt_xEntrada6 = "00:00"
     Me.txt_xEntrada6.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida6.Text <> "1" And txt_xSalida6.Text <> "2" And txt_xSalida6.Text <> "3" And txt_xSalida6.Text <> "4" And txt_xSalida6.Text <> "0" Then
    Select Case Len(txt_xSalida6.Value)
        Case 1
        txt_xSalida6.Value = txt_xSalida6.Value & ":"
        Me.txt_xSalida6.MaxLength = 4
          End Select
        
    End If
If txt_xSalida6.Value > 9 And txt_xSalida6.Value < 24 Then
    Select Case Len(txt_xSalida6.Value)
        Case 2
        txt_xSalida6.Value = txt_xSalida6.Value & ":"
        Me.txt_xSalida6.MaxLength = 5
        End Select
End If
If txt_xSalida6.Value > 23 And txt_xSalida6.Value < 30 Or txt_xSalida6.Value = 0 Or txt_xSalida6.Value = 3 Or txt_xSalida6.Value = 4 Then
    txt_xSalida6 = "00:00"
     Me.txt_xSalida6.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada7.Text <> "1" And txt_xEntrada7.Text <> "2" And txt_xEntrada7.Text <> "3" And txt_xEntrada7.Text <> "4" And txt_xEntrada7.Text <> "0" Then
    Select Case Len(txt_xEntrada7.Value)
        Case 1
        txt_xEntrada7.Value = txt_xEntrada7.Value & ":"
        Me.txt_xEntrada7.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada7.Value > 9 And txt_xEntrada7.Value < 24 Then
    Select Case Len(txt_xEntrada7.Value)
        Case 2
        txt_xEntrada7.Value = txt_xEntrada7.Value & ":"
        Me.txt_xEntrada7.MaxLength = 5
        End Select
End If
If txt_xEntrada7.Value > 23 And txt_xEntrada7.Value < 30 Or txt_xEntrada7.Value = 0 Or txt_xEntrada7.Value = 3 Or txt_xEntrada7.Value = 4 Then
    txt_xEntrada7 = "00:00"
     Me.txt_xEntrada7.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida7.Text <> "1" And txt_xSalida7.Text <> "2" And txt_xSalida7.Text <> "3" And txt_xSalida7.Text <> "4" And txt_xSalida7.Text <> "0" Then
    Select Case Len(txt_xSalida7.Value)
        Case 1
        txt_xSalida7.Value = txt_xSalida7.Value & ":"
        Me.txt_xSalida7.MaxLength = 4
          End Select
        
    End If
If txt_xSalida7.Value > 9 And txt_xSalida7.Value < 24 Then
    Select Case Len(txt_xSalida7.Value)
        Case 2
        txt_xSalida7.Value = txt_xSalida7.Value & ":"
        Me.txt_xSalida7.MaxLength = 5
        End Select
End If
If txt_xSalida7.Value > 23 And txt_xSalida7.Value < 30 Or txt_xSalida7.Value = 0 Or txt_xSalida7.Value = 3 Or txt_xSalida7.Value = 4 Then
    txt_xSalida7 = "00:00"
     Me.txt_xSalida7.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada8.Text <> "1" And txt_xEntrada8.Text <> "2" And txt_xEntrada8.Text <> "3" And txt_xEntrada8.Text <> "4" And txt_xEntrada8.Text <> "0" Then
    Select Case Len(txt_xEntrada8.Value)
        Case 1
        txt_xEntrada8.Value = txt_xEntrada8.Value & ":"
        Me.txt_xEntrada8.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada8.Value > 9 And txt_xEntrada8.Value < 24 Then
    Select Case Len(txt_xEntrada8.Value)
        Case 2
        txt_xEntrada8.Value = txt_xEntrada8.Value & ":"
        Me.txt_xEntrada8.MaxLength = 5
        End Select
End If
If txt_xEntrada8.Value > 23 And txt_xEntrada8.Value < 30 Or txt_xEntrada8.Value = 0 Or txt_xEntrada8.Value = 3 Or txt_xEntrada8.Value = 4 Then
    txt_xEntrada8 = "00:00"
     Me.txt_xEntrada8.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida8.Text <> "1" And txt_xSalida8.Text <> "2" And txt_xSalida8.Text <> "3" And txt_xSalida8.Text <> "4" And txt_xSalida8.Text <> "0" Then
    Select Case Len(txt_xSalida8.Value)
        Case 1
        txt_xSalida8.Value = txt_xSalida8.Value & ":"
        Me.txt_xSalida8.MaxLength = 4
          End Select
        
    End If
If txt_xSalida8.Value > 9 And txt_xSalida8.Value < 24 Then
    Select Case Len(txt_xSalida8.Value)
        Case 2
        txt_xSalida8.Value = txt_xSalida8.Value & ":"
        Me.txt_xSalida8.MaxLength = 5
        End Select
End If
If txt_xSalida8.Value > 23 And txt_xSalida8.Value < 30 Or txt_xSalida8.Value = 0 Or txt_xSalida8.Value = 3 Or txt_xSalida8.Value = 4 Then
    txt_xSalida8 = "00:00"
     Me.txt_xSalida8.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada9.Text <> "1" And txt_xEntrada9.Text <> "2" And txt_xEntrada9.Text <> "3" And txt_xEntrada9.Text <> "4" And txt_xEntrada9.Text <> "0" Then
    Select Case Len(txt_xEntrada9.Value)
        Case 1
        txt_xEntrada9.Value = txt_xEntrada9.Value & ":"
        Me.txt_xEntrada9.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada9.Value > 9 And txt_xEntrada9.Value < 24 Then
    Select Case Len(txt_xEntrada9.Value)
        Case 2
        txt_xEntrada9.Value = txt_xEntrada9.Value & ":"
        Me.txt_xEntrada9.MaxLength = 5
        End Select
End If
If txt_xEntrada9.Value > 23 And txt_xEntrada9.Value < 30 Or txt_xEntrada9.Value = 0 Or txt_xEntrada9.Value = 3 Or txt_xEntrada9.Value = 4 Then
    txt_xEntrada9 = "00:00"
     Me.txt_xEntrada9.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida9.Text <> "1" And txt_xSalida9.Text <> "2" And txt_xSalida9.Text <> "3" And txt_xSalida9.Text <> "4" And txt_xSalida9.Text <> "0" Then
    Select Case Len(txt_xSalida9.Value)
        Case 1
        txt_xSalida9.Value = txt_xSalida9.Value & ":"
        Me.txt_xSalida9.MaxLength = 4
          End Select
        
    End If
If txt_xSalida9.Value > 9 And txt_xSalida9.Value < 24 Then
    Select Case Len(txt_xSalida9.Value)
        Case 2
        txt_xSalida9.Value = txt_xSalida9.Value & ":"
        Me.txt_xSalida9.MaxLength = 5
        End Select
End If
If txt_xSalida9.Value > 23 And txt_xSalida9.Value < 30 Or txt_xSalida9.Value = 0 Or txt_xSalida9.Value = 3 Or txt_xSalida9.Value = 4 Then
    txt_xSalida9 = "00:00"
     Me.txt_xSalida9.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada10.Text <> "1" And txt_xEntrada10.Text <> "2" And txt_xEntrada10.Text <> "3" And txt_xEntrada10.Text <> "4" And txt_xEntrada10.Text <> "0" Then
    Select Case Len(txt_xEntrada10.Value)
        Case 1
        txt_xEntrada10.Value = txt_xEntrada10.Value & ":"
        Me.txt_xEntrada10.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada10.Value > 9 And txt_xEntrada10.Value < 24 Then
    Select Case Len(txt_xEntrada10.Value)
        Case 2
        txt_xEntrada10.Value = txt_xEntrada10.Value & ":"
        Me.txt_xEntrada10.MaxLength = 5
        End Select
End If
If txt_xEntrada10.Value > 23 And txt_xEntrada10.Value < 30 Or txt_xEntrada10.Value = 0 Or txt_xEntrada10.Value = 3 Or txt_xEntrada10.Value = 4 Then
    txt_xEntrada10 = "00:00"
     Me.txt_xEntrada10.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida10.Text <> "1" And txt_xSalida10.Text <> "2" And txt_xSalida10.Text <> "3" And txt_xSalida10.Text <> "4" And txt_xSalida10.Text <> "0" Then
    Select Case Len(txt_xSalida10.Value)
        Case 1
        txt_xSalida10.Value = txt_xSalida10.Value & ":"
        Me.txt_xSalida10.MaxLength = 4
          End Select
        
    End If
If txt_xSalida10.Value > 9 And txt_xSalida10.Value < 24 Then
    Select Case Len(txt_xSalida10.Value)
        Case 2
        txt_xSalida10.Value = txt_xSalida10.Value & ":"
        Me.txt_xSalida10.MaxLength = 5
        End Select
End If
If txt_xSalida10.Value > 23 And txt_xSalida10.Value < 30 Or txt_xSalida10.Value = 0 Or txt_xSalida10.Value = 3 Or txt_xSalida10.Value = 4 Then
    txt_xSalida10 = "00:00"
     Me.txt_xSalida10.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada11.Text <> "1" And txt_xEntrada11.Text <> "2" And txt_xEntrada11.Text <> "3" And txt_xEntrada11.Text <> "4" And txt_xEntrada11.Text <> "0" Then
    Select Case Len(txt_xEntrada11.Value)
        Case 1
        txt_xEntrada11.Value = txt_xEntrada11.Value & ":"
        Me.txt_xEntrada11.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada11.Value > 9 And txt_xEntrada11.Value < 24 Then
    Select Case Len(txt_xEntrada11.Value)
        Case 2
        txt_xEntrada11.Value = txt_xEntrada11.Value & ":"
        Me.txt_xEntrada11.MaxLength = 5
        End Select
End If
If txt_xEntrada11.Value > 23 And txt_xEntrada11.Value < 30 Or txt_xEntrada11.Value = 0 Or txt_xEntrada11.Value = 3 Or txt_xEntrada11.Value = 4 Then
    txt_xEntrada11 = "00:00"
     Me.txt_xEntrada11.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida11.Text <> "1" And txt_xSalida11.Text <> "2" And txt_xSalida11.Text <> "3" And txt_xSalida11.Text <> "4" And txt_xSalida11.Text <> "0" Then
    Select Case Len(txt_xSalida11.Value)
        Case 1
        txt_xSalida11.Value = txt_xSalida11.Value & ":"
        Me.txt_xSalida11.MaxLength = 4
          End Select
        
    End If
If txt_xSalida11.Value > 9 And txt_xSalida11.Value < 24 Then
    Select Case Len(txt_xSalida11.Value)
        Case 2
        txt_xSalida11.Value = txt_xSalida11.Value & ":"
        Me.txt_xSalida11.MaxLength = 5
        End Select
End If
If txt_xSalida11.Value > 23 And txt_xSalida11.Value < 30 Or txt_xSalida11.Value = 0 Or txt_xSalida11.Value = 3 Or txt_xSalida11.Value = 4 Then
    txt_xSalida11 = "00:00"
     Me.txt_xSalida11.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada12.Text <> "1" And txt_xEntrada12.Text <> "2" And txt_xEntrada12.Text <> "3" And txt_xEntrada12.Text <> "4" And txt_xEntrada12.Text <> "0" Then
    Select Case Len(txt_xEntrada12.Value)
        Case 1
        txt_xEntrada12.Value = txt_xEntrada12.Value & ":"
        Me.txt_xEntrada12.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada12.Value > 9 And txt_xEntrada12.Value < 24 Then
    Select Case Len(txt_xEntrada12.Value)
        Case 2
        txt_xEntrada12.Value = txt_xEntrada12.Value & ":"
        Me.txt_xEntrada12.MaxLength = 5
        End Select
End If
If txt_xEntrada12.Value > 23 And txt_xEntrada12.Value < 30 Or txt_xEntrada12.Value = 0 Or txt_xEntrada12.Value = 3 Or txt_xEntrada12.Value = 4 Then
    txt_xEntrada12 = "00:00"
     Me.txt_xEntrada12.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida12.Text <> "1" And txt_xSalida12.Text <> "2" And txt_xSalida12.Text <> "3" And txt_xSalida12.Text <> "4" And txt_xSalida12.Text <> "0" Then
    Select Case Len(txt_xSalida12.Value)
        Case 1
        txt_xSalida12.Value = txt_xSalida12.Value & ":"
        Me.txt_xSalida12.MaxLength = 4
          End Select
        
    End If
If txt_xSalida12.Value > 9 And txt_xSalida12.Value < 24 Then
    Select Case Len(txt_xSalida12.Value)
        Case 2
        txt_xSalida12.Value = txt_xSalida12.Value & ":"
        Me.txt_xSalida12.MaxLength = 5
        End Select
End If
If txt_xSalida12.Value > 23 And txt_xSalida12.Value < 30 Or txt_xSalida12.Value = 0 Or txt_xSalida12.Value = 3 Or txt_xSalida12.Value = 4 Then
    txt_xSalida12 = "00:00"
     Me.txt_xSalida12.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada13.Text <> "1" And txt_xEntrada13.Text <> "2" And txt_xEntrada13.Text <> "3" And txt_xEntrada13.Text <> "4" And txt_xEntrada13.Text <> "0" Then
    Select Case Len(txt_xEntrada13.Value)
        Case 1
        txt_xEntrada13.Value = txt_xEntrada13.Value & ":"
        Me.txt_xEntrada13.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada13.Value > 9 And txt_xEntrada13.Value < 24 Then
    Select Case Len(txt_xEntrada13.Value)
        Case 2
        txt_xEntrada13.Value = txt_xEntrada13.Value & ":"
        Me.txt_xEntrada13.MaxLength = 5
        End Select
End If
If txt_xEntrada13.Value > 23 And txt_xEntrada13.Value < 30 Or txt_xEntrada13.Value = 0 Or txt_xEntrada13.Value = 3 Or txt_xEntrada13.Value = 4 Then
    txt_xEntrada13 = "00:00"
     Me.txt_xEntrada13.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida13.Text <> "1" And txt_xSalida13.Text <> "2" And txt_xSalida13.Text <> "3" And txt_xSalida13.Text <> "4" And txt_xSalida13.Text <> "0" Then
    Select Case Len(txt_xSalida13.Value)
        Case 1
        txt_xSalida13.Value = txt_xSalida13.Value & ":"
        Me.txt_xSalida13.MaxLength = 4
          End Select
        
    End If
If txt_xSalida13.Value > 9 And txt_xSalida13.Value < 24 Then
    Select Case Len(txt_xSalida13.Value)
        Case 2
        txt_xSalida13.Value = txt_xSalida13.Value & ":"
        Me.txt_xSalida13.MaxLength = 5
        End Select
End If
If txt_xSalida13.Value > 23 And txt_xSalida13.Value < 30 Or txt_xSalida13.Value = 0 Or txt_xSalida13.Value = 3 Or txt_xSalida13.Value = 4 Then
    txt_xSalida13 = "00:00"
     Me.txt_xSalida13.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada14.Text <> "1" And txt_xEntrada14.Text <> "2" And txt_xEntrada14.Text <> "3" And txt_xEntrada14.Text <> "4" And txt_xEntrada14.Text <> "0" Then
    Select Case Len(txt_xEntrada14.Value)
        Case 1
        txt_xEntrada14.Value = txt_xEntrada14.Value & ":"
        Me.txt_xEntrada14.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada14.Value > 9 And txt_xEntrada14.Value < 24 Then
    Select Case Len(txt_xEntrada14.Value)
        Case 2
        txt_xEntrada14.Value = txt_xEntrada14.Value & ":"
        Me.txt_xEntrada14.MaxLength = 5
        End Select
End If
If txt_xEntrada14.Value > 23 And txt_xEntrada14.Value < 30 Or txt_xEntrada14.Value = 0 Or txt_xEntrada14.Value = 3 Or txt_xEntrada14.Value = 4 Then
    txt_xEntrada14 = "00:00"
     Me.txt_xEntrada14.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida14.Text <> "1" And txt_xSalida14.Text <> "2" And txt_xSalida14.Text <> "3" And txt_xSalida14.Text <> "4" And txt_xSalida14.Text <> "0" Then
    Select Case Len(txt_xSalida14.Value)
        Case 1
        txt_xSalida14.Value = txt_xSalida14.Value & ":"
        Me.txt_xSalida14.MaxLength = 4
          End Select
        
    End If
If txt_xSalida14.Value > 9 And txt_xSalida14.Value < 24 Then
    Select Case Len(txt_xSalida14.Value)
        Case 2
        txt_xSalida14.Value = txt_xSalida14.Value & ":"
        Me.txt_xSalida14.MaxLength = 5
        End Select
End If
If txt_xSalida14.Value > 23 And txt_xSalida14.Value < 30 Or txt_xSalida14.Value = 0 Or txt_xSalida14.Value = 3 Or txt_xSalida14.Value = 4 Then
    txt_xSalida14 = "00:00"
     Me.txt_xSalida14.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada15.Text <> "1" And txt_xEntrada15.Text <> "2" And txt_xEntrada15.Text <> "3" And txt_xEntrada15.Text <> "4" And txt_xEntrada15.Text <> "0" Then
    Select Case Len(txt_xEntrada15.Value)
        Case 1
        txt_xEntrada15.Value = txt_xEntrada15.Value & ":"
        Me.txt_xEntrada15.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada15.Value > 9 And txt_xEntrada15.Value < 24 Then
    Select Case Len(txt_xEntrada15.Value)
        Case 2
        txt_xEntrada15.Value = txt_xEntrada15.Value & ":"
        Me.txt_xEntrada15.MaxLength = 5
        End Select
End If
If txt_xEntrada15.Value > 23 And txt_xEntrada15.Value < 30 Or txt_xEntrada15.Value = 0 Or txt_xEntrada15.Value = 3 Or txt_xEntrada15.Value = 4 Then
    txt_xEntrada15 = "00:00"
     Me.txt_xEntrada15.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida15.Text <> "1" And txt_xSalida15.Text <> "2" And txt_xSalida15.Text <> "3" And txt_xSalida15.Text <> "4" And txt_xSalida15.Text <> "0" Then
    Select Case Len(txt_xSalida15.Value)
        Case 1
        txt_xSalida15.Value = txt_xSalida15.Value & ":"
        Me.txt_xSalida15.MaxLength = 4
          End Select
        
    End If
If txt_xSalida15.Value > 9 And txt_xSalida15.Value < 24 Then
    Select Case Len(txt_xSalida15.Value)
        Case 2
        txt_xSalida15.Value = txt_xSalida15.Value & ":"
        Me.txt_xSalida15.MaxLength = 5
        End Select
End If
If txt_xSalida15.Value > 23 And txt_xSalida15.Value < 30 Or txt_xSalida15.Value = 0 Or txt_xSalida15.Value = 3 Or txt_xSalida15.Value = 4 Then
    txt_xSalida15 = "00:00"
     Me.txt_xSalida15.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xEntrada16.Text <> "1" And txt_xEntrada16.Text <> "2" And txt_xEntrada16.Text <> "3" And txt_xEntrada16.Text <> "4" And txt_xEntrada16.Text <> "0" Then
    Select Case Len(txt_xEntrada16.Value)
        Case 1
        txt_xEntrada16.Value = txt_xEntrada16.Value & ":"
        Me.txt_xEntrada16.MaxLength = 4
          End Select
        
    End If
If txt_xEntrada16.Value > 9 And txt_xEntrada16.Value < 24 Then
    Select Case Len(txt_xEntrada16.Value)
        Case 2
        txt_xEntrada16.Value = txt_xEntrada16.Value & ":"
        Me.txt_xEntrada16.MaxLength = 5
        End Select
End If
If txt_xEntrada16.Value > 23 And txt_xEntrada16.Value < 30 Or txt_xEntrada16.Value = 0 Or txt_xEntrada16.Value = 3 Or txt_xEntrada16.Value = 4 Then
    txt_xEntrada16 = "00:00"
     Me.txt_xEntrada16.MaxLength = 4
End If
End Sub
Private Sub txt_xSalida16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_xSalida16.Text <> "1" And txt_xSalida16.Text <> "2" And txt_xSalida16.Text <> "3" And txt_xSalida16.Text <> "4" And txt_xSalida16.Text <> "0" Then
    Select Case Len(txt_xSalida16.Value)
        Case 1
        txt_xSalida16.Value = txt_xSalida16.Value & ":"
        Me.txt_xSalida16.MaxLength = 4
          End Select
        
    End If
If txt_xSalida16.Value > 9 And txt_xSalida16.Value < 24 Then
    Select Case Len(txt_xSalida16.Value)
        Case 2
        txt_xSalida16.Value = txt_xSalida16.Value & ":"
        Me.txt_xSalida16.MaxLength = 5
        End Select
End If
If txt_xSalida16.Value > 23 And txt_xSalida16.Value < 30 Or txt_xSalida16.Value = 0 Or txt_xSalida16.Value = 3 Or txt_xSalida16.Value = 4 Then
    txt_xSalida16 = "00:00"
     Me.txt_xSalida16.MaxLength = 4
End If
End Sub
Private Sub txt_xEntrada1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada1, KeyAscii)
End Sub
Private Sub txt_xSalida1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida1, KeyAscii)
End Sub
Private Sub txt_xEntrada2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada2, KeyAscii)
End Sub
Private Sub txt_xSalida2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida2, KeyAscii)
End Sub
Private Sub txt_xEntrada3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada3, KeyAscii)
End Sub
Private Sub txt_xSalida3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida3, KeyAscii)
End Sub
Private Sub txt_xEntrada4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada4, KeyAscii)
End Sub
Private Sub txt_xSalida4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida4, KeyAscii)
End Sub
Private Sub txt_xEntrada5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada5, KeyAscii)
End Sub
Private Sub txt_xSalida5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida5, KeyAscii)
End Sub
Private Sub txt_xEntrada6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada6, KeyAscii)
End Sub
Private Sub txt_xSalida6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida6, KeyAscii)
End Sub
Private Sub txt_xEntrada7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada7, KeyAscii)
End Sub
Private Sub txt_xSalida7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida7, KeyAscii)
End Sub
Private Sub txt_xEntrada8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada8, KeyAscii)
End Sub
Private Sub txt_xSalida8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida8, KeyAscii)
End Sub
Private Sub txt_xEntrada9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada9, KeyAscii)
End Sub
Private Sub txt_xSalida9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida9, KeyAscii)
End Sub
Private Sub txt_xEntrada10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada10, KeyAscii)
End Sub
Private Sub txt_xSalida10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida10, KeyAscii)
End Sub
Private Sub txt_xEntrada11_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada11, KeyAscii)
End Sub
Private Sub txt_xSalida11_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida11, KeyAscii)
End Sub
Private Sub txt_xEntrada12_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada12, KeyAscii)
End Sub
Private Sub txt_xSalida12_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida12, KeyAscii)
End Sub
Private Sub txt_xEntrada13_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada13, KeyAscii)
End Sub
Private Sub txt_xSalida13_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida13, KeyAscii)
End Sub
Private Sub txt_xEntrada14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada14, KeyAscii)
End Sub
Private Sub txt_xSalida14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida14, KeyAscii)
End Sub
Private Sub txt_xEntrada15_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada15, KeyAscii)
End Sub
Private Sub txt_xSalida15_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida15, KeyAscii)
End Sub
Private Sub txt_xEntrada16_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada16, KeyAscii)
End Sub
Private Sub txt_xSalida16_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida16, KeyAscii)
End Sub
Private Sub txt_xEntrada1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada1, KeyCode)
End Sub
Private Sub txt_xSalida1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida1, KeyCode)
End Sub
Private Sub txt_xEntrada2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada2, KeyCode)
End Sub
Private Sub txt_xSalida2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida2, KeyCode)
End Sub
Private Sub txt_xEntrada3_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada3, KeyCode)
End Sub
Private Sub txt_xSalida3_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida3, KeyCode)
End Sub
Private Sub txt_xEntrada4_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada4, KeyCode)
End Sub
Private Sub txt_xSalida4_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida4, KeyCode)
End Sub
Private Sub txt_xEntrada5_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada5, KeyCode)
End Sub
Private Sub txt_xSalida5_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida5, KeyCode)
End Sub
Private Sub txt_xEntrada6_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada6, KeyCode)
End Sub
Private Sub txt_xSalida6_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida6, KeyCode)
End Sub
Private Sub txt_xEntrada7_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada7, KeyCode)
End Sub
Private Sub txt_xSalida7_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida7, KeyCode)
End Sub
Private Sub txt_xEntrada8_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada8, KeyCode)
End Sub
Private Sub txt_xSalida8_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida8, KeyCode)
End Sub
Private Sub txt_xEntrada9_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada9, KeyCode)
End Sub
Private Sub txt_xSalida9_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida9, KeyCode)
End Sub
Private Sub txt_xEntrada10_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada10, KeyCode)
End Sub
Private Sub txt_xSalida10_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida10, KeyCode)
End Sub
Private Sub txt_xEntrada11_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada11, KeyCode)
End Sub
Private Sub txt_xSalida11_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida11, KeyCode)
End Sub
Private Sub txt_xEntrada12_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada12, KeyCode)
End Sub
Private Sub txt_xSalida12_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida12, KeyCode)
End Sub
Private Sub txt_xEntrada13_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada13, KeyCode)
End Sub
Private Sub txt_xSalida13_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida13, KeyCode)
End Sub
Private Sub txt_xEntrada14_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada14, KeyCode)
End Sub
Private Sub txt_xSalida14_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida14, KeyCode)
End Sub
Private Sub txt_xEntrada15_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada15, KeyCode)
End Sub
Private Sub txt_xSalida15_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida15, KeyCode)
End Sub
Private Sub txt_xEntrada16_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada16, KeyCode)
End Sub
Private Sub txt_xSalida16_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida16, KeyCode)
End Sub
