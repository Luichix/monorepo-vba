VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Hora_Marca 
   Caption         =   "CONSULTA DE LABORES"
   ClientHeight    =   10950
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   16360
   OleObjectBlob   =   "frm_Hora_Marca.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Hora_Marca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Limpiar_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name Like "txt_fecha" & "*" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name Like "txt_xEntrada" & "*" Or Ctrl.Name Like "txt_xSalida" & "*" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl


End Sub

Private Sub btn_Rango_Click()
banderaPeriodo = 9
  Call LanzarPeriodo(Me, "btn_Rango")
End Sub
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
Limpiar_Filtro

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
 Orden_Filtro
 Hoja2.Protect (Seguridad)
 
 
         MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
             
End Sub
Private Sub LimpiarHora()

Dim Ctrl As Control
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

                        If Me.txt_xEntrada1 <> Formato Or Me.txt_xSalida1 <> Formato Then
                        If LEntrada1 >= LSalida1 Then
                            Me.txt_xEntrada1.BackColor = &HC0C0FF
                            Me.txt_xSalida1.BackColor = &HC0C0FF
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada1.BackColor = &HFFFFFF
                            Me.txt_xSalida1.BackColor = &HFFFFFF
                            Me.txt_xEntrada1.SetFocus
                            Exit Sub
                        End If
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
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
                            MsgBox "Ingrese los datos correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada16.BackColor = &HFFFFFF
                            Me.txt_xSalida16.BackColor = &HFFFFFF
                            Me.txt_xEntrada16.SetFocus
                            Exit Sub
                        End If
                End If
                        

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.Cursor = xlWait
    Registrar_Hora
    UserForm_Initialize
Application.Cursor = xlDefault
                    
     
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Recursos Humanos"
    Application.Cursor = xlDefault
 End If


End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub



Private Sub CommandButton1_Click()
banderaPersonal = 5
Call LanzarListadoPersonal(Me, "label4")
Dim Entrada As String
Dim Salida As String
Dim Vacaciones As String
Dim Ausencia As String
Dim Feriado As String
Dim Nada As Integer
Dim Dias(37) As Variant
Dim Hora(37) As Variant

Me.SpinButton2.Value = Hoja58.Cells(2, 11)
'Me.SpinButton1.Value = Hoja58.Cells(3, 11)

Me.txt_id.Text = Hoja58.Cells(6, 11)
Me.txt_nombre.Text = Hoja58.Cells(6, 12)

'''''''''''''''''''''''''''''''''''''
Me.label_año1.Caption = "AÑO"
Me.label_año2.Caption = Hoja58.Cells(2, 11)


'''''''''''''''''''''''''''''''''''''
Me.Label1.Caption = Hoja58.Cells(1, 2)
Me.Label2.Caption = Hoja58.Cells(1, 3)
Me.Label3.Caption = Hoja58.Cells(1, 4)
Me.Label4.Caption = Hoja58.Cells(1, 5)
Me.Label5.Caption = Hoja58.Cells(1, 6)
Me.Label6.Caption = Hoja58.Cells(1, 7)
Me.Label7.Caption = Hoja58.Cells(1, 8)


''''''''''''''''''''''''''''''''''''''
'Me.Label8.Caption = " " & Format(Hoja58.Cells(4, 2), "dd;@")
'Me.Label9.Caption = " " & Format(Hoja58.Cells(4, 3), "dd;@")
'Me.Label10.Caption = " " & Format(Hoja58.Cells(4, 4), "dd;@")
'Me.Label11.Caption = " " & Format(Hoja58.Cells(4, 5), "dd;@")
'Me.Label12.Caption = " " & Format(Hoja58.Cells(4, 6), "dd;@")
'Me.Label13.Caption = " " & Format(Hoja58.Cells(4, 7), "dd;@")
'Me.Label14.Caption = " " & Format(Hoja58.Cells(4, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label15.Caption = " " & Format(Hoja58.Cells(6, 2), "dd;@")
'Me.Label16.Caption = " " & Format(Hoja58.Cells(6, 3), "dd;@")
'Me.Label17.Caption = " " & Format(Hoja58.Cells(6, 4), "dd;@")
'Me.Label18.Caption = " " & Format(Hoja58.Cells(6, 5), "dd;@")
'Me.Label19.Caption = " " & Format(Hoja58.Cells(6, 6), "dd;@")
'Me.Label20.Caption = " " & Format(Hoja58.Cells(6, 7), "dd;@")
'Me.Label21.Caption = " " & Format(Hoja58.Cells(6, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label22.Caption = " " & Format(Hoja58.Cells(8, 2), "dd;@")
'Me.Label23.Caption = " " & Format(Hoja58.Cells(8, 3), "dd;@")
'Me.Label24.Caption = " " & Format(Hoja58.Cells(8, 4), "dd;@")
'Me.Label25.Caption = " " & Format(Hoja58.Cells(8, 5), "dd;@")
'Me.Label26.Caption = " " & Format(Hoja58.Cells(8, 6), "dd;@")
'Me.Label27.Caption = " " & Format(Hoja58.Cells(8, 7), "dd;@")
'Me.Label28.Caption = " " & Format(Hoja58.Cells(8, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label29.Caption = " " & Format(Hoja58.Cells(10, 2), "dd;@")
'Me.Label30.Caption = " " & Format(Hoja58.Cells(10, 3), "dd;@")
'Me.Label31.Caption = " " & Format(Hoja58.Cells(10, 4), "dd;@")
'Me.Label32.Caption = " " & Format(Hoja58.Cells(10, 5), "dd;@")
'Me.Label33.Caption = " " & Format(Hoja58.Cells(10, 6), "dd;@")
'Me.Label34.Caption = " " & Format(Hoja58.Cells(10, 7), "dd;@")
'Me.Label35.Caption = " " & Format(Hoja58.Cells(10, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label36.Caption = " " & Format(Hoja58.Cells(12, 2), "dd;@")
'Me.Label37.Caption = " " & Format(Hoja58.Cells(12, 3), "dd;@")
'Me.Label38.Caption = " " & Format(Hoja58.Cells(12, 4), "dd;@")
'Me.Label39.Caption = " " & Format(Hoja58.Cells(12, 5), "dd;@")
'Me.Label40.Caption = " " & Format(Hoja58.Cells(12, 6), "dd;@")
'Me.Label41.Caption = " " & Format(Hoja58.Cells(12, 7), "dd;@")
'Me.Label42.Caption = " " & Format(Hoja58.Cells(12, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label43.Caption = " " & Format(Hoja58.Cells(14, 2), "dd;@")
'Me.Label44.Caption = " " & Format(Hoja58.Cells(14, 3), "dd;@")


Me.TextBox1.Text = Format(Hoja58.Cells(4, 2), "dd;@")
Me.TextBox2.Text = Format(Hoja58.Cells(4, 3), "dd;@")
Me.TextBox3.Text = Format(Hoja58.Cells(4, 4), "dd;@")
Me.TextBox4.Text = Format(Hoja58.Cells(4, 5), "dd;@")
Me.TextBox5.Text = Format(Hoja58.Cells(4, 6), "dd;@")
Me.TextBox6.Text = Format(Hoja58.Cells(4, 7), "dd;@")
Me.TextBox7.Text = Format(Hoja58.Cells(4, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox8.Text = Format(Hoja58.Cells(6, 2), "dd;@")
Me.TextBox9.Text = Format(Hoja58.Cells(6, 3), "dd;@")
Me.TextBox10.Text = Format(Hoja58.Cells(6, 4), "dd;@")
Me.TextBox11.Text = Format(Hoja58.Cells(6, 5), "dd;@")
Me.TextBox12.Text = Format(Hoja58.Cells(6, 6), "dd;@")
Me.TextBox13.Text = Format(Hoja58.Cells(6, 7), "dd;@")
Me.TextBox14.Text = Format(Hoja58.Cells(6, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox15.Text = Format(Hoja58.Cells(8, 2), "dd;@")
Me.TextBox16.Text = Format(Hoja58.Cells(8, 3), "dd;@")
Me.TextBox17.Text = Format(Hoja58.Cells(8, 4), "dd;@")
Me.TextBox18.Text = Format(Hoja58.Cells(8, 5), "dd;@")
Me.TextBox19.Text = Format(Hoja58.Cells(8, 6), "dd;@")
Me.TextBox20.Text = Format(Hoja58.Cells(8, 7), "dd;@")
Me.TextBox21.Text = Format(Hoja58.Cells(8, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox22.Text = Format(Hoja58.Cells(10, 2), "dd;@")
Me.TextBox23.Text = Format(Hoja58.Cells(10, 3), "dd;@")
Me.TextBox24.Text = Format(Hoja58.Cells(10, 4), "dd;@")
Me.TextBox25.Text = Format(Hoja58.Cells(10, 5), "dd;@")
Me.TextBox26.Text = Format(Hoja58.Cells(10, 6), "dd;@")
Me.TextBox27.Text = Format(Hoja58.Cells(10, 7), "dd;@")
Me.TextBox28.Text = Format(Hoja58.Cells(10, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox29.Text = Format(Hoja58.Cells(12, 2), "dd;@")
Me.TextBox30.Text = Format(Hoja58.Cells(12, 3), "dd;@")
Me.TextBox31.Text = Format(Hoja58.Cells(12, 4), "dd;@")
Me.TextBox32.Text = Format(Hoja58.Cells(12, 5), "dd;@")
Me.TextBox33.Text = Format(Hoja58.Cells(12, 6), "dd;@")
Me.TextBox34.Text = Format(Hoja58.Cells(12, 7), "dd;@")
Me.TextBox35.Text = Format(Hoja58.Cells(12, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox36.Text = Format(Hoja58.Cells(14, 2), "dd;@")
Me.TextBox37.Text = Format(Hoja58.Cells(14, 3), "dd;@")



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox1.Text = Format(Hoja58.Cells(5, 2), "hh:mm")
'Me.TextBox2.Text = Format(Hoja58.Cells(5, 3), "hh:mm")
'Me.TextBox3.Text = Format(Hoja58.Cells(5, 4), "hh:mm")
'Me.TextBox4.Text = Format(Hoja58.Cells(5, 5), "hh:mm")
'Me.TextBox5.Text = Format(Hoja58.Cells(5, 6), "hh:mm")
'Me.TextBox6.Text = Format(Hoja58.Cells(5, 7), "hh:mm")
'Me.TextBox7.Text = Format(Hoja58.Cells(5, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox8.Text = Format(Hoja58.Cells(7, 2), "hh:mm")
'Me.TextBox9.Text = Format(Hoja58.Cells(7, 3), "hh:mm")
'Me.TextBox10.Text = Format(Hoja58.Cells(7, 4), "hh:mm")
'Me.TextBox11.Text = Format(Hoja58.Cells(7, 5), "hh:mm")
'Me.TextBox12.Text = Format(Hoja58.Cells(7, 6), "hh:mm")
'Me.TextBox13.Text = Format(Hoja58.Cells(7, 7), "hh:mm")
'Me.TextBox14.Text = Format(Hoja58.Cells(7, 8), "hh:mm")
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox15.Text = Format(Hoja58.Cells(9, 2), "hh:mm")
'Me.TextBox16.Text = Format(Hoja58.Cells(9, 3), "hh:mm")
'Me.TextBox17.Text = Format(Hoja58.Cells(9, 4), "hh:mm")
'Me.TextBox18.Text = Format(Hoja58.Cells(9, 5), "hh:mm")
'Me.TextBox19.Text = Format(Hoja58.Cells(9, 6), "hh:mm")
'Me.TextBox20.Text = Format(Hoja58.Cells(9, 7), "hh:mm")
'Me.TextBox21.Text = Format(Hoja58.Cells(9, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox22.Text = Format(Hoja58.Cells(11, 2), "hh:mm")
'Me.TextBox23.Text = Format(Hoja58.Cells(11, 3), "hh:mm")
'Me.TextBox24.Text = Format(Hoja58.Cells(11, 4), "hh:mm")
'Me.TextBox25.Text = Format(Hoja58.Cells(11, 5), "hh:mm")
'Me.TextBox26.Text = Format(Hoja58.Cells(11, 6), "hh:mm")
'Me.TextBox27.Text = Format(Hoja58.Cells(11, 7), "hh:mm")
'Me.TextBox28.Text = Format(Hoja58.Cells(11, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox29.Text = Format(Hoja58.Cells(13, 2), "hh:mm")
'Me.TextBox30.Text = Format(Hoja58.Cells(13, 3), "hh:mm")
'Me.TextBox31.Text = Format(Hoja58.Cells(13, 4), "hh:mm")
'Me.TextBox32.Text = Format(Hoja58.Cells(13, 5), "hh:mm")
'Me.TextBox33.Text = Format(Hoja58.Cells(13, 6), "hh:mm")
'Me.TextBox34.Text = Format(Hoja58.Cells(13, 7), "hh:mm")
'Me.TextBox35.Text = Format(Hoja58.Cells(13, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox36.Text = Format(Hoja58.Cells(15, 2), "hh:mm")
'Me.TextBox37.Text = Format(Hoja58.Cells(15, 3), "hh:mm")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Hora(1) = Hoja58.Cells(5, 2)
Hora(2) = Hoja58.Cells(5, 3)
Hora(3) = Hoja58.Cells(5, 4)
Hora(4) = Hoja58.Cells(5, 5)
Hora(5) = Hoja58.Cells(5, 6)
Hora(6) = Hoja58.Cells(5, 7)
Hora(7) = Hoja58.Cells(5, 8)
Hora(8) = Hoja58.Cells(7, 2)
Hora(9) = Hoja58.Cells(7, 3)
Hora(10) = Hoja58.Cells(7, 4)
Hora(11) = Hoja58.Cells(7, 5)
Hora(12) = Hoja58.Cells(7, 6)
Hora(13) = Hoja58.Cells(7, 7)
Hora(14) = Hoja58.Cells(7, 8)
Hora(15) = Hoja58.Cells(9, 2)
Hora(16) = Hoja58.Cells(9, 3)
Hora(17) = Hoja58.Cells(9, 4)
Hora(18) = Hoja58.Cells(9, 5)
Hora(19) = Hoja58.Cells(9, 6)
Hora(20) = Hoja58.Cells(9, 7)
Hora(21) = Hoja58.Cells(9, 8)
Hora(22) = Hoja58.Cells(11, 2)
Hora(23) = Hoja58.Cells(11, 3)
Hora(24) = Hoja58.Cells(11, 4)
Hora(25) = Hoja58.Cells(11, 5)
Hora(26) = Hoja58.Cells(11, 6)
Hora(27) = Hoja58.Cells(11, 7)
Hora(28) = Hoja58.Cells(11, 8)
Hora(29) = Hoja58.Cells(13, 2)
Hora(30) = Hoja58.Cells(13, 3)
Hora(31) = Hoja58.Cells(13, 4)
Hora(32) = Hoja58.Cells(13, 5)
Hora(33) = Hoja58.Cells(13, 6)
Hora(34) = Hoja58.Cells(13, 7)
Hora(35) = Hoja58.Cells(13, 8)
Hora(36) = Hoja58.Cells(15, 2)
Hora(37) = Hoja58.Cells(15, 3)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dias(1) = Hoja58.Cells(5, 2)
Dias(2) = Hoja58.Cells(5, 3)
Dias(3) = Hoja58.Cells(5, 4)
Dias(4) = Hoja58.Cells(5, 5)
Dias(5) = Hoja58.Cells(5, 6)
Dias(6) = Hoja58.Cells(5, 7)
Dias(7) = Hoja58.Cells(5, 8)
Dias(8) = Hoja58.Cells(7, 2)
Dias(9) = Hoja58.Cells(7, 3)
Dias(10) = Hoja58.Cells(7, 4)
Dias(11) = Hoja58.Cells(7, 5)
Dias(12) = Hoja58.Cells(7, 6)
Dias(13) = Hoja58.Cells(7, 7)
Dias(14) = Hoja58.Cells(7, 8)
Dias(15) = Hoja58.Cells(9, 2)
Dias(16) = Hoja58.Cells(9, 3)
Dias(17) = Hoja58.Cells(9, 4)
Dias(18) = Hoja58.Cells(9, 5)
Dias(19) = Hoja58.Cells(9, 6)
Dias(20) = Hoja58.Cells(9, 7)
Dias(21) = Hoja58.Cells(9, 8)
Dias(22) = Hoja58.Cells(11, 2)
Dias(23) = Hoja58.Cells(11, 3)
Dias(24) = Hoja58.Cells(11, 4)
Dias(25) = Hoja58.Cells(11, 5)
Dias(26) = Hoja58.Cells(11, 6)
Dias(27) = Hoja58.Cells(11, 7)
Dias(28) = Hoja58.Cells(11, 8)
Dias(29) = Hoja58.Cells(13, 2)
Dias(30) = Hoja58.Cells(13, 3)
Dias(31) = Hoja58.Cells(13, 4)
Dias(32) = Hoja58.Cells(13, 5)
Dias(33) = Hoja58.Cells(13, 6)
Dias(34) = Hoja58.Cells(13, 7)
Dias(35) = Hoja58.Cells(13, 8)
Dias(36) = Hoja58.Cells(15, 2)
Dias(37) = Hoja58.Cells(15, 3)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Entrada = "PENDIENTE HORA DE ENTRADA"
Salida = "PENDIENTE HORA DE SALIDA"
Vacaciones = "REVISAR REGISTRO"
Ausencia = "AUSENCIA"
Feriado = "DIA FERIADO"
Nada = 0

   
    If Dias(1) = Entrada Or Dias(1) = Salida Then
        TextBox1.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(1) = Vacaciones Then
        TextBox1.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(1) = Ausencia Then
        TextBox1.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(1) = Nada Then
        TextBox1.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(1) = Feriado Then
        TextBox1.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(1) >= Nada And Dias(1) <= 1 Then
        TextBox1.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox1.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox2 = Entrada Or TextBox2 = Salida Then
        TextBox2.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(2) = Vacaciones Then
        TextBox2.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(2) = Ausencia Then
        TextBox2.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(2) = Nada Then
        TextBox2.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(2) = Feriado Then
        TextBox2.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(2) >= Nada And Dias(2) <= 1 Then
        TextBox2.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox2.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox3 = Entrada Or TextBox3 = Salida Then
        TextBox3.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(3) = Vacaciones Then
        TextBox3.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(3) = Ausencia Then
        TextBox3.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(3) = Nada Then
        TextBox3.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(3) = Feriado Then
        TextBox3.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(3) >= Nada And Dias(3) <= 1 Then
        TextBox3.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox3.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox4 = Entrada Or TextBox4 = Salida Then
        TextBox4.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(4) = Vacaciones Then
        TextBox4.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(4) = Ausencia Then
        TextBox4.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(4) = Nada Then
        TextBox4.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(4) = Feriado Then
        TextBox4.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(4) >= Nada And Dias(4) <= 1 Then
        TextBox4.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox4.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox5 = Entrada Or TextBox5 = Salida Then
        TextBox5.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(5) = Vacaciones Then
        TextBox5.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(5) = Ausencia Then
        TextBox5.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(5) = Nada Then
        TextBox5.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(5) = Feriado Then
        TextBox5.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(5) >= Nada And Dias(5) <= 1 Then
        TextBox5.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox5.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox6 = Entrada Or TextBox6 = Salida Then
        TextBox6.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(6) = Vacaciones Then
        TextBox6.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(6) = Ausencia Then
        TextBox6.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(6) = Nada Then
        TextBox6.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(6) = Feriado Then
        TextBox6.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(6) >= Nada And Dias(6) <= 1 Then
        TextBox6.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox6.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox7 = Entrada Or TextBox7 = Salida Then
        TextBox7.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(7) = Vacaciones Then
        TextBox7.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(7) = Ausencia Then
        TextBox7.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(7) = Nada Then
        TextBox7.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(7) = Feriado Then
        TextBox7.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(7) >= Nada And Dias(7) <= 1 Then
        TextBox7.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox7.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox8 = Entrada Or TextBox8 = Salida Then
        TextBox8.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(8) = Vacaciones Then
        TextBox8.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(8) = Ausencia Then
        TextBox8.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(8) = Nada Then
        TextBox8.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(8) = Feriado Then
        TextBox8.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(8) >= Nada And Dias(8) <= 1 Then
        TextBox8.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox8.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox9 = Entrada Or TextBox9 = Salida Then
        TextBox9.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(9) = Vacaciones Then
        TextBox9.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(9) = Ausencia Then
        TextBox9.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(9) = Nada Then
        TextBox9.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(9) = Feriado Then
        TextBox9.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(9) >= Nada And Dias(9) <= 1 Then
        TextBox9.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox9.BackColor = &HDBDB86                              'Azul
    End If
  If TextBox10 = Entrada Or TextBox10 = Salida Then
        TextBox10.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(10) = Vacaciones Then
        TextBox10.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(10) = Ausencia Then
        TextBox10.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(10) = Nada Then
        TextBox10.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(10) = Feriado Then
        TextBox10.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(10) >= Nada And Dias(10) <= 1 Then
        TextBox10.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox10.BackColor = &HDBDB86                              'Azul
    End If
  If TextBox11 = Entrada Or TextBox11 = Salida Then
        TextBox11.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(11) = Vacaciones Then
        TextBox11.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(11) = Ausencia Then
        TextBox11.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(11) = Nada Then
        TextBox11.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(11) = Feriado Then
        TextBox11.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(11) >= Nada And Dias(11) <= 1 Then
        TextBox11.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox11.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox12 = Entrada Or TextBox12 = Salida Then
        TextBox12.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(12) = Vacaciones Then
        TextBox12.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(12) = Ausencia Then
        TextBox12.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(12) = Nada Then
        TextBox12.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(12) = Feriado Then
        TextBox12.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(12) >= Nada And Dias(12) <= 1 Then
        TextBox12.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox12.BackColor = &HDBDB86                              'Azul
    End If
  If TextBox13 = Entrada Or TextBox13 = Salida Then
        TextBox13.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(13) = Vacaciones Then
        TextBox13.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(13) = Ausencia Then
        TextBox13.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(13) = Nada Then
        TextBox13.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(13) = Feriado Then
        TextBox13.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(13) >= Nada And Dias(13) <= 1 Then
        TextBox13.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox13.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox14 = Entrada Or TextBox14 = Salida Then
        TextBox14.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(14) = Vacaciones Then
        TextBox14.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(14) = Ausencia Then
        TextBox14.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(14) = Nada Then
        TextBox14.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(14) = Feriado Then
        TextBox14.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(14) >= Nada And Dias(14) <= 1 Then
        TextBox14.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox14.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox15 = Entrada Or TextBox15 = Salida Then
        TextBox15.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(15) = Vacaciones Then
        TextBox15.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(15) = Ausencia Then
        TextBox15.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(15) = Nada Then
        TextBox15.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(15) = Feriado Then
        TextBox15.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(15) >= Nada And Dias(15) <= 1 Then
        TextBox15.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox15.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox16 = Entrada Or TextBox16 = Salida Then
        TextBox16.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(16) = Vacaciones Then
        TextBox16.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(16) = Ausencia Then
        TextBox16.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(16) = Nada Then
        TextBox16.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(16) = Feriado Then
        TextBox16.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(16) >= Nada And Dias(16) <= 1 Then
        TextBox16.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox16.BackColor = &HDBDB86                              'Azul
    End If
   If TextBox17 = Entrada Or TextBox17 = Salida Then
        TextBox17.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(17) = Vacaciones Then
        TextBox17.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(17) = Ausencia Then
        TextBox17.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(17) = Nada Then
        TextBox17.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(17) = Feriado Then
        TextBox17.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(17) >= Nada And Dias(17) <= 1 Then
        TextBox17.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox17.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox18 = Entrada Or TextBox18 = Salida Then
        TextBox18.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(18) = Vacaciones Then
        TextBox18.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(18) = Ausencia Then
        TextBox18.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(18) = Nada Then
        TextBox18.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(18) = Feriado Then
        TextBox18.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(18) >= Nada And Dias(18) <= 1 Then
        TextBox18.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox18.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox19 = Entrada Or TextBox19 = Salida Then
        TextBox19.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(19) = Vacaciones Then
        TextBox19.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(19) = Ausencia Then
        TextBox19.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(19) = Nada Then
        TextBox19.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(19) = Feriado Then
        TextBox19.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(19) >= Nada And Dias(19) <= 1 Then
        TextBox19.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox19.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox20 = Entrada Or TextBox20 = Salida Then
        TextBox20.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(20) = Vacaciones Then
        TextBox20.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(20) = Ausencia Then
        TextBox20.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(20) = Nada Then
        TextBox20.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(20) = Feriado Then
        TextBox20.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(20) >= Nada And Dias(20) <= 1 Then
        TextBox20.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox20.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox21 = Entrada Or TextBox21 = Salida Then
        TextBox21.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(21) = Vacaciones Then
        TextBox21.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(21) = Ausencia Then
        TextBox21.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(21) = Nada Then
        TextBox21.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(21) = Feriado Then
        TextBox21.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(21) >= Nada And Dias(21) <= 1 Then
        TextBox21.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox21.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox22 = Entrada Or TextBox22 = Salida Then
        TextBox22.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(22) = Vacaciones Then
        TextBox22.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(22) = Ausencia Then
        TextBox22.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(22) = Nada Then
        TextBox22.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(22) = Feriado Then
        TextBox22.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(22) >= Nada And Dias(22) <= 1 Then
        TextBox22.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox22.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox23 = Entrada Or TextBox23 = Salida Then
        TextBox23.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(23) = Vacaciones Then
        TextBox23.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(23) = Ausencia Then
        TextBox23.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(23) = Nada Then
        TextBox23.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(23) = Feriado Then
        TextBox23.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(23) >= Nada And Dias(23) <= 1 Then
        TextBox23.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox23.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox24 = Entrada Or TextBox24 = Salida Then
        TextBox24.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(24) = Vacaciones Then
        TextBox24.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(24) = Ausencia Then
        TextBox24.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(24) = Nada Then
        TextBox24.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(24) = Feriado Then
        TextBox24.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(24) >= Nada And Dias(24) <= 1 Then
        TextBox24.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox24.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox25 = Entrada Or TextBox25 = Salida Then
        TextBox25.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(25) = Vacaciones Then
        TextBox25.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(25) = Ausencia Then
        TextBox25.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(25) = Nada Then
        TextBox25.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(25) = Feriado Then
        TextBox25.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(25) >= Nada And Dias(25) <= 1 Then
        TextBox25.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox25.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox26 = Entrada Or TextBox26 = Salida Then
        TextBox26.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(26) = Vacaciones Then
        TextBox26.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(26) = Ausencia Then
        TextBox26.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(26) = Nada Then
        TextBox26.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(26) = Feriado Then
        TextBox26.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(26) >= Nada And Dias(26) <= 1 Then
        TextBox26.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox26.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox27 = Entrada Or TextBox27 = Salida Then
        TextBox27.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(27) = Vacaciones Then
        TextBox27.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(27) = Ausencia Then
        TextBox27.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(27) = Nada Then
        TextBox27.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(27) = Feriado Then
        TextBox27.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(27) >= Nada And Dias(27) <= 1 Then
        TextBox27.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox27.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox28 = Entrada Or TextBox28 = Salida Then
        TextBox28.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(28) = Vacaciones Then
        TextBox28.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(28) = Ausencia Then
        TextBox28.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(28) = Nada Then
        TextBox28.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(28) = Feriado Then
        TextBox28.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(28) >= Nada And Dias(28) <= 1 Then
        TextBox28.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox28.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox29 = Entrada Or TextBox29 = Salida Then
        TextBox29.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(29) = Vacaciones Then
        TextBox29.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(29) = Ausencia Then
        TextBox29.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(29) = Nada Then
        TextBox29.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(29) = Feriado Then
        TextBox29.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(29) >= Nada And Dias(29) <= 1 Then
        TextBox29.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox29.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox30 = Entrada Or TextBox30 = Salida Then
        TextBox30.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(30) = Vacaciones Then
        TextBox30.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(30) = Ausencia Then
        TextBox30.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(30) = Nada Then
        TextBox30.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(30) = Feriado Then
        TextBox30.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(30) >= Nada And Dias(30) <= 1 Then
        TextBox30.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox30.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox31 = Entrada Or TextBox31 = Salida Then
        TextBox31.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(31) = Vacaciones Then
        TextBox31.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(31) = Ausencia Then
        TextBox31.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(31) = Nada Then
        TextBox31.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(31) = Feriado Then
        TextBox31.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(31) >= Nada And Dias(31) <= 1 Then
        TextBox31.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox31.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox32 = Entrada Or TextBox32 = Salida Then
        TextBox32.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(32) = Vacaciones Then
        TextBox32.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(32) = Ausencia Then
        TextBox32.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(32) = Nada Then
        TextBox32.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(32) = Feriado Then
        TextBox32.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(32) >= Nada And Dias(32) <= 1 Then
        TextBox32.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox32.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox33 = Entrada Or TextBox33 = Salida Then
        TextBox33.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(33) = Vacaciones Then
        TextBox33.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(33) = Ausencia Then
        TextBox33.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(33) = Nada Then
        TextBox33.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(33) = Feriado Then
        TextBox33.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(33) >= Nada And Dias(33) <= 1 Then
        TextBox33.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox33.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox34 = Entrada Or TextBox34 = Salida Then
        TextBox34.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(34) = Vacaciones Then
        TextBox34.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(34) = Ausencia Then
        TextBox34.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(34) = Nada Then
        TextBox34.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(34) = Feriado Then
        TextBox34.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(34) >= Nada And Dias(34) <= 1 Then
        TextBox34.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox34.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox35 = Entrada Or TextBox35 = Salida Then
        TextBox35.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(35) = Vacaciones Then
        TextBox35.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(35) = Ausencia Then
        TextBox35.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(35) = Nada Then
        TextBox35.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(35) = Feriado Then
        TextBox35.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(35) >= Nada And Dias(35) <= 1 Then
        TextBox35.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox35.BackColor = &HDBDB86                              'Azul
    End If
        If TextBox36 = Entrada Or TextBox36 = Salida Then
        TextBox36.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(36) = Vacaciones Then
        TextBox36.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(36) = Ausencia Then
        TextBox36.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(36) = Nada Then
        TextBox36.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(36) = Feriado Then
        TextBox36.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(36) >= Nada And Dias(36) <= 1 Then
        TextBox36.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox36.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox37 = Entrada Or TextBox37 = Salida Then
        TextBox37.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(37) = Vacaciones Then
        TextBox37.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(37) = Ausencia Then
        TextBox37.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(37) = Nada Then
        TextBox37.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(37) = Feriado Then
        TextBox37.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(37) >= Nada And Dias(37) <= 1 Then
        TextBox37.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox37.BackColor = &HDBDB86                              'Azul
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Fila As Long
Dim Final As Long
Dim Estado As String
Dim uf As Long
Dim STRG As String
Dim X As Long

On Error Resume Next

Estado = Me.txt_id.Text

Me.lbx_Hora.ColumnCount = 5
Me.lbx_Hora.ColumnWidths = "80 pt;50 pt;60 pt;50 pt"
Me.lbx_Hora.RowSource = "Tbl_Tiempo"

uf = Hoja2.Range("A" & Rows.Count).End(xlUp).Row

Hoja2.AutoFilterMode = False
Me.lbx_Hora = Empty
Me.lbx_Hora.RowSource = Empty

For Fila = 2 To uf
    STRG = Hoja2.Cells(Fila, 2).Value 'Variable para descripción

    If UCase(STRG) Like Estado Then
        Me.lbx_Hora.AddItem
        Me.lbx_Hora.List(X, 0) = Hoja2.Cells(Fila, 1).Value
        Me.lbx_Hora.List(X, 1) = Format(Hoja2.Cells(Fila, 5), "hh:mm")
        Me.lbx_Hora.List(X, 2) = Format(Hoja2.Cells(Fila, 6), "hh:mm")
        Me.lbx_Hora.List(X, 3) = Hoja2.Cells(Fila, 14).Text
        Me.lbx_Hora.List(X, 4) = Hoja2.Cells(Fila, 15).Text
        X = X + 1
   End If
Next

Me.lbx_Hora.ColumnCount = 5
Me.lbx_Hora.ColumnWidths = "80 pt;50 pt;60 pt;50 pt"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
End Sub




Private Sub cboMes_Change()

Dim Entrada As String
Dim Salida As String
Dim Vacaciones As String
Dim Ausencia As String
Dim Feriado As String
Dim Nada As Integer
Dim Dias(37) As Variant
Dim Hora(37) As Variant


Hoja58.Cells(3, 11) = Me.cboMes.ListIndex + 1

Me.txt_id.Text = Hoja58.Cells(6, 11)
Me.txt_nombre.Text = Hoja58.Cells(6, 12)


'''''''''''''''''''''''''''''''''''''
Me.Label1.Caption = Hoja58.Cells(1, 2)
Me.Label2.Caption = Hoja58.Cells(1, 3)
Me.Label3.Caption = Hoja58.Cells(1, 4)
Me.Label4.Caption = Hoja58.Cells(1, 5)
Me.Label5.Caption = Hoja58.Cells(1, 6)
Me.Label6.Caption = Hoja58.Cells(1, 7)
Me.Label7.Caption = Hoja58.Cells(1, 8)


Me.TextBox1.Text = Format(Hoja58.Cells(4, 2), "dd;@")
Me.TextBox2.Text = Format(Hoja58.Cells(4, 3), "dd;@")
Me.TextBox3.Text = Format(Hoja58.Cells(4, 4), "dd;@")
Me.TextBox4.Text = Format(Hoja58.Cells(4, 5), "dd;@")
Me.TextBox5.Text = Format(Hoja58.Cells(4, 6), "dd;@")
Me.TextBox6.Text = Format(Hoja58.Cells(4, 7), "dd;@")
Me.TextBox7.Text = Format(Hoja58.Cells(4, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox8.Text = Format(Hoja58.Cells(6, 2), "dd;@")
Me.TextBox9.Text = Format(Hoja58.Cells(6, 3), "dd;@")
Me.TextBox10.Text = Format(Hoja58.Cells(6, 4), "dd;@")
Me.TextBox11.Text = Format(Hoja58.Cells(6, 5), "dd;@")
Me.TextBox12.Text = Format(Hoja58.Cells(6, 6), "dd;@")
Me.TextBox13.Text = Format(Hoja58.Cells(6, 7), "dd;@")
Me.TextBox14.Text = Format(Hoja58.Cells(6, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox15.Text = Format(Hoja58.Cells(8, 2), "dd;@")
Me.TextBox16.Text = Format(Hoja58.Cells(8, 3), "dd;@")
Me.TextBox17.Text = Format(Hoja58.Cells(8, 4), "dd;@")
Me.TextBox18.Text = Format(Hoja58.Cells(8, 5), "dd;@")
Me.TextBox19.Text = Format(Hoja58.Cells(8, 6), "dd;@")
Me.TextBox20.Text = Format(Hoja58.Cells(8, 7), "dd;@")
Me.TextBox21.Text = Format(Hoja58.Cells(8, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox22.Text = Format(Hoja58.Cells(10, 2), "dd;@")
Me.TextBox23.Text = Format(Hoja58.Cells(10, 3), "dd;@")
Me.TextBox24.Text = Format(Hoja58.Cells(10, 4), "dd;@")
Me.TextBox25.Text = Format(Hoja58.Cells(10, 5), "dd;@")
Me.TextBox26.Text = Format(Hoja58.Cells(10, 6), "dd;@")
Me.TextBox27.Text = Format(Hoja58.Cells(10, 7), "dd;@")
Me.TextBox28.Text = Format(Hoja58.Cells(10, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox29.Text = Format(Hoja58.Cells(12, 2), "dd;@")
Me.TextBox30.Text = Format(Hoja58.Cells(12, 3), "dd;@")
Me.TextBox31.Text = Format(Hoja58.Cells(12, 4), "dd;@")
Me.TextBox32.Text = Format(Hoja58.Cells(12, 5), "dd;@")
Me.TextBox33.Text = Format(Hoja58.Cells(12, 6), "dd;@")
Me.TextBox34.Text = Format(Hoja58.Cells(12, 7), "dd;@")
Me.TextBox35.Text = Format(Hoja58.Cells(12, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox36.Text = Format(Hoja58.Cells(14, 2), "dd;@")
Me.TextBox37.Text = Format(Hoja58.Cells(14, 3), "dd;@")


'''''''''''''''''''''''''''''''''''''
'Me.Label8.Caption = " " & Format(Hoja58.Cells(4, 2), "dd;@")
'Me.Label9.Caption = " " & Format(Hoja58.Cells(4, 3), "dd;@")
'Me.Label10.Caption = " " & Format(Hoja58.Cells(4, 4), "dd;@")
'Me.Label11.Caption = " " & Format(Hoja58.Cells(4, 5), "dd;@")
'Me.Label12.Caption = " " & Format(Hoja58.Cells(4, 6), "dd;@")
'Me.Label13.Caption = " " & Format(Hoja58.Cells(4, 7), "dd;@")
'Me.Label14.Caption = " " & Format(Hoja58.Cells(4, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label15.Caption = " " & Format(Hoja58.Cells(6, 2), "dd;@")
'Me.Label16.Caption = " " & Format(Hoja58.Cells(6, 3), "dd;@")
'Me.Label17.Caption = " " & Format(Hoja58.Cells(6, 4), "dd;@")
'Me.Label18.Caption = " " & Format(Hoja58.Cells(6, 5), "dd;@")
'Me.Label19.Caption = " " & Format(Hoja58.Cells(6, 6), "dd;@")
'Me.Label20.Caption = " " & Format(Hoja58.Cells(6, 7), "dd;@")
'Me.Label21.Caption = " " & Format(Hoja58.Cells(6, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label22.Caption = " " & Format(Hoja58.Cells(8, 2), "dd;@")
'Me.Label23.Caption = " " & Format(Hoja58.Cells(8, 3), "dd;@")
'Me.Label24.Caption = " " & Format(Hoja58.Cells(8, 4), "dd;@")
'Me.Label25.Caption = " " & Format(Hoja58.Cells(8, 5), "dd;@")
'Me.Label26.Caption = " " & Format(Hoja58.Cells(8, 6), "dd;@")
'Me.Label27.Caption = " " & Format(Hoja58.Cells(8, 7), "dd;@")
'Me.Label28.Caption = " " & Format(Hoja58.Cells(8, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label29.Caption = " " & Format(Hoja58.Cells(10, 2), "dd;@")
'Me.Label30.Caption = " " & Format(Hoja58.Cells(10, 3), "dd;@")
'Me.Label31.Caption = " " & Format(Hoja58.Cells(10, 4), "dd;@")
'Me.Label32.Caption = " " & Format(Hoja58.Cells(10, 5), "dd;@")
'Me.Label33.Caption = " " & Format(Hoja58.Cells(10, 6), "dd;@")
'Me.Label34.Caption = " " & Format(Hoja58.Cells(10, 7), "dd;@")
'Me.Label35.Caption = " " & Format(Hoja58.Cells(10, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label36.Caption = " " & Format(Hoja58.Cells(12, 2), "dd;@")
'Me.Label37.Caption = " " & Format(Hoja58.Cells(12, 3), "dd;@")
'Me.Label38.Caption = " " & Format(Hoja58.Cells(12, 4), "dd;@")
'Me.Label39.Caption = " " & Format(Hoja58.Cells(12, 5), "dd;@")
'Me.Label40.Caption = " " & Format(Hoja58.Cells(12, 6), "dd;@")
'Me.Label41.Caption = " " & Format(Hoja58.Cells(12, 7), "dd;@")
'Me.Label42.Caption = " " & Format(Hoja58.Cells(12, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label43.Caption = " " & Format(Hoja58.Cells(14, 2), "dd;@")
'Me.Label44.Caption = " " & Format(Hoja58.Cells(14, 3), "dd;@")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox1.Text = Format(Hoja58.Cells(5, 2), "hh:mm")
'Me.TextBox2.Text = Format(Hoja58.Cells(5, 3), "hh:mm")
'Me.TextBox3.Text = Format(Hoja58.Cells(5, 4), "hh:mm")
'Me.TextBox4.Text = Format(Hoja58.Cells(5, 5), "hh:mm")
'Me.TextBox5.Text = Format(Hoja58.Cells(5, 6), "hh:mm")
'Me.TextBox6.Text = Format(Hoja58.Cells(5, 7), "hh:mm")
'Me.TextBox7.Text = Format(Hoja58.Cells(5, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox8.Text = Format(Hoja58.Cells(7, 2), "hh:mm")
'Me.TextBox9.Text = Format(Hoja58.Cells(7, 3), "hh:mm")
'Me.TextBox10.Text = Format(Hoja58.Cells(7, 4), "hh:mm")
'Me.TextBox11.Text = Format(Hoja58.Cells(7, 5), "hh:mm")
'Me.TextBox12.Text = Format(Hoja58.Cells(7, 6), "hh:mm")
'Me.TextBox13.Text = Format(Hoja58.Cells(7, 7), "hh:mm")
'Me.TextBox14.Text = Format(Hoja58.Cells(7, 8), "hh:mm")
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox15.Text = Format(Hoja58.Cells(9, 2), "hh:mm")
'Me.TextBox16.Text = Format(Hoja58.Cells(9, 3), "hh:mm")
'Me.TextBox17.Text = Format(Hoja58.Cells(9, 4), "hh:mm")
'Me.TextBox18.Text = Format(Hoja58.Cells(9, 5), "hh:mm")
'Me.TextBox19.Text = Format(Hoja58.Cells(9, 6), "hh:mm")
'Me.TextBox20.Text = Format(Hoja58.Cells(9, 7), "hh:mm")
'Me.TextBox21.Text = Format(Hoja58.Cells(9, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox22.Text = Format(Hoja58.Cells(11, 2), "hh:mm")
'Me.TextBox23.Text = Format(Hoja58.Cells(11, 3), "hh:mm")
'Me.TextBox24.Text = Format(Hoja58.Cells(11, 4), "hh:mm")
'Me.TextBox25.Text = Format(Hoja58.Cells(11, 5), "hh:mm")
'Me.TextBox26.Text = Format(Hoja58.Cells(11, 6), "hh:mm")
'Me.TextBox27.Text = Format(Hoja58.Cells(11, 7), "hh:mm")
'Me.TextBox28.Text = Format(Hoja58.Cells(11, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox29.Text = Format(Hoja58.Cells(13, 2), "hh:mm")
'Me.TextBox30.Text = Format(Hoja58.Cells(13, 3), "hh:mm")
'Me.TextBox31.Text = Format(Hoja58.Cells(13, 4), "hh:mm")
'Me.TextBox32.Text = Format(Hoja58.Cells(13, 5), "hh:mm")
'Me.TextBox33.Text = Format(Hoja58.Cells(13, 6), "hh:mm")
'Me.TextBox34.Text = Format(Hoja58.Cells(13, 7), "hh:mm")
'Me.TextBox35.Text = Format(Hoja58.Cells(13, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox36.Text = Format(Hoja58.Cells(15, 2), "hh:mm")
'Me.TextBox37.Text = Format(Hoja58.Cells(15, 3), "hh:mm")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Hora(1) = Hoja58.Cells(5, 2)
Hora(2) = Hoja58.Cells(5, 3)
Hora(3) = Hoja58.Cells(5, 4)
Hora(4) = Hoja58.Cells(5, 5)
Hora(5) = Hoja58.Cells(5, 6)
Hora(6) = Hoja58.Cells(5, 7)
Hora(7) = Hoja58.Cells(5, 8)
Hora(8) = Hoja58.Cells(7, 2)
Hora(9) = Hoja58.Cells(7, 3)
Hora(10) = Hoja58.Cells(7, 4)
Hora(11) = Hoja58.Cells(7, 5)
Hora(12) = Hoja58.Cells(7, 6)
Hora(13) = Hoja58.Cells(7, 7)
Hora(14) = Hoja58.Cells(7, 8)
Hora(15) = Hoja58.Cells(9, 2)
Hora(16) = Hoja58.Cells(9, 3)
Hora(17) = Hoja58.Cells(9, 4)
Hora(18) = Hoja58.Cells(9, 5)
Hora(19) = Hoja58.Cells(9, 6)
Hora(20) = Hoja58.Cells(9, 7)
Hora(21) = Hoja58.Cells(9, 8)
Hora(22) = Hoja58.Cells(11, 2)
Hora(23) = Hoja58.Cells(11, 3)
Hora(24) = Hoja58.Cells(11, 4)
Hora(25) = Hoja58.Cells(11, 5)
Hora(26) = Hoja58.Cells(11, 6)
Hora(27) = Hoja58.Cells(11, 7)
Hora(28) = Hoja58.Cells(11, 8)
Hora(29) = Hoja58.Cells(13, 2)
Hora(30) = Hoja58.Cells(13, 3)
Hora(31) = Hoja58.Cells(13, 4)
Hora(32) = Hoja58.Cells(13, 5)
Hora(33) = Hoja58.Cells(13, 6)
Hora(34) = Hoja58.Cells(13, 7)
Hora(35) = Hoja58.Cells(13, 8)
Hora(36) = Hoja58.Cells(15, 2)
Hora(37) = Hoja58.Cells(15, 3)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dias(1) = Hoja58.Cells(5, 2)
Dias(2) = Hoja58.Cells(5, 3)
Dias(3) = Hoja58.Cells(5, 4)
Dias(4) = Hoja58.Cells(5, 5)
Dias(5) = Hoja58.Cells(5, 6)
Dias(6) = Hoja58.Cells(5, 7)
Dias(7) = Hoja58.Cells(5, 8)
Dias(8) = Hoja58.Cells(7, 2)
Dias(9) = Hoja58.Cells(7, 3)
Dias(10) = Hoja58.Cells(7, 4)
Dias(11) = Hoja58.Cells(7, 5)
Dias(12) = Hoja58.Cells(7, 6)
Dias(13) = Hoja58.Cells(7, 7)
Dias(14) = Hoja58.Cells(7, 8)
Dias(15) = Hoja58.Cells(9, 2)
Dias(16) = Hoja58.Cells(9, 3)
Dias(17) = Hoja58.Cells(9, 4)
Dias(18) = Hoja58.Cells(9, 5)
Dias(19) = Hoja58.Cells(9, 6)
Dias(20) = Hoja58.Cells(9, 7)
Dias(21) = Hoja58.Cells(9, 8)
Dias(22) = Hoja58.Cells(11, 2)
Dias(23) = Hoja58.Cells(11, 3)
Dias(24) = Hoja58.Cells(11, 4)
Dias(25) = Hoja58.Cells(11, 5)
Dias(26) = Hoja58.Cells(11, 6)
Dias(27) = Hoja58.Cells(11, 7)
Dias(28) = Hoja58.Cells(11, 8)
Dias(29) = Hoja58.Cells(13, 2)
Dias(30) = Hoja58.Cells(13, 3)
Dias(31) = Hoja58.Cells(13, 4)
Dias(32) = Hoja58.Cells(13, 5)
Dias(33) = Hoja58.Cells(13, 6)
Dias(34) = Hoja58.Cells(13, 7)
Dias(35) = Hoja58.Cells(13, 8)
Dias(36) = Hoja58.Cells(15, 2)
Dias(37) = Hoja58.Cells(15, 3)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Entrada = "PENDIENTE HORA DE ENTRADA"
Salida = "PENDIENTE HORA DE SALIDA"
Vacaciones = "REVISAR REGISTRO"
Ausencia = "AUSENCIA"
Feriado = "DIA FERIADO"
Nada = 0

   
    If Dias(1) = Entrada Or Dias(1) = Salida Then
        TextBox1.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(1) = Vacaciones Then
        TextBox1.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(1) = Ausencia Then
        TextBox1.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(1) = Nada Then
        TextBox1.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(1) = Feriado Then
        TextBox1.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(1) >= Nada And Dias(1) <= 1 Then
        TextBox1.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox1.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox2 = Entrada Or TextBox2 = Salida Then
        TextBox2.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(2) = Vacaciones Then
        TextBox2.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(2) = Ausencia Then
        TextBox2.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(2) = Nada Then
        TextBox2.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(2) = Feriado Then
        TextBox2.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(2) >= Nada And Dias(2) <= 1 Then
        TextBox2.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox2.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox3 = Entrada Or TextBox3 = Salida Then
        TextBox3.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(3) = Vacaciones Then
        TextBox3.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(3) = Ausencia Then
        TextBox3.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(3) = Nada Then
        TextBox3.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(3) = Feriado Then
        TextBox3.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(3) >= Nada And Dias(3) <= 1 Then
        TextBox3.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox3.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox4 = Entrada Or TextBox4 = Salida Then
        TextBox4.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(4) = Vacaciones Then
        TextBox4.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(4) = Ausencia Then
        TextBox4.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(4) = Nada Then
        TextBox4.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(4) = Feriado Then
        TextBox4.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(4) >= Nada And Dias(4) <= 1 Then
        TextBox4.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox4.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox5 = Entrada Or TextBox5 = Salida Then
        TextBox5.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(5) = Vacaciones Then
        TextBox5.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(5) = Ausencia Then
        TextBox5.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(5) = Nada Then
        TextBox5.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(5) = Feriado Then
        TextBox5.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(5) >= Nada And Dias(5) <= 1 Then
        TextBox5.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox5.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox6 = Entrada Or TextBox6 = Salida Then
        TextBox6.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(6) = Vacaciones Then
        TextBox6.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(6) = Ausencia Then
        TextBox6.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(6) = Nada Then
        TextBox6.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(6) = Feriado Then
        TextBox6.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(6) >= Nada And Dias(6) <= 1 Then
        TextBox6.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox6.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox7 = Entrada Or TextBox7 = Salida Then
        TextBox7.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(7) = Vacaciones Then
        TextBox7.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(7) = Ausencia Then
        TextBox7.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(7) = Nada Then
        TextBox7.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(7) = Feriado Then
        TextBox7.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(7) >= Nada And Dias(7) <= 1 Then
        TextBox7.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox7.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox8 = Entrada Or TextBox8 = Salida Then
        TextBox8.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(8) = Vacaciones Then
        TextBox8.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(8) = Ausencia Then
        TextBox8.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(8) = Nada Then
        TextBox8.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(8) = Feriado Then
        TextBox8.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(8) >= Nada And Dias(8) <= 1 Then
        TextBox8.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox8.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox9 = Entrada Or TextBox9 = Salida Then
        TextBox9.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(9) = Vacaciones Then
        TextBox9.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(9) = Ausencia Then
        TextBox9.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(9) = Nada Then
        TextBox9.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(9) = Feriado Then
        TextBox9.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(9) >= Nada And Dias(9) <= 1 Then
        TextBox9.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox9.BackColor = &HDBDB86                              'Azul
    End If
  If TextBox10 = Entrada Or TextBox10 = Salida Then
        TextBox10.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(10) = Vacaciones Then
        TextBox10.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(10) = Ausencia Then
        TextBox10.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(10) = Nada Then
        TextBox10.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(10) = Feriado Then
        TextBox10.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(10) >= Nada And Dias(10) <= 1 Then
        TextBox10.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox10.BackColor = &HDBDB86                              'Azul
    End If
  If TextBox11 = Entrada Or TextBox11 = Salida Then
        TextBox11.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(11) = Vacaciones Then
        TextBox11.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(11) = Ausencia Then
        TextBox11.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(11) = Nada Then
        TextBox11.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(11) = Feriado Then
        TextBox11.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(11) >= Nada And Dias(11) <= 1 Then
        TextBox11.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox11.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox12 = Entrada Or TextBox12 = Salida Then
        TextBox12.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(12) = Vacaciones Then
        TextBox12.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(12) = Ausencia Then
        TextBox12.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(12) = Nada Then
        TextBox12.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(12) = Feriado Then
        TextBox12.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(12) >= Nada And Dias(12) <= 1 Then
        TextBox12.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox12.BackColor = &HDBDB86                              'Azul
    End If
  If TextBox13 = Entrada Or TextBox13 = Salida Then
        TextBox13.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(13) = Vacaciones Then
        TextBox13.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(13) = Ausencia Then
        TextBox13.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(13) = Nada Then
        TextBox13.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(13) = Feriado Then
        TextBox13.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(13) >= Nada And Dias(13) <= 1 Then
        TextBox13.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox13.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox14 = Entrada Or TextBox14 = Salida Then
        TextBox14.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(14) = Vacaciones Then
        TextBox14.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(14) = Ausencia Then
        TextBox14.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(14) = Nada Then
        TextBox14.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(14) = Feriado Then
        TextBox14.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(14) >= Nada And Dias(14) <= 1 Then
        TextBox14.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox14.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox15 = Entrada Or TextBox15 = Salida Then
        TextBox15.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(15) = Vacaciones Then
        TextBox15.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(15) = Ausencia Then
        TextBox15.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(15) = Nada Then
        TextBox15.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(15) = Feriado Then
        TextBox15.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(15) >= Nada And Dias(15) <= 1 Then
        TextBox15.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox15.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox16 = Entrada Or TextBox16 = Salida Then
        TextBox16.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(16) = Vacaciones Then
        TextBox16.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(16) = Ausencia Then
        TextBox16.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(16) = Nada Then
        TextBox16.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(16) = Feriado Then
        TextBox16.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(16) >= Nada And Dias(16) <= 1 Then
        TextBox16.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox16.BackColor = &HDBDB86                              'Azul
    End If
   If TextBox17 = Entrada Or TextBox17 = Salida Then
        TextBox17.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(17) = Vacaciones Then
        TextBox17.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(17) = Ausencia Then
        TextBox17.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(17) = Nada Then
        TextBox17.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(17) = Feriado Then
        TextBox17.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(17) >= Nada And Dias(17) <= 1 Then
        TextBox17.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox17.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox18 = Entrada Or TextBox18 = Salida Then
        TextBox18.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(18) = Vacaciones Then
        TextBox18.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(18) = Ausencia Then
        TextBox18.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(18) = Nada Then
        TextBox18.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(18) = Feriado Then
        TextBox18.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(18) >= Nada And Dias(18) <= 1 Then
        TextBox18.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox18.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox19 = Entrada Or TextBox19 = Salida Then
        TextBox19.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(19) = Vacaciones Then
        TextBox19.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(19) = Ausencia Then
        TextBox19.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(19) = Nada Then
        TextBox19.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(19) = Feriado Then
        TextBox19.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(19) >= Nada And Dias(19) <= 1 Then
        TextBox19.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox19.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox20 = Entrada Or TextBox20 = Salida Then
        TextBox20.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(20) = Vacaciones Then
        TextBox20.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(20) = Ausencia Then
        TextBox20.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(20) = Nada Then
        TextBox20.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(20) = Feriado Then
        TextBox20.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(20) >= Nada And Dias(20) <= 1 Then
        TextBox20.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox20.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox21 = Entrada Or TextBox21 = Salida Then
        TextBox21.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(21) = Vacaciones Then
        TextBox21.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(21) = Ausencia Then
        TextBox21.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(21) = Nada Then
        TextBox21.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(21) = Feriado Then
        TextBox21.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(21) >= Nada And Dias(21) <= 1 Then
        TextBox21.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox21.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox22 = Entrada Or TextBox22 = Salida Then
        TextBox22.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(22) = Vacaciones Then
        TextBox22.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(22) = Ausencia Then
        TextBox22.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(22) = Nada Then
        TextBox22.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(22) = Feriado Then
        TextBox22.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(22) >= Nada And Dias(22) <= 1 Then
        TextBox22.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox22.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox23 = Entrada Or TextBox23 = Salida Then
        TextBox23.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(23) = Vacaciones Then
        TextBox23.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(23) = Ausencia Then
        TextBox23.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(23) = Nada Then
        TextBox23.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(23) = Feriado Then
        TextBox23.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(23) >= Nada And Dias(23) <= 1 Then
        TextBox23.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox23.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox24 = Entrada Or TextBox24 = Salida Then
        TextBox24.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(24) = Vacaciones Then
        TextBox24.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(24) = Ausencia Then
        TextBox24.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(24) = Nada Then
        TextBox24.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(24) = Feriado Then
        TextBox24.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(24) >= Nada And Dias(24) <= 1 Then
        TextBox24.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox24.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox25 = Entrada Or TextBox25 = Salida Then
        TextBox25.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(25) = Vacaciones Then
        TextBox25.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(25) = Ausencia Then
        TextBox25.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(25) = Nada Then
        TextBox25.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(25) = Feriado Then
        TextBox25.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(25) >= Nada And Dias(25) <= 1 Then
        TextBox25.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox25.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox26 = Entrada Or TextBox26 = Salida Then
        TextBox26.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(26) = Vacaciones Then
        TextBox26.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(26) = Ausencia Then
        TextBox26.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(26) = Nada Then
        TextBox26.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(26) = Feriado Then
        TextBox26.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(26) >= Nada And Dias(26) <= 1 Then
        TextBox26.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox26.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox27 = Entrada Or TextBox27 = Salida Then
        TextBox27.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(27) = Vacaciones Then
        TextBox27.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(27) = Ausencia Then
        TextBox27.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(27) = Nada Then
        TextBox27.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(27) = Feriado Then
        TextBox27.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(27) >= Nada And Dias(27) <= 1 Then
        TextBox27.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox27.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox28 = Entrada Or TextBox28 = Salida Then
        TextBox28.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(28) = Vacaciones Then
        TextBox28.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(28) = Ausencia Then
        TextBox28.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(28) = Nada Then
        TextBox28.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(28) = Feriado Then
        TextBox28.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(28) >= Nada And Dias(28) <= 1 Then
        TextBox28.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox28.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox29 = Entrada Or TextBox29 = Salida Then
        TextBox29.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(29) = Vacaciones Then
        TextBox29.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(29) = Ausencia Then
        TextBox29.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(29) = Nada Then
        TextBox29.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(29) = Feriado Then
        TextBox29.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(29) >= Nada And Dias(29) <= 1 Then
        TextBox29.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox29.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox30 = Entrada Or TextBox30 = Salida Then
        TextBox30.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(30) = Vacaciones Then
        TextBox30.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(30) = Ausencia Then
        TextBox30.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(30) = Nada Then
        TextBox30.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(30) = Feriado Then
        TextBox30.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(30) >= Nada And Dias(30) <= 1 Then
        TextBox30.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox30.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox31 = Entrada Or TextBox31 = Salida Then
        TextBox31.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(31) = Vacaciones Then
        TextBox31.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(31) = Ausencia Then
        TextBox31.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(31) = Nada Then
        TextBox31.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(31) = Feriado Then
        TextBox31.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(31) >= Nada And Dias(31) <= 1 Then
        TextBox31.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox31.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox32 = Entrada Or TextBox32 = Salida Then
        TextBox32.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(32) = Vacaciones Then
        TextBox32.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(32) = Ausencia Then
        TextBox32.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(32) = Nada Then
        TextBox32.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(32) = Feriado Then
        TextBox32.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(32) >= Nada And Dias(32) <= 1 Then
        TextBox32.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox32.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox33 = Entrada Or TextBox33 = Salida Then
        TextBox33.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(33) = Vacaciones Then
        TextBox33.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(33) = Ausencia Then
        TextBox33.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(33) = Nada Then
        TextBox33.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(33) = Feriado Then
        TextBox33.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(33) >= Nada And Dias(33) <= 1 Then
        TextBox33.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox33.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox34 = Entrada Or TextBox34 = Salida Then
        TextBox34.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(34) = Vacaciones Then
        TextBox34.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(34) = Ausencia Then
        TextBox34.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(34) = Nada Then
        TextBox34.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(34) = Feriado Then
        TextBox34.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(34) >= Nada And Dias(34) <= 1 Then
        TextBox34.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox34.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox35 = Entrada Or TextBox35 = Salida Then
        TextBox35.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(35) = Vacaciones Then
        TextBox35.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(35) = Ausencia Then
        TextBox35.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(35) = Nada Then
        TextBox35.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(35) = Feriado Then
        TextBox35.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(35) >= Nada And Dias(35) <= 1 Then
        TextBox35.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox35.BackColor = &HDBDB86                              'Azul
    End If
        If TextBox36 = Entrada Or TextBox36 = Salida Then
        TextBox36.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(36) = Vacaciones Then
        TextBox36.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(36) = Ausencia Then
        TextBox36.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(36) = Nada Then
        TextBox36.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(36) = Feriado Then
        TextBox36.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(36) >= Nada And Dias(36) <= 1 Then
        TextBox36.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox36.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox37 = Entrada Or TextBox37 = Salida Then
        TextBox37.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(37) = Vacaciones Then
        TextBox37.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(37) = Ausencia Then
        TextBox37.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(37) = Nada Then
        TextBox37.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(37) = Feriado Then
        TextBox37.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(37) >= Nada And Dias(37) <= 1 Then
        TextBox37.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox37.BackColor = &HDBDB86                              'Azul
    End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub


Private Sub CommandButton3_Click()

End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub txt_nombre_Click()

End Sub



Private Sub SpinButton2_Change()
Dim Entrada As String
Dim Salida As String
Dim Vacaciones As String
Dim Ausencia As String
Dim Feriado As String
Dim Nada As Integer
Dim Dias(37) As Variant
Dim Hora(37) As Variant


Hoja58.Cells(2, 11) = Me.SpinButton2.Value
Me.txt_id.Text = Hoja58.Cells(6, 11)
Me.txt_nombre.Text = Hoja58.Cells(6, 12)

'''''''''''''''''''''''''''''''''''''
Me.label_año1.Caption = "AÑO"
Me.label_año2.Caption = Hoja58.Cells(2, 11)


'''''''''''''''''''''''''''''''''''''
Me.Label1.Caption = Hoja58.Cells(1, 2)
Me.Label2.Caption = Hoja58.Cells(1, 3)
Me.Label3.Caption = Hoja58.Cells(1, 4)
Me.Label4.Caption = Hoja58.Cells(1, 5)
Me.Label5.Caption = Hoja58.Cells(1, 6)
Me.Label6.Caption = Hoja58.Cells(1, 7)
Me.Label7.Caption = Hoja58.Cells(1, 8)

Me.TextBox1.Text = Format(Hoja58.Cells(4, 2), "dd;@")
Me.TextBox2.Text = Format(Hoja58.Cells(4, 3), "dd;@")
Me.TextBox3.Text = Format(Hoja58.Cells(4, 4), "dd;@")
Me.TextBox4.Text = Format(Hoja58.Cells(4, 5), "dd;@")
Me.TextBox5.Text = Format(Hoja58.Cells(4, 6), "dd;@")
Me.TextBox6.Text = Format(Hoja58.Cells(4, 7), "dd;@")
Me.TextBox7.Text = Format(Hoja58.Cells(4, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox8.Text = Format(Hoja58.Cells(6, 2), "dd;@")
Me.TextBox9.Text = Format(Hoja58.Cells(6, 3), "dd;@")
Me.TextBox10.Text = Format(Hoja58.Cells(6, 4), "dd;@")
Me.TextBox11.Text = Format(Hoja58.Cells(6, 5), "dd;@")
Me.TextBox12.Text = Format(Hoja58.Cells(6, 6), "dd;@")
Me.TextBox13.Text = Format(Hoja58.Cells(6, 7), "dd;@")
Me.TextBox14.Text = Format(Hoja58.Cells(6, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox15.Text = Format(Hoja58.Cells(8, 2), "dd;@")
Me.TextBox16.Text = Format(Hoja58.Cells(8, 3), "dd;@")
Me.TextBox17.Text = Format(Hoja58.Cells(8, 4), "dd;@")
Me.TextBox18.Text = Format(Hoja58.Cells(8, 5), "dd;@")
Me.TextBox19.Text = Format(Hoja58.Cells(8, 6), "dd;@")
Me.TextBox20.Text = Format(Hoja58.Cells(8, 7), "dd;@")
Me.TextBox21.Text = Format(Hoja58.Cells(8, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox22.Text = Format(Hoja58.Cells(10, 2), "dd;@")
Me.TextBox23.Text = Format(Hoja58.Cells(10, 3), "dd;@")
Me.TextBox24.Text = Format(Hoja58.Cells(10, 4), "dd;@")
Me.TextBox25.Text = Format(Hoja58.Cells(10, 5), "dd;@")
Me.TextBox26.Text = Format(Hoja58.Cells(10, 6), "dd;@")
Me.TextBox27.Text = Format(Hoja58.Cells(10, 7), "dd;@")
Me.TextBox28.Text = Format(Hoja58.Cells(10, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox29.Text = Format(Hoja58.Cells(12, 2), "dd;@")
Me.TextBox30.Text = Format(Hoja58.Cells(12, 3), "dd;@")
Me.TextBox31.Text = Format(Hoja58.Cells(12, 4), "dd;@")
Me.TextBox32.Text = Format(Hoja58.Cells(12, 5), "dd;@")
Me.TextBox33.Text = Format(Hoja58.Cells(12, 6), "dd;@")
Me.TextBox34.Text = Format(Hoja58.Cells(12, 7), "dd;@")
Me.TextBox35.Text = Format(Hoja58.Cells(12, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox36.Text = Format(Hoja58.Cells(14, 2), "dd;@")
Me.TextBox37.Text = Format(Hoja58.Cells(14, 3), "dd;@")


''''''''''''''''''''''''''''''''''''''
'Me.Label8.Caption = " " & Format(Hoja58.Cells(4, 2), "dd;@")
'Me.Label9.Caption = " " & Format(Hoja58.Cells(4, 3), "dd;@")
'Me.Label10.Caption = " " & Format(Hoja58.Cells(4, 4), "dd;@")
'Me.Label11.Caption = " " & Format(Hoja58.Cells(4, 5), "dd;@")
'Me.Label12.Caption = " " & Format(Hoja58.Cells(4, 6), "dd;@")
'Me.Label13.Caption = " " & Format(Hoja58.Cells(4, 7), "dd;@")
'Me.Label14.Caption = " " & Format(Hoja58.Cells(4, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label15.Caption = " " & Format(Hoja58.Cells(6, 2), "dd;@")
'Me.Label16.Caption = " " & Format(Hoja58.Cells(6, 3), "dd;@")
'Me.Label17.Caption = " " & Format(Hoja58.Cells(6, 4), "dd;@")
'Me.Label18.Caption = " " & Format(Hoja58.Cells(6, 5), "dd;@")
'Me.Label19.Caption = " " & Format(Hoja58.Cells(6, 6), "dd;@")
'Me.Label20.Caption = " " & Format(Hoja58.Cells(6, 7), "dd;@")
'Me.Label21.Caption = " " & Format(Hoja58.Cells(6, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label22.Caption = " " & Format(Hoja58.Cells(8, 2), "dd;@")
'Me.Label23.Caption = " " & Format(Hoja58.Cells(8, 3), "dd;@")
'Me.Label24.Caption = " " & Format(Hoja58.Cells(8, 4), "dd;@")
'Me.Label25.Caption = " " & Format(Hoja58.Cells(8, 5), "dd;@")
'Me.Label26.Caption = " " & Format(Hoja58.Cells(8, 6), "dd;@")
'Me.Label27.Caption = " " & Format(Hoja58.Cells(8, 7), "dd;@")
'Me.Label28.Caption = " " & Format(Hoja58.Cells(8, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label29.Caption = " " & Format(Hoja58.Cells(10, 2), "dd;@")
'Me.Label30.Caption = " " & Format(Hoja58.Cells(10, 3), "dd;@")
'Me.Label31.Caption = " " & Format(Hoja58.Cells(10, 4), "dd;@")
'Me.Label32.Caption = " " & Format(Hoja58.Cells(10, 5), "dd;@")
'Me.Label33.Caption = " " & Format(Hoja58.Cells(10, 6), "dd;@")
'Me.Label34.Caption = " " & Format(Hoja58.Cells(10, 7), "dd;@")
'Me.Label35.Caption = " " & Format(Hoja58.Cells(10, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label36.Caption = " " & Format(Hoja58.Cells(12, 2), "dd;@")
'Me.Label37.Caption = " " & Format(Hoja58.Cells(12, 3), "dd;@")
'Me.Label38.Caption = " " & Format(Hoja58.Cells(12, 4), "dd;@")
'Me.Label39.Caption = " " & Format(Hoja58.Cells(12, 5), "dd;@")
'Me.Label40.Caption = " " & Format(Hoja58.Cells(12, 6), "dd;@")
'Me.Label41.Caption = " " & Format(Hoja58.Cells(12, 7), "dd;@")
'Me.Label42.Caption = " " & Format(Hoja58.Cells(12, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label43.Caption = " " & Format(Hoja58.Cells(14, 2), "dd;@")
'Me.Label44.Caption = " " & Format(Hoja58.Cells(14, 3), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox1.Text = Format(Hoja58.Cells(5, 2), "hh:mm")
'Me.TextBox2.Text = Format(Hoja58.Cells(5, 3), "hh:mm")
'Me.TextBox3.Text = Format(Hoja58.Cells(5, 4), "hh:mm")
'Me.TextBox4.Text = Format(Hoja58.Cells(5, 5), "hh:mm")
'Me.TextBox5.Text = Format(Hoja58.Cells(5, 6), "hh:mm")
'Me.TextBox6.Text = Format(Hoja58.Cells(5, 7), "hh:mm")
'Me.TextBox7.Text = Format(Hoja58.Cells(5, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox8.Text = Format(Hoja58.Cells(7, 2), "hh:mm")
'Me.TextBox9.Text = Format(Hoja58.Cells(7, 3), "hh:mm")
'Me.TextBox10.Text = Format(Hoja58.Cells(7, 4), "hh:mm")
'Me.TextBox11.Text = Format(Hoja58.Cells(7, 5), "hh:mm")
'Me.TextBox12.Text = Format(Hoja58.Cells(7, 6), "hh:mm")
'Me.TextBox13.Text = Format(Hoja58.Cells(7, 7), "hh:mm")
'Me.TextBox14.Text = Format(Hoja58.Cells(7, 8), "hh:mm")
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox15.Text = Format(Hoja58.Cells(9, 2), "hh:mm")
'Me.TextBox16.Text = Format(Hoja58.Cells(9, 3), "hh:mm")
'Me.TextBox17.Text = Format(Hoja58.Cells(9, 4), "hh:mm")
'Me.TextBox18.Text = Format(Hoja58.Cells(9, 5), "hh:mm")
'Me.TextBox19.Text = Format(Hoja58.Cells(9, 6), "hh:mm")
'Me.TextBox20.Text = Format(Hoja58.Cells(9, 7), "hh:mm")
'Me.TextBox21.Text = Format(Hoja58.Cells(9, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox22.Text = Format(Hoja58.Cells(11, 2), "hh:mm")
'Me.TextBox23.Text = Format(Hoja58.Cells(11, 3), "hh:mm")
'Me.TextBox24.Text = Format(Hoja58.Cells(11, 4), "hh:mm")
'Me.TextBox25.Text = Format(Hoja58.Cells(11, 5), "hh:mm")
'Me.TextBox26.Text = Format(Hoja58.Cells(11, 6), "hh:mm")
'Me.TextBox27.Text = Format(Hoja58.Cells(11, 7), "hh:mm")
'Me.TextBox28.Text = Format(Hoja58.Cells(11, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox29.Text = Format(Hoja58.Cells(13, 2), "hh:mm")
'Me.TextBox30.Text = Format(Hoja58.Cells(13, 3), "hh:mm")
'Me.TextBox31.Text = Format(Hoja58.Cells(13, 4), "hh:mm")
'Me.TextBox32.Text = Format(Hoja58.Cells(13, 5), "hh:mm")
'Me.TextBox33.Text = Format(Hoja58.Cells(13, 6), "hh:mm")
'Me.TextBox34.Text = Format(Hoja58.Cells(13, 7), "hh:mm")
'Me.TextBox35.Text = Format(Hoja58.Cells(13, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox36.Text = Format(Hoja58.Cells(15, 2), "hh:mm")
'Me.TextBox37.Text = Format(Hoja58.Cells(15, 3), "hh:mm")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Hora(1) = Hoja58.Cells(5, 2)
Hora(2) = Hoja58.Cells(5, 3)
Hora(3) = Hoja58.Cells(5, 4)
Hora(4) = Hoja58.Cells(5, 5)
Hora(5) = Hoja58.Cells(5, 6)
Hora(6) = Hoja58.Cells(5, 7)
Hora(7) = Hoja58.Cells(5, 8)
Hora(8) = Hoja58.Cells(7, 2)
Hora(9) = Hoja58.Cells(7, 3)
Hora(10) = Hoja58.Cells(7, 4)
Hora(11) = Hoja58.Cells(7, 5)
Hora(12) = Hoja58.Cells(7, 6)
Hora(13) = Hoja58.Cells(7, 7)
Hora(14) = Hoja58.Cells(7, 8)
Hora(15) = Hoja58.Cells(9, 2)
Hora(16) = Hoja58.Cells(9, 3)
Hora(17) = Hoja58.Cells(9, 4)
Hora(18) = Hoja58.Cells(9, 5)
Hora(19) = Hoja58.Cells(9, 6)
Hora(20) = Hoja58.Cells(9, 7)
Hora(21) = Hoja58.Cells(9, 8)
Hora(22) = Hoja58.Cells(11, 2)
Hora(23) = Hoja58.Cells(11, 3)
Hora(24) = Hoja58.Cells(11, 4)
Hora(25) = Hoja58.Cells(11, 5)
Hora(26) = Hoja58.Cells(11, 6)
Hora(27) = Hoja58.Cells(11, 7)
Hora(28) = Hoja58.Cells(11, 8)
Hora(29) = Hoja58.Cells(13, 2)
Hora(30) = Hoja58.Cells(13, 3)
Hora(31) = Hoja58.Cells(13, 4)
Hora(32) = Hoja58.Cells(13, 5)
Hora(33) = Hoja58.Cells(13, 6)
Hora(34) = Hoja58.Cells(13, 7)
Hora(35) = Hoja58.Cells(13, 8)
Hora(36) = Hoja58.Cells(15, 2)
Hora(37) = Hoja58.Cells(15, 3)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dias(1) = Hoja58.Cells(5, 2)
Dias(2) = Hoja58.Cells(5, 3)
Dias(3) = Hoja58.Cells(5, 4)
Dias(4) = Hoja58.Cells(5, 5)
Dias(5) = Hoja58.Cells(5, 6)
Dias(6) = Hoja58.Cells(5, 7)
Dias(7) = Hoja58.Cells(5, 8)
Dias(8) = Hoja58.Cells(7, 2)
Dias(9) = Hoja58.Cells(7, 3)
Dias(10) = Hoja58.Cells(7, 4)
Dias(11) = Hoja58.Cells(7, 5)
Dias(12) = Hoja58.Cells(7, 6)
Dias(13) = Hoja58.Cells(7, 7)
Dias(14) = Hoja58.Cells(7, 8)
Dias(15) = Hoja58.Cells(9, 2)
Dias(16) = Hoja58.Cells(9, 3)
Dias(17) = Hoja58.Cells(9, 4)
Dias(18) = Hoja58.Cells(9, 5)
Dias(19) = Hoja58.Cells(9, 6)
Dias(20) = Hoja58.Cells(9, 7)
Dias(21) = Hoja58.Cells(9, 8)
Dias(22) = Hoja58.Cells(11, 2)
Dias(23) = Hoja58.Cells(11, 3)
Dias(24) = Hoja58.Cells(11, 4)
Dias(25) = Hoja58.Cells(11, 5)
Dias(26) = Hoja58.Cells(11, 6)
Dias(27) = Hoja58.Cells(11, 7)
Dias(28) = Hoja58.Cells(11, 8)
Dias(29) = Hoja58.Cells(13, 2)
Dias(30) = Hoja58.Cells(13, 3)
Dias(31) = Hoja58.Cells(13, 4)
Dias(32) = Hoja58.Cells(13, 5)
Dias(33) = Hoja58.Cells(13, 6)
Dias(34) = Hoja58.Cells(13, 7)
Dias(35) = Hoja58.Cells(13, 8)
Dias(36) = Hoja58.Cells(15, 2)
Dias(37) = Hoja58.Cells(15, 3)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Entrada = "PENDIENTE HORA DE ENTRADA"
Salida = "PENDIENTE HORA DE SALIDA"
Vacaciones = "REVISAR REGISTRO"
Ausencia = "AUSENCIA"
Feriado = "DIA FERIADO"
Nada = 0

   
    If Dias(1) = Entrada Or Dias(1) = Salida Then
        TextBox1.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(1) = Vacaciones Then
        TextBox1.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(1) = Ausencia Then
        TextBox1.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(1) = Nada Then
        TextBox1.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(1) = Feriado Then
        TextBox1.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(1) >= Nada And Dias(1) <= 1 Then
        TextBox1.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox1.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox2 = Entrada Or TextBox2 = Salida Then
        TextBox2.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(2) = Vacaciones Then
        TextBox2.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(2) = Ausencia Then
        TextBox2.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(2) = Nada Then
        TextBox2.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(2) = Feriado Then
        TextBox2.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(2) >= Nada And Dias(2) <= 1 Then
        TextBox2.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox2.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox3 = Entrada Or TextBox3 = Salida Then
        TextBox3.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(3) = Vacaciones Then
        TextBox3.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(3) = Ausencia Then
        TextBox3.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(3) = Nada Then
        TextBox3.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(3) = Feriado Then
        TextBox3.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(3) >= Nada And Dias(3) <= 1 Then
        TextBox3.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox3.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox4 = Entrada Or TextBox4 = Salida Then
        TextBox4.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(4) = Vacaciones Then
        TextBox4.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(4) = Ausencia Then
        TextBox4.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(4) = Nada Then
        TextBox4.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(4) = Feriado Then
        TextBox4.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(4) >= Nada And Dias(4) <= 1 Then
        TextBox4.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox4.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox5 = Entrada Or TextBox5 = Salida Then
        TextBox5.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(5) = Vacaciones Then
        TextBox5.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(5) = Ausencia Then
        TextBox5.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(5) = Nada Then
        TextBox5.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(5) = Feriado Then
        TextBox5.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(5) >= Nada And Dias(5) <= 1 Then
        TextBox5.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox5.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox6 = Entrada Or TextBox6 = Salida Then
        TextBox6.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(6) = Vacaciones Then
        TextBox6.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(6) = Ausencia Then
        TextBox6.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(6) = Nada Then
        TextBox6.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(6) = Feriado Then
        TextBox6.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(6) >= Nada And Dias(6) <= 1 Then
        TextBox6.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox6.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox7 = Entrada Or TextBox7 = Salida Then
        TextBox7.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(7) = Vacaciones Then
        TextBox7.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(7) = Ausencia Then
        TextBox7.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(7) = Nada Then
        TextBox7.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(7) = Feriado Then
        TextBox7.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(7) >= Nada And Dias(7) <= 1 Then
        TextBox7.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox7.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox8 = Entrada Or TextBox8 = Salida Then
        TextBox8.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(8) = Vacaciones Then
        TextBox8.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(8) = Ausencia Then
        TextBox8.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(8) = Nada Then
        TextBox8.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(8) = Feriado Then
        TextBox8.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(8) >= Nada And Dias(8) <= 1 Then
        TextBox8.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox8.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox9 = Entrada Or TextBox9 = Salida Then
        TextBox9.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(9) = Vacaciones Then
        TextBox9.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(9) = Ausencia Then
        TextBox9.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(9) = Nada Then
        TextBox9.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(9) = Feriado Then
        TextBox9.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(9) >= Nada And Dias(9) <= 1 Then
        TextBox9.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox9.BackColor = &HDBDB86                              'Azul
    End If
  If TextBox10 = Entrada Or TextBox10 = Salida Then
        TextBox10.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(10) = Vacaciones Then
        TextBox10.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(10) = Ausencia Then
        TextBox10.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(10) = Nada Then
        TextBox10.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(10) = Feriado Then
        TextBox10.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(10) >= Nada And Dias(10) <= 1 Then
        TextBox10.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox10.BackColor = &HDBDB86                              'Azul
    End If
  If TextBox11 = Entrada Or TextBox11 = Salida Then
        TextBox11.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(11) = Vacaciones Then
        TextBox11.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(11) = Ausencia Then
        TextBox11.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(11) = Nada Then
        TextBox11.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(11) = Feriado Then
        TextBox11.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(11) >= Nada And Dias(11) <= 1 Then
        TextBox11.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox11.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox12 = Entrada Or TextBox12 = Salida Then
        TextBox12.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(12) = Vacaciones Then
        TextBox12.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(12) = Ausencia Then
        TextBox12.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(12) = Nada Then
        TextBox12.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(12) = Feriado Then
        TextBox12.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(12) >= Nada And Dias(12) <= 1 Then
        TextBox12.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox12.BackColor = &HDBDB86                              'Azul
    End If
  If TextBox13 = Entrada Or TextBox13 = Salida Then
        TextBox13.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(13) = Vacaciones Then
        TextBox13.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(13) = Ausencia Then
        TextBox13.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(13) = Nada Then
        TextBox13.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(13) = Feriado Then
        TextBox13.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(13) >= Nada And Dias(13) <= 1 Then
        TextBox13.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox13.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox14 = Entrada Or TextBox14 = Salida Then
        TextBox14.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(14) = Vacaciones Then
        TextBox14.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(14) = Ausencia Then
        TextBox14.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(14) = Nada Then
        TextBox14.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(14) = Feriado Then
        TextBox14.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(14) >= Nada And Dias(14) <= 1 Then
        TextBox14.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox14.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox15 = Entrada Or TextBox15 = Salida Then
        TextBox15.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(15) = Vacaciones Then
        TextBox15.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(15) = Ausencia Then
        TextBox15.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(15) = Nada Then
        TextBox15.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(15) = Feriado Then
        TextBox15.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(15) >= Nada And Dias(15) <= 1 Then
        TextBox15.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox15.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox16 = Entrada Or TextBox16 = Salida Then
        TextBox16.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(16) = Vacaciones Then
        TextBox16.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(16) = Ausencia Then
        TextBox16.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(16) = Nada Then
        TextBox16.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(16) = Feriado Then
        TextBox16.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(16) >= Nada And Dias(16) <= 1 Then
        TextBox16.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox16.BackColor = &HDBDB86                              'Azul
    End If
   If TextBox17 = Entrada Or TextBox17 = Salida Then
        TextBox17.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(17) = Vacaciones Then
        TextBox17.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(17) = Ausencia Then
        TextBox17.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(17) = Nada Then
        TextBox17.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(17) = Feriado Then
        TextBox17.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(17) >= Nada And Dias(17) <= 1 Then
        TextBox17.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox17.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox18 = Entrada Or TextBox18 = Salida Then
        TextBox18.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(18) = Vacaciones Then
        TextBox18.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(18) = Ausencia Then
        TextBox18.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(18) = Nada Then
        TextBox18.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(18) = Feriado Then
        TextBox18.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(18) >= Nada And Dias(18) <= 1 Then
        TextBox18.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox18.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox19 = Entrada Or TextBox19 = Salida Then
        TextBox19.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(19) = Vacaciones Then
        TextBox19.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(19) = Ausencia Then
        TextBox19.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(19) = Nada Then
        TextBox19.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(19) = Feriado Then
        TextBox19.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(19) >= Nada And Dias(19) <= 1 Then
        TextBox19.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox19.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox20 = Entrada Or TextBox20 = Salida Then
        TextBox20.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(20) = Vacaciones Then
        TextBox20.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(20) = Ausencia Then
        TextBox20.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(20) = Nada Then
        TextBox20.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(20) = Feriado Then
        TextBox20.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(20) >= Nada And Dias(20) <= 1 Then
        TextBox20.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox20.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox21 = Entrada Or TextBox21 = Salida Then
        TextBox21.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(21) = Vacaciones Then
        TextBox21.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(21) = Ausencia Then
        TextBox21.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(21) = Nada Then
        TextBox21.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(21) = Feriado Then
        TextBox21.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(21) >= Nada And Dias(21) <= 1 Then
        TextBox21.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox21.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox22 = Entrada Or TextBox22 = Salida Then
        TextBox22.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(22) = Vacaciones Then
        TextBox22.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(22) = Ausencia Then
        TextBox22.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(22) = Nada Then
        TextBox22.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(22) = Feriado Then
        TextBox22.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(22) >= Nada And Dias(22) <= 1 Then
        TextBox22.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox22.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox23 = Entrada Or TextBox23 = Salida Then
        TextBox23.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(23) = Vacaciones Then
        TextBox23.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(23) = Ausencia Then
        TextBox23.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(23) = Nada Then
        TextBox23.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(23) = Feriado Then
        TextBox23.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(23) >= Nada And Dias(23) <= 1 Then
        TextBox23.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox23.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox24 = Entrada Or TextBox24 = Salida Then
        TextBox24.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(24) = Vacaciones Then
        TextBox24.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(24) = Ausencia Then
        TextBox24.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(24) = Nada Then
        TextBox24.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(24) = Feriado Then
        TextBox24.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(24) >= Nada And Dias(24) <= 1 Then
        TextBox24.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox24.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox25 = Entrada Or TextBox25 = Salida Then
        TextBox25.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(25) = Vacaciones Then
        TextBox25.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(25) = Ausencia Then
        TextBox25.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(25) = Nada Then
        TextBox25.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(25) = Feriado Then
        TextBox25.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(25) >= Nada And Dias(25) <= 1 Then
        TextBox25.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox25.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox26 = Entrada Or TextBox26 = Salida Then
        TextBox26.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(26) = Vacaciones Then
        TextBox26.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(26) = Ausencia Then
        TextBox26.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(26) = Nada Then
        TextBox26.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(26) = Feriado Then
        TextBox26.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(26) >= Nada And Dias(26) <= 1 Then
        TextBox26.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox26.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox27 = Entrada Or TextBox27 = Salida Then
        TextBox27.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(27) = Vacaciones Then
        TextBox27.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(27) = Ausencia Then
        TextBox27.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(27) = Nada Then
        TextBox27.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(27) = Feriado Then
        TextBox27.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(27) >= Nada And Dias(27) <= 1 Then
        TextBox27.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox27.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox28 = Entrada Or TextBox28 = Salida Then
        TextBox28.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(28) = Vacaciones Then
        TextBox28.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(28) = Ausencia Then
        TextBox28.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(28) = Nada Then
        TextBox28.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(28) = Feriado Then
        TextBox28.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(28) >= Nada And Dias(28) <= 1 Then
        TextBox28.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox28.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox29 = Entrada Or TextBox29 = Salida Then
        TextBox29.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(29) = Vacaciones Then
        TextBox29.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(29) = Ausencia Then
        TextBox29.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(29) = Nada Then
        TextBox29.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(29) = Feriado Then
        TextBox29.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(29) >= Nada And Dias(29) <= 1 Then
        TextBox29.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox29.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox30 = Entrada Or TextBox30 = Salida Then
        TextBox30.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(30) = Vacaciones Then
        TextBox30.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(30) = Ausencia Then
        TextBox30.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(30) = Nada Then
        TextBox30.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(30) = Feriado Then
        TextBox30.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(30) >= Nada And Dias(30) <= 1 Then
        TextBox30.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox30.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox31 = Entrada Or TextBox31 = Salida Then
        TextBox31.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(31) = Vacaciones Then
        TextBox31.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(31) = Ausencia Then
        TextBox31.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(31) = Nada Then
        TextBox31.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(31) = Feriado Then
        TextBox31.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(31) >= Nada And Dias(31) <= 1 Then
        TextBox31.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox31.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox32 = Entrada Or TextBox32 = Salida Then
        TextBox32.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(32) = Vacaciones Then
        TextBox32.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(32) = Ausencia Then
        TextBox32.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(32) = Nada Then
        TextBox32.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(32) = Feriado Then
        TextBox32.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(32) >= Nada And Dias(32) <= 1 Then
        TextBox32.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox32.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox33 = Entrada Or TextBox33 = Salida Then
        TextBox33.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(33) = Vacaciones Then
        TextBox33.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(33) = Ausencia Then
        TextBox33.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(33) = Nada Then
        TextBox33.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(33) = Feriado Then
        TextBox33.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(33) >= Nada And Dias(33) <= 1 Then
        TextBox33.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox33.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox34 = Entrada Or TextBox34 = Salida Then
        TextBox34.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(34) = Vacaciones Then
        TextBox34.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(34) = Ausencia Then
        TextBox34.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(34) = Nada Then
        TextBox34.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(34) = Feriado Then
        TextBox34.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(34) >= Nada And Dias(34) <= 1 Then
        TextBox34.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox34.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox35 = Entrada Or TextBox35 = Salida Then
        TextBox35.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(35) = Vacaciones Then
        TextBox35.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(35) = Ausencia Then
        TextBox35.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(35) = Nada Then
        TextBox35.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(35) = Feriado Then
        TextBox35.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(35) >= Nada And Dias(35) <= 1 Then
        TextBox35.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox35.BackColor = &HDBDB86                              'Azul
    End If
        If TextBox36 = Entrada Or TextBox36 = Salida Then
        TextBox36.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(36) = Vacaciones Then
        TextBox36.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(36) = Ausencia Then
        TextBox36.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(36) = Nada Then
        TextBox36.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(36) = Feriado Then
        TextBox36.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(36) >= Nada And Dias(36) <= 1 Then
        TextBox36.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox36.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox37 = Entrada Or TextBox37 = Salida Then
        TextBox37.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(37) = Vacaciones Then
        TextBox37.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(37) = Ausencia Then
        TextBox37.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(37) = Nada Then
        TextBox37.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(37) = Feriado Then
        TextBox37.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(37) >= Nada And Dias(37) <= 1 Then
        TextBox37.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox37.BackColor = &HDBDB86                              'Azul
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

Private Sub Limpiar_Filtro()

Range("A1").Select

          If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

           If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData


End Sub
Private Sub Orden_Filtro()

Range("A1").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
Range("B1").Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes

End Sub
Private Sub UserForm_Initialize()
Dim Entrada As String
Dim Salida As String
Dim Vacaciones As String
Dim Ausencia As String
Dim Feriado As String
Dim Nada As Integer
Dim Dias(37) As Variant
Dim Hora(37) As Variant


 With frm_Hora_Marca.cboMes
        .AddItem 1
        .List(0, 1) = "Enero"
        .AddItem 2
        .List(1, 1) = "Febrero"
        .AddItem 3
        .List(2, 1) = "Marzo"
        .AddItem 4
        .List(3, 1) = "Abril"
        .AddItem 5
        .List(4, 1) = "Mayo"
        .AddItem 6
        .List(5, 1) = "Junio"
        .AddItem 7
        .List(6, 1) = "Julio"
        .AddItem 8
        .List(7, 1) = "Agosto"
        .AddItem 9
        .List(8, 1) = "Septiembre"
        .AddItem 10
        .List(9, 1) = "Octubre"
        .AddItem 11
        .List(10, 1) = "Noviembre"
        .AddItem 12
        .List(11, 1) = "Diciembre"
    End With
    
   frm_Hora_Marca.cboMes.ListIndex = VBA.Month(VBA.Date) - 1
       
   frm_Hora_Marca.SpinButton2.Value = VBA.Year(VBA.Date)
    
   frm_Hora_Marca.label_año2.Caption = VBA.Year(VBA.Date)
        
  frm_Hora_Marca.lblHoy.Caption = VBA.Date

Hoja58.Cells(3, 11) = VBA.Month(VBA.Date)

Me.txt_id.Text = Hoja58.Cells(6, 11)
Me.txt_nombre.Text = Hoja58.Cells(6, 12)

'''''''''''''''''''''''''''''''''''''
Me.label_año1.Caption = "AÑO"
Me.label_año2.Caption = Hoja58.Cells(2, 11)

'''''''''''''''''''''''''''''''''''''
Me.Label1.Caption = Hoja58.Cells(1, 2)
Me.Label2.Caption = Hoja58.Cells(1, 3)
Me.Label3.Caption = Hoja58.Cells(1, 4)
Me.Label4.Caption = Hoja58.Cells(1, 5)
Me.Label5.Caption = Hoja58.Cells(1, 6)
Me.Label6.Caption = Hoja58.Cells(1, 7)
Me.Label7.Caption = Hoja58.Cells(1, 8)


Me.TextBox1.Text = Format(Hoja58.Cells(4, 2), "dd;@")
Me.TextBox2.Text = Format(Hoja58.Cells(4, 3), "dd;@")
Me.TextBox3.Text = Format(Hoja58.Cells(4, 4), "dd;@")
Me.TextBox4.Text = Format(Hoja58.Cells(4, 5), "dd;@")
Me.TextBox5.Text = Format(Hoja58.Cells(4, 6), "dd;@")
Me.TextBox6.Text = Format(Hoja58.Cells(4, 7), "dd;@")
Me.TextBox7.Text = Format(Hoja58.Cells(4, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox8.Text = Format(Hoja58.Cells(6, 2), "dd;@")
Me.TextBox9.Text = Format(Hoja58.Cells(6, 3), "dd;@")
Me.TextBox10.Text = Format(Hoja58.Cells(6, 4), "dd;@")
Me.TextBox11.Text = Format(Hoja58.Cells(6, 5), "dd;@")
Me.TextBox12.Text = Format(Hoja58.Cells(6, 6), "dd;@")
Me.TextBox13.Text = Format(Hoja58.Cells(6, 7), "dd;@")
Me.TextBox14.Text = Format(Hoja58.Cells(6, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox15.Text = Format(Hoja58.Cells(8, 2), "dd;@")
Me.TextBox16.Text = Format(Hoja58.Cells(8, 3), "dd;@")
Me.TextBox17.Text = Format(Hoja58.Cells(8, 4), "dd;@")
Me.TextBox18.Text = Format(Hoja58.Cells(8, 5), "dd;@")
Me.TextBox19.Text = Format(Hoja58.Cells(8, 6), "dd;@")
Me.TextBox20.Text = Format(Hoja58.Cells(8, 7), "dd;@")
Me.TextBox21.Text = Format(Hoja58.Cells(8, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox22.Text = Format(Hoja58.Cells(10, 2), "dd;@")
Me.TextBox23.Text = Format(Hoja58.Cells(10, 3), "dd;@")
Me.TextBox24.Text = Format(Hoja58.Cells(10, 4), "dd;@")
Me.TextBox25.Text = Format(Hoja58.Cells(10, 5), "dd;@")
Me.TextBox26.Text = Format(Hoja58.Cells(10, 6), "dd;@")
Me.TextBox27.Text = Format(Hoja58.Cells(10, 7), "dd;@")
Me.TextBox28.Text = Format(Hoja58.Cells(10, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox29.Text = Format(Hoja58.Cells(12, 2), "dd;@")
Me.TextBox30.Text = Format(Hoja58.Cells(12, 3), "dd;@")
Me.TextBox31.Text = Format(Hoja58.Cells(12, 4), "dd;@")
Me.TextBox32.Text = Format(Hoja58.Cells(12, 5), "dd;@")
Me.TextBox33.Text = Format(Hoja58.Cells(12, 6), "dd;@")
Me.TextBox34.Text = Format(Hoja58.Cells(12, 7), "dd;@")
Me.TextBox35.Text = Format(Hoja58.Cells(12, 8), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox36.Text = Format(Hoja58.Cells(14, 2), "dd;@")
Me.TextBox37.Text = Format(Hoja58.Cells(14, 3), "dd;@")


''''''''''''''''''''''''''''''''''''''
'Me.Label8.Caption = " " & Format(Hoja58.Cells(4, 2), "dd;@")
'Me.Label9.Caption = " " & Format(Hoja58.Cells(4, 3), "dd;@")
'Me.Label10.Caption = " " & Format(Hoja58.Cells(4, 4), "dd;@")
'Me.Label11.Caption = " " & Format(Hoja58.Cells(4, 5), "dd;@")
'Me.Label12.Caption = " " & Format(Hoja58.Cells(4, 6), "dd;@")
'Me.Label13.Caption = " " & Format(Hoja58.Cells(4, 7), "dd;@")
'Me.Label14.Caption = " " & Format(Hoja58.Cells(4, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label15.Caption = " " & Format(Hoja58.Cells(6, 2), "dd;@")
'Me.Label16.Caption = " " & Format(Hoja58.Cells(6, 3), "dd;@")
'Me.Label17.Caption = " " & Format(Hoja58.Cells(6, 4), "dd;@")
'Me.Label18.Caption = " " & Format(Hoja58.Cells(6, 5), "dd;@")
'Me.Label19.Caption = " " & Format(Hoja58.Cells(6, 6), "dd;@")
'Me.Label20.Caption = " " & Format(Hoja58.Cells(6, 7), "dd;@")
'Me.Label21.Caption = " " & Format(Hoja58.Cells(6, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label22.Caption = " " & Format(Hoja58.Cells(8, 2), "dd;@")
'Me.Label23.Caption = " " & Format(Hoja58.Cells(8, 3), "dd;@")
'Me.Label24.Caption = " " & Format(Hoja58.Cells(8, 4), "dd;@")
'Me.Label25.Caption = " " & Format(Hoja58.Cells(8, 5), "dd;@")
'Me.Label26.Caption = " " & Format(Hoja58.Cells(8, 6), "dd;@")
'Me.Label27.Caption = " " & Format(Hoja58.Cells(8, 7), "dd;@")
'Me.Label28.Caption = " " & Format(Hoja58.Cells(8, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label29.Caption = " " & Format(Hoja58.Cells(10, 2), "dd;@")
'Me.Label30.Caption = " " & Format(Hoja58.Cells(10, 3), "dd;@")
'Me.Label31.Caption = " " & Format(Hoja58.Cells(10, 4), "dd;@")
'Me.Label32.Caption = " " & Format(Hoja58.Cells(10, 5), "dd;@")
'Me.Label33.Caption = " " & Format(Hoja58.Cells(10, 6), "dd;@")
'Me.Label34.Caption = " " & Format(Hoja58.Cells(10, 7), "dd;@")
'Me.Label35.Caption = " " & Format(Hoja58.Cells(10, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label36.Caption = " " & Format(Hoja58.Cells(12, 2), "dd;@")
'Me.Label37.Caption = " " & Format(Hoja58.Cells(12, 3), "dd;@")
'Me.Label38.Caption = " " & Format(Hoja58.Cells(12, 4), "dd;@")
'Me.Label39.Caption = " " & Format(Hoja58.Cells(12, 5), "dd;@")
'Me.Label40.Caption = " " & Format(Hoja58.Cells(12, 6), "dd;@")
'Me.Label41.Caption = " " & Format(Hoja58.Cells(12, 7), "dd;@")
'Me.Label42.Caption = " " & Format(Hoja58.Cells(12, 8), "dd;@")
'
''''''''''''''''''''''''''''''''''''''
'Me.Label43.Caption = " " & Format(Hoja58.Cells(14, 2), "dd;@")
'Me.Label44.Caption = " " & Format(Hoja58.Cells(14, 3), "dd;@")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox1.Text = Format(Hoja58.Cells(5, 2), "hh:mm")
'Me.TextBox2.Text = Format(Hoja58.Cells(5, 3), "hh:mm")
'Me.TextBox3.Text = Format(Hoja58.Cells(5, 4), "hh:mm")
'Me.TextBox4.Text = Format(Hoja58.Cells(5, 5), "hh:mm")
'Me.TextBox5.Text = Format(Hoja58.Cells(5, 6), "hh:mm")
'Me.TextBox6.Text = Format(Hoja58.Cells(5, 7), "hh:mm")
'Me.TextBox7.Text = Format(Hoja58.Cells(5, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox8.Text = Format(Hoja58.Cells(7, 2), "hh:mm")
'Me.TextBox9.Text = Format(Hoja58.Cells(7, 3), "hh:mm")
'Me.TextBox10.Text = Format(Hoja58.Cells(7, 4), "hh:mm")
'Me.TextBox11.Text = Format(Hoja58.Cells(7, 5), "hh:mm")
'Me.TextBox12.Text = Format(Hoja58.Cells(7, 6), "hh:mm")
'Me.TextBox13.Text = Format(Hoja58.Cells(7, 7), "hh:mm")
'Me.TextBox14.Text = Format(Hoja58.Cells(7, 8), "hh:mm")
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox15.Text = Format(Hoja58.Cells(9, 2), "hh:mm")
'Me.TextBox16.Text = Format(Hoja58.Cells(9, 3), "hh:mm")
'Me.TextBox17.Text = Format(Hoja58.Cells(9, 4), "hh:mm")
'Me.TextBox18.Text = Format(Hoja58.Cells(9, 5), "hh:mm")
'Me.TextBox19.Text = Format(Hoja58.Cells(9, 6), "hh:mm")
'Me.TextBox20.Text = Format(Hoja58.Cells(9, 7), "hh:mm")
'Me.TextBox21.Text = Format(Hoja58.Cells(9, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox22.Text = Format(Hoja58.Cells(11, 2), "hh:mm")
'Me.TextBox23.Text = Format(Hoja58.Cells(11, 3), "hh:mm")
'Me.TextBox24.Text = Format(Hoja58.Cells(11, 4), "hh:mm")
'Me.TextBox25.Text = Format(Hoja58.Cells(11, 5), "hh:mm")
'Me.TextBox26.Text = Format(Hoja58.Cells(11, 6), "hh:mm")
'Me.TextBox27.Text = Format(Hoja58.Cells(11, 7), "hh:mm")
'Me.TextBox28.Text = Format(Hoja58.Cells(11, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox29.Text = Format(Hoja58.Cells(13, 2), "hh:mm")
'Me.TextBox30.Text = Format(Hoja58.Cells(13, 3), "hh:mm")
'Me.TextBox31.Text = Format(Hoja58.Cells(13, 4), "hh:mm")
'Me.TextBox32.Text = Format(Hoja58.Cells(13, 5), "hh:mm")
'Me.TextBox33.Text = Format(Hoja58.Cells(13, 6), "hh:mm")
'Me.TextBox34.Text = Format(Hoja58.Cells(13, 7), "hh:mm")
'Me.TextBox35.Text = Format(Hoja58.Cells(13, 8), "hh:mm")
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.TextBox36.Text = Format(Hoja58.Cells(15, 2), "hh:mm")
'Me.TextBox37.Text = Format(Hoja58.Cells(15, 3), "hh:mm")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Hora(1) = Hoja58.Cells(5, 2)
Hora(2) = Hoja58.Cells(5, 3)
Hora(3) = Hoja58.Cells(5, 4)
Hora(4) = Hoja58.Cells(5, 5)
Hora(5) = Hoja58.Cells(5, 6)
Hora(6) = Hoja58.Cells(5, 7)
Hora(7) = Hoja58.Cells(5, 8)
Hora(8) = Hoja58.Cells(7, 2)
Hora(9) = Hoja58.Cells(7, 3)
Hora(10) = Hoja58.Cells(7, 4)
Hora(11) = Hoja58.Cells(7, 5)
Hora(12) = Hoja58.Cells(7, 6)
Hora(13) = Hoja58.Cells(7, 7)
Hora(14) = Hoja58.Cells(7, 8)
Hora(15) = Hoja58.Cells(9, 2)
Hora(16) = Hoja58.Cells(9, 3)
Hora(17) = Hoja58.Cells(9, 4)
Hora(18) = Hoja58.Cells(9, 5)
Hora(19) = Hoja58.Cells(9, 6)
Hora(20) = Hoja58.Cells(9, 7)
Hora(21) = Hoja58.Cells(9, 8)
Hora(22) = Hoja58.Cells(11, 2)
Hora(23) = Hoja58.Cells(11, 3)
Hora(24) = Hoja58.Cells(11, 4)
Hora(25) = Hoja58.Cells(11, 5)
Hora(26) = Hoja58.Cells(11, 6)
Hora(27) = Hoja58.Cells(11, 7)
Hora(28) = Hoja58.Cells(11, 8)
Hora(29) = Hoja58.Cells(13, 2)
Hora(30) = Hoja58.Cells(13, 3)
Hora(31) = Hoja58.Cells(13, 4)
Hora(32) = Hoja58.Cells(13, 5)
Hora(33) = Hoja58.Cells(13, 6)
Hora(34) = Hoja58.Cells(13, 7)
Hora(35) = Hoja58.Cells(13, 8)
Hora(36) = Hoja58.Cells(15, 2)
Hora(37) = Hoja58.Cells(15, 3)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dias(1) = Hoja58.Cells(5, 2)
Dias(2) = Hoja58.Cells(5, 3)
Dias(3) = Hoja58.Cells(5, 4)
Dias(4) = Hoja58.Cells(5, 5)
Dias(5) = Hoja58.Cells(5, 6)
Dias(6) = Hoja58.Cells(5, 7)
Dias(7) = Hoja58.Cells(5, 8)
Dias(8) = Hoja58.Cells(7, 2)
Dias(9) = Hoja58.Cells(7, 3)
Dias(10) = Hoja58.Cells(7, 4)
Dias(11) = Hoja58.Cells(7, 5)
Dias(12) = Hoja58.Cells(7, 6)
Dias(13) = Hoja58.Cells(7, 7)
Dias(14) = Hoja58.Cells(7, 8)
Dias(15) = Hoja58.Cells(9, 2)
Dias(16) = Hoja58.Cells(9, 3)
Dias(17) = Hoja58.Cells(9, 4)
Dias(18) = Hoja58.Cells(9, 5)
Dias(19) = Hoja58.Cells(9, 6)
Dias(20) = Hoja58.Cells(9, 7)
Dias(21) = Hoja58.Cells(9, 8)
Dias(22) = Hoja58.Cells(11, 2)
Dias(23) = Hoja58.Cells(11, 3)
Dias(24) = Hoja58.Cells(11, 4)
Dias(25) = Hoja58.Cells(11, 5)
Dias(26) = Hoja58.Cells(11, 6)
Dias(27) = Hoja58.Cells(11, 7)
Dias(28) = Hoja58.Cells(11, 8)
Dias(29) = Hoja58.Cells(13, 2)
Dias(30) = Hoja58.Cells(13, 3)
Dias(31) = Hoja58.Cells(13, 4)
Dias(32) = Hoja58.Cells(13, 5)
Dias(33) = Hoja58.Cells(13, 6)
Dias(34) = Hoja58.Cells(13, 7)
Dias(35) = Hoja58.Cells(13, 8)
Dias(36) = Hoja58.Cells(15, 2)
Dias(37) = Hoja58.Cells(15, 3)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Entrada = "PENDIENTE HORA DE ENTRADA"
Salida = "PENDIENTE HORA DE SALIDA"
Vacaciones = "REVISAR REGISTRO"
Ausencia = "AUSENCIA"
Feriado = "DIA FERIADO"
Nada = 0

' &HFAE1CD  ROJO

   
    If Dias(1) = Entrada Or Dias(1) = Salida Then
        TextBox1.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(1) = Vacaciones Then
        TextBox1.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(1) = Ausencia Then
        TextBox1.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(1) = Nada Then
        TextBox1.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(1) = Feriado Then
        TextBox1.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(1) >= Nada And Dias(1) <= 1 Then
        TextBox1.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox1.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox2 = Entrada Or TextBox2 = Salida Then
        TextBox2.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(2) = Vacaciones Then
        TextBox2.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(2) = Ausencia Then
        TextBox2.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(2) = Nada Then
        TextBox2.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(2) = Feriado Then
        TextBox2.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(2) >= Nada And Dias(2) <= 1 Then
        TextBox2.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox2.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox3 = Entrada Or TextBox3 = Salida Then
        TextBox3.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(3) = Vacaciones Then
        TextBox3.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(3) = Ausencia Then
        TextBox3.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(3) = Nada Then
        TextBox3.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(3) = Feriado Then
        TextBox3.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(3) >= Nada And Dias(3) <= 1 Then
        TextBox3.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox3.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox4 = Entrada Or TextBox4 = Salida Then
        TextBox4.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(4) = Vacaciones Then
        TextBox4.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(4) = Ausencia Then
        TextBox4.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(4) = Nada Then
        TextBox4.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(4) = Feriado Then
        TextBox4.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(4) >= Nada And Dias(4) <= 1 Then
        TextBox4.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox4.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox5 = Entrada Or TextBox5 = Salida Then
        TextBox5.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(5) = Vacaciones Then
        TextBox5.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(5) = Ausencia Then
        TextBox5.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(5) = Nada Then
        TextBox5.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(5) = Feriado Then
        TextBox5.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(5) >= Nada And Dias(5) <= 1 Then
        TextBox5.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox5.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox6 = Entrada Or TextBox6 = Salida Then
        TextBox6.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(6) = Vacaciones Then
        TextBox6.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(6) = Ausencia Then
        TextBox6.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(6) = Nada Then
        TextBox6.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(6) = Feriado Then
        TextBox6.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(6) >= Nada And Dias(6) <= 1 Then
        TextBox6.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox6.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox7 = Entrada Or TextBox7 = Salida Then
        TextBox7.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(7) = Vacaciones Then
        TextBox7.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(7) = Ausencia Then
        TextBox7.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(7) = Nada Then
        TextBox7.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(7) = Feriado Then
        TextBox7.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(7) >= Nada And Dias(7) <= 1 Then
        TextBox7.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox7.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox8 = Entrada Or TextBox8 = Salida Then
        TextBox8.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(8) = Vacaciones Then
        TextBox8.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(8) = Ausencia Then
        TextBox8.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(8) = Nada Then
        TextBox8.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(8) = Feriado Then
        TextBox8.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(8) >= Nada And Dias(8) <= 1 Then
        TextBox8.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox8.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox9 = Entrada Or TextBox9 = Salida Then
        TextBox9.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(9) = Vacaciones Then
        TextBox9.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(9) = Ausencia Then
        TextBox9.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(9) = Nada Then
        TextBox9.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(9) = Feriado Then
        TextBox9.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(9) >= Nada And Dias(9) <= 1 Then
        TextBox9.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox9.BackColor = &HDBDB86                              'Azul
    End If
  If TextBox10 = Entrada Or TextBox10 = Salida Then
        TextBox10.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(10) = Vacaciones Then
        TextBox10.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(10) = Ausencia Then
        TextBox10.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(10) = Nada Then
        TextBox10.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(10) = Feriado Then
        TextBox10.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(10) >= Nada And Dias(10) <= 1 Then
        TextBox10.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox10.BackColor = &HDBDB86                              'Azul
    End If
  If TextBox11 = Entrada Or TextBox11 = Salida Then
        TextBox11.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(11) = Vacaciones Then
        TextBox11.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(11) = Ausencia Then
        TextBox11.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(11) = Nada Then
        TextBox11.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(11) = Feriado Then
        TextBox11.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(11) >= Nada And Dias(11) <= 1 Then
        TextBox11.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox11.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox12 = Entrada Or TextBox12 = Salida Then
        TextBox12.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(12) = Vacaciones Then
        TextBox12.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(12) = Ausencia Then
        TextBox12.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(12) = Nada Then
        TextBox12.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(12) = Feriado Then
        TextBox12.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(12) >= Nada And Dias(12) <= 1 Then
        TextBox12.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox12.BackColor = &HDBDB86                              'Azul
    End If
  If TextBox13 = Entrada Or TextBox13 = Salida Then
        TextBox13.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(13) = Vacaciones Then
        TextBox13.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(13) = Ausencia Then
        TextBox13.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(13) = Nada Then
        TextBox13.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(13) = Feriado Then
        TextBox13.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(13) >= Nada And Dias(13) <= 1 Then
        TextBox13.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox13.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox14 = Entrada Or TextBox14 = Salida Then
        TextBox14.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(14) = Vacaciones Then
        TextBox14.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(14) = Ausencia Then
        TextBox14.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(14) = Nada Then
        TextBox14.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(14) = Feriado Then
        TextBox14.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(14) >= Nada And Dias(14) <= 1 Then
        TextBox14.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox14.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox15 = Entrada Or TextBox15 = Salida Then
        TextBox15.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(15) = Vacaciones Then
        TextBox15.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(15) = Ausencia Then
        TextBox15.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(15) = Nada Then
        TextBox15.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(15) = Feriado Then
        TextBox15.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(15) >= Nada And Dias(15) <= 1 Then
        TextBox15.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox15.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox16 = Entrada Or TextBox16 = Salida Then
        TextBox16.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(16) = Vacaciones Then
        TextBox16.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(16) = Ausencia Then
        TextBox16.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(16) = Nada Then
        TextBox16.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(16) = Feriado Then
        TextBox16.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(16) >= Nada And Dias(16) <= 1 Then
        TextBox16.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox16.BackColor = &HDBDB86                              'Azul
    End If
   If TextBox17 = Entrada Or TextBox17 = Salida Then
        TextBox17.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(17) = Vacaciones Then
        TextBox17.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(17) = Ausencia Then
        TextBox17.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(17) = Nada Then
        TextBox17.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(17) = Feriado Then
        TextBox17.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(17) >= Nada And Dias(17) <= 1 Then
        TextBox17.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox17.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox18 = Entrada Or TextBox18 = Salida Then
        TextBox18.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(18) = Vacaciones Then
        TextBox18.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(18) = Ausencia Then
        TextBox18.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(18) = Nada Then
        TextBox18.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(18) = Feriado Then
        TextBox18.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(18) >= Nada And Dias(18) <= 1 Then
        TextBox18.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox18.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox19 = Entrada Or TextBox19 = Salida Then
        TextBox19.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(19) = Vacaciones Then
        TextBox19.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(19) = Ausencia Then
        TextBox19.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(19) = Nada Then
        TextBox19.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(19) = Feriado Then
        TextBox19.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(19) >= Nada And Dias(19) <= 1 Then
        TextBox19.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox19.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox20 = Entrada Or TextBox20 = Salida Then
        TextBox20.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(20) = Vacaciones Then
        TextBox20.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(20) = Ausencia Then
        TextBox20.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(20) = Nada Then
        TextBox20.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(20) = Feriado Then
        TextBox20.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(20) >= Nada And Dias(20) <= 1 Then
        TextBox20.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox20.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox21 = Entrada Or TextBox21 = Salida Then
        TextBox21.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(21) = Vacaciones Then
        TextBox21.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(21) = Ausencia Then
        TextBox21.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(21) = Nada Then
        TextBox21.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(21) = Feriado Then
        TextBox21.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(21) >= Nada And Dias(21) <= 1 Then
        TextBox21.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox21.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox22 = Entrada Or TextBox22 = Salida Then
        TextBox22.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(22) = Vacaciones Then
        TextBox22.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(22) = Ausencia Then
        TextBox22.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(22) = Nada Then
        TextBox22.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(22) = Feriado Then
        TextBox22.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(22) >= Nada And Dias(22) <= 1 Then
        TextBox22.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox22.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox23 = Entrada Or TextBox23 = Salida Then
        TextBox23.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(23) = Vacaciones Then
        TextBox23.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(23) = Ausencia Then
        TextBox23.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(23) = Nada Then
        TextBox23.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(23) = Feriado Then
        TextBox23.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(23) >= Nada And Dias(23) <= 1 Then
        TextBox23.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox23.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox24 = Entrada Or TextBox24 = Salida Then
        TextBox24.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(24) = Vacaciones Then
        TextBox24.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(24) = Ausencia Then
        TextBox24.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(24) = Nada Then
        TextBox24.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(24) = Feriado Then
        TextBox24.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(24) >= Nada And Dias(24) <= 1 Then
        TextBox24.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox24.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox25 = Entrada Or TextBox25 = Salida Then
        TextBox25.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(25) = Vacaciones Then
        TextBox25.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(25) = Ausencia Then
        TextBox25.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(25) = Nada Then
        TextBox25.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(25) = Feriado Then
        TextBox25.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(25) >= Nada And Dias(25) <= 1 Then
        TextBox25.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox25.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox26 = Entrada Or TextBox26 = Salida Then
        TextBox26.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(26) = Vacaciones Then
        TextBox26.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(26) = Ausencia Then
        TextBox26.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(26) = Nada Then
        TextBox26.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(26) = Feriado Then
        TextBox26.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(26) >= Nada And Dias(26) <= 1 Then
        TextBox26.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox26.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox27 = Entrada Or TextBox27 = Salida Then
        TextBox27.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(27) = Vacaciones Then
        TextBox27.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(27) = Ausencia Then
        TextBox27.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(27) = Nada Then
        TextBox27.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(27) = Feriado Then
        TextBox27.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(27) >= Nada And Dias(27) <= 1 Then
        TextBox27.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox27.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox28 = Entrada Or TextBox28 = Salida Then
        TextBox28.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(28) = Vacaciones Then
        TextBox28.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(28) = Ausencia Then
        TextBox28.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(28) = Nada Then
        TextBox28.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(28) = Feriado Then
        TextBox28.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(28) >= Nada And Dias(28) <= 1 Then
        TextBox28.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox28.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox29 = Entrada Or TextBox29 = Salida Then
        TextBox29.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(29) = Vacaciones Then
        TextBox29.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(29) = Ausencia Then
        TextBox29.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(29) = Nada Then
        TextBox29.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(29) = Feriado Then
        TextBox29.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(29) >= Nada And Dias(29) <= 1 Then
        TextBox29.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox29.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox30 = Entrada Or TextBox30 = Salida Then
        TextBox30.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(30) = Vacaciones Then
        TextBox30.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(30) = Ausencia Then
        TextBox30.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(30) = Nada Then
        TextBox30.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(30) = Feriado Then
        TextBox30.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(30) >= Nada And Dias(30) <= 1 Then
        TextBox30.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox30.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox31 = Entrada Or TextBox31 = Salida Then
        TextBox31.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(31) = Vacaciones Then
        TextBox31.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(31) = Ausencia Then
        TextBox31.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(31) = Nada Then
        TextBox31.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(31) = Feriado Then
        TextBox31.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(31) >= Nada And Dias(31) <= 1 Then
        TextBox31.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox31.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox32 = Entrada Or TextBox32 = Salida Then
        TextBox32.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(32) = Vacaciones Then
        TextBox32.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(32) = Ausencia Then
        TextBox32.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(32) = Nada Then
        TextBox32.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(32) = Feriado Then
        TextBox32.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(32) >= Nada And Dias(32) <= 1 Then
        TextBox32.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox32.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox33 = Entrada Or TextBox33 = Salida Then
        TextBox33.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(33) = Vacaciones Then
        TextBox33.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(33) = Ausencia Then
        TextBox33.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(33) = Nada Then
        TextBox33.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(33) = Feriado Then
        TextBox33.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(33) >= Nada And Dias(33) <= 1 Then
        TextBox33.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox33.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox34 = Entrada Or TextBox34 = Salida Then
        TextBox34.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(34) = Vacaciones Then
        TextBox34.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(34) = Ausencia Then
        TextBox34.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(34) = Nada Then
        TextBox34.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(34) = Feriado Then
        TextBox34.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(34) >= Nada And Dias(34) <= 1 Then
        TextBox34.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox34.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox35 = Entrada Or TextBox35 = Salida Then
        TextBox35.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(35) = Vacaciones Then
        TextBox35.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(35) = Ausencia Then
        TextBox35.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(35) = Nada Then
        TextBox35.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(35) = Feriado Then
        TextBox35.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(35) >= Nada And Dias(35) <= 1 Then
        TextBox35.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox35.BackColor = &HDBDB86                              'Azul
    End If
        If TextBox36 = Entrada Or TextBox36 = Salida Then
        TextBox36.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(36) = Vacaciones Then
        TextBox36.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(36) = Ausencia Then
        TextBox36.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(36) = Nada Then
        TextBox36.BackColor = &HFAE1CD                              'rojo
        ElseIf Dias(36) = Feriado Then
        TextBox36.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(36) >= Nada And Dias(36) <= 1 Then
        TextBox36.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox36.BackColor = &HDBDB86                              'Azul
    End If
    If TextBox37 = Entrada Or TextBox37 = Salida Then
        TextBox37.BackColor = &HC8DF9F                              ' verde
        ElseIf Dias(37) = Vacaciones Then
        TextBox37.BackColor = &HD1D7FE                              'A - rojo
        ElseIf Dias(37) = Ausencia Then
        TextBox37.BackColor = &HC0E0FF                              'naranja
        ElseIf Hora(37) = Nada Then
        TextBox37.BackColor = &HCFF9FC                              'R - amarillo
        ElseIf Dias(37) = Feriado Then
        TextBox37.BackColor = &HDBDB86                              'Azul
        ElseIf Dias(37) >= Nada And Dias(37) <= 1 Then
        TextBox37.BackColor = &HFFFFFF                              'blanco
        Else
        TextBox37.BackColor = &HDBDB86                              'Azul
    End If
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Fila As Long
Dim Final As Long
Dim Estado As String
Dim uf As Long
Dim STRG As String
Dim X As Long
On Error Resume Next

Estado = Me.txt_id.Text

Me.lbx_Hora.ColumnCount = 5
Me.lbx_Hora.ColumnWidths = "80 pt;50 pt;60 pt;50 pt"
Me.lbx_Hora.RowSource = "Tbl_Tiempo"

uf = Hoja2.Range("A" & Rows.Count).End(xlUp).Row

Hoja2.AutoFilterMode = False
Me.lbx_Hora = Empty
Me.lbx_Hora.RowSource = Empty

For Fila = 2 To uf
    STRG = Hoja2.Cells(Fila, 2).Value 'Variable para descripción

    If UCase(STRG) Like Estado Then
        Me.lbx_Hora.AddItem
        Me.lbx_Hora.List(X, 0) = Hoja2.Cells(Fila, 1).Value
        Me.lbx_Hora.List(X, 1) = Format(Hoja2.Cells(Fila, 5), "hh:mm")
        Me.lbx_Hora.List(X, 2) = Format(Hoja2.Cells(Fila, 6), "hh:mm")
        Me.lbx_Hora.List(X, 3) = Hoja2.Cells(Fila, 14).Text
        Me.lbx_Hora.List(X, 4) = Hoja2.Cells(Fila, 15).Text
        X = X + 1
   End If
Next

Me.lbx_Hora.ColumnCount = 5
Me.lbx_Hora.ColumnWidths = "80 pt;50 pt;60 pt;50 pt"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub



