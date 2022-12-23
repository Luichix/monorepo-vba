VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_hora_Multiple 
   Caption         =   "CONTROL DE ENTRADAS Y SALIDAS"
   ClientHeight    =   10365
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   18520
   OleObjectBlob   =   "frm_hora_Multiple.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_hora_Multiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Registrar_Hora()
Dim Fecha As Date
Dim Titulo As String
Dim Seguridad As String
Dim Formato As String

Dim xEnter1 As Date
Dim xExit1 As Date
Dim yEnter1 As Date
Dim yExit1 As Date

Dim xEnter2 As Date
Dim xExit2 As Date
Dim yEnter2 As Date
Dim yExit2 As Date

Dim xEnter3 As Date
Dim xExit3 As Date
Dim yEnter3 As Date
Dim yExit3 As Date

Dim xEnter4 As Date
Dim xExit4 As Date
Dim yEnter4 As Date
Dim yExit4 As Date

Dim xEnter5 As Date
Dim xExit5 As Date
Dim yEnter5 As Date
Dim yExit5 As Date

Dim xEnter6 As Date
Dim xExit6 As Date
Dim yEnter6 As Date
Dim yExit6 As Date

Dim xEnter7 As Date
Dim xExit7 As Date
Dim yEnter7 As Date
Dim yExit7 As Date

Dim xEnter8 As Date
Dim xExit8 As Date
Dim yEnter8 As Date
Dim yExit8 As Date

Dim xEnter9 As Date
Dim xExit9 As Date
Dim yEnter9 As Date
Dim yExit9 As Date

Dim xEnter10 As Date
Dim xExit10 As Date
Dim yEnter10 As Date
Dim yExit10 As Date

Dim xEnter11 As Date
Dim xExit11 As Date
Dim yEnter11 As Date
Dim yExit11 As Date

Dim xEnter12 As Date
Dim xExit12 As Date
Dim yEnter12 As Date
Dim yExit12 As Date

Dim xEnter13 As Date
Dim xExit13 As Date
Dim yEnter13 As Date
Dim yExit13 As Date

Dim xEnter14 As Date
Dim xExit14 As Date
Dim yEnter14 As Date
Dim yExit14 As Date

Dim xEnter15 As Date
Dim xExit15 As Date
Dim yEnter15 As Date
Dim yExit15 As Date

Dim xEnter16 As Date
Dim xExit16 As Date
Dim yEnter16 As Date
Dim yExit16 As Date

Dim xEnter17 As Date
Dim xExit17 As Date
Dim yEnter17 As Date
Dim yExit17 As Date



Seguridad = Hoja83.Range("L1").Text

Hoja2.Unprotect (Seguridad)

Titulo = "Gestor de Recursos Humanos"
Formato = "00:00"
    

xEnter1 = Me.txt_xEntrada1.Value
xExit1 = Me.txt_xSalida1.Value
yEnter1 = Me.txt_yEntrada1.Value
yExit1 = Me.txt_ySalida1.Value
           
           
    If Me.txt_xEntrada1 <> Formato Or Me.txt_xSalida1 <> Formato Or Me.txt_yEntrada1 = Formato Or Me.txt_ySalida1 = Formato Then
        If Me.txt_xEntrada1 <> Formato And Me.txt_xSalida1 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha1)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter1
                Hoja2.Cells(3, 6) = xExit1
        End If
        If Me.txt_yEntrada1 <> Formato And Me.txt_ySalida1 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha1)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter1
                Hoja2.Cells(3, 6) = yExit1
        End If
    End If
    
    
xEnter2 = Me.txt_xEntrada2.Value
xExit2 = Me.txt_xSalida2.Value
yEnter2 = Me.txt_yEntrada2.Value
yExit2 = Me.txt_ySalida2.Value
           
           
    If Me.txt_xEntrada2 <> Formato Or Me.txt_xSalida2 <> Formato Or Me.txt_yEntrada2 = Formato Or Me.txt_ySalida2 = Formato Then
        If Me.txt_xEntrada2 <> Formato And Me.txt_xSalida2 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha2)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter2
                Hoja2.Cells(3, 6) = xExit2
        End If
        If Me.txt_yEntrada2 <> Formato And Me.txt_ySalida2 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha2)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter2
                Hoja2.Cells(3, 6) = yExit2
        End If
    End If
    
xEnter3 = Me.txt_xEntrada3.Value
xExit3 = Me.txt_xSalida3.Value
yEnter3 = Me.txt_yEntrada3.Value
yExit3 = Me.txt_ySalida3.Value
           
           
    If Me.txt_xEntrada3 <> Formato Or Me.txt_xSalida3 <> Formato Or Me.txt_yEntrada3 = Formato Or Me.txt_ySalida3 = Formato Then
        If Me.txt_xEntrada3 <> Formato And Me.txt_xSalida3 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha3)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter3
                Hoja2.Cells(3, 6) = xExit3
        End If
        If Me.txt_yEntrada3 <> Formato And Me.txt_ySalida3 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha3)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter3
                Hoja2.Cells(3, 6) = yExit3
        End If
    End If
    
xEnter4 = Me.txt_xEntrada4.Value
xExit4 = Me.txt_xSalida4.Value
yEnter4 = Me.txt_yEntrada4.Value
yExit4 = Me.txt_ySalida4.Value
           
           
    If Me.txt_xEntrada4 <> Formato Or Me.txt_xSalida4 <> Formato Or Me.txt_yEntrada4 = Formato Or Me.txt_ySalida4 = Formato Then
        If Me.txt_xEntrada4 <> Formato And Me.txt_xSalida4 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha4)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter4
                Hoja2.Cells(3, 6) = xExit4
        End If
        If Me.txt_yEntrada4 <> Formato And Me.txt_ySalida4 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha4)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter4
                Hoja2.Cells(3, 6) = yExit4
        End If
    End If
    
xEnter5 = Me.txt_xEntrada5.Value
xExit5 = Me.txt_xSalida5.Value
yEnter5 = Me.txt_yEntrada5.Value
yExit5 = Me.txt_ySalida5.Value
           
           
    If Me.txt_xEntrada5 <> Formato Or Me.txt_xSalida5 <> Formato Or Me.txt_yEntrada5 = Formato Or Me.txt_ySalida5 = Formato Then
        If Me.txt_xEntrada5 <> Formato And Me.txt_xSalida5 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha5)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter5
                Hoja2.Cells(3, 6) = xExit5
        End If
        If Me.txt_yEntrada5 <> Formato And Me.txt_ySalida5 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha5)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter5
                Hoja2.Cells(3, 6) = yExit5
        End If
    End If
    
xEnter6 = Me.txt_xEntrada6.Value
xExit6 = Me.txt_xSalida6.Value
yEnter6 = Me.txt_yEntrada6.Value
yExit6 = Me.txt_ySalida6.Value
           
           
    If Me.txt_xEntrada6 <> Formato Or Me.txt_xSalida6 <> Formato Or Me.txt_yEntrada6 = Formato Or Me.txt_ySalida6 = Formato Then
        If Me.txt_xEntrada6 <> Formato And Me.txt_xSalida6 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha6)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter6
                Hoja2.Cells(3, 6) = xExit6
        End If
        If Me.txt_yEntrada6 <> Formato And Me.txt_ySalida6 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha6)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter6
                Hoja2.Cells(3, 6) = yExit6
        End If
    End If
    
xEnter7 = Me.txt_xEntrada7.Value
xExit7 = Me.txt_xSalida7.Value
yEnter7 = Me.txt_yEntrada7.Value
yExit7 = Me.txt_ySalida7.Value
           
           
    If Me.txt_xEntrada7 <> Formato Or Me.txt_xSalida7 <> Formato Or Me.txt_yEntrada7 = Formato Or Me.txt_ySalida7 = Formato Then
        If Me.txt_xEntrada7 <> Formato And Me.txt_xSalida7 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha7)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter7
                Hoja2.Cells(3, 6) = xExit7
        End If
        If Me.txt_yEntrada7 <> Formato And Me.txt_ySalida7 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha7)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter7
                Hoja2.Cells(3, 6) = yExit7
        End If
    End If
    
xEnter8 = Me.txt_xEntrada8.Value
xExit8 = Me.txt_xSalida8.Value
yEnter8 = Me.txt_yEntrada8.Value
yExit8 = Me.txt_ySalida8.Value
           
           
    If Me.txt_xEntrada8 <> Formato Or Me.txt_xSalida8 <> Formato Or Me.txt_yEntrada8 = Formato Or Me.txt_ySalida8 = Formato Then
        If Me.txt_xEntrada8 <> Formato And Me.txt_xSalida8 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha8)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter8
                Hoja2.Cells(3, 6) = xExit8
        End If
        If Me.txt_yEntrada8 <> Formato And Me.txt_ySalida8 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha8)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter8
                Hoja2.Cells(3, 6) = yExit8
        End If
    End If
    
xEnter9 = Me.txt_xEntrada9.Value
xExit9 = Me.txt_xSalida9.Value
yEnter9 = Me.txt_yEntrada9.Value
yExit9 = Me.txt_ySalida9.Value
           
           
    If Me.txt_xEntrada9 <> Formato Or Me.txt_xSalida9 <> Formato Or Me.txt_yEntrada9 = Formato Or Me.txt_ySalida9 = Formato Then
        If Me.txt_xEntrada9 <> Formato And Me.txt_xSalida9 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha9)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter9
                Hoja2.Cells(3, 6) = xExit9
        End If
        If Me.txt_yEntrada9 <> Formato And Me.txt_ySalida9 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha9)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter9
                Hoja2.Cells(3, 6) = yExit9
        End If
    End If
    
xEnter10 = Me.txt_xEntrada10.Value
xExit10 = Me.txt_xSalida10.Value
yEnter10 = Me.txt_yEntrada10.Value
yExit10 = Me.txt_ySalida10.Value
           
           
    If Me.txt_xEntrada10 <> Formato Or Me.txt_xSalida10 <> Formato Or Me.txt_yEntrada10 = Formato Or Me.txt_ySalida10 = Formato Then
        If Me.txt_xEntrada10 <> Formato And Me.txt_xSalida10 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha10)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter10
                Hoja2.Cells(3, 6) = xExit10
        End If
        If Me.txt_yEntrada10 <> Formato And Me.txt_ySalida10 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha10)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter10
                Hoja2.Cells(3, 6) = yExit10
        End If
    End If
    
xEnter11 = Me.txt_xEntrada11.Value
xExit11 = Me.txt_xSalida11.Value
yEnter11 = Me.txt_yEntrada11.Value
yExit11 = Me.txt_ySalida11.Value
           
           
    If Me.txt_xEntrada11 <> Formato Or Me.txt_xSalida11 <> Formato Or Me.txt_yEntrada11 = Formato Or Me.txt_ySalida11 = Formato Then
        If Me.txt_xEntrada11 <> Formato And Me.txt_xSalida11 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha11)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter11
                Hoja2.Cells(3, 6) = xExit11
        End If
        If Me.txt_yEntrada11 <> Formato And Me.txt_ySalida11 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha11)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter11
                Hoja2.Cells(3, 6) = yExit11
        End If
    End If
    
xEnter12 = Me.txt_xEntrada12.Value
xExit12 = Me.txt_xSalida12.Value
yEnter12 = Me.txt_yEntrada12.Value
yExit12 = Me.txt_ySalida12.Value
           
           
    If Me.txt_xEntrada12 <> Formato Or Me.txt_xSalida12 <> Formato Or Me.txt_yEntrada12 = Formato Or Me.txt_ySalida12 = Formato Then
        If Me.txt_xEntrada12 <> Formato And Me.txt_xSalida12 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha12)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter12
                Hoja2.Cells(3, 6) = xExit12
        End If
        If Me.txt_yEntrada12 <> Formato And Me.txt_ySalida12 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha12)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter12
                Hoja2.Cells(3, 6) = yExit12
        End If
    End If
    
xEnter13 = Me.txt_xEntrada13.Value
xExit13 = Me.txt_xSalida13.Value
yEnter13 = Me.txt_yEntrada13.Value
yExit13 = Me.txt_ySalida13.Value
           
           
    If Me.txt_xEntrada13 <> Formato Or Me.txt_xSalida13 <> Formato Or Me.txt_yEntrada13 = Formato Or Me.txt_ySalida13 = Formato Then
        If Me.txt_xEntrada13 <> Formato And Me.txt_xSalida13 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha13)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter13
                Hoja2.Cells(3, 6) = xExit13
        End If
        If Me.txt_yEntrada13 <> Formato And Me.txt_ySalida13 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha13)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter13
                Hoja2.Cells(3, 6) = yExit13
        End If
    End If
    
xEnter14 = Me.txt_xEntrada14.Value
xExit14 = Me.txt_xSalida14.Value
yEnter14 = Me.txt_yEntrada14.Value
yExit14 = Me.txt_ySalida14.Value
           
           
    If Me.txt_xEntrada14 <> Formato Or Me.txt_xSalida14 <> Formato Or Me.txt_yEntrada14 = Formato Or Me.txt_ySalida14 = Formato Then
        If Me.txt_xEntrada14 <> Formato And Me.txt_xSalida14 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha14)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter14
                Hoja2.Cells(3, 6) = xExit14
        End If
        If Me.txt_yEntrada14 <> Formato And Me.txt_ySalida14 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha14)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter14
                Hoja2.Cells(3, 6) = yExit14
        End If
    End If
    
xEnter15 = Me.txt_xEntrada15.Value
xExit15 = Me.txt_xSalida15.Value
yEnter15 = Me.txt_yEntrada15.Value
yExit15 = Me.txt_ySalida15.Value
           
           
    If Me.txt_xEntrada15 <> Formato Or Me.txt_xSalida15 <> Formato Or Me.txt_yEntrada15 = Formato Or Me.txt_ySalida15 = Formato Then
        If Me.txt_xEntrada15 <> Formato And Me.txt_xSalida15 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha15)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter15
                Hoja2.Cells(3, 6) = xExit15
        End If
        If Me.txt_yEntrada15 <> Formato And Me.txt_ySalida15 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha15)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter15
                Hoja2.Cells(3, 6) = yExit15
        End If
    End If
    
xEnter16 = Me.txt_xEntrada16.Value
xExit16 = Me.txt_xSalida16.Value
yEnter16 = Me.txt_yEntrada16.Value
yExit16 = Me.txt_ySalida16.Value
           
           
    If Me.txt_xEntrada16 <> Formato Or Me.txt_xSalida16 <> Formato Or Me.txt_yEntrada16 = Formato Or Me.txt_ySalida16 = Formato Then
        If Me.txt_xEntrada16 <> Formato And Me.txt_xSalida16 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha16)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = xEnter16
                Hoja2.Cells(3, 6) = xExit16
        End If
        If Me.txt_yEntrada16 <> Formato And Me.txt_ySalida16 <> Formato Then
                
                Hoja2.Select
                Hoja2.Rows("3:3").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja2.Cells(3, 1) = CDate(frm_hora_Multiple.txt_fecha16)
                Hoja2.Cells(3, 2) = Me.txt_id.Text
                Hoja2.Cells(3, 5) = yEnter16
                Hoja2.Cells(3, 6) = yExit16
        End If
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
Dim NEntrada1 As Date
Dim NSalida1 As Date

Dim LEntrada2 As Date
Dim LSalida2 As Date
Dim NEntrada2 As Date
Dim NSalida2 As Date

Dim LEntrada3 As Date
Dim LSalida3 As Date
Dim NEntrada3 As Date
Dim NSalida3 As Date

Dim LEntrada4 As Date
Dim LSalida4 As Date
Dim NEntrada4 As Date
Dim NSalida4 As Date

Dim LEntrada5 As Date
Dim LSalida5 As Date
Dim NEntrada5 As Date
Dim NSalida5 As Date

Dim LEntrada6 As Date
Dim LSalida6 As Date
Dim NEntrada6 As Date
Dim NSalida6 As Date

Dim LEntrada7 As Date
Dim LSalida7 As Date
Dim NEntrada7 As Date
Dim NSalida7 As Date

Dim LEntrada8 As Date
Dim LSalida8 As Date
Dim NEntrada8 As Date
Dim NSalida8 As Date

Dim LEntrada9 As Date
Dim LSalida9 As Date
Dim NEntrada9 As Date
Dim NSalida9 As Date

Dim LEntrada10 As Date
Dim LSalida10 As Date
Dim NEntrada10 As Date
Dim NSalida10 As Date

Dim LEntrada11 As Date
Dim LSalida11 As Date
Dim NEntrada11 As Date
Dim NSalida11 As Date

Dim LEntrada12 As Date
Dim LSalida12 As Date
Dim NEntrada12 As Date
Dim NSalida12 As Date

Dim LEntrada13 As Date
Dim LSalida13 As Date
Dim NEntrada13 As Date
Dim NSalida13 As Date

Dim LEntrada14 As Date
Dim LSalida14 As Date
Dim NEntrada14 As Date
Dim NSalida14 As Date

Dim LEntrada15 As Date
Dim LSalida15 As Date
Dim NEntrada15 As Date
Dim NSalida15 As Date

Dim LEntrada16 As Date
Dim LSalida16 As Date
Dim NEntrada16 As Date
Dim NSalida16 As Date


Formato = "00:00"
Titulo = "Gestor de Recursos Humanos"



    If Me.txt_id.Text = "" Or Me.txt_nombre.Text = "" Then
            MsgBox "Debe seleccionar un colaborador del Listado..!", vbInformation, "Gestor de Recursos Humanos"
            Exit Sub
    End If
    
    If Me.txt_fecha1 = "" Then
        If Me.txt_xEntrada1 <> Formato Or Me.txt_xSalida1 <> Formato Or Me.txt_yEntrada1 <> Formato Or Me.txt_ySalida1 <> Formato Then
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
    
    If Me.txt_xEntrada1 = Formato Or Me.txt_xSalida1 = Formato Or Me.txt_xEntrada1 = "" Or Me.txt_xSalida1 = "" Then
        If Me.txt_yEntrada1 <> "00:00" Or Me.txt_ySalida1 <> "00:00" Then
            Me.txt_xEntrada1.BackColor = &HC0C0FF
            Me.txt_xSalida1.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 01..!", vbInformation, Titulo
            Me.txt_xEntrada1.BackColor = &HFFFFFF
            Me.txt_xSalida1.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
                                If Me.txt_xEntrada1.Text = "" Or Me.txt_xEntrada1.Text = Formato Then
                                    Me.txt_xEntrada1.BackColor = &HC0C0FF
                                    MsgBox "Ingrese los datos correctamente", vbInformation, Titulo
                                    Me.txt_xEntrada1.BackColor = &HFFFFFF
                                    Me.txt_xEntrada1.SetFocus
                                    Exit Sub
                                End If

                                If Me.txt_xSalida1.Text = "" Or Me.txt_xSalida1.Text = Formato Then
                                    Me.txt_xSalida1.BackColor = &HC0C0FF
                                    MsgBox "Ingrese los datos correctamente...!", vbInformation, Titulo
                                    Me.txt_xSalida1.BackColor = &HFFFFFF
                                    Me.txt_xSalida1.SetFocus
                                    Exit Sub
                                End If

    

LEntrada1 = Me.txt_xEntrada1.Value
LSalida1 = Me.txt_xSalida1.Value
NEntrada1 = Me.txt_yEntrada1.Value
NSalida1 = Me.txt_ySalida1.Value

                        
                        If LEntrada1 >= LSalida1 Then
                            Me.txt_xEntrada1.BackColor = &HC0C0FF
                            Me.txt_xSalida1.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente..!", vbInformation, Titulo
                            Me.txt_xEntrada1.BackColor = &HFFFFFF
                            Me.txt_xSalida1.BackColor = &HFFFFFF
                            Me.txt_xEntrada1.SetFocus
                            Exit Sub
                        End If
                        
                If Me.txt_yEntrada1 <> Formato Or Me.txt_ySalida1 <> Formato Then
                         If LSalida1 >= NEntrada1 Then
                            Me.txt_xEntrada1.BackColor = &HC0C0FF
                            Me.txt_xSalida1.BackColor = &HC0C0FF
                            Me.txt_yEntrada1.BackColor = &HC0C0FF
                            Me.txt_ySalida1.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 01...!", vbInformation, Titulo
                            Me.txt_xEntrada1.BackColor = &HFFFFFF
                            Me.txt_xSalida1.BackColor = &HFFFFFF
                            Me.txt_yEntrada1.BackColor = &HFFFFFF
                            Me.txt_ySalida1.BackColor = &HFFFFFF
                            Me.txt_xEntrada1.SetFocus
                            Exit Sub
                        End If
                        If NEntrada1 >= NSalida1 Then
                            Me.txt_yEntrada1.BackColor = &HC0C0FF
                            Me.txt_ySalida1.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 01..!", vbInformation, Titulo
                            Me.txt_yEntrada1.BackColor = &HFFFFFF
                            Me.txt_ySalida1.BackColor = &HFFFFFF
                            Me.txt_yEntrada1.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha2 = "" Then
        If Me.txt_xEntrada2 <> Formato Or Me.txt_xSalida2 <> Formato Or Me.txt_yEntrada2 <> Formato Or Me.txt_ySalida2 <> Formato Then
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
    If Me.txt_xEntrada2 = Formato Or Me.txt_xSalida2 = Formato Or Me.txt_xEntrada2 = "" Or Me.txt_xSalida2 = "" Then
        If Me.txt_yEntrada2 <> Formato Or Me.txt_ySalida2 <> Formato Then
            Me.txt_xEntrada2.BackColor = &HC0C0FF
            Me.txt_xSalida2.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada2.BackColor = &HFFFFFF
            Me.txt_xSalida2.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada2 = Me.txt_xEntrada2.Value
LSalida2 = Me.txt_xSalida2.Value
NEntrada2 = Me.txt_yEntrada2.Value
NSalida2 = Me.txt_ySalida2.Value

                        
                If Me.txt_xEntrada2 <> Formato Or Me.txt_xSalida2 <> Formato Or Me.txt_yEntrada2 <> Formato Or Me.txt_ySalida2 <> Formato Then
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
                        
                If Me.txt_yEntrada2 <> Formato Or Me.txt_ySalida2 <> Formato Then
                         If LSalida2 >= NEntrada2 Then
                            Me.txt_xEntrada2.BackColor = &HC0C0FF
                            Me.txt_xSalida2.BackColor = &HC0C0FF
                            Me.txt_yEntrada2.BackColor = &HC0C0FF
                            Me.txt_ySalida2.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada2.BackColor = &HFFFFFF
                            Me.txt_xSalida2.BackColor = &HFFFFFF
                            Me.txt_yEntrada2.BackColor = &HFFFFFF
                            Me.txt_ySalida2.BackColor = &HFFFFFF
                            Me.txt_xEntrada2.SetFocus
                            Exit Sub
                        End If
                        If NEntrada2 >= NSalida2 Then
                            Me.txt_yEntrada2.BackColor = &HC0C0FF
                            Me.txt_ySalida2.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada2.BackColor = &HFFFFFF
                            Me.txt_ySalida2.BackColor = &HFFFFFF
                            Me.txt_yEntrada2.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha3 = "" Then
        If Me.txt_xEntrada3 <> Formato Or Me.txt_xSalida3 <> Formato Or Me.txt_yEntrada3 <> Formato Or Me.txt_ySalida3 <> Formato Then
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
    If Me.txt_xEntrada3 = Formato Or Me.txt_xSalida3 = Formato Or Me.txt_xEntrada3 = "" Or Me.txt_xSalida3 = "" Then
        If Me.txt_yEntrada3 <> Formato Or Me.txt_ySalida3 <> Formato Then
            Me.txt_xEntrada3.BackColor = &HC0C0FF
            Me.txt_xSalida3.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada3.BackColor = &HFFFFFF
            Me.txt_xSalida3.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada3 = Me.txt_xEntrada3.Value
LSalida3 = Me.txt_xSalida3.Value
NEntrada3 = Me.txt_yEntrada3.Value
NSalida3 = Me.txt_ySalida3.Value

                        
                If Me.txt_xEntrada3 <> Formato Or Me.txt_xSalida3 <> Formato Or Me.txt_yEntrada3 <> Formato Or Me.txt_ySalida3 <> Formato Then
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
                        
                If Me.txt_yEntrada3 <> Formato Or Me.txt_ySalida3 <> Formato Then
                         If LSalida3 >= NEntrada3 Then
                            Me.txt_xEntrada3.BackColor = &HC0C0FF
                            Me.txt_xSalida3.BackColor = &HC0C0FF
                            Me.txt_yEntrada3.BackColor = &HC0C0FF
                            Me.txt_ySalida3.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada3.BackColor = &HFFFFFF
                            Me.txt_xSalida3.BackColor = &HFFFFFF
                            Me.txt_yEntrada3.BackColor = &HFFFFFF
                            Me.txt_ySalida3.BackColor = &HFFFFFF
                            Me.txt_xEntrada3.SetFocus
                            Exit Sub
                        End If
                        If NEntrada3 >= NSalida3 Then
                            Me.txt_yEntrada3.BackColor = &HC0C0FF
                            Me.txt_ySalida3.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada3.BackColor = &HFFFFFF
                            Me.txt_ySalida3.BackColor = &HFFFFFF
                            Me.txt_yEntrada3.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha4 = "" Then
        If Me.txt_xEntrada4 <> Formato Or Me.txt_xSalida4 <> Formato Or Me.txt_yEntrada4 <> Formato Or Me.txt_ySalida4 <> Formato Then
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
    If Me.txt_xEntrada4 = Formato Or Me.txt_xSalida4 = Formato Or Me.txt_xEntrada4 = "" Or Me.txt_xSalida4 = "" Then
        If Me.txt_yEntrada4 <> Formato Or Me.txt_ySalida4 <> Formato Then
            Me.txt_xEntrada4.BackColor = &HC0C0FF
            Me.txt_xSalida4.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada4.BackColor = &HFFFFFF
            Me.txt_xSalida4.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada4 = Me.txt_xEntrada4.Value
LSalida4 = Me.txt_xSalida4.Value
NEntrada4 = Me.txt_yEntrada4.Value
NSalida4 = Me.txt_ySalida4.Value

                        
                If Me.txt_xEntrada4 <> Formato Or Me.txt_xSalida4 <> Formato Or Me.txt_yEntrada4 <> Formato Or Me.txt_ySalida4 <> Formato Then
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
                        
                If Me.txt_yEntrada4 <> Formato Or Me.txt_ySalida4 <> Formato Then
                         If LSalida4 >= NEntrada4 Then
                            Me.txt_xEntrada4.BackColor = &HC0C0FF
                            Me.txt_xSalida4.BackColor = &HC0C0FF
                            Me.txt_yEntrada4.BackColor = &HC0C0FF
                            Me.txt_ySalida4.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada4.BackColor = &HFFFFFF
                            Me.txt_xSalida4.BackColor = &HFFFFFF
                            Me.txt_yEntrada4.BackColor = &HFFFFFF
                            Me.txt_ySalida4.BackColor = &HFFFFFF
                            Me.txt_xEntrada4.SetFocus
                            Exit Sub
                        End If
                        If NEntrada4 >= NSalida4 Then
                            Me.txt_yEntrada4.BackColor = &HC0C0FF
                            Me.txt_ySalida4.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada4.BackColor = &HFFFFFF
                            Me.txt_ySalida4.BackColor = &HFFFFFF
                            Me.txt_yEntrada4.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha5 = "" Then
        If Me.txt_xEntrada5 <> Formato Or Me.txt_xSalida5 <> Formato Or Me.txt_yEntrada5 <> Formato Or Me.txt_ySalida5 <> Formato Then
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
    If Me.txt_xEntrada5 = Formato Or Me.txt_xSalida5 = Formato Or Me.txt_xEntrada5 = "" Or Me.txt_xSalida5 = "" Then
        If Me.txt_yEntrada5 <> Formato Or Me.txt_ySalida5 <> Formato Then
            Me.txt_xEntrada5.BackColor = &HC0C0FF
            Me.txt_xSalida5.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada5.BackColor = &HFFFFFF
            Me.txt_xSalida5.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada5 = Me.txt_xEntrada5.Value
LSalida5 = Me.txt_xSalida5.Value
NEntrada5 = Me.txt_yEntrada5.Value
NSalida5 = Me.txt_ySalida5.Value

                        
                If Me.txt_xEntrada5 <> Formato Or Me.txt_xSalida5 <> Formato Or Me.txt_yEntrada5 <> Formato Or Me.txt_ySalida5 <> Formato Then
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
                        
                If Me.txt_yEntrada5 <> Formato Or Me.txt_ySalida5 <> Formato Then
                         If LSalida5 >= NEntrada5 Then
                            Me.txt_xEntrada5.BackColor = &HC0C0FF
                            Me.txt_xSalida5.BackColor = &HC0C0FF
                            Me.txt_yEntrada5.BackColor = &HC0C0FF
                            Me.txt_ySalida5.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada5.BackColor = &HFFFFFF
                            Me.txt_xSalida5.BackColor = &HFFFFFF
                            Me.txt_yEntrada5.BackColor = &HFFFFFF
                            Me.txt_ySalida5.BackColor = &HFFFFFF
                            Me.txt_xEntrada5.SetFocus
                            Exit Sub
                        End If
                        If NEntrada5 >= NSalida5 Then
                            Me.txt_yEntrada5.BackColor = &HC0C0FF
                            Me.txt_ySalida5.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada5.BackColor = &HFFFFFF
                            Me.txt_ySalida5.BackColor = &HFFFFFF
                            Me.txt_yEntrada5.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha6 = "" Then
        If Me.txt_xEntrada6 <> Formato Or Me.txt_xSalida6 <> Formato Or Me.txt_yEntrada6 <> Formato Or Me.txt_ySalida6 <> Formato Then
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
    If Me.txt_xEntrada6 = Formato Or Me.txt_xSalida6 = Formato Or Me.txt_xEntrada6 = "" Or Me.txt_xSalida6 = "" Then
        If Me.txt_yEntrada6 <> Formato Or Me.txt_ySalida6 <> Formato Then
            Me.txt_xEntrada6.BackColor = &HC0C0FF
            Me.txt_xSalida6.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada6.BackColor = &HFFFFFF
            Me.txt_xSalida6.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada6 = Me.txt_xEntrada6.Value
LSalida6 = Me.txt_xSalida6.Value
NEntrada6 = Me.txt_yEntrada6.Value
NSalida6 = Me.txt_ySalida6.Value

                        
                If Me.txt_xEntrada6 <> Formato Or Me.txt_xSalida6 <> Formato Or Me.txt_yEntrada6 <> Formato Or Me.txt_ySalida6 <> Formato Then
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
                        
                If Me.txt_yEntrada6 <> Formato Or Me.txt_ySalida6 <> Formato Then
                         If LSalida6 >= NEntrada6 Then
                            Me.txt_xEntrada6.BackColor = &HC0C0FF
                            Me.txt_xSalida6.BackColor = &HC0C0FF
                            Me.txt_yEntrada6.BackColor = &HC0C0FF
                            Me.txt_ySalida6.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada6.BackColor = &HFFFFFF
                            Me.txt_xSalida6.BackColor = &HFFFFFF
                            Me.txt_yEntrada6.BackColor = &HFFFFFF
                            Me.txt_ySalida6.BackColor = &HFFFFFF
                            Me.txt_xEntrada6.SetFocus
                            Exit Sub
                        End If
                        If NEntrada6 >= NSalida6 Then
                            Me.txt_yEntrada6.BackColor = &HC0C0FF
                            Me.txt_ySalida6.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada6.BackColor = &HFFFFFF
                            Me.txt_ySalida6.BackColor = &HFFFFFF
                            Me.txt_yEntrada6.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha7 = "" Then
        If Me.txt_xEntrada7 <> Formato Or Me.txt_xSalida7 <> Formato Or Me.txt_yEntrada7 <> Formato Or Me.txt_ySalida7 <> Formato Then
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
    If Me.txt_xEntrada7 = Formato Or Me.txt_xSalida7 = Formato Or Me.txt_xEntrada7 = "" Or Me.txt_xSalida7 = "" Then
        If Me.txt_yEntrada7 <> Formato Or Me.txt_ySalida7 <> Formato Then
            Me.txt_xEntrada7.BackColor = &HC0C0FF
            Me.txt_xSalida7.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada7.BackColor = &HFFFFFF
            Me.txt_xSalida7.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada7 = Me.txt_xEntrada7.Value
LSalida7 = Me.txt_xSalida7.Value
NEntrada7 = Me.txt_yEntrada7.Value
NSalida7 = Me.txt_ySalida7.Value

                        
                If Me.txt_xEntrada7 <> Formato Or Me.txt_xSalida7 <> Formato Or Me.txt_yEntrada7 <> Formato Or Me.txt_ySalida7 <> Formato Then
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
                        
                If Me.txt_yEntrada7 <> Formato Or Me.txt_ySalida7 <> Formato Then
                         If LSalida7 >= NEntrada7 Then
                            Me.txt_xEntrada7.BackColor = &HC0C0FF
                            Me.txt_xSalida7.BackColor = &HC0C0FF
                            Me.txt_yEntrada7.BackColor = &HC0C0FF
                            Me.txt_ySalida7.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada7.BackColor = &HFFFFFF
                            Me.txt_xSalida7.BackColor = &HFFFFFF
                            Me.txt_yEntrada7.BackColor = &HFFFFFF
                            Me.txt_ySalida7.BackColor = &HFFFFFF
                            Me.txt_xEntrada7.SetFocus
                            Exit Sub
                        End If
                        If NEntrada7 >= NSalida7 Then
                            Me.txt_yEntrada7.BackColor = &HC0C0FF
                            Me.txt_ySalida7.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada7.BackColor = &HFFFFFF
                            Me.txt_ySalida7.BackColor = &HFFFFFF
                            Me.txt_yEntrada7.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha8 = "" Then
        If Me.txt_xEntrada8 <> Formato Or Me.txt_xSalida8 <> Formato Or Me.txt_yEntrada8 <> Formato Or Me.txt_ySalida8 <> Formato Then
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
    If Me.txt_xEntrada8 = Formato Or Me.txt_xSalida8 = Formato Or Me.txt_xEntrada8 = "" Or Me.txt_xSalida8 = "" Then
        If Me.txt_yEntrada8 <> Formato Or Me.txt_ySalida8 <> Formato Then
            Me.txt_xEntrada8.BackColor = &HC0C0FF
            Me.txt_xSalida8.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada8.BackColor = &HFFFFFF
            Me.txt_xSalida8.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada8 = Me.txt_xEntrada8.Value
LSalida8 = Me.txt_xSalida8.Value
NEntrada8 = Me.txt_yEntrada8.Value
NSalida8 = Me.txt_ySalida8.Value

                        
                If Me.txt_xEntrada8 <> Formato Or Me.txt_xSalida8 <> Formato Or Me.txt_yEntrada8 <> Formato Or Me.txt_ySalida8 <> Formato Then
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
                        
                If Me.txt_yEntrada8 <> Formato Or Me.txt_ySalida8 <> Formato Then
                         If LSalida8 >= NEntrada8 Then
                            Me.txt_xEntrada8.BackColor = &HC0C0FF
                            Me.txt_xSalida8.BackColor = &HC0C0FF
                            Me.txt_yEntrada8.BackColor = &HC0C0FF
                            Me.txt_ySalida8.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada8.BackColor = &HFFFFFF
                            Me.txt_xSalida8.BackColor = &HFFFFFF
                            Me.txt_yEntrada8.BackColor = &HFFFFFF
                            Me.txt_ySalida8.BackColor = &HFFFFFF
                            Me.txt_xEntrada8.SetFocus
                            Exit Sub
                        End If
                        If NEntrada8 >= NSalida8 Then
                            Me.txt_yEntrada8.BackColor = &HC0C0FF
                            Me.txt_ySalida8.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada8.BackColor = &HFFFFFF
                            Me.txt_ySalida8.BackColor = &HFFFFFF
                            Me.txt_yEntrada8.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha9 = "" Then
        If Me.txt_xEntrada9 <> Formato Or Me.txt_xSalida9 <> Formato Or Me.txt_yEntrada9 <> Formato Or Me.txt_ySalida9 <> Formato Then
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
    If Me.txt_xEntrada9 = Formato Or Me.txt_xSalida9 = Formato Or Me.txt_xEntrada9 = "" Or Me.txt_xSalida9 = "" Then
        If Me.txt_yEntrada9 <> Formato Or Me.txt_ySalida9 <> Formato Then
            Me.txt_xEntrada9.BackColor = &HC0C0FF
            Me.txt_xSalida9.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada9.BackColor = &HFFFFFF
            Me.txt_xSalida9.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada9 = Me.txt_xEntrada9.Value
LSalida9 = Me.txt_xSalida9.Value
NEntrada9 = Me.txt_yEntrada9.Value
NSalida9 = Me.txt_ySalida9.Value

                        
                If Me.txt_xEntrada9 <> Formato Or Me.txt_xSalida9 <> Formato Or Me.txt_yEntrada9 <> Formato Or Me.txt_ySalida9 <> Formato Then
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
                        
                If Me.txt_yEntrada9 <> Formato Or Me.txt_ySalida9 <> Formato Then
                         If LSalida9 >= NEntrada9 Then
                            Me.txt_xEntrada9.BackColor = &HC0C0FF
                            Me.txt_xSalida9.BackColor = &HC0C0FF
                            Me.txt_yEntrada9.BackColor = &HC0C0FF
                            Me.txt_ySalida9.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada9.BackColor = &HFFFFFF
                            Me.txt_xSalida9.BackColor = &HFFFFFF
                            Me.txt_yEntrada9.BackColor = &HFFFFFF
                            Me.txt_ySalida9.BackColor = &HFFFFFF
                            Me.txt_xEntrada9.SetFocus
                            Exit Sub
                        End If
                        If NEntrada9 >= NSalida9 Then
                            Me.txt_yEntrada9.BackColor = &HC0C0FF
                            Me.txt_ySalida9.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada9.BackColor = &HFFFFFF
                            Me.txt_ySalida9.BackColor = &HFFFFFF
                            Me.txt_yEntrada9.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha10 = "" Then
        If Me.txt_xEntrada10 <> Formato Or Me.txt_xSalida10 <> Formato Or Me.txt_yEntrada10 <> Formato Or Me.txt_ySalida10 <> Formato Then
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
    If Me.txt_xEntrada10 = Formato Or Me.txt_xSalida10 = Formato Or Me.txt_xEntrada10 = "" Or Me.txt_xSalida10 = "" Then
        If Me.txt_yEntrada10 <> Formato Or Me.txt_ySalida10 <> Formato Then
            Me.txt_xEntrada10.BackColor = &HC0C0FF
            Me.txt_xSalida10.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada10.BackColor = &HFFFFFF
            Me.txt_xSalida10.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada10 = Me.txt_xEntrada10.Value
LSalida10 = Me.txt_xSalida10.Value
NEntrada10 = Me.txt_yEntrada10.Value
NSalida10 = Me.txt_ySalida10.Value

                        
                If Me.txt_xEntrada10 <> Formato Or Me.txt_xSalida10 <> Formato Or Me.txt_yEntrada10 <> Formato Or Me.txt_ySalida10 <> Formato Then
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
                        
                If Me.txt_yEntrada10 <> Formato Or Me.txt_ySalida10 <> Formato Then
                         If LSalida10 >= NEntrada10 Then
                            Me.txt_xEntrada10.BackColor = &HC0C0FF
                            Me.txt_xSalida10.BackColor = &HC0C0FF
                            Me.txt_yEntrada10.BackColor = &HC0C0FF
                            Me.txt_ySalida10.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada10.BackColor = &HFFFFFF
                            Me.txt_xSalida10.BackColor = &HFFFFFF
                            Me.txt_yEntrada10.BackColor = &HFFFFFF
                            Me.txt_ySalida10.BackColor = &HFFFFFF
                            Me.txt_xEntrada10.SetFocus
                            Exit Sub
                        End If
                        If NEntrada10 >= NSalida10 Then
                            Me.txt_yEntrada10.BackColor = &HC0C0FF
                            Me.txt_ySalida10.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada10.BackColor = &HFFFFFF
                            Me.txt_ySalida10.BackColor = &HFFFFFF
                            Me.txt_yEntrada10.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha11 = "" Then
        If Me.txt_xEntrada11 <> Formato Or Me.txt_xSalida11 <> Formato Or Me.txt_yEntrada11 <> Formato Or Me.txt_ySalida11 <> Formato Then
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
    If Me.txt_xEntrada11 = Formato Or Me.txt_xSalida11 = Formato Or Me.txt_xEntrada11 = "" Or Me.txt_xSalida11 = "" Then
        If Me.txt_yEntrada11 <> Formato Or Me.txt_ySalida11 <> Formato Then
            Me.txt_xEntrada11.BackColor = &HC0C0FF
            Me.txt_xSalida11.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada11.BackColor = &HFFFFFF
            Me.txt_xSalida11.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada11 = Me.txt_xEntrada11.Value
LSalida11 = Me.txt_xSalida11.Value
NEntrada11 = Me.txt_yEntrada11.Value
NSalida11 = Me.txt_ySalida11.Value

                        
                If Me.txt_xEntrada11 <> Formato Or Me.txt_xSalida11 <> Formato Or Me.txt_yEntrada11 <> Formato Or Me.txt_ySalida11 <> Formato Then
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
                        
                If Me.txt_yEntrada11 <> Formato Or Me.txt_ySalida11 <> Formato Then
                         If LSalida11 >= NEntrada11 Then
                            Me.txt_xEntrada11.BackColor = &HC0C0FF
                            Me.txt_xSalida11.BackColor = &HC0C0FF
                            Me.txt_yEntrada11.BackColor = &HC0C0FF
                            Me.txt_ySalida11.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada11.BackColor = &HFFFFFF
                            Me.txt_xSalida11.BackColor = &HFFFFFF
                            Me.txt_yEntrada11.BackColor = &HFFFFFF
                            Me.txt_ySalida11.BackColor = &HFFFFFF
                            Me.txt_xEntrada11.SetFocus
                            Exit Sub
                        End If
                        If NEntrada11 >= NSalida11 Then
                            Me.txt_yEntrada11.BackColor = &HC0C0FF
                            Me.txt_ySalida11.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada11.BackColor = &HFFFFFF
                            Me.txt_ySalida11.BackColor = &HFFFFFF
                            Me.txt_yEntrada11.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha12 = "" Then
        If Me.txt_xEntrada12 <> Formato Or Me.txt_xSalida12 <> Formato Or Me.txt_yEntrada12 <> Formato Or Me.txt_ySalida12 <> Formato Then
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
    If Me.txt_xEntrada12 = Formato Or Me.txt_xSalida12 = Formato Or Me.txt_xEntrada12 = "" Or Me.txt_xSalida12 = "" Then
        If Me.txt_yEntrada12 <> Formato Or Me.txt_ySalida12 <> Formato Then
            Me.txt_xEntrada12.BackColor = &HC0C0FF
            Me.txt_xSalida12.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada12.BackColor = &HFFFFFF
            Me.txt_xSalida12.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada12 = Me.txt_xEntrada12.Value
LSalida12 = Me.txt_xSalida12.Value
NEntrada12 = Me.txt_yEntrada12.Value
NSalida12 = Me.txt_ySalida12.Value

                        
                If Me.txt_xEntrada12 <> Formato Or Me.txt_xSalida12 <> Formato Or Me.txt_yEntrada12 <> Formato Or Me.txt_ySalida12 <> Formato Then
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
                        
                If Me.txt_yEntrada12 <> Formato Or Me.txt_ySalida12 <> Formato Then
                         If LSalida12 >= NEntrada12 Then
                            Me.txt_xEntrada12.BackColor = &HC0C0FF
                            Me.txt_xSalida12.BackColor = &HC0C0FF
                            Me.txt_yEntrada12.BackColor = &HC0C0FF
                            Me.txt_ySalida12.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada12.BackColor = &HFFFFFF
                            Me.txt_xSalida12.BackColor = &HFFFFFF
                            Me.txt_yEntrada12.BackColor = &HFFFFFF
                            Me.txt_ySalida12.BackColor = &HFFFFFF
                            Me.txt_xEntrada12.SetFocus
                            Exit Sub
                        End If
                        If NEntrada12 >= NSalida12 Then
                            Me.txt_yEntrada12.BackColor = &HC0C0FF
                            Me.txt_ySalida12.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada12.BackColor = &HFFFFFF
                            Me.txt_ySalida12.BackColor = &HFFFFFF
                            Me.txt_yEntrada12.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha13 = "" Then
        If Me.txt_xEntrada13 <> Formato Or Me.txt_xSalida13 <> Formato Or Me.txt_yEntrada13 <> Formato Or Me.txt_ySalida13 <> Formato Then
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
    If Me.txt_xEntrada13 = Formato Or Me.txt_xSalida13 = Formato Or Me.txt_xEntrada13 = "" Or Me.txt_xSalida13 = "" Then
        If Me.txt_yEntrada13 <> Formato Or Me.txt_ySalida13 <> Formato Then
            Me.txt_xEntrada13.BackColor = &HC0C0FF
            Me.txt_xSalida13.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada13.BackColor = &HFFFFFF
            Me.txt_xSalida13.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada13 = Me.txt_xEntrada13.Value
LSalida13 = Me.txt_xSalida13.Value
NEntrada13 = Me.txt_yEntrada13.Value
NSalida13 = Me.txt_ySalida13.Value

                        
                If Me.txt_xEntrada13 <> Formato Or Me.txt_xSalida13 <> Formato Or Me.txt_yEntrada13 <> Formato Or Me.txt_ySalida13 <> Formato Then
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
                        
                If Me.txt_yEntrada13 <> Formato Or Me.txt_ySalida13 <> Formato Then
                         If LSalida13 >= NEntrada13 Then
                            Me.txt_xEntrada13.BackColor = &HC0C0FF
                            Me.txt_xSalida13.BackColor = &HC0C0FF
                            Me.txt_yEntrada13.BackColor = &HC0C0FF
                            Me.txt_ySalida13.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada13.BackColor = &HFFFFFF
                            Me.txt_xSalida13.BackColor = &HFFFFFF
                            Me.txt_yEntrada13.BackColor = &HFFFFFF
                            Me.txt_ySalida13.BackColor = &HFFFFFF
                            Me.txt_xEntrada13.SetFocus
                            Exit Sub
                        End If
                        If NEntrada13 >= NSalida13 Then
                            Me.txt_yEntrada13.BackColor = &HC0C0FF
                            Me.txt_ySalida13.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada13.BackColor = &HFFFFFF
                            Me.txt_ySalida13.BackColor = &HFFFFFF
                            Me.txt_yEntrada13.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha14 = "" Then
        If Me.txt_xEntrada14 <> Formato Or Me.txt_xSalida14 <> Formato Or Me.txt_yEntrada14 <> Formato Or Me.txt_ySalida14 <> Formato Then
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
    If Me.txt_xEntrada14 = Formato Or Me.txt_xSalida14 = Formato Or Me.txt_xEntrada14 = "" Or Me.txt_xSalida14 = "" Then
        If Me.txt_yEntrada14 <> Formato Or Me.txt_ySalida14 <> Formato Then
            Me.txt_xEntrada14.BackColor = &HC0C0FF
            Me.txt_xSalida14.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada14.BackColor = &HFFFFFF
            Me.txt_xSalida14.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada14 = Me.txt_xEntrada14.Value
LSalida14 = Me.txt_xSalida14.Value
NEntrada14 = Me.txt_yEntrada14.Value
NSalida14 = Me.txt_ySalida14.Value

                        
                If Me.txt_xEntrada14 <> Formato Or Me.txt_xSalida14 <> Formato Or Me.txt_yEntrada14 <> Formato Or Me.txt_ySalida14 <> Formato Then
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
                        
                If Me.txt_yEntrada14 <> Formato Or Me.txt_ySalida14 <> Formato Then
                         If LSalida14 >= NEntrada14 Then
                            Me.txt_xEntrada14.BackColor = &HC0C0FF
                            Me.txt_xSalida14.BackColor = &HC0C0FF
                            Me.txt_yEntrada14.BackColor = &HC0C0FF
                            Me.txt_ySalida14.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada14.BackColor = &HFFFFFF
                            Me.txt_xSalida14.BackColor = &HFFFFFF
                            Me.txt_yEntrada14.BackColor = &HFFFFFF
                            Me.txt_ySalida14.BackColor = &HFFFFFF
                            Me.txt_xEntrada14.SetFocus
                            Exit Sub
                        End If
                        If NEntrada14 >= NSalida14 Then
                            Me.txt_yEntrada14.BackColor = &HC0C0FF
                            Me.txt_ySalida14.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada14.BackColor = &HFFFFFF
                            Me.txt_ySalida14.BackColor = &HFFFFFF
                            Me.txt_yEntrada14.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha15 = "" Then
        If Me.txt_xEntrada15 <> Formato Or Me.txt_xSalida15 <> Formato Or Me.txt_yEntrada15 <> Formato Or Me.txt_ySalida15 <> Formato Then
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
    If Me.txt_xEntrada15 = Formato Or Me.txt_xSalida15 = Formato Or Me.txt_xEntrada15 = "" Or Me.txt_xSalida15 = "" Then
        If Me.txt_yEntrada15 <> Formato Or Me.txt_ySalida15 <> Formato Then
            Me.txt_xEntrada15.BackColor = &HC0C0FF
            Me.txt_xSalida15.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada15.BackColor = &HFFFFFF
            Me.txt_xSalida15.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada15 = Me.txt_xEntrada15.Value
LSalida15 = Me.txt_xSalida15.Value
NEntrada15 = Me.txt_yEntrada15.Value
NSalida15 = Me.txt_ySalida15.Value

                        
                If Me.txt_xEntrada15 <> Formato Or Me.txt_xSalida15 <> Formato Or Me.txt_yEntrada15 <> Formato Or Me.txt_ySalida15 <> Formato Then
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
                        
                If Me.txt_yEntrada15 <> Formato Or Me.txt_ySalida15 <> Formato Then
                         If LSalida15 >= NEntrada15 Then
                            Me.txt_xEntrada15.BackColor = &HC0C0FF
                            Me.txt_xSalida15.BackColor = &HC0C0FF
                            Me.txt_yEntrada15.BackColor = &HC0C0FF
                            Me.txt_ySalida15.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada15.BackColor = &HFFFFFF
                            Me.txt_xSalida15.BackColor = &HFFFFFF
                            Me.txt_yEntrada15.BackColor = &HFFFFFF
                            Me.txt_ySalida15.BackColor = &HFFFFFF
                            Me.txt_xEntrada15.SetFocus
                            Exit Sub
                        End If
                        If NEntrada15 >= NSalida15 Then
                            Me.txt_yEntrada15.BackColor = &HC0C0FF
                            Me.txt_ySalida15.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada15.BackColor = &HFFFFFF
                            Me.txt_ySalida15.BackColor = &HFFFFFF
                            Me.txt_yEntrada15.SetFocus
                            Exit Sub
                        End If
                End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    If Me.txt_fecha16 = "" Then
        If Me.txt_xEntrada16 <> Formato Or Me.txt_xSalida16 <> Formato Or Me.txt_yEntrada16 <> Formato Or Me.txt_ySalida16 <> Formato Then
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
    If Me.txt_xEntrada16 = Formato Or Me.txt_xSalida16 = Formato Or Me.txt_xEntrada16 = "" Or Me.txt_xSalida16 = "" Then
        If Me.txt_yEntrada16 <> Formato Or Me.txt_ySalida16 <> Formato Then
            Me.txt_xEntrada16.BackColor = &HC0C0FF
            Me.txt_xSalida16.BackColor = &HC0C0FF
            MsgBox "Ingrese las datos correctamente o limpie los datos ingresados incorrectamente: Registros 02..!", vbInformation, Titulo
            Me.txt_xEntrada16.BackColor = &HFFFFFF
            Me.txt_xSalida16.BackColor = &HFFFFFF
            Exit Sub
        End If
    End If
    
LEntrada16 = Me.txt_xEntrada16.Value
LSalida16 = Me.txt_xSalida16.Value
NEntrada16 = Me.txt_yEntrada16.Value
NSalida16 = Me.txt_ySalida16.Value

                        
                If Me.txt_xEntrada16 <> Formato Or Me.txt_xSalida16 <> Formato Or Me.txt_yEntrada16 <> Formato Or Me.txt_ySalida16 <> Formato Then
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
                        
                If Me.txt_yEntrada16 <> Formato Or Me.txt_ySalida16 <> Formato Then
                         If LSalida16 >= NEntrada16 Then
                            Me.txt_xEntrada16.BackColor = &HC0C0FF
                            Me.txt_xSalida16.BackColor = &HC0C0FF
                            Me.txt_yEntrada16.BackColor = &HC0C0FF
                            Me.txt_ySalida16.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02...!", vbInformation, Titulo
                            Me.txt_xEntrada16.BackColor = &HFFFFFF
                            Me.txt_xSalida16.BackColor = &HFFFFFF
                            Me.txt_yEntrada16.BackColor = &HFFFFFF
                            Me.txt_ySalida16.BackColor = &HFFFFFF
                            Me.txt_xEntrada16.SetFocus
                            Exit Sub
                        End If
                        If NEntrada16 >= NSalida16 Then
                            Me.txt_yEntrada16.BackColor = &HC0C0FF
                            Me.txt_ySalida16.BackColor = &HC0C0FF
                            MsgBox "Ingrese los registros horarios correctamente: Registros 02..!", vbInformation, Titulo
                            Me.txt_yEntrada16.BackColor = &HFFFFFF
                            Me.txt_ySalida16.BackColor = &HFFFFFF
                            Me.txt_yEntrada16.SetFocus
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
Private Sub btn_Calendario_Click()
Dim Seguridad As String

    If Me.txt_id.Text = "" Then
        MsgBox "Seleccione un Codigo de Personal", vbInformation, "Gestor de Personal"
        Exit Sub
    Else
    

Seguridad = Hoja83.Range("L1").Text

Hoja58.Unprotect (Seguridad)
Hoja58.Cells(6, 11) = Me.txt_id.Text
        frm_Calendario_Asistencia.Show
Hoja58.Unprotect (Seguridad)
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
Private Sub btn_Limpiar_Click()
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
Private Sub btn_limpiar1_Click()
Dim Ctrl As Control
    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_fecha1" Then
            Ctrl.Value = ""
        End If
    Next Ctrl

    For Each Ctrl In Me.Controls
        If Ctrl.Name = "txt_xEntrada1" Or Ctrl.Name = "txt_xSalida1" Or Ctrl.Name = "txt_yEntrada1" Or Ctrl.Name = "txt_ySalida1" Then
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
        If Ctrl.Name = "txt_xEntrada2" Or Ctrl.Name = "txt_xSalida2" Or Ctrl.Name = "txt_yEntrada2" Or Ctrl.Name = "txt_ySalida2" Then
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
        If Ctrl.Name = "txt_xEntrada3" Or Ctrl.Name = "txt_xSalida3" Or Ctrl.Name = "txt_yEntrada3" Or Ctrl.Name = "txt_ySalida3" Then
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
        If Ctrl.Name = "txt_xEntrada4" Or Ctrl.Name = "txt_xSalida4" Or Ctrl.Name = "txt_yEntrada4" Or Ctrl.Name = "txt_ySalida4" Then
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
        If Ctrl.Name = "txt_xEntrada5" Or Ctrl.Name = "txt_xSalida5" Or Ctrl.Name = "txt_yEntrada5" Or Ctrl.Name = "txt_ySalida5" Then
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
        If Ctrl.Name = "txt_xEntrada6" Or Ctrl.Name = "txt_xSalida6" Or Ctrl.Name = "txt_yEntrada6" Or Ctrl.Name = "txt_ySalida6" Then
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
        If Ctrl.Name = "txt_xEntrada7" Or Ctrl.Name = "txt_xSalida7" Or Ctrl.Name = "txt_yEntrada7" Or Ctrl.Name = "txt_ySalida7" Then
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
        If Ctrl.Name = "txt_xEntrada8" Or Ctrl.Name = "txt_xSalida8" Or Ctrl.Name = "txt_yEntrada8" Or Ctrl.Name = "txt_ySalida8" Then
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
        If Ctrl.Name = "txt_xEntrada9" Or Ctrl.Name = "txt_xSalida9" Or Ctrl.Name = "txt_yEntrada9" Or Ctrl.Name = "txt_ySalida9" Then
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
        If Ctrl.Name = "txt_xEntrada10" Or Ctrl.Name = "txt_xSalida10" Or Ctrl.Name = "txt_yEntrada10" Or Ctrl.Name = "txt_ySalida10" Then
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
        If Ctrl.Name = "txt_xEntrada11" Or Ctrl.Name = "txt_xSalida11" Or Ctrl.Name = "txt_yEntrada11" Or Ctrl.Name = "txt_ySalida11" Then
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
        If Ctrl.Name = "txt_xEntrada12" Or Ctrl.Name = "txt_xSalida12" Or Ctrl.Name = "txt_yEntrada12" Or Ctrl.Name = "txt_ySalida12" Then
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
        If Ctrl.Name = "txt_xEntrada13" Or Ctrl.Name = "txt_xSalida13" Or Ctrl.Name = "txt_yEntrada13" Or Ctrl.Name = "txt_ySalida13" Then
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
        If Ctrl.Name = "txt_xEntrada14" Or Ctrl.Name = "txt_xSalida14" Or Ctrl.Name = "txt_yEntrada14" Or Ctrl.Name = "txt_ySalida14" Then
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
        If Ctrl.Name = "txt_xEntrada15" Or Ctrl.Name = "txt_xSalida15" Or Ctrl.Name = "txt_yEntrada15" Or Ctrl.Name = "txt_ySalida15" Then
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
        If Ctrl.Name = "txt_xEntrada16" Or Ctrl.Name = "txt_xSalida16" Or Ctrl.Name = "txt_yEntrada16" Or Ctrl.Name = "txt_ySalida16" Then
            Ctrl.Value = "00:00"
        End If
    Next Ctrl
End Sub





Private Sub txt_hora1_Change()

End Sub



Private Sub btn_salir_Click()
Unload Me
End Sub







'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada1.Text <> "1" And txt_yEntrada1.Text <> "2" And txt_yEntrada1.Text <> "3" And txt_yEntrada1.Text <> "4" And txt_yEntrada1.Text <> "0" Then
    
    Select Case Len(txt_yEntrada1.Value)
        Case 1
        txt_yEntrada1.Value = txt_yEntrada1.Value & ":"
        Me.txt_yEntrada1.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada1.Value > 9 And txt_yEntrada1.Value < 24 Then
    Select Case Len(txt_yEntrada1.Value)
        Case 2
        txt_yEntrada1.Value = txt_yEntrada1.Value & ":"
        Me.txt_yEntrada1.MaxLength = 5
        End Select
End If
If txt_yEntrada1.Value > 23 And txt_yEntrada1.Value < 30 Or txt_yEntrada1.Value = 0 Or txt_yEntrada1.Value = 3 Or txt_yEntrada1.Value = 4 Then
    txt_yEntrada1 = "00:00"
     Me.txt_yEntrada1.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida1.Text <> "1" And txt_ySalida1.Text <> "2" And txt_ySalida1.Text <> "3" And txt_ySalida1.Text <> "4" And txt_ySalida1.Text <> "0" Then
    
    Select Case Len(txt_ySalida1.Value)
        Case 1
        txt_ySalida1.Value = txt_ySalida1.Value & ":"
        Me.txt_ySalida1.MaxLength = 4
          End Select
        
End If
If txt_ySalida1.Value > 9 And txt_ySalida1.Value < 24 Then
    Select Case Len(txt_ySalida1.Value)
        Case 2
        txt_ySalida1.Value = txt_ySalida1.Value & ":"
        Me.txt_ySalida1.MaxLength = 5
        End Select
End If
If txt_ySalida1.Value > 23 And txt_ySalida1.Value < 30 Or txt_ySalida1.Value = 0 Or txt_ySalida1.Value = 3 Or txt_ySalida1.Value = 4 Then
    txt_ySalida1 = "00:00"
     Me.txt_ySalida1.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada2.Text <> "1" And txt_yEntrada2.Text <> "2" And txt_yEntrada2.Text <> "3" And txt_yEntrada2.Text <> "4" And txt_yEntrada2.Text <> "0" Then
    
    Select Case Len(txt_yEntrada2.Value)
        Case 1
        txt_yEntrada2.Value = txt_yEntrada2.Value & ":"
        Me.txt_yEntrada2.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada2.Value > 9 And txt_yEntrada2.Value < 24 Then
    Select Case Len(txt_yEntrada2.Value)
        Case 2
        txt_yEntrada2.Value = txt_yEntrada2.Value & ":"
        Me.txt_yEntrada2.MaxLength = 5
        End Select
End If
If txt_yEntrada2.Value > 23 And txt_yEntrada2.Value < 30 Or txt_yEntrada2.Value = 0 Or txt_yEntrada2.Value = 3 Or txt_yEntrada2.Value = 4 Then
    txt_yEntrada2 = "00:00"
     Me.txt_yEntrada2.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida2.Text <> "1" And txt_ySalida2.Text <> "2" And txt_ySalida2.Text <> "3" And txt_ySalida2.Text <> "4" And txt_ySalida2.Text <> "0" Then
    
    Select Case Len(txt_ySalida2.Value)
        Case 1
        txt_ySalida2.Value = txt_ySalida2.Value & ":"
        Me.txt_ySalida2.MaxLength = 4
          End Select
        
End If
If txt_ySalida2.Value > 9 And txt_ySalida2.Value < 24 Then
    Select Case Len(txt_ySalida2.Value)
        Case 2
        txt_ySalida2.Value = txt_ySalida2.Value & ":"
        Me.txt_ySalida2.MaxLength = 5
        End Select
End If
If txt_ySalida2.Value > 23 And txt_ySalida2.Value < 30 Or txt_ySalida2.Value = 0 Or txt_ySalida2.Value = 3 Or txt_ySalida2.Value = 4 Then
    txt_ySalida2 = "00:00"
     Me.txt_ySalida2.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada3.Text <> "1" And txt_yEntrada3.Text <> "2" And txt_yEntrada3.Text <> "3" And txt_yEntrada3.Text <> "4" And txt_yEntrada3.Text <> "0" Then
    
    Select Case Len(txt_yEntrada3.Value)
        Case 1
        txt_yEntrada3.Value = txt_yEntrada3.Value & ":"
        Me.txt_yEntrada3.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada3.Value > 9 And txt_yEntrada3.Value < 24 Then
    Select Case Len(txt_yEntrada3.Value)
        Case 2
        txt_yEntrada3.Value = txt_yEntrada3.Value & ":"
        Me.txt_yEntrada3.MaxLength = 5
        End Select
End If
If txt_yEntrada3.Value > 23 And txt_yEntrada3.Value < 30 Or txt_yEntrada3.Value = 0 Or txt_yEntrada3.Value = 3 Or txt_yEntrada3.Value = 4 Then
    txt_yEntrada3 = "00:00"
     Me.txt_yEntrada3.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida3.Text <> "1" And txt_ySalida3.Text <> "2" And txt_ySalida3.Text <> "3" And txt_ySalida3.Text <> "4" And txt_ySalida3.Text <> "0" Then
    
    Select Case Len(txt_ySalida3.Value)
        Case 1
        txt_ySalida3.Value = txt_ySalida3.Value & ":"
        Me.txt_ySalida3.MaxLength = 4
          End Select
        
End If
If txt_ySalida3.Value > 9 And txt_ySalida3.Value < 24 Then
    Select Case Len(txt_ySalida3.Value)
        Case 2
        txt_ySalida3.Value = txt_ySalida3.Value & ":"
        Me.txt_ySalida3.MaxLength = 5
        End Select
End If
If txt_ySalida3.Value > 23 And txt_ySalida3.Value < 30 Or txt_ySalida3.Value = 0 Or txt_ySalida3.Value = 3 Or txt_ySalida3.Value = 4 Then
    txt_ySalida3 = "00:00"
     Me.txt_ySalida3.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada4.Text <> "1" And txt_yEntrada4.Text <> "2" And txt_yEntrada4.Text <> "3" And txt_yEntrada4.Text <> "4" And txt_yEntrada4.Text <> "0" Then
    
    Select Case Len(txt_yEntrada4.Value)
        Case 1
        txt_yEntrada4.Value = txt_yEntrada4.Value & ":"
        Me.txt_yEntrada4.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada4.Value > 9 And txt_yEntrada4.Value < 24 Then
    Select Case Len(txt_yEntrada4.Value)
        Case 2
        txt_yEntrada4.Value = txt_yEntrada4.Value & ":"
        Me.txt_yEntrada4.MaxLength = 5
        End Select
End If
If txt_yEntrada4.Value > 23 And txt_yEntrada4.Value < 30 Or txt_yEntrada4.Value = 0 Or txt_yEntrada4.Value = 3 Or txt_yEntrada4.Value = 4 Then
    txt_yEntrada4 = "00:00"
     Me.txt_yEntrada4.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida4.Text <> "1" And txt_ySalida4.Text <> "2" And txt_ySalida4.Text <> "3" And txt_ySalida4.Text <> "4" And txt_ySalida4.Text <> "0" Then
    
    Select Case Len(txt_ySalida4.Value)
        Case 1
        txt_ySalida4.Value = txt_ySalida4.Value & ":"
        Me.txt_ySalida4.MaxLength = 4
          End Select
        
End If
If txt_ySalida4.Value > 9 And txt_ySalida4.Value < 24 Then
    Select Case Len(txt_ySalida4.Value)
        Case 2
        txt_ySalida4.Value = txt_ySalida4.Value & ":"
        Me.txt_ySalida4.MaxLength = 5
        End Select
End If
If txt_ySalida4.Value > 23 And txt_ySalida4.Value < 30 Or txt_ySalida4.Value = 0 Or txt_ySalida4.Value = 3 Or txt_ySalida4.Value = 4 Then
    txt_ySalida4 = "00:00"
     Me.txt_ySalida4.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada5.Text <> "1" And txt_yEntrada5.Text <> "2" And txt_yEntrada5.Text <> "3" And txt_yEntrada5.Text <> "4" And txt_yEntrada5.Text <> "0" Then
    
    Select Case Len(txt_yEntrada5.Value)
        Case 1
        txt_yEntrada5.Value = txt_yEntrada5.Value & ":"
        Me.txt_yEntrada5.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada5.Value > 9 And txt_yEntrada5.Value < 24 Then
    Select Case Len(txt_yEntrada5.Value)
        Case 2
        txt_yEntrada5.Value = txt_yEntrada5.Value & ":"
        Me.txt_yEntrada5.MaxLength = 5
        End Select
End If
If txt_yEntrada5.Value > 23 And txt_yEntrada5.Value < 30 Or txt_yEntrada5.Value = 0 Or txt_yEntrada5.Value = 3 Or txt_yEntrada5.Value = 4 Then
    txt_yEntrada5 = "00:00"
     Me.txt_yEntrada5.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida5.Text <> "1" And txt_ySalida5.Text <> "2" And txt_ySalida5.Text <> "3" And txt_ySalida5.Text <> "4" And txt_ySalida5.Text <> "0" Then
    
    Select Case Len(txt_ySalida5.Value)
        Case 1
        txt_ySalida5.Value = txt_ySalida5.Value & ":"
        Me.txt_ySalida5.MaxLength = 4
          End Select
        
End If
If txt_ySalida5.Value > 9 And txt_ySalida5.Value < 24 Then
    Select Case Len(txt_ySalida5.Value)
        Case 2
        txt_ySalida5.Value = txt_ySalida5.Value & ":"
        Me.txt_ySalida5.MaxLength = 5
        End Select
End If
If txt_ySalida5.Value > 23 And txt_ySalida5.Value < 30 Or txt_ySalida5.Value = 0 Or txt_ySalida5.Value = 3 Or txt_ySalida5.Value = 4 Then
    txt_ySalida5 = "00:00"
     Me.txt_ySalida5.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada6.Text <> "1" And txt_yEntrada6.Text <> "2" And txt_yEntrada6.Text <> "3" And txt_yEntrada6.Text <> "4" And txt_yEntrada6.Text <> "0" Then
    
    Select Case Len(txt_yEntrada6.Value)
        Case 1
        txt_yEntrada6.Value = txt_yEntrada6.Value & ":"
        Me.txt_yEntrada6.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada6.Value > 9 And txt_yEntrada6.Value < 24 Then
    Select Case Len(txt_yEntrada6.Value)
        Case 2
        txt_yEntrada6.Value = txt_yEntrada6.Value & ":"
        Me.txt_yEntrada6.MaxLength = 5
        End Select
End If
If txt_yEntrada6.Value > 23 And txt_yEntrada6.Value < 30 Or txt_yEntrada6.Value = 0 Or txt_yEntrada6.Value = 3 Or txt_yEntrada6.Value = 4 Then
    txt_yEntrada6 = "00:00"
     Me.txt_yEntrada6.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida6.Text <> "1" And txt_ySalida6.Text <> "2" And txt_ySalida6.Text <> "3" And txt_ySalida6.Text <> "4" And txt_ySalida6.Text <> "0" Then
    
    Select Case Len(txt_ySalida6.Value)
        Case 1
        txt_ySalida6.Value = txt_ySalida6.Value & ":"
        Me.txt_ySalida6.MaxLength = 4
          End Select
        
End If
If txt_ySalida6.Value > 9 And txt_ySalida6.Value < 24 Then
    Select Case Len(txt_ySalida6.Value)
        Case 2
        txt_ySalida6.Value = txt_ySalida6.Value & ":"
        Me.txt_ySalida6.MaxLength = 5
        End Select
End If
If txt_ySalida6.Value > 23 And txt_ySalida6.Value < 30 Or txt_ySalida6.Value = 0 Or txt_ySalida6.Value = 3 Or txt_ySalida6.Value = 4 Then
    txt_ySalida6 = "00:00"
     Me.txt_ySalida6.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada7.Text <> "1" And txt_yEntrada7.Text <> "2" And txt_yEntrada7.Text <> "3" And txt_yEntrada7.Text <> "4" And txt_yEntrada7.Text <> "0" Then
    
    Select Case Len(txt_yEntrada7.Value)
        Case 1
        txt_yEntrada7.Value = txt_yEntrada7.Value & ":"
        Me.txt_yEntrada7.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada7.Value > 9 And txt_yEntrada7.Value < 24 Then
    Select Case Len(txt_yEntrada7.Value)
        Case 2
        txt_yEntrada7.Value = txt_yEntrada7.Value & ":"
        Me.txt_yEntrada7.MaxLength = 5
        End Select
End If
If txt_yEntrada7.Value > 23 And txt_yEntrada7.Value < 30 Or txt_yEntrada7.Value = 0 Or txt_yEntrada7.Value = 3 Or txt_yEntrada7.Value = 4 Then
    txt_yEntrada7 = "00:00"
     Me.txt_yEntrada7.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida7.Text <> "1" And txt_ySalida7.Text <> "2" And txt_ySalida7.Text <> "3" And txt_ySalida7.Text <> "4" And txt_ySalida7.Text <> "0" Then
    
    Select Case Len(txt_ySalida7.Value)
        Case 1
        txt_ySalida7.Value = txt_ySalida7.Value & ":"
        Me.txt_ySalida7.MaxLength = 4
          End Select
        
End If
If txt_ySalida7.Value > 9 And txt_ySalida7.Value < 24 Then
    Select Case Len(txt_ySalida7.Value)
        Case 2
        txt_ySalida7.Value = txt_ySalida7.Value & ":"
        Me.txt_ySalida7.MaxLength = 5
        End Select
End If
If txt_ySalida7.Value > 23 And txt_ySalida7.Value < 30 Or txt_ySalida7.Value = 0 Or txt_ySalida7.Value = 3 Or txt_ySalida7.Value = 4 Then
    txt_ySalida7 = "00:00"
     Me.txt_ySalida7.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada8.Text <> "1" And txt_yEntrada8.Text <> "2" And txt_yEntrada8.Text <> "3" And txt_yEntrada8.Text <> "4" And txt_yEntrada8.Text <> "0" Then
    
    Select Case Len(txt_yEntrada8.Value)
        Case 1
        txt_yEntrada8.Value = txt_yEntrada8.Value & ":"
        Me.txt_yEntrada8.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada8.Value > 9 And txt_yEntrada8.Value < 24 Then
    Select Case Len(txt_yEntrada8.Value)
        Case 2
        txt_yEntrada8.Value = txt_yEntrada8.Value & ":"
        Me.txt_yEntrada8.MaxLength = 5
        End Select
End If
If txt_yEntrada8.Value > 23 And txt_yEntrada8.Value < 30 Or txt_yEntrada8.Value = 0 Or txt_yEntrada8.Value = 3 Or txt_yEntrada8.Value = 4 Then
    txt_yEntrada8 = "00:00"
     Me.txt_yEntrada8.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida8.Text <> "1" And txt_ySalida8.Text <> "2" And txt_ySalida8.Text <> "3" And txt_ySalida8.Text <> "4" And txt_ySalida8.Text <> "0" Then
    
    Select Case Len(txt_ySalida8.Value)
        Case 1
        txt_ySalida8.Value = txt_ySalida8.Value & ":"
        Me.txt_ySalida8.MaxLength = 4
          End Select
        
End If
If txt_ySalida8.Value > 9 And txt_ySalida8.Value < 24 Then
    Select Case Len(txt_ySalida8.Value)
        Case 2
        txt_ySalida8.Value = txt_ySalida8.Value & ":"
        Me.txt_ySalida8.MaxLength = 5
        End Select
End If
If txt_ySalida8.Value > 23 And txt_ySalida8.Value < 30 Or txt_ySalida8.Value = 0 Or txt_ySalida8.Value = 3 Or txt_ySalida8.Value = 4 Then
    txt_ySalida8 = "00:00"
     Me.txt_ySalida8.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada9.Text <> "1" And txt_yEntrada9.Text <> "2" And txt_yEntrada9.Text <> "3" And txt_yEntrada9.Text <> "4" And txt_yEntrada9.Text <> "0" Then
    
    Select Case Len(txt_yEntrada9.Value)
        Case 1
        txt_yEntrada9.Value = txt_yEntrada9.Value & ":"
        Me.txt_yEntrada9.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada9.Value > 9 And txt_yEntrada9.Value < 24 Then
    Select Case Len(txt_yEntrada9.Value)
        Case 2
        txt_yEntrada9.Value = txt_yEntrada9.Value & ":"
        Me.txt_yEntrada9.MaxLength = 5
        End Select
End If
If txt_yEntrada9.Value > 23 And txt_yEntrada9.Value < 30 Or txt_yEntrada9.Value = 0 Or txt_yEntrada9.Value = 3 Or txt_yEntrada9.Value = 4 Then
    txt_yEntrada9 = "00:00"
     Me.txt_yEntrada9.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida9.Text <> "1" And txt_ySalida9.Text <> "2" And txt_ySalida9.Text <> "3" And txt_ySalida9.Text <> "4" And txt_ySalida9.Text <> "0" Then
    
    Select Case Len(txt_ySalida9.Value)
        Case 1
        txt_ySalida9.Value = txt_ySalida9.Value & ":"
        Me.txt_ySalida9.MaxLength = 4
          End Select
        
End If
If txt_ySalida9.Value > 9 And txt_ySalida9.Value < 24 Then
    Select Case Len(txt_ySalida9.Value)
        Case 2
        txt_ySalida9.Value = txt_ySalida9.Value & ":"
        Me.txt_ySalida9.MaxLength = 5
        End Select
End If
If txt_ySalida9.Value > 23 And txt_ySalida9.Value < 30 Or txt_ySalida9.Value = 0 Or txt_ySalida9.Value = 3 Or txt_ySalida9.Value = 4 Then
    txt_ySalida9 = "00:00"
     Me.txt_ySalida9.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada10.Text <> "1" And txt_yEntrada10.Text <> "2" And txt_yEntrada10.Text <> "3" And txt_yEntrada10.Text <> "4" And txt_yEntrada10.Text <> "0" Then
    
    Select Case Len(txt_yEntrada10.Value)
        Case 1
        txt_yEntrada10.Value = txt_yEntrada10.Value & ":"
        Me.txt_yEntrada10.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada10.Value > 9 And txt_yEntrada10.Value < 24 Then
    Select Case Len(txt_yEntrada10.Value)
        Case 2
        txt_yEntrada10.Value = txt_yEntrada10.Value & ":"
        Me.txt_yEntrada10.MaxLength = 5
        End Select
End If
If txt_yEntrada10.Value > 23 And txt_yEntrada10.Value < 30 Or txt_yEntrada10.Value = 0 Or txt_yEntrada10.Value = 3 Or txt_yEntrada10.Value = 4 Then
    txt_yEntrada10 = "00:00"
     Me.txt_yEntrada10.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida10.Text <> "1" And txt_ySalida10.Text <> "2" And txt_ySalida10.Text <> "3" And txt_ySalida10.Text <> "4" And txt_ySalida10.Text <> "0" Then
    
    Select Case Len(txt_ySalida10.Value)
        Case 1
        txt_ySalida10.Value = txt_ySalida10.Value & ":"
        Me.txt_ySalida10.MaxLength = 4
          End Select
        
End If
If txt_ySalida10.Value > 9 And txt_ySalida10.Value < 24 Then
    Select Case Len(txt_ySalida10.Value)
        Case 2
        txt_ySalida10.Value = txt_ySalida10.Value & ":"
        Me.txt_ySalida10.MaxLength = 5
        End Select
End If
If txt_ySalida10.Value > 23 And txt_ySalida10.Value < 30 Or txt_ySalida10.Value = 0 Or txt_ySalida10.Value = 3 Or txt_ySalida10.Value = 4 Then
    txt_ySalida10 = "00:00"
     Me.txt_ySalida10.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada11.Text <> "1" And txt_yEntrada11.Text <> "2" And txt_yEntrada11.Text <> "3" And txt_yEntrada11.Text <> "4" And txt_yEntrada11.Text <> "0" Then
    
    Select Case Len(txt_yEntrada11.Value)
        Case 1
        txt_yEntrada11.Value = txt_yEntrada11.Value & ":"
        Me.txt_yEntrada11.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada11.Value > 9 And txt_yEntrada11.Value < 24 Then
    Select Case Len(txt_yEntrada11.Value)
        Case 2
        txt_yEntrada11.Value = txt_yEntrada11.Value & ":"
        Me.txt_yEntrada11.MaxLength = 5
        End Select
End If
If txt_yEntrada11.Value > 23 And txt_yEntrada11.Value < 30 Or txt_yEntrada11.Value = 0 Or txt_yEntrada11.Value = 3 Or txt_yEntrada11.Value = 4 Then
    txt_yEntrada11 = "00:00"
     Me.txt_yEntrada11.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida11.Text <> "1" And txt_ySalida11.Text <> "2" And txt_ySalida11.Text <> "3" And txt_ySalida11.Text <> "4" And txt_ySalida11.Text <> "0" Then
    
    Select Case Len(txt_ySalida11.Value)
        Case 1
        txt_ySalida11.Value = txt_ySalida11.Value & ":"
        Me.txt_ySalida11.MaxLength = 4
          End Select
        
End If
If txt_ySalida11.Value > 9 And txt_ySalida11.Value < 24 Then
    Select Case Len(txt_ySalida11.Value)
        Case 2
        txt_ySalida11.Value = txt_ySalida11.Value & ":"
        Me.txt_ySalida11.MaxLength = 5
        End Select
End If
If txt_ySalida11.Value > 23 And txt_ySalida11.Value < 30 Or txt_ySalida11.Value = 0 Or txt_ySalida11.Value = 3 Or txt_ySalida11.Value = 4 Then
    txt_ySalida11 = "00:00"
     Me.txt_ySalida11.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada12.Text <> "1" And txt_yEntrada12.Text <> "2" And txt_yEntrada12.Text <> "3" And txt_yEntrada12.Text <> "4" And txt_yEntrada12.Text <> "0" Then
    
    Select Case Len(txt_yEntrada12.Value)
        Case 1
        txt_yEntrada12.Value = txt_yEntrada12.Value & ":"
        Me.txt_yEntrada12.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada12.Value > 9 And txt_yEntrada12.Value < 24 Then
    Select Case Len(txt_yEntrada12.Value)
        Case 2
        txt_yEntrada12.Value = txt_yEntrada12.Value & ":"
        Me.txt_yEntrada12.MaxLength = 5
        End Select
End If
If txt_yEntrada12.Value > 23 And txt_yEntrada12.Value < 30 Or txt_yEntrada12.Value = 0 Or txt_yEntrada12.Value = 3 Or txt_yEntrada12.Value = 4 Then
    txt_yEntrada12 = "00:00"
     Me.txt_yEntrada12.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida12.Text <> "1" And txt_ySalida12.Text <> "2" And txt_ySalida12.Text <> "3" And txt_ySalida12.Text <> "4" And txt_ySalida12.Text <> "0" Then
    
    Select Case Len(txt_ySalida12.Value)
        Case 1
        txt_ySalida12.Value = txt_ySalida12.Value & ":"
        Me.txt_ySalida12.MaxLength = 4
          End Select
        
End If
If txt_ySalida12.Value > 9 And txt_ySalida12.Value < 24 Then
    Select Case Len(txt_ySalida12.Value)
        Case 2
        txt_ySalida12.Value = txt_ySalida12.Value & ":"
        Me.txt_ySalida12.MaxLength = 5
        End Select
End If
If txt_ySalida12.Value > 23 And txt_ySalida12.Value < 30 Or txt_ySalida12.Value = 0 Or txt_ySalida12.Value = 3 Or txt_ySalida12.Value = 4 Then
    txt_ySalida12 = "00:00"
     Me.txt_ySalida12.MaxLength = 4
End If
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada13.Text <> "1" And txt_yEntrada13.Text <> "2" And txt_yEntrada13.Text <> "3" And txt_yEntrada13.Text <> "4" And txt_yEntrada13.Text <> "0" Then
    
    Select Case Len(txt_yEntrada13.Value)
        Case 1
        txt_yEntrada13.Value = txt_yEntrada13.Value & ":"
        Me.txt_yEntrada13.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada13.Value > 9 And txt_yEntrada13.Value < 24 Then
    Select Case Len(txt_yEntrada13.Value)
        Case 2
        txt_yEntrada13.Value = txt_yEntrada13.Value & ":"
        Me.txt_yEntrada13.MaxLength = 5
        End Select
End If
If txt_yEntrada13.Value > 23 And txt_yEntrada13.Value < 30 Or txt_yEntrada13.Value = 0 Or txt_yEntrada13.Value = 3 Or txt_yEntrada13.Value = 4 Then
    txt_yEntrada13 = "00:00"
     Me.txt_yEntrada13.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida13.Text <> "1" And txt_ySalida13.Text <> "2" And txt_ySalida13.Text <> "3" And txt_ySalida13.Text <> "4" And txt_ySalida13.Text <> "0" Then
    
    Select Case Len(txt_ySalida13.Value)
        Case 1
        txt_ySalida13.Value = txt_ySalida13.Value & ":"
        Me.txt_ySalida13.MaxLength = 4
          End Select
        
End If
If txt_ySalida13.Value > 9 And txt_ySalida13.Value < 24 Then
    Select Case Len(txt_ySalida13.Value)
        Case 2
        txt_ySalida13.Value = txt_ySalida13.Value & ":"
        Me.txt_ySalida13.MaxLength = 5
        End Select
End If
If txt_ySalida13.Value > 23 And txt_ySalida13.Value < 30 Or txt_ySalida13.Value = 0 Or txt_ySalida13.Value = 3 Or txt_ySalida13.Value = 4 Then
    txt_ySalida13 = "00:00"
     Me.txt_ySalida13.MaxLength = 4
End If
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada14.Text <> "1" And txt_yEntrada14.Text <> "2" And txt_yEntrada14.Text <> "3" And txt_yEntrada14.Text <> "4" And txt_yEntrada14.Text <> "0" Then
    
    Select Case Len(txt_yEntrada14.Value)
        Case 1
        txt_yEntrada14.Value = txt_yEntrada14.Value & ":"
        Me.txt_yEntrada14.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada14.Value > 9 And txt_yEntrada14.Value < 24 Then
    Select Case Len(txt_yEntrada14.Value)
        Case 2
        txt_yEntrada14.Value = txt_yEntrada14.Value & ":"
        Me.txt_yEntrada14.MaxLength = 5
        End Select
End If
If txt_yEntrada14.Value > 23 And txt_yEntrada14.Value < 30 Or txt_yEntrada14.Value = 0 Or txt_yEntrada14.Value = 3 Or txt_yEntrada14.Value = 4 Then
    txt_yEntrada14 = "00:00"
     Me.txt_yEntrada14.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida14.Text <> "1" And txt_ySalida14.Text <> "2" And txt_ySalida14.Text <> "3" And txt_ySalida14.Text <> "4" And txt_ySalida14.Text <> "0" Then
    
    Select Case Len(txt_ySalida14.Value)
        Case 1
        txt_ySalida14.Value = txt_ySalida14.Value & ":"
        Me.txt_ySalida14.MaxLength = 4
          End Select
        
End If
If txt_ySalida14.Value > 9 And txt_ySalida14.Value < 24 Then
    Select Case Len(txt_ySalida14.Value)
        Case 2
        txt_ySalida14.Value = txt_ySalida14.Value & ":"
        Me.txt_ySalida14.MaxLength = 5
        End Select
End If
If txt_ySalida14.Value > 23 And txt_ySalida14.Value < 30 Or txt_ySalida14.Value = 0 Or txt_ySalida14.Value = 3 Or txt_ySalida14.Value = 4 Then
    txt_ySalida14 = "00:00"
     Me.txt_ySalida14.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada15.Text <> "1" And txt_yEntrada15.Text <> "2" And txt_yEntrada15.Text <> "3" And txt_yEntrada15.Text <> "4" And txt_yEntrada15.Text <> "0" Then
    
    Select Case Len(txt_yEntrada15.Value)
        Case 1
        txt_yEntrada15.Value = txt_yEntrada15.Value & ":"
        Me.txt_yEntrada15.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada15.Value > 9 And txt_yEntrada15.Value < 24 Then
    Select Case Len(txt_yEntrada15.Value)
        Case 2
        txt_yEntrada15.Value = txt_yEntrada15.Value & ":"
        Me.txt_yEntrada15.MaxLength = 5
        End Select
End If
If txt_yEntrada15.Value > 23 And txt_yEntrada15.Value < 30 Or txt_yEntrada15.Value = 0 Or txt_yEntrada15.Value = 3 Or txt_yEntrada15.Value = 4 Then
    txt_yEntrada15 = "00:00"
     Me.txt_yEntrada15.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida15.Text <> "1" And txt_ySalida15.Text <> "2" And txt_ySalida15.Text <> "3" And txt_ySalida15.Text <> "4" And txt_ySalida15.Text <> "0" Then
    
    Select Case Len(txt_ySalida15.Value)
        Case 1
        txt_ySalida15.Value = txt_ySalida15.Value & ":"
        Me.txt_ySalida15.MaxLength = 4
          End Select
        
End If
If txt_ySalida15.Value > 9 And txt_ySalida15.Value < 24 Then
    Select Case Len(txt_ySalida15.Value)
        Case 2
        txt_ySalida15.Value = txt_ySalida15.Value & ":"
        Me.txt_ySalida15.MaxLength = 5
        End Select
End If
If txt_ySalida15.Value > 23 And txt_ySalida15.Value < 30 Or txt_ySalida15.Value = 0 Or txt_ySalida15.Value = 3 Or txt_ySalida15.Value = 4 Then
    txt_ySalida15 = "00:00"
     Me.txt_ySalida15.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub txt_yEntrada16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_yEntrada16.Text <> "1" And txt_yEntrada16.Text <> "2" And txt_yEntrada16.Text <> "3" And txt_yEntrada16.Text <> "4" And txt_yEntrada16.Text <> "0" Then
    
    Select Case Len(txt_yEntrada16.Value)
        Case 1
        txt_yEntrada16.Value = txt_yEntrada16.Value & ":"
        Me.txt_yEntrada16.MaxLength = 4
          End Select
        
    End If
If txt_yEntrada16.Value > 9 And txt_yEntrada16.Value < 24 Then
    Select Case Len(txt_yEntrada16.Value)
        Case 2
        txt_yEntrada16.Value = txt_yEntrada16.Value & ":"
        Me.txt_yEntrada16.MaxLength = 5
        End Select
End If
If txt_yEntrada16.Value > 23 And txt_yEntrada16.Value < 30 Or txt_yEntrada16.Value = 0 Or txt_yEntrada16.Value = 3 Or txt_yEntrada16.Value = 4 Then
    txt_yEntrada16 = "00:00"
     Me.txt_yEntrada16.MaxLength = 4
End If
End Sub
Private Sub txt_ySalida16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If txt_ySalida16.Text <> "1" And txt_ySalida16.Text <> "2" And txt_ySalida16.Text <> "3" And txt_ySalida16.Text <> "4" And txt_ySalida16.Text <> "0" Then
    
    Select Case Len(txt_ySalida16.Value)
        Case 1
        txt_ySalida16.Value = txt_ySalida16.Value & ":"
        Me.txt_ySalida16.MaxLength = 4
          End Select
        
End If
If txt_ySalida16.Value > 9 And txt_ySalida16.Value < 24 Then
    Select Case Len(txt_ySalida16.Value)
        Case 2
        txt_ySalida16.Value = txt_ySalida16.Value & ":"
        Me.txt_ySalida16.MaxLength = 5
        End Select
End If
If txt_ySalida16.Value > 23 And txt_ySalida16.Value < 30 Or txt_ySalida16.Value = 0 Or txt_ySalida16.Value = 3 Or txt_ySalida16.Value = 4 Then
    txt_ySalida16 = "00:00"
     Me.txt_ySalida16.MaxLength = 4
End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub txt_xEntrada1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada1, KeyAscii)
End Sub
Private Sub txt_xSalida1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida1, KeyAscii)
End Sub
Private Sub txt_yEntrada1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada1, KeyAscii)
End Sub
Private Sub txt_ySalida1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida1, KeyAscii)
End Sub

Private Sub txt_xEntrada2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada2, KeyAscii)
End Sub
Private Sub txt_xSalida2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida2, KeyAscii)
End Sub
Private Sub txt_yEntrada2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada2, KeyAscii)
End Sub
Private Sub txt_ySalida2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida2, KeyAscii)
End Sub

Private Sub txt_xEntrada3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada3, KeyAscii)
End Sub
Private Sub txt_xSalida3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida3, KeyAscii)
End Sub
Private Sub txt_yEntrada3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada3, KeyAscii)
End Sub
Private Sub txt_ySalida3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida3, KeyAscii)
End Sub

Private Sub txt_xEntrada4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada4, KeyAscii)
End Sub
Private Sub txt_xSalida4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida4, KeyAscii)
End Sub
Private Sub txt_yEntrada4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada4, KeyAscii)
End Sub
Private Sub txt_ySalida4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida4, KeyAscii)
End Sub

Private Sub txt_xEntrada5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada5, KeyAscii)
End Sub
Private Sub txt_xSalida5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida5, KeyAscii)
End Sub
Private Sub txt_yEntrada5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada5, KeyAscii)
End Sub
Private Sub txt_ySalida5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida5, KeyAscii)
End Sub

Private Sub txt_xEntrada6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada6, KeyAscii)
End Sub
Private Sub txt_xSalida6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida6, KeyAscii)
End Sub
Private Sub txt_yEntrada6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada6, KeyAscii)
End Sub
Private Sub txt_ySalida6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida6, KeyAscii)
End Sub

Private Sub txt_xEntrada7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada7, KeyAscii)
End Sub
Private Sub txt_xSalida7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida7, KeyAscii)
End Sub
Private Sub txt_yEntrada7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada7, KeyAscii)
End Sub
Private Sub txt_ySalida7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida7, KeyAscii)
End Sub

Private Sub txt_xEntrada8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada8, KeyAscii)
End Sub
Private Sub txt_xSalida8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida8, KeyAscii)
End Sub
Private Sub txt_yEntrada8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada8, KeyAscii)
End Sub
Private Sub txt_ySalida8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida8, KeyAscii)
End Sub

Private Sub txt_xEntrada9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada9, KeyAscii)
End Sub
Private Sub txt_xSalida9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida9, KeyAscii)
End Sub
Private Sub txt_yEntrada9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada9, KeyAscii)
End Sub
Private Sub txt_ySalida9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida9, KeyAscii)
End Sub

Private Sub txt_xEntrada10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada10, KeyAscii)
End Sub
Private Sub txt_xSalida10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida10, KeyAscii)
End Sub
Private Sub txt_yEntrada10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada10, KeyAscii)
End Sub
Private Sub txt_ySalida10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida10, KeyAscii)
End Sub

Private Sub txt_xEntrada11_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada11, KeyAscii)
End Sub
Private Sub txt_xSalida11_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida11, KeyAscii)
End Sub
Private Sub txt_yEntrada11_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada11, KeyAscii)
End Sub
Private Sub txt_ySalida11_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida11, KeyAscii)
End Sub

Private Sub txt_xEntrada12_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada12, KeyAscii)
End Sub
Private Sub txt_xSalida12_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida12, KeyAscii)
End Sub
Private Sub txt_yEntrada12_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada12, KeyAscii)
End Sub
Private Sub txt_ySalida12_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida12, KeyAscii)
End Sub

Private Sub txt_xEntrada13_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada13, KeyAscii)
End Sub
Private Sub txt_xSalida13_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida13, KeyAscii)
End Sub
Private Sub txt_yEntrada13_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada13, KeyAscii)
End Sub
Private Sub txt_ySalida13_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida13, KeyAscii)
End Sub

Private Sub txt_xEntrada14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada14, KeyAscii)
End Sub
Private Sub txt_xSalida14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida14, KeyAscii)
End Sub
Private Sub txt_yEntrada14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada14, KeyAscii)
End Sub
Private Sub txt_ySalida14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida14, KeyAscii)
End Sub

Private Sub txt_xEntrada15_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada15, KeyAscii)
End Sub
Private Sub txt_xSalida15_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida15, KeyAscii)
End Sub
Private Sub txt_yEntrada15_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada15, KeyAscii)
End Sub
Private Sub txt_ySalida15_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida15, KeyAscii)
End Sub

Private Sub txt_xEntrada16_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xEntrada16, KeyAscii)
End Sub
Private Sub txt_xSalida16_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_xSalida16, KeyAscii)
End Sub
Private Sub txt_yEntrada16_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_yEntrada16, KeyAscii)
End Sub
Private Sub txt_ySalida16_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = DoblePunto(txt_ySalida16, KeyAscii)
End Sub



Private Sub txt_xEntrada1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada1, KeyCode)
End Sub
Private Sub txt_xSalida1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida1, KeyCode)
End Sub
Private Sub txt_yEntrada1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada1, KeyCode)
End Sub
Private Sub txt_ySalida1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida1, KeyCode)
End Sub

Private Sub txt_xEntrada2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada2, KeyCode)
End Sub
Private Sub txt_xSalida2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida2, KeyCode)
End Sub
Private Sub txt_yEntrada2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada2, KeyCode)
End Sub
Private Sub txt_ySalida2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida2, KeyCode)
End Sub

Private Sub txt_xEntrada3_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada3, KeyCode)
End Sub
Private Sub txt_xSalida3_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida3, KeyCode)
End Sub
Private Sub txt_yEntrada3_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada3, KeyCode)
End Sub
Private Sub txt_ySalida3_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida3, KeyCode)
End Sub

Private Sub txt_xEntrada4_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada4, KeyCode)
End Sub
Private Sub txt_xSalida4_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida4, KeyCode)
End Sub
Private Sub txt_yEntrada4_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada4, KeyCode)
End Sub
Private Sub txt_ySalida4_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida4, KeyCode)
End Sub

Private Sub txt_xEntrada5_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada5, KeyCode)
End Sub
Private Sub txt_xSalida5_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida5, KeyCode)
End Sub
Private Sub txt_yEntrada5_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada5, KeyCode)
End Sub
Private Sub txt_ySalida5_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida5, KeyCode)
End Sub

Private Sub txt_xEntrada6_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada6, KeyCode)
End Sub
Private Sub txt_xSalida6_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida6, KeyCode)
End Sub
Private Sub txt_yEntrada6_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada6, KeyCode)
End Sub
Private Sub txt_ySalida6_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida6, KeyCode)
End Sub

Private Sub txt_xEntrada7_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada7, KeyCode)
End Sub
Private Sub txt_xSalida7_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida7, KeyCode)
End Sub
Private Sub txt_yEntrada7_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada7, KeyCode)
End Sub
Private Sub txt_ySalida7_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida7, KeyCode)
End Sub

Private Sub txt_xEntrada8_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada8, KeyCode)
End Sub
Private Sub txt_xSalida8_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida8, KeyCode)
End Sub
Private Sub txt_yEntrada8_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada8, KeyCode)
End Sub
Private Sub txt_ySalida8_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida8, KeyCode)
End Sub

Private Sub txt_xEntrada9_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada9, KeyCode)
End Sub
Private Sub txt_xSalida9_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida9, KeyCode)
End Sub
Private Sub txt_yEntrada9_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada9, KeyCode)
End Sub
Private Sub txt_ySalida9_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida9, KeyCode)
End Sub

Private Sub txt_xEntrada10_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada10, KeyCode)
End Sub
Private Sub txt_xSalida10_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida10, KeyCode)
End Sub
Private Sub txt_yEntrada10_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada10, KeyCode)
End Sub
Private Sub txt_ySalida10_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida10, KeyCode)
End Sub

Private Sub txt_xEntrada11_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada11, KeyCode)
End Sub
Private Sub txt_xSalida11_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida11, KeyCode)
End Sub
Private Sub txt_yEntrada11_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada11, KeyCode)
End Sub
Private Sub txt_ySalida11_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida11, KeyCode)
End Sub

Private Sub txt_xEntrada12_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada12, KeyCode)
End Sub
Private Sub txt_xSalida12_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida12, KeyCode)
End Sub
Private Sub txt_yEntrada12_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada12, KeyCode)
End Sub
Private Sub txt_ySalida12_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida12, KeyCode)
End Sub

Private Sub txt_xEntrada13_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada13, KeyCode)
End Sub
Private Sub txt_xSalida13_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida13, KeyCode)
End Sub
Private Sub txt_yEntrada13_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada13, KeyCode)
End Sub
Private Sub txt_ySalida13_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida13, KeyCode)
End Sub

Private Sub txt_xEntrada14_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada14, KeyCode)
End Sub
Private Sub txt_xSalida14_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida14, KeyCode)
End Sub
Private Sub txt_yEntrada14_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada14, KeyCode)
End Sub
Private Sub txt_ySalida14_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida14, KeyCode)
End Sub

Private Sub txt_xEntrada15_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada15, KeyCode)
End Sub
Private Sub txt_xSalida15_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida15, KeyCode)
End Sub
Private Sub txt_yEntrada15_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada15, KeyCode)
End Sub
Private Sub txt_ySalida15_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida15, KeyCode)
End Sub

Private Sub txt_xEntrada16_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xEntrada16, KeyCode)
End Sub
Private Sub txt_xSalida16_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_xSalida16, KeyCode)
End Sub
Private Sub txt_yEntrada16_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_yEntrada16, KeyCode)
End Sub
Private Sub txt_ySalida16_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = TextoBorrado(txt_ySalida16, KeyCode)
End Sub


Private Sub btn_personal_Click()
Me.txt_id.BackColor = &H80000005
banderaPersonal = 8
Call LanzarListadoPersonal(Me, "btn_Fecha_Horas")
End Sub


Private Sub UserForm_Click()

End Sub
