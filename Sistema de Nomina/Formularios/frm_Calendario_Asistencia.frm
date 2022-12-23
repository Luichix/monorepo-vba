VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Calendario_Asistencia 
   Caption         =   "CONSULTA DE LABORES"
   ClientHeight    =   9840.001
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   10330
   OleObjectBlob   =   "frm_Calendario_Asistencia.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Calendario_Asistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



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

Me.Label47.Caption = Hoja58.Cells(6, 11)
Me.Label46.Caption = "-  " & Hoja58.Cells(6, 12)

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


'''''''''''''''''''''''''''''''''''''
Me.Label8.Caption = " " & Format(Hoja58.Cells(4, 2), "dd;@")
Me.Label9.Caption = " " & Format(Hoja58.Cells(4, 3), "dd;@")
Me.Label10.Caption = " " & Format(Hoja58.Cells(4, 4), "dd;@")
Me.Label11.Caption = " " & Format(Hoja58.Cells(4, 5), "dd;@")
Me.Label12.Caption = " " & Format(Hoja58.Cells(4, 6), "dd;@")
Me.Label13.Caption = " " & Format(Hoja58.Cells(4, 7), "dd;@")
Me.Label14.Caption = " " & Format(Hoja58.Cells(4, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label15.Caption = " " & Format(Hoja58.Cells(6, 2), "dd;@")
Me.Label16.Caption = " " & Format(Hoja58.Cells(6, 3), "dd;@")
Me.Label17.Caption = " " & Format(Hoja58.Cells(6, 4), "dd;@")
Me.Label18.Caption = " " & Format(Hoja58.Cells(6, 5), "dd;@")
Me.Label19.Caption = " " & Format(Hoja58.Cells(6, 6), "dd;@")
Me.Label20.Caption = " " & Format(Hoja58.Cells(6, 7), "dd;@")
Me.Label21.Caption = " " & Format(Hoja58.Cells(6, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label22.Caption = " " & Format(Hoja58.Cells(8, 2), "dd;@")
Me.Label23.Caption = " " & Format(Hoja58.Cells(8, 3), "dd;@")
Me.Label24.Caption = " " & Format(Hoja58.Cells(8, 4), "dd;@")
Me.Label25.Caption = " " & Format(Hoja58.Cells(8, 5), "dd;@")
Me.Label26.Caption = " " & Format(Hoja58.Cells(8, 6), "dd;@")
Me.Label27.Caption = " " & Format(Hoja58.Cells(8, 7), "dd;@")
Me.Label28.Caption = " " & Format(Hoja58.Cells(8, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label29.Caption = " " & Format(Hoja58.Cells(10, 2), "dd;@")
Me.Label30.Caption = " " & Format(Hoja58.Cells(10, 3), "dd;@")
Me.Label31.Caption = " " & Format(Hoja58.Cells(10, 4), "dd;@")
Me.Label32.Caption = " " & Format(Hoja58.Cells(10, 5), "dd;@")
Me.Label33.Caption = " " & Format(Hoja58.Cells(10, 6), "dd;@")
Me.Label34.Caption = " " & Format(Hoja58.Cells(10, 7), "dd;@")
Me.Label35.Caption = " " & Format(Hoja58.Cells(10, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label36.Caption = " " & Format(Hoja58.Cells(12, 2), "dd;@")
Me.Label37.Caption = " " & Format(Hoja58.Cells(12, 3), "dd;@")
Me.Label38.Caption = " " & Format(Hoja58.Cells(12, 4), "dd;@")
Me.Label39.Caption = " " & Format(Hoja58.Cells(12, 5), "dd;@")
Me.Label40.Caption = " " & Format(Hoja58.Cells(12, 6), "dd;@")
Me.Label41.Caption = " " & Format(Hoja58.Cells(12, 7), "dd;@")
Me.Label42.Caption = " " & Format(Hoja58.Cells(12, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label43.Caption = " " & Format(Hoja58.Cells(14, 2), "dd;@")
Me.Label44.Caption = " " & Format(Hoja58.Cells(14, 3), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox1.Text = Format(Hoja58.Cells(5, 2), "hh:mm")
Me.TextBox2.Text = Format(Hoja58.Cells(5, 3), "hh:mm")
Me.TextBox3.Text = Format(Hoja58.Cells(5, 4), "hh:mm")
Me.TextBox4.Text = Format(Hoja58.Cells(5, 5), "hh:mm")
Me.TextBox5.Text = Format(Hoja58.Cells(5, 6), "hh:mm")
Me.TextBox6.Text = Format(Hoja58.Cells(5, 7), "hh:mm")
Me.TextBox7.Text = Format(Hoja58.Cells(5, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox8.Text = Format(Hoja58.Cells(7, 2), "hh:mm")
Me.TextBox9.Text = Format(Hoja58.Cells(7, 3), "hh:mm")
Me.TextBox10.Text = Format(Hoja58.Cells(7, 4), "hh:mm")
Me.TextBox11.Text = Format(Hoja58.Cells(7, 5), "hh:mm")
Me.TextBox12.Text = Format(Hoja58.Cells(7, 6), "hh:mm")
Me.TextBox13.Text = Format(Hoja58.Cells(7, 7), "hh:mm")
Me.TextBox14.Text = Format(Hoja58.Cells(7, 8), "hh:mm")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox15.Text = Format(Hoja58.Cells(9, 2), "hh:mm")
Me.TextBox16.Text = Format(Hoja58.Cells(9, 3), "hh:mm")
Me.TextBox17.Text = Format(Hoja58.Cells(9, 4), "hh:mm")
Me.TextBox18.Text = Format(Hoja58.Cells(9, 5), "hh:mm")
Me.TextBox19.Text = Format(Hoja58.Cells(9, 6), "hh:mm")
Me.TextBox20.Text = Format(Hoja58.Cells(9, 7), "hh:mm")
Me.TextBox21.Text = Format(Hoja58.Cells(9, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox22.Text = Format(Hoja58.Cells(11, 2), "hh:mm")
Me.TextBox23.Text = Format(Hoja58.Cells(11, 3), "hh:mm")
Me.TextBox24.Text = Format(Hoja58.Cells(11, 4), "hh:mm")
Me.TextBox25.Text = Format(Hoja58.Cells(11, 5), "hh:mm")
Me.TextBox26.Text = Format(Hoja58.Cells(11, 6), "hh:mm")
Me.TextBox27.Text = Format(Hoja58.Cells(11, 7), "hh:mm")
Me.TextBox28.Text = Format(Hoja58.Cells(11, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox29.Text = Format(Hoja58.Cells(13, 2), "hh:mm")
Me.TextBox30.Text = Format(Hoja58.Cells(13, 3), "hh:mm")
Me.TextBox31.Text = Format(Hoja58.Cells(13, 4), "hh:mm")
Me.TextBox32.Text = Format(Hoja58.Cells(13, 5), "hh:mm")
Me.TextBox33.Text = Format(Hoja58.Cells(13, 6), "hh:mm")
Me.TextBox34.Text = Format(Hoja58.Cells(13, 7), "hh:mm")
Me.TextBox35.Text = Format(Hoja58.Cells(13, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox36.Text = Format(Hoja58.Cells(15, 2), "hh:mm")
Me.TextBox37.Text = Format(Hoja58.Cells(15, 3), "hh:mm")

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

   
    If TextBox1 = Entrada Or TextBox1 = Salida Then
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

Me.Label47.Caption = Hoja58.Cells(6, 11)
Me.Label46.Caption = "-  " & Hoja58.Cells(6, 12)


'''''''''''''''''''''''''''''''''''''
Me.Label1.Caption = Hoja58.Cells(1, 2)
Me.Label2.Caption = Hoja58.Cells(1, 3)
Me.Label3.Caption = Hoja58.Cells(1, 4)
Me.Label4.Caption = Hoja58.Cells(1, 5)
Me.Label5.Caption = Hoja58.Cells(1, 6)
Me.Label6.Caption = Hoja58.Cells(1, 7)
Me.Label7.Caption = Hoja58.Cells(1, 8)


'''''''''''''''''''''''''''''''''''''
Me.Label8.Caption = " " & Format(Hoja58.Cells(4, 2), "dd;@")
Me.Label9.Caption = " " & Format(Hoja58.Cells(4, 3), "dd;@")
Me.Label10.Caption = " " & Format(Hoja58.Cells(4, 4), "dd;@")
Me.Label11.Caption = " " & Format(Hoja58.Cells(4, 5), "dd;@")
Me.Label12.Caption = " " & Format(Hoja58.Cells(4, 6), "dd;@")
Me.Label13.Caption = " " & Format(Hoja58.Cells(4, 7), "dd;@")
Me.Label14.Caption = " " & Format(Hoja58.Cells(4, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label15.Caption = " " & Format(Hoja58.Cells(6, 2), "dd;@")
Me.Label16.Caption = " " & Format(Hoja58.Cells(6, 3), "dd;@")
Me.Label17.Caption = " " & Format(Hoja58.Cells(6, 4), "dd;@")
Me.Label18.Caption = " " & Format(Hoja58.Cells(6, 5), "dd;@")
Me.Label19.Caption = " " & Format(Hoja58.Cells(6, 6), "dd;@")
Me.Label20.Caption = " " & Format(Hoja58.Cells(6, 7), "dd;@")
Me.Label21.Caption = " " & Format(Hoja58.Cells(6, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label22.Caption = " " & Format(Hoja58.Cells(8, 2), "dd;@")
Me.Label23.Caption = " " & Format(Hoja58.Cells(8, 3), "dd;@")
Me.Label24.Caption = " " & Format(Hoja58.Cells(8, 4), "dd;@")
Me.Label25.Caption = " " & Format(Hoja58.Cells(8, 5), "dd;@")
Me.Label26.Caption = " " & Format(Hoja58.Cells(8, 6), "dd;@")
Me.Label27.Caption = " " & Format(Hoja58.Cells(8, 7), "dd;@")
Me.Label28.Caption = " " & Format(Hoja58.Cells(8, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label29.Caption = " " & Format(Hoja58.Cells(10, 2), "dd;@")
Me.Label30.Caption = " " & Format(Hoja58.Cells(10, 3), "dd;@")
Me.Label31.Caption = " " & Format(Hoja58.Cells(10, 4), "dd;@")
Me.Label32.Caption = " " & Format(Hoja58.Cells(10, 5), "dd;@")
Me.Label33.Caption = " " & Format(Hoja58.Cells(10, 6), "dd;@")
Me.Label34.Caption = " " & Format(Hoja58.Cells(10, 7), "dd;@")
Me.Label35.Caption = " " & Format(Hoja58.Cells(10, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label36.Caption = " " & Format(Hoja58.Cells(12, 2), "dd;@")
Me.Label37.Caption = " " & Format(Hoja58.Cells(12, 3), "dd;@")
Me.Label38.Caption = " " & Format(Hoja58.Cells(12, 4), "dd;@")
Me.Label39.Caption = " " & Format(Hoja58.Cells(12, 5), "dd;@")
Me.Label40.Caption = " " & Format(Hoja58.Cells(12, 6), "dd;@")
Me.Label41.Caption = " " & Format(Hoja58.Cells(12, 7), "dd;@")
Me.Label42.Caption = " " & Format(Hoja58.Cells(12, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label43.Caption = " " & Format(Hoja58.Cells(14, 2), "dd;@")
Me.Label44.Caption = " " & Format(Hoja58.Cells(14, 3), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox1.Text = Format(Hoja58.Cells(5, 2), "hh:mm")
Me.TextBox2.Text = Format(Hoja58.Cells(5, 3), "hh:mm")
Me.TextBox3.Text = Format(Hoja58.Cells(5, 4), "hh:mm")
Me.TextBox4.Text = Format(Hoja58.Cells(5, 5), "hh:mm")
Me.TextBox5.Text = Format(Hoja58.Cells(5, 6), "hh:mm")
Me.TextBox6.Text = Format(Hoja58.Cells(5, 7), "hh:mm")
Me.TextBox7.Text = Format(Hoja58.Cells(5, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox8.Text = Format(Hoja58.Cells(7, 2), "hh:mm")
Me.TextBox9.Text = Format(Hoja58.Cells(7, 3), "hh:mm")
Me.TextBox10.Text = Format(Hoja58.Cells(7, 4), "hh:mm")
Me.TextBox11.Text = Format(Hoja58.Cells(7, 5), "hh:mm")
Me.TextBox12.Text = Format(Hoja58.Cells(7, 6), "hh:mm")
Me.TextBox13.Text = Format(Hoja58.Cells(7, 7), "hh:mm")
Me.TextBox14.Text = Format(Hoja58.Cells(7, 8), "hh:mm")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox15.Text = Format(Hoja58.Cells(9, 2), "hh:mm")
Me.TextBox16.Text = Format(Hoja58.Cells(9, 3), "hh:mm")
Me.TextBox17.Text = Format(Hoja58.Cells(9, 4), "hh:mm")
Me.TextBox18.Text = Format(Hoja58.Cells(9, 5), "hh:mm")
Me.TextBox19.Text = Format(Hoja58.Cells(9, 6), "hh:mm")
Me.TextBox20.Text = Format(Hoja58.Cells(9, 7), "hh:mm")
Me.TextBox21.Text = Format(Hoja58.Cells(9, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox22.Text = Format(Hoja58.Cells(11, 2), "hh:mm")
Me.TextBox23.Text = Format(Hoja58.Cells(11, 3), "hh:mm")
Me.TextBox24.Text = Format(Hoja58.Cells(11, 4), "hh:mm")
Me.TextBox25.Text = Format(Hoja58.Cells(11, 5), "hh:mm")
Me.TextBox26.Text = Format(Hoja58.Cells(11, 6), "hh:mm")
Me.TextBox27.Text = Format(Hoja58.Cells(11, 7), "hh:mm")
Me.TextBox28.Text = Format(Hoja58.Cells(11, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox29.Text = Format(Hoja58.Cells(13, 2), "hh:mm")
Me.TextBox30.Text = Format(Hoja58.Cells(13, 3), "hh:mm")
Me.TextBox31.Text = Format(Hoja58.Cells(13, 4), "hh:mm")
Me.TextBox32.Text = Format(Hoja58.Cells(13, 5), "hh:mm")
Me.TextBox33.Text = Format(Hoja58.Cells(13, 6), "hh:mm")
Me.TextBox34.Text = Format(Hoja58.Cells(13, 7), "hh:mm")
Me.TextBox35.Text = Format(Hoja58.Cells(13, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox36.Text = Format(Hoja58.Cells(15, 2), "hh:mm")
Me.TextBox37.Text = Format(Hoja58.Cells(15, 3), "hh:mm")

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

   
    If TextBox1 = Entrada Or TextBox1 = Salida Then
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
Me.Label47.Caption = Hoja58.Cells(6, 11)
Me.Label46.Caption = "-  " & Hoja58.Cells(6, 12)

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


'''''''''''''''''''''''''''''''''''''
Me.Label8.Caption = " " & Format(Hoja58.Cells(4, 2), "dd;@")
Me.Label9.Caption = " " & Format(Hoja58.Cells(4, 3), "dd;@")
Me.Label10.Caption = " " & Format(Hoja58.Cells(4, 4), "dd;@")
Me.Label11.Caption = " " & Format(Hoja58.Cells(4, 5), "dd;@")
Me.Label12.Caption = " " & Format(Hoja58.Cells(4, 6), "dd;@")
Me.Label13.Caption = " " & Format(Hoja58.Cells(4, 7), "dd;@")
Me.Label14.Caption = " " & Format(Hoja58.Cells(4, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label15.Caption = " " & Format(Hoja58.Cells(6, 2), "dd;@")
Me.Label16.Caption = " " & Format(Hoja58.Cells(6, 3), "dd;@")
Me.Label17.Caption = " " & Format(Hoja58.Cells(6, 4), "dd;@")
Me.Label18.Caption = " " & Format(Hoja58.Cells(6, 5), "dd;@")
Me.Label19.Caption = " " & Format(Hoja58.Cells(6, 6), "dd;@")
Me.Label20.Caption = " " & Format(Hoja58.Cells(6, 7), "dd;@")
Me.Label21.Caption = " " & Format(Hoja58.Cells(6, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label22.Caption = " " & Format(Hoja58.Cells(8, 2), "dd;@")
Me.Label23.Caption = " " & Format(Hoja58.Cells(8, 3), "dd;@")
Me.Label24.Caption = " " & Format(Hoja58.Cells(8, 4), "dd;@")
Me.Label25.Caption = " " & Format(Hoja58.Cells(8, 5), "dd;@")
Me.Label26.Caption = " " & Format(Hoja58.Cells(8, 6), "dd;@")
Me.Label27.Caption = " " & Format(Hoja58.Cells(8, 7), "dd;@")
Me.Label28.Caption = " " & Format(Hoja58.Cells(8, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label29.Caption = " " & Format(Hoja58.Cells(10, 2), "dd;@")
Me.Label30.Caption = " " & Format(Hoja58.Cells(10, 3), "dd;@")
Me.Label31.Caption = " " & Format(Hoja58.Cells(10, 4), "dd;@")
Me.Label32.Caption = " " & Format(Hoja58.Cells(10, 5), "dd;@")
Me.Label33.Caption = " " & Format(Hoja58.Cells(10, 6), "dd;@")
Me.Label34.Caption = " " & Format(Hoja58.Cells(10, 7), "dd;@")
Me.Label35.Caption = " " & Format(Hoja58.Cells(10, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label36.Caption = " " & Format(Hoja58.Cells(12, 2), "dd;@")
Me.Label37.Caption = " " & Format(Hoja58.Cells(12, 3), "dd;@")
Me.Label38.Caption = " " & Format(Hoja58.Cells(12, 4), "dd;@")
Me.Label39.Caption = " " & Format(Hoja58.Cells(12, 5), "dd;@")
Me.Label40.Caption = " " & Format(Hoja58.Cells(12, 6), "dd;@")
Me.Label41.Caption = " " & Format(Hoja58.Cells(12, 7), "dd;@")
Me.Label42.Caption = " " & Format(Hoja58.Cells(12, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label43.Caption = " " & Format(Hoja58.Cells(14, 2), "dd;@")
Me.Label44.Caption = " " & Format(Hoja58.Cells(14, 3), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox1.Text = Format(Hoja58.Cells(5, 2), "hh:mm")
Me.TextBox2.Text = Format(Hoja58.Cells(5, 3), "hh:mm")
Me.TextBox3.Text = Format(Hoja58.Cells(5, 4), "hh:mm")
Me.TextBox4.Text = Format(Hoja58.Cells(5, 5), "hh:mm")
Me.TextBox5.Text = Format(Hoja58.Cells(5, 6), "hh:mm")
Me.TextBox6.Text = Format(Hoja58.Cells(5, 7), "hh:mm")
Me.TextBox7.Text = Format(Hoja58.Cells(5, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox8.Text = Format(Hoja58.Cells(7, 2), "hh:mm")
Me.TextBox9.Text = Format(Hoja58.Cells(7, 3), "hh:mm")
Me.TextBox10.Text = Format(Hoja58.Cells(7, 4), "hh:mm")
Me.TextBox11.Text = Format(Hoja58.Cells(7, 5), "hh:mm")
Me.TextBox12.Text = Format(Hoja58.Cells(7, 6), "hh:mm")
Me.TextBox13.Text = Format(Hoja58.Cells(7, 7), "hh:mm")
Me.TextBox14.Text = Format(Hoja58.Cells(7, 8), "hh:mm")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox15.Text = Format(Hoja58.Cells(9, 2), "hh:mm")
Me.TextBox16.Text = Format(Hoja58.Cells(9, 3), "hh:mm")
Me.TextBox17.Text = Format(Hoja58.Cells(9, 4), "hh:mm")
Me.TextBox18.Text = Format(Hoja58.Cells(9, 5), "hh:mm")
Me.TextBox19.Text = Format(Hoja58.Cells(9, 6), "hh:mm")
Me.TextBox20.Text = Format(Hoja58.Cells(9, 7), "hh:mm")
Me.TextBox21.Text = Format(Hoja58.Cells(9, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox22.Text = Format(Hoja58.Cells(11, 2), "hh:mm")
Me.TextBox23.Text = Format(Hoja58.Cells(11, 3), "hh:mm")
Me.TextBox24.Text = Format(Hoja58.Cells(11, 4), "hh:mm")
Me.TextBox25.Text = Format(Hoja58.Cells(11, 5), "hh:mm")
Me.TextBox26.Text = Format(Hoja58.Cells(11, 6), "hh:mm")
Me.TextBox27.Text = Format(Hoja58.Cells(11, 7), "hh:mm")
Me.TextBox28.Text = Format(Hoja58.Cells(11, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox29.Text = Format(Hoja58.Cells(13, 2), "hh:mm")
Me.TextBox30.Text = Format(Hoja58.Cells(13, 3), "hh:mm")
Me.TextBox31.Text = Format(Hoja58.Cells(13, 4), "hh:mm")
Me.TextBox32.Text = Format(Hoja58.Cells(13, 5), "hh:mm")
Me.TextBox33.Text = Format(Hoja58.Cells(13, 6), "hh:mm")
Me.TextBox34.Text = Format(Hoja58.Cells(13, 7), "hh:mm")
Me.TextBox35.Text = Format(Hoja58.Cells(13, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox36.Text = Format(Hoja58.Cells(15, 2), "hh:mm")
Me.TextBox37.Text = Format(Hoja58.Cells(15, 3), "hh:mm")

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

   
    If TextBox1 = Entrada Or TextBox1 = Salida Then
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

Private Sub ToggleButton1_Click()

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


 With frm_Calendario_Asistencia.cboMes
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
    
    frm_Calendario_Asistencia.cboMes.ListIndex = VBA.Month(VBA.Date) - 1
       
    frm_Calendario_Asistencia.SpinButton2.Value = VBA.Year(VBA.Date)
    
    frm_Calendario_Asistencia.label_año2.Caption = VBA.Year(VBA.Date)
        

Hoja58.Cells(3, 11) = VBA.Month(VBA.Date)

Me.Label47.Caption = Hoja58.Cells(6, 11)
Me.Label46.Caption = "-  " & Hoja58.Cells(6, 12)

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


'''''''''''''''''''''''''''''''''''''
Me.Label8.Caption = " " & Format(Hoja58.Cells(4, 2), "dd;@")
Me.Label9.Caption = " " & Format(Hoja58.Cells(4, 3), "dd;@")
Me.Label10.Caption = " " & Format(Hoja58.Cells(4, 4), "dd;@")
Me.Label11.Caption = " " & Format(Hoja58.Cells(4, 5), "dd;@")
Me.Label12.Caption = " " & Format(Hoja58.Cells(4, 6), "dd;@")
Me.Label13.Caption = " " & Format(Hoja58.Cells(4, 7), "dd;@")
Me.Label14.Caption = " " & Format(Hoja58.Cells(4, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label15.Caption = " " & Format(Hoja58.Cells(6, 2), "dd;@")
Me.Label16.Caption = " " & Format(Hoja58.Cells(6, 3), "dd;@")
Me.Label17.Caption = " " & Format(Hoja58.Cells(6, 4), "dd;@")
Me.Label18.Caption = " " & Format(Hoja58.Cells(6, 5), "dd;@")
Me.Label19.Caption = " " & Format(Hoja58.Cells(6, 6), "dd;@")
Me.Label20.Caption = " " & Format(Hoja58.Cells(6, 7), "dd;@")
Me.Label21.Caption = " " & Format(Hoja58.Cells(6, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label22.Caption = " " & Format(Hoja58.Cells(8, 2), "dd;@")
Me.Label23.Caption = " " & Format(Hoja58.Cells(8, 3), "dd;@")
Me.Label24.Caption = " " & Format(Hoja58.Cells(8, 4), "dd;@")
Me.Label25.Caption = " " & Format(Hoja58.Cells(8, 5), "dd;@")
Me.Label26.Caption = " " & Format(Hoja58.Cells(8, 6), "dd;@")
Me.Label27.Caption = " " & Format(Hoja58.Cells(8, 7), "dd;@")
Me.Label28.Caption = " " & Format(Hoja58.Cells(8, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label29.Caption = " " & Format(Hoja58.Cells(10, 2), "dd;@")
Me.Label30.Caption = " " & Format(Hoja58.Cells(10, 3), "dd;@")
Me.Label31.Caption = " " & Format(Hoja58.Cells(10, 4), "dd;@")
Me.Label32.Caption = " " & Format(Hoja58.Cells(10, 5), "dd;@")
Me.Label33.Caption = " " & Format(Hoja58.Cells(10, 6), "dd;@")
Me.Label34.Caption = " " & Format(Hoja58.Cells(10, 7), "dd;@")
Me.Label35.Caption = " " & Format(Hoja58.Cells(10, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label36.Caption = " " & Format(Hoja58.Cells(12, 2), "dd;@")
Me.Label37.Caption = " " & Format(Hoja58.Cells(12, 3), "dd;@")
Me.Label38.Caption = " " & Format(Hoja58.Cells(12, 4), "dd;@")
Me.Label39.Caption = " " & Format(Hoja58.Cells(12, 5), "dd;@")
Me.Label40.Caption = " " & Format(Hoja58.Cells(12, 6), "dd;@")
Me.Label41.Caption = " " & Format(Hoja58.Cells(12, 7), "dd;@")
Me.Label42.Caption = " " & Format(Hoja58.Cells(12, 8), "dd;@")

'''''''''''''''''''''''''''''''''''''
Me.Label43.Caption = " " & Format(Hoja58.Cells(14, 2), "dd;@")
Me.Label44.Caption = " " & Format(Hoja58.Cells(14, 3), "dd;@")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox1.Text = Format(Hoja58.Cells(5, 2), "hh:mm")
Me.TextBox2.Text = Format(Hoja58.Cells(5, 3), "hh:mm")
Me.TextBox3.Text = Format(Hoja58.Cells(5, 4), "hh:mm")
Me.TextBox4.Text = Format(Hoja58.Cells(5, 5), "hh:mm")
Me.TextBox5.Text = Format(Hoja58.Cells(5, 6), "hh:mm")
Me.TextBox6.Text = Format(Hoja58.Cells(5, 7), "hh:mm")
Me.TextBox7.Text = Format(Hoja58.Cells(5, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox8.Text = Format(Hoja58.Cells(7, 2), "hh:mm")
Me.TextBox9.Text = Format(Hoja58.Cells(7, 3), "hh:mm")
Me.TextBox10.Text = Format(Hoja58.Cells(7, 4), "hh:mm")
Me.TextBox11.Text = Format(Hoja58.Cells(7, 5), "hh:mm")
Me.TextBox12.Text = Format(Hoja58.Cells(7, 6), "hh:mm")
Me.TextBox13.Text = Format(Hoja58.Cells(7, 7), "hh:mm")
Me.TextBox14.Text = Format(Hoja58.Cells(7, 8), "hh:mm")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox15.Text = Format(Hoja58.Cells(9, 2), "hh:mm")
Me.TextBox16.Text = Format(Hoja58.Cells(9, 3), "hh:mm")
Me.TextBox17.Text = Format(Hoja58.Cells(9, 4), "hh:mm")
Me.TextBox18.Text = Format(Hoja58.Cells(9, 5), "hh:mm")
Me.TextBox19.Text = Format(Hoja58.Cells(9, 6), "hh:mm")
Me.TextBox20.Text = Format(Hoja58.Cells(9, 7), "hh:mm")
Me.TextBox21.Text = Format(Hoja58.Cells(9, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox22.Text = Format(Hoja58.Cells(11, 2), "hh:mm")
Me.TextBox23.Text = Format(Hoja58.Cells(11, 3), "hh:mm")
Me.TextBox24.Text = Format(Hoja58.Cells(11, 4), "hh:mm")
Me.TextBox25.Text = Format(Hoja58.Cells(11, 5), "hh:mm")
Me.TextBox26.Text = Format(Hoja58.Cells(11, 6), "hh:mm")
Me.TextBox27.Text = Format(Hoja58.Cells(11, 7), "hh:mm")
Me.TextBox28.Text = Format(Hoja58.Cells(11, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox29.Text = Format(Hoja58.Cells(13, 2), "hh:mm")
Me.TextBox30.Text = Format(Hoja58.Cells(13, 3), "hh:mm")
Me.TextBox31.Text = Format(Hoja58.Cells(13, 4), "hh:mm")
Me.TextBox32.Text = Format(Hoja58.Cells(13, 5), "hh:mm")
Me.TextBox33.Text = Format(Hoja58.Cells(13, 6), "hh:mm")
Me.TextBox34.Text = Format(Hoja58.Cells(13, 7), "hh:mm")
Me.TextBox35.Text = Format(Hoja58.Cells(13, 8), "hh:mm")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.TextBox36.Text = Format(Hoja58.Cells(15, 2), "hh:mm")
Me.TextBox37.Text = Format(Hoja58.Cells(15, 3), "hh:mm")

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

   ' AMARILLO = &HD1D7FE                              'A - rojo
   ' ROJO = &HD1D7FE                         '     'rojo
   
    If TextBox1 = Entrada Or TextBox1 = Salida Then
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



