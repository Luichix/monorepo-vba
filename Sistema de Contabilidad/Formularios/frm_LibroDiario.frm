VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_LibroDiario 
   Caption         =   "Libro Diario"
   ClientHeight    =   9225.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   9630.001
   OleObjectBlob   =   "frm_LibroDiario.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_LibroDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Long 'Representa los items en el ListBox de este formulario

Private Sub btn_Editar_Click()

With Me.lbx_DebeHaber

    If .ListIndex = -1 Then
        MsgBox "Seleccione un registro para editar", vbInformation
        Exit Sub
    End If

    If Me.btn_Editar.Caption = "Editar" Then

        If MsgBox("¿Seguro que quiere modificar esta operación?", vbQuestion + vbYesNo) = vbYes Then

            .Locked = True

            Call OcultarBotonesRestantes

            Me.cbo_CodCuenta.Value = .List(.ListIndex, 0)
            Me.txt_Concepto.Text = .List(.ListIndex, 4)

                If Me.opt_Cargar.Value = True Then
                   Me.txt_Monto.Value = .List(.ListIndex, 2)
                Else
                    Me.txt_Monto.Value = .List(.ListIndex, 3)
                End If

                Me.btn_Editar.Caption = "Guardar cambios"
                Exit Sub
        Else
            .ListIndex = -1
            Exit Sub
        End If
    End If



                    If Me.btn_Editar.Caption = "Guardar cambios" Then
                        .Locked = False

                        .List(.ListIndex, 0) = Me.cbo_CodCuenta.Value
                        .List(.ListIndex, 1) = Me.txt_NombreCuenta.Text
                        .List(.ListIndex, 4) = Me.txt_Concepto.Text


                            If Me.opt_Cargar.Value = True Then
                                .List(.ListIndex, 2) = Me.txt_Monto.Value
                                Call sumarDebe
                            Else
                                .List(.ListIndex, 3) = Me.txt_Monto.Value
                                Call sumarhaber
                            End If

                            Call LimpiarControlesContables("libro_diario", Me)
                            MsgBox "Cambios guardados satisfactoriamente!", vbInformation
                            Me.btn_Editar.Caption = "Editar"
                            Call MostrarBotonesRestantes
                    End If

End With

End Sub

Private Sub btn_Eliminar_Click()

With Me.lbx_DebeHaber

    If .ListIndex = -1 Then
        MsgBox "Seleccione un registro para eliminar", vbInformation
        Exit Sub
    End If

    If MsgBox("¿Seguro que quiere eliminar esta operación?", vbQuestion + vbYesNo) = vbYes Then


                If Me.opt_Cargar.Value = True Then
                    .RemoveItem (.ListIndex)
                    .ListIndex = -1
                    Call sumarDebe
                Else
                    .RemoveItem (.ListIndex)
                    .ListIndex = -1
                    Call sumarhaber
                End If

                i = i - 1

            MsgBox "La operación ha sido eliminada!", vbInformation

        Else
                .ListIndex = -1
                Exit Sub
    End If


        If .ListCount = Empty Then
            Call OcultarBotonesEditarEliminar
        End If
End With

End Sub

Private Sub btn_LimpiarItems_Click()
    Call LimpiarControlesContables("libro_diario", Me)
    Call LimpiarItems
    Call OcultarBotonesEditarEliminar
End Sub

Private Sub btn_Registrar_Click()


If Me.btn_Editar.Caption = "Guardar cambios" Then
    MsgBox "Debe guardar los cambios realizados", vbInformation
    Exit Sub
End If

If ControlesVaciosContables("libro_diario", Me, Frame1, True) = True Then Exit Sub

        If Me.chk_ISR.Value = True Then
            Call CalcularRetencionISR
        End If

        If Me.chk_IVA.Value = True Then
            Call CalcularIVA
        End If


With Me

        'Busca un item en el ListBox, que si está repetido, no permite agregarlo nuevamente
        'obligando al usuario a seleccionar uno diferente
        For i = 0 To .lbx_DebeHaber.ListCount - 1
            If .lbx_DebeHaber.List(i, 0) = .cbo_CodCuenta Then
                MsgBox "Esta cuenta ya se agregó, elija una diferente"
                .lbx_DebeHaber.ListIndex = i
                Exit Sub
            End If
        Next


        .lbx_DebeHaber.AddItem .cbo_CodCuenta.Value
        .lbx_DebeHaber.List(i, 1) = .txt_NombreCuenta.Text

        If .opt_Cargar.Value = True Then
            .lbx_DebeHaber.List(i, 2) = .txt_Monto.Value

            Call sumarDebe

        Else
            .lbx_DebeHaber.List(i, 3) = .txt_Monto.Value

            Call sumarhaber

        End If

            .lbx_DebeHaber.List(i, 4) = .txt_Concepto.Text

        i = i + 1


        .lbx_DebeHaber.ListIndex = -1 'Elimina la selección del ListBox

 End With

    Call LimpiarControlesContables("libro_diario", Me)
    Call MostrarBotonesEditarEliminar
    Me.chk_ISR.Value = False
    Me.chk_IVA.Value = False

End Sub

Private Sub btn_ListadoCuentas_Click()
banderaListadoCuentas = 2
    Call LanzarListadoCuentas(Me, "lbl_LanzarListadoCuentas")
End Sub

Private Sub btn_EnviarADiario_Click()
Dim Final As Long

On Error Resume Next

If Me.lbx_DebeHaber.ListCount = Empty Then
    MsgBox "No hay movimientos para procesar", vbInformation
    Exit Sub
End If

If Me.lbl_Diferencia.Caption <> 0 Then
    MsgBox "La partida aún no está cuadrada!", vbCritical
    Exit Sub
End If



Final = nReg(Hoja3, 2, 3)


With Hoja3

    If MsgBox("¿Seguro que desea continuar?", vbQuestion + vbYesNo) = vbYes Then

            Application.ScreenUpdating = False
            


        For i = 0 To Me.lbx_DebeHaber.ListCount - 1
        
         Hoja3.Select
                Hoja3.Rows("2:2").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                
            .Cells(2, 1) = Hoja3.Cells(3, 1) + 1
            .Cells(2, 2) = Me.txt_Asiento.Value
            .Cells(2, 3) = CDate(Me.txt_Fecha)
                
            .Cells(2, 4) = Me.lbx_DebeHaber.List(i, 4) ' Concepto
            .Cells(2, 5) = Me.lbx_DebeHaber.List(i, 0) ' Cuenta
            .Cells(2, 6) = Me.lbx_DebeHaber.List(i, 1) ' Nombre de Cuenta

            ' DEBE
            Me.lbx_DebeHaber.List(i, 2) = _
            Replace(Me.lbx_DebeHaber.List(i, 2), Application.ThousandsSeparator, "")  ' Elimino el separador de miles
            Me.lbx_DebeHaber.List(i, 2) = _
            Replace(Me.lbx_DebeHaber.List(i, 2), Application.DecimalSeparator, ".")  'sustituyo el separador decimal

            .Cells(2, 7) = Me.lbx_DebeHaber.List(i, 2) ' Debe
            .Cells(2, 7).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

            ' HABER
            Me.lbx_DebeHaber.List(i, 3) = _
            Replace(Me.lbx_DebeHaber.List(i, 3), Application.ThousandsSeparator, "")  'elimino el separador de miles
            Me.lbx_DebeHaber.List(i, 3) = _
            Replace(Me.lbx_DebeHaber.List(i, 3), Application.DecimalSeparator, ".")  'Sustituyo el separador decimal

            .Cells(2, 8) = Me.lbx_DebeHaber.List(i, 3) ' Haber
            .Cells(2, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            
            .Cells(2, 9) = Hoja91.Range("G1").Text

            Final = Final + 1
        Next



            Application.ScreenUpdating = True

        Else
            Exit Sub
    End If

End With

    Call LimpiarItems
    Call CorrelativoPartidas
    Call OcultarBotonesEditarEliminar

End Sub

Private Sub btn_Salir_Click()
    Unload Me
End Sub

Private Sub cbo_CodCuenta_Change()
Dim Fila As Long
Dim Final As Long
Dim encontrado As Boolean

Me.cbo_CodCuenta.BackColor = RGB(255, 255, 255)


If Me.cbo_CodCuenta = Empty Then
    Me.txt_NombreCuenta = Empty
    Exit Sub
End If


Final = nReg(Hoja2, 2, 1) - 1

    For Fila = 2 To Final
        If Hoja2.Cells(Fila, 1) = Val(Me.cbo_CodCuenta) Then
            encontrado = True
            Me.txt_NombreCuenta = Hoja2.Cells(Fila, 2)
            Exit For
        End If
    Next

    If encontrado = False Then
        Me.txt_NombreCuenta = Empty
    End If

End Sub

Private Sub cbo_CodCuenta_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As Long

Do While Me.cbo_CodCuenta.ListCount > 0
    Me.cbo_CodCuenta.RemoveItem 0
Loop

    Final = nReg(Hoja2, 2, 1) - 1

        For Fila = 2 To Final
            Lista = Hoja2.Cells(Fila, 1)
            If Lista > 1000 Then
            Me.cbo_CodCuenta.AddItem Lista
            End If
        Next

End Sub

Private Sub cbo_CodCuenta_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub chk_ISR_Click()

    If Me.chk_ISR.Value = True Then
        Me.chk_IVA.Value = False
    End If


End Sub

Private Sub chk_IVA_Click()

    If Me.chk_IVA.Value = True Then
        Me.chk_ISR.Value = False
    End If

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub lbx_DebeHaber_Click()
On Error GoTo optAbonar

        If Me.lbx_DebeHaber.List(Me.lbx_DebeHaber.ListIndex, 2) Then
                Me.opt_Cargar = True
            Else
optAbonar:
                Me.opt_Abonar = True
        End If

End Sub

Private Sub opt_Abonar_Change()
    Me.chk_ISR.Value = False
    Me.chk_ISR.Visible = False
End Sub

Private Sub opt_Cargar_Change()
    Me.chk_ISR.Value = False
    Me.chk_ISR.Visible = True
End Sub

Private Sub txt_Concepto_Change()
Me.txt_Concepto.BackColor = RGB(255, 255, 255)
Me.txt_Concepto = UCase(Me.txt_Concepto)
End Sub

Private Sub btn_LanzarCalendario_Click()
banderaCalendario = 1
    Call LanzarCalendario(Me, "txt_Fecha")
End Sub

Private Sub txt_Monto_Change()
Me.txt_Monto.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txt_Monto_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
Me.txt_Monto = FormatNumber(Me.txt_Monto, 2)
End Sub

Private Sub txt_Monto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Application.DecimalSeparator = "." Then
    If KeyAscii <> 46 And KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
Else
    If KeyAscii <> 44 And KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txt_NombreCuenta_Change()
Me.txt_NombreCuenta.BackColor = RGB(255, 255, 255)
End Sub

Private Sub UserForm_Activate()
Me.txt_Fecha.Text = Date
Me.cbo_CodCuenta.SetFocus
End Sub


Private Sub UserForm_Initialize()

    Call OcultarBotonesEditarEliminar

    Call CorrelativoPartidas

    Me.lbx_DebeHaber.ColumnCount = 5

    Me.lbx_DebeHaber.ColumnWidths = "45 pt;250 pt;75 pt;75 pt;0 pt"
End Sub

Private Sub LimpiarItems()
    Me.lbx_DebeHaber.Clear
    i = 0
    Me.cbo_CodCuenta.BackColor = RGB(255, 255, 255)
    Me.txt_NombreCuenta.BackColor = RGB(255, 255, 255)
    Me.txt_Monto.BackColor = RGB(255, 255, 255)
    Me.txt_Concepto.BackColor = RGB(255, 255, 255)
    Me.lbl_SumaDebe.Caption = "0.00"
    Me.lbl_SumaHaber.Caption = "0.00"
    Me.lbl_Diferencia.Caption = "0.00"
    Me.lbl_Diferencia.ForeColor = RGB(255, 255, 255)
    Me.opt_Cargar.Value = True
    Me.chk_IVA.Value = False
    Me.chk_ISR.Value = False
End Sub

Private Sub CorrelativoPartidas()
Dim Final As Long

        Final = nReg(Hoja3, 2, 3) - 1

        If Hoja3.Cells(2, 2) = Empty Then
                Me.txt_Asiento.Value = 1
            Else
                Me.txt_Asiento.Value = Hoja3.Cells(2, 2) + 1
        End If
End Sub

Private Sub CalcularRetencionISR()
Dim nCodigoCta As Long
Dim sNombreCta As String
Dim sConcepto As String
Dim valorISR As Currency

valorISR = 0

With Me.lbx_DebeHaber

        valorISR = (Me.txt_Monto.Value / 100) * 10



        If Me.opt_Cargar.Value = True Then
            nCodigoCta = 1160202
            sNombreCta = "RETENCIÓN ISR 10%"
            sConcepto = "IMPUESTO SOBRE LA RENTA RETENIDO SEGÚN ARTÍCULO 156"

            .AddItem nCodigoCta
            .List(i, 1) = sNombreCta
            .List(i, 2) = FormatNumber(-valorISR, 2)
            .List(i, 4) = sConcepto

        End If

        i = i + 1

End With

End Sub

Private Sub CalcularIVA()
Dim nCodigoCta1 As Long
Dim nCodigoCta2 As Long
Dim sNombreCta1 As String
Dim sNombreCta2 As String
Dim sConcepto1 As String
Dim sConcepto2 As String
Dim valorIVA As Currency

valorIVA = 0

With Me.lbx_DebeHaber

        valorIVA = (Me.txt_Monto.Value / 100) * 15


        If Me.opt_Cargar.Value = True Then
            nCodigoCta1 = 1170101
            sNombreCta1 = "IVA POR PAGAR 15%"
            sConcepto1 = "CRÉDITO FISCAL"

            .AddItem nCodigoCta1
            .List(i, 1) = sNombreCta1
            .List(i, 2) = FormatNumber(valorIVA, 2)
            .List(i, 4) = sConcepto1

        Else
             nCodigoCta2 = 20201
            sNombreCta2 = "IVA ACREDITABLE 15%"
            sConcepto2 = "DÉBITO FISCAL"

            .AddItem nCodigoCta2
            .List(i, 1) = sNombreCta2
            .List(i, 3) = FormatNumber(valorIVA, 2)
            .List(i, 4) = sConcepto2
            .List(i, 4) = sConcepto2

        End If

        i = i + 1

End With

End Sub

Private Sub OcultarBotonesEditarEliminar()
    Me.btn_Editar.Visible = False
    Me.btn_Eliminar.Visible = False
End Sub

Private Sub MostrarBotonesEditarEliminar()
    Me.btn_Editar.Visible = True
    Me.btn_Eliminar.Visible = True
End Sub

Private Sub OcultarBotonesRestantes()
    Me.btn_Eliminar.Visible = False
    Me.btn_Registrar.Visible = False
    Me.btn_EnviarADiario.Visible = False
    Me.btn_LimpiarItems.Visible = False
    Me.chk_ISR.Visible = False
    Me.chk_IVA.Visible = False
End Sub

Private Sub MostrarBotonesRestantes()
    Me.btn_Eliminar.Visible = True
    Me.btn_Registrar.Visible = True
    Me.btn_EnviarADiario.Visible = True
    Me.btn_LimpiarItems.Visible = True

        If Me.opt_Cargar.Value = True Then
            Me.chk_ISR.Visible = True
        End If

    Me.chk_IVA.Visible = True
End Sub

