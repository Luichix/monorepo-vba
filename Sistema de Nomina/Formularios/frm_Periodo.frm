VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Periodo 
   Caption         =   "Seleccione una fecha"
   ClientHeight    =   3630
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   4540
   OleObjectBlob   =   "frm_Periodo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Periodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit

Private Sub cboMes_Click()
    ModuloPeriodo.CambioDeMesPeriodo
End Sub



Private Sub Frame1_Click()

End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Lbl1.Font.Bold = True
   Lbl1.ForeColor = VBA.RGB(255, 0, 0)
   
   Lbl2.Font.Bold = True
   Lbl2.ForeColor = VBA.RGB(17, 114, 155)


End Sub
Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Lbl2.Font.Bold = True
   Lbl2.ForeColor = VBA.RGB(255, 0, 0)
   
   Lbl1.Font.Bold = True
   Lbl1.ForeColor = VBA.RGB(17, 114, 155)


End Sub

Private Sub lbl1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.Lbl1.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.Lbl1.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl1 As Control
    Set Control_lbl1 = frm_Periodo.Lbl1
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl1)
End Sub

Private Sub lbl10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl10.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl10.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl10 As Control
    Set Control_lbl10 = frm_Periodo.lbl10
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl10)
End Sub

Private Sub lbl11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl11.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl11.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl11 As Control
    Set Control_lbl11 = frm_Periodo.lbl11
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl11)
End Sub

Private Sub lbl12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl12.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl12.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl12 As Control
    Set Control_lbl12 = frm_Periodo.lbl12
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl12)
End Sub

Private Sub lbl13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl13.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl13.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl13 As Control
    Set Control_lbl13 = frm_Periodo.lbl13
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl13)
End Sub

Private Sub lbl14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl14.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl14.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl14 As Control
    Set Control_lbl14 = frm_Periodo.lbl14
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl14)
End Sub

Private Sub lbl15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl15.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl15.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl15 As Control
    Set Control_lbl15 = frm_Periodo.lbl15
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl15)
End Sub

Private Sub lbl16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl16.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl16.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl16 As Control
    Set Control_lbl16 = frm_Periodo.lbl16
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl16)
End Sub

Private Sub lbl17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl17.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl17.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl17 As Control
    Set Control_lbl17 = frm_Periodo.lbl17
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl17)
End Sub

Private Sub lbl18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl18.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl18.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl18 As Control
    Set Control_lbl18 = frm_Periodo.lbl18
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl18)
End Sub

Private Sub lbl19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl19.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl19.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl19 As Control
    Set Control_lbl19 = frm_Periodo.lbl19
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl19)
End Sub

Private Sub lbl2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.Lbl2.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.Lbl2.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl2 As Control
    Set Control_lbl2 = frm_Periodo.Lbl2
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl2)
End Sub

Private Sub lbl20_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl20.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl20.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl20 As Control
    Set Control_lbl20 = frm_Periodo.lbl20
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl20)
End Sub

Private Sub lbl21_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl21.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl21.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl21 As Control
    Set Control_lbl21 = frm_Periodo.lbl21
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl21)
End Sub

Private Sub lbl22_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl22.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl22.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl22 As Control
    Set Control_lbl22 = frm_Periodo.lbl22
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl22)
End Sub

Private Sub lbl23_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl23.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl23.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl23 As Control
    Set Control_lbl23 = frm_Periodo.lbl23
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl23)
End Sub

Private Sub lbl24_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl24.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl24.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl24 As Control
    Set Control_lbl24 = frm_Periodo.lbl24
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl24)
End Sub

Private Sub lbl25_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl25.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl25.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl25 As Control
    Set Control_lbl25 = frm_Periodo.lbl25
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl25)
End Sub

Private Sub lbl26_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl26.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl26.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl26 As Control
    Set Control_lbl26 = frm_Periodo.lbl26
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl26)
End Sub

Private Sub lbl27_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl27.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl27.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl27 As Control
    Set Control_lbl27 = frm_Periodo.lbl27
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl27)
End Sub

Private Sub lbl28_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl28.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl28.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl28 As Control
    Set Control_lbl28 = frm_Periodo.lbl28
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl28)
End Sub

Private Sub lbl29_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl29.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl29.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl29 As Control
    Set Control_lbl29 = frm_Periodo.lbl29
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl29)
End Sub

Private Sub lbl3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.Lbl3.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.Lbl3.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl3 As Control
    Set Control_lbl3 = frm_Periodo.Lbl3
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl3)
End Sub

Private Sub lbl30_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl30.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl30.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl30 As Control
    Set Control_lbl30 = frm_Periodo.lbl30
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl30)
End Sub

Private Sub lbl31_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl31.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl31.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl31 As Control
    Set Control_lbl31 = frm_Periodo.lbl31
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl31)
End Sub

Private Sub lbl32_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl32.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl32.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl32 As Control
    Set Control_lbl32 = frm_Periodo.lbl32
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl32)
End Sub

Private Sub lbl33_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl33.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl33.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl33 As Control
    Set Control_lbl33 = frm_Periodo.lbl33
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl33)
End Sub

Private Sub lbl34_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl34.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl34.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl34 As Control
    Set Control_lbl34 = frm_Periodo.lbl34
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl34)
End Sub

Private Sub lbl35_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl35.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl35.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl35 As Control
    Set Control_lbl35 = frm_Periodo.lbl35
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl35)
End Sub

Private Sub lbl36_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl36.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl36.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl36 As Control
    Set Control_lbl36 = frm_Periodo.lbl36
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl36)
End Sub

Private Sub lbl37_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl37.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl37.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl37 As Control
    Set Control_lbl37 = frm_Periodo.lbl37
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl37)
End Sub

Private Sub lbl38_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl38.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl38.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl38 As Control
    Set Control_lbl38 = frm_Periodo.lbl38
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl38)
End Sub

Private Sub lbl39_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl39.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl39.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl39 As Control
    Set Control_lbl39 = frm_Periodo.lbl39
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl39)
End Sub

Private Sub lbl4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl4.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl4.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl4 As Control
    Set Control_lbl4 = frm_Periodo.lbl4
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl4)
End Sub

Private Sub lbl40_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl40.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl40.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl40 As Control
    Set Control_lbl40 = frm_Periodo.lbl40
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl40)
End Sub

Private Sub lbl41_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl41.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl41.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl41 As Control
    Set Control_lbl41 = frm_Periodo.lbl41
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl41)
End Sub

Private Sub lbl42_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl42.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl42.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl42 As Control
    Set Control_lbl42 = frm_Periodo.lbl42
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl42)
End Sub

Private Sub lbl5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl5.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl5.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl5 As Control
    Set Control_lbl5 = frm_Periodo.lbl5
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl5)
End Sub

Private Sub lbl6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl6.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl6.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl6 As Control
    Set Control_lbl6 = frm_Periodo.lbl6
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl6)
End Sub

Private Sub lbl7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl7.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl7.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl7 As Control
    Set Control_lbl7 = frm_Periodo.lbl7
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl7)
End Sub

Private Sub lbl8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl8.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl8.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl8 As Control
    Set Control_lbl8 = frm_Periodo.lbl8
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl8)
End Sub

Private Sub lbl9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frm_Periodo.lbl9.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frm_Periodo.lbl9.Caption)
        Mes = VBA.CLng(frm_Periodo.cboMes.Value)
        Ano = VBA.CLng(frm_Periodo.lblAno.Caption)
        
        Unload frm_Periodo
        Call ModuloPeriodo.RecibeElPeriodo(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control_lbl9 As Control
    Set Control_lbl9 = frm_Periodo.lbl9
    
    Call ModuloPeriodo.MarcarDiaPeriodo(Control_lbl9)
End Sub

Private Sub cmdSalirConEscapePeriodo_Click()
    Call ModuloPeriodo.SalirConEscapePeriodo
End Sub

Private Sub lblHoy_Click()
    ModuloPeriodo.UnClickEnHoyEsPeriodo
End Sub

Private Sub mrcDias_Click()

End Sub

Private Sub spbAño_Change()
    ModuloPeriodo.CambioDeAnoPeriodo
End Sub

Private Sub UserForm_Initialize()
    Call ModuloPeriodo.InicializaFormularioPeriodo
End Sub
