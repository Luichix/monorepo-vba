VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_pieza 
   Caption         =   "REGISTRO DE PIEZAS"
   ClientHeight    =   10404
   ClientLeft      =   20
   ClientTop       =   310
   ClientWidth     =   18290
   OleObjectBlob   =   "frm_pieza.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_pieza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_caja_Click()
 banderaCaja = 1
    Call LanzarListadoCaja(Me, "label41")
End Sub

Private Sub btn_cestado_Click()

End Sub

Private Sub btn_Fecha_Click()
 banderaCalendario = 2
    Call LanzarCalendario(Me, "lbl_fecha")
    Me.txt_f1.SetFocus
End Sub

Private Sub btn_fija_Click()
        
        
        If Me.txt_f0.Value = "" And Me.txt_f1.Text = "" And Me.txt_f2.Text = "" Then
            MsgBox ("Hay datos vacios en el registro")
               Exit Sub
        End If
        
    
        If Me.txt_Fecha = "" Or Me.txt_id = "" Or Me.txt_caja = "" Or Me.txt_pieza = "" Then
            MsgBox ("Hay datos vacios en el registro")
               Exit Sub
        Else
        
    
            For x = 0 To lbx_pieza.ListCount - 1
            If txt_f2.Text = "" Then
            
               If txt_f1.Text = "" Then
                    If lbx_pieza.List(x, 1) = txt_id.Text & "-" & UCase(txt_f0.Text) Then
                    MsgBox "Esta pieza ya se agregó, elija una diferente"
                    lbx_pieza.ListIndex = x
                    Exit Sub
                End If
            
                Else
                If lbx_pieza.List(x, 1) = txt_id.Text & "-" & UCase(txt_f0.Text) & txt_f1.Text Then
                    MsgBox "Esta pieza ya se agregó, elija una diferente"
                    lbx_pieza.ListIndex = x
                    Exit Sub
                End If
              
              End If

                        
            
                If lbx_pieza.List(x, 1) = txt_id.Text & "-" & UCase(txt_f0.Text) & txt_f1.Text & "-" & txt_f2.Text Then
                    MsgBox "Esta pieza ya se agregó, elija una diferente"
                    lbx_pieza.ListIndex = x
                    Exit Sub
                End If
            
            End If
            
            Next
            
            
    
            lbx_pieza.AddItem
            lbx_pieza.List(x, 0) = txt_Fecha
            
            If txt_f2.Text = "" Then
            
                        If txt_f1.Text = "" Then
            
                            lbx_pieza.List(x, 1) = txt_id.Text & "-" & UCase(txt_f0.Text)
                            lbx_pieza.List(x, 2) = txt_pieza.Text & " " & UCase(txt_f0.Text)
                        Else
                        
                            lbx_pieza.List(x, 1) = txt_id.Text & "-" & UCase(txt_f0.Text) & txt_f1.Text
                            lbx_pieza.List(x, 2) = txt_pieza.Text & " " & UCase(txt_f0.Text) & " - " & txt_f1.Text
                        End If

            Else
            
            lbx_pieza.List(x, 1) = txt_id.Text & "-" & UCase(txt_f0.Text) & txt_f1.Text & "-" & txt_f2.Text
            lbx_pieza.List(x, 2) = txt_pieza.Text & " " & UCase(txt_f0.Text) & " - " & txt_f1.Text & "-" & txt_f2.Text
            
            End If
            
            lbx_pieza.List(x, 3) = txt_fllave.Value
            lbx_pieza.List(x, 4) = txt_caja.Text
            lbx_pieza.List(x, 5) = txt_id.Text
            lbx_pieza.List(x, 6) = cbx_estado.Text
    
            x = x + 1
        
        End If

     
    Me.txt_f1.Value = Empty
    Me.txt_f2.Value = Empty
    Me.txt_fllave.Value = 1
    Me.txt_f0.Value = Empty
    
  
        lbx_pieza.ListIndex = -1
        lbx_pieza.ColumnCount = 7
        lbx_pieza.ColumnWidths = "45 pt;70 pt;150 pt;50 pt;150 pt;70 pt;70 pt"

End Sub

Private Sub btn_fllave_Click()
    If Me.txt_fllave = Empty Then
        Me.txt_fllave.Value = 1
    Else
        Me.txt_fllave.Value = Me.txt_fllave.Value + 1
    End If

End Sub

Private Sub btn_grabar_Click()
  
 If Me.txt_Fecha = "" Or Me.txt_id = "" Or Me.txt_caja = "" Or Me.txt_pieza = "" Then
            MsgBox ("Hay datos vacios en el registro")
               Exit Sub
 End If
  
 For i = 1 To 81
    If Controls("txt_p" & i) = "" Then
    
    Else
    
        For x = 0 To lbx_pieza.ListCount - 1
            If lbx_pieza.List(x, 1) = txt_id.Text & "-" & Controls("btn_p" & i).Caption Then
                MsgBox "Esta pieza ya se agregó, elija una diferente"
                lbx_pieza.ListIndex = x
                Exit Sub
            End If
        Next

        lbx_pieza.AddItem
        lbx_pieza.List(x, 0) = txt_Fecha
        lbx_pieza.List(x, 1) = txt_id.Text & "-" & Controls("btn_p" & i).Caption
        lbx_pieza.List(x, 2) = txt_pieza.Text & " - " & Controls("btn_p" & i).Caption
        lbx_pieza.List(x, 3) = Controls("txt_p" & i).Text
        lbx_pieza.List(x, 4) = txt_caja.Text
        lbx_pieza.List(x, 5) = txt_id.Text
        lbx_pieza.List(x, 6) = Me.cbx_estado.Text

        x = x + 1

    End If
 Next i
  
        lbx_pieza.ListIndex = -1
        lbx_pieza.ColumnCount = 7
        lbx_pieza.ColumnWidths = "45 pt;70 pt;150 pt;50 pt;150 pt;70 pt;70 pt"
        
        Limpiar
 
 
End Sub

Private Sub btn_Limpiar_Click()
    Limpiar
End Sub

Private Sub Limpiar()
  
    Me.txt_p1.Value = Empty
    Me.txt_p2.Value = Empty
    Me.txt_p3.Value = Empty
    Me.txt_p4.Value = Empty
    Me.txt_p5.Value = Empty
    Me.txt_p6.Value = Empty
    Me.txt_p7.Value = Empty
    Me.txt_p8.Value = Empty
    Me.txt_p9.Value = Empty
    Me.txt_p10.Value = Empty
    Me.txt_p11.Value = Empty
    Me.txt_p12.Value = Empty
    Me.txt_p13.Value = Empty
    Me.txt_p14.Value = Empty
    Me.txt_p15.Value = Empty
    Me.txt_p16.Value = Empty
    Me.txt_p17.Value = Empty
    Me.txt_p18.Value = Empty
    Me.txt_p19.Value = Empty
    Me.txt_p20.Value = Empty
    Me.txt_p21.Value = Empty
    Me.txt_p22.Value = Empty
    Me.txt_p23.Value = Empty
    Me.txt_p24.Value = Empty
    Me.txt_p25.Value = Empty
    Me.txt_p26.Value = Empty
    Me.txt_p27.Value = Empty
    Me.txt_p28.Value = Empty
    Me.txt_p29.Value = Empty
    Me.txt_p30.Value = Empty
    Me.txt_p31.Value = Empty
    Me.txt_p32.Value = Empty
    
        
    Me.txt_p33.Value = Empty
    Me.txt_p34.Value = Empty
    Me.txt_p35.Value = Empty
    Me.txt_p36.Value = Empty
    Me.txt_p37.Value = Empty
    Me.txt_p38.Value = Empty
    Me.txt_p39.Value = Empty
    Me.txt_p40.Value = Empty
    Me.txt_p41.Value = Empty
    Me.txt_p42.Value = Empty
    Me.txt_p43.Value = Empty
    Me.txt_p44.Value = Empty
    Me.txt_p45.Value = Empty
    Me.txt_p46.Value = Empty
    Me.txt_p47.Value = Empty
    Me.txt_p48.Value = Empty
    Me.txt_p49.Value = Empty
    Me.txt_p50.Value = Empty
    Me.txt_p51.Value = Empty
    Me.txt_p52.Value = Empty
    Me.txt_p53.Value = Empty
    Me.txt_p54.Value = Empty
    Me.txt_p55.Value = Empty
    Me.txt_p56.Value = Empty
    Me.txt_p57.Value = Empty
    Me.txt_p58.Value = Empty
    Me.txt_p59.Value = Empty
    Me.txt_p60.Value = Empty
    Me.txt_p61.Value = Empty
    Me.txt_p62.Value = Empty
    Me.txt_p63.Value = Empty
    Me.txt_p64.Value = Empty
    Me.txt_p65.Value = Empty
    Me.txt_p66.Value = Empty
    Me.txt_p67.Value = Empty
    Me.txt_p68.Value = Empty
    Me.txt_p69.Value = Empty
    Me.txt_p70.Value = Empty
    Me.txt_p71.Value = Empty
    Me.txt_p72.Value = Empty
    Me.txt_p73.Value = Empty
    Me.txt_p74.Value = Empty
    Me.txt_p75.Value = Empty
    Me.txt_p76.Value = Empty
    
    Me.txt_p77.Value = Empty
    Me.txt_p78.Value = Empty
    Me.txt_p79.Value = Empty
    Me.txt_p80.Value = Empty
    Me.txt_p81.Value = Empty
    
    Me.txt_fllave.Value = Empty
    Me.txt_f1.Value = Empty
    Me.txt_f2.Value = Empty
    
    
End Sub


Private Sub btn_pieza_Click()
    banderaCategoria = 1
    Call LanzarListadoCategoria(Me, "label41")
End Sub
Private Sub btn_p1_Click()
    If Me.txt_p1 = Empty Then
        Me.txt_p1.Value = 1
    Else
        Me.txt_p1.Value = Me.txt_p1.Value + 1
    End If
End Sub
Private Sub btn_p2_Click()
    If Me.txt_p2 = Empty Then
        Me.txt_p2.Value = 1
    Else
        Me.txt_p2.Value = Me.txt_p2.Value + 1
    End If
End Sub
Private Sub btn_p3_Click()
    If Me.txt_p3 = Empty Then
        Me.txt_p3.Value = 1
    Else
        Me.txt_p3.Value = Me.txt_p3.Value + 1
    End If
End Sub
Private Sub btn_p4_Click()
    If Me.txt_p4 = Empty Then
        Me.txt_p4.Value = 1
    Else
        Me.txt_p4.Value = Me.txt_p4.Value + 1
    End If
End Sub
Private Sub btn_p5_Click()
    If Me.txt_p5 = Empty Then
        Me.txt_p5.Value = 1
    Else
        Me.txt_p5.Value = Me.txt_p5.Value + 1
    End If
End Sub
Private Sub btn_p6_Click()
    If Me.txt_p6 = Empty Then
        Me.txt_p6.Value = 1
    Else
        Me.txt_p6.Value = Me.txt_p6.Value + 1
    End If
End Sub
Private Sub btn_p7_Click()
    If Me.txt_p7 = Empty Then
        Me.txt_p7.Value = 1
    Else
        Me.txt_p7.Value = Me.txt_p7.Value + 1
    End If
End Sub
Private Sub btn_p8_Click()
    If Me.txt_p8 = Empty Then
        Me.txt_p8.Value = 1
    Else
        Me.txt_p8.Value = Me.txt_p8.Value + 1
    End If
End Sub
Private Sub btn_p9_Click()
    If Me.txt_p9 = Empty Then
        Me.txt_p9.Value = 1
    Else
        Me.txt_p9.Value = Me.txt_p9.Value + 1
    End If
End Sub
Private Sub btn_p10_Click()
    If Me.txt_p10 = Empty Then
        Me.txt_p10.Value = 1
    Else
        Me.txt_p10.Value = Me.txt_p10.Value + 1
    End If
End Sub
Private Sub btn_p11_Click()
    If Me.txt_p11 = Empty Then
        Me.txt_p11.Value = 1
    Else
        Me.txt_p11.Value = Me.txt_p11.Value + 1
    End If
End Sub
Private Sub btn_p12_Click()
    If Me.txt_p12 = Empty Then
        Me.txt_p12.Value = 1
    Else
        Me.txt_p12.Value = Me.txt_p12.Value + 1
    End If
End Sub
Private Sub btn_p13_Click()
    If Me.txt_p13 = Empty Then
        Me.txt_p13.Value = 1
    Else
        Me.txt_p13.Value = Me.txt_p13.Value + 1
    End If
End Sub
Private Sub btn_p14_Click()
    If Me.txt_p14 = Empty Then
        Me.txt_p14.Value = 1
    Else
        Me.txt_p14.Value = Me.txt_p14.Value + 1
    End If
End Sub
Private Sub btn_p15_Click()
    If Me.txt_p15 = Empty Then
        Me.txt_p15.Value = 1
    Else
        Me.txt_p15.Value = Me.txt_p15.Value + 1
    End If
End Sub
Private Sub btn_p16_Click()
    If Me.txt_p16 = Empty Then
        Me.txt_p16.Value = 1
    Else
        Me.txt_p16.Value = Me.txt_p16.Value + 1
    End If
End Sub
Private Sub btn_p17_Click()
    If Me.txt_p17 = Empty Then
        Me.txt_p17.Value = 1
    Else
        Me.txt_p17.Value = Me.txt_p17.Value + 1
    End If
End Sub
Private Sub btn_p18_Click()
    If Me.txt_p18 = Empty Then
        Me.txt_p18.Value = 1
    Else
        Me.txt_p18.Value = Me.txt_p18.Value + 1
    End If
End Sub
Private Sub btn_p19_Click()
    If Me.txt_p19 = Empty Then
        Me.txt_p19.Value = 1
    Else
        Me.txt_p19.Value = Me.txt_p19.Value + 1
    End If
End Sub
Private Sub btn_p20_Click()
    If Me.txt_p20 = Empty Then
        Me.txt_p20.Value = 1
    Else
        Me.txt_p20.Value = Me.txt_p20.Value + 1
    End If
End Sub
Private Sub btn_p21_Click()
    If Me.txt_p21 = Empty Then
        Me.txt_p21.Value = 1
    Else
        Me.txt_p21.Value = Me.txt_p21.Value + 1
    End If
End Sub
Private Sub btn_p22_Click()
    If Me.txt_p22 = Empty Then
        Me.txt_p22.Value = 1
    Else
        Me.txt_p22.Value = Me.txt_p22.Value + 1
    End If
End Sub
Private Sub btn_p23_Click()
    If Me.txt_p23 = Empty Then
        Me.txt_p23.Value = 1
    Else
        Me.txt_p23.Value = Me.txt_p23.Value + 1
    End If
End Sub
Private Sub btn_p24_Click()
    If Me.txt_p24 = Empty Then
        Me.txt_p24.Value = 1
    Else
        Me.txt_p24.Value = Me.txt_p24.Value + 1
    End If
End Sub
Private Sub btn_p25_Click()
    If Me.txt_p25 = Empty Then
        Me.txt_p25.Value = 1
    Else
        Me.txt_p25.Value = Me.txt_p25.Value + 1
    End If
End Sub
Private Sub btn_p26_Click()
    If Me.txt_p26 = Empty Then
        Me.txt_p26.Value = 1
    Else
        Me.txt_p26.Value = Me.txt_p26.Value + 1
    End If
End Sub
Private Sub btn_p27_Click()
    If Me.txt_p27 = Empty Then
        Me.txt_p27.Value = 1
    Else
        Me.txt_p27.Value = Me.txt_p27.Value + 1
    End If
End Sub
Private Sub btn_p28_Click()
    If Me.txt_p28 = Empty Then
        Me.txt_p28.Value = 1
    Else
        Me.txt_p28.Value = Me.txt_p28.Value + 1
    End If
End Sub
Private Sub btn_p29_Click()
    If Me.txt_p29 = Empty Then
        Me.txt_p29.Value = 1
    Else
        Me.txt_p29.Value = Me.txt_p1.Value + 1
    End If
End Sub
Private Sub btn_p30_Click()
    If Me.txt_p30 = Empty Then
        Me.txt_p30.Value = 1
    Else
        Me.txt_p30.Value = Me.txt_p30.Value + 1
    End If
End Sub
Private Sub btn_p31_Click()
    If Me.txt_p31 = Empty Then
        Me.txt_p31.Value = 1
    Else
        Me.txt_p31.Value = Me.txt_p31.Value + 1
    End If
End Sub
Private Sub btn_p32_Click()
    If Me.txt_p32 = Empty Then
        Me.txt_p32.Value = 1
    Else
        Me.txt_p32.Value = Me.txt_p32.Value + 1
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub btn_p33_Click()
    If Me.txt_p33 = Empty Then
        Me.txt_p33.Value = 1
    Else
        Me.txt_p33.Value = Me.txt_p33.Value + 1
    End If
End Sub
Private Sub btn_p34_Click()
    If Me.txt_p34 = Empty Then
        Me.txt_p34.Value = 1
    Else
        Me.txt_p34.Value = Me.txt_p34.Value + 1
    End If
End Sub
Private Sub btn_p35_Click()
    If Me.txt_p35 = Empty Then
        Me.txt_p35.Value = 1
    Else
        Me.txt_p35.Value = Me.txt_p35.Value + 1
    End If
End Sub
Private Sub btn_p36_Click()
    If Me.txt_p36 = Empty Then
        Me.txt_p36.Value = 1
    Else
        Me.txt_p36.Value = Me.txt_p36.Value + 1
    End If
End Sub
Private Sub btn_p37_Click()
    If Me.txt_p37 = Empty Then
        Me.txt_p37.Value = 1
    Else
        Me.txt_p37.Value = Me.txt_p37.Value + 1
    End If
End Sub
Private Sub btn_p38_Click()
    If Me.txt_p38 = Empty Then
        Me.txt_p38.Value = 1
    Else
        Me.txt_p38.Value = Me.txt_p38.Value + 1
    End If
End Sub
Private Sub btn_p39_Click()
    If Me.txt_p39 = Empty Then
        Me.txt_p39.Value = 1
    Else
        Me.txt_p39.Value = Me.txt_p39.Value + 1
    End If
End Sub
Private Sub btn_p40_Click()
    If Me.txt_p40 = Empty Then
        Me.txt_p40.Value = 1
    Else
        Me.txt_p40.Value = Me.txt_p40.Value + 1
    End If
End Sub
Private Sub btn_p41_Click()
    If Me.txt_p41 = Empty Then
        Me.txt_p41.Value = 1
    Else
        Me.txt_p41.Value = Me.txt_p41.Value + 1
    End If
End Sub
Private Sub btn_p42_Click()
    If Me.txt_p42 = Empty Then
        Me.txt_p42.Value = 1
    Else
        Me.txt_p42.Value = Me.txt_p42.Value + 1
    End If
End Sub
Private Sub btn_p43_Click()
    If Me.txt_p43 = Empty Then
        Me.txt_p43.Value = 1
    Else
        Me.txt_p43.Value = Me.txt_p43.Value + 1
    End If
End Sub
Private Sub btn_p44_Click()
    If Me.txt_p44 = Empty Then
        Me.txt_p44.Value = 1
    Else
        Me.txt_p44.Value = Me.txt_p44.Value + 1
    End If
End Sub
Private Sub btn_p45_Click()
    If Me.txt_p45 = Empty Then
        Me.txt_p45.Value = 1
    Else
        Me.txt_p45.Value = Me.txt_p45.Value + 1
    End If
End Sub
Private Sub btn_p46_Click()
    If Me.txt_p46 = Empty Then
        Me.txt_p46.Value = 1
    Else
        Me.txt_p46.Value = Me.txt_p46.Value + 1
    End If
End Sub
Private Sub btn_p47_Click()
    If Me.txt_p47 = Empty Then
        Me.txt_p47.Value = 1
    Else
        Me.txt_p47.Value = Me.txt_p47.Value + 1
    End If
End Sub
Private Sub btn_p48_Click()
    If Me.txt_p48 = Empty Then
        Me.txt_p48.Value = 1
    Else
        Me.txt_p48.Value = Me.txt_p48.Value + 1
    End If
End Sub
Private Sub btn_p49_Click()
    If Me.txt_p49 = Empty Then
        Me.txt_p49.Value = 1
    Else
        Me.txt_p49.Value = Me.txt_p49.Value + 1
    End If
End Sub
Private Sub btn_p50_Click()
    If Me.txt_p50 = Empty Then
        Me.txt_p50.Value = 1
    Else
        Me.txt_p50.Value = Me.txt_p50.Value + 1
    End If
End Sub
Private Sub btn_p51_Click()
    If Me.txt_p51 = Empty Then
        Me.txt_p51.Value = 1
    Else
        Me.txt_p51.Value = Me.txt_p51.Value + 1
    End If
End Sub
Private Sub btn_p52_Click()
    If Me.txt_p52 = Empty Then
        Me.txt_p52.Value = 1
    Else
        Me.txt_p52.Value = Me.txt_p52.Value + 1
    End If
End Sub
Private Sub btn_p53_Click()
    If Me.txt_p53 = Empty Then
        Me.txt_p53.Value = 1
    Else
        Me.txt_p53.Value = Me.txt_p53.Value + 1
    End If
End Sub
Private Sub btn_p54_Click()
    If Me.txt_p54 = Empty Then
        Me.txt_p54.Value = 1
    Else
        Me.txt_p54.Value = Me.txt_p54.Value + 1
    End If
End Sub
Private Sub btn_p55_Click()
    If Me.txt_p55 = Empty Then
        Me.txt_p55.Value = 1
    Else
        Me.txt_p55.Value = Me.txt_p55.Value + 1
    End If
End Sub
Private Sub btn_p56_Click()
    If Me.txt_p56 = Empty Then
        Me.txt_p56.Value = 1
    Else
        Me.txt_p56.Value = Me.txt_p56.Value + 1
    End If
End Sub
Private Sub btn_p57_Click()
    If Me.txt_p57 = Empty Then
        Me.txt_p57.Value = 1
    Else
        Me.txt_p57.Value = Me.txt_p57.Value + 1
    End If
End Sub
Private Sub btn_p58_Click()
    If Me.txt_p58 = Empty Then
        Me.txt_p58.Value = 1
    Else
        Me.txt_p58.Value = Me.txt_p58.Value + 1
    End If
End Sub
Private Sub btn_p59_Click()
    If Me.txt_p59 = Empty Then
        Me.txt_p59.Value = 1
    Else
        Me.txt_p59.Value = Me.txt_p59.Value + 1
    End If
End Sub
Private Sub btn_p60_Click()
    If Me.txt_p60 = Empty Then
        Me.txt_p60.Value = 1
    Else
        Me.txt_p60.Value = Me.txt_p60.Value + 1
    End If
End Sub
Private Sub btn_p61_Click()
    If Me.txt_p61 = Empty Then
        Me.txt_p61.Value = 1
    Else
        Me.txt_p61.Value = Me.txt_p61.Value + 1
    End If
End Sub
Private Sub btn_p62_Click()
    If Me.txt_p62 = Empty Then
        Me.txt_p62.Value = 1
    Else
        Me.txt_p62.Value = Me.txt_p62.Value + 1
    End If
End Sub
Private Sub btn_p63_Click()
    If Me.txt_p63 = Empty Then
        Me.txt_p63.Value = 1
    Else
        Me.txt_p63.Value = Me.txt_p63.Value + 1
    End If
End Sub
Private Sub btn_p64_Click()
    If Me.txt_p64 = Empty Then
        Me.txt_p64.Value = 1
    Else
        Me.txt_p64.Value = Me.txt_p64.Value + 1
    End If
End Sub
Private Sub btn_p65_Click()
    If Me.txt_p65 = Empty Then
        Me.txt_p65.Value = 1
    Else
        Me.txt_p65.Value = Me.txt_p65.Value + 1
    End If
End Sub
Private Sub btn_p66_Click()
    If Me.txt_p66 = Empty Then
        Me.txt_p66.Value = 1
    Else
        Me.txt_p66.Value = Me.txt_p66.Value + 1
    End If
End Sub
Private Sub btn_p67_Click()
    If Me.txt_p67 = Empty Then
        Me.txt_p67.Value = 1
    Else
        Me.txt_p67.Value = Me.txt_p67.Value + 1
    End If
End Sub
Private Sub btn_p68_Click()
    If Me.txt_p68 = Empty Then
        Me.txt_p68.Value = 1
    Else
        Me.txt_p68.Value = Me.txt_p68.Value + 1
    End If
End Sub
Private Sub btn_p69_Click()
    If Me.txt_p69 = Empty Then
        Me.txt_p69.Value = 1
    Else
        Me.txt_p69.Value = Me.txt_p69.Value + 1
    End If
End Sub
Private Sub btn_p70_Click()
    If Me.txt_p70 = Empty Then
        Me.txt_p70.Value = 1
    Else
        Me.txt_p70.Value = Me.txt_p70.Value + 1
    End If
End Sub
Private Sub btn_p71_Click()
    If Me.txt_p71 = Empty Then
        Me.txt_p71.Value = 1
    Else
        Me.txt_p71.Value = Me.txt_p71.Value + 1
    End If
End Sub
Private Sub btn_p72_Click()
    If Me.txt_p72 = Empty Then
        Me.txt_p72.Value = 1
    Else
        Me.txt_p72.Value = Me.txt_p72.Value + 1
    End If
End Sub
Private Sub btn_p73_Click()
    If Me.txt_p73 = Empty Then
        Me.txt_p73.Value = 1
    Else
        Me.txt_p73.Value = Me.txt_p73.Value + 1
    End If
End Sub
Private Sub btn_p74_Click()
    If Me.txt_p74 = Empty Then
        Me.txt_p74.Value = 1
    Else
        Me.txt_p74.Value = Me.txt_p74.Value + 1
    End If
End Sub
Private Sub btn_p75_Click()
    If Me.txt_p75 = Empty Then
        Me.txt_p75.Value = 1
    Else
        Me.txt_p75.Value = Me.txt_p75.Value + 1
    End If
End Sub
Private Sub btn_p76_Click()
    If Me.txt_p76 = Empty Then
        Me.txt_p76.Value = 1
    Else
        Me.txt_p76.Value = Me.txt_p76.Value + 1
    End If
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub btn_p77_Click()
    If Me.txt_p77 = Empty Then
        Me.txt_p77.Value = 1
    Else
        Me.txt_p77.Value = Me.txt_p77.Value + 1
    End If
End Sub
Private Sub btn_p78_Click()
    If Me.txt_p78 = Empty Then
        Me.txt_p78.Value = 1
    Else
        Me.txt_p78.Value = Me.txt_p78.Value + 1
    End If
End Sub
Private Sub btn_p79_Click()
    If Me.txt_p79 = Empty Then
        Me.txt_p79.Value = 1
    Else
        Me.txt_p79.Value = Me.txt_p79.Value + 1
    End If
End Sub
Private Sub btn_p80_Click()
    If Me.txt_p80 = Empty Then
        Me.txt_p80.Value = 1
    Else
        Me.txt_p80.Value = Me.txt_p80.Value + 1
    End If
End Sub
Private Sub btn_p81_Click()
    If Me.txt_p81 = Empty Then
        Me.txt_p81.Value = 1
    Else
        Me.txt_p81.Value = Me.txt_p81.Value + 1
    End If
End Sub

Private Sub btn_registro_Click()
On Error GoTo Salir

Application.ScreenUpdating = False

    If Me.lbx_pieza.ListCount = 0 Then

            MsgBox "No hay registros en el cuadro de lista", , "Gestor de Inventario de Herramientas"
            Exit Sub

    End If
    
If MsgBox("Son correctos los datos?" + Chr(13) + "Desea procesar el registro?", vbYesNo, "Gestor de Inventarios") = vbNo Then
        Exit Sub
    Else

        Registrar_Pieza
        MsgBox "Datos registrados con éxito!!!", , "Gestor de Inventario de Herramientas"
        Unload Me
                ThisWorkbook.Save

End If

    Hoja0.Activate
    Hoja0.Select
     Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventario de Herramientas"
 End If
 
End Sub

Private Sub Registrar_Pieza()
Dim xFecha As Date
Dim xCodigo As String
Dim xPieza As String
Dim xCaja As String
Dim xItem As String
Dim xCantidad As Integer
Dim xEstado As String
Dim Indice As Long


        Indice = Hoja5.Range("U2").Value
        'Envía los datos a la hoja de ENTRADAS
            Hoja11.Activate
            Hoja11.Select
                    
            For i = 0 To Me.lbx_pieza.ListCount - 1
            
                    Hoja11.Rows("2:2").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                    
                    xFecha = Me.lbx_pieza.List(i, 0)
                    xCodigo = Me.lbx_pieza.List(i, 1)
                    xPieza = Me.lbx_pieza.List(i, 2)
                    xCantidad = Me.lbx_pieza.List(i, 3)
                    xCaja = Me.lbx_pieza.List(i, 4)
                    xItem = Me.lbx_pieza.List(i, 5)
                    xEstado = Me.lbx_pieza.List(i, 6)
                
                    Hoja11.Cells(2, 1) = Indice + i
                    Hoja11.Cells(2, 2) = CDate(xFecha)
                    Hoja11.Cells(2, 3) = xCodigo
                    Hoja11.Cells(2, 4) = xPieza
                    Hoja11.Cells(2, 5) = xCantidad
                    Hoja11.Cells(2, 6) = xCaja
                    Hoja11.Cells(2, 7) = xItem
                    Hoja11.Cells(2, 8) = "Activo"
                    Hoja11.Cells(2, 9) = xEstado
                    
                    Hoja5.Range("U2") = Indice + i
                    
                    
            Next
     
End Sub





Private Sub cbx_estado_Change()
If Me.cbx_estado.Text = "Dañado" Then
    Me.cbx_estado.BackColor = &H8080FF
ElseIf Me.cbx_estado.Text = "Faltante" Then
    Me.cbx_estado.BackColor = &H80FFFF
Else
    Me.cbx_estado.BackColor = &H80FF80
End If

Me.txt_f1.SetFocus

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub img_l77_Click()
    Me.txt_p77.Value = Empty
End Sub
Private Sub img_l78_Click()
    Me.txt_p78.Value = Empty
End Sub
Private Sub img_l79_Click()
    Me.txt_p79.Value = Empty
End Sub
Private Sub img_l80_Click()
    Me.txt_p80.Value = Empty
End Sub
Private Sub img_l81_Click()
    Me.txt_p81.Value = Empty
End Sub

Private Sub btn_borrar_Click()

On Error GoTo Errores

Me.lbx_pieza.RemoveItem (lbx_pieza.ListIndex)
Me.lbx_pieza.ListIndex = -1 ' Eliminar la "barra de selección"

Exit Sub

Errores:
MsgBox "Debe seleccionar una pieza del listado"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub img_fllave_Click()
    Me.txt_fllave.Value = Empty
End Sub

Private Sub img_l1_Click()
    Me.txt_p1.Value = Empty
End Sub
Private Sub img_l2_Click()
    Me.txt_p2.Value = Empty
End Sub
Private Sub img_l3_Click()
    Me.txt_p3.Value = Empty
End Sub
Private Sub img_l4_Click()
    Me.txt_p4.Value = Empty
End Sub
Private Sub img_l5_Click()
    Me.txt_p5.Value = Empty
End Sub
Private Sub img_l6_Click()
    Me.txt_p6.Value = Empty
End Sub
Private Sub img_l7_Click()
    Me.txt_p7.Value = Empty
End Sub
Private Sub img_l8_Click()
    Me.txt_p8.Value = Empty
End Sub
Private Sub img_l9_Click()
    Me.txt_p9.Value = Empty
End Sub
Private Sub img_l10_Click()
    Me.txt_p10.Value = Empty
End Sub
Private Sub img_l11_Click()
    Me.txt_p11.Value = Empty
End Sub
Private Sub img_l12_Click()
    Me.txt_p12.Value = Empty
End Sub
Private Sub img_l13_Click()
    Me.txt_p13.Value = Empty
End Sub
Private Sub img_l14_Click()
    Me.txt_p14.Value = Empty
End Sub
Private Sub img_l15_Click()
    Me.txt_p15.Value = Empty
End Sub
Private Sub img_l16_Click()
    Me.txt_p16.Value = Empty
End Sub
Private Sub img_l17_Click()
    Me.txt_p17.Value = Empty
End Sub
Private Sub img_l18_Click()
    Me.txt_p18.Value = Empty
End Sub
Private Sub img_l19_Click()
    Me.txt_p19.Value = Empty
End Sub
Private Sub img_l20_Click()
    Me.txt_p20.Value = Empty
End Sub
Private Sub img_l21_Click()
    Me.txt_p21.Value = Empty
End Sub
Private Sub img_l22_Click()
    Me.txt_p22.Value = Empty
End Sub
Private Sub img_l23_Click()
    Me.txt_p23.Value = Empty
End Sub
Private Sub img_l24_Click()
    Me.txt_p24.Value = Empty
End Sub
Private Sub img_l25_Click()
    Me.txt_p25.Value = Empty
End Sub
Private Sub img_l26_Click()
    Me.txt_p26.Value = Empty
End Sub
Private Sub img_l27_Click()
    Me.txt_p27.Value = Empty
End Sub
Private Sub img_l28_Click()
    Me.txt_p28.Value = Empty
End Sub
Private Sub img_l29_Click()
    Me.txt_p29.Value = Empty
End Sub
Private Sub img_l30_Click()
    Me.txt_p30.Value = Empty
End Sub
Private Sub img_l31_Click()
    Me.txt_p31.Value = Empty
End Sub
Private Sub img_l32_Click()
    Me.txt_p32.Value = Empty
End Sub
Private Sub img_l33_Click()
    Me.txt_p33.Value = Empty
End Sub
Private Sub img_l34_Click()
    Me.txt_p34.Value = Empty
End Sub
Private Sub img_l35_Click()
    Me.txt_p35.Value = Empty
End Sub
Private Sub img_l36_Click()
    Me.txt_p36.Value = Empty
End Sub
Private Sub img_l37_Click()
    Me.txt_p37.Value = Empty
End Sub
Private Sub img_l38_Click()
    Me.txt_p38.Value = Empty
End Sub
Private Sub img_l39_Click()
    Me.txt_p39.Value = Empty
End Sub
Private Sub img_l40_Click()
    Me.txt_p40.Value = Empty
End Sub
Private Sub img_l41_Click()
    Me.txt_p41.Value = Empty
End Sub
Private Sub img_l42_Click()
    Me.txt_p42.Value = Empty
End Sub
Private Sub img_l43_Click()
    Me.txt_p43.Value = Empty
End Sub
Private Sub img_l44_Click()
    Me.txt_p44.Value = Empty
End Sub
Private Sub img_l45_Click()
    Me.txt_p45.Value = Empty
End Sub
Private Sub img_l46_Click()
    Me.txt_p46.Value = Empty
End Sub
Private Sub img_l47_Click()
    Me.txt_p47.Value = Empty
End Sub
Private Sub img_l48_Click()
    Me.txt_p48.Value = Empty
End Sub
Private Sub img_l49_Click()
    Me.txt_p49.Value = Empty
End Sub
Private Sub img_l50_Click()
    Me.txt_p50.Value = Empty
End Sub
Private Sub img_l51_Click()
    Me.txt_p51.Value = Empty
End Sub
Private Sub img_l52_Click()
    Me.txt_p52.Value = Empty
End Sub
Private Sub img_l53_Click()
    Me.txt_p53.Value = Empty
End Sub
Private Sub img_l54_Click()
    Me.txt_p54.Value = Empty
End Sub
Private Sub img_l55_Click()
    Me.txt_p55.Value = Empty
End Sub
Private Sub img_l56_Click()
    Me.txt_p56.Value = Empty
End Sub
Private Sub img_l57_Click()
    Me.txt_p57.Value = Empty
End Sub
Private Sub img_l58_Click()
    Me.txt_p58.Value = Empty
End Sub
Private Sub img_l59_Click()
    Me.txt_p59.Value = Empty
End Sub
Private Sub img_l60_Click()
    Me.txt_p60.Value = Empty
End Sub
Private Sub img_l61_Click()
    Me.txt_p61.Value = Empty
End Sub
Private Sub img_l62_Click()
    Me.txt_p62.Value = Empty
End Sub
Private Sub img_l63_Click()
    Me.txt_p63.Value = Empty
End Sub
Private Sub img_l64_Click()
    Me.txt_p64.Value = Empty
End Sub
Private Sub img_l65_Click()
    Me.txt_p65.Value = Empty
End Sub
Private Sub img_l66_Click()
    Me.txt_p66.Value = Empty
End Sub
Private Sub img_l67_Click()
    Me.txt_p67.Value = Empty
End Sub
Private Sub img_l68_Click()
    Me.txt_p68.Value = Empty
End Sub
Private Sub img_l69_Click()
    Me.txt_p69.Value = Empty
End Sub
Private Sub img_l70_Click()
    Me.txt_p70.Value = Empty
End Sub
Private Sub img_l71_Click()
    Me.txt_p71.Value = Empty
End Sub
Private Sub img_l72_Click()
    Me.txt_p72.Value = Empty
End Sub
Private Sub img_l73_Click()
    Me.txt_p73.Value = Empty
End Sub
Private Sub img_l74_Click()
    Me.txt_p74.Value = Empty
End Sub
Private Sub img_l75_Click()
    Me.txt_p75.Value = Empty
End Sub
Private Sub img_l76_Click()
    Me.txt_p76.Value = Empty
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub img_l1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l1.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l2.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l3.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l4.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l5.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l6_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l6.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l7.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l8_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l8.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l9_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l9.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l10_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l10.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l11_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l11.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l12_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l12.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l13_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l13.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l14_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l14.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l15_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l15.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l16_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l16.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l17_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l17.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l18_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l18.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l19_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l19.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l20_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l20.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l21_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l21.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l22_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l22.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l23_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l23.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l24_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l24.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l25_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l25.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l26_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l26.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l27_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l27.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l28_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l28.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l29_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l29.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l30_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l30.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l31_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l31.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l32_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l32.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l33_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l33.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l34_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l34.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l35_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l35.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l36_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l36.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l37_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l37.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l38_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l38.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l39_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l39.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l40_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l40.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l41_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l41.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l42_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l42.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l43_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l43.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l44_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l44.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l45_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l45.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l46_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l46.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l47_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l47.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l48_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l48.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l49_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l49.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l50_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l50.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l51_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l51.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l52_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l52.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l53_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l53.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l54_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l54.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l55_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l55.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l56_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l56.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l57_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l57.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l58_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l58.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l59_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l59.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l60_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l60.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l61_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l61.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l62_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l62.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l63_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l63.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l64_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l64.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l65_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l65.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l66_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l66.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l67_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l67.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l68_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l68.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l69_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l69.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l70_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l70.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l71_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l71.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l72_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l72.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l73_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l73.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l74_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l74.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l75_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l75.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l76_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l76.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l1.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l2.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l3.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l4.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l5_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l5.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l6.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l7_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l7.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l8.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l9_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l9.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l10_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l10.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l11_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l11.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l12_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l12.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l13_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l13.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l14_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l14.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l15_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l15.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l16_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l16.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l17_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l17.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l18_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l18.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l19_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l19.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l20_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l20.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l21_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l21.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l22_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l22.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l23_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l23.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l24_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l24.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l25_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l25.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l26_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l26.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l27_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l27.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l28_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l28.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l29_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l29.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l30_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l30.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l31_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l31.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l32_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l32.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l33_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l33.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l34_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l34.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l35_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l35.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l36_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l36.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l37_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l37.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l38_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l38.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l39_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l39.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l40_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l40.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l41_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l41.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l42_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l42.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l43_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l43.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l44_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l44.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l45_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l45.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l46_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l46.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l47_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l47.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l48_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l48.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l49_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l49.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l50_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l50.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l51_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l51.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l52_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l52.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l53_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l53.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l54_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l54.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l55_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l55.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l56_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l56.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l57_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l57.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l58_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l58.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l59_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l59.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l60_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l60.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l61_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l61.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l62_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l62.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l63_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l63.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l64_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l64.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l65_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l65.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l66_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l66.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l67_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l67.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l68_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l68.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l69_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l69.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l70_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l70.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l71_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l71.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l72_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l72.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l73_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l73.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l74_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l74.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l75_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l75.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub img_l76_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l76.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l77_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l77.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l77_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l77.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_l78_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l78.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l78_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l78.SpecialEffect = fmSpecialEffectSunken
End Sub
Private Sub img_l79_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l79.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l79_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l79.SpecialEffect = fmSpecialEffectSunken
End Sub
Private Sub img_l80_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l80.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l80_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l80.SpecialEffect = fmSpecialEffectSunken
End Sub
Private Sub img_l81_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l81.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_l81_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_l81.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub img_fllave_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_fllave.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub img_fllave_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.img_fllave.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub lbx_pieza_Click()

End Sub

Private Sub txt_f0_Change()

End Sub

Private Sub txt_pieza_Change()

End Sub

Private Sub UserForm_Click()

End Sub
