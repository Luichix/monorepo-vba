VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Pedido_Select 
   Caption         =   "GESTOR DE VENTAS"
   ClientHeight    =   4980
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   6570
   OleObjectBlob   =   "frm_Pedido_Select.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Pedido_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_materiales_Click()
If frm_Factura.txt_nservicio <> Empty Then
    MsgBox "No puede grabarse un servicio como encargo, limpie la facturación y registre los datos nuevamente...!", vbInformation, "GESTOR DE PEDIDOS"
    Unload Me
    Exit Sub
End If
Unload Me
frm_Grabar.Show


End Sub
Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Lbl1.Visible = True
 Lbl2.Visible = False
    Lbl3.Visible = False
    
    Frame1.SpecialEffect = fmSpecialEffectSunken
     Frame2.SpecialEffect = fmSpecialEffectFlat
      Frame3.SpecialEffect = fmSpecialEffectFlat

End Sub
Private Sub cmd_productos_Click()
    If frm_Factura.txt_nPedido <> Empty Then
    MsgBox "No puede grabarse un encargo como servicio, limpie la facturación y registre los datos nuevamente...!", vbInformation, "GESTOR DE PEDIDOS"
    Unload Me
    Exit Sub
    End If
    Unload Me
    frm_Servicio.Show
End Sub
Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Lbl2.Visible = True
 Lbl1.Visible = False
    Lbl3.Visible = False
   Frame2.SpecialEffect = fmSpecialEffectSunken
    Frame3.SpecialEffect = fmSpecialEffectFlat
     Frame1.SpecialEffect = fmSpecialEffectFlat

End Sub
Private Sub cmd_salir_Click()
Unload Me
End Sub

Private Sub Frame3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Lbl3.Visible = True
 Lbl2.Visible = False
    Lbl1.Visible = False
    Frame3.SpecialEffect = fmSpecialEffectSunken
    Frame2.SpecialEffect = fmSpecialEffectFlat
     Frame1.SpecialEffect = fmSpecialEffectFlat

End Sub
Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub UserForm_Initialize()
EliminarTitulo Me.Caption
    Me.Height = Me.Height - 20
    Lbl1.Visible = True
    Lbl2.Visible = True
    Lbl3.Visible = True
        Frame3.SpecialEffect = fmSpecialEffectFlat
    Frame2.SpecialEffect = fmSpecialEffectFlat
     Frame1.SpecialEffect = fmSpecialEffectFlat
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Lbl1.Visible = True
    Lbl2.Visible = True
    Lbl3.Visible = True
            Frame3.SpecialEffect = fmSpecialEffectFlat
    Frame2.SpecialEffect = fmSpecialEffectFlat
     Frame1.SpecialEffect = fmSpecialEffectFlat
End Sub


