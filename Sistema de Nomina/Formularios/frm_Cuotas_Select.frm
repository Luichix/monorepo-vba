VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Cuotas_Select 
   Caption         =   "GESTOR DE VENTAS"
   ClientHeight    =   6180
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   6430
   OleObjectBlob   =   "frm_Cuotas_Select.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Cuotas_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmd_materiales_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub cmd_materiales_Click()
Unload Me
frm_Cuota_Nivelada.Show
End Sub
Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)


    
    Frame1.SpecialEffect = fmSpecialEffectSunken
     Frame2.SpecialEffect = fmSpecialEffectFlat
      Frame3.SpecialEffect = fmSpecialEffectFlat

End Sub
Private Sub cmd_productos_Click()
    Unload Me
    frm_Cuota_Variable.Show
End Sub
Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)


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


