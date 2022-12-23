VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_hora_Select 
   Caption         =   "GESTOR DE VENTAS"
   ClientHeight    =   6555
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   7950
   OleObjectBlob   =   "frm_hora_Select.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_hora_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmd_materiales_Click()
Unload Me
Application.ScreenUpdating = False

Dim Seguridad As String
Seguridad = Hoja83.Range("L1").Text

Hoja58.Unprotect (Seguridad)
    frm_Hora_Marca.Show
Hoja58.Protect (Seguridad)

 Application.ScreenUpdating = True
End Sub
Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Frame1.SpecialEffect = fmSpecialEffectSunken
Frame2.SpecialEffect = fmSpecialEffectFlat
Frame3.SpecialEffect = fmSpecialEffectFlat
Frame5.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub cmd_productos_Click()
    Unload Me
    frm_hora_Multiple.Show
End Sub
Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Frame2.SpecialEffect = fmSpecialEffectSunken
Frame3.SpecialEffect = fmSpecialEffectFlat
Frame1.SpecialEffect = fmSpecialEffectFlat
Frame5.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub cmd_salir_Click()
Unload Me
End Sub

Private Sub Frame3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Frame3.SpecialEffect = fmSpecialEffectSunken
Frame2.SpecialEffect = fmSpecialEffectFlat
Frame5.SpecialEffect = fmSpecialEffectFlat
Frame1.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub Frame5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Frame5.SpecialEffect = fmSpecialEffectSunken
    Frame2.SpecialEffect = fmSpecialEffectFlat
    Frame1.SpecialEffect = fmSpecialEffectFlat
    Frame3.SpecialEffect = fmSpecialEffectFlat
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Image1_Click()
Unload Me
Importar_Data
End Sub

Private Sub UserForm_Initialize()
EliminarTitulo Me.Caption
    Me.Height = Me.Height - 20
    Frame3.SpecialEffect = fmSpecialEffectFlat
    Frame2.SpecialEffect = fmSpecialEffectFlat
    Frame1.SpecialEffect = fmSpecialEffectFlat
    Frame5.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Frame3.SpecialEffect = fmSpecialEffectFlat
    Frame2.SpecialEffect = fmSpecialEffectFlat
    Frame1.SpecialEffect = fmSpecialEffectFlat
    Frame5.SpecialEffect = fmSpecialEffectFlat
End Sub


