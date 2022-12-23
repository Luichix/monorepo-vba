VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_General_Select 
   Caption         =   "GESTOR DE VENTAS"
   ClientHeight    =   7050
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   7570
   OleObjectBlob   =   "frm_General_Select.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_General_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_PDF_Click()
Unload Me

    Mensaje = MsgBox("Esta seguro que desea generar el archivo PDF, no creara copia e incluira el monto por hora" + Chr(13) + "¿Desea Continuar?", _
    vbYesNo + vbQuestion, "Historico")
    
        MsgBox "Espere un momento... Click para continuar..."
Application.Cursor = xlWait
With frm_General
    .General_Quincena
    .Cargar_PDF
    .GENERAL_GUARDAR_PDF
End With
Application.Cursor = xlDefault
        MsgBox "Comprobante de pago elaborado con éxito!!!", vbInformation, Titulo

End Sub

Private Sub cmd_reporte_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub cmd_reporte_Click()
Unload Me
        MsgBox "Espere un momento... Click para continuar..."
      
 Application.Cursor = xlWait
               
With frm_General
    .General_Quincena
    .Cargar_Todo
    
End With

 Application.Cursor = xlDefault
 
        MsgBox "Comprobante de pago elaborado con éxito!!!", vbInformation, Titulo
        
       
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Frame1.SpecialEffect = fmSpecialEffectSunken
     Frame2.SpecialEffect = fmSpecialEffectFlat
      Frame3.SpecialEffect = fmSpecialEffectFlat
              Frame5.SpecialEffect = fmSpecialEffectFlat

End Sub
Private Sub cmd_Correo_Click()
Unload Me
        MsgBox "Espere un momento... Click para continuar..."

Application.Cursor = xlWait
 With frm_General
   .General_Quincena
    .Cargar_PDF
    .GENERAL_ENVIAR_PDF
End With

Application.Cursor = xlDefault

        MsgBox "Comprobante de pago elaborado con éxito!!!", vbInformation, Titulo

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

Lbl3.Visible = True
 Lbl2.Visible = False
    Lbl1.Visible = False
    Frame3.SpecialEffect = fmSpecialEffectSunken
    Frame2.SpecialEffect = fmSpecialEffectFlat
     Frame1.SpecialEffect = fmSpecialEffectFlat
        Frame5.SpecialEffect = fmSpecialEffectFlat

End Sub
Private Sub CommandButton1_Click()
Unload Me
End Sub
Private Sub Frame5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)


   Frame5.SpecialEffect = fmSpecialEffectSunken
    Frame3.SpecialEffect = fmSpecialEffectFlat
     Frame1.SpecialEffect = fmSpecialEffectFlat
        Frame2.SpecialEffect = fmSpecialEffectFlat

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


