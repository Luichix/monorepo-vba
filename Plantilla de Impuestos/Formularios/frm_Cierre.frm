VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Cierre 
   Caption         =   "REGISTRO"
   ClientHeight    =   4770
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   7410
   OleObjectBlob   =   "frm_Cierre.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Cierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub btn_Fecha_Click()
Me.txt_Fecha.BackColor = &H80000005
banderaCalendario = 3
  Call LanzarCalendario(Me, "txt_Fecha")
  
End Sub

Private Sub cmd_materiales_Click()
    Dim Titulo As String
Application.ScreenUpdating = False

Titulo = "CIERRE MENSUAL"
   
If Me.txt_Fecha.Text = "" Then
    Me.txt_Fecha.BackColor = &HC0C0FF
    MsgBox "Ingrese la fecha de Cierre..!", vbInformation, Titulo
    Me.btn_Fecha.SetFocus
    Exit Sub
End If
    
    If MsgBox("¿Esta seguro que desea realizar el cierre de Caja Mensual?" + Chr(13) + "¡Si lo hace solo se hara revisón de los movimientos del dia!", vbYesNo, "Gestor de Caja") = vbNo Then
        Exit Sub
    Else
        
    CierreZ
    MsgBox "Cierre Diario realizado con exito..!", vbInformation, Titulo
    
        Unload Me
        
Application.ScreenUpdating = True

    End If


End Sub
Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Frame1.SpecialEffect = fmSpecialEffectSunken
      Frame3.SpecialEffect = fmSpecialEffectFlat
       Frame5.SpecialEffect = fmSpecialEffectFlat


End Sub

Private Sub cmd_salir_Click()
Unload Me
End Sub

Private Sub Frame3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Frame3.SpecialEffect = fmSpecialEffectSunken
     Frame1.SpecialEffect = fmSpecialEffectFlat
      Frame5.SpecialEffect = fmSpecialEffectFlat


End Sub

Private Sub Frame5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Frame5.SpecialEffect = fmSpecialEffectSunken
     Frame1.SpecialEffect = fmSpecialEffectFlat
      Frame3SpecialEffect = fmSpecialEffectFlat


End Sub
Private Sub CommandButton1_Click()
Unload Me
End Sub


Private Sub txt_Cierre_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_cierre, KeyAscii)
End Sub


Private Sub Image1_Click()
    Dim Titulo As String
Application.ScreenUpdating = False

Titulo = "CIERRE DIARIO"
   
If Me.txt_Fecha.Text = "" Then
    Me.txt_Fecha.BackColor = &HC0C0FF
    MsgBox "Ingrese la fecha de Cierre..!", vbInformation, Titulo
    Me.btn_Fecha.SetFocus
    Exit Sub
End If
    
    If MsgBox("¿Esta seguro que desea realizar el cierre de Caja Diario?" + Chr(13) + "¡Si lo hace solo se hara revisón de los movimientos del dia!", vbYesNo, "Gestor de Caja") = vbNo Then
        Exit Sub
    Else
        
    CierreX
    MsgBox "Cierre Diario realizado con exito..!", vbInformation, Titulo
    
        Unload Me
        
Application.ScreenUpdating = True

    End If

End Sub

Private Sub UserForm_Initialize()
EliminarTitulo Me.Caption
    Me.Height = Me.Height - 20


End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
            Frame3.SpecialEffect = fmSpecialEffectFlat
     Frame1.SpecialEffect = fmSpecialEffectFlat
     Frame5.SpecialEffect = fmSpecialEffectFlat
End Sub


