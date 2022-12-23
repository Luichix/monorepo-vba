VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_CierreCaja 
   Caption         =   "GESTOR DE VENTAS"
   ClientHeight    =   4725
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   6670
   OleObjectBlob   =   "frm_CierreCaja.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_CierreCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_materiales_Click()
Dim Titulo As String

Titulo = "GESTOR DE CAJA"

    If Hoja26.Cells(2, 1) = "" Then
        MsgBox "No se ha registrado ninguna transacción.", vbInformation, Titulo
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro que desea realizar el cierre de Caja X?" + Chr(13) + "¡Si lo hace solo se hara revisón de los movimientos del dia!", vbYesNo, "Gestor de Caja") = vbNo Then
        Exit Sub
    Else

    Unload Me
    frm_ArqueoCaja.lbl_cierre.Caption = "CIERRE X"
    frm_ArqueoCaja.Show
    
    
    End If


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
Dim Titulo As String

Titulo = "GESTOR DE CAJA"

    If Hoja26.Cells(2, 1) = "" Then
        MsgBox "No se ha registrado ninguna transacción.", vbInformation, Titulo
        Exit Sub
    End If
        
    If MsgBox("¿Esta seguro que desea realizar el cierre de Caja Z?" + Chr(13) + "¡Si lo hace se limpiaran los movimientos del dia!", vbYesNo, "Gestor de Caja") = vbNo Then
        Exit Sub
    Else
    Unload Me
     If Hoja92.Range("H1") = "ADMINISTRADOR" Then
        frm_ArqueoCaja.lbl_cierre.Caption = "CIERRE Z"
        frm_ArqueoCaja.Show
        Else
         MsgBox "Debe ingresar desde una cuenta Administrativa para poder realizar el Cierre Z de Caja.", vbInformation, "GESTOR DE CAJA"
     End If
    
    
    
    End If
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


