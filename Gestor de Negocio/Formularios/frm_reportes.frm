VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_reportes 
   Caption         =   "SISTEMA DE REPORTES"
   ClientHeight    =   4130
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   9090.001
   OleObjectBlob   =   "frm_reportes.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnotas_Click()
Hoja18.Select
Hoja18.Cells(1, 1).Select
Unload Me
End Sub

Private Sub cmdnotas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
cmdnotas.BorderColor = &H80000003 '&H80FF& '&H8000000D
lbl2.Visible = True
End Sub

Private Sub cmdprofesor_Click()
Reporte_Planilla
Unload Me
End Sub

Private Sub cmdprofesor_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
cmdprofesor.BorderColor = &H80000003 '&H80FF&    '&H8000000D
lbl1.Visible = True
End Sub
Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdsalir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
cmdsalir.BorderColor = &H80000003 '&H80FF& '&H8000000D
lbl3.Visible = True
End Sub

Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub UserForm_Initialize()
'    EliminarTitulo Me.Caption
'    Me.Height = Me.Height - 20
    
    lbl1.Visible = False
    lbl2.Visible = False
    lbl3.Visible = False
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    cmdprofesor.BorderColor = &HFFFFFF
    cmdnotas.BorderColor = &HFFFFFF
    cmdsalir.BorderColor = &HFFFFFF

    lbl1.Visible = False
    lbl2.Visible = False
    lbl3.Visible = False


End Sub
