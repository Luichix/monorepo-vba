VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Application.ScreenUpdating = False
Call OcultarHojas
Application.ScreenUpdating = True


End Sub

Private Sub Workbook_Open()
   
    Call OcultarHojas
    'Call MostrarHojas
    SplashForm.Show
       
   ThisWorkbook.Save
   
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
Application.DisplayAlerts = False
ActiveSheet.Delete

MsgBox "No esta Permitido insertar Nuevas Hojas de Cálculo...!", vbInformation, "GESTOR DE VENTA"

Application.DisplayAlerts = True
End Sub

