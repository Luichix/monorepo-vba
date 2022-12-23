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


Private Sub Workbook_BeforePrint(Cancel As Boolean)
Cancel = True
End Sub

Private Sub Workbook_Open()
   
    Call OcultarHojas
    'Call MostrarHojas
    SplashForm.Show
       
   ThisWorkbook.Save
   
End Sub
'
Private Sub Workbook_NewSheet(ByVal Sh As Object)
Application.DisplayAlerts = False
ActiveSheet.Delete

MsgBox "No esta Permitido insertar Nuevas Hojas de Cálculo...!", vbInformation, "GESTOR DE VENTA"

Application.DisplayAlerts = True
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Select Case SaveAsUI
Case True: 'Se ha pulsado Guardar Como...
Cancel = SaveAsUI
MsgBox "No tiene permisos para Guardar Copia del Archivo...!", vbInformation, "GESTOR DE VENTAS"

'Case False: 'Se ha pulsado Guardar...
'Cancel = False
'MsgBox "No tienes permisos para Guardar el Archivo...!", vbInformation, "GESTOR DE VENTAS"
'
End Select
End Sub

