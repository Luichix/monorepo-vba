Attribute VB_Name = "Ribbon"
Option Explicit
Option Base 1
Public CintaDeRibbon As IRibbonUI
Public RetVal(62) As Boolean

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If


Sub CargarCinta(CintaDeExcel As IRibbonUI)
    Set CintaDeRibbon = CintaDeExcel
    frm_Iniciosesion.Show
End Sub

'////////////////////// Llamadas desde la Cinta para ejectuar cada formulario ///////////////////////////////////


Sub Boton1(Control As IRibbonControl)
   Hoja0.Select
End Sub

Sub Boton2(Control As IRibbonControl)
Hoja2.Select
    frm_CatalogoCuentas.Show
End Sub

Sub Boton3(Control As IRibbonControl)
Hoja3.Select
    frm_LibroDiario.Show
End Sub

Sub Boton4(Control As IRibbonControl)
  EnviarAMayor
End Sub

Sub Boton5(Control As IRibbonControl)
ConstruirBalancedeComprobacion
End Sub

Sub Boton6(Control As IRibbonControl)
Estado_Resultado
End Sub

Sub Boton7(Control As IRibbonControl)
BalanceGeneral
End Sub

Sub Boton8(Control As IRibbonControl)
frm_NuevoUsuario.Show
End Sub

Sub Boton9(Control As IRibbonControl)
frm_EliminarUsuario.Show
End Sub

Sub Boton10(Control As IRibbonControl)
    frm_Iniciosesion.Show
End Sub

Sub Boton11(Control As IRibbonControl)
    ThisWorkbook.Save
End Sub



'//////////////////// Retornos del estado de cada botón ////////////////////////


Public Sub DesactivarBoton1(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(1)
    
End Sub


Public Sub DesactivarBoton2(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(2)
    
End Sub

Public Sub DesactivarBoton3(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(3)
    
End Sub

Public Sub DesactivarBoton4(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(4)
    
End Sub


Public Sub DesactivarBoton5(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(5)
    
End Sub


Public Sub DesactivarBoton6(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(6)
    
End Sub


Public Sub DesactivarBoton7(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(7)
    
End Sub

Public Sub DesactivarBoton8(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(8)
    
End Sub

Public Sub DesactivarBoton9(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(9)
    
End Sub
Public Sub DesactivarBoton10(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(10)
    
End Sub
Public Sub DesactivarBoton11(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(11)
    
End Sub
Public Sub DesactivarBoton12(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(12)
    
End Sub
