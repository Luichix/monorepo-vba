Attribute VB_Name = "Ribbon"
Option Explicit
Option Base 1
Public CintaDeRibbon As IRibbonUI
Public RetVal(54) As Boolean

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If


Sub CargarCinta(CintaDeExcel As IRibbonUI)
    Set CintaDeRibbon = CintaDeExcel
End Sub

'////////////////////// Llamadas desde la Cinta para ejectuar cada formulario ///////////////////////////////////


Sub Boton1(Control As IRibbonControl)
   Hoja6.Select
   Hoja6.Cells(1, 1).Select
End Sub

Sub Boton2(Control As IRibbonControl)
   Hoja3.Select
   Hoja3.Cells(1, 1).Select
   frm_Factura.Show
End Sub

Sub Boton3(Control As IRibbonControl)
    Hoja4.Select
    Hoja4.Cells(1, 1).Select
   frm_Nota_Credito.Show
End Sub

Sub Boton4(Control As IRibbonControl)
    

    frm_Cierre.Show
     
End Sub

Sub Boton5(Control As IRibbonControl)
    ThisWorkbook.Save
End Sub


'//////////////////// Retornos del estado de cada botón ////////////////////////


Public Sub DesactivarBoton1(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = True
    
End Sub


Public Sub DesactivarBoton2(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = True
    
End Sub

Public Sub DesactivarBoton3(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = True
    
End Sub

Public Sub DesactivarBoton4(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = True
    
End Sub


Public Sub DesactivarBoton5(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = True
    
End Sub
