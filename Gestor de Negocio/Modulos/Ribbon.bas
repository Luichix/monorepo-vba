Attribute VB_Name = "Ribbon"
Option Explicit
Option Base 1
Public CintaDeRibbon As IRibbonUI
Public RetVal(42) As Boolean

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If


Sub CargarCinta(CintaDeExcel As IRibbonUI)
    Set CintaDeRibbon = CintaDeExcel
    form_iniciosesion.Show
End Sub

'////////////////////// Llamadas desde la Cinta para ejectuar cada formulario ///////////////////////////////////


Sub Boton1(Control As IRibbonControl)
   Hoja20.Select
End Sub

Sub Boton2(Control As IRibbonControl)
   Hoja10.Select
   Hoja10.Cells(1, 1).Select
   frm_fCompras.Show
End Sub

Sub Boton3(Control As IRibbonControl)
    Hoja11.Select
   frm_Transferencias.Show
End Sub

Sub Boton4(Control As IRibbonControl)
    Hoja12.Select
    frm_ConsultaProducto.Show
End Sub

Sub Boton5(Control As IRibbonControl)
    Hoja23.Select
    frm_RegistrarProveedor.Show
End Sub

Sub Boton6(Control As IRibbonControl)
    Hoja4.Select
    frm_RegistrarClientes.Show
End Sub

Sub Boton7(Control As IRibbonControl)
    Hoja26.Select
    frm_Factura.Show
End Sub

Sub Boton8(Control As IRibbonControl)
    Hoja5.Select
    Hoja5.Cells(1, 1).Select
    frm_Personal.Show
End Sub

Sub Boton9(Control As IRibbonControl)
    Hoja6.Select
    Hoja6.Cells(1, 1).Select
    form_registropagos.Show
End Sub

Sub Boton10(Control As IRibbonControl)
    Hoja16.Select
    Hoja16.Cells(1, 1).Select
    frm_reportes.Show
End Sub

Sub Boton11(Control As IRibbonControl)
    Hoja17.Select
    Hoja17.Cells(1, 1).Select
    frm_ausencias.Show
End Sub

Sub Boton12(Control As IRibbonControl)
    Hoja33.Select
    Hoja33.Cells(1, 1).Select
    frm_horas.Show
    
End Sub

Sub Boton13(Control As IRibbonControl)
    Hoja7.Select
    frm_egresosganaderia.Show
End Sub
Sub Boton14(Control As IRibbonControl)
    Hoja8.Select
    frm_egresosmadera.Show
End Sub
Sub Boton15(Control As IRibbonControl)
    Hoja13.Select
    frm_egresosgastos.Show
End Sub
Sub Boton16(Control As IRibbonControl)
    Hoja14.Select
    frm_egresoslegales.Show
End Sub
Sub Boton17(Control As IRibbonControl)
    Hoja15.Select
    frm_egresosfamiliares.Show
End Sub
Sub Boton18(Control As IRibbonControl)
ThisWorkbook.Save
End Sub
Sub Boton19(Control As IRibbonControl)
   Hoja20.Select
End Sub
Sub Boton20(Control As IRibbonControl)
    Load frm_Ganado
    Hoja29.Select
    frm_Ganado.Show
End Sub
Sub Boton21(Control As IRibbonControl)
    Hoja30.Select
End Sub
Sub Boton22(Control As IRibbonControl)
    Hoja32.Select
    frm_enfermeria.Show
End Sub
Sub Boton23(Control As IRibbonControl)
    Hoja31.Select
End Sub
Sub Boton24(Control As IRibbonControl)
ThisWorkbook.Save
End Sub
Sub Boton25(Control As IRibbonControl)
    Hoja20.Select
    form_iniciosesion.Show
End Sub
Sub Boton26(Control As IRibbonControl)
    Hoja20.Select
    form_iniciosesion.Show
End Sub
Sub Boton27(Control As IRibbonControl)
   Hoja20.Select
End Sub
Sub Boton28(Control As IRibbonControl)
    Hoja41.Select
    Load frm_CatalogoCuentas
    frm_CatalogoCuentas.Show
End Sub
Sub Boton29(Control As IRibbonControl)
    Hoja42.Select
    Load frm_LibroDiario
    frm_LibroDiario.Show
End Sub
Sub Boton30(Control As IRibbonControl)
    Hoja43.Select
End Sub
Sub Boton31(Control As IRibbonControl)
   Hoja44.Select
End Sub
Sub Boton32(Control As IRibbonControl)
    Hoja46.Select
End Sub
Sub Boton33(Control As IRibbonControl)
   Hoja45.Select
End Sub
Sub Boton34(Control As IRibbonControl)
    form_iniciosesion.Show
End Sub
Sub Boton35(Control As IRibbonControl)
   ThisWorkbook.Save
End Sub
Sub Boton36(Control As IRibbonControl)
   REPORTE1
End Sub
Sub Boton37(Control As IRibbonControl)
 
End Sub
Sub Boton38(Control As IRibbonControl)
 
End Sub
Sub Boton39(Control As IRibbonControl)
 
End Sub
Sub Boton40(Control As IRibbonControl)
 
End Sub
Sub Boton41(Control As IRibbonControl)
 
End Sub
Sub Boton42(Control As IRibbonControl)
 
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
Public Sub DesactivarBoton13(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(13)
    
End Sub
Public Sub DesactivarBoton14(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(14)
    
End Sub
Public Sub DesactivarBoton15(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(15)
    
End Sub
Public Sub DesactivarBoton16(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(16)
    
End Sub
Public Sub DesactivarBoton17(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(17)
    
End Sub
Public Sub DesactivarBoton18(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(18)
       
End Sub
Public Sub DesactivarBoton19(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(19)
    
End Sub
Public Sub DesactivarBoton20(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(20)
    
End Sub
Public Sub DesactivarBoton21(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(21)
    
End Sub
Public Sub DesactivarBoton22(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(22)
    
End Sub
Public Sub DesactivarBoton23(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(23)
    
End Sub
Public Sub DesactivarBoton24(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(24)
    
End Sub
Public Sub DesactivarBoton25(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(25)
    
End Sub
Public Sub DesactivarBoton26(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(26)
    
End Sub
Public Sub DesactivarBoton27(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(27)
    
End Sub
Public Sub DesactivarBoton28(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(28)
    
End Sub
Public Sub DesactivarBoton29(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(29)
    
End Sub
Public Sub DesactivarBoton30(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(30)
    
End Sub
Public Sub DesactivarBoton31(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(31)
    
End Sub
Public Sub DesactivarBoton32(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(32)
    
End Sub
Public Sub DesactivarBoton33(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(33)
    
End Sub
Public Sub DesactivarBoton34(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(34)
    
End Sub
Public Sub DesactivarBoton35(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(35)
    
End Sub

Public Sub DesactivarBoton36(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(36)
    
End Sub

Public Sub DesactivarBoton37(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(37)
    
End Sub
Public Sub DesactivarBoton38(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(38)
    
End Sub
Public Sub DesactivarBoton39(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(39)
    
End Sub
Public Sub DesactivarBoton40(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(40)
    
End Sub
Public Sub DesactivarBoton41(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(41)
    
End Sub
Public Sub DesactivarBoton42(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(42)
    
End Sub
