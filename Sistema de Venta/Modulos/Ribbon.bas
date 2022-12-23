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
    form_iniciosesion.Show
End Sub

'////////////////////// Llamadas desde la Cinta para ejectuar cada formulario ///////////////////////////////////


Sub Boton1(Control As IRibbonControl)
   Hoja0.Select
End Sub

Sub Boton2(Control As IRibbonControl)
  If Hoja3.Visible = xlSheetVisible Then
    Hoja3.Select
    Hoja3.Cells(1, 1).Select
  End If
   frm_fCompras.Show
End Sub

Sub Boton3(Control As IRibbonControl)
 If Hoja4.Visible = xlSheetVisible Then
    Hoja4.Select
    Hoja4.Cells(1, 1).Select
End If

   frm_Transferencias.Show
End Sub

Sub Boton4(Control As IRibbonControl)

    frm_Consulta_Materiales.Show
     
End Sub

Sub Boton5(Control As IRibbonControl)
    If Hoja8.Visible = xlSheetVisible Then
        Hoja8.Select
        Hoja8.Cells(1, 1).Select
    End If
    frm_RegistrarProveedor.Show
End Sub

Sub Boton6(Control As IRibbonControl)
    If Hoja7.Visible = xlSheetVisible Then
        Hoja7.Select
        Hoja7.Cells(1, 1).Select
    End If
    frm_RegistrarClientes.Show
End Sub

Sub Boton7(Control As IRibbonControl)
    If Hoja25.Visible = xlSheetVisible Then
    Hoja25.Select
    Hoja25.Cells(1, 1).Select
    End If
    If Hoja92.Range("H1") = "ADMINISTRADOR" Then
    frm_Devolucion.Show
    Else
    MsgBox "Debe ingresar desde una cuenta Administrativa para realizar devolución de efectivo.", vbInformation, "GESTOR DE DEVOLUCIÓNES"
    End If
End Sub

Sub Boton8(Control As IRibbonControl)
    Hoja5.Select
    Hoja5.Cells(1, 1).Select
    frm_Personal.Show
End Sub

Sub Boton9(Control As IRibbonControl)
'&H886815
    frm_ListaPedido.BackColor = &HB3891C
    frm_ListaPedido.Frame1.BackColor = &HB3891C
    frm_ListaPedido.Frame2.BackColor = &HB3891C
    frm_ListaPedido.btn_Facturar.Visible = False
    frm_ListaPedido.btn_Salir.Visible = False
    frm_ListaPedido.btn_Salir2.Visible = True
    frm_ListaPedido.Show
End Sub

Sub Boton10(Control As IRibbonControl)
    If Hoja92.Range("H1") = "ADMINISTRADOR" Then
    frm_ingreso.Show
     Else
    MsgBox "Debe ingresar desde una cuenta Administrativa para realizar ingresos de efectivo.", vbInformation, "GESTOR DE CAJA"
    End If
End Sub

Sub Boton11(Control As IRibbonControl)
    If Hoja92.Range("H1") = "ADMINISTRADOR" Then
    frm_egreso.Show
        Else
    MsgBox "Debe ingresar desde una cuenta Administrativa para realizar egresos de efectivo.", vbInformation, "GESTOR DE CAJA"
    End If
End Sub

Sub Boton12(Control As IRibbonControl)
    Hoja33.Select
    Hoja33.Cells(1, 1).Select
    frm_horas.Show
    
End Sub

Sub Boton13(Control As IRibbonControl)
    Hoja7.Select
    Hoja7.Cells(1, 1).Select
    frm_egresosganaderia.Show
End Sub
Sub Boton14(Control As IRibbonControl)
    Hoja8.Select
    Hoja8.Cells(1, 1).Select
    frm_egresosmadera.Show
End Sub
Sub Boton15(Control As IRibbonControl)
    Hoja13.Select
    Hoja13.Cells(1, 1).Select
    frm_egresosgastos.Show
End Sub
Sub Boton16(Control As IRibbonControl)
    Hoja14.Select
    Hoja14.Cells(1, 1).Select
    frm_egresoslegales.Show
End Sub
Sub Boton17(Control As IRibbonControl)
    Hoja15.Select
    Hoja15.Cells(1, 1).Select
    frm_egresosfamiliares.Show
End Sub
Sub Boton18(Control As IRibbonControl)
Application.EnableEvents = False
ThisWorkbook.Save
Application.EnableEvents = True
End Sub
Sub Boton19(Control As IRibbonControl)

End Sub
Sub Boton20(Control As IRibbonControl)
    Load frm_Ganado
    Hoja29.Select
    Hoja29.Cells(1, 1).Select
    frm_Ganado.Show
End Sub
Sub Boton21(Control As IRibbonControl)
    Hoja30.Select
    Hoja30.Cells(1, 1).Select
End Sub
Sub Boton22(Control As IRibbonControl)
    Hoja32.Select
    Hoja32.Cells(1, 1).Select
    frm_enfermeria.Show
End Sub
Sub Boton23(Control As IRibbonControl)
    Hoja31.Select
    Hoja31.Cells(1, 1).Select
End Sub
Sub Boton24(Control As IRibbonControl)
ThisWorkbook.Save
End Sub
Sub Boton25(Control As IRibbonControl)
    Hoja0.Select
    form_iniciosesion.Show
End Sub
Sub Boton26(Control As IRibbonControl)
    Hoja0.Select
    form_iniciosesion.Show
End Sub
Sub Boton27(Control As IRibbonControl)
 
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
    Hoja43.Cells(1, 1).Select
End Sub
Sub Boton31(Control As IRibbonControl)
   Hoja44.Select
   Hoja44.Cells(1, 1).Select
End Sub
Sub Boton32(Control As IRibbonControl)
    Hoja46.Select
    Hoja46.Cells(1, 1).Select
End Sub
Sub Boton33(Control As IRibbonControl)
   Hoja45.Select
   Hoja45.Cells(1, 1).Select
End Sub
Sub Boton34(Control As IRibbonControl)
    Hoja0.Select
    Hoja0.Cells(1, 1).Select
    form_iniciosesion.Show
End Sub
Sub Boton35(Control As IRibbonControl)
   ThisWorkbook.Save
End Sub
Sub Boton36(Control As IRibbonControl)
    If Hoja6.Visible = xlSheetVisible Then
        Hoja6.Select
        Hoja6.Cells(1, 1).Select
    End If
   frm_reportes_producto.Show
End Sub
Sub Boton37(Control As IRibbonControl)
     Hoja0.Select
   Hoja0.Cells(1, 1).Select
   frm_ArqueoResumen.Show
End Sub
Sub Boton38(Control As IRibbonControl)
 If Hoja1.Visible = xlSheetVisible Then
 Hoja1.Select
 Hoja1.Cells(1, 1).Select
 End If
 frm_Codigo.Show
End Sub
Sub Boton39(Control As IRibbonControl)
 If Hoja1.Visible = xlSheetVisible Then
 Hoja1.Select
 Hoja1.Cells(1, 1).Select
 End If
End Sub
Sub Boton40(Control As IRibbonControl)
 If Hoja91.Visible = xlSheetVisible Then
 Hoja91.Select
 Hoja91.Cells(1, 1).Select
 End If
 frm_NuevoUsuario.Show
End Sub
Sub Boton41(Control As IRibbonControl)
 If Hoja91.Visible = xlSheetVisible Then
 Hoja91.Select
 Hoja91.Cells(1, 1).Select
 End If
 frm_EliminarUsuario.Show
End Sub
Sub Boton42(Control As IRibbonControl)
 If Hoja91.Visible = xlSheetVisible Then
 Hoja91.Select
 Hoja91.Cells(1, 1).Select
 End If
 frm_Modificar_Permisos.Show
 
End Sub
Sub Boton43(Control As IRibbonControl)
    Hoja0.Select
    Hoja0.Cells(1, 1).Select
    form_iniciosesion.Show
End Sub
Sub Boton44(Control As IRibbonControl)
 Application.EnableEvents = False
 ThisWorkbook.Save
 Application.EnableEvents = True
End Sub
Sub Boton45(Control As IRibbonControl)
    If Hoja9.Visible = xlSheetVisible Then
        Hoja9.Select
        Hoja9.Cells(1, 1).Select
    End If
End Sub
Sub Boton46(Control As IRibbonControl)
    If Hoja2.Visible = xlSheetVisible Then
     Hoja2.Select
    Hoja2.Cells(1, 1).Select
    End If
    
    frm_Factura.Show
End Sub
Sub Boton47(Control As IRibbonControl)
    Hoja3.Select
    Hoja3.Cells(1, 1).Select
End Sub
Sub Boton48(Control As IRibbonControl)
    If Hoja2.Visible = xlSheetVisible Then
     Hoja2.Select
    Hoja2.Cells(1, 1).Select
    End If
   
    frm_CierreCaja.Show

End Sub
Sub Boton49(Control As IRibbonControl)
    Hoja35.Select
    Hoja35.Cells(1, 1).Select
End Sub
Sub Boton50(Control As IRibbonControl)
    Hoja36.Select
    Hoja36.Cells(1, 1).Select
End Sub
Sub Boton51(Control As IRibbonControl)
    Hoja34.Select
    Hoja34.Cells(1, 1).Select
End Sub
Sub Boton52(Control As IRibbonControl)
 
End Sub
Sub Boton53(Control As IRibbonControl)
    Hoja18.Select
    Hoja18.Cells(1, 1).Select
End Sub
Sub Boton54(Control As IRibbonControl)
    Hoja16.Select
    Hoja16.Cells(1, 1).Select
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
Public Sub DesactivarBoton43(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(43)
    
End Sub
Public Sub DesactivarBoton44(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(44)
    
End Sub
Public Sub DesactivarBoton45(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(45)
    
End Sub
Public Sub DesactivarBoton46(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(46)
    
End Sub
Public Sub DesactivarBoton47(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(47)
    
End Sub
Public Sub DesactivarBoton48(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(48)
    
End Sub
Public Sub DesactivarBoton49(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(49)
    
End Sub
Public Sub DesactivarBoton50(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(50)
    
End Sub
Public Sub DesactivarBoton51(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(51)
    
End Sub
Public Sub DesactivarBoton52(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(52)
    
End Sub
Public Sub DesactivarBoton53(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(53)
    
End Sub
Public Sub DesactivarBoton54(Control As IRibbonControl, ByRef ValorBloqueo)
       ValorBloqueo = RetVal(54)
    
End Sub
