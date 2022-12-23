Attribute VB_Name = "llamarControl"
Option Explicit
Public banderaPersonal As Long
Public banderaEstado As Long
Public banderaHerramienta As Long
Public banderaCaja As Long
Public banderaCategoria As Long

Public Function LanzarListadoPersonal(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_personal
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_personal.StartUpPosition = 0
            frm_personal.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_personal.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_personal.Show

End Function
Sub InsertarPersonal()

If frm_personal.lbx_personal.ListIndex = -1 Then
    MsgBox "Debe seleccionar un colaborador", vbInformation
    frm_personal.lbx_personal.SetFocus
    Exit Sub
End If

Select Case banderaPersonal
    Case 1
        With frm_cajon
            .txt_id = frm_personal.lbx_personal.Column(1)
            .txt_personal = frm_personal.lbx_personal.Column(2)
            Unload frm_personal
            
        End With
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Public Function LanzarEstado(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_Estado
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_Estado.StartUpPosition = 0
            frm_Estado.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_Estado.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_Estado.Show

End Function
Sub InsertaEstado()

If frm_Estado.lbx_cuenta.ListIndex = -1 Then
    MsgBox "Debe seleccionar un estado", vbInformation
    frm_Estado.lbx_cuenta.SetFocus
    Exit Sub
End If

Select Case banderaEstado
    Case 1
       frm_herramienta.txt_cestado = frm_Estado.lbx_cuenta.Column(0)
      
        Unload frm_Estado
    Case 2
       frm_herramienta.txt_hestado = frm_Estado.lbx_cuenta.Column(0)
      
        Unload frm_Estado
       
       
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub

Public Function LanzarListadoHerramienta(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_Listado
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_Listado.StartUpPosition = 0
            frm_Listado.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_Listado.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_Listado.Show

End Function
Sub InsertarHerramienta()

If frm_Listado.lbx_herramienta.ListIndex = -1 Then
    MsgBox "Debe seleccionar una herramienta", vbInformation
    frm_Listado.lbx_herramienta.SetFocus
    Exit Sub
End If

Select Case banderaHerramienta
    Case 1
        With frm_herramienta
            .txt_codigo = frm_Listado.lbx_herramienta.Column(1)
            .txt_herramienta = frm_Listado.lbx_herramienta.Column(2)
            Unload frm_Listado
            
        End With

    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub

Public Function LanzarListadoCaja(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_caja
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_caja.StartUpPosition = 0
            frm_caja.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_caja.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_caja.Show

End Function
Sub InsertarCaja()

If frm_caja.lbx_cuenta.ListIndex = -1 Then
    MsgBox "Debe seleccionar una caja de herramienta", vbInformation
    frm_caja.lbx_cuenta.SetFocus
    Exit Sub
End If

Select Case banderaCaja
    Case 1
        With frm_pieza
            .txt_caja = frm_caja.lbx_cuenta.Column(2)
            Unload frm_caja
            
        End With
    Case 2
        With frm_cajon
            .txt_caja = frm_caja.lbx_cuenta.Column(1)
            .txt_id = frm_caja.lbx_cuenta.Column(2)
            .txt_personal = frm_caja.lbx_cuenta.Column(3)
            .txt_puesto = frm_caja.lbx_cuenta.Column(4)
            .txt_area = frm_caja.lbx_cuenta.Column(5)
            .txt_estado = frm_caja.lbx_cuenta.Column(6)
            .txt_activo = frm_caja.lbx_cuenta.Column(7)
            .txt_observacion = frm_caja.lbx_cuenta.Column(8)
            
            Unload frm_caja
            
        End With
    Case 3
        With frm_detalle
            .txt_caja = frm_caja.lbx_cuenta.Column(1)
            Unload frm_caja
            
        End With
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub

Public Function LanzarListadoCategoria(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_categoria
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_categoria.StartUpPosition = 0
            frm_categoria.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_categoria.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frm_categoria.Show

End Function
Sub InsertarCategoria()

If frm_categoria.lbx_cuenta.ListIndex = -1 Then
    MsgBox "Debe seleccionar una categoria", vbInformation
    frm_categoria.lbx_cuenta.SetFocus
    Exit Sub
End If

Select Case banderaCategoria
    Case 1
        With frm_pieza
            .txt_id = frm_categoria.lbx_cuenta.Column(0)
            .txt_pieza = frm_categoria.lbx_cuenta.Column(1)
            Unload frm_categoria
            
        End With
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select

End Sub
Public Function LanzarContenido(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_contenido
    
'    For Each xCtrl In CualquierFormulario.Controls
'        If xCtrl.Name = xTextBox Then
'            frm_contenido.StartUpPosition = 0
'            frm_contenido.Left = CualquierFormulario.Left + xCtrl.Left + 5
'            frm_contenido.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
'        End If
'    Next

End Function
Public Function LanzarPieza(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As Control

     Load frm_ajuste
    
'    For Each xCtrl In CualquierFormulario.Controls
'        If xCtrl.Name = xTextBox Then
'            frm_ajuste.StartUpPosition = 0
'            frm_ajuste.Left = CualquierFormulario.Left + xCtrl.Left + 5
'            frm_ajuste.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
'        End If
'    Next

End Function
