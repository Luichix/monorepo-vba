VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_contenido 
   Caption         =   "Gestor"
   ClientHeight    =   7632
   ClientLeft      =   70
   ClientTop       =   300
   ClientWidth     =   19790
   OleObjectBlob   =   "frm_contenido.frx":0000
End
Attribute VB_Name = "frm_contenido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_agregarpieza_Click()
Dim x As String
Dim Pieza As String


x = Me.txt_juego.Text
Hoja6.Select
Range("A1").Select
    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like x Then
            encontrado = True
            Pieza = ActiveCell.Offset(0, 1)
            Exit Do
                                 
        End If
    Loop
    If encontrado = True Then
        
    Load frm_pieza
    
    With frm_pieza
            .txt_caja = frm_detalle.txt_caja.Text
            .txt_id = frm_contenido.txt_juego.Text
            .txt_pieza = Pieza
    End With
        frm_pieza.Show
        txt_herramienta_Change
        
    End If

    If encontrado = False Then
        MsgBox "Esta herramienta no posee componentes", vbInformation, "Gestor de inventarios"
        Exit Sub
    End If
    
End Sub
Private Sub btn_ajuste_Click()
If Me.lbx_pieza.ListIndex = -1 Then
    MsgBox "Debe seleccionar una pieza", vbInformation
    Me.lbx_pieza.SetFocus
    Exit Sub
End If
    
   Call LanzarPieza(Me, "label50")
    With frm_ajuste
            .txt_item = frm_contenido.lbx_pieza.Column(2)
            .txt_pieza = frm_contenido.lbx_pieza.Column(3)
            .txt_cantidad = frm_contenido.lbx_pieza.Column(4)
            .txt_estado = frm_contenido.lbx_pieza.Column(6)
            .txt_detalle = frm_contenido.lbx_pieza.Column(7)
            .txt_numero = frm_contenido.lbx_pieza.Column(0)
    End With
        frm_ajuste.Show
        
        txt_herramienta_Change

End Sub

Private Sub btn_cancelar_Click()
Unload Me
End Sub

Private Sub btn_Fecha_Click()
 banderaCalendario = 1
    Call LanzarCalendario(Me, "txt_Fecha")
End Sub
Private Sub btn_modificar_Click()
On Error GoTo Salir

Application.ScreenUpdating = False

    If Me.txt_Fecha.Text = Empty Or _
        Me.txt_item.Text = Empty Or _
        Me.txt_herramienta.Text = Empty Or _
        Me.txt_activo.Text = Empty Or _
        Me.txt_detalle.Text = Empty Or _
        Me.txt_cantidad.Text = Empty Then
                If Me.txt_numero.Text = Empty Then
                    MsgBox "Notifique al programador, error en la estructura de datos", vbCritical, "luisreynaldo.pch@gmail.com"
                    Exit Sub
                End If
            MsgBox "Hay campos vacíos en el registro", , "Gestor de Inventario de Herramientas"
            Exit Sub
        
    End If
    
    If Me.lbx_pieza.ListCount <> 0 And Me.txt_activo = "Inactivo" Then
        MsgBox "Primero deshabilite los componentes", vbExclamation, "Gestor"
        Exit Sub
    End If
    
If MsgBox("Son correctos los datos?" + Chr(13) + "Desea procesar el registro?", vbYesNo, "Gestor de Inventarios") = vbNo Then
        Exit Sub
    Else

        Modificar
End If

Hoja0.Activate
Hoja0.Select
     Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventario de Herramientas"
 End If

End Sub

Private Sub Modificar()
Dim x As String

x = Me.txt_numero.Text

Hoja3.Select
Range("A1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like x And ActiveCell.Offset(0, 2) Like frm_detalle.txt_caja And ActiveCell.Offset(0, 3) Like Me.txt_item Then
            
            If ActiveCell.Offset(0, 9) = "" And Me.txt_detalle.Text <> "Bueno" Then
                    If Me.txt_activo.Text = "Inactivo" Then
                        ActiveCell.Offset(0, 6) = Me.txt_activo.Text
                        ActiveCell.Offset(0, 7) = Me.txt_detalle.Text
                        ActiveCell.Offset(0, 10) = CDate(Me.txt_Fecha)
                        MsgBox "Registro ha sido inhabilitado correctamente..!", vbInformation, "Gestor"
                    Else
                        ActiveCell.Offset(0, 6) = Me.txt_activo.Text
                        ActiveCell.Offset(0, 7) = Me.txt_detalle.Text
                        ActiveCell.Offset(0, 9) = CDate(Me.txt_Fecha)
                        MsgBox "Registro ha sido modificado Correctamente..!", vbInformation, "Gestor"
                    End If
                    Unload Me
                    ThisWorkbook.Save
                    Exit Do
                    Exit Sub
                    
            ElseIf ActiveCell.Offset(0, 9) <> "" Then
                ActiveCell.Offset(0, 6) = Me.txt_activo.Text
                If Me.txt_activo.Text = "Inactivo" Then
                    ActiveCell.Offset(0, 6) = Me.txt_activo.Text
                    ActiveCell.Offset(0, 10) = CDate(Me.txt_Fecha)
                    MsgBox "Registro ha sido inhabilitado correctamente..!", vbInformation, "Gestor"
                    Unload Me
                    ThisWorkbook.Save
                Else
                    Me.txt_detalle.BackColor = &H8080FF
                    MsgBox "No se puede modificar el detalle del registro.!", vbExclamation, "Gestor"
                End If
                    Exit Do
                    Exit Sub
            ElseIf Me.txt_activo.Text = "Inactivo" Then
                        ActiveCell.Offset(0, 6) = Me.txt_activo.Text
                        ActiveCell.Offset(0, 10) = CDate(Me.txt_Fecha)
                        MsgBox "Registro ha sido inhabilitado correctamente..!", vbInformation, "Gestor"
                        Unload Me
                        ThisWorkbook.Save
                        Exit Do
                        Exit Sub
            Else
                MsgBox "No se ha modificado el registro.!", vbExclamation, "Gestor"
                Exit Do
                Exit Sub
             
            End If
            
        End If
    Loop
    
End Sub





Private Sub txt_cantidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 48 And KeyAscii < 58 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub


Private Sub txt_herramienta_Change()
On Error Resume Next
Dim Estado As String
Dim Codigo As String

uf = Hoja11.Range("A" & Rows.Count).End(xlUp).Row

    Estado = "Activo"
    
    Hoja11.AutoFilterMode = False
    Me.lbx_pieza.Clear
    Me.lbx_pieza.RowSource = Clear
    
    For Fila = 2 To uf
        STRG = Hoja11.Cells(Fila, 6).Value 'Variable para descripción
        Codigo = Hoja11.Cells(Fila, 7).Value 'Variable para codigo
        Componente = Hoja11.Cells(Fila, 8).Value
        
        If UCase(STRG) Like UCase(frm_detalle.txt_caja.Value) And UCase(Codigo) Like UCase(txt_juego) And UCase(Componente) Like UCase(Estado) Then
            Me.lbx_pieza.AddItem
            Me.lbx_pieza.List(x, 0) = Hoja11.Cells(Fila, 1).Value
            Me.lbx_pieza.List(x, 1) = Hoja11.Cells(Fila, 2).Value
            Me.lbx_pieza.List(x, 2) = Hoja11.Cells(Fila, 3).Value
            Me.lbx_pieza.List(x, 3) = Hoja11.Cells(Fila, 4).Value
            Me.lbx_pieza.List(x, 4) = Hoja11.Cells(Fila, 5).Value
            'Me.lbx_pieza.List(x, 5) = Hoja11.Cells(Fila, 6).Value
            Me.lbx_pieza.List(x, 5) = Hoja11.Cells(Fila, 7).Value
            Me.lbx_pieza.List(x, 6) = Hoja11.Cells(Fila, 8).Value
            Me.lbx_pieza.List(x, 7) = Hoja11.Cells(Fila, 9).Value
    
        
            x = x + 1
       '----------------------------------------------------------------------------------
       
        End If
        '----------------------------------------------------------------------------------
    Next
    Me.lbx_pieza.ColumnCount = 8
    Me.lbx_pieza.ColumnWidths = "40 pt;80 pt;80 pt;200 pt;50 pt;70 pt"

End Sub

Private Sub UserForm_Activate()

On Error Resume Next
Dim Estado As String
Dim Codigo As String

uf = Hoja11.Range("A" & Rows.Count).End(xlUp).Row

If txt_juego = "" Then
    frm_contenido.Width = 280
    Exit Sub
Else
 frm_contenido.Width = 1000
     
    Estado = "Activo"
    
    Hoja11.AutoFilterMode = False
    Me.lbx_pieza.Clear
    Me.lbx_pieza.RowSource = Clear
    
    For Fila = 2 To uf
        STRG = Hoja11.Cells(Fila, 6).Value 'Variable para descripción
        Codigo = Hoja11.Cells(Fila, 7).Value 'Variable para codigo
        Componente = Hoja11.Cells(Fila, 8).Value
        
        If UCase(STRG) Like UCase(frm_detalle.txt_caja.Value) And UCase(Codigo) Like UCase(txt_juego) And UCase(Componente) Like UCase(Estado) Then
            Me.lbx_pieza.AddItem
            Me.lbx_pieza.List(x, 0) = Hoja11.Cells(Fila, 1).Value
            Me.lbx_pieza.List(x, 1) = Hoja11.Cells(Fila, 2).Value
            Me.lbx_pieza.List(x, 2) = Hoja11.Cells(Fila, 3).Value
            Me.lbx_pieza.List(x, 3) = Hoja11.Cells(Fila, 4).Value
            Me.lbx_pieza.List(x, 4) = Hoja11.Cells(Fila, 5).Value
            'Me.lbx_pieza.List(x, 5) = Hoja11.Cells(Fila, 6).Value
            Me.lbx_pieza.List(x, 5) = Hoja11.Cells(Fila, 7).Value
            Me.lbx_pieza.List(x, 6) = Hoja11.Cells(Fila, 8).Value
            Me.lbx_pieza.List(x, 7) = Hoja11.Cells(Fila, 9).Value
    
        
            x = x + 1
       '----------------------------------------------------------------------------------
       
        End If
        '----------------------------------------------------------------------------------
    Next
    Me.lbx_pieza.ColumnCount = 8
    Me.lbx_pieza.ColumnWidths = "40 pt;80 pt;80 pt;200 pt;50 pt;70 pt"
End If
End Sub

