VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_detalle 
   Caption         =   "Registro de Herramientas"
   ClientHeight    =   8784.001
   ClientLeft      =   100
   ClientTop       =   380
   ClientWidth     =   16780
   OleObjectBlob   =   "frm_detalle.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cancelar_Click()
Unload Me
End Sub
Private Sub btn_Fecha_Click()
 banderaCalendario = 1
    Call LanzarCalendario(Me, "txt_Fecha")
    Me.txt_caja.SetFocus
End Sub
Private Sub btn_caja_Click()
 banderaCaja = 3
    Call LanzarListadoCaja(Me, "label1")
End Sub

Private Sub btn_pieza_Click()

If Me.lbx_herramienta.ListIndex = -1 Then
    MsgBox "Debe seleccionar una herramienta", vbInformation
    Me.lbx_herramienta.SetFocus
    Exit Sub
End If
    
   Call LanzarContenido(Me, "label1")
    With frm_contenido
            .txt_item = frm_detalle.lbx_herramienta.Column(2)
            .txt_herramienta = frm_detalle.lbx_herramienta.Column(3)
            .txt_cantidad = frm_detalle.lbx_herramienta.Column(4)
            .txt_activo = frm_detalle.lbx_herramienta.Column(5)
            .txt_detalle = frm_detalle.lbx_herramienta.Column(6)
            .txt_numero = frm_detalle.lbx_herramienta.Column(0)
            .txt_juego = frm_detalle.lbx_herramienta.Column(7)
    End With
        frm_contenido.Show
txt_caja_Change
End Sub

Private Sub btn_Registrar_Click()
frm_herramienta.Show
txt_caja_Change
End Sub



Private Sub Frame1_Click()

End Sub

Private Sub txt_caja_Change()
On Error Resume Next
Dim Estado As String
Dim Codigo As String

uf = Hoja3.Range("A" & Rows.Count).End(xlUp).Row

Hoja3.AutoFilterMode = False
Me.lbx_herramienta.Clear
Me.lbx_herramienta.RowSource = Clear

Estado = "Activo"

For Fila = 2 To uf
    STRG = Hoja3.Cells(Fila, 3).Value 'Variable para descripción
    Codigo = Hoja3.Cells(Fila, 7).Value 'Variable para codigo
    
    If UCase(STRG) Like UCase(txt_caja.Value) And UCase(Codigo) Like UCase(Estado) Then
        Me.lbx_herramienta.AddItem
        Me.lbx_herramienta.List(x, 0) = Hoja3.Cells(Fila, 1).Value
        Me.lbx_herramienta.List(x, 1) = Hoja3.Cells(Fila, 2).Value
        'Me.lbx_herramienta.List(X, 2) = Hoja3.Cells(Fila, 3).Value
        Me.lbx_herramienta.List(x, 2) = Hoja3.Cells(Fila, 4).Value
        Me.lbx_herramienta.List(x, 3) = Hoja3.Cells(Fila, 5).Value
        Me.lbx_herramienta.List(x, 4) = Hoja3.Cells(Fila, 6).Value
        Me.lbx_herramienta.List(x, 5) = Hoja3.Cells(Fila, 7).Value
        Me.lbx_herramienta.List(x, 6) = Hoja3.Cells(Fila, 8).Value
        Me.lbx_herramienta.List(x, 7) = Hoja3.Cells(Fila, 9).Value

    
        x = x + 1
   '----------------------------------------------------------------------------------
   
    End If
    '----------------------------------------------------------------------------------
Next
Me.lbx_herramienta.ColumnCount = 8
Me.lbx_herramienta.ColumnWidths = "40 pt;80 pt;110 pt;250 pt;70 pt"
End Sub

