VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_EliminarAbono 
   Caption         =   "GESTOR DE VENTA DE ACTIVOS"
   ClientHeight    =   9690.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   18530
   OleObjectBlob   =   "frm_EliminarAbono.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_EliminarAbono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_Cargar_Click()

   Call InsertarEliminarAbono

    If frm_EliminarAbono.lbx_ListadoAbono.ListIndex = -1 Then
        frm_EliminarAbono.lbx_ListadoAbono.SetFocus
        Exit Sub
    End If

    With frm_EliminarAbono
        .Visualizar_EliminarAbono
    End With
  
    With frm_Motivo_EliminarA
        .lbl_nFactura = "N° " & frm_EliminarAbono.txt_referencia.Text
        .txt_id = frm_EliminarAbono.txt_idpersonal.Text
        .txt_nombre = frm_EliminarAbono.txt_nombre.Text
        .txt_detalle = frm_EliminarAbono.txt_Observa.Text
    End With
    
    frm_Motivo_EliminarA.Show
    UserForm_Initialize
End Sub





Private Sub btn_Regresar_Click()
Me.Frame1.Visible = True
Me.Frame2.Visible = False
Me.btn_Visualizar.Visible = True
Me.btn_regresar.Visible = False
Me.btn_Salir.Visible = True
Me.btn_Cargar.Visible = True
Me.btn_Salir2.Visible = False

End Sub
Private Sub btn_salir_Click()
Unload Me
End Sub
Private Sub btn_Salir2_Click()
Unload Me
End Sub
Private Sub btn_visualizar_Click()
   Call InsertarEliminarAbono

    If frm_EliminarAbono.lbx_ListadoAbono.ListIndex = -1 Then
        frm_EliminarAbono.lbx_ListadoAbono.SetFocus
        Exit Sub
    End If

    With frm_EliminarAbono
        .Visualizar_EliminarAbono
    End With

Me.Frame1.Visible = False
Me.Frame2.Visible = True
Me.btn_Visualizar.Visible = False
Me.btn_regresar.Visible = True
Me.btn_Salir.Visible = False
Me.btn_Cargar.Visible = False
Me.btn_Salir2.Visible = True

End Sub
Public Sub Visualizar_EliminarAbono()
Dim Fila As Long
Dim Final As Long
Dim Referencia As String
On Error Resume Next

Referencia = txt_referencia.Text

frm_EliminarAbono.ListBox1.ColumnCount = 5
frm_EliminarAbono.ListBox1.ColumnWidths = "70 pt;170 pt;70 pt;100 pt;50 pt"
frm_EliminarAbono.ListBox1.RowSource = "Tbl_deposito"

uf = Hoja12.Range("A" & Rows.Count).End(xlUp).Row

Hoja12.AutoFilterMode = False
frm_EliminarAbono.ListBox1 = Clear
frm_EliminarAbono.ListBox1.RowSource = Clear

For Fila = 2 To uf
    STRG = Hoja12.Cells(Fila, 9).Value 'Variable para descripción

    If UCase(STRG) Like Referencia Then
        frm_EliminarAbono.ListBox1.AddItem
        frm_EliminarAbono.ListBox1.List(X, 0) = Hoja12.Cells(Fila, 5).Text
        frm_EliminarAbono.ListBox1.List(X, 1) = Hoja12.Cells(Fila, 6).Text
        frm_EliminarAbono.ListBox1.List(X, 2) = Hoja12.Cells(Fila, 7).Value
        frm_EliminarAbono.ListBox1.List(X, 2) = Replace(frm_EliminarAbono.ListBox1.List(X, 2), ",", ".")
        frm_EliminarAbono.ListBox1.List(X, 3) = Hoja12.Cells(Fila, 8).Text
        frm_EliminarAbono.ListBox1.List(X, 4) = Hoja12.Cells(Fila, 10).Text
        frm_EliminarAbono.ListBox1.List(X, 5) = Hoja12.Cells(Fila, 9).Text


        X = X + 1
   End If
Next

frm_EliminarAbono.ListBox1.ColumnCount = 5
frm_EliminarAbono.ListBox1.ColumnWidths = "70 pt;170 pt;70 pt;110 pt;50 pt"

End Sub



Private Sub txt_referencia_Change()
Dim Fila As Long
Dim Final As Long
Dim Actividad As String



If txt_referencia.Text = "" Then
    Me.txt_Observa.Text = ""
    
End If

Final = GetUltimoR(Hoja8)

    For Fila = 2 To Final
        If txt_referencia.Text = Hoja8.Cells(Fila, 17) Then
             Me.txt_Observa.Text = Hoja8.Cells(Fila, 7)
            Exit For
        End If
    Next

End Sub

Private Sub UserForm_Initialize()
Dim Fila As Long
Dim Final As Long
Dim Estado As String
On Error Resume Next

banderaEliminarAbono = 1

Estado = "ACTIVO"

Me.lbx_ListadoAbono.ColumnCount = 10
Me.lbx_ListadoAbono.ColumnWidths = "50 pt;240 pt;90 pt;60 pt;85 pt;85 pt;85 pt;85 pt;85 pt;0 pt"
Me.lbx_ListadoAbono.RowSource = "Tbl_abono"

uf = Hoja8.Range("A" & Rows.Count).End(xlUp).Row

Hoja8.AutoFilterMode = False
Me.lbx_ListadoAbono = Clear
Me.lbx_ListadoAbono.RowSource = Clear

For Fila = 2 To uf
    STRG = Hoja8.Cells(Fila, 19).Value 'Variable para descripción

    If UCase(STRG) Like Estado Then
        Me.lbx_ListadoAbono.AddItem
        Me.lbx_ListadoAbono.List(X, 0) = Hoja8.Cells(Fila, 1).Value
        Me.lbx_ListadoAbono.List(X, 1) = Hoja8.Cells(Fila, 4).Value
        Me.lbx_ListadoAbono.List(X, 2) = Hoja8.Cells(Fila, 8).Value
        Me.lbx_ListadoAbono.List(X, 2) = Replace(Me.lbx_ListadoAbono.List(X, 2), ",", ".")
        Me.lbx_ListadoAbono.List(X, 2) = Format(Me.lbx_ListadoAbono.List(X, 2), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 3) = Hoja8.Cells(Fila, 9).Value & "%"
        Me.lbx_ListadoAbono.List(X, 4) = Hoja8.Cells(Fila, 10).Value
        Me.lbx_ListadoAbono.List(X, 4) = Replace(Me.lbx_ListadoAbono.List(X, 4), ",", ".")
        Me.lbx_ListadoAbono.List(X, 4) = Format(Me.lbx_ListadoAbono.List(X, 4), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 5) = Hoja8.Cells(Fila, 11).Value
        Me.lbx_ListadoAbono.List(X, 5) = Replace(Me.lbx_ListadoAbono.List(X, 5), ",", ".")
        Me.lbx_ListadoAbono.List(X, 5) = Format(Me.lbx_ListadoAbono.List(X, 5), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 6) = Hoja8.Cells(Fila, 14).Text
        Me.lbx_ListadoAbono.List(X, 7) = Hoja8.Cells(Fila, 15).Value
        Me.lbx_ListadoAbono.List(X, 7) = Replace(Me.lbx_ListadoAbono.List(X, 7), ",", ".")
        Me.lbx_ListadoAbono.List(X, 7) = Format(Me.lbx_ListadoAbono.List(X, 7), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 8) = Hoja8.Cells(Fila, 16).Value
        Me.lbx_ListadoAbono.List(X, 8) = Replace(Me.lbx_ListadoAbono.List(X, 8), ",", ".")
        Me.lbx_ListadoAbono.List(X, 8) = Format(Me.lbx_ListadoAbono.List(X, 8), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 9) = Hoja8.Cells(Fila, 17).Value

        X = X + 1

   End If
Next

Me.lbx_ListadoAbono.ColumnCount = 10
Me.lbx_ListadoAbono.ColumnWidths = "50 pt;240 pt;90 pt;60 pt;85 pt;85 pt;85 pt;85 pt;85 pt;0 pt"
Me.txt_busqueda.SetFocus
End Sub
Private Sub txt_busqueda_Change()
Dim Fila As Long
Dim Final As Long
Dim Estado As String
On Error Resume Next

Estado = "ACTIVO"

If txt_busqueda = "" Then


Me.lbx_ListadoAbono.ColumnCount = 10
Me.lbx_ListadoAbono.ColumnWidths = "50 pt;240 pt;90 pt;60 pt;85 pt;85 pt;85 pt;85 pt;85 pt;0 pt"
Me.lbx_ListadoAbono.RowSource = "Tbl_abono"

uf = Hoja8.Range("A" & Rows.Count).End(xlUp).Row

Hoja8.AutoFilterMode = False
Me.lbx_ListadoAbono = Clear
Me.lbx_ListadoAbono.RowSource = Clear

For Fila = 2 To uf
    STRG = Hoja8.Cells(Fila, 19).Value 'Variable para descripción

    If UCase(STRG) Like Estado Then
        Me.lbx_ListadoAbono.AddItem
        Me.lbx_ListadoAbono.List(X, 0) = Hoja8.Cells(Fila, 1).Value
        Me.lbx_ListadoAbono.List(X, 1) = Hoja8.Cells(Fila, 4).Value
        Me.lbx_ListadoAbono.List(X, 2) = Hoja8.Cells(Fila, 8).Value
        Me.lbx_ListadoAbono.List(X, 2) = Replace(Me.lbx_ListadoAbono.List(X, 2), ",", ".")
        Me.lbx_ListadoAbono.List(X, 2) = Format(Me.lbx_ListadoAbono.List(X, 2), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 3) = Hoja8.Cells(Fila, 9).Value & "%"
        Me.lbx_ListadoAbono.List(X, 4) = Hoja8.Cells(Fila, 10).Value
        Me.lbx_ListadoAbono.List(X, 4) = Replace(Me.lbx_ListadoAbono.List(X, 4), ",", ".")
        Me.lbx_ListadoAbono.List(X, 4) = Format(Me.lbx_ListadoAbono.List(X, 4), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 5) = Hoja8.Cells(Fila, 11).Value
        Me.lbx_ListadoAbono.List(X, 5) = Replace(Me.lbx_ListadoAbono.List(X, 5), ",", ".")
        Me.lbx_ListadoAbono.List(X, 5) = Format(Me.lbx_ListadoAbono.List(X, 5), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 6) = Hoja8.Cells(Fila, 14).Text
        Me.lbx_ListadoAbono.List(X, 7) = Hoja8.Cells(Fila, 15).Value
        Me.lbx_ListadoAbono.List(X, 7) = Replace(Me.lbx_ListadoAbono.List(X, 7), ",", ".")
        Me.lbx_ListadoAbono.List(X, 7) = Format(Me.lbx_ListadoAbono.List(X, 7), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 8) = Hoja8.Cells(Fila, 16).Value
        Me.lbx_ListadoAbono.List(X, 8) = Replace(Me.lbx_ListadoAbono.List(X, 8), ",", ".")
        Me.lbx_ListadoAbono.List(X, 8) = Format(Me.lbx_ListadoAbono.List(X, 8), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 9) = Hoja8.Cells(Fila, 17).Value

        X = X + 1

   End If
Next

Me.lbx_ListadoAbono.ColumnCount = 10
Me.lbx_ListadoAbono.ColumnWidths = "50 pt;240 pt;90 pt;60 pt;85 pt;85 pt;85 pt;85 pt;85 pt;0 pt"

    Exit Sub
End If

Estado = "ACTIVO"

Me.lbx_ListadoAbono.ColumnCount = 10
Me.lbx_ListadoAbono.ColumnWidths = "50 pt;240 pt;90 pt;60 pt;85 pt;85 pt;85 pt;85 pt;85 pt;0 pt"
Me.lbx_ListadoAbono.RowSource = "Tbl_abono"


uf = Hoja8.Range("A" & Rows.Count).End(xlUp).Row

Hoja8.AutoFilterMode = False
Me.lbx_ListadoAbono = Clear
Me.lbx_ListadoAbono.RowSource = Clear

For Fila = 2 To uf
    STRG = Hoja8.Cells(Fila, 19).Value 'Variable para descripción
    Codigo = Hoja8.Cells(Fila, 4).Value 'Variable para codigo

     If UCase(STRG) Like Estado And UCase(Codigo) Like "*" & UCase(txt_busqueda.Value) & "*" Then
        Me.lbx_ListadoAbono.AddItem
        Me.lbx_ListadoAbono.List(X, 0) = Hoja8.Cells(Fila, 1).Value
        Me.lbx_ListadoAbono.List(X, 1) = Hoja8.Cells(Fila, 4).Value
        Me.lbx_ListadoAbono.List(X, 2) = Hoja8.Cells(Fila, 8).Value
        Me.lbx_ListadoAbono.List(X, 2) = Replace(Me.lbx_ListadoAbono.List(X, 2), ",", ".")
        Me.lbx_ListadoAbono.List(X, 2) = Format(Me.lbx_ListadoAbono.List(X, 2), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 3) = Hoja8.Cells(Fila, 9).Value & "%"
        Me.lbx_ListadoAbono.List(X, 4) = Hoja8.Cells(Fila, 10).Value
        Me.lbx_ListadoAbono.List(X, 4) = Replace(Me.lbx_ListadoAbono.List(X, 4), ",", ".")
        Me.lbx_ListadoAbono.List(X, 4) = Format(Me.lbx_ListadoAbono.List(X, 4), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 5) = Hoja8.Cells(Fila, 11).Value
        Me.lbx_ListadoAbono.List(X, 5) = Replace(Me.lbx_ListadoAbono.List(X, 5), ",", ".")
        Me.lbx_ListadoAbono.List(X, 5) = Format(Me.lbx_ListadoAbono.List(X, 5), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 6) = Hoja8.Cells(Fila, 14).Text
        Me.lbx_ListadoAbono.List(X, 7) = Hoja8.Cells(Fila, 15).Value
        Me.lbx_ListadoAbono.List(X, 7) = Replace(Me.lbx_ListadoAbono.List(X, 7), ",", ".")
        Me.lbx_ListadoAbono.List(X, 7) = Format(Me.lbx_ListadoAbono.List(X, 7), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 8) = Hoja8.Cells(Fila, 16).Value
        Me.lbx_ListadoAbono.List(X, 8) = Replace(Me.lbx_ListadoAbono.List(X, 8), ",", ".")
        Me.lbx_ListadoAbono.List(X, 8) = Format(Me.lbx_ListadoAbono.List(X, 8), "#,##0.00")
        Me.lbx_ListadoAbono.List(X, 9) = Hoja8.Cells(Fila, 17).Value

        X = X + 1

   End If
Next

Me.lbx_ListadoAbono.ColumnCount = 10
Me.lbx_ListadoAbono.ColumnWidths = "50 pt;240 pt;90 pt;60 pt;85 pt;85 pt;85 pt;85 pt;85 pt;0 pt"


End Sub


