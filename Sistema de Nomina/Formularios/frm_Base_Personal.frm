VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Base_Personal 
   Caption         =   "GESTIÓN DE PERSONAL"
   ClientHeight    =   10620
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   16050
   OleObjectBlob   =   "frm_Base_Personal.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Base_Personal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Unload Me

End Sub

Private Sub UserForm_Initialize()
Dim Fila As Long
Dim Final As Long
Dim Estado As String
On Error Resume Next

Estado = "ACTIVO"

Me.lbx_Personal.ColumnCount = 8
Me.lbx_Personal.ColumnWidths = "45 pt;240 pt;75 pt;100 pt;75 pt;75 pt"
Me.lbx_Personal.RowSource = "Tbl_personal"


uf = Hoja1.Range("A" & Rows.Count).End(xlUp).Row

Hoja1.AutoFilterMode = False
Me.lbx_Personal = Clear
Me.lbx_Personal.RowSource = Clear

For Fila = 2 To uf
    STRG = Hoja1.Cells(Fila, 16).Value 'Variable para descripción

    If UCase(STRG) Like Estado Then
        Me.lbx_Personal.AddItem
        Me.lbx_Personal.List(X, 0) = Hoja1.Cells(Fila, 1).Value
        Me.lbx_Personal.List(X, 1) = Hoja1.Cells(Fila, 2).Value
        Me.lbx_Personal.List(X, 2) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_Personal.List(X, 3) = Hoja1.Cells(Fila, 11).Value
        Me.lbx_Personal.List(X, 4) = Hoja1.Cells(Fila, 12).Value
        Me.lbx_Personal.List(X, 5) = Hoja1.Cells(Fila, 13).Value
        Me.lbx_Personal.List(X, 6) = Hoja1.Cells(Fila, 16).Value
        Me.lbx_Personal.List(X, 7) = Hoja1.Cells(Fila, 18).Value
        
        X = X + 1
   End If
Next

Me.lbx_Personal.ColumnCount = 8
Me.lbx_Personal.ColumnWidths = "45 pt;240 pt;75 pt;100 pt;75 pt;75 pt"

Me.TextBox1.SetFocus
End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub
Private Sub Ckx_Inactivo_Click()
Dim Fila As Long
Dim Final As Long
Dim Estado As String
Dim Inactivo As String
On Error Resume Next

If Me.Ckx_Inactivo.Value = False Then

Estado = "ACTIVO"
Inactivo = "INACTIVO"

Me.lbx_Personal.ColumnCount = 7
Me.lbx_Personal.ColumnWidths = "45 pt;240 pt;75 pt;100 pt;75 pt;75 pt"
Me.lbx_Personal.RowSource = "Tbl_personal"

uf = Hoja1.Range("A" & Rows.Count).End(xlUp).Row

Hoja1.AutoFilterMode = False
Me.lbx_Personal = Clear
Me.lbx_Personal.RowSource = Clear

For Fila = 2 To uf
    STRG = Hoja1.Cells(Fila, 16).Value 'Variable para descripción

    If UCase(STRG) Like Estado Then
        Me.lbx_Personal.AddItem
        Me.lbx_Personal.List(X, 0) = Hoja1.Cells(Fila, 1).Value
        Me.lbx_Personal.List(X, 1) = Hoja1.Cells(Fila, 2).Value
        Me.lbx_Personal.List(X, 2) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_Personal.List(X, 3) = Hoja1.Cells(Fila, 11).Value
        Me.lbx_Personal.List(X, 4) = Hoja1.Cells(Fila, 12).Value
        Me.lbx_Personal.List(X, 5) = Hoja1.Cells(Fila, 13).Value
        Me.lbx_Personal.List(X, 6) = Hoja1.Cells(Fila, 16).Value
        Me.lbx_Personal.List(X, 7) = Hoja1.Cells(Fila, 18).Value
        
        X = X + 1
   End If
Next


ElseIf Me.Ckx_Inactivo.Value = True Then

Me.lbx_Personal.ColumnCount = 7
Me.lbx_Personal.ColumnWidths = "45 pt;240 pt;75 pt;100 pt;75 pt;75 pt"
Me.lbx_Personal.RowSource = "Tbl_personal"


uf = Hoja1.Range("A" & Rows.Count).End(xlUp).Row

Hoja1.AutoFilterMode = False
Me.lbx_Personal = Clear
Me.lbx_Personal.RowSource = Clear

For Fila = 2 To uf
        Me.lbx_Personal.AddItem
        Me.lbx_Personal.List(X, 0) = Hoja1.Cells(Fila, 1).Value
        Me.lbx_Personal.List(X, 1) = Hoja1.Cells(Fila, 2).Value
        Me.lbx_Personal.List(X, 2) = Hoja1.Cells(Fila, 4).Value
        Me.lbx_Personal.List(X, 3) = Hoja1.Cells(Fila, 11).Value
        Me.lbx_Personal.List(X, 4) = Hoja1.Cells(Fila, 12).Value
        Me.lbx_Personal.List(X, 5) = Hoja1.Cells(Fila, 13).Value
        Me.lbx_Personal.List(X, 6) = Hoja1.Cells(Fila, 16).Value
        Me.lbx_Personal.List(X, 7) = Hoja1.Cells(Fila, 18).Value
        
        X = X + 1
Next
End If
Me.lbx_Personal.ColumnCount = 7
Me.lbx_Personal.ColumnWidths = "45 pt;240 pt;75 pt;100 pt;75 pt;75 pt"

Me.TextBox1.SetFocus
End Sub
Private Sub TextBox1_Change()
Dim Actividad As String
On Error Resume Next


uf = Hoja1.Range("A" & Rows.Count).End(xlUp).Row

Hoja1.AutoFilterMode = False
Me.lbx_Personal = Clear
Me.lbx_Personal.RowSource = Clear

            If Me.Ckx_Inactivo.Value = False Then
            
                For Fila = 2 To uf
                    STRG = Hoja1.Cells(Fila, 2).Value 'Variable para descripción
                    Codigo = Hoja1.Cells(Fila, 1).Value 'Variable para codigo

                        If UCase(STRG) Like "*" & UCase(TextBox1.Value) & "*" Then
                            If Hoja1.Cells(Fila, 16).Text = "ACTIVO" Then
                            Me.lbx_Personal.AddItem
                                    Me.lbx_Personal.List(X, 0) = Hoja1.Cells(Fila, 1).Value
                                    Me.lbx_Personal.List(X, 1) = Hoja1.Cells(Fila, 2).Value
                                    Me.lbx_Personal.List(X, 2) = Hoja1.Cells(Fila, 4).Value
                                    Me.lbx_Personal.List(X, 3) = Hoja1.Cells(Fila, 11).Value
                                    Me.lbx_Personal.List(X, 4) = Hoja1.Cells(Fila, 12).Value
                                    Me.lbx_Personal.List(X, 5) = Hoja1.Cells(Fila, 13).Value
                                    Me.lbx_Personal.List(X, 6) = Hoja1.Cells(Fila, 16).Value
                                    Me.lbx_Personal.List(X, 7) = Hoja1.Cells(Fila, 18).Value
                            X = X + 1
                            End If
                       '----------------------------------------------------------------------------------
                        'He añadido todo este fragmento para que me busque al mismo tiempo por codigo.
                        ElseIf Codigo Like "*" & UCase(TextBox1.Value) & "*" Then
                            If Hoja1.Cells(Fila, 16).Text = "ACTIVO" Then
                            Me.lbx_Personal.AddItem
                                    Me.lbx_Personal.List(X, 0) = Hoja1.Cells(Fila, 1).Value
                                    Me.lbx_Personal.List(X, 1) = Hoja1.Cells(Fila, 2).Value
                                    Me.lbx_Personal.List(X, 2) = Hoja1.Cells(Fila, 4).Value
                                    Me.lbx_Personal.List(X, 3) = Hoja1.Cells(Fila, 11).Value
                                    Me.lbx_Personal.List(X, 4) = Hoja1.Cells(Fila, 12).Value
                                    Me.lbx_Personal.List(X, 5) = Hoja1.Cells(Fila, 13).Value
                                    Me.lbx_Personal.List(X, 6) = Hoja1.Cells(Fila, 16).Value
                                    Me.lbx_Personal.List(X, 7) = Hoja1.Cells(Fila, 18).Value
                            X = X + 1
                            End If
                        End If
                        '----------------------------------------------------------------------------------
                    Next
            ElseIf Me.Ckx_Inactivo.Value = True Then
            
                    For Fila = 2 To uf
                        STRG = Hoja1.Cells(Fila, 2).Value 'Variable para descripción
                        Codigo = Hoja1.Cells(Fila, 1).Value 'Variable para codigo
                     
                        If UCase(STRG) Like "*" & UCase(TextBox1.Value) & "*" Then
                            Me.lbx_Personal.AddItem
                                    Me.lbx_Personal.List(X, 0) = Hoja1.Cells(Fila, 1).Value
                                    Me.lbx_Personal.List(X, 1) = Hoja1.Cells(Fila, 2).Value
                                    Me.lbx_Personal.List(X, 2) = Hoja1.Cells(Fila, 4).Value
                                    Me.lbx_Personal.List(X, 3) = Hoja1.Cells(Fila, 11).Value
                                    Me.lbx_Personal.List(X, 4) = Hoja1.Cells(Fila, 12).Value
                                    Me.lbx_Personal.List(X, 5) = Hoja1.Cells(Fila, 13).Value
                                    Me.lbx_Personal.List(X, 6) = Hoja1.Cells(Fila, 16).Value
                                    Me.lbx_Personal.List(X, 7) = Hoja1.Cells(Fila, 18).Value
                            X = X + 1
                                                  
                        ElseIf Codigo Like "*" & UCase(TextBox1.Value) & "*" Then
                            Me.lbx_Personal.AddItem
                                    Me.lbx_Personal.List(X, 0) = Hoja1.Cells(Fila, 1).Value
                                    Me.lbx_Personal.List(X, 1) = Hoja1.Cells(Fila, 2).Value
                                    Me.lbx_Personal.List(X, 2) = Hoja1.Cells(Fila, 4).Value
                                    Me.lbx_Personal.List(X, 3) = Hoja1.Cells(Fila, 11).Value
                                    Me.lbx_Personal.List(X, 4) = Hoja1.Cells(Fila, 12).Value
                                    Me.lbx_Personal.List(X, 5) = Hoja1.Cells(Fila, 13).Value
                                    Me.lbx_Personal.List(X, 6) = Hoja1.Cells(Fila, 16).Value
                                    Me.lbx_Personal.List(X, 7) = Hoja1.Cells(Fila, 18).Value
                            X = X + 1
                        End If
                    Next
            End If
            
            
Me.lbx_Personal.ColumnCount = 7
Me.lbx_Personal.ColumnWidths = "45 pt;240 pt;75 pt;100 pt;75 pt;75 pt"

End Sub


