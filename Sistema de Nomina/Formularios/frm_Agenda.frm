VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Agenda 
   Caption         =   "Gestor de Recursos Humanos"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   18480
   OleObjectBlob   =   "frm_Agenda.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Agenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public banderaEmpresa As Long
Public Sub AgregarNotas()
     frm_Nota.Show
     Notas
End Sub

Public Sub EliminarNotas()
Dim Agenda As String
    If Me.lbx_Notas.ListIndex = -1 Then
        MsgBox "Debe seleccionar un registro a Eliminar..!", vbInformation
        Exit Sub
    End If

     
    If MsgBox("Esta seguro que desea inhabilitar este registro..?" + Chr(13) + "Desea proceder..?", vbYesNo, "Agenda") = vbNo Then
        Exit Sub
    Else
    
       Eliminar_Nota
       
       With frm_Agenda
        .lbx_Notas.ListIndex = -1
       End With
  
    End If

End Sub

Private Sub btn_ANotas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btn_ANotas.SpecialEffect = fmSpecialEffectSunken
End Sub
Private Sub btn_ANotas_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btn_ANotas.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub btn_ENotas_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btn_ENotas.SpecialEffect = fmSpecialEffectSunken
End Sub
Private Sub btn_ENotas_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btn_ENotas.SpecialEffect = fmSpecialEffectFlat
End Sub
Private Sub btn_ANotas_Click()
 Me.AgregarNotas

End Sub
Private Sub btn_ENotas_Click()
 Me.EliminarNotas
End Sub
Private Sub Eliminar_Nota()
Dim X As String
Dim encontrado As Boolean
Dim Seguridad As String

Seguridad = Hoja83.Range("L1")

Hoja9.Unprotect (Seguridad)

X = Me.lbx_Notas.Column(2)

Hoja9.Select
Range("D1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like X Then
            ActiveCell.Offset(0, 2) = "INACTIVO"
            encontrado = True
            Exit Do
                                 
        End If
    Loop
    If encontrado = True Then
        
        MsgBox "Registro Inhabilitado Correctamente..!", vbInformation, "Agenda"
    
    End If
    
     If encontrado = False Then

        MsgBox "El registro seleccionado no ha sido encontrado.!", vbInformation, "Agenda"
        
     End If
     
     Notas

Hoja9.Unprotect (Seguridad)

End Sub

Private Sub btn_ANotas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.lbx_Notas.ListIndex = -1
SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub btn_ENotas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub btn_salir_Click()
Unload Me
End Sub
Public Sub Notas()
Dim Fila As Long
Dim Final As Long
Dim Estado As String

Dim uf As Long
Dim X As Long

Dim Referencia As String


On Error Resume Next

Estado = "ACTIVO"

frm_Agenda.lbx_Notas.ColumnCount = 5
frm_Agenda.lbx_Notas.ColumnWidths = "35 pt;200 pt;100 pt;850 pt"
frm_Agenda.lbx_Notas.RowSource = "Tbl_Notas"

uf = Hoja3.Range("A" & Rows.Count).End(xlUp).Row

Hoja9.AutoFilterMode = False
frm_Agenda.lbx_Notas = Empty
frm_Agenda.lbx_Notas.RowSource = Empty

For Fila = 2 To uf
   Referencia = Hoja9.Cells(Fila, 6).Value
    If UCase(Referencia) Like Estado Then
        frm_Agenda.lbx_Notas.AddItem
        frm_Agenda.lbx_Notas.List(X, 0) = Hoja9.Cells(Fila, 2).Text
        frm_Agenda.lbx_Notas.List(X, 1) = Hoja9.Cells(Fila, 3).Text
        frm_Agenda.lbx_Notas.List(X, 2) = Hoja9.Cells(Fila, 5).Text
        frm_Agenda.lbx_Notas.List(X, 3) = Hoja9.Cells(Fila, 4).Text
        frm_Agenda.lbx_Notas.List(X, 4) = Hoja9.Cells(Fila, 7).Text
        X = X + 1
   End If
Next
frm_Agenda.lbx_Notas.ColumnCount = 5
frm_Agenda.lbx_Notas.ColumnWidths = "35 pt;200 pt;100 pt;850 pt"

End Sub

Private Sub Frame4_Click()

End Sub

Private Sub UserForm_Initialize()
Dim Fila As Long
Dim Final As Long
Dim Estado As String

Dim uf As Long
Dim X As Long

Dim Referencia As String


On Error Resume Next

Estado = "ACTIVO"

frm_Agenda.lbx_Notas.ColumnCount = 5
frm_Agenda.lbx_Notas.ColumnWidths = "35 pt;200 pt;100 pt;850 pt"
frm_Agenda.lbx_Notas.RowSource = "Tbl_Notas"

uf = Hoja3.Range("A" & Rows.Count).End(xlUp).Row

Hoja9.AutoFilterMode = False
frm_Agenda.lbx_Notas = Empty
frm_Agenda.lbx_Notas.RowSource = Empty

For Fila = 2 To uf
   Referencia = Hoja9.Cells(Fila, 6).Value
    If UCase(Referencia) Like Estado Then
        frm_Agenda.lbx_Notas.AddItem
        frm_Agenda.lbx_Notas.List(X, 0) = Hoja9.Cells(Fila, 2).Text
        frm_Agenda.lbx_Notas.List(X, 1) = Hoja9.Cells(Fila, 3).Text
        frm_Agenda.lbx_Notas.List(X, 2) = Hoja9.Cells(Fila, 5).Text
        frm_Agenda.lbx_Notas.List(X, 3) = Hoja9.Cells(Fila, 4).Text
        frm_Agenda.lbx_Notas.List(X, 4) = Hoja9.Cells(Fila, 7).Text
        X = X + 1
   End If
Next
frm_Agenda.lbx_Notas.ColumnCount = 5
frm_Agenda.lbx_Notas.ColumnWidths = "35 pt;200 pt;100 pt;850 pt"
End Sub
