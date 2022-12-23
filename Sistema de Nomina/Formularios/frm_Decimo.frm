VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Decimo 
   Caption         =   "GESTOR DE VENTAS"
   ClientHeight    =   3945
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   8310.001
   OleObjectBlob   =   "frm_Decimo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Decimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private Sub btn_Cargar_Click()
Dim Fecha As Date
Dim Mes As Date
Dim Año As Date
Dim Dia As Date
Dim Seguridad As String
Dim Repetido As String
Dim Fila As Long
Dim Final As Long
Dim encontrado As Boolean

Seguridad = Hoja83.Range("L1").Text

Hoja23.Unprotect (Seguridad)
Hoja1.Unprotect (Seguridad)

Dia = 15

If Me.cboMes.ListIndex = 0 Then
Mes = 4
ElseIf Me.cboMes.ListIndex = 1 Then
Mes = 8
Else
Mes = 12
End If

Año = Me.label_año2.Value

Fecha = DateSerial(Año, Mes, Dia)

Hoja23.Cells(2, 7) = Fecha

Reporte_Decimo

Hoja23.Protect (Seguridad)
Hoja1.Protect (Seguridad)
MsgBox "Cambios Generados con Exito..!", vbInformation, "Gestor de Recursos Humanos"


Unload Me

End Sub

Private Sub btn_Listado_Click()

Dim Fecha As Date
Dim Mes As Date
Dim Año As Date
Dim Dia As Date
Dim Seguridad As String
Dim Repetido As String
Dim Fila As Long
Dim Final As Long
Dim encontrado As Boolean

Seguridad = Hoja83.Range("L1").Text

Hoja23.Unprotect (Seguridad)
Hoja1.Unprotect (Seguridad)

Dia = 15

If Me.cboMes.ListIndex = 0 Then
Mes = 4
ElseIf Me.cboMes.ListIndex = 1 Then
Mes = 8
Else
Mes = 12
End If

Año = Me.label_año2.Value

Fecha = DateSerial(Año, Mes, Dia)

Hoja23.Cells(2, 7) = Fecha

Hoja23.Select

Limpiar_Filtro
Orden_Filtro


Hoja1.Select
ActiveSheet.ListObjects("Tbl_personal").ShowTotals = False

Fila = 2

Do While Hoja1.Cells(Fila, 1) <> Empty
   Fila = Fila + 1
Loop
   Final = Fila - 1

For Fila = 2 To Final

Repetido = Hoja1.Cells(Fila, 1)

Hoja23.Select
Hoja23.Range("A5").Select
Do Until IsEmpty(ActiveCell)
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.Value Like Repetido Then
        encontrado = True
        Exit Do
    Else
        encontrado = False
    End If
Loop

If encontrado = True Then

Else
    Hoja23.Select
    Hoja23.Cells(6, 1).Select
    Selection.ListObject.ListRows.Add (1)

    Hoja23.Cells(6, 1) = Hoja1.Cells(Fila, 1)

End If

Next


Hoja1.Select
ActiveSheet.ListObjects("Tbl_personal").ShowTotals = True

Hoja23.Select

Orden_Filtro


Hoja23.Protect (Seguridad)
Hoja1.Protect (Seguridad)
MsgBox "Cambios Generados con Exito..!", vbInformation, "Gestor de Recursos Humanos"


Unload Me


End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub



Private Sub SpinButton2_Change()
frm_Decimo.label_año2.Value = frm_Decimo.SpinButton2.Value
End Sub

Private Sub UserForm_Initialize()
 With frm_Decimo.cboMes
        .AddItem 1
        .List(0, 1) = "Abril"
        .AddItem 2
        .List(1, 1) = "Agosto"
        .AddItem 3
        .List(2, 1) = "Diciembre"
    End With
    
    frm_Decimo.cboMes.ListIndex = 0
       
    frm_Decimo.SpinButton2.Value = VBA.Year(VBA.Date)
    
    frm_Decimo.label_año2.Value = VBA.Year(VBA.Date)
    
End Sub
Private Sub Limpiar_Filtro()

Range("A5").Select

          If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

           If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData


End Sub
Private Sub Orden_Filtro()

Range("A5").Sort Key1:=Range("A5"), Order1:=xlAscending, Header:=xlYes

End Sub

