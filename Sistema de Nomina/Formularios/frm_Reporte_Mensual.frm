VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Reporte_Mensual 
   Caption         =   "GESTOR DE VENTAS"
   ClientHeight    =   4740
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   4770
   OleObjectBlob   =   "frm_Reporte_Mensual.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Reporte_Mensual"
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

Hoja6.Unprotect (Seguridad)
Hoja1.Unprotect (Seguridad)

Dia = 1
Mes = Me.cboMes.ListIndex + 1
Año = Me.label_año2.Value

Fecha = DateSerial(Año, Mes, Dia)

Hoja6.Cells(2, 3) = Fecha

Hoja6.Select

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

Hoja6.Select
Hoja6.Range("A4").Select
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
    Hoja6.Select
    Hoja6.Cells(5, 1).Select
    Selection.ListObject.ListRows.Add (1)

    Hoja6.Cells(5, 1) = Hoja1.Cells(Fila, 1)

End If

Next

Hoja6.Select
Orden_Filtro


Hoja6.Protect (Seguridad)
Hoja1.Protect (Seguridad)
MsgBox "Cambios Generados con Exito..!", vbInformation, "Gestor de Recursos Humanos"


Unload Me

End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub

Private Sub SpinButton2_Change()
frm_Reporte_Mensual.label_año2.Value = frm_Reporte_Mensual.SpinButton2.Value
End Sub

Private Sub UserForm_Initialize()
 With frm_Reporte_Mensual.cboMes
        .AddItem 1
        .List(0, 1) = "Enero"
        .AddItem 2
        .List(1, 1) = "Febrero"
        .AddItem 3
        .List(2, 1) = "Marzo"
        .AddItem 4
        .List(3, 1) = "Abril"
        .AddItem 5
        .List(4, 1) = "Mayo"
        .AddItem 6
        .List(5, 1) = "Junio"
        .AddItem 7
        .List(6, 1) = "Julio"
        .AddItem 8
        .List(7, 1) = "Agosto"
        .AddItem 9
        .List(8, 1) = "Septiembre"
        .AddItem 10
        .List(9, 1) = "Octubre"
        .AddItem 11
        .List(10, 1) = "Noviembre"
        .AddItem 12
        .List(11, 1) = "Diciembre"
    End With
    
    frm_Reporte_Mensual.cboMes.ListIndex = VBA.Month(VBA.Date) - 1
       
    frm_Reporte_Mensual.SpinButton2.Value = VBA.Year(VBA.Date)
    
    frm_Reporte_Mensual.label_año2.Value = VBA.Year(VBA.Date)
    
End Sub
Private Sub Limpiar_Filtro()

Range("A4").Select

          If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

           If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData


End Sub
Private Sub Orden_Filtro()

Range("A4").Sort Key1:=Range("A4"), Order1:=xlAscending, Header:=xlYes

End Sub
