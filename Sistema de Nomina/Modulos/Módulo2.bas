Attribute VB_Name = "Módulo2"
Private Sub Limpiar_Filtro()

Range("A1").Select

          If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

           If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData


End Sub
Private Sub Orden_Filtro()

Range("A1").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
Range("B1").Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes

End Sub

Sub ChangeCursor()
 
 Application.Cursor = xlWait

 Application.Cursor = xlDefault
 
End Sub

Sub Fecha()
Dim Fecha As String
Dim Mes As Date
Dim Año As Date
Dim Dia As Date

'Dia = 1
'Mes = VBA.Month(Date)
'Año = VBA.Year(Date)
'
'Fecha = DateSerial(Año, Mes, Dia)
'
'
'
'
'MsgBox Fecha

Dia = 1
Año = 2








End Sub

