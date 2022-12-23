Attribute VB_Name = "ModuloPeriodo"
Option Explicit
Option Private Module
Private SenalCambioMes As Long

Public Sub RecibeElPeriodo(Dia As Long, Mes As Long, Ano As Long)
    Dim FechaRecibida As Date
    FechaRecibida = VBA.DateSerial((VBA.CInt(Ano)), (VBA.CInt(Mes)), (VBA.CInt(Dia)))
    
    'DIRECCIONE LA FECHA AL CONTROL O CELDA QUE REQUIERA
    
       Call InsertarPeriodo(FechaRecibida)
    
End Sub

'********************************** NO MODIFICAR SI NO SABE **********************************
'*************************************|||||||||||||||||||*************************************
'***************************************|||||||||||||||***************************************
'*****************************************|||||||||||*****************************************
'*******************************************|||||||*******************************************
Public Sub InicializaFormularioPeriodo()
    SenalCambioMes = 1
    
    With frm_Periodo.cboMes
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
    
    frm_Periodo.cboMes.ListIndex = VBA.Month(VBA.Date) - 1
    
    frm_Periodo.spbAño.Value = VBA.Year(VBA.Date)
    
    frm_Periodo.lblAno.Caption = VBA.Year(VBA.Date)
    
    Dim Ano As Long, Mes As Long
    Ano = VBA.Year(VBA.Date)
    Mes = VBA.Month(VBA.Date)
    Call ModuloPeriodo.CargarLosDiasPeriodo(Ano, Mes)
    
    frm_Periodo.lblHoy.Caption = VBA.Date
End Sub

Public Sub CargarLosDiasPeriodo(Ano As Long, Mes As Long)
'    Dim FechaDelPrimerDia As Date
'    Dim FechaDelUltimoDia As Date
'    Dim DiaSemanaPrimerDia As Long
'    Dim VariableControl As Control
'    Dim contador As Long
'
'    FechaDelPrimerDia = VBA.DateSerial(Ano, Mes, 1)
'    FechaDelUltimoDia = Application.WorksheetFunction.EoMonth(VBA.DateSerial(Ano, Mes, 1), 0)
'    DiaSemanaPrimerDia = Application.WorksheetFunction.Weekday(FechaDelPrimerDia, 2)
'    contador = 1
'
'    For Each VariableControl In frm_Periodo.mrcDias.Controls
'        VariableControl.Caption = "-"
'        If VariableControl.Tag >= DiaSemanaPrimerDia And contador <= VBA.Day(FechaDelUltimoDia) Then
'            VariableControl.Caption = contador
'            contador = contador + 1
'        End If
'    Next VariableControl
End Sub

Public Sub CambioDeMesPeriodo()
    If SenalCambioMes > 1 Then
        Dim MesEnElCombo As Long, AnoEnElLabel As Long
        
        If Not (IsNull(frm_Periodo.cboMes.Value)) And Not (IsNull(frm_Periodo.lblAno.Caption)) Then
            MesEnElCombo = VBA.CLng(frm_Periodo.cboMes.Value)
            AnoEnElLabel = VBA.CLng(frm_Periodo.lblAno.Caption)
            Call ModuloPeriodo.DesMarcarDiaPeriodos
            Call ModuloPeriodo.CargarLosDiasPeriodo(AnoEnElLabel, MesEnElCombo)
        End If
    End If
    SenalCambioMes = SenalCambioMes + 1
End Sub

Public Sub CambioDeAnoPeriodo()
    Dim MesEnElCombo As Long, AnoEnElLabel As Long
    
    frm_Periodo.lblAno.Caption = frm_Periodo.spbAño.Value
    
    MesEnElCombo = VBA.CLng(frm_Periodo.cboMes.Value)
    AnoEnElLabel = VBA.CLng(frm_Periodo.lblAno.Caption)
    Call ModuloPeriodo.DesMarcarDiaPeriodos
    Call ModuloPeriodo.CargarLosDiasPeriodo(AnoEnElLabel, MesEnElCombo)
    
End Sub

Public Sub UnClickEnHoyEsPeriodo()
    Dim Mes As Long, Ano As Long
    Dim FechaActual As Date
    
    FechaActual = VBA.CDate(frm_Periodo.lblHoy.Caption)
    Mes = VBA.CLng(VBA.Month(FechaActual))
    Ano = VBA.CLng(VBA.Year(FechaActual))
    
    frm_Periodo.lblAno.Caption = Ano
    frm_Periodo.cboMes.ListIndex = Mes - 1
    frm_Periodo.spbAño.Value = Ano
    frm_Periodo.spbAño.SetFocus
    
    Call ModuloPeriodo.DesMarcarDiaPeriodos
    Call ModuloPeriodo.CargarLosDiasPeriodo(Ano, Mes)
    
End Sub

Sub SalirConEscapePeriodo()
    Unload frm_Periodo
End Sub

Sub MarcarDiaPeriodo(ControlDeEtiqueta As Control)
    Call ModuloPeriodo.DesMarcarDiaPeriodos
    ControlDeEtiqueta.Font.Bold = True
    ControlDeEtiqueta.ForeColor = VBA.RGB(255, 0, 0)
End Sub

Sub DesMarcarDiaPeriodos()
    Dim ControlEtiqueta As Control
    
    For Each ControlEtiqueta In frm_Periodo.mrcDias.Controls
        ControlEtiqueta.Font.Bold = False
        ControlEtiqueta.ForeColor = VBA.RGB(0, 0, 0)
    Next ControlEtiqueta
End Sub

'*******************************************|||||||*******************************************
'*****************************************|||||||||||*****************************************
'***************************************|||||||||||||||***************************************
'*************************************||||||||||||||||||**************************************
'********************************** NO MODIFICAR SI NO SABE **********************************

' Nota del autor -----------------------------------------------------------------------------

' Creado por Andrés Rojas Moncada - Autor del canal Excel Hecho Fácil en YouTube

' Versión 1.0 - 20 de julio de 2015

' URL del canal: www.youtube.com/jarmoncada01

' Si quieres usarlo, solo copia y pega el presente módulo en conjunto con el UserForm y listo.

' Para ver algunos ejemplos sobre el uso de este calendario, observa este video.

' Enlace: |||||||||||||||||||| https://www.youtube.com/watch?v=FkjsuN2zqSU ||||||||||||||||||||

' ! Muchas gracias y espero lo disfruten ! ---------------------------------------------------


