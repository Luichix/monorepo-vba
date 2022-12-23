Attribute VB_Name = "Procedimientos"
Option Explicit
Public banderaListadoCuentas As Long
Public nGrupo As Long

Sub ValidarCuenta()
Dim Fila As Long
Dim Final As Long
Dim encontrado As Boolean

With frm_CatalogoCuentas
    Final = nReg(Hoja40, 2, 1) - 1

    For Fila = 2 To Final
        If Hoja40.Cells(Fila, 1) = Val(Mid(.cbo_CodCuenta, 1, 1)) _
        Or .cbo_CodCuenta = Empty Then
            encontrado = True
            nGrupo = Hoja40.Cells(Fila, 1)
            Exit Sub
        End If
    Next

        If encontrado = False Then
            MsgBox "La cuenta: " & .cbo_CodCuenta & _
            " aún no se ha establecido en los parámetros", vbInformation
            .cbo_CodCuenta = Empty
            .cbo_CodCuenta.BackColor = RGB(211, 255, 211)
            .cbo_CodCuenta.SetFocus
            Exit Sub
        End If
End With
End Sub

Sub CodCuentaATexto()
Dim Celda As Object
Dim miRangoDinamico As String
Dim Rango As Range
Dim Final As Long

Final = nReg(Hoja41, 2, 1) - 1

miRangoDinamico = "A" & 2 & ":" & "A" & Final

Hoja41.Range(miRangoDinamico).NumberFormat = "@"

Set Rango = Hoja41.Range(miRangoDinamico)
        
    For Each Celda In Rango
        Celda.Value = CStr(Celda)
    Next Celda
End Sub

Sub CodCuentaANumero()
Dim Celda As Object
Dim miRangoDinamico As String
Dim Rango As Range
Dim Final As Long

Final = nReg(Hoja41, 2, 1) - 1

miRangoDinamico = "A" & 2 & ":" & "A" & Final

Hoja41.Range(miRangoDinamico).NumberFormat = "General"

Set Rango = Hoja41.Range(miRangoDinamico)
        
    For Each Celda In Rango
        Celda.Value = Val(Celda)
    Next Celda
    
End Sub

Sub IndexarCodCuentasPLAN()
    Call CodCuentaATexto
    
    Hoja41.Range("A:C").Sort key1:=Hoja41.Range("A2"), _
    order1:=xlAscending, Header:=xlYes
  
    Call CodCuentaANumero
End Sub

Sub InsertarCuentadesdeListBox()

If frm_ListadoCuentas.lbx_Cuentas.ListIndex = -1 Then
    MsgBox "Debe seleccionar una cuenta", vbInformation
    frm_ListadoCuentas.lbx_Cuentas.SetFocus
    Exit Sub
End If

Select Case banderaListadoCuentas
    Case 1
        With frm_CatalogoCuentas
            .cbo_CodCuenta = frm_ListadoCuentas.lbx_Cuentas.Column(0)
            .txt_NombreCuenta = frm_ListadoCuentas.lbx_Cuentas.Column(1)
            Unload frm_ListadoCuentas
        End With
    
    Case 2
        With frm_LibroDiario
            .cbo_CodCuenta = frm_ListadoCuentas.lbx_Cuentas.Column(0)
            .txt_NombreCuenta = frm_ListadoCuentas.lbx_Cuentas.Column(1)
            Unload frm_ListadoCuentas
        End With
    
    Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
End Select



End Sub

Sub BuscarItemEnListBox()
Dim i As Long

Select Case banderaListadoCuentas
  
    Case 1
        For i = 0 To frm_ListadoCuentas.lbx_Cuentas.ListCount - 1
            If frm_ListadoCuentas.lbx_Cuentas.List(i, 0) = frm_CatalogoCuentas.cbo_CodCuenta Then
                frm_ListadoCuentas.lbx_Cuentas.ListIndex = i
                Exit For
            End If
        Next
    
    Case 2
        For i = 0 To frm_ListadoCuentas.lbx_Cuentas.ListCount - 1
            If frm_ListadoCuentas.lbx_Cuentas.List(i, 0) = frm_LibroDiario.cbo_CodCuenta Then
                frm_ListadoCuentas.lbx_Cuentas.ListIndex = i
                Exit For
            End If
        Next
        
    End Select
End Sub

Sub CambiarTamanoListboxCuentas()

If banderaListadoCuentas = 1 Then
    frm_ListadoCuentas.Height = 125.25
    frm_ListadoCuentas.lbx_Cuentas.Height = 75
End If

End Sub

Sub sumarDebe()
Dim item As Long
Dim totDebe As Currency

On Error Resume Next

With frm_LibroDiario

    totDebe = 0

        For item = 0 To frm_LibroDiario.lbx_DebeHaber.ListCount - 1
            .lbx_DebeHaber.List(item, 2) = _
            Replace(.lbx_DebeHaber.List(item, 2), Application.ThousandsSeparator, "")  'Aquí elimino el separador de miles
            
            .lbx_DebeHaber.List(item, 2) = _
            Replace(.lbx_DebeHaber.List(item, 2), ",", ".") 'Ahora sustituyo la coma decimal por el punto decimal, para poder hacer la sumatoria con la variable totDebe, ya que con la coma decimal, no se suman los decimales
            
            totDebe = totDebe + Val(.lbx_DebeHaber.List(item, 2))
                        
            .lbx_DebeHaber.List(item, 2) = _
            Replace(.lbx_DebeHaber.List(item, 2), ".", Application.DecimalSeparator)  'Aquí devuelvo el formato decimal para que no afecte al ListBox
            
            .lbx_DebeHaber.List(item, 2) = FormatNumber(.lbx_DebeHaber.List(item, 2), 2) 'Aqui doy formato de moneda para que aparezcan los separadores de miles y decimales
            
        Next item

    .lbl_SumaDebe.Caption = totDebe
    
    
        .lbl_Diferencia.Caption = .lbl_SumaDebe.Caption - .lbl_SumaHaber.Caption
        
            If .lbl_SumaDebe.Caption - .lbl_SumaHaber.Caption = 0 Then
                .lbl_Diferencia.ForeColor = RGB(255, 255, 255)
            Else
                .lbl_Diferencia.ForeColor = RGB(255, 0, 0)
            End If
        
        
        .lbl_Diferencia.Caption = FormatNumber(.lbl_Diferencia.Caption, 2)
        .lbl_SumaDebe.Caption = FormatNumber(.lbl_SumaDebe.Caption, 2)
        .lbl_SumaHaber.Caption = FormatNumber(.lbl_SumaHaber.Caption, 2)
    
    
End With

End Sub

Sub sumarhaber()
Dim item As Long
Dim totHaber As Currency

On Error Resume Next
    
With frm_LibroDiario

    totHaber = 0

        For item = 0 To frm_LibroDiario.lbx_DebeHaber.ListCount - 1
        
        
            .lbx_DebeHaber.List(item, 3) = _
            Replace(.lbx_DebeHaber.List(item, 3), Application.ThousandsSeparator, "")  'Aquí elimino el separador de miles
            
            .lbx_DebeHaber.List(item, 3) = _
            Replace(.lbx_DebeHaber.List(item, 3), ",", ".") 'Ahora sustituyo la coma decimal por el punto decimal, para poder hacer la sumatoria con la variable totHaber, ya que con la coma decimal, no se suman los decimales
            
            totHaber = totHaber + Val(.lbx_DebeHaber.List(item, 3))
                        
            .lbx_DebeHaber.List(item, 3) = _
            Replace(.lbx_DebeHaber.List(item, 3), ".", Application.DecimalSeparator)  'Aquí devuelvo el formato decimal para que no afecte al ListBox
            
            .lbx_DebeHaber.List(item, 3) = FormatNumber(.lbx_DebeHaber.List(item, 3), 2)
        
        Next item

    .lbl_SumaHaber.Caption = totHaber
    
    
        .lbl_Diferencia.Caption = .lbl_SumaDebe.Caption - .lbl_SumaHaber.Caption
        
            If .lbl_SumaDebe.Caption - .lbl_SumaHaber.Caption = 0 Then
                .lbl_Diferencia.ForeColor = RGB(255, 255, 255)
            Else
                .lbl_Diferencia.ForeColor = RGB(255, 0, 0)
            End If
        
        
        .lbl_Diferencia.Caption = FormatNumber(.lbl_Diferencia.Caption, 2)
        .lbl_SumaDebe.Caption = FormatNumber(.lbl_SumaDebe.Caption, 2)
        .lbl_SumaHaber.Caption = FormatNumber(.lbl_SumaHaber.Caption, 2)
    
    
End With
    
End Sub

Sub EnviarAMayor()
    Dim ccCelda As Range, ccRango As Range
    Dim ldCelda As Range, ldRango As Range
    Dim lmFila As Long
    
     
    
 Application.ScreenUpdating = False
 
 Hoja43.Activate ' Libro Mayor
 
    Cells.Select
    Selection.Clear
 

lmFila = 2
    
  
Hoja41.Activate ' Catálogo de Cuentas
    
    Set ccRango = Hoja41.Range(Cells(2, 1), Cells(2, 1).End(xlDown)) 'Preparando el Rango del Catálogo de Cuentas
    
    For Each ccCelda In ccRango ' Checando cada celda en el catálogo de cuentas Hoja41
        

            
            If Len(ccCelda) = 3 Then
            
                Hoja42.Activate ' Libro Diario
                
                Set ldRango = Hoja42.Range(Cells(2, 4), Cells(2, 4).End(xlDown)) 'Preparando el Rango del Libro Diario
                For Each ldCelda In ldRango
                    If ccCelda = Val(Mid(ldCelda.Offset(0, 0), 1, 3)) Then  ' Comparo la CELDA de la Hoja41 Catálogo de Cuentas, con la Hoja42 Libro Diario.
                       ' y escribo los datos en la hoj4 Libro Mayor
        With Hoja43
                        .Cells(1, 1) = "CUENTA"
                        .Cells(1, 2) = "NOMBRE DE LA CUENTA"
                        .Cells(1, 3) = "#"
                        .Cells(1, 4) = "FECHA"
                        .Cells(1, 5) = "DEBE"
                        .Cells(1, 6) = "HABER"
                        
                        .Cells(lmFila, 1) = ccCelda.Offset(0, 0).Value ' No. de Cuenta proviene de la hoja 2
                        .Cells(lmFila, 2) = ccCelda.Offset(0, 1).Value ' Nombre de la Cuenta proviene de la hoja 2
                        
                        
                            If ldCelda.Offset(0, -3) = Empty Then
                                    .Cells(lmFila, 3) = ldCelda.Offset(0, -3).End(xlUp) 'No. de Partida
                                Else
                                    .Cells(lmFila, 3) = ldCelda.Offset(0, -3) 'No. de Partida
                            End If
                        
                            If ldCelda.Offset(0, -2) = Empty Then
                                    .Cells(lmFila, 4) = Format(ldCelda.Offset(0, -2).End(xlUp), "MM/DD/YYYY") 'Fecha
                                Else
                                    .Cells(lmFila, 4) = Format(ldCelda.Offset(0, -2), "MM/DD/YYYY") 'Fecha
                            End If
                       
                        .Cells(lmFila, 5) = ldCelda.Offset(0, 2) 'DEBE
                        .Cells(lmFila, 5).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        .Cells(lmFila, 6) = ldCelda.Offset(0, 3) 'HABER
                        .Cells(lmFila, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        
                        .Range(.Cells(1, 1), .Cells(1, 6)).HorizontalAlignment = xlCenter
                        .Range(.Cells(1, 1), .Cells(1, 6)).Interior.Color = RGB(190, 190, 90)
                        .Range(.Cells(1, 1), .Cells(1, 6)).Font.Color = RGB(255, 255, 255)
                        .Range(.Cells(1, 1), .Cells(1, 6)).Font.Bold = True
        End With
                        
                        lmFila = lmFila + 1
                    End If
                Next ldCelda
            End If
    Next ccCelda
    
    Call SepararCuentasMayor
    Call SumarDebeMayor
    Call SumarHaberMayor
    Call LimpiarRepetidosMayor

Application.ScreenUpdating = True
    

End Sub

Sub SepararCuentasMayor()
Dim Fila As Long
Dim Final As Long
    
 Hoja43.Activate
 
    Final = nReg(Hoja43, 2, 1) - 2 ' Le resto 2, para que no inserte un encabezado sin datos al final
    
    
    With Hoja43
        For Fila = Final To 2 Step -1
             If .Cells(Fila + 1, 1) <> .Cells(Fila, 1) Then
                Rows(.Cells(Fila + 1, 1).Row & ":" & .Cells(Fila + 1, 1).Row + 1).Insert
        
                .Cells(Fila + 2, 1) = "CUENTA"
                .Cells(Fila + 2, 2) = "NOMBRE DE LA CUENTA"
                .Cells(Fila + 2, 3) = "#"
                .Cells(Fila + 2, 4) = "FECHA"
                .Cells(Fila + 2, 5) = "DEBE"
                .Cells(Fila + 2, 6) = "HABER"
                .Range(.Cells(Fila + 2, 1), .Cells(Fila + 2, 6)).HorizontalAlignment = xlCenter
                .Range(.Cells(Fila + 2, 1), .Cells(Fila + 2, 6)).Interior.Color = RGB(190, 190, 90)
                .Range(.Cells(Fila + 2, 1), .Cells(Fila + 2, 6)).Font.Color = RGB(255, 255, 255)
                .Range(.Cells(Fila + 2, 1), .Cells(Fila + 2, 6)).Font.Bold = True
            End If
        Next
   End With
    
End Sub

Sub SumarDebeMayor()
Dim i As Long
Dim vDebe As Currency
Dim Final As Long

On Error Resume Next

Hoja43.Range("A1").Activate
      
      Do Until IsEmpty(ActiveCell) And IsEmpty(ActiveCell.Offset(1, 0))
         ActiveCell.Offset(2, 0).Select
      Loop
    
        Final = ActiveCell.Row
        
        
       
Hoja43.Range("E2").Activate

    For i = 1 To Evaluate("CountBlank(A1:A" & Final & ")")
    vDebe = 0
            Do Until IsEmpty(ActiveCell.Offset(0, -4))
                vDebe = vDebe + ActiveCell.Value
                ActiveCell.Offset(1, 0).Select
            Loop

            If vDebe <> 0 Then
                ActiveCell.Value = vDebe
                ActiveCell.Borders(xlEdgeTop).Color = RGB(0, 0, 0)
                ActiveCell.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                ActiveCell.Font.Bold = True
                ActiveCell.Offset(1, 0).Select
           
            Else
                ActiveCell.Offset(1, 0).Select
            End If
    Next i
End Sub
Sub SumarHaberMayor()
Dim i As Long
Dim Final As Long
Dim vHaber As Currency

On Error Resume Next

Hoja43.Range("A1").Activate
      
      Do Until IsEmpty(ActiveCell) And IsEmpty(ActiveCell.Offset(1, 0))
         ActiveCell.Offset(2, 0).Select
      Loop
    
        Final = ActiveCell.Row



Hoja43.Range("F2").Activate

    For i = 1 To Evaluate("CountBlank(A1:A" & Final & ")")
    
        vHaber = 0
        
            Do Until IsEmpty(ActiveCell.Offset(0, -5))
                vHaber = vHaber + ActiveCell.Value
                ActiveCell.Offset(1, 0).Select
            Loop
            
        If vHaber <> 0 Then
            ActiveCell.Value = vHaber
            ActiveCell.Borders(xlEdgeTop).Color = RGB(0, 0, 0)
            ActiveCell.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            ActiveCell.Font.Bold = True
            ActiveCell.Offset(1, 0).Select
        Else
            ActiveCell.Offset(1, 0).Select
        End If
        
    Next i
End Sub

Sub LimpiarRepetidosMayor()
Dim Fila As Long
Dim Final As Long
    
Hoja43.Range("A1").Activate
      
      Do Until IsEmpty(ActiveCell) And IsEmpty(ActiveCell.Offset(1, 0))
         ActiveCell.Offset(2, 0).Select
      Loop
    
        Final = ActiveCell.Row
        

        For Fila = Final To 2 Step -1
            If Hoja43.Cells(Fila + 1, 1) = Hoja43.Cells(Fila, 1) Then
                Hoja43.Cells(Fila + 1, 1) = Empty
                Hoja43.Cells(Fila + 1, 2) = Empty
            End If
        Next

Hoja43.Range("A1").Activate

End Sub

Sub ConstruirBalancedeComprobacion()
    Dim ccCelda As Range, ccRango As Range
    Dim ldCelda As Range, ldRango As Range
    Dim bcFila As Long
    
Application.ScreenUpdating = False

    Hoja44.Activate ' Activar la hoja Balance de Comprobación
    
    Cells.Select
    Selection.Clear
 
    
bcFila = 2
    
Hoja41.Activate ' Catálogo de Cuentas
    
    Set ccRango = Hoja41.Range(Cells(2, 1), Cells(2, 1).End(xlDown)) 'Preparando el Rango del Catálogo de Cuentas
    
    For Each ccCelda In ccRango
        With Hoja44
            If Len(ccCelda) = 3 Then
                Hoja42.Activate ' Libro Diario
                Set ldRango = Hoja42.Range(Cells(2, 4), Cells(2, 4).End(xlDown)) 'Preparando el Rango del Libro Diario
                For Each ldCelda In ldRango
                    If ccCelda = Val(Mid(ldCelda.Offset(0, 0), 1, 3)) Then  ' Comparo la CELDA de la Hoja41 Catálogo de Cuentas, con la Hoja42 Libro Diario.
                        .Cells(1, 1) = "CUENTA"
                        .Cells(1, 2) = "NOMBRE DE LA CUENTA"
                        .Cells(1, 3) = "DEBE"
                        .Cells(1, 4) = "HABER"
                        .Cells(1, 5) = "SALDO DEUDOR"
                        .Cells(1, 6) = "SALDO ACREEDOR"
                        
                        'Dando formato al encabezado
                        .Range(.Cells(1, 1), .Cells(1, 6)).HorizontalAlignment = xlCenter
                        .Range(.Cells(1, 1), .Cells(1, 6)).Interior.Color = RGB(100, 190, 190)
                        .Range(.Cells(1, 1), .Cells(1, 6)).Font.Color = RGB(255, 255, 255)
                        .Range(.Cells(1, 1), .Cells(1, 6)).Font.Bold = True
                        
                        .Cells(bcFila, 1) = ccCelda.Offset(0, 0).Value 'Cuenta
                        .Cells(bcFila, 2) = ccCelda.Offset(0, 1).Value 'Nombre de Cuenta
                     
                        .Cells(bcFila, 3) = ldCelda.Offset(0, 2) 'DEBE
                        .Cells(bcFila, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        .Cells(bcFila, 4) = ldCelda.Offset(0, 3) 'HABER
                        .Cells(bcFila, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

                        
                        
                    bcFila = bcFila + 1
                    End If
                Next ldCelda
            End If
        End With
    Next ccCelda
    
Call SepararCuentasComprobacion
Call SumarDebeHaberComprobacion
Call ConsolidarBalanceComprobacion
Call TotalizarBalanceComprobacion

Application.ScreenUpdating = True
    
End Sub

Sub SepararCuentasComprobacion()
Dim Fila As Long
Dim Final As Long

Hoja44.Activate ' Balance de Comprobación
    
    Final = nReg(Hoja44, 2, 1) - 1
    
    With Hoja44
        For Fila = Final To 2 Step -1
             If .Cells(Fila + 1, 1) <> .Cells(Fila, 1) Then
                Rows(.Cells(Fila + 1, 1).Row & ":" & .Cells(Fila + 1, 1).Row).Insert
            End If
        Next
   End With
End Sub


Sub SumarDebeHaberComprobacion()
Dim i As Long
Dim vDebeHaber As Currency
Dim sDeudor As Currency
Dim Final As Long
Dim j As Long

Hoja44.Range("A1").Activate
      
      Do Until IsEmpty(ActiveCell) And IsEmpty(ActiveCell.Offset(1, 0))
         ActiveCell.Offset(2, 0).Select
      Loop
    
        Final = ActiveCell.Row
        

 
For j = 3 To 5

    Hoja44.Cells(2, j).Activate

        For i = 1 To Evaluate("CountBlank(A1:A" & Final & ")")
        vDebeHaber = 0
                Do Until IsEmpty(ActiveCell.Offset(0, 1 - j))
                    vDebeHaber = vDebeHaber + ActiveCell.Value
                    ActiveCell.Font.Bold = True
                    ActiveCell.Offset(1, 0).Select
                Loop


            If j = 5 And ActiveCell.Offset(0, -1) <> Empty Or ActiveCell.Offset(0, -2) <> Empty Then
                                                                            
                ActiveCell.Offset(0, -4) = ActiveCell.Offset(-1, -4)
                ActiveCell.Offset(0, -3) = ActiveCell.Offset(-1, -3)
            End If
                
                
                
    
                If vDebeHaber <> 0 Then
                    ActiveCell.Value = vDebeHaber
                    ActiveCell.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    ActiveCell.Offset(1, 0).Select
               
                Else
                    ActiveCell.Offset(1, 0).Select
                End If
        Next i
Next j

End Sub

Sub ConsolidarBalanceComprobacion()
Dim Fila As Long
Dim Final As Long

    
Hoja44.Range("A1").Activate
      
      Do Until IsEmpty(ActiveCell)
         ActiveCell.Offset(1, 0).Select
      Loop
      

    
        Final = ActiveCell.Row
        
        
           

        For Fila = Final To 2 Step -1
            If Hoja44.Cells(Fila, 3).Font.Bold = True Then
                Hoja44.Cells(Fila, 1).EntireRow.Delete
             End If
        Next

Hoja44.Range("A1").Activate

End Sub

Sub TotalizarBalanceComprobacion()
Dim vTotalMoneda As Currency
Dim i As Long


' Totalizo los valores de moneda
        For i = 3 To 6
        
            Hoja44.Cells(2, i).Activate
        
        vTotalMoneda = 0
        
                    Do Until IsEmpty(ActiveCell.Offset(0, 1 - i))
                    
                    If i = 5 Then ' Genero el saldo deudor
                        If ActiveCell.Offset(0, -2) - ActiveCell.Offset(0, -1).Value > 0 Then
                            ActiveCell.Value = ActiveCell.Offset(0, -2) - ActiveCell.Offset(0, -1).Value
                            ActiveCell.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        End If
                    End If
                    
                    If i = 6 Then ' Genero el saldo acreedor
                        If ActiveCell.Offset(0, -3) - ActiveCell.Offset(0, -2).Value < 0 Then
                            ActiveCell.Value = ActiveCell.Offset(0, -3) - ActiveCell.Offset(0, -2).Value
                            ActiveCell.Value = Replace(ActiveCell.Value, "-", "")
                            ActiveCell.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        End If
                    End If
                    
                        vTotalMoneda = vTotalMoneda + ActiveCell.Value
                        ActiveCell.Offset(1, 0).Select
                    Loop
                    
                    If i = 6 Then 'Trazo una línea en la parte superior de los totales
                        Hoja44.Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(0, -5)).Borders(xlEdgeTop).Color = RGB(190, 190, 190)
                    End If
                        
                        ActiveCell.Value = vTotalMoneda
                        ActiveCell.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        ActiveCell.Font.Bold = True
                        ActiveCell.Offset(1, 0).Select
        
        Next i

Hoja44.Range("A1").Activate
End Sub

Sub BalanceGeneral()
Dim Celda As Range, Rango As Range, rBuscar As Range
Dim Fila As Long, ultFila As Long
Dim totACorriente As Currency, totAnoCorriente As Currency, totActivos As Currency
Dim totPCorriente As Currency, totPnoCorriente As Currency, totPatrimonio As Currency, totPasivos As Currency
Dim sEmpresa As String, FechaBalance As String, TipoMoneda As String
Dim fRepresentante As String, fContador As String, fAuditor As String

Application.ScreenUpdating = False


  
   '/////////////// ENCABEZADO ////////////////////


sEmpresa = InputBox("Nombre de la Empresa: ")
If sEmpresa = Empty Then Exit Sub
FechaBalance = InputBox("Período del Balance: ")
If FechaBalance = Empty Then Exit Sub
TipoMoneda = InputBox("Escriba la moneda expresada: ")
If TipoMoneda = Empty Then Exit Sub

   '/////////////// CAPTURANDO NOMBRES PARA LAS FIRMAS ////////////////////

fRepresentante = InputBox("Firma del Representante Legal o Apoderado: ")
If fRepresentante = Empty Then Exit Sub
fContador = InputBox("Firma del Contador General: ")
If fContador = Empty Then Exit Sub
fAuditor = InputBox("Firma del Auditor Externo: ")
If fAuditor = Empty Then Exit Sub


   '/////////////// ACTIVO ////////////////////

ultFila = 0
Fila = 7
    
With Hoja45
    
Hoja44.Activate ' Balance de Comprobación
    
    Set Rango = Hoja44.Range(Cells(2, 1), Cells(2, 1).End(xlDown)) 'Preparando el Rango del Balance de Comprobación

.Activate
    
    .Cells(1, 1) = sEmpresa
    .Cells(1, 1).Font.Bold = True
    .Cells(1, 1).HorizontalAlignment = xlCenter
    .Range(Cells(1, 1), Cells(1, 7)).Merge
    
    .Cells(2, 1) = FechaBalance
    .Cells(2, 1).Font.Bold = True
    .Cells(2, 1).HorizontalAlignment = xlCenter
    .Range(Cells(2, 1), Cells(2, 7)).Merge
    
    .Cells(3, 1) = TipoMoneda
    .Cells(3, 1).HorizontalAlignment = xlCenter
    .Range(Cells(3, 1), Cells(3, 7)).Merge
        
    .Cells(Fila - 2, 1) = "ACTIVO"
    .Cells(Fila - 2, 1).HorizontalAlignment = xlCenter
    .Range(Cells(Fila - 2, 1), Cells(Fila - 2, 3)).Merge
    .Cells(Fila - 2, 1).Font.Bold = True
    .Cells(Fila - 1, 1) = "Corriente"
    .Cells(Fila - 1, 1).Font.Bold = True
    
    totACorriente = 0
    For Each Celda In Rango
        If Mid(Celda, 1, 2) = 10 Or Mid(Celda, 1, 2) = 11 Then
            .Cells(Fila, 1) = Celda.Offset(0, 1).Value 'Nombre de Cuenta
            .Cells(Fila, 2) = Celda.Offset(0, 4).Value 'Saldo Deudor
            .Cells(Fila, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            totACorriente = totACorriente + Celda.Offset(0, 4).Value
            Fila = Fila + 1
        End If
    Next Celda
    .Cells(Fila - 1, 2).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    
    Fila = Fila + 2
    
    .Cells(Fila - 1, 1) = "No Corriente"
    .Cells(Fila - 1, 1).Font.Bold = True
    totAnoCorriente = 0
    For Each Celda In Rango
        If Mid(Celda, 1, 2) = 12 Then
            .Cells(Fila, 1) = Celda.Offset(0, 1).Value 'Nombre de Cuenta
            .Cells(Fila, 2) = Celda.Offset(0, 4).Value 'Saldo Deudor
            .Cells(Fila, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            totAnoCorriente = totAnoCorriente + Celda.Offset(0, 4).Value
            Fila = Fila + 1
        End If
    Next Celda
    .Cells(Fila - 1, 2).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    
   '/////////////// PASIVO ////////////////////
   
Fila = 7

    .Cells(Fila - 2, 5) = "PASIVO"
    .Cells(Fila - 2, 5).HorizontalAlignment = xlCenter
    .Range(Cells(Fila - 2, 5), Cells(Fila - 2, 7)).Merge
    .Cells(Fila - 2, 5).Font.Bold = True
    .Cells(Fila - 1, 5) = "Corriente"
    .Cells.Cells(Fila - 1, 5).Font.Bold = True
    totPCorriente = 0
    For Each Celda In Rango
        If Mid(Celda, 1, 2) = 20 Then
            .Cells(Fila, 5) = Celda.Offset(0, 1).Value 'Nombre de Cuenta
            .Cells(Fila, 6) = Celda.Offset(0, 5).Value 'Saldo Acreedor
            .Cells(Fila, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            totPCorriente = totPCorriente + Celda.Offset(0, 5).Value
            Fila = Fila + 1
        End If
    Next Celda
    .Cells(Fila - 1, 6).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    
    Fila = Fila + 2
    
    .Cells(Fila - 1, 5) = "No Corriente"
    .Cells(Fila - 1, 5).Font.Bold = True
    totPnoCorriente = 0
    For Each Celda In Rango
        If Mid(Celda, 1, 2) = 21 Then
            .Cells(Fila, 5) = Celda.Offset(0, 1).Value 'Nombre de Cuenta
            .Cells(Fila, 6) = Celda.Offset(0, 5).Value 'Saldo Acreedor
            .Cells(Fila, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            totPnoCorriente = totPnoCorriente + Celda.Offset(0, 5).Value
            Fila = Fila + 1
        End If
    Next Celda
    .Cells(Fila - 1, 6).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    
    Fila = Fila + 2
    
    .Cells(Fila - 1, 5) = "Patrimonio"
    .Cells(Fila - 1, 5).Font.Bold = True
    totPatrimonio = 0
    For Each Celda In Rango
        If Mid(Celda, 1, 2) = 30 Or Mid(Celda, 1, 2) = 40 Then
            .Cells(Fila, 5) = Celda.Offset(0, 1).Value 'Nombre de Cuenta
            .Cells(Fila, 6) = Celda.Offset(0, 5).Value 'Saldo Acreedor
            .Cells(Fila, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            totPatrimonio = totPatrimonio + Celda.Offset(0, 5).Value
            Fila = Fila + 1
        End If
    Next Celda
    .Cells(Fila - 1, 6).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    
   .Activate
   
   .Range("A1").Select
   
   '/////////////// TOTALES ////////////////////

totActivos = totACorriente + totAnoCorriente
totPasivos = totPCorriente + totPnoCorriente + totPatrimonio

Set rBuscar = .Range("A:A").Find("Corriente", LookIn:=xlValues)
    rBuscar.Offset(0, 2).Value = totACorriente
    rBuscar.Offset(0, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    rBuscar.Offset(0, 2).Font.Bold = True
Set rBuscar = .Range("A:A").Find("No Corriente", LookIn:=xlValues)
    rBuscar.Offset(0, 2).Value = totAnoCorriente
    rBuscar.Offset(0, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    rBuscar.Offset(0, 2).Font.Bold = True

Set rBuscar = .Range("E:E").Find("Corriente", LookIn:=xlValues)
    rBuscar.Offset(0, 2).Value = totPCorriente
    rBuscar.Offset(0, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    rBuscar.Offset(0, 2).Font.Bold = True
Set rBuscar = .Range("E:E").Find("No Corriente", LookIn:=xlValues)
    rBuscar.Offset(0, 2).Value = totPnoCorriente
    rBuscar.Offset(0, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    rBuscar.Offset(0, 2).Font.Bold = True
Set rBuscar = .Range("E:E").Find("Patrimonio", LookIn:=xlValues)
    rBuscar.Offset(0, 2).Value = totPatrimonio
    rBuscar.Offset(0, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    rBuscar.Offset(0, 2).Font.Bold = True

'ThisWorkbook.Save
ultFila = .Cells.SpecialCells(xlCellTypeLastCell).Row + 3

.Cells(ultFila, 1) = "Total Activos:"
.Cells(ultFila, 1).Font.Bold = True
.Cells(ultFila, 3) = totActivos
.Cells(ultFila, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
.Cells(ultFila, 3).Font.Bold = True
.Cells(ultFila, 3).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
.Cells(ultFila, 3).Borders(xlEdgeBottom).LineStyle = xlDouble

.Cells(ultFila, 5) = "Total Pasivo y Patrimonio:"
.Cells(ultFila, 5).Font.Bold = True
.Cells(ultFila, 7) = totPasivos
.Cells(ultFila, 7).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
.Cells(ultFila, 7).Font.Bold = True
.Cells(ultFila, 7).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
.Cells(ultFila, 7).Borders(xlEdgeBottom).LineStyle = xlDouble


   '/////////////// FIRMAS ////////////////////

ultFila = ultFila + 5

.Cells(ultFila, 1) = fRepresentante
.Cells(ultFila, 1).HorizontalAlignment = xlCenter
.Cells(ultFila, 3) = fContador
.Cells(ultFila, 3).HorizontalAlignment = xlCenter
.Range(Cells(ultFila, 3), Cells(ultFila, 4)).Merge
.Cells(ultFila, 6) = fAuditor
.Cells(ultFila, 6).HorizontalAlignment = xlCenter
.Range(Cells(ultFila, 6), Cells(ultFila, 7)).Merge

ultFila = ultFila + 1

.Cells(ultFila, 1) = "REPRESENTANTE LEGAL O APODERADO"
.Cells(ultFila, 1).HorizontalAlignment = xlCenter
.Cells(ultFila - 1, 1).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
.Cells(ultFila, 3) = "CONTADOR GENERAL"
.Cells(ultFila, 3).HorizontalAlignment = xlCenter
.Range(Cells(ultFila, 3), Cells(ultFila, 4)).Merge
.Range(Cells(ultFila - 1, 3), Cells(ultFila, 4)).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
.Cells(ultFila, 6) = "AUDITOR EXTERNO, REG.#"
.Cells(ultFila, 6).HorizontalAlignment = xlCenter
.Range(Cells(ultFila, 6), Cells(ultFila, 7)).Merge
.Range(Cells(ultFila - 1, 6), Cells(ultFila, 7)).Borders(xlEdgeTop).Color = RGB(0, 0, 0)

.btn_BalanceGeneral.Caption = "Limpiar"

End With

Application.ScreenUpdating = True

End Sub








