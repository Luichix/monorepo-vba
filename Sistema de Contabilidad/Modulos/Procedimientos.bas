Attribute VB_Name = "Procedimientos"
Option Explicit
Public banderaListadoCuentas As Long
Public nGrupo As Long

Sub ValidarCuenta()
Dim Fila As Long
Dim Final As Long
Dim encontrado As Boolean

With frm_CatalogoCuentas
    Final = nReg(Hoja2, 2, 1) - 1

    For Fila = 2 To Final
        If Hoja2.Cells(Fila, 1) = Val(Mid(.cbo_CodCuenta, 1, 1)) _
        Or .cbo_CodCuenta = Empty Then
            encontrado = True
            nGrupo = Hoja2.Cells(Fila, 1)
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

Sub Run_CatalogoCuentas()
    Load frm_CatalogoCuentas
    frm_CatalogoCuentas.Show
End Sub
Sub Run_LibroDiario()
    Load frm_LibroDiario
    frm_LibroDiario.Show
End Sub

Sub CodCuentaATexto()
Dim Celda As Object
Dim miRangoDinamico As String
Dim Rango As Range
Dim Final As Long

Final = nReg(Hoja2, 2, 1) - 1

miRangoDinamico = "A" & 2 & ":" & "A" & Final

Hoja2.Range(miRangoDinamico).NumberFormat = "@"

Set Rango = Hoja2.Range(miRangoDinamico)
        
    For Each Celda In Rango
        Celda.Value = CStr(Celda)
    Next Celda
End Sub

Sub CodCuentaANumero()
Dim Celda As Object
Dim miRangoDinamico As String
Dim Rango As Range
Dim Final As Long

Final = nReg(Hoja2, 2, 1) - 1

miRangoDinamico = "A" & 2 & ":" & "A" & Final

Hoja2.Range(miRangoDinamico).NumberFormat = "General"

Set Rango = Hoja2.Range(miRangoDinamico)
        
    For Each Celda In Rango
        Celda.Value = Val(Celda)
    Next Celda
    
End Sub

Sub IndexarCodCuentasPLAN()
    Call CodCuentaATexto
    
    Hoja2.Range("A:C").Sort key1:=Hoja2.Range("A2"), _
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
    
        If frm_ListadoCuentas.lbx_Cuentas.Column(0) < 1000 Then
            MsgBox "Debe seleccionar una subcuenta..!", vbInformation
            frm_ListadoCuentas.lbx_Cuentas.SetFocus
            Exit Sub
        End If

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
 
 Hoja4.Activate ' Libro Mayor
 
    Cells.Select
    Selection.Clear
 

lmFila = 2
    
  
Hoja2.Activate ' Catálogo de Cuentas
    
    Set ccRango = Hoja2.Range(Cells(2, 1), Cells(2, 1).End(xlDown)) 'Preparando el Rango del Catálogo de Cuentas
    
    For Each ccCelda In ccRango ' Checando cada celda en el catálogo de cuentas Hoja2
        

            
            If Len(ccCelda) = 5 Then
            
                Hoja3.Activate ' Libro Diario
                
                Set ldRango = Hoja3.Range(Cells(2, 5), Cells(2, 5).End(xlDown)) 'Preparando el Rango del Libro Diario
                For Each ldCelda In ldRango
                    If ccCelda = Val(Mid(ldCelda.Offset(0, 0), 1, 5)) Then  ' Comparo la CELDA de la hoja2 Catálogo de Cuentas, con la Hoja3 Libro Diario.
                       ' y escribo los datos en la hoj4 Libro Mayor
        With Hoja4
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
    
 Hoja4.Activate
 
    Final = nReg(Hoja4, 2, 1) - 2 ' Le resto 2, para que no inserte un encabezado sin datos al final
    
    
    With Hoja4
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

Hoja4.Range("A1").Activate
      
      Do Until IsEmpty(ActiveCell) And IsEmpty(ActiveCell.Offset(1, 0))
         ActiveCell.Offset(2, 0).Select
      Loop
    
        Final = ActiveCell.Row
        
        
       
Hoja4.Range("E2").Activate

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

Hoja4.Range("A1").Activate
      
      Do Until IsEmpty(ActiveCell) And IsEmpty(ActiveCell.Offset(1, 0))
         ActiveCell.Offset(2, 0).Select
      Loop
    
        Final = ActiveCell.Row



Hoja4.Range("F2").Activate

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
    
Hoja4.Range("A1").Activate
      
      Do Until IsEmpty(ActiveCell) And IsEmpty(ActiveCell.Offset(1, 0))
         ActiveCell.Offset(2, 0).Select
      Loop
    
        Final = ActiveCell.Row
        

        For Fila = Final To 2 Step -1
            If Hoja4.Cells(Fila + 1, 1) = Hoja4.Cells(Fila, 1) Then
                Hoja4.Cells(Fila + 1, 1) = Empty
                Hoja4.Cells(Fila + 1, 2) = Empty
            End If
        Next

Hoja4.Range("A1").Activate

End Sub

Sub ConstruirBalancedeComprobacion()
    Dim ccCelda As Range, ccRango As Range
    Dim ldCelda As Range, ldRango As Range
    Dim bcFila As Long
    
Application.ScreenUpdating = False

    Hoja5.Activate ' Activar la hoja Balance de Comprobación
    
    Cells.Select
    Selection.Clear
 
    
bcFila = 2
    
Hoja2.Activate ' Catálogo de Cuentas
    
    Set ccRango = Hoja2.Range(Cells(2, 1), Cells(2, 1).End(xlDown)) 'Preparando el Rango del Catálogo de Cuentas
    
    For Each ccCelda In ccRango
        With Hoja5
            If Len(ccCelda) = 5 Then
                Hoja3.Activate ' Libro Diario
                Set ldRango = Hoja3.Range(Cells(2, 5), Cells(2, 5).End(xlDown)) 'Preparando el Rango del Libro Diario
                For Each ldCelda In ldRango
                    If ccCelda = Val(Mid(ldCelda.Offset(0, 0), 1, 5)) Then  ' Comparo la CELDA de la hoja2 Catálogo de Cuentas, con la Hoja3 Libro Diario.
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

Hoja5.Activate ' Balance de Comprobación
    
    Final = nReg(Hoja5, 2, 1) - 1
    
    With Hoja5
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

Hoja5.Range("A1").Activate
      
      Do Until IsEmpty(ActiveCell) And IsEmpty(ActiveCell.Offset(1, 0))
         ActiveCell.Offset(2, 0).Select
      Loop
    
        Final = ActiveCell.Row
        

 
For j = 3 To 5

    Hoja5.Cells(2, j).Activate

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

    
Hoja5.Range("A1").Activate
      
      Do Until IsEmpty(ActiveCell)
         ActiveCell.Offset(1, 0).Select
      Loop
      

    
        Final = ActiveCell.Row
        
        
           

        For Fila = Final To 2 Step -1
            If Hoja5.Cells(Fila, 3).Font.Bold = True Then
                Hoja5.Cells(Fila, 1).EntireRow.Delete
             End If
        Next

Hoja5.Range("A1").Activate

End Sub

Sub TotalizarBalanceComprobacion()
Dim vTotalMoneda As Currency
Dim i As Long


' Totalizo los valores de moneda
        For i = 3 To 6
        
            Hoja5.Cells(2, i).Activate
        
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
                        Hoja5.Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(0, -5)).Borders(xlEdgeTop).Color = RGB(190, 190, 190)
                    End If
                        
                        ActiveCell.Value = vTotalMoneda
                        ActiveCell.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        ActiveCell.Font.Bold = True
                        ActiveCell.Offset(1, 0).Select
        
        Next i

Hoja5.Range("A1").Activate
End Sub

Sub BalanceGeneral()
Dim Celda As Range, Rango As Range, rBuscar As Range
Dim Fila As Long, ultFila As Long
Dim xFinal As Long, zFinal As Long
Dim totACorriente As Currency, totAnoCorriente As Currency, totActivos As Currency
Dim totPCorriente As Currency, totPnoCorriente As Currency, totPatrimonio As Currency, totPasivos As Currency
Dim sEmpresa As String, FechaBalance As String, TipoMoneda As String
Dim fRepresentante As String, fContador As String, fAuditor As String

Application.ScreenUpdating = False


  
   '/////////////// ENCABEZADO ////////////////////

 Hoja6.Activate ' Libro Mayor
 
    Cells.Select
    Selection.Clear
    
sEmpresa = Hoja91.Range("H4").Text

FechaBalance = "BALANCE GENERAL"

TipoMoneda = "ELABORADO EN " & UCase(Format(Date, "MMMM")) & " " & VBA.Year(Date)




   '/////////////// CAPTURANDO NOMBRES PARA LAS FIRMAS ////////////////////

fRepresentante = "Elaborado por"
fContador = "Revisado por"
fAuditor = "Autorizado por"



   '/////////////// ACTIVO ////////////////////

ultFila = 0
Fila = 7
    
With Hoja6
    
Hoja5.Activate ' Balance de Comprobación
    
    Set Rango = Hoja5.Range(Cells(2, 1), Cells(2, 1).End(xlDown)) 'Preparando el Rango del Balance de Comprobación

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
    .Cells(3, 1).Font.Bold = True
    .Cells(3, 1).HorizontalAlignment = xlCenter
    .Range(Cells(3, 1), Cells(3, 7)).Merge
        
    .Cells(Fila - 2, 1) = "ACTIVO"
    .Cells(Fila - 2, 1).HorizontalAlignment = xlCenter
    .Range(Cells(Fila - 2, 1), Cells(Fila - 2, 3)).Merge
    .Cells(Fila - 2, 1).Font.Bold = True
    .Cells(Fila - 1, 1) = "Activo Circulante"
    .Cells(Fila - 1, 1).Font.Bold = True
    
    totACorriente = 0
    For Each Celda In Rango
        If Mid(Celda, 1, 3) = 101 Then ' Or Mid(Celda, 1, 2) = 11
            .Cells(Fila, 1) = Celda.Offset(0, 1).Value 'Nombre de Cuenta
            .Cells(Fila, 2) = Celda.Offset(0, 4).Value 'Saldo Deudor
            .Cells(Fila, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            totACorriente = totACorriente + Celda.Offset(0, 4).Value
            Fila = Fila + 1
        End If
    Next Celda
    .Cells(Fila - 1, 2).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    
    Fila = Fila + 2
    
    .Cells(Fila - 1, 1) = "Activo Fijo"
    .Cells(Fila - 1, 1).Font.Bold = True
    totAnoCorriente = 0
    For Each Celda In Rango
        If Mid(Celda, 1, 3) = 102 Then
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
    .Cells(Fila - 1, 5) = "Pasivo Circulante"
    .Cells.Cells(Fila - 1, 5).Font.Bold = True
    totPCorriente = 0
    For Each Celda In Rango
        If Mid(Celda, 1, 3) = 201 Then
            .Cells(Fila, 5) = Celda.Offset(0, 1).Value 'Nombre de Cuenta
            .Cells(Fila, 6) = Celda.Offset(0, 5).Value 'Saldo Acreedor
            .Cells(Fila, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            totPCorriente = totPCorriente + Celda.Offset(0, 5).Value
            Fila = Fila + 1
        End If
    Next Celda
    .Cells(Fila - 1, 6).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    
    Fila = Fila + 2
    
    .Cells(Fila - 1, 5) = "Pasivo Fijo"
    .Cells(Fila - 1, 5).Font.Bold = True
    totPnoCorriente = 0
    For Each Celda In Rango
        If Mid(Celda, 1, 3) = 202 Then
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
        If Mid(Celda, 1, 2) = 30 Then  'Or Mid(Celda, 1, 2) = 40
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

Set rBuscar = .Range("A:A").Find("Activo Circulante", LookIn:=xlValues)
    rBuscar.Offset(0, 2).Value = totACorriente
    rBuscar.Offset(0, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    rBuscar.Offset(0, 2).Font.Bold = True
Set rBuscar = .Range("A:A").Find("Activo Fijo", LookIn:=xlValues)
    rBuscar.Offset(0, 2).Value = totAnoCorriente
    rBuscar.Offset(0, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    rBuscar.Offset(0, 2).Font.Bold = True

Set rBuscar = .Range("E:E").Find("Pasivo Circulante", LookIn:=xlValues)
    rBuscar.Offset(0, 2).Value = totPCorriente
    rBuscar.Offset(0, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    rBuscar.Offset(0, 2).Font.Bold = True
Set rBuscar = .Range("E:E").Find("Pasivo Fijo", LookIn:=xlValues)
    rBuscar.Offset(0, 2).Value = totPnoCorriente
    rBuscar.Offset(0, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    rBuscar.Offset(0, 2).Font.Bold = True
Set rBuscar = .Range("E:E").Find("Patrimonio", LookIn:=xlValues)
    rBuscar.Offset(0, 2).Value = totPatrimonio
    rBuscar.Offset(0, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    rBuscar.Offset(0, 2).Font.Bold = True

'ThisWorkbook.Save



xFinal = Hoja6.Range("A" & Rows.Count).End(xlUp).Row
zFinal = Hoja6.Range("E" & Rows.Count).End(xlUp).Row

If xFinal > zFinal Then
    ultFila = xFinal + 3
ElseIf zFinal > xFinal Then
    ultFila = zFinal + 3
Else
    ultFila = xFinal + 3
End If



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


.Cells(ultFila - 1, 1).Borders(xlEdgeTop).Color = RGB(0, 0, 0)

.Range(Cells(ultFila, 3), Cells(ultFila, 4)).Merge
.Range(Cells(ultFila - 1, 3), Cells(ultFila, 4)).Borders(xlEdgeTop).Color = RGB(0, 0, 0)

.Range(Cells(ultFila, 6), Cells(ultFila, 7)).Merge
.Range(Cells(ultFila - 1, 6), Cells(ultFila, 7)).Borders(xlEdgeTop).Color = RGB(0, 0, 0)



End With

Application.ScreenUpdating = True

End Sub

Sub Estado_Resultado()

Application.ScreenUpdating = False

Dim Celda As Range, Rango As Range, rBuscar As Range
Dim Empresa As String, Estado As String, Periodo As String
Dim Fila As Long, Final As Long
Dim xValor As Currency, xCosto As Currency, xGasto As Currency
Dim xBruta As Currency, xIngreso As Currency, xAnteImpuesto As Currency
Dim xImpuesto As Currency, xUtilidad As Currency
Dim Responsable As String


 Hoja7.Activate
    Cells.Select
    Selection.Clear

Empresa = Hoja91.Range("H4").Text
Estado = "ESTADO DE RESULTADO"
Periodo = "ESTABLECER PERIODO"

xValor = 0

With Hoja7
    
    .Cells(1, 1) = Empresa
    .Cells(2, 1) = Estado
    .Cells(3, 1) = Periodo
    
    .Cells(4, 2) = 1
    .Cells(4, 3) = 2
    
    .Cells(5, 1) = "INGRESOS POR VENTAS"
    .Cells(6, 1) = "COSTOS DE PRODUCCIÓN"
    .Cells(7, 1) = "MATERIALES DIRECTOS"
    .Cells(8, 1) = "MANO DE OBRA"
    .Cells(9, 1) = "COSTOS INDIRECTOS DE FABRICACIÓN"
    .Cells(10, 1) = "UTILIDAD BRUTA"
    .Cells(11, 1) = "GASTOS OPERATIVOS"
    .Cells(12, 1) = "GASTOS DE ADMINISTRACIÓN"
    .Cells(13, 1) = "GASTOS DE VENTA"
    .Cells(14, 1) = "GASTOS FINANCIEROS"
    .Cells(15, 1) = "UTILIDAD ANTES DE IMPUESTOS"
    .Cells(16, 1) = "IMPUESTOS DE OPERACIÓN"
    .Cells(17, 1) = "UTILIDAD NETA"
    
    For Fila = 7 To 9
        .Cells(Fila, 1).InsertIndent 2
    Next
    For Fila = 12 To 14
        .Cells(Fila, 1).InsertIndent 2
    Next
        .Cells(16, 1).InsertIndent 2
 
    Hoja5.Activate
    
    Set Rango = Hoja5.Range(Cells(2, 1), Cells(2, 1).End(xlDown))

.Activate

    For Each Celda In Rango
        If Mid(Celda, 1, 3) = 401 Then
            .Cells(5, 3) = Celda.Offset(0, 5).Value 'Saldo Deudor
            .Cells(5, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End If
    Next Celda
    
    If .Cells(5, 3) = Empty Then
        .Cells(5, 3) = xValor
        .Cells(5, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    End If
    
    
    For Each Celda In Rango
        If Mid(Celda, 1, 5) = 60101 Then
            .Cells(7, 2) = Celda.Offset(0, 4).Value 'Saldo Acreedor
            .Cells(7, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End If
    Next Celda
    
    If .Cells(7, 2) = Empty Then
        .Cells(7, 2) = xValor
        .Cells(7, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    End If
    
        For Each Celda In Rango
        If Mid(Celda, 1, 5) = 60102 Then
            .Cells(8, 2) = Celda.Offset(0, 4).Value 'Saldo Acreedor
            .Cells(8, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End If
    Next Celda
    
    If .Cells(8, 2) = Empty Then
        .Cells(8, 2) = xValor
        .Cells(8, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    End If
    
        For Each Celda In Rango
        If Mid(Celda, 1, 5) = 60103 Then
            .Cells(9, 2) = Celda.Offset(0, 4).Value 'Saldo Acreedor
            .Cells(9, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End If
    Next Celda
    
    If .Cells(9, 2) = Empty Then
        .Cells(9, 2) = xValor
        .Cells(9, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    End If
    
    xCosto = .Cells(7, 2) + .Cells(8, 2) + .Cells(9, 2)
    
    .Cells(9, 3) = xCosto
    .Cells(9, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    xIngreso = .Cells(5, 3)
    xBruta = xIngreso - xCosto
    
    .Cells(10, 3) = xBruta
    .Cells(10, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    
    For Each Celda In Rango
        If Mid(Celda, 1, 5) = 50101 Then
            .Cells(12, 2) = Celda.Offset(0, 4).Value 'Saldo Acreedor
            .Cells(12, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End If
    Next Celda
    
    If .Cells(12, 2) = Empty Then
        .Cells(12, 2) = xValor
        .Cells(12, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    End If
    
    For Each Celda In Rango
        If Mid(Celda, 1, 5) = 50102 Then
            .Cells(13, 2) = Celda.Offset(0, 4).Value 'Saldo Acreedor
            .Cells(13, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End If
    Next Celda
    
    If .Cells(13, 2) = Empty Then
        .Cells(13, 2) = xValor
        .Cells(13, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    End If
    
    For Each Celda In Rango
        If Mid(Celda, 1, 5) = 50103 Then
            .Cells(14, 2) = Celda.Offset(0, 4).Value 'Saldo Acreedor
            .Cells(14, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End If
    Next Celda
    
    If .Cells(14, 2) = Empty Then
        .Cells(14, 2) = xValor
        .Cells(14, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    End If
    
    xGasto = .Cells(12, 2) + .Cells(13, 2) + .Cells(14, 2)
    
    .Cells(14, 3) = xGasto
    .Cells(14, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    xAnteImpuesto = xBruta - xGasto
    
    .Cells(15, 3) = xAnteImpuesto
    .Cells(15, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    For Each Celda In Rango
        If Mid(Celda, 1, 5) = 20104 Then
            .Cells(16, 3) = Celda.Offset(0, 5).Value 'Saldo Acreedor
            .Cells(16, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End If
    Next Celda
    
    If .Cells(16, 3) = Empty Then
        .Cells(16, 3) = xValor
        .Cells(16, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    End If
    
    xImpuesto = .Cells(16, 3)
    
    xUtilidad = xAnteImpuesto - xImpuesto
    
    .Cells(17, 3) = xUtilidad
    .Cells(17, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    
End With
    
    Hoja7.Range(Cells(1, 1), Cells(1, 3)).Select
        Selection.Merge
         Selection.Font.Bold = True
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    
    Hoja7.Range(Cells(2, 1), Cells(2, 3)).Select
        Selection.Merge
         Selection.Font.Bold = True
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    
    Hoja7.Range(Cells(3, 1), Cells(3, 3)).Select
        Selection.Merge
         Selection.Font.Bold = True
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With

    Hoja7.Cells(6, 1).Select
     
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    Hoja7.Range(Cells(9, 2), Cells(9, 3)).Select
     
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    Hoja7.Range(Cells(14, 2), Cells(14, 3)).Select
     
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    
    Hoja7.Cells(11, 1).Select
     
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    Hoja7.Cells(16, 3).Select
     
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    Hoja7.Cells(17, 3).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With

Hoja7.Cells(21, 1) = "Elaborado por:  "
Hoja7.Cells(22, 1) = "Revisado por:  "
Hoja7.Cells(23, 1) = "Autorizado por:  "

    Hoja7.Range(Cells(21, 1), Cells(23, 1)).Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
    End With
    
    For Final = 21 To 23
    
    Hoja7.Range(Cells(Final, 2), Cells(Final, 3)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Next
    
    Hoja7.Cells(1, 1).Select
    
    Application.ScreenUpdating = True

End Sub






