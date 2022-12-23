Attribute VB_Name = "Módulo2"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort. _
        SortFields.Add Key:=Range("IDPERSONAL[CODIGO DE EMPLEADO]"), SortOn:= _
        xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort. _
        SortFields.Add Key:=Range("IDPERSONAL[CODIGO DE EMPLEADO]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Sheets("ID PERSONAL").Select
    ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort. _
        SortFields.Add Key:=Range("IDPERSONAL[CODIGO DE EMPLEADO]"), SortOn:= _
        xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort. _
        SortFields.Add Key:=Range("IDPERSONAL[CODIGO DE EMPLEADO]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ID PERSONAL").ListObjects("IDPERSONAL").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
