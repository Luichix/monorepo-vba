VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Cuota_Abono 
   Caption         =   "GESTOR DE RECURSOS HUMANOS"
   ClientHeight    =   8490.001
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   18180
   OleObjectBlob   =   "frm_Cuota_Abono.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Cuota_Abono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim i As Long
Dim j As Long
Dim a As Long
Dim b As Long

Private Sub btn_Cargar_Click()
Dim Seguridad As String
On Error GoTo Salir

Seguridad = Hoja83.Range("L1").Text
    If Me.txt_Fecha.Text = Empty Then
            MsgBox "No se ha ingresado la fecha de deposito...!", vbInformation, "Gestor de Recursos Humanos"
            Exit Sub
    End If
    If Me.txtTotal.Text = Empty Then
            MsgBox "No se ha cargado ninguna cuenta a depositar", vbInformation, "Gestor de Recursos Humanos"
            Exit Sub
    ElseIf Me.txtTotal.Text < 0 Then
            MsgBox "Reportar problema a un usuario administrativo", vbInformation, "Gestor de Recursos Humanos"
            Exit Sub
    
    End If


Hoja12.Unprotect (Seguridad)
        Procesar_Abono
Hoja12.Protect (Seguridad)

        MsgBox "Registro grabado con éxito!!!", , "Gestor de Recursos Humanos"
        Unload Me


                    
     
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Ventas"
 End If


End Sub



Private Sub btn_Fecha_Click()
banderaPeriodo = 5
    Call LanzarPeriodo(Me, "txt_Fecha")
End Sub

Private Sub btn_Limpiar_Click()
Me.ListBox1.Clear
sumarImporte
Me.txt_Fecha.Text = Empty
End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub

Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.ListBox1.ListIndex = -1
SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub Image2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub Image2_Click()
On Error Resume Next
Me.EliminarItem
End Sub
Private Sub Image1_Click()
frm_ListadoAbono.Show
End Sub
Public Sub EliminarItem()
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "Seleccionar un producto para eliminar", vbInformation
        Exit Sub
    End If

Me.ListBox1.RemoveItem (ListBox1.ListIndex)
Me.ListBox1.ListIndex = -1

Me.sumarImporte
            
End Sub
Public Sub sumarImporte()
Dim sTotal As Currency


sTotal = 0
        For i = 0 To Me.ListBox1.ListCount - 1
            
            sTotal = sTotal + Val(Me.ListBox1.List(i, 6))

        Next

Me.txtTotal.Text = Format(CDbl(sTotal), "0.00")

End Sub

Public Sub Cargar_CuotaAbono()
On Error Resume Next
Dim encontrado As Boolean
Dim busqueda As Boolean
Dim Referencia As String
Dim Criterio As String
Dim Conteo As String
Dim Fila As Long
Dim Final As Long
Dim Concepto As String
Dim Abono As Currency
Dim Duplicado As String
Dim Seguridad As String


Seguridad = Hoja83.Range("L1").Text

Titulo = "Gestor de Recursos Humanos"
Referencia = frm_ListadoAbono.txt_referencia.Text


With frm_Cuota_Abono

Hoja12.Unprotect (Seguridad)
Hoja12.Select
Limpiar_Filtro
Orden_Filtro
Range("I1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like Referencia Then
            encontrado = True
            Criterio = ActiveCell.Offset(0, 2).Value
                             Hoja12.Select
                            Range("J1").Select
                            
                                Do Until IsEmpty(ActiveCell)
                                    ActiveCell.Offset(1, 0).Select
                                    If ActiveCell.Value Like Criterio Then
                                        busqueda = True
                                        Conteo = ActiveCell.Offset(0, 0).Value
                                        Concepto = ActiveCell.Offset(0, -5).Value
                                        Abono = ActiveCell.Offset(0, -4).Value
                                        Exit Do
                                    End If
                                Loop
            
            Exit Do
        End If
    Loop
            
            

            
                 For j = 0 To frm_Cuota_Abono.ListBox1.ListCount - 1
                        Duplicado = Val(frm_Cuota_Abono.ListBox1.List(j, 7))
                    If Duplicado Like Conteo Then
                        MsgBox "Este abono a sido cargado anteriormente..!", vbInformation, Titulo

                        Exit Sub
                    End If

                 Next j
                 
                                         Me.ListBox1.AddItem
                        Me.ListBox1.List(i, 0) = frm_ListadoAbono.txt_idpersonal.Text
                        Me.ListBox1.List(i, 1) = frm_ListadoAbono.txt_nombre.Text
                        Me.ListBox1.List(i, 2) = frm_ListadoAbono.txt_referencia.Text
                        Me.ListBox1.List(i, 3) = Concepto
                        Me.ListBox1.List(i, 4) = frm_ListadoAbono.lbx_ListadoAbono.Column(6)
                        Me.ListBox1.List(i, 5) = frm_ListadoAbono.lbx_ListadoAbono.Column(8)
                        Me.ListBox1.List(i, 5) = Replace(Me.ListBox1.List(i, 5), ",", ".")
                        Me.ListBox1.List(i, 5) = Format(Me.ListBox1.List(i, 5), "#,##0.00")
                        Me.ListBox1.List(i, 6) = Abono
                        Me.ListBox1.List(i, 6) = Replace(Me.ListBox1.List(i, 6), ",", ".")
                        Me.ListBox1.List(i, 6) = Format(Me.ListBox1.List(i, 6), "#,##0.00")
                        Me.ListBox1.List(i, 7) = Conteo
                
                            i = i + 1
                            

        End With
        

        sumarImporte
    
Hoja12.Protect (Seguridad)

        Unload frm_ListadoAbono

End Sub
Public Sub Cargar_Todo()
On Error Resume Next
Dim encontrado As Boolean
Dim busqueda As Boolean
Dim Referencia As String
Dim Criterio As String
Dim Conteo As String
Dim Fila As Long
Dim Final As Long
Dim Concepto As String
Dim Abono As Currency
Dim Duplicado As String
Dim Seguridad As String

Titulo = "Gestor de Recursos Humanos"

With frm_Cuota_Abono

    For a = 0 To frm_ListadoAbono.lbx_ListadoAbono.ListCount
    
    If frm_ListadoAbono.lbx_ListadoAbono.List(a, 8) = 0 Then
    
    Else

    Referencia = frm_ListadoAbono.lbx_ListadoAbono.List(a, 9)

Seguridad = Hoja83.Range("L1").Text
Hoja12.Unprotect (Seguridad)

Hoja12.Select
Limpiar_Filtro
Orden_Filtro

Range("I1").Select

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like Referencia Then
            encontrado = True
            Criterio = ActiveCell.Offset(0, 2).Value
            
                             Hoja12.Select
                            Range("J1").Select

                                Do Until IsEmpty(ActiveCell)
                                    ActiveCell.Offset(1, 0).Select
                                    If ActiveCell.Value Like Criterio Then
                                        busqueda = True
                                        Conteo = ActiveCell.Offset(0, 0).Value
                                        Concepto = ActiveCell.Offset(0, -5).Value
                                        Abono = ActiveCell.Offset(0, -4).Value
                                        Exit Do
                                    End If
                                Loop

            Exit Do
        End If
    Loop

                 For j = 0 To frm_Cuota_Abono.ListBox1.ListCount - 1
                        Duplicado = Val(frm_Cuota_Abono.ListBox1.List(j, 7))
                    If Duplicado Like Conteo Then
                            sumarImporte
                            Unload frm_ListadoAbono
                        Exit Sub
                    End If

                 Next j

                        Me.ListBox1.AddItem
                        Me.ListBox1.List(i, 0) = frm_ListadoAbono.lbx_ListadoAbono.List(a, 0)
                        Me.ListBox1.List(i, 1) = frm_ListadoAbono.lbx_ListadoAbono.List(a, 1)
                        Me.ListBox1.List(i, 2) = Referencia
                        Me.ListBox1.List(i, 3) = Concepto
                        Me.ListBox1.List(i, 4) = frm_ListadoAbono.lbx_ListadoAbono.List(a, 6)
                        Me.ListBox1.List(i, 5) = frm_ListadoAbono.lbx_ListadoAbono.List(a, 8)
                        Me.ListBox1.List(i, 5) = Replace(Me.ListBox1.List(i, 5), ",", ".")
                        Me.ListBox1.List(i, 5) = Format(Me.ListBox1.List(i, 5), "#,##0.00")
                        Me.ListBox1.List(i, 6) = Abono
                        Me.ListBox1.List(i, 6) = Replace(Me.ListBox1.List(i, 6), ",", ".")
                        Me.ListBox1.List(i, 6) = Format(Me.ListBox1.List(i, 6), "#,##0.00")
                        Me.ListBox1.List(i, 7) = Conteo

                            i = i + 1

        End If
    Next a

End With

        sumarImporte

Hoja12.Protect (Seguridad)

        Unload frm_ListadoAbono

End Sub
Private Sub Procesar_Abono()
On Error Resume Next
Dim Referencia As String
Dim Registrado As String

Registrado = "REGISTRADO"

                For b = 0 To frm_Cuota_Abono.ListBox1.ListCount - 1
                
                Referencia = frm_Cuota_Abono.ListBox1.List(b, 7)
                        
                            Hoja12.Select
                            Limpiar_Filtro
                            Orden_Filtro
                            Range("J1").Select
                            
                                Do Until IsEmpty(ActiveCell)
                                    ActiveCell.Offset(1, 0).Select
                                    If ActiveCell.Value Like Referencia Then
                                        encontrado = True
                                        ActiveCell.Offset(0, 0).Value = Registrado
                                        ActiveCell.Offset(0, -2).Value = CDate(frm_Cuota_Abono.txt_Fecha)
                                                                   
                                    Exit Do
                                    End If
                                Loop

                Next b



End Sub
Private Sub Limpiar_CuotaAbono()
    Me.ListBox1.Clear
    Me.txtTotal = Empty
End Sub

Private Sub UserForm_Initialize()
With ListBox1
    Me.ListBox1.ColumnCount = 8
    Me.ListBox1.ColumnWidths = "50 pt;250 pt;90 pt;200 pt;85 pt;85 pt;70 pt;0 pt"
End With
End Sub
Private Sub Limpiar_Filtro()

Range("A1").Select

          If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

           If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData


End Sub
Private Sub Orden_Filtro()

Range("A1").Sort Key1:=Range("A1"), Order1:=xlDescending, Header:=xlYes

End Sub


