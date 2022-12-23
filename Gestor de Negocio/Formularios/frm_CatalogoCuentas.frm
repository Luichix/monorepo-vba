VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_CatalogoCuentas 
   Caption         =   "Administrar Catálogo de Cuentas"
   ClientHeight    =   1870
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   7020
   OleObjectBlob   =   "frm_CatalogoCuentas.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_CatalogoCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btn_Agregar_Click()
Dim Fila As Long
Dim Final As Long

Application.ScreenUpdating = False


        If ControlesVaciosContables("plan_contable", Me) = True Then Exit Sub

        Final = nReg(Hoja41, 2, 1)

With Hoja41

        For Fila = 2 To Final
            If .Cells(Fila, 1) = Val(Me.cbo_CodCuenta) Then
                MsgBox ("Código de Cuenta ya existe!" + Chr(13) + "Ingrese uno diferente"), vbInformation
                Me.cbo_CodCuenta.SetFocus
                Me.cbo_CodCuenta.BackColor = RGB(211, 255, 211)
                Exit Sub
            End If
        Next
                          
        .Unprotect "355365847"
        
        If Len(Me.cbo_CodCuenta) = 1 Or Len(Me.cbo_CodCuenta) = 2 Then
            .Range(.Cells(Final, 1), .Cells(Final, 3)).Font.Bold = True
            
            ElseIf Len(Me.cbo_CodCuenta) = 3 Then
                .Range(.Cells(Final, 1), .Cells(Final, 3)).Interior.Color = RGB(190, 190, 190)
                .Range(.Cells(Final, 1), .Cells(Final, 3)).Font.Color = RGB(255, 255, 255)
                .Range(.Cells(Final, 1), .Cells(Final, 3)).Font.Bold = True
        End If



                .Cells(Final, 1) = Me.cbo_CodCuenta.Value
                .Cells(Final, 2) = Me.txt_NombreCuenta.Text
                .Cells(Final, 3) = nGrupo
                Call IndexarCodCuentasPLAN
        .Protect "355365847"
End With
                
                Me.cbo_CodCuenta = Me.cbo_CodCuenta + 1
                Me.txt_NombreCuenta = Empty
                Me.txt_NombreCuenta.SetFocus
                
Application.ScreenUpdating = True

End Sub

Private Sub btn_Eliminar_Click()
Dim Fila As Long
Dim Final As Long

Application.ScreenUpdating = False
    
    Final = nReg(Hoja41, 2, 1) - 1
    
If Me.cbo_CodCuenta = Empty Then
    MsgBox "Seleccione un registro para eliminar", vbInformation
    Me.cbo_CodCuenta.BackColor = RGB(211, 255, 211)
    Me.cbo_CodCuenta.SetFocus
    Exit Sub
End If

If Me.txt_NombreCuenta = Empty Then
    MsgBox "La cuenta " & Me.cbo_CodCuenta & " no existe!", vbInformation
    Me.cbo_CodCuenta = Empty
    Me.cbo_CodCuenta.BackColor = RGB(211, 255, 211)
    Me.cbo_CodCuenta.SetFocus
    Exit Sub
End If


With Hoja41


    For Fila = 2 To Final
        
            If .Cells(Fila, 1) = Val(Me.cbo_CodCuenta) _
                And Mid(.Cells(Fila + 1, 1), 1, 1) = Val(Me.cbo_CodCuenta) Then
                MsgBox "Este elemento contable tiene rubros asociados y no puede ser eliminado!", vbCritical
                Exit Sub
                    ElseIf .Cells(Fila, 1) = Val(Me.cbo_CodCuenta) _
                        And Mid(.Cells(Fila + 1, 1), 1, 2) = Val(Me.cbo_CodCuenta) Then
                        MsgBox "Este rubro tiene cuentas de mayor asociadas y no puede ser eliminado!", vbCritical
                        Exit Sub
                        
                    ElseIf .Cells(Fila, 1) = Val(Me.cbo_CodCuenta) _
                        And Mid(.Cells(Fila + 1, 1), 1, 3) = Val(Me.cbo_CodCuenta) Then
                        MsgBox "Esta cuenta de mayor tiene cuentas asociadas y no puede ser eliminada!", vbCritical
                        Exit Sub
                    
                    ElseIf .Cells(Fila, 1) = Val(Me.cbo_CodCuenta) _
                        And Mid(.Cells(Fila + 1, 1), 1, 5) = Val(Me.cbo_CodCuenta) Then
                        MsgBox "Esta cuenta tiene subcuentas asociadas y no puede ser eliminada!", vbCritical
                        Exit Sub
            End If
    Next

            If MsgBox("¿Seguro que quiere eliminar esta Cuenta?", vbQuestion + vbYesNo) = vbYes Then
                     For Fila = 2 To Final
                        If .Cells(Fila, 1) = Val(Me.cbo_CodCuenta) Then
                            .Unprotect "355365847"
                            .Cells(Fila, 1).EntireRow.Delete
                            .Protect "355365847"
                            Exit For
                         End If
                     Next
                    
                     MsgBox "La cuenta: " & Me.cbo_CodCuenta & " ha sido eliminada!", vbInformation
                     Call LimpiarControlesContables("plan_contable", Me)
                     Me.cbo_CodCuenta.SetFocus
                Else
                        Exit Sub
            End If
End With

Application.ScreenUpdating = True

End Sub

Private Sub btn_ListadoCuentas_Click()
banderaListadoCuentas = 1
    Call LanzarListadoCuentas(Me, "lbl_LanzarListadoCuentas")
End Sub

Private Sub btn_Salir_Click()
    Unload Me
End Sub

Private Sub cbo_CodCuenta_Change()
Dim Fila As Long
Dim Final As Long
Dim encontrado As Boolean
    
    Me.cbo_CodCuenta.BackColor = RGB(255, 255, 255)
    
    Call ValidarCuenta
    
If Me.cbo_CodCuenta = Empty Then
    Me.txt_NombreCuenta = Empty
    Exit Sub
End If

    
Final = nReg(Hoja41, 2, 1) - 1

    For Fila = 2 To Final
        If Hoja41.Cells(Fila, 1) = Val(Me.cbo_CodCuenta) Then
            encontrado = True
            Me.txt_NombreCuenta = Hoja41.Cells(Fila, 2)
            Exit For
        End If
    Next
    
    If encontrado = False Then
        Me.txt_NombreCuenta = Empty
    End If

End Sub

Private Sub cbo_CodCuenta_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As Long


Do While Me.cbo_CodCuenta.ListCount > 0
    Me.cbo_CodCuenta.RemoveItem 0
Loop

    
    Final = nReg(Hoja41, 2, 1) - 1

        For Fila = 2 To Final
            Lista = Hoja41.Cells(Fila, 1)
            Me.cbo_CodCuenta.AddItem Lista
        Next

End Sub

Private Sub cbo_CodCuenta_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_NombreCuenta_Change()
    Me.txt_NombreCuenta.BackColor = RGB(255, 255, 255)
    Me.txt_NombreCuenta = UCase(Me.txt_NombreCuenta)
End Sub


Private Sub UserForm_Click()

End Sub
