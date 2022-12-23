Attribute VB_Name = "Funciones"


Public Function nReg(Hoja As Worksheet, nFila As Long, nColumna As Long)
    Do Until IsEmpty(Hoja.Cells(nFila, nColumna))
        nFila = nFila + 1
    Loop
    nReg = nFila
End Function

Public Function LimpiarControlesContables(xTag As String, xForm As UserForm)
Dim xCtrl As Control

    For Each xCtrl In xForm.Controls
        If xCtrl.Tag = xTag Then
            xCtrl = Empty
        End If
    Next
End Function

Public Function ControlesVaciosContables(xTag As String, xForm As UserForm, Optional xContenedor As Object, _
Optional Switch As Boolean) As Boolean

Dim xCtrl As Control

If Switch = True Then
  For Each xCtrl In xContenedor.Controls
        
        If xCtrl.Tag = xTag And xCtrl = Empty Then
            ControlesVaciosContables = True
            MsgBox "Debe rellenar el campo: " & UCase(xCtrl.ControlTipText), vbInformation
            xCtrl.SetFocus
            xCtrl.BackColor = RGB(211, 255, 211)
            Exit Function
        End If
    Next

Else

    For Each xCtrl In xForm.Controls
        If xCtrl.Tag = xTag And xCtrl = Empty Then
            ControlesVaciosContables = True
            MsgBox "Debe rellenar el campo: " & UCase(xCtrl.ControlTipText), vbInformation
            xCtrl.SetFocus
            xCtrl.BackColor = RGB(211, 255, 211)
            Exit Function
        End If
    Next
End If

End Function


Public Function LanzarListadoCuentas(CualquierFormulario As Object, CualquierControl As String)
Dim xCtrl As Control

    Load frm_ListadoCuentas
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = CualquierControl Then
            frm_ListadoCuentas.StartUpPosition = 0
            frm_ListadoCuentas.Left = CualquierFormulario.Left + xCtrl.Left
            frm_ListadoCuentas.Top = CualquierFormulario.Top
        End If
    Next
    
    frm_ListadoCuentas.Show

End Function


