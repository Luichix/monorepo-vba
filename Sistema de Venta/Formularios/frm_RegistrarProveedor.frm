VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_RegistrarProveedor 
   Caption         =   "AGENDA "
   ClientHeight    =   5820
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   8980.001
   OleObjectBlob   =   "frm_RegistrarProveedor.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_RegistrarProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim Fila As Long
    Dim Final As Long
    Dim Registro As Integer
    Dim Titulo As String
    Dim IdProveedor As Integer

On Error GoTo Salir

    Application.ScreenUpdating = False
    Titulo = "Gestor de Ventas"
    
    IdProveedor = Hoja93.Range("E2").Value + 1
        
    If Me.txt_Proveedor.Text = "" Then
        Me.txt_Proveedor.BackColor = &HC0C0FF
        MsgBox "Ingrese el nombre del Proveedor", vbInformation, Titulo
        Me.txt_Proveedor.SetFocus
        Exit Sub
    End If
        If Me.txt_nRegistroFiscal.Text = "" Then
            Me.txt_nRegistroFiscal.BackColor = &HC0C0FF
            MsgBox "Ingrese el Número de Registro Fiscal del Proveedor, vbInformation, Titulo"
            Me.txt_nRegistroFiscal.SetFocus
            Exit Sub
        End If
                If Me.txt_telf.Text = "" Then
                    Me.txt_telf.BackColor = &HC0C0FF
                    MsgBox "Ingrese el número telefono del Proveedor", vbInformation, Titulo
                    Me.txt_telf.SetFocus
                    Exit Sub
                 End If
                        If Me.txt_direccion.Text = "" Then
                           Me.txt_direccion.BackColor = &HC0C0FF
                           MsgBox "Ingrese la dirección del Proveedor", vbInformation, Titulo
                           Me.txt_direccion.SetFocus
                           Exit Sub
                        End If

        'Determina el final del listado de Clientes
        Final = GetNuevoR(Hoja8)

        'Validación para impedir Clientes repetidos
        For Fila = 2 To Final
            If Hoja8.Cells(Fila, 1) = UCase(Me.txt_Proveedor.Text) Then
                MsgBox "Proveedor ya existe en la Base de Datos", vbInformation, Titulo
                LimpiarControles
                Me.txt_Proveedor.SetFocus
                Exit Sub
                Exit For
            End If
        Next

      If MsgBox("Son correctos los datos?" + Chr(13) + "Desea proceder?", vbOKCancel, Titulo) = vbNo Then
          Exit Sub

      Else
           If Hoja8.Visible = xlSheetVisible Then

                Hoja8.Select
                    Hoja8.Range("A2:F2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja8.Range("A3:F3").Select
                    Selection.Copy
                    Hoja8.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                Hoja8.Cells(2, 1) = IdProveedor
                Hoja8.Cells(2, 2) = UCase(Me.txt_Proveedor.Text)
                Hoja8.Cells(2, 3) = UCase(Me.txt_nRegistroFiscal)
                Hoja8.Cells(2, 4) = "'" & Me.txt_telf
                Hoja8.Cells(2, 5) = UCase(Me.txt_direccion)
                Hoja8.Cells(2, 6) = Hoja92.Range("G1")
                '-----------------------------------------------

            ElseIf Hoja8.Visible = xlSheetVeryHidden Then
                Hoja8.Visible = xlSheetVisible

                  Hoja8.Select
                    Hoja8.Range("A2:F2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja8.Range("A3:F3").Select
                    Selection.Copy
                    Hoja8.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                Hoja8.Cells(2, 1) = IdProveedor
                Hoja8.Cells(2, 2) = UCase(Me.txt_Proveedor.Text)
                Hoja8.Cells(2, 3) = UCase(Me.txt_nRegistroFiscal)
                Hoja8.Cells(2, 4) = "'" & Me.txt_telf
                Hoja8.Cells(2, 5) = UCase(Me.txt_direccion)
                Hoja8.Cells(2, 6) = Hoja92.Range("G1")
                '-----------------------------------------------

               Hoja8.Visible = xlSheetVeryHidden

            End If
      End If
    
    MsgBox "Registro Realizado Correctamente", , Titulo
    LimpiarControles
    Hoja93.Range("E2").Value = Hoja93.Range("E2").Value + 1

    lbl_proveedor.Caption = "ID PROVEEDOR - " & Hoja93.Range("E2").Value + 1
    
     Application.ScreenUpdating = True
     
     Unload Me
         Application.EnableEvents = False
        ThisWorkbook.Save
    Application.EnableEvents = True
     
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, Titulo
 End If

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub LimpiarControles()
    Dim xTextBox As Control
        
        For Each xTextBox In Controls
            If xTextBox.Name Like "txt*" Then
                xTextBox = Empty
                Me.txt_Proveedor.SetFocus
            End If
        Next

End Sub


Private Sub UserForm_Initialize()

lbl_proveedor.Caption = "ID PROVEEDOR - " & Hoja93.Range("E2").Value + 1

End Sub

