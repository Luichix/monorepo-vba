VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_RegistrarClientes 
   Caption         =   "GESTOR DE VENTAS"
   ClientHeight    =   5820
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   8910.001
   OleObjectBlob   =   "frm_RegistrarClientes.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_RegistrarClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim Fila As Long
    Dim Final As Long
    Dim Registro As Integer
    Dim Titulo As String
    Dim xTextBox As Control
    Dim IdCliente As Integer

On Error GoTo Salir

    Application.ScreenUpdating = False
    Titulo = "Gestor de Ventas"
    
    IdCliente = Hoja93.Range("D2").Value + 1
        
    If Me.txt_cliente.Text = "" Then
        Me.txt_cliente.BackColor = &HC0C0FF
        MsgBox "Ingrese el nombre del cliente", vbInformation, Titulo
        Me.txt_cliente.SetFocus
        Exit Sub
    End If
        If Me.txt_Ruc.Text = "" Then
            Me.txt_Ruc.BackColor = &HC0C0FF
            MsgBox "Ingrese la identificación del cliente", vbInformation, Titulo
            Me.txt_Ruc.SetFocus
            Exit Sub
        End If
                If Me.txt_telf.Text = "" Then
                    Me.txt_telf.BackColor = &HC0C0FF
                    MsgBox "Ingrese el número telefono del cliente", vbInformation, Titulo
                    Me.txt_telf.SetFocus
                    Exit Sub
                 End If
                        If Me.txt_direccion.Text = "" Then
                           Me.txt_direccion.BackColor = &HC0C0FF
                           MsgBox "Ingrese la dirección del cliente", vbInformation, Titulo
                           Me.txt_direccion.SetFocus
                           Exit Sub
                        End If

        'Determina el final del listado de Clientes
        Final = GetNuevoR(Hoja7)

        'Validación para impedir Clientes repetidos
        For Fila = 2 To Final
            If Hoja7.Cells(Fila, 1) = UCase(Me.txt_cliente.Text) Then
                MsgBox "Cliente ya existe en la Base de Datos", vbInformation, Titulo
                LimpiarControles
                Me.txt_cliente.SetFocus
                Exit Sub
                Exit For
            End If
        Next

      If MsgBox("Son correctos los datos?" + Chr(13) + "Desea proceder?", vbOKCancel, Titulo) = vbNo Then
          Exit Sub

      Else
           If Hoja7.Visible = xlSheetVisible Then

                Hoja7.Select
                    Hoja7.Range("A2:F2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja7.Range("A3:F3").Select
                    Selection.Copy
                    Hoja7.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                Hoja7.Cells(2, 1) = IdCliente
                Hoja7.Cells(2, 2) = UCase(Me.txt_cliente.Text)
                Hoja7.Cells(2, 3) = UCase(Me.txt_Ruc)
                Hoja7.Cells(2, 4) = "'" & Me.txt_telf
                Hoja7.Cells(2, 5) = UCase(Me.txt_direccion)
                Hoja7.Cells(2, 6) = Hoja92.Range("G1")
                '-----------------------------------------------

            ElseIf Hoja7.Visible = xlSheetVeryHidden Then
                Hoja7.Visible = xlSheetVisible

                 Hoja7.Select
                    Hoja7.Range("A2:F2").Select
                    Selection.ListObject.ListRows.Add (1)
                    Hoja7.Range("A3:F3").Select
                    Selection.Copy
                    Hoja7.Range("A2").Select
                    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                    Application.CutCopyMode = False

                Hoja7.Cells(2, 1) = IdCliente
                Hoja7.Cells(2, 2) = UCase(Me.txt_cliente.Text)
                Hoja7.Cells(2, 3) = UCase(Me.txt_Ruc)
                Hoja7.Cells(2, 4) = "'" & Me.txt_telf
                Hoja7.Cells(2, 5) = UCase(Me.txt_direccion)
                Hoja7.Cells(2, 6) = Hoja92.Range("G1")
                '-----------------------------------------------

               Hoja7.Visible = xlSheetVeryHidden

            End If
      End If
    
    MsgBox "Registro Realizado Correctamente", , Titulo
    LimpiarControles
    Hoja93.Range("D2").Value = Hoja93.Range("D2").Value + 1

    lbl_cliente.Caption = "ID CLIENTE - " & Hoja93.Range("D2").Value + 1
    
    Unload Me
    
     Application.ScreenUpdating = True
     
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
                Me.txt_cliente.SetFocus
            End If
        Next

End Sub


Private Sub UserForm_Initialize()

lbl_cliente.Caption = "ID CLIENTE - " & Hoja93.Range("D2").Value + 1

End Sub
