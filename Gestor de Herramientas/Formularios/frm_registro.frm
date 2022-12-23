VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_registro 
   Caption         =   "GESTOR DE HERRAMIENTAS"
   ClientHeight    =   4320
   ClientLeft      =   40
   ClientTop       =   310
   ClientWidth     =   10500
   OleObjectBlob   =   "frm_registro.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frm_registro"
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
    Dim iTem As Integer

On Error GoTo Salir

    Application.ScreenUpdating = False
    Titulo = "Gestor de Herramientas"
    
    IdCliente = Hoja5.Range("G2").Value + 1
 
    
        
    If Me.txt_nombre.Text = "" Then
        Me.txt_nombre.BackColor = &HC0C0FF
        MsgBox "Ingrese el nombre de la herramienta", vbInformation, Titulo
        Me.txt_nombre.SetFocus
        Exit Sub
    End If
                       If Me.txt_detalle.Text = "" Then
                           Me.txt_detalle.BackColor = &HC0C0FF
                           MsgBox "Ingrese el detalle de la herramienta", vbInformation, Titulo
                           Me.txt_detalle.SetFocus
                           Exit Sub
                        End If
        
        Final = GetNuevoR(Hoja1)

        'Validación para impedir Clientes repetidos
        For Fila = 2 To Final
            If Hoja1.Cells(Fila, 1) = UCase(Me.txt_nombre.Text) Then
                MsgBox "Herramienta ya existe en la Base de Datos", vbInformation, Titulo
                LimpiarControles
                Me.txt_nombre.SetFocus
                Exit Sub
                Exit For
            End If
        Next

      If MsgBox("Son correctos los datos?" + Chr(13) + "Desea proceder?", vbOKCancel, Titulo) = vbNo Then
          Exit Sub

      Else
         
           Hoja1.Activate
            Hoja1.Select


                    Hoja1.Rows("2:2").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                    
                Hoja1.Cells(2, 1) = Hoja1.Cells(3, 1) + 1
                Hoja1.Cells(2, 2) = "H0" & IdCliente
                Hoja1.Cells(2, 3) = UCase(Me.txt_nombre.Text)
                Hoja1.Cells(2, 4) = UCase(Me.txt_detalle.Text)
                '-----------------------------------------------
      End If
    
    MsgBox "Registro Realizado Correctamente", , Titulo
    LimpiarControles
    Hoja5.Range("G2").Value = Hoja5.Range("G2").Value + 1

    lbl_cliente.Caption = "ITEM HERRAMIENTA - H0" & Hoja5.Range("G2").Value + 1
    
    Unload Me
    Hoja0.Activate
    Hoja0.Select
    
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
                Me.txt_nombre.SetFocus
            End If
        Next

End Sub


Private Sub UserForm_Initialize()

lbl_cliente.Caption = "ITEM HERRAMIENTA - H0" & Hoja5.Range("G2").Value + 1

End Sub
