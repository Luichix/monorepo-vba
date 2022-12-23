VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_clientes2 
   Caption         =   "Registro de Clientes"
   ClientHeight    =   4150
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   8230.001
   OleObjectBlob   =   "frm_clientes2.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_clientes2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ArchivoIMG As String

Private Sub cmd_Agregar_Click()
    Dim i As Integer
    
    If cbo_Nombre.Text = "" Then
        MsgBox "Nombre inválido", vbInformation + vbOKOnly
        cbo_Nombre.SetFocus
        Exit Sub
    End If
    
    If Not (Mid(cbo_Nombre.Text, 1, 1) Like "[a-z]" Or Mid(cbo_Nombre.Text, 1, 1) Like "[A-Z]") Then
        MsgBox "Nombre inválido", vbInformation + vbOKOnly
        cbo_Nombre.SetFocus
        Exit Sub
    End If
    
    For i = 2 To Len(cbo_Nombre.Text)
        If Mid(cbo_Nombre.Text, i, 1) Like "#" Then
            MsgBox "Nombre inválido", vbInformation + vbOKOnly
            cbo_Nombre.SetFocus
            Exit Sub
        End If
    Next
    
    Sheets("Clientes").Activate
    
    Dim fCliente As Integer
    fCliente = nCliente(cbo_Nombre.Text)
    
    If fCliente = 0 Then
        Do While Not IsEmpty(ActiveCell)
            ActiveCell.Offset(1, 0).Activate ' si el registro no existe, se va al final.
        Loop
    Else
        Cells(fCliente, 1).Select  ' cuando ya existe el registro, cumple esta condición.
    End If
  
    'Aqui es cuando agregamos o modificamos el registro
    Application.ScreenUpdating = False
    ActiveCell = cbo_Nombre
    ActiveCell.Offset(0, 1) = txt_Direccion
    ActiveCell.Offset(0, 2) = txt_telefono
    ActiveCell.Offset(0, 3) = txt_ID
    ActiveCell.Offset(0, 4) = txt_email
    ActiveCell.Offset(0, 5) = ArchivoIMG
        
    Application.ScreenUpdating = True
   
    LimpiarFormulario
    
    cbo_Nombre.SetFocus

End Sub
Private Sub cmd_Eliminar_Click()
    Dim fCliente As Integer
    fCliente = nCliente(cbo_Nombre.Text)
    
    If fCliente = 0 Then
        MsgBox "El cliente que usted quiere eliminar no existe", vbInformation + vbOKOnly
        cbo_Nombre.SetFocus
        Exit Sub
    End If
    
    If MsgBox("¿Seguro que quiere eliminar este cliente?", vbQuestion + vbYesNo) = vbYes Then
        
        Cells(fCliente, 1).Select
        
        ActiveCell.EntireRow.Delete
        
        LimpiarFormulario
        
        MsgBox "Cliente eliminado", vbInformation + vbOKOnly
        cbo_Nombre.SetFocus
        
   End If

End Sub
Private Sub cmd_Cerrar_Click()
End
End Sub
Private Sub cbo_Nombre_Change()
On Error Resume Next


    If nCliente(cbo_Nombre.Text) <> 0 Then
        
        Sheets("Clientes").Activate
        
        Cells(cbo_Nombre.ListIndex + 2, 1).Select
        txt_Direccion = ActiveCell.Offset(0, 1)
        txt_telefono = ActiveCell.Offset(0, 2)
        txt_ID = ActiveCell.Offset(0, 3)
        txt_email = ActiveCell.Offset(0, 4)
        
        fotografia.Picture = LoadPicture("")
        fotografia.Picture = LoadPicture(ActiveCell.Offset(0, 5))
        
        ArchivoIMG = ActiveCell.Offset(0, 5)
       
    Else
        txt_Direccion = ""
        txt_telefono = ""
        txt_ID = ""
        txt_email = ""
        ArchivoIMG = ""
        fotografia.Picture = LoadPicture("")
    End If
End Sub
Private Sub cbo_Nombre_Enter()
CargarLista
End Sub
Sub CargarLista()
   cbo_Nombre.Clear
    
    Sheets("Clientes").Select
    Range("A2").Select
    Do While Not IsEmpty(ActiveCell)
        cbo_Nombre.AddItem ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
    Loop
End Sub

Sub LimpiarFormulario()
    CargarLista
    
    cbo_Nombre = ""
    txt_Direccion = ""
    txt_telefono = ""
    txt_ID = ""
    txt_email = ""
    ArchivoIMG = ""
End Sub
Private Sub cmd_Imagen_Click()
On Error Resume Next
        
        ArchivoIMG = Application.GetOpenFilename("Imágenes jpg,*.jpg,Imágenes bmp,*.bmp", 0, "Seleccionar Imágen para Registro de Clientes")
        fotografia.Picture = LoadPicture("")
        fotografia.Picture = LoadPicture(ArchivoIMG)

End Sub
