VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Ganado 
   Caption         =   "GESTIÓN DE ANIMALES"
   ClientHeight    =   10200
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   10320
   OleObjectBlob   =   "frm_Ganado.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Ganado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim ArchivoIMG(4) As String

Private Sub btn_Registrar_Click()
    
If Me.Text_fecha1.Text = Empty Or _
    Me.txt_nombre = Empty Or _
    txt_codigo = Empty Or _
    cbx_raza = Empty Or _
    Text_fecha2.Text = Empty Or _
    cbx_sexo = Empty Or _
    cbx_origen = Empty Or _
    cbx_ubicacion = Empty Or _
    cbx_proposito = Empty Or _
    cbx_rodeo = Empty Then
    
           
            MsgBox "Hay campos vacíos en el registro", , "Gestor de Ganaderia"
            Exit Sub
    
End If
    

If MsgBox("¿Son Correctos los Datos?" + Chr(13) + "¿Desea Continuar?", vbYesNo, "Gestor de Ganaderia") = vbNo Then
        Exit Sub
    Else
        RegistrarAnimal
        
        MsgBox "Registro procesado con éxito!!!", , "Gestor de Ganaderia"
        
End If

       

End Sub

Private Sub btn_Salir_Click()
Unload Me
End Sub
Private Sub RegistrarAnimal()
Dim Comprb As Long
Dim Hembra As String
Dim Macho As String
Dim Criollo As String
Dim Compra As String


'Aquí manejo el correlativo del comprobante
Hoja22.Range("E2").Value = Hoja22.Range("E2").Value + 1
Comprb = Hoja22.Range("E2").Value


    Hoja29.Select
    Hoja29.Range("A2:W2").Select
    Selection.ListObject.ListRows.Add (1)
    Hoja29.Range("A3:w2").Select
    Selection.Copy
    Hoja29.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
            Hoja29.Cells(2, 1) = Comprb                 'Número de Registro
            Hoja29.Cells(2, 8) = CDate(Me.Text_fecha1)  'Fecha de Nacimiento
            Hoja29.Cells(2, 2) = CDate(Me.Text_fecha2)  'Fecha de Incorporación
            Hoja29.Cells(2, 5) = txt_nombre             'Nombre
            Hoja29.Cells(2, 4) = txt_codigo.Value       'Código
            Hoja29.Cells(2, 6) = cbx_raza               'Raza
            Hoja29.Cells(2, 3) = cbx_ubicacion          'Ubicación
            Hoja29.Cells(2, 7) = cbx_proposito          'Proposito
            Hoja29.Cells(2, 11) = cbx_rodeo             'Rodeo
            Hoja29.Cells(2, 10) = cbx_sexo              'Sexo
            Hoja29.Cells(2, 13) = cbx_origen            'Origen
            
            
            If ComboBox1 = "" Then
                Hoja29.Cells(2, 14) = "DESCONOCIDO"
                Hoja29.Cells(2, 15) = "DESCONOCIDO"
            Else
                Hoja29.Cells(2, 14) = ComboBox1.Value     'Codigo Madre
                 Hoja29.Cells(2, 15) = cbx_madre          'Nombre Madre
            End If
            
            If ComboBox2 = "" Then
                Hoja29.Cells(2, 16) = "DESCONOCIDO"
                Hoja29.Cells(2, 17) = "DESCONOCIDO"
             
            Else
                Hoja29.Cells(2, 16) = ComboBox2.Value            'Codigo Padre
                Hoja29.Cells(2, 17) = cbx_padre            'Nombre Padre
            End If
                               
            Hoja29.Cells(2, 23) = ArchivoIMG(1)             'Foto del Animal
            Hoja29.Cells(2, 20) = ArchivoIMG(2)             'Fierro 1
            Hoja29.Cells(2, 21) = ArchivoIMG(3)             'Fierro 2
            Hoja29.Cells(2, 22) = ArchivoIMG(4)             'Fierro 3
           
       
        LimpiarControles
        txt_nombre.SetFocus
     
     
   Me.Text_fecha1 = Date
   Me.Text_fecha2 = Date
   Me.Label16.Caption = "No. " & Hoja22.Range("E2").Value + 1 'Llamamos el número de la factura
    
End Sub


Private Sub CheckBox1_Click()
    If CheckBox1 = True Then
        cmd_Agregarfierro1.Enabled = True
        Else
        cmd_Agregarfierro1.Enabled = False
        img_fierro1.Picture = LoadPicture("")
    End If
End Sub

Private Sub CheckBox2_Click()
    If CheckBox2 = True Then
        cmd_Agregarfierro2.Enabled = True
        Else
        cmd_Agregarfierro2.Enabled = False
        img_fierro2.Picture = LoadPicture("")
    End If
End Sub

Private Sub CheckBox3_Click()
    If CheckBox3 = True Then
        cmd_Agregarfierro3.Enabled = True
        Else
        cmd_Agregarfierro3.Enabled = False
        img_fierro3.Picture = LoadPicture("")
    End If
End Sub

Private Sub cmd_Agregar_Click()
On Error Resume Next

        ArchivoIMG(1) = Application.GetOpenFilename("Imágenes jpg,*.jpg,Imágenes bmp,*.bmp", 0, "Seleccionar Imágen para Registro del Animal")
        img_animales.Picture = LoadPicture("")
        img_animales.Picture = LoadPicture(ArchivoIMG(1))

End Sub
Private Sub cmd_Agregarfierro1_Click()
On Error Resume Next

        ArchivoIMG(2) = Application.GetOpenFilename("Imágenes jpg,*.jpg,Imágenes bmp,*.bmp", 0, "Seleccionar Imágen para Registro de Clientes")
        img_fierro1.Picture = LoadPicture("")
        img_fierro1.Picture = LoadPicture(ArchivoIMG(2))

End Sub
Private Sub cmd_Agregarfierro2_Click()
On Error Resume Next

        ArchivoIMG(3) = Application.GetOpenFilename("Imágenes jpg,*.jpg,Imágenes bmp,*.bmp", 0, "Seleccionar Imágen para Registro de Clientes")
        img_fierro2.Picture = LoadPicture("")
        img_fierro2.Picture = LoadPicture(ArchivoIMG(3))
        
End Sub
Private Sub cmd_Agregarfierro3_Click()
On Error Resume Next

        ArchivoIMG(4) = Application.GetOpenFilename("Imágenes jpg,*.jpg,Imágenes bmp,*.bmp", 0, "Seleccionar Imágen para Registro de Clientes")
        img_fierro3.Picture = LoadPicture("")
        img_fierro3.Picture = LoadPicture(ArchivoIMG(4))

End Sub



Private Sub UserForm_Initialize()
 Me.Text_fecha1 = Date
 Me.Text_fecha2 = Date
 Me.Label16.Caption = "No  " & Hoja22.Range("E2").Value + 1 'Llamamos el número de la factura
 Me.cmd_Agregarfierro1.Enabled = False
 Me.cmd_Agregarfierro2.Enabled = False
 Me.cmd_Agregarfierro3.Enabled = False
End Sub
Private Sub btn_FechaDat1_Click()
banderaCalendario = 7
    Call LanzarCalendario(Me, "lbl_FechaSal")
End Sub
Private Sub btn_FechaAni2_Click()
banderaCalendario = 8
    Call LanzarCalendario(Me, "lbl_FechaSal")
End Sub
Private Sub LimpiarControles()
    txt_nombre = Empty
    txt_codigo = Empty
    Text_fecha1 = Empty
    Text_fecha2 = Empty
    cbx_raza = Empty
    ComboBox1 = Empty
    ComboBox2 = Empty
    cbx_origen = Empty
    cbx_sexo = Empty
    CheckBox1 = False
    CheckBox2 = False
    CheckBox3 = False
    cbx_ubicacion = Empty
    cbx_proposito = Empty
    cbx_rodeo = Empty
    img_animales.Picture = LoadPicture("")
    img_fierro1.Picture = LoadPicture("")
    img_fierro2.Picture = LoadPicture("")
    img_fierro3.Picture = LoadPicture("")
 End Sub



Private Sub cbx_ubicacion_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To cbx_ubicacion.ListCount
    cbx_ubicacion.RemoveItem 0
Next Fila

        For Fila = 2 To 5
                Lista = Hoja1.Cells(Fila, 36)
                cbx_ubicacion.AddItem (Lista)
        Next
End Sub
Private Sub cbx_proposito_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To cbx_proposito.ListCount
    cbx_proposito.RemoveItem 0
Next Fila

        For Fila = 2 To 3
                Lista = Hoja1.Cells(Fila, 34)
                cbx_proposito.AddItem (Lista)
        Next
End Sub
Private Sub cbx_rodeo_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To cbx_rodeo.ListCount
    cbx_rodeo.RemoveItem 0
Next Fila
  
        For Fila = 2 To 8
                Lista = Hoja1.Cells(Fila, 32)
                cbx_rodeo.AddItem (Lista)
        Next
End Sub

Private Sub cbx_raza_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To cbx_raza.ListCount
    cbx_raza.RemoveItem 0
Next Fila
  
        For Fila = 2 To 7
                Lista = Hoja1.Cells(Fila, 28)
                cbx_raza.AddItem (Lista)
        Next
End Sub
Private Sub cbx_sexo_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To cbx_sexo.ListCount
    cbx_sexo.RemoveItem 0
Next Fila
  
        For Fila = 2 To 3
                Lista = Hoja1.Cells(Fila, 30)
                cbx_sexo.AddItem (Lista)
        Next
End Sub
Private Sub cbx_origen_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To cbx_origen.ListCount
    cbx_origen.RemoveItem 0
Next Fila
  
        For Fila = 2 To 3
                Lista = Hoja1.Cells(Fila, 40)
                cbx_origen.AddItem (Lista)
        Next
End Sub

Private Sub ComboBox1_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox1.ListCount
    ComboBox1.RemoveItem 0
Next Fila
  
Final = GetUltimoR(Hoja29)

        For Fila = 2 To Final
            If Hoja29.Cells(Fila, 10) = "HEMBRA" Then
                Lista = Hoja29.Cells(Fila, 4)
                ComboBox1.AddItem (Lista)
            End If
        Next
End Sub
Private Sub ComboBox2_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To ComboBox2.ListCount
    ComboBox2.RemoveItem 0
Next Fila
  
Final = GetUltimoR(Hoja29)

        For Fila = 2 To Final
            If Hoja29.Cells(Fila, 10) = "MACHO" Then
                Lista = Hoja29.Cells(Fila, 4)
                ComboBox2.AddItem (Lista)
            End If
        Next
End Sub
Private Sub cbx_madre_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To cbx_madre.ListCount
    cbx_madre.RemoveItem 0
Next Fila
  
Final = GetUltimoR(Hoja29)

        For Fila = 2 To Final
            If Hoja29.Cells(Fila, 10) = "HEMBRA" Then
                Lista = Hoja29.Cells(Fila, 5)
                cbx_madre.AddItem (Lista)
            End If
        Next
End Sub
Private Sub cbx_padre_Enter()
Dim Fila As Long
Dim Final As Long
Dim Lista As String


For Fila = 1 To cbx_padre.ListCount
    cbx_padre.RemoveItem 0
Next Fila
  
Final = GetUltimoR(Hoja29)

        For Fila = 2 To Final
            If Hoja29.Cells(Fila, 10) = "MACHO" Then
                Lista = Hoja29.Cells(Fila, 5)
                cbx_padre.AddItem (Lista)
            End If
        Next
End Sub

Private Sub ComboBox1_Change()
Dim Fila As Long
Dim Final As Long
Dim Registro As Integer

Me.ComboBox1.BackColor = &H80000005

If ComboBox1.Text = "" Then
    cbx_madre = Empty
End If

    
Final = GetUltimoR(Hoja29)

    For Fila = 2 To Final
        If ComboBox1.Text = Hoja29.Cells(Fila, 4) Then
            Me.cbx_madre.Text = Hoja29.Cells(Fila, 5)
            Exit For
        
        End If
    Next
End Sub
Private Sub cbx_madre_Change()
Dim Fila As Long
Dim Final As Long
Dim Registro As Integer

Me.cbx_madre.BackColor = &H80000005

If cbx_madre.Text = "" Then
    ComboBox1 = Empty
End If

    
Final = GetUltimoR(Hoja29)

    For Fila = 2 To Final
        If cbx_madre.Text = Hoja29.Cells(Fila, 5) Then
            Me.ComboBox1.Text = Hoja29.Cells(Fila, 4)
            Exit For
        
        End If
    Next
End Sub
Private Sub ComboBox2_Change()
Dim Fila As Long
Dim Final As Long
Dim Registro As Integer

Me.ComboBox2.BackColor = &H80000005

If ComboBox2.Text = "" Then
    cbx_padre = Empty
End If

    
Final = GetUltimoR(Hoja29)

    For Fila = 2 To Final
        If ComboBox2.Text = Hoja29.Cells(Fila, 4) Then
            Me.cbx_padre.Text = Hoja29.Cells(Fila, 5)
            Exit For
        
        End If
    Next
End Sub
Private Sub cbx_padre_Change()
Dim Fila As Long
Dim Final As Long
Dim Registro As Integer

Me.cbx_padre.BackColor = &H80000005

If cbx_padre.Text = "" Then
    ComboBox2 = Empty
End If

    
Final = GetUltimoR(Hoja29)

    For Fila = 2 To Final
        If cbx_padre.Text = Hoja29.Cells(Fila, 5) Then
            Me.ComboBox2.Text = Hoja29.Cells(Fila, 4)
            Exit For
        
        End If
    Next
End Sub
