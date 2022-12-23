VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_herramienta 
   Caption         =   "Registro de Herramientas"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   10670
   OleObjectBlob   =   "frm_herramienta.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_herramienta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cancelar_Click()
Unload Me
End Sub
Private Sub btn_Fecha_Click()
 banderaCalendario = 4
    Call LanzarCalendario(Me, "txt_Fecha")
    Me.txt_cantidad.SetFocus
End Sub
Private Sub btn_herramienta_Click()
 banderaHerramienta = 1
    Call LanzarListadoHerramienta(Me, "lbl_fecha")
    Me.txt_cantidad.SetFocus
End Sub
Private Sub btn_Registrar_Click()
On Error GoTo Salir

Application.ScreenUpdating = False

    If Me.txt_Fecha.Text = Empty Or _
        Me.txt_codigo.Text = Empty Or _
        Me.txt_herramienta.Text = Empty Or _
        Me.txt_cantidad.Text = Empty Then

            MsgBox "Hay campos vacíos en el registro", , "Gestor de Inventario de Herramientas"
            Exit Sub
    
    End If
    
If MsgBox("Son correctos los datos?" + Chr(13) + "Desea procesar el registro?", vbYesNo, "Gestor de Inventarios") = vbNo Then
        Exit Sub
    Else

        ProcesarHerramienta
        ThisWorkbook.Save
        MsgBox "Datos registrados con éxito!!!", , "Gestor de Inventario de Herramientas"
        Unload Me
        
End If
    Hoja0.Activate
    Hoja0.Select
     Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventario de Herramientas"
 End If
 
End Sub

Private Sub ProcesarHerramienta()
Dim xCodigo As String
Dim xHerramienta As String
Dim xCantidad As Double
Dim xEstado As String
Dim xDetalle As String
Dim Indice As Long
                    
            Indice = Hoja5.Range("T2").Value
            Hoja3.Activate
            Hoja3.Select
                                
                    Hoja3.Rows("2:2").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                    
                    xCodigo = Me.txt_codigo.Text
                    xHerramienta = Me.txt_herramienta.Text
                    xCantidad = Me.txt_cantidad.Value
                    xEstado = "Activo"
                    xDetalle = "Bueno"
                    
                    Hoja3.Cells(2, 1) = Indice + 1
                    Hoja3.Cells(2, 2) = CDate(Me.txt_Fecha)
                    Hoja3.Cells(2, 3) = frm_detalle.txt_caja.Text
                    Hoja3.Cells(2, 4) = xCodigo
                    Hoja3.Cells(2, 5) = xHerramienta
                    Hoja3.Cells(2, 6) = xCantidad
                    Hoja3.Cells(2, 7) = xEstado
                    Hoja3.Cells(2, 8) = xDetalle
                    
                    Hoja5.Range("T2") = Indice + 1
                    
     
End Sub


Private Sub txt_cantidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii > 47 And KeyAscii < 58 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End Sub
