VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Nota_Credito 
   Caption         =   "NOTA DE CREDITO"
   ClientHeight    =   8910.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   8330.001
   OleObjectBlob   =   "frm_Nota_Credito.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Nota_Credito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub btn_Fecha_Click()
Me.txt_Fecha.BackColor = &H80000005
banderaCalendario = 2
  Call LanzarCalendario(Me, "txt_Fecha")
  
End Sub

Private Sub btn_personal_Click()
banderaCategoria = 2
frm_Categoria.Show
End Sub

Private Sub txt_Gravada_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_Gravada, KeyAscii)
End Sub
Private Sub txt_Exenta_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_Exenta, KeyAscii)
End Sub
Private Sub txt_Descuento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = ValidarDecimales(Me.txt_Descuento, KeyAscii)
End Sub
Private Sub btn_Salir_Click()
Unload Me
End Sub
Private Sub CommandButton3_Click()
On Error GoTo Salir
Dim Titulo As String

Titulo = "Registro de Factura"
 
If Me.txt_Fecha.Text = "" Then
    Me.txt_Fecha.BackColor = &HC0C0FF
    MsgBox "Ingrese la fecha..!", vbInformation, Titulo
    Me.btn_Fecha.SetFocus
    Exit Sub
End If

        If Me.txt_Concepto.Text = "" Then
            Me.txt_Concepto.BackColor = &HC0C0FF
            MsgBox "Seleccione un concepto del listado..!", vbInformation, Titulo
            Me.btn_personal.SetFocus
            Exit Sub
        End If
        
                          If Me.txt_Factura.Text = "" Then
                            Me.txt_Factura.BackColor = &HC0C0FF
                            MsgBox "Ingrese el numero de factura..!", vbInformation, Titulo
                            Me.txt_Factura.SetFocus
                            Exit Sub
                        End If

  Verificador

    Me.txt_Factura.SetFocus
   
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Recursos Humanos"
 End If
End Sub
Private Sub Registrar_Comision()
Dim Comprb As Long
Dim Fecha As Date
Dim Titulo As String

Titulo = "Registro de Factura"
 
    
Fecha = Me.txt_Fecha.Text

                Hoja4.Select
                Hoja4.Rows("2:2").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

                Hoja4.Cells(2, 1) = Format(Fecha, "MM/DD/YYYY")
                Hoja4.Cells(2, 2) = Me.txt_Factura.Value
                Hoja4.Cells(2, 3) = Me.txt_Descuento.Value
                Hoja4.Cells(2, 4) = Me.txt_Gravada.Value
                Hoja4.Cells(2, 5) = Me.txt_Exenta.Value
                Hoja4.Cells(2, 9) = Me.txt_Concepto

                    
         MsgBox "Registro procesado con éxito!!!", vbInformation, Titulo
             


End Sub
Private Sub Verificador()
Dim referencia As String
Dim encontrado As Boolean

Hoja4.Activate
Hoja4.Select

Hoja4.Range("B1").Select

referencia = Me.txt_Factura

Do Until IsEmpty(ActiveCell)
ActiveCell.Offset(1, 0).Select
    If ActiveCell.Value Like referencia Then
        encontrado = True
        Exit Do
    End If
Loop

If encontrado = True Then
    MsgBox "El número de nota de credito ya ha sido registrado anteriormente..!", vbCritical, "Registro"
    Exit Sub
End If

If encontrado = False Then
  
    Registrar_Comision
       LimpiarControles
    
End If


End Sub


Private Sub LimpiarControles()
Me.txt_Factura.Text = Empty
Me.txt_Gravada.Text = Empty
Me.txt_Exenta.Text = Empty
Me.txt_Descuento.Value = Empty

End Sub

