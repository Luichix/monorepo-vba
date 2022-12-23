VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Motivo_EliminarA 
   Caption         =   "CUENTA"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   9080.001
   OleObjectBlob   =   "frm_Motivo_EliminarA.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Motivo_EliminarA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_Eliminar_Click()

Dim Seguridad As String
On Error GoTo Salir

Seguridad = Hoja83.Range("L1").Text
    If Me.txt_motivo.Text = Empty Then
            MsgBox "Debe de especificar el motivo porque se elimina la cuenta...!", vbInformation, "Gestor de Recursos Humanos"
            Exit Sub
    End If
    
      If Me.opt_finalizado = False And Me.opt_anulado = False Then
            MsgBox "Debe de seleccionar una de las opciones de eliminación...!", vbInformation, "Gestor de Recursos Humanos"
            Exit Sub
    End If

    If Me.opt_finalizado = True Then
     If frm_EliminarAbono.txt_Valor_actual = 0 Then
        Else
         MsgBox "Esta cuenta aun no ha finalizado...!", vbInformation, "Gestor de Recursos Humanos"
         Exit Sub
    End If
    End If

Hoja8.Unprotect (Seguridad)
        Eliminar_Cuenta
Hoja8.Protect (Seguridad)

        
       
        With frm_EliminarAbono
            .txt_busqueda = Empty
            .txt_busqueda = "a"
            .txt_busqueda = Empty
        End With
        
       
        Unload Me
       


                    
     
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Ventas"
 End If


End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub
Private Sub Eliminar_Cuenta()
On Error Resume Next
Dim Referencia As String
Dim Inabilitar As String

Inabilitar = "ELIMINADO"


                
                Referencia = frm_EliminarAbono.txt_referencia.Text

                            Hoja8.Select
                            Range("Q1").Select
                            
                                Do Until IsEmpty(ActiveCell)
                                    ActiveCell.Offset(1, 0).Select
                                    If ActiveCell.Value Like Referencia Then
                                        encontrado = True
                                    Exit Do
                                    End If
                                Loop

                    If encontrado = True Then
                            If Me.opt_finalizado = True Then
                                        ActiveCell.Offset(0, 2).Value = Inabilitar
                                        ActiveCell.Offset(0, 3).Value = opt_finalizado.Caption & ": " & UCase(frm_Motivo_EliminarA.txt_motivo.Text)
                                        ActiveCell.Offset(0, 4).Value = Hoja83.Range("G1").Text
                                        
                            
                            ElseIf Me.opt_anulado = True Then
                                        ActiveCell.Offset(0, 2).Value = Inabilitar
                                        ActiveCell.Offset(0, 3).Value = opt_anulado.Caption & ": " & UCase(frm_Motivo_EliminarA.txt_motivo.Text)
                                        ActiveCell.Offset(0, 4).Value = Hoja83.Range("G1").Text
                            End If
                            
                            MsgBox "Registro grabado con éxito!!!", , "Gestor de Recursos Humanos"
                    End If
End Sub
