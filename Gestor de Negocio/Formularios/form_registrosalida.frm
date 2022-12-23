VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_registrosalida 
   Caption         =   "REGISTRO DE SALIDA"
   ClientHeight    =   3500
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   7080
   OleObjectBlob   =   "form_registrosalida.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "form_registrosalida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_Registrar_Click()

  Dim Fecha As String
  Dim DESCRIPCION As String
  Dim CANTIDAD As Double
  Dim COSTO As Double
  
  Fecha = Text_fecha
  DESCRIPCION = Text_descripcion
  CANTIDAD = Text_cantidad
  COSTO = Text_costo
  
    Hoja11.Select
    Hoja11.Range("A2:I2").Select
    Selection.ListObject.ListRows.Add (1)

     ActiveSheet.Cells(2, 1) = Fecha
     ActiveSheet.Cells(2, 3) = lista_area
     ActiveSheet.Cells(2, 5) = DESCRIPCION
     ActiveSheet.Cells(2, 6) = CANTIDAD
     ActiveSheet.Cells(2, 8) = COSTO
     ActiveCell = Format(Fecha, "MM/DD/YYYY")
         
       
        Text_descripcion = Empty
        Text_cantidad = Empty
        Text_costo = Empty
        
        Text_fecha.SetFocus
        
                    
End Sub





Private Sub btn_Salir_Click()
End

End Sub

Private Sub Text_cantidad_Change()

End Sub

Private Sub Text_fecha_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)


Select Case Len(Text_fecha.Value)
Case 2
Text_fecha.Value = Text_fecha.Value & "/"
Case 5
Text_fecha.Value = Text_fecha.Value & "/"

End Select


End Sub



