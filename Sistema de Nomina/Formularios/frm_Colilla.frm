VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Colilla 
   Caption         =   "Gestor de Recursos Humanos"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5120
   OleObjectBlob   =   "frm_Colilla.frx":0000
End
Attribute VB_Name = "frm_Colilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_Cargar_Click()
On Error GoTo Salir
Dim Titulo As String
Dim Seguridad As String

Seguridad = Hoja83.Range("L1").Text
Titulo = "Gestor de Recursos Humanos"

Application.Cursor = xlWait
Application.ScreenUpdating = False

    If Me.txt_Fecha.Text = Empty Then
            MsgBox "Seleccione la fecha del reporte..!", vbInformation, "Gestor de Recursos Humanos"
            Exit Sub
    End If
   

    Hoja3.Unprotect (Seguridad)
    Hoja4.Unprotect (Seguridad)
    Hoja5.Unprotect (Seguridad)
    
    Fecha_Quincena
    Reporte_General
    
    Hoja3.Protect (Seguridad)
    Hoja4.Protect (Seguridad)
    Hoja5.Protect (Seguridad)
    
    Unload Me
        
    Application.Cursor = xlDefault
     'Application.ScreenUpdating = True
                    
     
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Ventas"
 End If
End Sub


Public Sub Fecha_Quincena()
Dim Dia As String
Dim Fecha As Date
Dim Quincena As Date
Dim Seguridad As String
Dim Periodo As Date


Hoja3.Activate
Hoja3.Range("C2").Select
Hoja3.Range("C2") = CDate(Me.txt_Fecha.Text)
Periodo = Day(Hoja3.Range("C2"))

Seguridad = Hoja83.Range("L1").Text

Hoja11.Unprotect (Seguridad)

Fecha = CDate(Me.txt_Fecha.Text)
Quincena = CDate(Me.txt_Fecha.Text) + 10
If Periodo = 11 Then
    Dia = "2da "
    
Else
    Dia = "1ra "

End If


Hoja11.Range("K2") = Dia & " " & Format(Quincena, "MMMM yyyy")

Hoja11.Range("J2") = "Reporte SP, " & Dia & Format(Quincena, "MMMM yyyy")




Hoja11.Protect (Seguridad)

End Sub

Private Sub btn_Fecha_Click()

banderaPeriodo = 1
  Call LanzarPeriodo(Me, "txt_Fecha")
  Me.btn_Cargar.SetFocus
End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub


