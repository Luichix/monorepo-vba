Attribute VB_Name = "Seguridad"
Option Explicit

Public banderaUnprotect As Long
Public banderaProtect As Long

Sub Unprotec()
Dim Seguridad As String

Seguridad = Hoja83.Range("L1").Text

Select Case banderaUnprotect
    
    Case 1 'Contratación
        Hoja1.Unprotect (Seguridad)
        Hoja10.Unprotect (Seguridad)
        Hoja3.Unprotect (Seguridad)
        Hoja4.Unprotect (Seguridad)
    
    Case 2 'Contratación
        Hoja1.Unprotect (Seguridad)
        Hoja10.Unprotect (Seguridad)
        Hoja3.Unprotect (Seguridad)
        Hoja4.Unprotect (Seguridad)
        
  Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
    End Select

End Sub

Sub Protect()
Seguridad As String

Seguridad = Hoja83.Range("L1")

Select Case banderaProtect

    Case 1 'Contratación
        Hoja1.Protect (Seguridad)
        Hoja10.Protect (Seguridad)
        Hoja3.Protect (Seguridad)
        Hoja4.Protect (Seguridad)
    
    Case 2 'Contratación
        Hoja1.Protect (Seguridad)
        Hoja10.Protect (Seguridad)
        Hoja3.Protect (Seguridad)
        Hoja4.Protect (Seguridad)
        
  Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
    End Select

End Sub

