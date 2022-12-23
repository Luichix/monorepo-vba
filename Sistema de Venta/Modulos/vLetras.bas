Attribute VB_Name = "vLetras"
'========================================================================
' Función: Convertir números a letras
'
' Creado por Otto Javier González
' www.youtube.com/ottojaviergonzalez
' Finalizado el 4 de Julio de 2013
'
' Visual Basic Para Microsoft Excel 2013
' Lista de reproducción del curso en YouTube:
' http://www.youtube.com/playlist?list=PLFNWPvtjBMjtnYLCp8KJwD1Ref7WLCIVZ
'
'========================================================================
Option Explicit
Function cMoneda(num As Double) As String
    Dim nEntero As Long
    Dim nDecimal As Double
    Dim texto As String
    Dim TipoMoneda As String
    
    TipoMoneda = Hoja94.Range("C4").Value
    
    nEntero = Int(num)
    nDecimal = Int(Round((num - nEntero) * 100)) 'Corrección de últimoo momento
    
    texto = cNumero(nEntero)
    

'Agrega los centavos
    If nDecimal <> 0 Then
        texto = texto + " Cordobas Con " + Str(nDecimal) + "/" + "100"
            
            ElseIf nDecimal = 0 Then
                texto = texto + " Cordobas Con " + "00" + "/" + "100"
                    
                    
    End If


' Agrega la moneda
    texto = texto + " " + TipoMoneda

    

    
    
cMoneda = texto
    

End Function
Function cNumero(ByVal num As Long) As String
    Dim texto As String
    
    Dim cUnidades, cDecenas, cCentenas
    Dim nUnidades, nDecenas, nCentenas As Byte
    
    Dim nMiles As Long
    Dim nMillones As Long
        
    
    cUnidades = Array("", "Uno", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve", "Diez", "Once", "Doce", "Trece", "Catorce", "Quince", "Dieciseis", "Diecisite", "Dieciocho", "Diecinueve", "Veinte", "Veintiuno", "Veintidós", "Veintitrés", "Veitnicuatro", "Veinticinco", "Veintiseis", "Veintisiete", "Veintiocho", "Veintinueve")
    cDecenas = Array("", "Diez", "Veinte", "Treinta", "Cuarenta", "Cincuenta", "Sesenta", "Setenta", "Ochenta", "Noventa", "Cien")
    cCentenas = Array("", "Ciento", "Doscientos", "Trescientos", "Cuatrocientos", "Quinientos", "Seiscientos", "Setecientos", "Ochocientos", "Novecientos")

    nMillones = num \ 1000000
    nMiles = (num \ 1000) Mod 1000
    nCentenas = (num \ 100) Mod 10
    nDecenas = (num \ 10) Mod 10
    nUnidades = num Mod 10
    
    
 'Evaluación de Millones

            If nMillones = 1 Then
                texto = "Un Millón" + IIf(num Mod 1000000 <> 0, " " + cNumero(num Mod 1000000), "")
                cNumero = texto
                Exit Function
            ElseIf nMillones >= 2 And nMillones <= 999 Then
                texto = cNumero(num \ 1000000) + " Millones" + IIf(num Mod 1000000 <> 0, " " + cNumero(num Mod 1000000), "")
                cNumero = texto
                Exit Function
            
            
 
'Evaluación de Miles

            ElseIf nMiles = 1 Then
                texto = "Mil" + IIf(num Mod 1000 <> 0, " " + cNumero(num Mod 1000), "")
                cNumero = texto
                Exit Function
            ElseIf nMiles >= 2 And nMiles <= 999 Then
                texto = cNumero(num \ 1000) + " Mil" + IIf(num Mod 1000 <> 0, " " + cNumero(num Mod 1000), "")
                cNumero = texto
                Exit Function
            
            End If
                
            
            
            
            


'Evaluación desde 0 a 999
            
            
            'Casos Especiales
            If num = 100 Then
                texto = cDecenas(10)
                cNumero = texto
                Exit Function
            ElseIf num = 0 Then
                texto = "Cero"
                cNumero = texto
                Exit Function
            End If
            
            
            
            If nCentenas <> 0 Then
                texto = cCentenas(nCentenas)
            End If
            
            
            If nDecenas <> 0 Then
                    If nDecenas = 1 Or nDecenas = 2 Then
                                If nCentenas <> 0 Then
                                    texto = texto + " "
                                End If
                        texto = texto + cUnidades(num Mod 100)
                        cNumero = texto
                        Exit Function
                Else
                
                        If nCentenas <> 0 Then
                            texto = texto + " "
                        End If
                
                texto = texto + cDecenas(nDecenas)
            End If
            End If
                
                
            If nUnidades <> 0 Then
                    If nDecenas <> 0 Then
                        texto = texto + " y "
                    ElseIf nCentenas <> 0 Then
                        texto = texto + " "
                    End If
            texto = texto + cUnidades(nUnidades)
            End If
            
            
            
            
            


cNumero = texto
End Function
