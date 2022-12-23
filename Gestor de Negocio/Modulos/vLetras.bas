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
    Dim Texto As String
    Dim TipoMoneda As String
    
    TipoMoneda = Hoja27.Range("C4").Value
    
    nEntero = Int(num)
    nDecimal = Int(Round((num - nEntero) * 100)) 'Corrección de últimoo momento
    
    Texto = cNumero(nEntero)
    

'Agrega los centavos
    If nDecimal <> 0 Then
        Texto = Texto + " Cordobas Con " + Str(nDecimal) + "/" + "100"
            
            ElseIf nDecimal = 0 Then
                Texto = Texto + " Cordobas Con " + "00" + "/" + "100"
                    
                    
    End If


' Agrega la moneda
    Texto = Texto + " " + TipoMoneda

    

    
    
cMoneda = Texto
    

End Function
Function cNumero(ByVal num As Long) As String
    Dim Texto As String
    
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
                Texto = "Un Millón" + IIf(num Mod 1000000 <> 0, " " + cNumero(num Mod 1000000), "")
                cNumero = Texto
                Exit Function
            ElseIf nMillones >= 2 And nMillones <= 999 Then
                Texto = cNumero(num \ 1000000) + " Millones" + IIf(num Mod 1000000 <> 0, " " + cNumero(num Mod 1000000), "")
                cNumero = Texto
                Exit Function
            
            
 
'Evaluación de Miles

            ElseIf nMiles = 1 Then
                Texto = "Mil" + IIf(num Mod 1000 <> 0, " " + cNumero(num Mod 1000), "")
                cNumero = Texto
                Exit Function
            ElseIf nMiles >= 2 And nMiles <= 999 Then
                Texto = cNumero(num \ 1000) + " Mil" + IIf(num Mod 1000 <> 0, " " + cNumero(num Mod 1000), "")
                cNumero = Texto
                Exit Function
            
            End If
                
            
            
            
            


'Evaluación desde 0 a 999
            
            
            'Casos Especiales
            If num = 100 Then
                Texto = cDecenas(10)
                cNumero = Texto
                Exit Function
            ElseIf num = 0 Then
                Texto = "Cero"
                cNumero = Texto
                Exit Function
            End If
            
            
            
            If nCentenas <> 0 Then
                Texto = cCentenas(nCentenas)
            End If
            
            
            If nDecenas <> 0 Then
                    If nDecenas = 1 Or nDecenas = 2 Then
                                If nCentenas <> 0 Then
                                    Texto = Texto + " "
                                End If
                        Texto = Texto + cUnidades(num Mod 100)
                        cNumero = Texto
                        Exit Function
                Else
                
                        If nCentenas <> 0 Then
                            Texto = Texto + " "
                        End If
                
                Texto = Texto + cDecenas(nDecenas)
            End If
            End If
                
                
            If nUnidades <> 0 Then
                    If nDecenas <> 0 Then
                        Texto = Texto + " y "
                    ElseIf nCentenas <> 0 Then
                        Texto = Texto + " "
                    End If
            Texto = Texto + cUnidades(nUnidades)
            End If
            
            
            
            
            


cNumero = Texto
End Function
