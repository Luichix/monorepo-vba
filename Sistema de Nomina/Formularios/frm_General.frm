VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_General 
   Caption         =   "Gestor de Recursos Humanos"
   ClientHeight    =   3705
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5520
   OleObjectBlob   =   "frm_General.frx":0000
End
Attribute VB_Name = "frm_General"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub btn_Cargar_Click()
On Error GoTo Salir
Dim Titulo As String
Dim Seguridad As String

Seguridad = Hoja83.Range("L1").Text

Titulo = "Gestor de Recursos Humanos"

 
Application.ScreenUpdating = False

    If Me.txt_Fecha = Empty Then
            MsgBox "Debe seleccionar la fecha del reporte..!", vbInformation, "Gestor de Recursos Humanos"
            Exit Sub
    End If
    
   
   
    Hoja3.Unprotect (Seguridad)
    Hoja4.Unprotect (Seguridad)
    Hoja5.Unprotect (Seguridad)

 MsgBox "Espere un momento... Click para continuar..."
      
 Application.Cursor = xlWait
               

    General_Quincena
    Cargar_Todo
    
 Application.Cursor = xlDefault
 
MsgBox "Comprobante de pago elaborado con éxito!!!", vbInformation, Titulo
    

    Hoja3.Protect (Seguridad)
    Hoja4.Protect (Seguridad)
    Hoja5.Protect (Seguridad)

    
     Application.ScreenUpdating = True
                    
     Unload Me
     
Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Ventas"
 End If
End Sub

Public Sub Cargar_Todo()
Dim Fila As Long
Dim Titulo As String
Dim Referencia As String
Dim encontrado As Boolean
Dim xFila As Long
Dim xFinal As Long
'Dim Empresa As String

'Empresa = "FLORENCIA INTERCOMERCIAL, S.A."

'VARIABLES DE TEXTO
    
Dim C1_Colilla As String
Dim C2_Datogeneral As String
Dim C3_Id As String
Dim C4_Colabo As String
Dim C5_Cedu As String
Dim C6_Cargo As String
Dim C7_Contrato As String
Dim C8_DataPago As String
Dim C9_Periodo As String
Dim C10_SalarioM As String
Dim C11_SalarioQ As String
'Dim C12_SalarioH As String
Dim C13_DatoLabores As String
Dim C14_Base As String
Dim C15_Ausencia As String
Dim C16_Ordinario As String
Dim C17_IngresoBruto As String
Dim C18_Deduce As String
Dim C19_Ingreso As String
Dim C20_Egreso As String
Dim C21_Horas As String
'Dim C22_Valor As String
Dim C23_Saldo As String
Dim C24_Cuota As String
Dim C25_Ordinal As String
Dim C26_Extra As String
Dim C27_Vesper1 As String
Dim C28_Vesper2 As String
Dim C29_Nocturne As String
Dim C30_Comision As String
Dim C31_PrestamoP As String
Dim C32_PrestamoB As String
Dim C33_Vales As String
Dim C34_Florencia As String
Dim C35_Otros As String
Dim C36_SeguroS As String
Dim C37_SeguroE As String
''''''''''''''''''''''''''''''''Dim C38_IR As String
Dim C39_TIngreso As String
Dim C40_TEgreso As String
Dim C41_NPago As String
Dim C42_Conforme As String
Dim C43_Ajuste As String
Dim C44_Viatico As String


    'TEXTO DE COLILLA

C1_Colilla = "COMPROBANTE DE PAGO"
C2_Datogeneral = "DATOS GENERALES"
C3_Id = "ID PERSONAL:"
C4_Colabo = "COLABORADOR:"
C5_Cedu = "CEDULA DE IDENTIDAD:"
C6_Cargo = "CARGO EN LA ENTIDAD:"
C7_Contrato = "TIPO DE CONTRATO:"
C8_DataPago = "DATOS DE PAGO"
C9_Periodo = "PERIODO DE PAGO:"
C10_SalarioM = "SALARIO MENSUAL:"
C11_SalarioQ = "SALARIO QUINCENAL:"
'C12_SalarioH = "SALARIO POR HORA:"
C13_DatoLabores = "DETALLE DE LABORES"
C14_Base = "(+)BASE SALARIAL:"
C15_Ausencia = "(-)AUSENCIA/TARDANZA:"
C16_Ordinario = "(=)INGRESO ORDINARIO:"
C17_IngresoBruto = "INGRESOS BRUTOS"
C18_Deduce = "DEDUCCIÓNES"
C19_Ingreso = "INGRESO"
C20_Egreso = "EGRESO"
C21_Horas = "HORAS"
'C22_Valor = "VALOR"
C23_Saldo = "SALDO"
C24_Cuota = "CUOTA"
C25_Ordinal = "INGRESO ORDINARIO:"
C26_Extra = "EXTRAS DIURNAS:"
C27_Vesper1 = "EXTRAS 5PM - 6PM:"
C28_Vesper2 = "EXTRAS 6PM - 8PM:"
C29_Nocturne = "EXTRAS 8+:"
C30_Comision = "COMISIÓNES:"
C31_PrestamoP = "PRESTAMOS PERSONALES:"
C32_PrestamoB = "PRESTAMOS BANCARIOS:"
C33_Vales = "VALES:"
C34_Florencia = "VENTAS FLORENCIA:"
C35_Otros = "OTRAS DEDUCCIONES:"
C36_SeguroS = "SEGURO SOCIAL:"
C37_SeguroE = "SEGURO EDUCATIVO:"
'''''''''''''''''''''''''''''''''''''''''''''''''C38_IR = "IMPUESTO S/RENTA:"
C39_TIngreso = "INGRESOS BRUTOS:"
C40_TEgreso = "TOTAL EGRESOS:"
C41_NPago = "NETO A PAGAR:"
C42_Conforme = "RECIBI CONFORME"
C43_Ajuste = "AJUSTE:"
C44_Viatico = "VIATICO:"



'VARIABLES DE VALORES

Dim Id_Personal As String
Dim Colaborador As String
Dim Cedula As String
Dim Cargo As String
Dim Contrato As String
Dim Periodo As String
Dim SalarioMensual As Currency
Dim SalarioQuincenal As Currency
'Dim SalarioHora As Currency
Dim IngresoBase As Currency
Dim IngresoOrdinario As Currency
Dim HAusencia As Date
'Dim VAusencia As Currency
Dim EAusencia As Currency
Dim HDiurnas As Date
'Dim VDiurnas As Currency
Dim IDiurnas As Currency
Dim HVesper1 As Date
'Dim VVesper1 As Currency
Dim IVesper1 As Currency
Dim HVesper2 As Date
'Dim VVesper2 As Currency
Dim IVesper2 As Currency
Dim HNocturne As Date
'Dim VNocturne As Currency
Dim INocturne As Currency
Dim SPrestamoP As Currency
Dim CPrestamoP As String
Dim EPrestamoP As Currency
Dim SPrestamoB As Currency
Dim CPrestamoB As String
Dim EPrestamoB As Currency
Dim SVale As Currency
Dim CVale As String
Dim EVale As Currency
Dim SVentaF As Currency
Dim CVentaF As String
Dim EVentaF As Currency
Dim SOtroD As Currency
Dim COtroD As String
Dim EOtroD As Currency
Dim Comision As Currency
Dim Ajuste As Currency
Dim Viatico As Currency
Dim SeguroSocial As Currency
Dim SeguroEducativo As Currency
'''''''''''''''''''''''''''''''''''''''''''Dim IR As Currency
Dim TIngresos As Currency
Dim TEgresos As Currency
Dim NPago As Currency



    'TEXTO DE COLILLA
Application.ScreenUpdating = False
Application.EnableEvents = False


    'LIMPIAR HOJA
    Hoja5.Activate

With Hoja5.PageSetup

 .TopMargin = Application.InchesToPoints(0.764)
 .BottomMargin = Application.InchesToPoints(0.764)
 .LeftMargin = Application.InchesToPoints(0.4)
 .RightMargin = Application.InchesToPoints(0.24)
 .HeaderMargin = Application.InchesToPoints(0.32)
 .FooterMargin = Application.InchesToPoints(0.32)
End With


    
    Hoja5.Cells.Select
    Selection.Clear
    
    'FORMATO LETRA CALIBRI 9
    With Selection.Font
        .Name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With


Hoja4.Activate
Hoja4.Cells(5, 1).Select

xFila = 5

                    Do While Hoja4.Cells(xFila, 1) <> Empty
                        xFila = xFila + 1
                    Loop
                    xFinal = xFila - 1


Fila = 1

For xFila = 5 To xFinal

   
    'INGRESO DE CAMPOS
    Hoja5.Activate
        Hoja5.Cells(1, 1).Select

    Hoja5.Cells(Fila, 1) = C1_Colilla
    Hoja5.Cells(Fila + 1, 1) = C2_Datogeneral
    Hoja5.Cells(Fila + 2, 1) = C3_Id
    Hoja5.Cells(Fila + 3, 1) = C4_Colabo
    Hoja5.Cells(Fila + 4, 1) = C5_Cedu
    Hoja5.Cells(Fila + 5, 1) = C6_Cargo
    
    'Hoja5.Cells(Fila + 6, 1) = C7_Contrato
       
    Hoja5.Cells(Fila + 1, 5) = C8_DataPago
    Hoja5.Cells(Fila + 2, 5) = C9_Periodo
    Hoja5.Cells(Fila + 4, 5) = C10_SalarioM
    Hoja5.Cells(Fila + 5, 5) = C11_SalarioQ
    'Hoja5.Cells(Fila + 6, 5) = C12_SalarioH
    
    Hoja5.Cells(Fila + 7, 1) = C13_DatoLabores
    Hoja5.Cells(Fila + 8, 1) = C14_Base
    Hoja5.Cells(Fila + 9, 1) = C15_Ausencia
    Hoja5.Cells(Fila + 8, 5) = C16_Ordinario
    
    Hoja5.Cells(Fila + 10, 1) = C17_IngresoBruto
    Hoja5.Cells(Fila + 10, 5) = C18_Deduce
    
    Hoja5.Cells(Fila + 10, 4) = C19_Ingreso
    Hoja5.Cells(Fila + 10, 8) = C20_Egreso
    Hoja5.Cells(Fila + 10, 2) = C21_Horas
    'Hoja5.Cells(Fila + 10, 3) = C22_Valor
    Hoja5.Cells(Fila + 10, 6) = C23_Saldo
    Hoja5.Cells(Fila + 10, 7) = C24_Cuota
    
    
    Hoja5.Cells(Fila + 11, 1) = C25_Ordinal
    Hoja5.Cells(Fila + 12, 1) = C26_Extra
    Hoja5.Cells(Fila + 13, 1) = C27_Vesper1
    Hoja5.Cells(Fila + 14, 1) = C28_Vesper2
    Hoja5.Cells(Fila + 15, 1) = C29_Nocturne
    Hoja5.Cells(Fila + 16, 1) = C30_Comision
    Hoja5.Cells(Fila + 17, 1) = C43_Ajuste
    
             
    Hoja5.Cells(Fila + 11, 5) = C31_PrestamoP
    Hoja5.Cells(Fila + 12, 5) = C32_PrestamoB
    Hoja5.Cells(Fila + 13, 5) = C33_Vales
    Hoja5.Cells(Fila + 14, 5) = C34_Florencia
    Hoja5.Cells(Fila + 15, 5) = C35_Otros
    Hoja5.Cells(Fila + 16, 5) = C44_Viatico
    Hoja5.Cells(Fila + 17, 5) = C36_SeguroS
    Hoja5.Cells(Fila + 18, 5) = C37_SeguroE
    '''''''''''''''''''''''''''''''''''''''''''''Hoja5.Cells(Fila + 19, 5) = C38_IR
    
    Hoja5.Cells(Fila + 20, 1) = C39_TIngreso
    Hoja5.Cells(Fila + 20, 5) = C40_TEgreso
    Hoja5.Cells(Fila + 21, 1) = C41_NPago
    Hoja5.Cells(Fila + 24, 1) = C42_Conforme
    
   
'DISEÑO Y FORMATO DE LA ESTRUCTURA "COLOR Y CENTRADO

Hoja5.Range(Cells(Fila, 1), Cells(Fila, 8)).Select
    Diseño_A
    Formato_A

Hoja5.Range(Cells(Fila + 1, 1), Cells(Fila + 1, 4)).Select
    Diseño_A
    Formato_A

Hoja5.Range(Cells(Fila + 1, 5), Cells(Fila + 1, 8)).Select
    Diseño_A
    Formato_A
    
Hoja5.Range(Cells(Fila + 2, 5), Cells(Fila + 3, 5)).Select
    Diseño_D

Hoja5.Range(Cells(Fila + 2, 6), Cells(Fila + 3, 8)).Select
    Diseño_A
    
Hoja5.Range(Cells(Fila + 7, 1), Cells(Fila + 7, 8)).Select
    Diseño_A
    Formato_A

Hoja5.Range(Cells(Fila + 10, 1), Cells(Fila + 10, 8)).Select
    Diseño_B
    Formato_A

Hoja5.Range(Cells(Fila + 20, 1), Cells(Fila + 21, 8)).Select
    Formato_A

Hoja5.Range(Cells(Fila + 20, 1), Cells(Fila + 20, 2)).Select
    Diseño_C
    
Hoja5.Range(Cells(Fila + 20, 3), Cells(Fila + 20, 4)).Select
    Diseño_C

Hoja5.Range(Cells(Fila + 20, 5), Cells(Fila + 20, 6)).Select
    Diseño_C
    
Hoja5.Range(Cells(Fila + 20, 7), Cells(Fila + 20, 8)).Select
    Diseño_C

Hoja5.Range(Cells(Fila + 21, 1), Cells(Fila + 21, 4)).Select
    Diseño_C
 
Hoja5.Range(Cells(Fila + 24, 1), Cells(Fila + 24, 8)).Select
    Diseño_A

'DISEÑO Y FORMATO DE ESTRUCTURA "BORDES"

Hoja5.Range(Cells(Fila, 1), Cells(Fila, 8)).Select
    Borde_A
    Borde_B

Hoja5.Range(Cells(Fila + 1, 1), Cells(Fila + 1, 8)).Select
    Borde_B
  
Hoja5.Range(Cells(Fila + 7, 1), Cells(Fila + 7, 8)).Select
    Borde_A
    Borde_B

Hoja5.Range(Cells(Fila + 10, 1), Cells(Fila + 10, 8)).Select
    Borde_A
    Borde_B
    
Hoja5.Range(Cells(Fila + 20, 1), Cells(Fila + 20, 8)).Select
    Borde_A
    Borde_B

Hoja5.Range(Cells(Fila + 21, 1), Cells(Fila + 21, 8)).Select
    Borde_B

Hoja5.Range(Cells(Fila + 1, 5), Cells(Fila + 6, 8)).Select
    Borde_I
    
Hoja5.Range(Cells(Fila + 8, 1), Cells(Fila + 21, 4)).Select
    Borde_D
    
Hoja5.Range(Cells(Fila + 23, 3), Cells(Fila + 23, 5)).Select
    Borde_B


'LLENADO DE DATOS DE COLILLA "DATOS DE VARIABLES"

Referencia = Hoja4.Cells(xFila, 1)


Hoja4.Select
Range("A4").Select

    Periodo = Hoja4.Cells(2, 1).Value

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like Referencia Then
            encontrado = True
            Id_Personal = ActiveCell.Offset(0, 0).Value
            Colaborador = ActiveCell.Offset(0, 1).Value
            Cedula = ActiveCell.Offset(0, 2).Value
            Cargo = ActiveCell.Offset(0, 3).Value
            ''Contrato = ActiveCell.Offset(0, 6).Value
            SalarioMensual = ActiveCell.Offset(0, 10).Value
            SalarioQuincenal = ActiveCell.Offset(0, 11).Value
'            SalarioHora = ActiveCell.Offset(0, 12).Value
            IngresoBase = ActiveCell.Offset(0, 26).Value
            HAusencia = ActiveCell.Offset(0, 21).Value
'            VAusencia = ActiveCell.Offset(0, 12).Value
            EAusencia = ActiveCell.Offset(0, 27).Value
            IngresoOrdinario = ActiveCell.Offset(0, 28).Value
            HDiurnas = ActiveCell.Offset(0, 22).Value
'            VDiurnas = ActiveCell.Offset(0, 13).Value
            IDiurnas = ActiveCell.Offset(0, 29).Value
            HVesper1 = ActiveCell.Offset(0, 23).Value
'            VVesper1 = ActiveCell.Offset(0, 14).Value
            IVesper1 = ActiveCell.Offset(0, 30).Value
            HVesper2 = ActiveCell.Offset(0, 24).Value
'            VVesper2 = ActiveCell.Offset(0, 15).Value
            IVesper2 = ActiveCell.Offset(0, 31).Value
            HNocturne = ActiveCell.Offset(0, 25).Value
'            VNocturne = ActiveCell.Offset(0, 16).Value
            INocturne = ActiveCell.Offset(0, 32).Value
            SPrestamoP = ActiveCell.Offset(0, 46).Value
            CPrestamoP = "'" & ActiveCell.Offset(0, 47).Value
            EPrestamoP = ActiveCell.Offset(0, 36).Value
            SPrestamoB = ActiveCell.Offset(0, 48).Value
            CPrestamoB = "'" & ActiveCell.Offset(0, 49).Value
            EPrestamoB = ActiveCell.Offset(0, 37).Value
            SVale = ActiveCell.Offset(0, 50).Value
            CVale = "'" & ActiveCell.Offset(0, 51).Value
            EVale = ActiveCell.Offset(0, 38).Value
            SVentaF = ActiveCell.Offset(0, 52).Value
            CVentaF = "'" & ActiveCell.Offset(0, 53).Value
            EVentaF = ActiveCell.Offset(0, 39).Value
            SOtroD = ActiveCell.Offset(0, 54).Value
            COtroD = "'" & ActiveCell.Offset(0, 55).Value
            EOtroD = ActiveCell.Offset(0, 40).Value
            Comision = ActiveCell.Offset(0, 33).Value
            Ajuste = ActiveCell.Offset(0, 34).Value
            Viatico = ActiveCell.Offset(0, 41).Value
            SeguroSocial = ActiveCell.Offset(0, 42).Value
            SeguroEducativo = ActiveCell.Offset(0, 43).Value
            ''''''''''''''''''''''''''''''''''''''''''''''''''''IR = ActiveCell.Offset(0, 45).Value
            TIngresos = ActiveCell.Offset(0, 35).Value
            TEgresos = ActiveCell.Offset(0, 44).Value
            NPago = ActiveCell.Offset(0, 45).Value
            Exit Do
        End If
    Loop
            
'LLENADO DE DATOS DE COLILLA "INGRESO DE DATOS"

    Hoja5.Activate

    Hoja5.Cells(Fila + 2, 2) = Id_Personal
    Hoja5.Cells(Fila + 3, 2) = Colaborador
    Hoja5.Cells(Fila + 4, 2) = Cedula
    Hoja5.Cells(Fila + 5, 2) = Cargo
    'Hoja5.Cells(Fila + 6, 2) = Contrato
    
    Hoja5.Cells(Fila + 2, 6) = Periodo
    Hoja5.Cells(Fila + 4, 6) = SalarioMensual
    Hoja5.Cells(Fila + 5, 6) = SalarioQuincenal
    'Hoja5.Cells(Fila + 6, 6) = SalarioHora
    
    Hoja5.Cells(Fila + 8, 4) = IngresoBase
    Hoja5.Cells(Fila + 9, 2) = HAusencia
    'Hoja5.Cells(Fila + 9, 3) = VAusencia
    Hoja5.Cells(Fila + 9, 4) = EAusencia
    Hoja5.Cells(Fila + 8, 8) = IngresoOrdinario
    
    Hoja5.Cells(Fila + 11, 4) = IngresoOrdinario
        
    Hoja5.Cells(Fila + 12, 2) = HDiurnas
   ' Hoja5.Cells(Fila + 12, 3) = VDiurnas
    Hoja5.Cells(Fila + 12, 4) = IDiurnas
    
    Hoja5.Cells(Fila + 13, 2) = HVesper1
    'Hoja5.Cells(Fila + 13, 3) = VVesper1
    Hoja5.Cells(Fila + 13, 4) = IVesper1
    
    Hoja5.Cells(Fila + 14, 2) = HVesper2
    'Hoja5.Cells(Fila + 14, 3) = VVesper2
    Hoja5.Cells(Fila + 14, 4) = IVesper2
    
    Hoja5.Cells(Fila + 15, 2) = HNocturne
    'Hoja5.Cells(Fila + 15, 3) = VNocturne
    Hoja5.Cells(Fila + 15, 4) = INocturne
    
    Hoja5.Cells(Fila + 16, 4) = Comision
    Hoja5.Cells(Fila + 17, 4) = Ajuste
    
    Hoja5.Cells(Fila + 11, 6) = SPrestamoP
    Hoja5.Cells(Fila + 11, 7) = CPrestamoP
    Hoja5.Cells(Fila + 11, 8) = EPrestamoP
    
    Hoja5.Cells(Fila + 12, 6) = SPrestamoB
    Hoja5.Cells(Fila + 12, 7) = CPrestamoB
    Hoja5.Cells(Fila + 12, 8) = EPrestamoB
    
    Hoja5.Cells(Fila + 13, 6) = SVale
    Hoja5.Cells(Fila + 13, 7) = CVale
    Hoja5.Cells(Fila + 13, 8) = EVale
        
    Hoja5.Cells(Fila + 14, 6) = SVentaF
    Hoja5.Cells(Fila + 14, 7) = CVentaF
    Hoja5.Cells(Fila + 14, 8) = EVentaF
    
    Hoja5.Cells(Fila + 15, 6) = SOtroD
    Hoja5.Cells(Fila + 15, 7) = COtroD
    Hoja5.Cells(Fila + 15, 8) = EOtroD
    
    Hoja5.Cells(Fila + 16, 8) = Viatico
    Hoja5.Cells(Fila + 17, 8) = SeguroSocial
    Hoja5.Cells(Fila + 18, 8) = SeguroEducativo
    ''''''''''''''''''''''''''''''''''''''''''''''''''''Hoja5.Cells(Fila + 19, 8) = IR
    
    Hoja5.Cells(Fila + 20, 3) = TIngresos
    Hoja5.Cells(Fila + 20, 7) = TEgresos
    Hoja5.Cells(Fila + 21, 5) = NPago
    
    
'FORMATO DE DATOS DE VARIABLES

    Hoja5.Cells(Fila + 2, 2).HorizontalAlignment = xlLeft
    Hoja5.Cells(Fila + 3, 2).HorizontalAlignment = xlLeft
    Hoja5.Cells(Fila + 4, 2).HorizontalAlignment = xlLeft
    Hoja5.Cells(Fila + 5, 2).HorizontalAlignment = xlLeft
    Hoja5.Cells(Fila + 6, 2).HorizontalAlignment = xlLeft
    
    Hoja5.Cells(Fila + 2, 6).WrapText = True
    Hoja5.Cells(Fila + 2, 6).Font.Size = 7
    Hoja5.Cells(Fila + 4, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 5, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 6, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 8, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 9, 2).NumberFormat = "[hh]:mm"
    Hoja5.Cells(Fila + 9, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 9, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 8, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 11, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 12, 2).NumberFormat = "[hh]:mm"
    Hoja5.Cells(Fila + 12, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 12, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 13, 2).NumberFormat = "[hh]:mm"
    Hoja5.Cells(Fila + 13, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 13, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 14, 2).NumberFormat = "[hh]:mm"
    Hoja5.Cells(Fila + 14, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 14, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 15, 2).NumberFormat = "[hh]:mm"
    Hoja5.Cells(Fila + 15, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 15, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 16, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 17, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 11, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 11, 7).HorizontalAlignment = xlCenter
    Hoja5.Cells(Fila + 11, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 12, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 12, 7).HorizontalAlignment = xlCenter
    Hoja5.Cells(Fila + 12, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 13, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 13, 7).HorizontalAlignment = xlCenter
    Hoja5.Cells(Fila + 13, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 14, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 14, 7).HorizontalAlignment = xlCenter
    Hoja5.Cells(Fila + 14, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        
    Hoja5.Cells(Fila + 15, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 15, 7).HorizontalAlignment = xlCenter
    Hoja5.Cells(Fila + 15, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
       
    Hoja5.Cells(Fila + 16, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 17, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 18, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 19, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 20, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 20, 7).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 21, 5).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    
'DUPLICADO DE COLILLA
'Hoja5.Activate
'
'Hoja5.Cells(Fila + 29, 1) = Empresa
'
'Hoja5.Range(Cells(Fila + 29, 1), Cells(Fila + 29, 8)).Select
'Diseño_A
'
'Hoja5.Range(Cells(Fila, 1), Cells(Fila + 24, 8)).Select
'Hoja5.Range(Cells(Fila, 1), Cells(Fila + 24, 8)).Copy
'
'Hoja5.Range(Cells(Fila + 30, 1), Cells(Fila + 30, 1)).Select
'    ActiveSheet.Paste
'    Application.CutCopyMode = False
        
  Fila = Fila + 26
  
    
Next

Hoja5.Activate

Application.ScreenUpdating = True
Application.EnableEvents = True

    Unload Me
End Sub

Public Sub Cargar_PDF()
Dim Fila As Long
Dim Titulo As String
Dim Referencia As String
Dim encontrado As Boolean
Dim xFila As Long
Dim xFinal As Long
Dim Validador As String
Dim Contrato_General As String
Dim Empresa As String

Empresa = "FLORENCIA INTERCOMERCIAL, S.A."
Contrato_General = "PLANILLA EN GENERAL"

'VARIABLES DE TEXTO
    
Dim C1_Colilla As String
Dim C2_Datogeneral As String
Dim C3_Id As String
Dim C4_Colabo As String
Dim C5_Cedu As String
Dim C6_Cargo As String
Dim C7_Contrato As String
Dim C8_DataPago As String
Dim C9_Periodo As String
Dim C10_SalarioM As String
Dim C11_SalarioQ As String
Dim C12_SalarioH As String
Dim C13_DatoLabores As String
Dim C14_Base As String
Dim C15_Ausencia As String
Dim C16_Ordinario As String
Dim C17_IngresoBruto As String
Dim C18_Deduce As String
Dim C19_Ingreso As String
Dim C20_Egreso As String
Dim C21_Horas As String
Dim C22_Valor As String
Dim C23_Saldo As String
Dim C24_Cuota As String
Dim C25_Ordinal As String
Dim C26_Extra As String
Dim C27_Vesper1 As String
Dim C28_Vesper2 As String
Dim C29_Nocturne As String
Dim C30_Comision As String
Dim C31_PrestamoP As String
Dim C32_PrestamoB As String
Dim C33_Vales As String
Dim C34_Florencia As String
Dim C35_Otros As String
Dim C36_SeguroS As String
Dim C37_SeguroE As String
''''''''''''''''''''''''''''''''''''''''''''''''''''Dim C38_IR As String
Dim C39_TIngreso As String
Dim C40_TEgreso As String
Dim C41_NPago As String
Dim C42_Conforme As String
Dim C43_Ajuste As String
Dim C44_Viatico As String


    'TEXTO DE COLILLA

C1_Colilla = "COMPROBANTE DE PAGO"
C2_Datogeneral = "DATOS GENERALES"
C3_Id = "ID PERSONAL:"
C4_Colabo = "COLABORADOR:"
C5_Cedu = "CEDULA DE IDENTIDAD:"
C6_Cargo = "CARGO EN LA ENTIDAD:"
C7_Contrato = "TIPO DE CONTRATO:"
C8_DataPago = "DATOS DE PAGO"
C9_Periodo = "PERIODO DE PAGO:"
C10_SalarioM = "SALARIO MENSUAL:"
C11_SalarioQ = "SALARIO QUINCENAL:"
C12_SalarioH = "SALARIO POR HORA:"
C13_DatoLabores = "DETALLE DE LABORES"
C14_Base = "(+)BASE SALARIAL:"
C15_Ausencia = "(-)AUSENCIA/TARDANZA:"
C16_Ordinario = "(=)INGRESO ORDINARIO:"
C17_IngresoBruto = "INGRESOS BRUTOS"
C18_Deduce = "DEDUCCIÓNES"
C19_Ingreso = "INGRESO"
C20_Egreso = "EGRESO"
C21_Horas = "HORAS"
C22_Valor = "VALOR"
C23_Saldo = "SALDO"
C24_Cuota = "CUOTA"
C25_Ordinal = "INGRESO ORDINARIO:"
C26_Extra = "EXTRAS DIURNAS:"
C27_Vesper1 = "EXTRAS 5PM - 6PM:"
C28_Vesper2 = "EXTRAS 6PM - 8PM:"
C29_Nocturne = "EXTRAS 8+:"
C30_Comision = "COMISIÓNES:"
C31_PrestamoP = "PRESTAMOS PERSONALES:"
C32_PrestamoB = "PRESTAMOS BANCARIOS:"
C33_Vales = "VALES:"
C34_Florencia = "VENTAS FLORENCIA:"
C35_Otros = "OTRAS DEDUCCIONES:"
C36_SeguroS = "SEGURO SOCIAL:"
C37_SeguroE = "SEGURO EDUCATIVO:"
''''''''''''''''''''''''''''''''''''''''''''''''''''C38_IR = "IMPUESTO S/RENTA:"
C39_TIngreso = "INGRESOS BRUTOS:"
C40_TEgreso = "TOTAL EGRESOS:"
C41_NPago = "NETO A PAGAR:"
C42_Conforme = "RECIBI CONFORME"
C43_Ajuste = "AJUSTE:"
C44_Viatico = "VIATICO:"



'VARIABLES DE VALORES

Dim Id_Personal As String
Dim Colaborador As String
Dim Cedula As String
Dim Cargo As String
Dim Contrato As String
Dim Periodo As String
Dim SalarioMensual As Currency
Dim SalarioQuincenal As Currency
Dim SalarioHora As Currency
Dim IngresoBase As Currency
Dim IngresoOrdinario As Currency
Dim HAusencia As Date
Dim VAusencia As Currency
Dim EAusencia As Currency
Dim HDiurnas As Date
Dim VDiurnas As Currency
Dim IDiurnas As Currency
Dim HVesper1 As Date
Dim VVesper1 As Currency
Dim IVesper1 As Currency
Dim HVesper2 As Date
Dim VVesper2 As Currency
Dim IVesper2 As Currency
Dim HNocturne As Date
Dim VNocturne As Currency
Dim INocturne As Currency
Dim SPrestamoP As Currency
Dim CPrestamoP As String
Dim EPrestamoP As Currency
Dim SPrestamoB As Currency
Dim CPrestamoB As String
Dim EPrestamoB As Currency
Dim SVale As Currency
Dim CVale As String
Dim EVale As Currency
Dim SVentaF As Currency
Dim CVentaF As String
Dim EVentaF As Currency
Dim SOtroD As Currency
Dim COtroD As String
Dim EOtroD As Currency
Dim Comision As Currency
Dim Ajuste As Currency
Dim Viatico As Currency
Dim SeguroSocial As Currency
Dim SeguroEducativo As Currency
''''''''''''''''''''''''''''''''''''''''''''''''''''Dim IR As Currency
Dim TIngresos As Currency
Dim TEgresos As Currency
Dim NPago As Currency



    'TEXTO DE COLILLA
Application.ScreenUpdating = False
Application.EnableEvents = False


    'LIMPIAR HOJA
    Hoja5.Activate

With Hoja5.PageSetup

 .TopMargin = Application.InchesToPoints(0.764)
 .BottomMargin = Application.InchesToPoints(0.764)
 .LeftMargin = Application.InchesToPoints(0.4)
 .RightMargin = Application.InchesToPoints(0.24)
 .HeaderMargin = Application.InchesToPoints(0.32)
 .FooterMargin = Application.InchesToPoints(0.32)
End With

    
    Hoja5.Cells.Select
    Selection.Clear
    
    'FORMATO LETRA CALIBRI 9
    With Selection.Font
        .Name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With


Hoja4.Activate
Hoja4.Cells(5, 1).Select

xFila = 5

                    Do While Hoja4.Cells(xFila, 1) <> Empty
                        xFila = xFila + 1
                    Loop
                    xFinal = xFila - 1


Fila = 1



For xFila = 5 To xFinal

  'INGRESO DE CAMPOS
    Hoja5.Activate
        Hoja5.Cells(1, 1).Select

    Hoja5.Cells(Fila, 1) = C1_Colilla
    
    Hoja5.Cells(Fila + 1, 1) = C2_Datogeneral
    Hoja5.Cells(Fila + 2, 1) = C3_Id
    Hoja5.Cells(Fila + 3, 1) = C4_Colabo
    Hoja5.Cells(Fila + 4, 1) = C5_Cedu
    Hoja5.Cells(Fila + 5, 1) = C6_Cargo
    'Hoja5.Cells(Fila + 6, 1) = C7_Contrato
       
    Hoja5.Cells(Fila + 1, 5) = C8_DataPago
    Hoja5.Cells(Fila + 2, 5) = C9_Periodo
    Hoja5.Cells(Fila + 4, 5) = C10_SalarioM
    Hoja5.Cells(Fila + 5, 5) = C11_SalarioQ
    Hoja5.Cells(Fila + 6, 5) = C12_SalarioH
    
    Hoja5.Cells(Fila + 7, 1) = C13_DatoLabores
    Hoja5.Cells(Fila + 8, 1) = C14_Base
    Hoja5.Cells(Fila + 9, 1) = C15_Ausencia
    Hoja5.Cells(Fila + 8, 5) = C16_Ordinario
    
    Hoja5.Cells(Fila + 10, 1) = C17_IngresoBruto
    Hoja5.Cells(Fila + 10, 5) = C18_Deduce
    
    Hoja5.Cells(Fila + 10, 4) = C19_Ingreso
    Hoja5.Cells(Fila + 10, 8) = C20_Egreso
    Hoja5.Cells(Fila + 10, 2) = C21_Horas
    Hoja5.Cells(Fila + 10, 3) = C22_Valor
    Hoja5.Cells(Fila + 10, 6) = C23_Saldo
    Hoja5.Cells(Fila + 10, 7) = C24_Cuota
    
    
    Hoja5.Cells(Fila + 11, 1) = C25_Ordinal
    Hoja5.Cells(Fila + 12, 1) = C26_Extra
    Hoja5.Cells(Fila + 13, 1) = C27_Vesper1
    Hoja5.Cells(Fila + 14, 1) = C28_Vesper2
    Hoja5.Cells(Fila + 15, 1) = C29_Nocturne
    Hoja5.Cells(Fila + 16, 1) = C30_Comision
    Hoja5.Cells(Fila + 17, 1) = C43_Ajuste
    
             
    Hoja5.Cells(Fila + 11, 5) = C31_PrestamoP
    Hoja5.Cells(Fila + 12, 5) = C32_PrestamoB
    Hoja5.Cells(Fila + 13, 5) = C33_Vales
    Hoja5.Cells(Fila + 14, 5) = C34_Florencia
    Hoja5.Cells(Fila + 15, 5) = C35_Otros
    Hoja5.Cells(Fila + 16, 5) = C44_Viatico
    Hoja5.Cells(Fila + 17, 5) = C36_SeguroS
    Hoja5.Cells(Fila + 18, 5) = C37_SeguroE
    ''''''''''''''''''''''''''''''''''''''''''''''''''''Hoja5.Cells(Fila + 19, 5) = C38_IR
    
    Hoja5.Cells(Fila + 20, 1) = C39_TIngreso
    Hoja5.Cells(Fila + 20, 5) = C40_TEgreso
    Hoja5.Cells(Fila + 21, 1) = C41_NPago
    Hoja5.Cells(Fila + 24, 1) = C42_Conforme
    
   
'DISEÑO Y FORMATO DE LA ESTRUCTURA "COLOR Y CENTRADO

Hoja5.Range(Cells(Fila, 1), Cells(Fila, 8)).Select
    Diseño_A
    Formato_A

Hoja5.Range(Cells(Fila + 1, 1), Cells(Fila + 1, 4)).Select
    Diseño_A
    Formato_A

Hoja5.Range(Cells(Fila + 1, 5), Cells(Fila + 1, 8)).Select
    Diseño_A
    Formato_A
    
Hoja5.Range(Cells(Fila + 2, 5), Cells(Fila + 3, 5)).Select
    Diseño_D

Hoja5.Range(Cells(Fila + 2, 6), Cells(Fila + 3, 8)).Select
    Diseño_A
    
Hoja5.Range(Cells(Fila + 7, 1), Cells(Fila + 7, 8)).Select
    Diseño_A
    Formato_A

Hoja5.Range(Cells(Fila + 10, 1), Cells(Fila + 10, 8)).Select
    Diseño_B
    Formato_A

Hoja5.Range(Cells(Fila + 20, 1), Cells(Fila + 21, 8)).Select
    Formato_A

Hoja5.Range(Cells(Fila + 20, 1), Cells(Fila + 20, 2)).Select
    Diseño_C
    
Hoja5.Range(Cells(Fila + 20, 3), Cells(Fila + 20, 4)).Select
    Diseño_C

Hoja5.Range(Cells(Fila + 20, 5), Cells(Fila + 20, 6)).Select
    Diseño_C
    
Hoja5.Range(Cells(Fila + 20, 7), Cells(Fila + 20, 8)).Select
    Diseño_C

Hoja5.Range(Cells(Fila + 21, 1), Cells(Fila + 21, 4)).Select
    Diseño_C
 
Hoja5.Range(Cells(Fila + 24, 1), Cells(Fila + 24, 8)).Select
    Diseño_A

'DISEÑO Y FORMATO DE ESTRUCTURA "BORDES"

Hoja5.Range(Cells(Fila, 1), Cells(Fila, 8)).Select
    Borde_A
    Borde_B

Hoja5.Range(Cells(Fila + 1, 1), Cells(Fila + 1, 8)).Select
    Borde_B
  
Hoja5.Range(Cells(Fila + 7, 1), Cells(Fila + 7, 8)).Select
    Borde_A
    Borde_B

Hoja5.Range(Cells(Fila + 10, 1), Cells(Fila + 10, 8)).Select
    Borde_A
    Borde_B
    
Hoja5.Range(Cells(Fila + 20, 1), Cells(Fila + 20, 8)).Select
    Borde_A
    Borde_B

Hoja5.Range(Cells(Fila + 21, 1), Cells(Fila + 21, 8)).Select
    Borde_B

Hoja5.Range(Cells(Fila + 1, 5), Cells(Fila + 6, 8)).Select
    Borde_I
    
Hoja5.Range(Cells(Fila + 8, 1), Cells(Fila + 21, 4)).Select
    Borde_D
    
Hoja5.Range(Cells(Fila + 23, 3), Cells(Fila + 23, 5)).Select
    Borde_B



'LLENADO DE DATOS DE COLILLA "DATOS DE VARIABLES"

Referencia = Hoja4.Cells(xFila, 1)


Hoja4.Select
Range("A4").Select

    Periodo = Hoja4.Cells(2, 1).Value

    Do Until IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value Like Referencia Then
            encontrado = True
            Id_Personal = ActiveCell.Offset(0, 0).Value
            Colaborador = ActiveCell.Offset(0, 1).Value
            Cedula = ActiveCell.Offset(0, 2).Value
            Cargo = ActiveCell.Offset(0, 3).Value
            ''Contrato = ActiveCell.Offset(0, 6).Value
            SalarioMensual = ActiveCell.Offset(0, 10).Value
            SalarioQuincenal = ActiveCell.Offset(0, 11).Value
'            SalarioHora = ActiveCell.Offset(0, 12).Value
            IngresoBase = ActiveCell.Offset(0, 26).Value
            HAusencia = ActiveCell.Offset(0, 21).Value
'            VAusencia = ActiveCell.Offset(0, 12).Value
            EAusencia = ActiveCell.Offset(0, 27).Value
            IngresoOrdinario = ActiveCell.Offset(0, 28).Value
            HDiurnas = ActiveCell.Offset(0, 22).Value
'            VDiurnas = ActiveCell.Offset(0, 13).Value
            IDiurnas = ActiveCell.Offset(0, 29).Value
            HVesper1 = ActiveCell.Offset(0, 23).Value
'            VVesper1 = ActiveCell.Offset(0, 14).Value
            IVesper1 = ActiveCell.Offset(0, 30).Value
            HVesper2 = ActiveCell.Offset(0, 24).Value
'            VVesper2 = ActiveCell.Offset(0, 15).Value
            IVesper2 = ActiveCell.Offset(0, 31).Value
            HNocturne = ActiveCell.Offset(0, 25).Value
'            VNocturne = ActiveCell.Offset(0, 16).Value
            INocturne = ActiveCell.Offset(0, 32).Value
            SPrestamoP = ActiveCell.Offset(0, 46).Value
            CPrestamoP = "'" & ActiveCell.Offset(0, 47).Value
            EPrestamoP = ActiveCell.Offset(0, 36).Value
            SPrestamoB = ActiveCell.Offset(0, 48).Value
            CPrestamoB = "'" & ActiveCell.Offset(0, 49).Value
            EPrestamoB = ActiveCell.Offset(0, 37).Value
            SVale = ActiveCell.Offset(0, 50).Value
            CVale = "'" & ActiveCell.Offset(0, 51).Value
            EVale = ActiveCell.Offset(0, 38).Value
            SVentaF = ActiveCell.Offset(0, 52).Value
            CVentaF = "'" & ActiveCell.Offset(0, 53).Value
            EVentaF = ActiveCell.Offset(0, 39).Value
            SOtroD = ActiveCell.Offset(0, 54).Value
            COtroD = "'" & ActiveCell.Offset(0, 55).Value
            EOtroD = ActiveCell.Offset(0, 40).Value
            Comision = ActiveCell.Offset(0, 33).Value
            Ajuste = ActiveCell.Offset(0, 34).Value
            Viatico = ActiveCell.Offset(0, 41).Value
            SeguroSocial = ActiveCell.Offset(0, 42).Value
            SeguroEducativo = ActiveCell.Offset(0, 43).Value
            ''''''''''''''''''''''''''''''''''''''''''''''''''''IR = ActiveCell.Offset(0, 45).Value
            TIngresos = ActiveCell.Offset(0, 35).Value
            TEgresos = ActiveCell.Offset(0, 44).Value
            NPago = ActiveCell.Offset(0, 45).Value
            Exit Do
        End If
    Loop
            
'LLENADO DE DATOS DE COLILLA "INGRESO DE DATOS"

    Hoja5.Activate

    Hoja5.Cells(Fila + 2, 2) = Id_Personal
    Hoja5.Cells(Fila + 3, 2) = Colaborador
    Hoja5.Cells(Fila + 4, 2) = Cedula
    Hoja5.Cells(Fila + 5, 2) = Cargo
    ''Hoja5.Cells(Fila + 6, 2) = Contrato
    
    Hoja5.Cells(Fila + 2, 6) = Periodo
    Hoja5.Cells(Fila + 4, 6) = SalarioMensual
    Hoja5.Cells(Fila + 5, 6) = SalarioQuincenal
    Hoja5.Cells(Fila + 6, 6) = SalarioHora
    
    Hoja5.Cells(Fila + 8, 4) = IngresoBase
    Hoja5.Cells(Fila + 9, 2) = HAusencia
    Hoja5.Cells(Fila + 9, 3) = VAusencia
    Hoja5.Cells(Fila + 9, 4) = EAusencia
    Hoja5.Cells(Fila + 8, 8) = IngresoOrdinario
    
    Hoja5.Cells(Fila + 11, 4) = IngresoOrdinario
        
    Hoja5.Cells(Fila + 12, 2) = HDiurnas
    Hoja5.Cells(Fila + 12, 3) = VDiurnas
    Hoja5.Cells(Fila + 12, 4) = IDiurnas
    
    Hoja5.Cells(Fila + 13, 2) = HVesper1
    Hoja5.Cells(Fila + 13, 3) = VVesper1
    Hoja5.Cells(Fila + 13, 4) = IVesper1
    
    Hoja5.Cells(Fila + 14, 2) = HVesper2
    Hoja5.Cells(Fila + 14, 3) = VVesper2
    Hoja5.Cells(Fila + 14, 4) = IVesper2
    
    Hoja5.Cells(Fila + 15, 2) = HNocturne
    Hoja5.Cells(Fila + 15, 3) = VNocturne
    Hoja5.Cells(Fila + 15, 4) = INocturne
    
    Hoja5.Cells(Fila + 16, 4) = Comision
    Hoja5.Cells(Fila + 17, 4) = Ajuste
    
    Hoja5.Cells(Fila + 11, 6) = SPrestamoP
    Hoja5.Cells(Fila + 11, 7) = CPrestamoP
    Hoja5.Cells(Fila + 11, 8) = EPrestamoP
    
    Hoja5.Cells(Fila + 12, 6) = SPrestamoB
    Hoja5.Cells(Fila + 12, 7) = CPrestamoB
    Hoja5.Cells(Fila + 12, 8) = EPrestamoB
    
    Hoja5.Cells(Fila + 13, 6) = SVale
    Hoja5.Cells(Fila + 13, 7) = CVale
    Hoja5.Cells(Fila + 13, 8) = EVale
        
    Hoja5.Cells(Fila + 14, 6) = SVentaF
    Hoja5.Cells(Fila + 14, 7) = CVentaF
    Hoja5.Cells(Fila + 14, 8) = EVentaF
    
    Hoja5.Cells(Fila + 15, 6) = SOtroD
    Hoja5.Cells(Fila + 15, 7) = COtroD
    Hoja5.Cells(Fila + 15, 8) = EOtroD
    
    Hoja5.Cells(Fila + 16, 8) = Viatico
    Hoja5.Cells(Fila + 17, 8) = SeguroSocial
    Hoja5.Cells(Fila + 18, 8) = SeguroEducativo
    ''''''''''''''''''''''''''''''''''''''''''''''''''''Hoja5.Cells(Fila + 19, 8) = IR
    
    Hoja5.Cells(Fila + 20, 3) = TIngresos
    Hoja5.Cells(Fila + 20, 7) = TEgresos
    Hoja5.Cells(Fila + 21, 5) = NPago
    
    
'FORMATO DE DATOS DE VARIABLES

    Hoja5.Cells(Fila + 2, 2).HorizontalAlignment = xlLeft
    Hoja5.Cells(Fila + 3, 2).HorizontalAlignment = xlLeft
    Hoja5.Cells(Fila + 4, 2).HorizontalAlignment = xlLeft
    Hoja5.Cells(Fila + 5, 2).HorizontalAlignment = xlLeft
    Hoja5.Cells(Fila + 6, 2).HorizontalAlignment = xlLeft
    
    Hoja5.Cells(Fila + 2, 6).WrapText = True
    Hoja5.Cells(Fila + 2, 6).Font.Size = 7
    Hoja5.Cells(Fila + 4, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 5, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 6, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 8, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 9, 2).NumberFormat = "[hh]:mm"
    Hoja5.Cells(Fila + 9, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 9, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 8, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 11, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 12, 2).NumberFormat = "[hh]:mm"
    Hoja5.Cells(Fila + 12, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 12, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 13, 2).NumberFormat = "[hh]:mm"
    Hoja5.Cells(Fila + 13, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 13, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 14, 2).NumberFormat = "[hh]:mm"
    Hoja5.Cells(Fila + 14, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 14, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 15, 2).NumberFormat = "[hh]:mm"
    Hoja5.Cells(Fila + 15, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 15, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 16, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 17, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 11, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 11, 7).HorizontalAlignment = xlCenter
    Hoja5.Cells(Fila + 11, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 12, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 12, 7).HorizontalAlignment = xlCenter
    Hoja5.Cells(Fila + 12, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 13, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 13, 7).HorizontalAlignment = xlCenter
    Hoja5.Cells(Fila + 13, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 14, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 14, 7).HorizontalAlignment = xlCenter
    Hoja5.Cells(Fila + 14, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        
    Hoja5.Cells(Fila + 15, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 15, 7).HorizontalAlignment = xlCenter
    Hoja5.Cells(Fila + 15, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
       
    Hoja5.Cells(Fila + 16, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 17, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 18, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 19, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    Hoja5.Cells(Fila + 20, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 20, 7).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Hoja5.Cells(Fila + 21, 5).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
'DUPLICADO DE COLILLA
Hoja5.Activate
   
  Fila = Fila + 26
  
Next

Hoja5.Activate

Application.ScreenUpdating = True
Application.EnableEvents = True
    
End Sub

Public Sub GENERAL_GUARDAR_PDF()
Dim mi_hoja As Worksheet
Dim mi_ventana As FileDialog
Dim mi_carpeta As String
Dim mi_archivo As Integer
Dim Nombre_PDF As String
Dim mi_referencia
Dim RutaArchivo As String

Set mi_ventana = Application.FileDialog(msoFileDialogFolderPicker)

If mi_ventana.Show = True Then
    mi_carpeta = mi_ventana.SelectedItems(1)
    mi_referencia = mi_ventana.SelectedItems(1)
Else
    MsgBox "No se indicó la carpeta donde guardar el PDF…" _
    & vbCrLf & vbCrLf & "Operación cancelada.", _
    vbCritical, "Carpeta de almacenamiento pdf"
    Exit Sub
End If

Nombre_PDF = ThisWorkbook.Name

mi_carpeta = mi_carpeta + "\" + Nombre_PDF + ".pdf"


If Len(Dir(mi_carpeta)) > 0 Then
    mi_archivo = MsgBox(mi_carpeta & " existente." _
    & vbCrLf & vbCrLf & "¿Desea reemplazarlo?", _
    vbYesNo + vbQuestion, "Archivo existente")
    
    On Error Resume Next
    
    If mi_archivo = vbYes Then
        Kill mi_carpeta
    Else
        MsgBox "Reemplazar el archivo PDF existente para continuar…" _
        & vbCrLf & vbCrLf & "Operación cancelada.", _
        vbCritical, "Confirmar guardar como"
        Exit Sub
    End If
    
    If Err.Number <> 0 Then
        MsgBox "El archivo pdf se encuentra abierto o protegido como sólo lectura." _
        & vbCrLf & vbCrLf, vbCritical, _
        "Error al guardar el archivo"
        Exit Sub
    End If
End If



RutaArchivo = mi_referencia & "\" & Nombre_PDF & ".pdf"

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=RutaArchivo, _
Quality:=xlQualityStandard, IncludeDocProperties:=True, _
IgnorePrintAreas:=False, OpenAfterPublish:=True

End Sub
Public Sub GENERAL_ENVIAR_PDF()
Dim mi_hoja As Worksheet
Dim mi_ventana As FileDialog
Dim mi_carpeta As String
Dim mi_archivo As Integer
Dim Nombre_PDF As String
Dim mi_referencia
Dim mi_outlook As Object
Dim mi_correo As Object
Dim RutaArchivo As String
Dim DisplayEmail As Object

Set mi_ventana = Application.FileDialog(msoFileDialogFolderPicker)

If mi_ventana.Show = True Then
    mi_carpeta = mi_ventana.SelectedItems(1)
    mi_referencia = mi_ventana.SelectedItems(1)
Else
    MsgBox "No se indicó la carpeta donde guardar el PDF…" _
    & vbCrLf & vbCrLf & "Operación cancelada.", _
    vbCritical, "Carpeta de almacenamiento pdf"
    Exit Sub
End If

Nombre_PDF = ThisWorkbook.Name

mi_carpeta = mi_carpeta + "\" + Nombre_PDF + ".pdf"


If Len(Dir(mi_carpeta)) > 0 Then
    mi_archivo = MsgBox(mi_carpeta & " existente." _
    & vbCrLf & vbCrLf & "¿Desea reemplazarlo?", _
    vbYesNo + vbQuestion, "Archivo existente")
    
    On Error Resume Next
    
    If mi_archivo = vbYes Then
        Kill mi_carpeta
    Else
        MsgBox "Reemplazar el archivo PDF existente para continuar…" _
        & vbCrLf & vbCrLf & "Operación cancelada.", _
        vbCritical, "Confirmar guardar como"
        Exit Sub
    End If
    
    If Err.Number <> 0 Then
        MsgBox "El archivo pdf se encuentra abierto o protegido como sólo lectura." _
        & vbCrLf & vbCrLf, vbCritical, _
        "Error al guardar el archivo"
        Exit Sub
    End If
End If



RutaArchivo = mi_referencia & "\" & Nombre_PDF & ".pdf"

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=RutaArchivo, _
Quality:=xlQualityStandard, IncludeDocProperties:=True, _
IgnorePrintAreas:=False, OpenAfterPublish:=True

Set mi_outlook = CreateObject("Outlook.Application")
Set mi_correo = mi_outlook.CreateItem(0)

With mi_correo
.Display
.To = ""
.CC = ""
.Subject = Nombre_PDF + ".pdf"
.Attachments.Add RutaArchivo
End With



End Sub

    
Public Sub Diseño_A()
    With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .MergeCells = True
    End With
End Sub
Public Sub Diseño_B()
    With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
End Sub
Public Sub Diseño_C()
    With Selection
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlCenter
    .MergeCells = True
    End With
End Sub
Public Sub Diseño_D()
    With Selection
    .VerticalAlignment = xlCenter
    .MergeCells = True
    End With
End Sub
Public Sub Formato_A()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
End Sub
Public Sub Borde_A()
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
Public Sub Borde_B()
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
Public Sub Borde_I()
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
Public Sub Borde_D()
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub
Public Sub General_Quincena()

Hoja3.Activate
Hoja3.Range("C2").Select
Hoja3.Range("C2") = CDate(Me.txt_Fecha)

End Sub


Private Sub btn_Fecha_Click()

banderaPeriodo = 2
  Call LanzarPeriodo(Me, "txt_Fecha")
  Me.btn_Cargar.SetFocus
End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub


