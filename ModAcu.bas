Attribute VB_Name = "ModAcu"
Type sc
    guarda As String * 260
End Type
Type te
    miarchivo As String * 64
End Type
Type Nacu
     Pdias As Integer
     Pnormal As Currency
     Pextra As Currency
     Pviaticos As Currency
     Pvaca As Currency
     Potras As Currency
     Pagui As Currency
     PPTU As Currency
     Pexenta As Currency
     PCompensa As Currency      ' NUEVO - LIQ
     PAntiguedad As Currency    ' NUEVO - LIQ
     PIndemniza As Currency     ' NUEVO - LIQ
     PSueldoVac As Currency     ' NUEVO - LIQ
     PBono As Currency          ' NUEVO - ESP
     DImpto As Currency
     DSubioAp As Currency
     DCrApl As Currency
     DImpret As Currency
     DCrPag As Currency
     DSubNoap As Currency
     DImss As Currency
     DPrestamos As Currency
     DTonacot As Currency
     DTelefono As Currency
     DOtrasded As Currency
End Type
Type NacuOt
     Pdias As Integer
     Pnormal As Currency
     Pextra As Currency
     Pviaticos As Currency
     Pvaca As Currency
     Potras As Currency
     Pagui As Currency
     PPTU As Currency
     Pexenta As Currency
     DImpto As Currency
     DSubioAp As Currency
     DCrApl As Currency
     DImpret As Currency
     DCrPag As Currency
     DSubNoap As Currency
     DImss As Currency
     DPrestamos As Currency
     DTonacot As Currency
     DTelefono As Currency
     DOtrasded As Currency
End Type
Type AxN
     Narch As String * 30
     Pdias As Integer
     Pnormal As Currency
     Pextra As Currency
     Pviaticos As Currency
     Pvaca As Currency
     Potras As Currency
     Pagui As Currency
     PPTU As Currency
     Pexenta As Currency
     PCompensa As Currency      ' NUEVO
     PAntiguedad As Currency    ' NUEVO
     PIndemniza As Currency     ' NUEVO
     PSueldoVac As Currency     ' NUEVO
     PBono As Currency          ' NUEVO
     DImpto As Currency
     DSubioAp As Currency
     DCrApl As Currency
     DImpret As Currency
     DCrPag As Currency
     DSubNoap As Currency
     DImss As Currency
     DPrestamos As Currency
     DTonacot As Currency
     DTelefono As Currency
     DOtrasded As Currency
End Type
 Type OtrasCh
      curp As String * 30
      otra As String * 30
      yotra As String * 30
      yporsi As String * 30
 End Type
 Public Otros_Rgtros As OtrasCh
 
Public Enum TipoNominaAcum
    taOrdinaria = 0
    taLiquidacion = 1
    taAguinaldo = 2
    taPTU = 3
    taBono = 4
End Enum

Public AxNom As AxN, Aumento1 As Long, Kincenal As Integer
Public AxPer As Nacu, Ftem As Long, r As Long, a_opago As Integer
Public ArAcum As Nacu, SUMA_CREDITO_MES As Currency
Public Ot_Acum As NacuOt, ta_r As Integer
Public temporal As te, DIR_CALC As String
Public SCont As sc, AcumSup_Cor As Currency, AcumSup_CorEx As Currency
Public Subidio_doble As Currency, DirecT_arifas As String
Sub derecha(ancho2, ltotal, cadena As String)
    ancho2 = 0
    ancho2 = (ltotal - Printer.TextWidth(cadena))
End Sub


Public Function ClasificarNominaAcum(ByVal NombreArchivo As String) As TipoNominaAcum

    Dim n As String

    n = UCase(Trim(NombreArchivo))

    If InStr(n, "LIQUID") > 0 _
    Or InStr(n, "LIQUI") > 0 _
    Or InStr(n, "LIQ") > 0 _
    Or InStr(n, "FINIQUITO") > 0 _
    Or InStr(n, "FINIQ") > 0 _
    Or InStr(n, "FINI") > 0 Then

        ClasificarNominaAcum = taLiquidacion

    ElseIf InStr(n, "AGUINALDO") > 0 _
        Or InStr(n, "AGUI") > 0 Then

        ClasificarNominaAcum = taAguinaldo

    ElseIf InStr(n, "PTU") > 0 Then

        ClasificarNominaAcum = taPTU

    ElseIf InStr(n, "BONO") > 0 Then

        ClasificarNominaAcum = taBono

    Else

        ClasificarNominaAcum = taOrdinaria

    End If

End Function

' comentario
