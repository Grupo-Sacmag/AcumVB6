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
    taLiquidacion = 1     ' histórico
    taAguinaldo = 2       ' histórico
    taPTU = 3             ' histórico
    taBono = 4            ' histórico
    taNuevaESP = 5        ' ESP / PTU / Otro (>= corte)
    taNuevaLIQ = 6
    taNuevaFIN = 7
End Enum

Public Type ConceptosNom
    Pextra As Currency
    Pviaticos As Currency
    Pvaca As Currency
    Potras As Currency
    Pagui As Currency
    PPTU As Currency
    PCompensa As Currency
    PAntiguedad As Currency
    PIndemniza As Currency
    PSueldoVac As Currency
    PBono As Currency
End Type

Public Const FECHA_CORTE_ESQUEMA As Date = #10/10/2026#
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


Public Function EsNominaNormal(ByVal Nombre As String) As Boolean
    Dim n As String
    n = UCase$(Trim$(Nombre))
    If Not (n Like "???[12]####.NOM") Then Exit Function
    Select Case Left$(n, 3)
        Case "ENE", "FEB", "MAR", "ABR", "MAY", "JUN", _
             "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"
            EsNominaNormal = True
    End Select
End Function

Public Function EsquemaNuevo(ByVal Archivo As String) As Boolean
    On Error GoTo sinFecha
    EsquemaNuevo = (DateValue(FileDateTime(Trim$(Archivo))) >= FECHA_CORTE_ESQUEMA)
    Exit Function
sinFecha:
    EsquemaNuevo = False
End Function

Public Function ClasificarNominaAcum(ByVal NombreArchivo As String) As TipoNominaAcum
    Dim n As String, tp As String, vr As Integer, mk As Integer
    n = UCase$(Trim$(NombreArchivo))

    If EsNominaNormal(n) Then
        ClasificarNominaAcum = taOrdinaria
        Exit Function
    End If

    If LeerMarcaTipo(NombreArchivo, tp, vr, mk) Then
        ' Con marca manda la marca. Versión 1 u "OTR" caen al esquema histórico (abajo).
        If vr >= 2 Then
            Select Case tp
                Case "LIQ": ClasificarNominaAcum = taNuevaLIQ: Exit Function
                Case "FIN": ClasificarNominaAcum = taNuevaFIN: Exit Function
                Case "ESP": ClasificarNominaAcum = taNuevaESP: Exit Function
            End Select
        End If
    ElseIf EsquemaNuevo(NombreArchivo) Then
        ' Sin marca (archivos posteriores al corte, anteriores a esta versión)
        Select Case Left$(n, 3)
            Case "LIQ": ClasificarNominaAcum = taNuevaLIQ: Exit Function
            Case "FIN": ClasificarNominaAcum = taNuevaFIN: Exit Function
            Case "ESP": ClasificarNominaAcum = taNuevaESP: Exit Function
        End Select
        ' cualquier otro nombre = "Otro" -> esquema histórico
    End If

    ' ---- Histórico / Otro: igual que antes ----
    If InStr(n, "LIQ") > 0 Or InStr(n, "FINI") > 0 Then
        ClasificarNominaAcum = taLiquidacion
    ElseIf InStr(n, "AGUI") > 0 Then
        ClasificarNominaAcum = taAguinaldo
    ElseIf InStr(n, "PTU") > 0 Then
        ClasificarNominaAcum = taPTU
    ElseIf InStr(n, "BONO") > 0 Then
        ClasificarNominaAcum = taBono
    Else
        ClasificarNominaAcum = taOrdinaria
    End If
End Function
' Usa la variable global "nomina" (el registro recién leído)
Public Sub CalcularConceptos(ByVal Tipo As TipoNominaAcum, ByVal Archivo As String, c As ConceptosNom)
    Dim hn As Currency, hd As Currency, ht As Currency, vi As Currency
    Dim vacio As ConceptosNom

    c = vacio
    hn = nomina.hs_nor: hd = nomina.hs_dbl: ht = nomina.hs_tri: vi = nomina.viaticos
    c.Pvaca = nomina.pvac
    c.Pagui = nomina.aguin
    c.PPTU = nomina.ptu
    c.Potras = nomina.otras

    Select Case Tipo
        Case taNuevaESP          ' ESP / Otro: col4 hs.Norm., col5 Aguinaldo, col6 PTU, col7 Bono
            c.Pextra = hn
            c.PBono = vi
        Case taNuevaLIQ          ' col4 Compens., col5 Aguinaldo, col6 P.Antigüedad (viene en ptu), col7 Indemnizac.
            c.PCompensa = hn
            c.PAntiguedad = nomina.ptu
            c.PPTU = 0
            c.PIndemniza = vi
        Case taNuevaFIN          ' col4 Sdo.Vac., col5 Aguinaldo, col6 P.Antigüedad (viene en ptu), col7 O.F.
            c.PSueldoVac = hn
            c.PAntiguedad = nomina.ptu
            c.PPTU = 0
            c.Pviaticos = vi
        Case taLiquidacion      ' histórico
            c.PCompensa = hn: c.PAntiguedad = hd: c.PIndemniza = ht: c.PSueldoVac = vi
        Case taBono             ' histórico
            c.PBono = nomina.otras: c.Potras = 0
        Case Else               ' ordinaria / histórico
            If UCase$(Left$(Trim$(Archivo), 3)) = "PRE" Then ht = vi: vi = 0
            c.Pextra = hn + hd + ht
            c.Pviaticos = vi
    End Select
End Sub

' Lee Archivo.TIP (TIPO|VERSION|MES). False si no existe o está mal formado.
Public Function LeerMarcaTipo(ByVal Archivo As String, Tipo As String, ver As Integer, Mes As Integer) As Boolean
    Dim f As Integer, ruta As String, linea As String, partes() As String

    Tipo = "": ver = 0: Mes = 0
    Archivo = Trim$(Archivo)
    If Len(Archivo) < 5 Then Exit Function
    ruta = Left$(Archivo, Len(Archivo) - 4) & ".TIP"

    On Error GoTo fin
    f = FreeFile
    Open ruta For Input As #f        ' si no existe, error 53 -> fin
    Line Input #f, linea
    Close #f
    On Error GoTo 0

    partes = Split(Trim$(linea), "|")
    If UBound(partes) < 1 Then Exit Function
    Tipo = UCase$(Trim$(partes(0)))
    ver = Val(partes(1))
    If UBound(partes) >= 2 Then Mes = Val(partes(2))
    LeerMarcaTipo = (ver > 0)
    Exit Function
fin:
    On Error Resume Next
    Close #f
End Function
' comentario
