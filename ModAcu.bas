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
     PexentaIncluida As Currency    ' Nueva separacion: exenta ya incluida en bruto
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
     PexentaIncluida As Currency    ' Nueva separacion: exenta ya incluida en bruto
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
    taLiquidacion = 1     ' hist?rico
    taAguinaldo = 2       ' hist?rico
    taPTU = 3             ' hist?rico
    taBono = 4            ' hist?rico
    taNuevaESP = 5        ' ESP / PTU / Otro (>= corte)
    taNuevaLIQ = 6
    taNuevaFIN = 7
    taSeparacion2026LIQ = 8
    taSeparacion2026FIN = 9
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



' Utilidades comunes. No alteran Type nom ni el formato .NOM.
Public Function NombreBaseNomAcum(ByVal rutaArchivo As String) As String
    Dim p As Long
    rutaArchivo = UCase$(Trim$(rutaArchivo))
    p = InStrRev(rutaArchivo, "\")
    NombreBaseNomAcum = Mid$(rutaArchivo, p + 1)
End Function

Public Function ArchivoMarcaSeparacion2026(ByVal Archivo As String) As Boolean
    Dim f As Integer
    Dim n As nom
    Dim abiertos As Boolean

    Archivo = Trim$(Archivo)
    If Len(Dir$(Archivo)) = 0 Then Exit Function
    If FileLen(Archivo) < Len(n) Then Exit Function
    f = FreeFile
    On Error GoTo ErrorMarca
    Open Archivo For Random Access Read As #f Len = Len(n)
    abiertos = True
    Get #f, 1, n
    ArchivoMarcaSeparacion2026 = (n.hstri = CCur(20261010))
    Close #f
    Exit Function
ErrorMarca:
    Dim numero As Long
    Dim descripcion As String
    numero = Err.Number
    descripcion = Err.Description
    If abiertos Then Close #f
    Err.Raise numero, "ArchivoMarcaSeparacion2026", descripcion
End Function

Public Function EsSeparacion2026Acum(ByVal Tipo As TipoNominaAcum) As Boolean
    EsSeparacion2026Acum = (Tipo = taSeparacion2026LIQ Or Tipo = taSeparacion2026FIN)
End Function

Public Function IngresosTotalesAcum(ByRef a As Nacu) As Currency
    ' Historicas: exenta adicional. Nuevas LIQ/FIN: exenta contenida en bruto.
    IngresosTotalesAcum = a.Pnormal + a.Pextra + a.Pviaticos + _
        a.Pvaca + a.Potras + a.Pagui + a.PPTU + a.Pexenta + _
        a.PCompensa + a.PAntiguedad + a.PIndemniza + _
        a.PSueldoVac + a.PBono - a.PexentaIncluida
End Function

Public Function IngresosTotalesDetalle(ByRef a As AxN) As Currency
    IngresosTotalesDetalle = a.Pnormal + a.Pextra + a.Pviaticos + _
        a.Pvaca + a.Potras + a.Pagui + a.PPTU + a.Pexenta + _
        a.PCompensa + a.PAntiguedad + a.PIndemniza + _
        a.PSueldoVac + a.PBono - a.PexentaIncluida
End Function

' Lee solo el mes de la fecha REAL guardada en .FEC, si existe y es valida.
Public Function MesFichaAcumulado(ByVal Archivo As String, ByRef mesPago As Integer) As Boolean
    Dim rutaFec As String
    Dim f As Integer
    Dim cab As String, fechaISO As String, Nombre As String
    Dim a As Long, m As Long, d As Long, fecha As Date
    Dim abierto As Boolean

    mesPago = 0
    Archivo = Trim$(Archivo)
    If Len(Archivo) < 5 Then Exit Function
    rutaFec = Left$(Archivo, Len(Archivo) - 4) & ".FEC"
    If Len(Dir$(rutaFec)) = 0 Then Exit Function

    f = FreeFile
    On Error GoTo FalloFicha
    Open rutaFec For Input As #f
    abierto = True
    Line Input #f, cab
    Line Input #f, fechaISO
    Line Input #f, Nombre
    Close #f
    abierto = False

    If Trim$(cab) <> "SACMAG-FECHA-2026" Then GoTo FalloFormato
    If UCase$(Trim$(Nombre)) <> NombreBaseNomAcum(Archivo) Then GoTo FalloFormato
    If Len(fechaISO) <> 10 Then GoTo FalloFormato
    If Mid$(fechaISO, 5, 1) <> "-" Or Mid$(fechaISO, 8, 1) <> "-" Then GoTo FalloFormato
    If Not IsNumeric(Left$(fechaISO, 4)) Then GoTo FalloFormato
    If Not IsNumeric(Mid$(fechaISO, 6, 2)) Then GoTo FalloFormato
    If Not IsNumeric(Right$(fechaISO, 2)) Then GoTo FalloFormato
    a = CLng(Left$(fechaISO, 4))
    m = CLng(Mid$(fechaISO, 6, 2))
    d = CLng(Right$(fechaISO, 2))
    If a < 1900 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then GoTo FalloFormato
    fecha = DateSerial(a, m, d)
    If Year(fecha) <> a Or Month(fecha) <> m Or Day(fecha) <> d Then GoTo FalloFormato
    mesPago = CInt(m)
    MesFichaAcumulado = True
    Exit Function

FalloFormato:
    Err.Raise vbObjectError + 6202, "MesFichaAcumulado", "Ficha .FEC incorrecta: " & rutaFec
FalloFicha:
    Dim numero As Long, descripcion As String
    numero = Err.Number: descripcion = Err.Description
    If abierto Then Close #f
    Err.Raise numero, "MesFichaAcumulado", descripcion
End Function

Public Function EsNominaNormal(ByVal Nombre As String) As Boolean
    Dim n As String
    n = NombreBaseNomAcum(Nombre)
    If Not (n Like "???[12]####.NOM") Then Exit Function
    Select Case Left$(n, 3)
        Case "ENE", "FEB", "MAR", "ABR", "MAY", "JUN", _
             "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"
            EsNominaNormal = True
    End Select
End Function
Public Function EsquemaNuevo(ByVal Archivo As String) As Boolean
    ' Compatibilidad del nombre, pero criterio determinista: leer la marca.
    ' FileDateTime NO indica el formato de los registros, cambia al copiar.
    EsquemaNuevo = ArchivoMarcaSeparacion2026(Archivo)
End Function
Public Function ClasificarNominaAcum(ByVal NombreArchivo As String) As TipoNominaAcum
    Dim n As String
    Dim tp As String
    Dim vr As Integer
    Dim mk As Integer

    n = NombreBaseNomAcum(NombreArchivo)

    If EsNominaNormal(n) Then
        ClasificarNominaAcum = taOrdinaria
        Exit Function
    End If

    ' La marca REAL va en el sexto Currency del registro NOM.
    ' El nombre puede comenzar con un prefijo numerico: 15FINIQUITO.NOM.
    If ArchivoMarcaSeparacion2026(NombreArchivo) Then
        If InStr(1, n, "FINI", vbTextCompare) > 0 Then
            ClasificarNominaAcum = taSeparacion2026FIN
            Exit Function
        End If
        If InStr(1, n, "LIQ", vbTextCompare) > 0 Then
            ClasificarNominaAcum = taSeparacion2026LIQ
            Exit Function
        End If
        Err.Raise vbObjectError + 6201, "ClasificarNominaAcum", _
                  "Archivo marcado como separacion nueva sin nombre LIQ/FINI: " & NombreArchivo
    End If

    ' Conserva interpretacion de versiones previas marcadas con .TIP.
    ' No usar la fecha de MODIFICACION para decidir el formato del binario.
    If LeerMarcaTipo(NombreArchivo, tp, vr, mk) Then
        If vr >= 2 Then
            Select Case tp
                Case "LIQ": ClasificarNominaAcum = taNuevaLIQ: Exit Function
                Case "FIN": ClasificarNominaAcum = taNuevaFIN: Exit Function
                Case "ESP": ClasificarNominaAcum = taNuevaESP: Exit Function
            End Select
        End If
    End If

    If InStr(1, n, "LIQ", vbTextCompare) > 0 Or _
       InStr(1, n, "FINI", vbTextCompare) > 0 Then
        ClasificarNominaAcum = taLiquidacion
    ElseIf InStr(1, n, "AGUI", vbTextCompare) > 0 Then
        ClasificarNominaAcum = taAguinaldo
    ElseIf InStr(1, n, "PTU", vbTextCompare) > 0 Then
        ClasificarNominaAcum = taPTU
    ElseIf InStr(1, n, "BONO", vbTextCompare) > 0 Then
        ClasificarNominaAcum = taBono
    Else
        ClasificarNominaAcum = taOrdinaria
    End If
End Function
' Usa la variable global "nomina" (el registro reci?n le?do)
Public Sub CalcularConceptos(ByVal Tipo As TipoNominaAcum, ByVal Archivo As String, c As ConceptosNom)
    Dim hn As Currency, hd As Currency, ht As Currency, vi As Currency
    Dim vacio As ConceptosNom

    c = vacio
    hn = nomina.hs_nor
    hd = nomina.hs_dbl
    ht = nomina.hs_tri
    vi = nomina.viaticos
    c.Pvaca = nomina.pvac
    c.Pagui = nomina.aguin
    c.PPTU = nomina.ptu
    c.Potras = nomina.otras

    Select Case Tipo
        Case taSeparacion2026LIQ
            ' Form8 2026: COMPENS=hs_nor; ANTIG=hs_dbl; INDEMN=hs_tri.
            c.PCompensa = nomina.hs_nor
            c.PAntiguedad = nomina.hs_dbl
            c.PIndemniza = nomina.hs_tri
            c.PPTU = 0

        Case taSeparacion2026FIN
            ' Form8 2026: SUELDO VAC=viaticos; ANTIG=hs_dbl; BONO=hsnor.
            c.PSueldoVac = nomina.viaticos
            c.PAntiguedad = nomina.hs_dbl
            c.PBono = nomina.hsnor
            c.PPTU = 0

        Case taNuevaESP
            ' Conserva la interpretacion anterior del formato .TIP.
            c.Pextra = hn
            c.PBono = vi

        Case taNuevaLIQ
            ' Esquema .TIP anterior: no confundir con 2026 marcado en NOM.
            c.PCompensa = hn
            c.PAntiguedad = nomina.ptu
            c.PPTU = 0
            c.PIndemniza = vi

        Case taNuevaFIN
            ' Esquema .TIP anterior: no confundir con 2026 marcado en NOM.
            c.PSueldoVac = hn
            c.PAntiguedad = nomina.ptu
            c.PPTU = 0
            c.Pviaticos = vi

        Case taLiquidacion
            ' Original: conservar interpretacion historica de finiquitos/liquidaciones.
            c.PCompensa = hn
            c.PAntiguedad = hd
            c.PIndemniza = ht
            c.PSueldoVac = vi

        Case taBono
            If nomina.viaticos <> 0 Then
                ' Estructura comprobada en 1BONO.NOM.
                c.PBono = nomina.viaticos
        
                ' Conservar Otras cuando exista.
                c.Potras = nomina.otras
            Else
                ' Compatibilidad con Bono historico.
                c.PBono = nomina.otras
                c.Potras = 0
            End If

        Case Else
            If UCase$(Left$(Trim$(Archivo), 3)) = "PRE" Then ht = vi: vi = 0
            c.Pextra = hn + hd + ht
            c.Pviaticos = vi
    End Select
End Sub
' Lee Archivo.TIP (TIPO|VERSION|MES). False si no existe o est? mal formado.
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

