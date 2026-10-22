VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form AxNom1 
   Caption         =   "Auxiliar de Personal"
   ClientHeight    =   4560
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7935
   Icon            =   "AxNom1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Axn1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   23
      FixedCols       =   2
      BackColorBkg    =   -2147483633
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Menu AuxEdic 
      Caption         =   "Edicion"
      Begin VB.Menu EditSelTot 
         Caption         =   "&Seleccionar Todo"
      End
      Begin VB.Menu AuxEdSep1 
         Caption         =   "-"
      End
      Begin VB.Menu AuxEdCop 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "AxNom1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoAch As String, FinAx2 As Long
Public Tipo As Boolean

' NUEVO: 5 columnas al final (23 a 27) para no mover las columnas existentes
Const NCOLS As Long = 28      ' total de columnas (0 a 27)
Const ULT_COL As Long = 27    ' ultima columna con importes

Private Function EntradaAxNomVisual(ByVal entradaOriginal As String) As String
    ' Unicamente orden de columnas visuales; no cambia AxNom ni el .DNO.
    ' Original: 0..9, 10..22, 23..27
    ' Pantalla: 0..9, 23..27, 10..22
    Dim original() As String
    Dim ordenada(0 To 27) As String
    Dim i As Long

    original = Split(entradaOriginal, vbTab)
    If UBound(original) <> 27 Then
        Err.Raise vbObjectError + 2027, "EntradaAxNomVisual", _
                  "La fila del detalle no tiene 28 columnas."
    End If

    For i = 0 To 9
        ordenada(i) = original(i)
    Next i
    For i = 0 To 4
        ordenada(10 + i) = original(23 + i)
    Next i
    For i = 10 To 22
        ordenada(i + 5) = original(i)
    Next i

    EntradaAxNomVisual = Join(ordenada, vbTab)
End Function

Sub ReAper()
    Axn1.Cols = NCOLS         ' NUEVO
    Axn1.Row = 0
    Axn1.Col = 0: Axn1.CellFontBold = True: Axn1.ColWidth(0) = 600: Axn1.CellAlignment = 3: Axn1.Text = "N�m."
    Axn1.ColAlignment(0) = 4
    Axn1.Col = 1: Axn1.CellFontBold = True: Axn1.ColWidth(1) = 3200: Axn1.CellAlignment = 3: Axn1.Text = "N�mbre."
    Axn1.Col = 2: Axn1.CellFontBold = True: Axn1.ColWidth(2) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Dias"
    Axn1.Col = 3: Axn1.CellFontBold = True: Axn1.ColWidth(3) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Sueldos"
    Axn1.Col = 4: Axn1.CellFontBold = True: Axn1.ColWidth(4) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Prem.Punt"
    Axn1.Col = 5: Axn1.CellFontBold = True: Axn1.ColWidth(5) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "O.F."
    Axn1.Col = 6: Axn1.CellFontBold = True: Axn1.ColWidth(6) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "P.Vacac."
    Axn1.Col = 7: Axn1.CellFontBold = True: Axn1.ColWidth(7) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Otras"
    Axn1.Col = 8: Axn1.CellFontBold = True: Axn1.ColWidth(8) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Aguinaldo"
    Axn1.Col = 9: Axn1.CellFontBold = True: Axn1.ColWidth(9) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "P.T.U."
    Axn1.Col = 15: Axn1.CellFontBold = True: Axn1.ColWidth(15) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Exentos"
    Axn1.Col = 16: Axn1.CellFontBold = True: Axn1.ColWidth(16) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Total"
    Axn1.Col = 17: Axn1.CellFontBold = True: Axn1.ColWidth(17) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Impto."
    Axn1.Col = 18: Axn1.CellFontBold = True: Axn1.ColWidth(18) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Subdio.Apl."
    Axn1.Col = 19: Axn1.CellFontBold = True: Axn1.ColWidth(19) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Cr.Apl."
    Axn1.Col = 20: Axn1.CellFontBold = True: Axn1.ColWidth(20) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Impto.Ret."
    Axn1.Col = 21: Axn1.CellFontBold = True: Axn1.ColWidth(21) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Cr.Pag."
    Axn1.Col = 22: Axn1.CellFontBold = True: Axn1.ColWidth(22) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Subdio.No apl."
    Axn1.Col = 23: Axn1.CellFontBold = True: Axn1.ColWidth(23) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "IMSS"
    Axn1.Col = 24: Axn1.CellFontBold = True: Axn1.ColWidth(24) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Prestamos"
    Axn1.Col = 25: Axn1.CellFontBold = True: Axn1.ColWidth(25) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Pensi�n Ali"
    Axn1.Col = 26: Axn1.CellFontBold = True: Axn1.ColWidth(26) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Fonacot"
    Axn1.Col = 27: Axn1.CellFontBold = True: Axn1.ColWidth(27) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Infonavit"
    ' ---- NUEVO: columnas LIQ y ESP ----
    Axn1.Col = 10: Axn1.CellFontBold = True: Axn1.ColWidth(10) = 1300: Axn1.CellAlignment = 3: Axn1.Text = "Compensaciones"
    Axn1.Col = 11: Axn1.CellFontBold = True: Axn1.ColWidth(11) = 1300: Axn1.CellAlignment = 3: Axn1.Text = "P. Antig�edad"
    Axn1.Col = 12: Axn1.CellFontBold = True: Axn1.ColWidth(12) = 1300: Axn1.CellAlignment = 3: Axn1.Text = "Indemnizaci�n"
    Axn1.Col = 13: Axn1.CellFontBold = True: Axn1.ColWidth(13) = 1600: Axn1.CellAlignment = 3: Axn1.Text = "Sueldos Vacaciones"
    Axn1.Col = 14: Axn1.CellFontBold = True: Axn1.ColWidth(14) = 1200: Axn1.CellAlignment = 3: Axn1.Text = "Bono"
    Axn1.Rows = 1
End Sub

Private Sub AuxEdCop_Click()
  Dim Temporal1
    Clipboard.Clear
    difer = Axn1.RowSel - Axn1.Row
    For i = Axn1.Row To Axn1.RowSel
        For f = 0 To Axn1.ColSel
            Temporal1 = Temporal1 + Axn1.TextMatrix(i, f)
            If f < Axn1.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
        Next f
        Temporal1 = Temporal1 & Chr(13) & Chr(10)
        Next i
        Clipboard.SetText Temporal1

End Sub

Private Sub EditSelTot_Click()
    Axn1.Row = 0
    Axn1.Col = 0
   Axn1.RowSel = Axn1.Rows - 1
   Axn1.ColSel = Axn1.Cols - 1
End Sub
Private Sub Form_Load()
    ' Limpiar la grilla antes de cargar nuevos datos
    Axn1.Clear
    Axn1.Rows = 1
    
    ' Reiniciar la colecci�n de nombres para evitar duplicados
    Set nombresAgregados = Nothing
    Set nombresAgregados = New Collection

    ' Inicializar variables
    'Axn1 = 1
    ReAper

    ' Asignar nombre
    AxNom1.Caption = Acu21.Dacu1.TextMatrix(Acu21.Dacu1.Row, 0) + " " + Acu21.Dacu1.TextMatrix(Acu21.Dacu1.Row, 1)

    ' Construir nombre de archivo seg�n el tipo
    If (Tipo = True) Then
        NoAch = "Auxiliar\Axna" + Trim(Acu21.Dacu1.TextMatrix(Acu21.Dacu1.Row, 0))
    Else
        NoAch = "Auxiliar\Axn" + LTrim(Acu21.Dacu1.TextMatrix(Acu21.Dacu1.Row, 0))
    End If

    Debug.Print "Archivo le�do: " & NoAch ' Para verificar que se est� leyendo el archivo correcto

    ' Cerrar y abrir archivo para lectura
    Close 7
    Open NoAch For Random As 7 Len = Len(AxNom)
    FinAx2 = LOF(7) / Len(AxNom)

    ' Cargar datos en la grilla
    Dim Nombre As String
    Dim repetido As Boolean

    For r = 1 To FinAx2
        Get 7, r, AxNom
        Nombre = Trim$(AxNom.Narch)
        If UCase$(Right$(Nombre, 4)) = ".NOM" Then Nombre = Left$(Nombre, Len(Nombre) - 4)

        ' Verificar si el nombre ya fue agregado
        repetido = False
        On Error Resume Next
        nombresAgregados.Item (Nombre)
        If Err.Number = 0 Then
            repetido = True
        End If
        On Error GoTo 0

        ' Agregar si no est� repetido
        If Not repetido Then
            nombresAgregados.Add Nombre, Nombre

            ' Calcular ingresos (NUEVO: incluye los conceptos LIQ / ESP)
            InGresos = IngresosTotalesDetalle(AxNom)

            ' Construir entrada para la grilla (columnas 0 a 22, igual que antes)
            entrada = Format(r, "###0") & Chr(9) & Nombre & Chr(9) & _
                      Format(AxNom.Pdias, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.Pnormal, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.Pextra, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.Pviaticos, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.Pvaca, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.Potras, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.Pagui, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.PPTU, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.Pexenta, "###,###,##0.00") & Chr(9) & _
                      Format(InGresos, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.DImpto, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.DSubioAp, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.DCrApl, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.DImpret, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.DCrPag, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.DSubNoap, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.DImss, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.DPrestamos, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.DTelefono, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.DTonacot, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.DOtrasded, "###,###,##0.00")

            ' NUEVO: columnas 23 a 27 (se arma en otra instrucci�n porque VB6
            ' permite m�ximo 24 continuaciones de l�nea "_" por instrucci�n)
            entrada = entrada & Chr(9) & _
                      Format(AxNom.PCompensa, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.PAntiguedad, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.PIndemniza, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.PSueldoVac, "###,###,##0.00") & Chr(9) & _
                      Format(AxNom.PBono, "###,###,##0.00")

            ' Agregar la entrada a la grilla
            Axn1.AddItem EntradaAxNomVisual(CStr(entrada))
        End If
    Next r

    ' Llamar a funciones adicionales
    cero_s
    Axn1.Rows = Axn1.Rows + 1
    sumas
End Sub


Private Sub Form_Resize()
    Axn1.Width = ScaleWidth
    Axn1.Height = ScaleHeight

End Sub
Sub cero_s()

   For r = 1 To Axn1.Rows - 1

       For i = 2 To ULT_COL

            If IsNumeric(Axn1.TextMatrix(r, i)) Then

                If Val(Axn1.TextMatrix(r, i)) = 0 Then

                    Axn1.TextMatrix(r, i) = ""

                End If

            End If

       Next i

   Next r

End Sub
Sub sumas()
    Dim sum As Currency
    Axn1.Row = Axn1.Rows - 1: Axn1.Col = 1: Axn1.CellFontBold = True
    Axn1.CellAlignment = 6: Axn1.TextMatrix((Axn1.Rows - 1), 1) = "Sumas  "
    For r = 2 To ULT_COL
       For i = 1 To Axn1.Rows - 2
          If IsNumeric(Axn1.TextMatrix(i, r)) Then sum = sum + CCur(Axn1.TextMatrix(i, r))
       Next i
       If sum <> 0 Then
            Axn1.Col = r
            Axn1.CellFontBold = True
            Axn1.TextMatrix((Axn1.Rows - 1), r) = Format(sum, "###,###,##0.00")
       Else
            Axn1.TextMatrix((Axn1.Rows - 1), r) = ""
       End If
       sum = 0
    Next r
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Clipboard.Clear
End Sub

' comentario