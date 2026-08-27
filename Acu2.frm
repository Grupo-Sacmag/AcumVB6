VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Acu21 
   Caption         =   "Acumulado de sueldos"
   ClientHeight    =   5235
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7065
   Icon            =   "Acu2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox ACUM 
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      Text            =   "Sopecitos"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid Dacu1 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   1
      Cols            =   27
      FixedCols       =   2
      BackColorBkg    =   -2147483633
   End
   Begin VB.Menu Arc 
      Caption         =   "&Archivo"
      Begin VB.Menu ArCamb 
         Caption         =   "&Cambio de Subdirectorio"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu ArcOrd 
         Caption         =   "&Ordenar"
         Begin VB.Menu ArOrAlf 
            Caption         =   "&Alfabetico"
         End
         Begin VB.Menu ArOrNum 
            Caption         =   "&Númerico"
         End
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu ArAcumulados 
         Caption         =   "Opciones de Acumulacion"
         Begin VB.Menu ArAcuTT 
            Caption         =   "&Total"
         End
         Begin VB.Menu ArNoincluye 
            Caption         =   "&No incluir la Ultima nomina de Diciembre"
         End
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu ArcImp 
         Caption         =   "&Impresión"
      End
      Begin VB.Menu ArcSep3 
         Caption         =   "-"
      End
      Begin VB.Menu ArSal 
         Caption         =   "&Salida"
      End
   End
   Begin VB.Menu Edi 
      Caption         =   "&Edición"
      Begin VB.Menu EditDupl 
         Caption         =   "&Buscar Duplicados"
      End
      Begin VB.Menu EditPer 
         Caption         =   "&Editar personal"
         Shortcut        =   ^E
      End
      Begin VB.Menu EdCop 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu Edsep1 
         Caption         =   "-"
      End
      Begin VB.Menu EDesc 
         Caption         =   "&Desactivar titulos"
         Shortcut        =   ^A
      End
      Begin VB.Menu EdAct 
         Caption         =   "&Activar titulos"
         Shortcut        =   ^B
      End
      Begin VB.Menu EdSep2 
         Caption         =   "-"
      End
      Begin VB.Menu EdSelt 
         Caption         =   "&Seleccionar todo"
      End
   End
   Begin VB.Menu ver 
      Caption         =   "&Ver"
      Begin VB.Menu acumTo 
         Caption         =   "&Acumulados Totales"
      End
      Begin VB.Menu VerAr 
         Caption         =   "&Archivos Acumulados"
         Shortcut        =   ^V
      End
      Begin VB.Menu Inicial 
         Caption         =   "&Mes Incial"
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   1
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   10
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   11
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   12
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   13
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   14
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   15
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   16
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   17
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   18
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   19
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   20
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   21
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   22
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   23
            Visible         =   0   'False
         End
         Begin VB.Menu mesInicial 
            Caption         =   "&mes"
            Index           =   24
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Final 
         Caption         =   "&Mes Final"
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   1
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   10
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   11
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   12
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   13
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   14
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   15
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   16
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   17
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   18
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   19
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   20
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   21
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   22
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   23
            Visible         =   0   'False
         End
         Begin VB.Menu mesFin 
            Caption         =   "&mes"
            Index           =   24
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu Info3 
      Caption         =   "&Informativa"
      Begin VB.Menu InfoGene 
         Caption         =   "&Generar Archivo"
      End
      Begin VB.Menu InfSep1 
         Caption         =   "-"
      End
      Begin VB.Menu InfoTring 
         Caption         =   "&Otros ingresos"
      End
   End
   Begin VB.Menu versionActual 
      Caption         =   "Versión"
   End
End
Attribute VB_Name = "Acu21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hoja As Integer
Dim Femp As Long, Femco As Long, FPer As Long, sum As Currency
Dim FiNax As Long, EXTRA As Integer, Arch1
Dim FAcum1 As Long, Fin_Otreg As Long, NoIncLuye As Integer
Dim busquedaInicial
Dim busquedaFinal
Dim nombreArchivo
Dim numeroArchivo
Dim filasTemporal
Dim Directorio1 As String
Dim Directorio2 As String
Sub EditREp()
   Dim CadAnt As String, CadActual As String
   CadAnt = Trim(Dacu1.TextMatrix(1, 1))
   For r = 2 To Dacu1.Rows - 2
        If r > (Dacu1.Rows - 2) Then Exit For
        CadActual = Trim(Dacu1.TextMatrix(r, 1))
        If CadAnt = CadActual Then
            Dacu1.Row = r: Dacu1.Col = 2
            MsgBox CadActual + " Desea Adicionar el Anterior al Actual", vbYesNo
            Dacu1_EnterCell
            Dacu1.TopRow = Dacu1.TopRow + (r - Dacu1.TopRow)
            If vbYes = 6 Then
                  Adicionar
                  Else
                  Rem nada
            End If
            CadAnt = CadActual: CadActual = ""
            Else
            CadAnt = CadActual: CadActual = ""
        End If
        
   Next r
   sumando
End Sub
Sub Adicionar()
Dim AntValor As Currency, ActValor As Currency
   For i = 2 To 22
          
         If IsNumeric(Dacu1.TextMatrix(r - 1, i)) Then
            AntValor = (Dacu1.TextMatrix(r - 1, i))
            
            Else
            AntValor = 0
         End If
         If IsNumeric(Dacu1.TextMatrix(r, i)) Then
                    ActValor = (Dacu1.TextMatrix(r, i))
                    ActValor = ActValor + AntValor
                    
                    Else
                    
                    ActValor = 0
                    ActValor = ActValor + AntValor
         End If
        If i = 2 Then
                Dacu1.TextMatrix(r, i) = Format(ActValor, "####0.00")
                
                Else
                Dacu1.TextMatrix(r, i) = Format(ActValor, "###,###,##0.00")
         End If
        
    Next i
 
   Dacu1.RemoveItem r - 1
   r = r - 1
End Sub
Sub inicio()

    On Error GoTo DirectorioNoDisponible

    Close 3

    Open "C:\GconTa\sccontr.soc" For Random As 3 Len = Len(SCont)

    Get 3, 1, SCont

    Close 3

    '------------------------------------------------------
    ' MISMO COMPORTAMIENTO QUE EL ORIGINAL
    '------------------------------------------------------
    If Trim$(SCont.guarda) <= " " Then

        ChDir "C:\GconTA"

    Else

        ' Intentar utilizar directamente la ruta guardada.
        '
        ' NO validamos archivos.
        ' Si existe, continúa normalmente.
        ' Si ya no existe, salta al manejador.
        ChDir Trim$(SCont.guarda)

    End If

    Exit Sub


DirectorioNoDisponible:

    Close 3
    Err.Clear

    MsgBox "El directorio guardado ya no está disponible." & _
           vbCrLf & vbCrLf & _
           "Es posible que haya sido movido, eliminado o renombrado." & _
           vbCrLf & _
           "Seleccione nuevamente el directorio de trabajo.", _
           vbExclamation, _
           "Directorio no disponible"

    Load Camdir
    Camdir.Show 1

    '------------------------------------------------------
    ' Camdir es modal.
    '
    ' Si el usuario guardó una nueva ruta, volvemos
    ' a ejecutar inicio para cargarla.
    '
    ' Si decidió cerrar la aplicación, Camdir hará End
    ' y nunca llegará a esta línea.
    '------------------------------------------------------
    inicio

End Sub

Sub Apertura()
    miarchivo = Dir("Auxiliar", vbDirectory)
    Rem ******   Verifica que Exista el archivo de auxiliares *****
    If miarchivo = "" Then
       Respuesta = MsgBox("No existe el Subdirectorio de Auxiliares desea crearlo ", vbYesNo, "Auxiliares ")
       If Respuesta = vbYes Then
              MkDir "AUXILIAR"
       End If
    End If
End Sub

Sub elecero()
   For r = 1 To Dacu1.Rows - 2
       For i = 2 To 22
         If Dacu1.TextMatrix(r, i) = "" Then Dacu1.TextMatrix(r, i) = 0
          If Dacu1.TextMatrix(r, i) = 0 Then
             Dacu1.TextMatrix(r, i) = ""
          End If
       Next i
   Next r
   sumando
End Sub
Sub sumando()
    Dacu1.Row = Dacu1.Rows - 1: Dacu1.Col = 1: Dacu1.CellFontBold = True
    Dacu1.CellAlignment = 6: Dacu1.TextMatrix((Dacu1.Rows - 1), 1) = "Sumas  "
    For r = 2 To 22
       For i = 1 To Dacu1.Rows - 2
          If Dacu1.TextMatrix(i, r) = "" Then
             Dacu1.TextMatrix(i, r) = ""
             Else
             sum = sum + Dacu1.TextMatrix(i, r)
          End If
       Next i
       If sum <> 0 Then
            Dacu1.Col = r
            Dacu1.CellFontBold = True
            Dacu1.TextMatrix((Dacu1.Rows - 1), r) = Format(sum, "###,###,##0.00")
            Else
            Dacu1.TextMatrix((Dacu1.Rows - 1), r) = ""
       End If
       sum = 0
   Next r
End Sub
Sub encabezado()

    Dacu1.Row = 0
    Dacu1.Col = 0: Dacu1.CellFontBold = True: Dacu1.ColWidth(0) = 600:  Dacu1.CellAlignment = 3: Dacu1.Text = "Núm."
    Dacu1.ColAlignment(0) = 4
    Dacu1.Col = 1: Dacu1.CellFontBold = True: Dacu1.ColWidth(1) = 3200: Dacu1.CellAlignment = 3: Dacu1.Text = "Nómbre."
    Dacu1.Col = 2: Dacu1.CellFontBold = True: Dacu1.ColWidth(2) = 1600:  Dacu1.CellAlignment = 3: Dacu1.Text = "Dias"
    Dacu1.Col = 3: Dacu1.CellFontBold = True: Dacu1.ColWidth(3) = 1600:  Dacu1.CellAlignment = 3: Dacu1.Text = "Sueldos"
    Dacu1.Col = 4: Dacu1.CellFontBold = True: Dacu1.ColWidth(4) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Prem.Punt"
    Dacu1.Col = 5: Dacu1.CellFontBold = True: Dacu1.ColWidth(5) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "O.F."
    Dacu1.Col = 6: Dacu1.CellFontBold = True: Dacu1.ColWidth(6) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "P.Vacac."
    Dacu1.Col = 7: Dacu1.CellFontBold = True: Dacu1.ColWidth(7) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Otras"
    Dacu1.Col = 8: Dacu1.CellFontBold = True: Dacu1.ColWidth(8) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Aguinaldo"
    Dacu1.Col = 9: Dacu1.CellFontBold = True: Dacu1.ColWidth(9) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "P.T.U."
    Dacu1.Col = 10: Dacu1.CellFontBold = True: Dacu1.ColWidth(10) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Exentos"
    Dacu1.Col = 11: Dacu1.CellFontBold = True: Dacu1.ColWidth(11) = 1600:  Dacu1.CellAlignment = 3: Dacu1.Text = "Total"
    Dacu1.Col = 12: Dacu1.CellFontBold = True: Dacu1.ColWidth(12) = 1600:  Dacu1.CellAlignment = 3: Dacu1.Text = "Impto."
    Dacu1.Col = 13: Dacu1.CellFontBold = True: Dacu1.ColWidth(13) = 1600:  Dacu1.CellAlignment = 3: Dacu1.Text = "Subdio.Apl."
    Dacu1.Col = 14: Dacu1.CellFontBold = True: Dacu1.ColWidth(14) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Cr.Apl."
    Dacu1.Col = 15: Dacu1.CellFontBold = True: Dacu1.ColWidth(15) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Impto.Ret."
    Dacu1.Col = 16: Dacu1.CellFontBold = True: Dacu1.ColWidth(16) = 1600:  Dacu1.CellAlignment = 3: Dacu1.Text = "Cr.Pag."
    Dacu1.Col = 17: Dacu1.CellFontBold = True: Dacu1.ColWidth(17) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Subdio.No apl."
    Dacu1.Col = 18: Dacu1.CellFontBold = True: Dacu1.ColWidth(18) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "IMSS"
    Dacu1.Col = 19: Dacu1.CellFontBold = True: Dacu1.ColWidth(19) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Prestamos"
    Dacu1.Col = 20: Dacu1.CellFontBold = True: Dacu1.ColWidth(20) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Pensión Ali"
    Dacu1.Col = 21: Dacu1.CellFontBold = True: Dacu1.ColWidth(21) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Fonacot"
    Dacu1.Col = 22: Dacu1.CellFontBold = True: Dacu1.ColWidth(22) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Infonavit"
    Dacu1.Col = 23: Dacu1.CellFontBold = True: Dacu1.ColWidth(23) = 1500:  Dacu1.CellAlignment = 3: Dacu1.Text = "Rfc"
    Dacu1.Col = 24: Dacu1.CellFontBold = True: Dacu1.ColWidth(24) = 1500:  Dacu1.CellAlignment = 3: Dacu1.Text = "Imss"
    Dacu1.Col = 25: Dacu1.CellFontBold = True: Dacu1.ColWidth(25) = 1500:  Dacu1.CellAlignment = 3: Dacu1.Text = "Fecha Alta"
    Dacu1.Col = 26: Dacu1.CellFontBold = True: Dacu1.ColWidth(26) = 1500:  Dacu1.CellAlignment = 3: Dacu1.Text = "Fecha Baja"

End Sub

Private Sub acumTo_Click()
  AxNom1.tipo = False
  Dacu1.Rows = 1
  Form_Load
End Sub

Private Sub ArAcuTT_Click()
    NoIncLuye = 0
    Dacu1.Clear
    Dacu1.Rows = 1
    Close
    Kill "Auxiliar\*.*"
    Acumula_todo
    
    NoIncLuye = 0
    encabezado
    elecero
End Sub

Private Sub ArCamb_Click()
    Load Camdir
    Acu21.Caption = "Acumulado de sueldos"
    Camdir.Show 1
    Dacu1.Clear
    Dacu1.Rows = 1
    
    Form_Load
End Sub

Private Sub patasdecabra()
  mesInicial_Click (1)
  mesFin_Click (24)
End Sub


Private Sub ACUM_Click()
    Dim Respuesta As Integer

    ' Asegurar que "case" es la predeterminada
    'Este if solamente se agrega cuando se inicia la app, pues SELECCIONA UNA OPCION es con la que inicia la app
    If ACUM.Text = "Selecciona una opcion" Then
        
    Select Case ACUM.Text
        Case "Selecciona una opcion"
            Directorio1 = ("PATA DE POLLO")
            Directorio2 = ("*.NOM")
            MsgBox "Acumulado General", vbInformation, "Mensaje"
    End Select
    
    'Si es diferente a la opcion predeterminada, empieza a rellenar Directorio1 y Directorio2 con
    Else
        Respuesta = MsgBox(ACUM.Text + " ¿Estás seguro que quieres esta opción?", vbYesNo + vbQuestion, "Confirmación")
    
        If Respuesta = vbNo Then Exit Sub ' Si el usuario cancela, salir de la función

    Select Case ACUM.Text

        Case "GENERAL"
        
            AxNom1.tipo = False
            Dacu1.Rows = 1
            Form_Load
                                  
                        
        Case "1 ENE-FEB"

            Dacu1.Rows = 1
            Directorio1 = ("ENE*.NOM")
            Directorio2 = ("FEB*.NOM")
            MsgBox "Acumulado del Primer bimestre", vbInformation, "Mensaje"
   
   
            Rem Acumula_todo
            Rem encabezado
            Rem elecero
            
            mesInicial_Click (1)
            mesFin_Click (4)
            
        Case "2 MAR-ABR"

            Dacu1.Rows = 1
            Directorio1 = ("MAR*.NOM")
            Directorio2 = ("ABR*.NOM")
            MsgBox "Acumulado del Segundo bimestre", vbInformation, "Mensaje"
           
            Rem Acumula_todo
            Rem encabezado
            Rem elecero
            
           
            mesInicial_Click (5)
            mesFin_Click (8)

        Case "3 MAY-JUN"

            Dacu1.Rows = 1
            Directorio1 = ("MAY*.NOM")
            Directorio2 = ("JUN*.NOM")
            MsgBox "Acumulado del Tercer bimestre", vbInformation, "Mensaje"
              
            Rem Acumula_todo
            Rem encabezado
            Rem elecero
            

            mesInicial_Click (9)
            mesFin_Click (12)
            
            
        Case "4 JUL-AGO"

            Dacu1.Rows = 1
            Directorio1 = ("JUL*.NOM")
            Directorio2 = ("AGO*.NOM")
            MsgBox "Acumulado del Cuarto bimestre", vbInformation, "Mensaje"
          
            Rem Acumula_todo
            Rem encabezado
            Rem elecero
            
            mesInicial_Click (13)
            mesFin_Click (16)
            
        Case "5 SEP-OCT"

            Dacu1.Rows = 1
            Directorio1 = ("SEP*.NOM")
            Directorio2 = ("OCT*.NOM")
            MsgBox "Acumulado del Quinto bimestre", vbInformation, "Mensaje"
   
            Rem Acumula_todo
            Rem encabezado
            Rem elecero
            
            mesInicial_Click (17)
            mesFin_Click (20)

        Case "6 NOV-DIC"

            Dacu1.Rows = 1
            Directorio1 = ("NOV*.NOM")
            Directorio2 = ("DIC*.NOM")
            MsgBox "Acumulado del Sexto bimestre", vbInformation, "Mensaje"
   
            Rem Acumula_todo
            Rem encabezado
            Rem elecero
            
            mesInicial_Click (21)
            mesFin_Click (24)
            
    End Select
    End If
End Sub
Sub lectura()
    Apertura
    Dim miarchivo As String
    r = 1
    Close 10
    


    ' Abrir el archivo temporal
    Open "AcuTemp" For Random As 10 Len = Len(temporal)
    Ftem = LOF(10) / Len(temporal)

    ' Si el archivo temporal ya tiene datos, eliminarlo
    If Ftem > 0 Then
        Close 10
        Ftem = 0
        Kill "AcuTemp"
        Open "AcuTemp" For Random As 10 Len = Len(temporal)
        On Error GoTo sigue5
        Kill "Auxiliar\."
sigue5:
    End If

    ' Procesar archivos del primer directorio (Directorio1)
    miarchivo = Dir(Directorio1) ' Obtener el primer archivo del patrón
    Do While miarchivo <> ""
        Ftem = Ftem + 1
        temporal.miarchivo = miarchivo
        Put 10, Ftem, temporal
        miarchivo = Dir() ' Obtener el siguiente archivo
    Loop

    ' Procesar archivos del segundo directorio (Directorio2)
    miarchivo = Dir(Directorio2) ' Obtener el primer archivo del patrón
    Do While miarchivo <> ""
        Ftem = Ftem + 1
        temporal.miarchivo = miarchivo
        Put 10, Ftem, temporal
        miarchivo = Dir() ' Obtener el siguiente archivo
    Loop

    Close 10
    ordenxmes
End Sub

Sub ordenxmes()
 registro = 1
 Open "AcuTemp" For Random As 10 Len = Len(temporal)
 Ftem = LOF(10) / Len(temporal)
 mirar = registro
 ultimo.texto = mm(1)
 For r = 1 To 12
     Get 10, registro, temporal
     miarch1 = temporal.miarchivo
     For i = mirar To Ftem: Get 10, i, temporal
         If Left(temporal.miarchivo, 3) = Left(mm(r), 3) Then
                miarch2 = temporal.miarchivo
                temporal.miarchivo = miarch1
                Put 10, i, temporal
                temporal.miarchivo = miarch2
                Put 10, registro, temporal
                registro = registro + 1
                Get 10, registro, temporal
                miarch1 = temporal.miarchivo
                miarch2 = ""
                ultimo.texto1 = mm(r)
         End If
     Next i
 Next r
 Close 10
End Sub
Sub abre()

    Close 1
    Open "Empresa.Dno" For Random As 1 Len = Len(empresa)
    
    Femp = LOF(1) / Len(empresa)
      
    If Femp < 1 Then
        MsgBox "No existen archivos de nomina" & Chr(13) & "cambie el subdirectorio"
        Close
        Exit Sub
    End If
    
    Get 1, Femp, empresa
    Acu21.Caption = Acu21.Caption + " " + RTrim(empresa.name) + " " + RTrim(empresa.ao)
    
    Close 2
    
    Open "EmpComp.dno" For Random As 2 Len = Len(Dat_ide)
    Femco = LOF(2) / Len(Dat_ide)
    
    Close 3
    Open "Personal.dno" For Random As 3 Len = Len(personal)
    FPer = LOF(3) / Len(personal)

End Sub
Sub reconocedornomina()
   EXTRA = 1
   If Left(UCase(Arch1), 3) = "ENE" Then EXTRA = 0
   If Left(UCase(Arch1), 3) = "FEB" Then EXTRA = 0
   If Left(UCase(Arch1), 3) = "MAR" Then EXTRA = 0
   If Left(UCase(Arch1), 3) = "ABR" Then EXTRA = 0
   If Left(UCase(Arch1), 3) = "MAY" Then EXTRA = 0
   If Left(UCase(Arch1), 3) = "JUN" Then EXTRA = 0
   If Left(UCase(Arch1), 3) = "JUL" Then EXTRA = 0
   If Left(UCase(Arch1), 3) = "AGO" Then EXTRA = 0
   If Left(UCase(Arch1), 3) = "SEP" Then EXTRA = 0
   If Left(UCase(Arch1), 3) = "OCT" Then EXTRA = 0
   If Left(UCase(Arch1), 3) = "NOV" Then EXTRA = 0
   If Left(UCase(Arch1), 3) = "DIC" Then EXTRA = 0
   If (NoIncLuye = 1) And (Left(UCase(Arch1), 4)) = "DIC2" Then
            EXTRA = 3
            
   End If
End Sub

Sub Acumula_todo()
'On Error GoTo manejo1
   lectura
   Open "AcuTemp" For Random As 10 Len = Len(temporal)
   Ftem = LOF(10) / Len(temporal)
   Close 5
   Close 7
   
   Open "personal.dno" For Random As 13 Len = Len(personal)
   Open "AcuNom2.dno" For Random As 5 Len = Len(ArAcum)
   facum = LOF(5) / Len(ArAcum)
      
    For r = 1 To FPer
       ceros
       Put 5, r, ArAcum
    Next r
   
    For r = 1 To Ftem: Get 10, r, temporal
    Rem On Error GoTo manejo1
       temporal.miarchivo = Trim(temporal.miarchivo)
       Open temporal.miarchivo For Random As 4 Len = Len(nomina)
       fnom = LOF(4) / Len(nomina)
       Arch1 = Trim(temporal.miarchivo)
       Arch1 = Mid(Arch1, 1, Len(Arch1) - 4) + ".cmp"
       Open Arch1 For Random As 6 Len = Len(Nom_Com)
       FNomcom = LOF(6) / Len(Nom_Com)
       Debug.Print Arch1
       reconocedornomina
       If EXTRA = 3 Then
            Rem GoTo manejo1
            Close 4, 6
        Else
         For i = 1 To fnom
                Get 4, i, nomina
                Get 6, i, Nom_Com
                Get 13, i, personal
                InGresos = nomina.sueldo + nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri + nomina.aguin + nomina.viaticos + nomina.pvac + nomina.otras + nomina.exentos + nomina.ptu
                If InGresos <> 0 Then
                    
                    If EXTRA = 1 Then
                        Rem If (Val(Mid(LTrim(personal.fab), 7, 4)) > 0) And (Val(Mid(LTrim(personal.fab), 7, 4)) <= empresa.ao) Then
                            Nom_Com.CredNe = 0: Nom_Com.CreTot = 0
                            Nom_Com.ImpTot = 0: Nom_Com.subapl = 0
                            Nom_Com.subdio = 0: Nom_Com.subNap = 0
                            Nom_Com.ImpTot = nomina.ispt
                            Rem Put 6, I, Nom_Com
                        Rem End If
                        Rem If (Val(Mid(LTrim(personal.fal), 7, 4)) > 0) And (Val(Mid(LTrim(personal.fal), 7, 4)) = empresa.ao) Then
                            Rem Nom_Com.CredNe = 0: Nom_Com.CreTot = 0
                            Rem Nom_Com.ImpTot = 0: Nom_Com.subapl = 0
                            Rem Nom_Com.subdio = 0: Nom_Com.subNap = 0
                            Rem Nom_Com.ImpTot = nomina.ispt
                            Rem Put 6, I, Nom_Com
                        Rem End If
                        Rem If UCase(Left(LTrim(Arch1), 3)) = "PTU" Then
                            Rem Nom_Com.CredNe = 0: Nom_Com.CreTot = 0
                            Rem Nom_Com.ImpTot = 0: Nom_Com.subapl = 0
                            Rem Nom_Com.subdio = 0: Nom_Com.subNap = 0
                            Rem Nom_Com.ImpTot = nomina.ispt
                            Rem Put 6, I, Nom_Com
                        Rem End If
                    End If
                    Get 5, i, ArAcum
                    NoAux = "Auxiliar\AxN" + LTrim(Str(i))
                    Open NoAux For Random As 7 Len = Len(AxNom)
                    FiNax = LOF(7) / Len(AxNom)
                    FiNax = FiNax + 1
                    ArAcum.Pdias = ArAcum.Pdias + nomina.dias
                    ArAcum.Pnormal = ArAcum.Pnormal + nomina.sueldo
                    If RTrim(Left(temporal.miarchivo, 3)) = UCase("PRE") Then
                        
                        nomina.hs_tri = nomina.viaticos: nomina.viaticos = 0
                    End If
                    ArAcum.Pextra = ArAcum.Pextra + (nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri)
                    ArAcum.Pviaticos = ArAcum.Pviaticos + nomina.viaticos
                    
                    ArAcum.Pagui = ArAcum.Pagui + nomina.aguin
                    ArAcum.Pvaca = ArAcum.Pvaca + nomina.pvac
                    ArAcum.Potras = ArAcum.Potras + nomina.otras
                    ArAcum.PPTU = ArAcum.PPTU + nomina.ptu
                    ArAcum.Pexenta = ArAcum.Pexenta + nomina.exentos
                    ArAcum.DImpto = ArAcum.DImpto + Nom_Com.ImpTot
                    ArAcum.DSubioAp = ArAcum.DSubioAp + Nom_Com.subapl
                    ArAcum.DCrApl = ArAcum.DCrApl + Nom_Com.CreTot
                    ArAcum.DImpret = ArAcum.DImpret + nomina.ispt
                    ArAcum.DCrPag = ArAcum.DCrPag + nomina.crdsal
                    ArAcum.DSubNoap = ArAcum.DSubNoap + Nom_Com.subNap
                    ArAcum.DImss = ArAcum.DImss + nomina.imss
                    ArAcum.DPrestamos = ArAcum.DPrestamos + nomina.prestamos
                    ArAcum.DTonacot = ArAcum.DTonacot + nomina.fonacot
                    ArAcum.DTelefono = ArAcum.DTelefono + nomina.telefono
                    ArAcum.DOtrasded = ArAcum.DOtrasded + nomina.otraded
                    Put 5, i, ArAcum
                    Auxliar
                    Put 7, FiNax, AxNom
                    
                    Close 7
                    AxPer.Pdias = nomina.dias
                    AxPer.Pnormal = nomina.sueldo
                    AxPer.Pextra = (nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri)
                    AxPer.Pviaticos = nomina.viaticos
                    AxPer.Pvaca = nomina.pvac
                    AxPer.Pagui = nomina.aguin
                    AxPer.Potras = nomina.otras
                    AxPer.PPTU = nomina.ptu
                    AxPer.Pexenta = AxPer.Pexenta + nomina.exentos
                    AxPer.DImpto = AxPer.DImpto + Nom_Com.ImpTot
                    AxPer.DSubioAp = AxPer.DSubioAp + Nom_Com.subapl
                    AxPer.DCrApl = AxPer.DCrApl + Nom_Com.CreTot
                    AxPer.DImpret = AxPer.DImpret + nomina.ispt
                    AxPer.DCrPag = AxPer.DCrPag + nomina.crdsal
                    AxPer.DSubNoap = AxPer.DSubNoap + Nom_Com.subNap
                    AxPer.DImss = AxPer.DImss + nomina.imss
                    AxPer.DPrestamos = AxPer.DPrestamos + nomina.prestamos
                    AxPer.DTonacot = AxPer.DTonacot + nomina.fonacot
                    AxPer.DTelefono = AxPer.DTelefono + nomina.telefono
                    AxPer.DOtrasded = AxPer.DOtrasded + nomina.otraded
                    
              End If
          Next i
          Close 4, 6
        End If
    Next r
    facum = LOF(5) / Len(ArAcum)
    Rem If NoIncLuye = 1 Then Stop
    For r = 1 To facum: Get 5, r, ArAcum
            Get 13, r, personal
            InGresos = ArAcum.Pnormal + ArAcum.Pextra + ArAcum.Pviaticos + ArAcum.Pvaca + ArAcum.Potras + ArAcum.Pagui + ArAcum.PPTU + ArAcum.Pexenta
            If InGresos <> 0 Then
            nombre = RTrim(personal.ape1) + " " + RTrim(personal.ape2) + " " + RTrim(personal.nom)
            entrada = Format(r, "###0") & Chr(9) & nombre & Chr(9) & Format(ArAcum.Pdias, "###,###,##0.00") & Chr(9) & Format(ArAcum.Pnormal, "###,###,##0.00") & Chr(9) & _
                      Format(ArAcum.Pextra, "###,###,##0.00") & Chr(9) & _
                      Format(ArAcum.Pviaticos, "###,###,##0.00") & Chr(9) & Format(ArAcum.Pvaca, "###,###,##0.00") & Chr(9) & Format(ArAcum.Potras, "###,###,##0.00") & Chr(9) & _
                      Format(ArAcum.Pagui, "###,###,##0.00") & Chr(9) & _
                      Format(ArAcum.PPTU, "###,###,##0.00") & Chr(9) & Format(ArAcum.Pexenta, "###,###,##0.00") & Chr(9) & _
                      Format(InGresos, "###,###,##0.00") & Chr(9) & Format(ArAcum.DImpto, "###,###,##0.00") _
                      & Chr(9) & Format(ArAcum.DSubioAp, "###,###,##0.00") & Chr(9) & Format(ArAcum.DCrApl, "###,###,##0.00") _
                      & Chr(9) & Format(ArAcum.DImpret, "###,###,##0.00") & Chr(9) & Format(ArAcum.DCrPag, "###,###,##0.00") _
                      & Chr(9) & Format(ArAcum.DSubNoap, "###,###,##0.00") & Chr(9) & Format(ArAcum.DImss, "###,###,##0.00") _
                      & Chr(9) & Format(ArAcum.DPrestamos, "###,###,##0.00") & Chr(9) & Format(ArAcum.DTelefono, "###,###,##0.00") _
                      & Chr(9) & Format(ArAcum.DTonacot, "###,###,##0.00") & Chr(9) & Format(ArAcum.DOtrasded, "###,###,##0.00") _
                      & Chr(9) & personal.rfc & Chr(9) & ("'" + personal.imss) & Chr(9) & (" " + personal.fal) & Chr(9) & (" " + personal.fab)
                      
                      
            Dacu1.AddItem entrada
            End If
    Next r
    Dacu1.Rows = Dacu1.Rows + 1
    Close: Exit Sub
'manejo1:
   'MsgBox "Falla al Acumular " & Chr(13) & "revisa archivos  *.Nom", vbCritical
 '   MsgBox (Err.Number & " " & Err.Description)
 'Close
End Sub
Sub Auxliar()
    
    AxNom.Narch = temporal.miarchivo
    AxNom.Pdias = nomina.dias
    AxNom.Pnormal = nomina.sueldo
    AxNom.Pextra = (nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri)
    AxNom.Pviaticos = nomina.viaticos
    AxNom.Pvaca = nomina.pvac
    AxNom.Pagui = nomina.aguin
    AxNom.Potras = nomina.otras
    AxNom.PPTU = nomina.ptu
    AxNom.Pexenta = nomina.exentos
    AxNom.DImpto = Nom_Com.ImpTot
    AxNom.DSubioAp = Nom_Com.subapl
    AxNom.DCrApl = Nom_Com.CreTot
    AxNom.DImpret = nomina.ispt
    AxNom.DCrPag = nomina.crdsal
    AxNom.DSubNoap = Nom_Com.subNap
    AxNom.DImss = nomina.imss
    AxNom.DPrestamos = nomina.prestamos
    AxNom.DTonacot = nomina.fonacot
    AxNom.DTelefono = nomina.telefono
    AxNom.DOtrasded = nomina.otraded

End Sub

Sub ceros()
    ArAcum.Pdias = 0
    ArAcum.Pnormal = 0
    ArAcum.Pextra = 0
    ArAcum.Pviaticos = 0
    ArAcum.Pvaca = 0
    ArAcum.Potras = 0
    ArAcum.PPTU = 0
    ArAcum.Pexenta = 0
    ArAcum.DImpto = 0
    ArAcum.DSubioAp = 0
    ArAcum.DCrApl = 0
    ArAcum.DImpret = 0
    ArAcum.DCrPag = 0
    ArAcum.DSubNoap = 0
    ArAcum.DImss = 0
    ArAcum.DPrestamos = 0
    ArAcum.DTonacot = 0
    ArAcum.DTelefono = 0
    ArAcum.DOtrasded = 0
    ArAcum.Pagui = 0

End Sub

Sub Titul_Imp()
    hoja = hoja + 1
    Titulo = RTrim(empresa.name)
    Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(Titulo) / 2)
    Printer.Print Titulo
    Titulo = "Acumulado de Sueldos de " + RTrim(ultimo.texto) + " a " + RTrim(ultimo.texto1) + " de " + RTrim(empresa.ao)
    Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(Titulo) / 2)
    Printer.Print Titulo
    Titulo = "Hoja.. " + Str(hoja)
    Printer.CurrentX = 10000 - Printer.TextWidth(Titulo)
    Printer.Print Titulo
    Printer.Line (1200, Printer.CurrentY)-(10000, (Printer.CurrentY + 20)), , BF
    Printer.Print
    
     
End Sub
Private Sub ArcImp_Click()
    ultimo.impresion = Printer.FontSize
    Printer.FontSize = 10
    If Dacu1.ColSel = 0 Then
       ultimo.ColIni = 1
       ultimo.ColFin = Dacu1.Cols - 1
       ultimo.RenIni = 1
       ultimo.RenFin = Dacu1.Rows - 2
       Else
       ultimo.ColIni = Dacu1.Col
       ultimo.ColFin = Dacu1.ColSel
       ultimo.RenIni = Dacu1.Row
       ultimo.RenFin = Dacu1.RowSel
    End If
    hoja = 0
    Titul_Imp
    For r = ultimo.RenIni To ultimo.RenFin
        For c = 0 To ultimo.ColFin
            Select Case c
                 Case 0
                  Printer.CurrentX = 1200
                  Printer.Print "Número ";
                  Printer.Print Format(Dacu1.TextMatrix(r, c), "####0")
                  Case 1, 23, 24
                  Printer.CurrentX = 1200
                  Printer.Print Dacu1.TextMatrix(0, c);
                  Printer.Print (" " + RTrim(Dacu1.TextMatrix(r, c)))
                  Case 2 To 22
                  If Dacu1.TextMatrix(r, c) <> "" Then
                      Printer.CurrentX = 2400
                      Printer.Print Dacu1.TextMatrix(0, c);
                      Balor = Dacu1.TextMatrix(r, c)
                      valor$ = Balor: ancho2 = 0
                      colocar ancho2, valor$, "###,###,##0.00"
                      Printer.CurrentX = 7000 + ancho2
                      Printer.Print Format(Dacu1.TextMatrix(r, c), "###,###,##0.00")

                  End If
            End Select
            If Printer.CurrentY > Printer.ScaleHeight - 3000 Then
                Printer.NewPage
                Titul_Imp
            End If
        Next c
           Printer.Line (1200, Printer.CurrentY)-(10000, (Printer.CurrentY + 20)), , BF
           Printer.Print
    Next r
    Printer.EndDoc
    Printer.FontSize = ultimo.impresion
End Sub

Private Sub ArNoincluye_Click()
    NoIncLuye = 1
    Dacu1.Clear
    Dacu1.Rows = 1
    Close
    Kill "Auxiliar\*.*"
    Acumula_todo
    NoIncLuye = 0
    encabezado
    elecero
End Sub

Private Sub ArOrAlf_Click()
    
    colanti = Dacu1.Col
    renati = Dacu1.Row
    Dacu1.Row = 1
    Dacu1.Col = 1
    Dacu1.RowSel = Dacu1.Rows - 2
    Dacu1.Sort = 1
    Dacu1.Col = colanti
    Dacu1.Row = renati
     
End Sub

Private Sub ArOrNum_Click()
    colanti = Dacu1.Col
    renati = Dacu1.Row
    Dacu1.Row = 1
    Dacu1.Col = 0
    Dacu1.RowSel = Dacu1.Rows - 2
    Dacu1.Sort = 1
    Dacu1.Col = colanti
    Dacu1.Row = renati

End Sub

Private Sub ArSal_Click()
 Close: End
End Sub

Private Sub Dacu1_EnterCell()
  If ((Dacu1.Row > 0) And (Dacu1.Col > 1)) Then
             Dacu1.CellBackColor = vbYellow
  End If
End Sub

Private Sub Dacu1_KeyDown(KeyCode As Integer, Shift As Integer)
      If vbKeyF2 Then
        If Dacu1.Col = 23 Then
             Text1.Text = RTrim(Dacu1.TextMatrix(Dacu1.Row, 23))
             Text1.SetFocus
        End If
      End If
End Sub

Private Sub Dacu1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      AxNom1.Show
   End If
End Sub

Private Sub Dacu1_LeaveCell()
    If ((Dacu1.Row > 0) And (Dacu1.Col > 1)) Then
        Dacu1.CellBackColor = vbWhite
    End If
End Sub


Private Sub EdAct_Click()
    Dacu1.FixedRows = 1
    Dacu1.FixedCols = 2
End Sub

Private Sub EdCop_Click()
    Dim Temporal1 As String
    Clipboard.Clear
    Dim difer As Integer
    difer = Dacu1.RowSel - Dacu1.Row

    For i = Dacu1.Row To Dacu1.RowSel
        For f = 0 To Dacu1.ColSel
            Temporal1 = Temporal1 & CStr(Dacu1.TextMatrix(i, f)) ' Convertir explícitamente a cadena
            If f < Dacu1.ColSel Then
                Temporal1 = Temporal1 & Chr(9) ' Separador de columnas (Tab)
            End If
        Next f
        Temporal1 = Temporal1 & Chr(13) & Chr(10) ' Nueva línea
    Next i

    Clipboard.SetText Temporal1
End Sub


Private Sub EDesc_Click()
    Dacu1.FixedRows = 0
    Dacu1.FixedCols = 0
End Sub

Private Sub EditDupl_Click()
      EditREp
End Sub

Private Sub EditPer_Click()
   Close 2
   DatosConta.Show 1
   Dacu1.TextMatrix(Dacu1.Row, 1) = RTrim(personal.ape1) + " " + RTrim(personal.ape2) _
                                   + " " + RTrim(personal.nom)
   Dacu1.TextMatrix(Dacu1.Row, 23) = personal.rfc
   Close 2
End Sub

Private Sub EdSelt_Click()
   Dacu1.Col = 0
   Dacu1.Row = 0
   Dacu1.ColSel = Dacu1.Cols - 1
   Dacu1.RowSel = Dacu1.Rows - 1
End Sub
Private Sub Form_Load()
    Close 1: Open "AcuTemp" For Random As 1 Len = Len(temporal)
    mlen = LOF(1) / Len(temporal)
    
    ' Ejecutar la lógica correspondiente a "ANUAL"
    ' Agregar opciones al ComboBox ACUM
        ACUM.Text = "Selecciona una opcion"
        ACUM.AddItem "GENERAL"
        ACUM.AddItem "1 ENE-FEB"
        ACUM.AddItem "2 MAR-ABR"
        ACUM.AddItem "3 MAY-JUN"
        ACUM.AddItem "4 JUL-AGO"
        ACUM.AddItem "5 SEP-OCT"
        ACUM.AddItem "6 NOV-DIC"
      ACUM_Click
    For g = 1 To mlen
        Get 1, g, temporal
        nombreArchivo = Left(Trim(temporal.miarchivo), 6)
        If (nombreArchivo = "ENE120") Then
            mesInicial.Item(1).Caption = "ENE1raQuin"
            mesFin.Item(1).Caption = "ENE1raQuin"
        End If
        If (nombreArchivo = "ENE220") Then
            mesInicial.Item(2).Caption = "ENE2daQuin"
            mesFin.Item(2).Caption = "ENE2daQuin"
            mesInicial.Item(2).Visible = True
            mesFin.Item(2).Visible = True
        End If
        If (nombreArchivo = "FEB120") Then
            mesInicial.Item(3).Caption = "FEB1raQuin"
            mesFin.Item(3).Caption = "FEB1raQuin"
            mesInicial.Item(3).Visible = True
            mesFin.Item(3).Visible = True
        End If
        If (nombreArchivo = "FEB220") Then
            mesInicial.Item(4).Caption = "FEB2daQuin"
            mesFin.Item(4).Caption = "FEB2daQuin"
            mesInicial.Item(4).Visible = True
            mesFin.Item(4).Visible = True
        End If
        If (nombreArchivo = "MAR120") Then
            mesInicial.Item(5).Caption = "MAR1raQuin"
            mesFin.Item(5).Caption = "MAR1raQuin"
            mesInicial.Item(5).Visible = True
            mesFin.Item(5).Visible = True
        End If
        If (nombreArchivo = "MAR220") Then
            mesInicial.Item(6).Caption = "MAR2daQuin"
            mesFin.Item(6).Caption = "MAR2daQuin"
            mesInicial.Item(6).Visible = True
            mesFin.Item(6).Visible = True
        End If
        If (nombreArchivo = "ABR120") Then
            mesInicial.Item(7).Caption = "ABR1raQuin"
            mesFin.Item(7).Caption = "ABR1raQuin"
            mesInicial.Item(7).Visible = True
            mesFin.Item(7).Visible = True
        End If
        If (nombreArchivo = "ABR220") Then
            mesInicial.Item(8).Caption = "ABR2daQuin"
            mesFin.Item(8).Caption = "ABR2daQuin"
            mesInicial.Item(8).Visible = True
            mesFin.Item(8).Visible = True
        End If
        If (nombreArchivo = "MAY120") Then
            mesInicial.Item(9).Caption = "MAY1raQuin"
            mesFin.Item(9).Caption = "MAY1raQuin"
            mesInicial.Item(9).Visible = True
            mesFin.Item(9).Visible = True
        End If
        If (nombreArchivo = "MAY220") Then
            mesInicial.Item(10).Caption = "MAY2daQuin"
            mesFin.Item(10).Caption = "MAY2daQuin"
            mesInicial.Item(10).Visible = True
            mesFin.Item(10).Visible = True
        End If
        If (nombreArchivo = "JUN120") Then
            mesInicial.Item(11).Caption = "JUN1raQuin"
            mesFin.Item(11).Caption = "JUN1raQuin"
            mesInicial.Item(11).Visible = True
            mesFin.Item(11).Visible = True
        End If
        If (nombreArchivo = "JUN220") Then
            mesInicial.Item(12).Caption = "JUN2daQuin"
            mesFin.Item(12).Caption = "JUN2daQuin"
            mesInicial.Item(12).Visible = True
            mesFin.Item(12).Visible = True
        End If
        If (nombreArchivo = "JUL120") Then
            mesInicial.Item(13).Caption = "JUL1raQuin"
            mesFin.Item(13).Caption = "JUL1raQuin"
            mesInicial.Item(13).Visible = True
            mesFin.Item(13).Visible = True
        End If
        If (nombreArchivo = "JUL220") Then
            mesInicial.Item(14).Caption = "JUL2daQuin"
            mesFin.Item(14).Caption = "JUL2daQuin"
            mesInicial.Item(14).Visible = True
            mesFin.Item(14).Visible = True
        End If
        If (nombreArchivo = "AGO120") Then
            mesInicial.Item(15).Caption = "AGO1raQuin"
            mesFin.Item(15).Caption = "AGO1raQuin"
            mesInicial.Item(15).Visible = True
            mesFin.Item(15).Visible = True
        End If
        If (nombreArchivo = "AGO220") Then
            mesInicial.Item(16).Caption = "AGO2daQuin"
            mesFin.Item(16).Caption = "AGO2daQuin"
            mesInicial.Item(16).Visible = True
            mesFin.Item(16).Visible = True
        End If
        If (nombreArchivo = "SEP120") Then
            mesInicial.Item(17).Caption = "SEP1raQuin"
            mesFin.Item(17).Caption = "SEP1raQuin"
            mesInicial.Item(17).Visible = True
            mesFin.Item(17).Visible = True
        End If
        If (nombreArchivo = "SEP220") Then
            mesInicial.Item(18).Caption = "SEP2daQuin"
            mesFin.Item(18).Caption = "SEP2daQuin"
            mesInicial.Item(18).Visible = True
            mesFin.Item(18).Visible = True
        End If
        If (nombreArchivo = "OCT120") Then
            mesInicial.Item(19).Caption = "OCT1raQuin"
            mesFin.Item(19).Caption = "OCT1raQuin"
            mesInicial.Item(19).Visible = True
            mesFin.Item(19).Visible = True
        End If
        If (nombreArchivo = "OCT220") Then
            mesInicial.Item(20).Caption = "OCT2daQuin"
            mesFin.Item(20).Caption = "OCT2daQuin"
            mesInicial.Item(20).Visible = True
            mesFin.Item(20).Visible = True
        End If
        
        If (nombreArchivo = "NOV120") Then
            mesInicial.Item(21).Caption = "NOV1raQuin"
            mesFin.Item(21).Caption = "NOV1raQuin"
            mesInicial.Item(21).Visible = True
            mesFin.Item(21).Visible = True
        End If
        If (nombreArchivo = "NOV220") Then
            mesInicial.Item(22).Caption = "NOV2daQuin"
            mesFin.Item(22).Caption = "NOV2daQuin"
            mesInicial.Item(22).Visible = True
            mesFin.Item(22).Visible = True
        End If
        If (nombreArchivo = "DIC120") Then
            mesInicial.Item(23).Caption = "DIC1raQuin"
            mesFin.Item(23).Caption = "DIC1raQuin"
            mesInicial.Item(23).Visible = True
            mesFin.Item(23).Visible = True
        End If
        If (nombreArchivo = "DIC220") Then
            mesInicial.Item(24).Caption = "DIC2daQuin"
            mesFin.Item(24).Caption = "DIC2daQuin"
            mesInicial.Item(24).Visible = True
            mesFin.Item(24).Visible = True
        End If

    Next
    Close 1
 
    
    mm(1) = "ENERO": mm(2) = "FEBRERO": mm(3) = "MARZO": mm(4) = "ABRIL"
    mm(5) = "MAYO": mm(6) = "JUNIO": mm(7) = "JULIO": mm(8) = "AGOSTO"
    mm(9) = "SEPTIEMBRE": mm(10) = "OCTUBRE": mm(11) = "NOVIEMBRE"
    mm(12) = "DICIEMBRE"
    dd(1) = 31: dd(2) = 28: dd(3) = 31: dd(4) = 30
    dd(5) = 31: dd(6) = 30: dd(7) = 31: dd(8) = 31
    dd(9) = 30: dd(10) = 31: dd(11) = 30: dd(12) = 31
    inicio
    abre
    If Femp < 1 Then
            Close
            Exit Sub
    End If
    Acumula_todo
    encabezado
    elecero
    busquedaInicial = ""
    busquedaFinal = ""
End Sub

Private Sub Form_Resize()
    Dacu1.Width = ScaleWidth
    Dacu1.Height = ScaleHeight
End Sub
Sub ArchivoInformativo()
 Dim DatoTe, A_o, Empleado As Integer, DA_TO As String
 Dim CALCULOANUAL As String, SubsidioAplicable As Currency
 Close 5
 Open "AcuNom2.dno" For Random As 5 Len = Len(ArAcum)
 facum = LOF(5) / Len(ArAcum)
 Close 3
 Close 4
 Open "PerOtre.dno" For Random As 4 Len = Len(Otros_Rgtros)
 Fin_Otreg = LOF(4) / Len(Otros_Rgtros)
 Open "Personal.dno" For Random As 3 Len = Len(personal)
 FPer = LOF(3) / Len(personal)
 Close 16
 Rem ******** Otros_Rgtros
 Open "AcuNom1.dno" For Random As 16 Len = Len(Ot_Acum)
 FAcum1 = LOF(16) / Len(Ot_Acum)
 If FAcum1 = 0 Then Close 16: Kill "AcuNom1.dno"
 ArchivoTexto = "Inf" + Trim(empresa.ao) + ".txt"
 Open ArchivoTexto For Output As 12
 For r = 1 To Dacu1.Rows - 2
    If Dacu1.TextMatrix(r, 0) > "" Then
       Empleado = Dacu1.TextMatrix(r, 0)
       Get 3, Empleado, personal
       Get 5, Empleado, ArAcum
       Get 4, Empleado, Otros_Rgtros
       
       If FAcum1 > 0 Then
            Get 16, Empleado, Ot_Acum
            Ot_Acum.Pdias = Ot_Acum.Pdias + ArAcum.Pdias
            AcumSup_Cor = 0
            AcumSup_Cor = CLng(Ot_Acum.Pnormal + Ot_Acum.Pagui + Ot_Acum.Pextra _
                 + Ot_Acum.Potras + Ot_Acum.PPTU + Ot_Acum.Pvaca + Ot_Acum.Pviaticos) - AcumSup_CorEx
      
       End If
     Rem Campo 1 ********************************************
        A_o = Mid(personal.fal, 7, 4)
     If Val(A_o) = empresa.ao Then
        mes = Mid(personal.fal, 4, 2)
        Else
        mes = "01"
     End If
       DatoTe = mes + "|"
     Rem  campo 2 ********************************************
     Rem Se hace calculo cuando trabajo todo el año **********
     Rem Se hace calculo cuando trabajo en otro lado y lo solicita *******
     
     If (Val(LTrim(Mid(personal.fab, 7, 4))) <= 0) Then
        DatoTe = DatoTe + "12|"
        CALCULOANUAL = "1"
        If mes <> "01" Then CALCULOANUAL = "2"
        Else
        
        If Val(Mid(LTrim(personal.fab), 7, 4)) = empresa.ao Then
             DatoTe = DatoTe + Mid(personal.fab, 4, 2) + "|"
             CALCULOANUAL = "2"
             Else
             DatoTe = DatoTe + "01|"
             CALCULOANUAL = "2"
        End If
        If Val(Mid(LTrim(personal.fal), 7, 4)) = empresa.ao Then
             Rem DatoTe = DatoTe + Mid(personal.fal, 4, 2) + "|"
             CALCULOANUAL = "2"
        End If
     End If
      
    End If
    If Ot_Acum.Pdias > 343 Then CALCULOANUAL = "1"
    Rem campo 3 ************************************************
    Rem RFC ****************************************************
     DA_TO = RTrim(personal.rfc)
     If Len(DA_TO) < 10 Then
          DA_TO = Mid(personal.ape1, 1, 2) + Mid(personal.ape2, 1, 1) + Mid(personal.nom, 1, 1) + "000000"
     End If
     EliminaGuion DA_TO, Empleado
     If Len(DA_TO) < 13 Then DA_TO = DA_TO + String(13 - Len(DA_TO), "0")
     DatoTe = DatoTe + DA_TO + "|"
     Rem campo 4 ***********************************************
     Rem CURP **************************************************
     DA_TO = RTrim("")
     
     DA_TO = RTrim(Otros_Rgtros.curp)
     EliminaGuion DA_TO, Empleado
     Rem If Len(DA_TO) < 18 Then DA_TO = DA_TO + String(18 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     
     Rem CAMPO 5 APELLIDO PATERNO ******************************
     DA_TO = RTrim(personal.ape1)
     Rem If Len(DA_TO) < 43 Then DA_TO = DA_TO + String(43 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     
     Rem CAMPO 6 APELLIDO MATERNO ******************************
     DA_TO = RTrim(personal.ape2)
     Rem If Len(DA_TO) < 43 Then DA_TO = DA_TO + String(43 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 7 Nombre ******************************
     DA_TO = RTrim(personal.nom)
     Rem If Len(DA_TO) < 43 Then DA_TO = DA_TO + String(43 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     
     Rem CAMPO 8 Area del salario minimo **************
     DA_TO = RTrim("01")
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 9  El patron realizo calculo anual **************
     DA_TO = RTrim(CALCULOANUAL)
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     
     Rem CAMPO 10 TARIFA UTILIZADA *****************************
     DA_TO = RTrim("1")
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 11 TARIFA 1991 UTILIZADA *****************************
     DA_TO = RTrim("2")
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 12 PROPORCION DEL SUBSIDIO APLICADA
     
     Rem Subidio_doble = (ArAcum.DSubioAp + Ot_Acum.DSubioAp) / (ArAcum.DSubioAp + Ot_Acum.DSubioAp + ArAcum.DSubNoap + Ot_Acum.DSubNoap)
     
     SubsidioAplicable = (1 - ((1 - empresa.psub) / 2))
     If AcumSup_Cor > 0 Then
         sub1001 = (ArAcum.DSubioAp + Ot_Acum.DSubioAp)
         sub1002 = (ArAcum.DSubioAp + Ot_Acum.DSubioAp + ArAcum.DSubNoap + Ot_Acum.DSubNoap)
         If sub1002 > 0 Then
          subsidio12 = (ArAcum.DSubioAp + Ot_Acum.DSubioAp) / (ArAcum.DSubioAp + Ot_Acum.DSubioAp + ArAcum.DSubNoap + Ot_Acum.DSubNoap)
          Else
          subsidio12 = 0
         End If
          SubsidioAplicable = (1 - ((1 - subsidio12) / 2))
     End If
     SubsidioAplicable = Format(SubsidioAplicable, "#.0000")
     subdioforma = Format(SubsidioAplicable, "#.0000")
     If subdioforma < 1 Then
            DA_TO = SubsidioAplicable
            If Len(DA_TO) < 6 Then DA_TO = DA_TO + String(6 - Len(DA_TO), "0")
            Else
            DA_TO = subdioforma
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 13 EL TRABAJADOR ES SINDICALIZADO ********************
     DA_TO = RTrim("2")
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 14 SI ES ASIMILADO O NO *****************************
     DA_TO = RTrim("0")
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 15 CLAVE DE LA ENTIDAD FEDERATIVA ********************
     DA_TO = RTrim("09")
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPOS 16 AL 25 OTROS patrones  ****************************
     If FAcum1 > 0 Then
        DA_TO = "DEO010228PZ1"
        Rem DA_TO = "CTA840227RT0"
        DatoTe = DatoTe + DA_TO + "|"
        INICIAR = 17
        Else
        INICIAR = 16
     End If
     
     For i = INICIAR To 25
        DA_TO = ""
        Rem If Len(DA_TO) < 13 Then DA_TO = DA_TO + String(13 - Len(DA_TO), " ")
        DatoTe = DatoTe + DA_TO + "|"
     Next i
     Rem CAMPO 26 Pagos aportaciones voluntarias *******************************
     Rem DA_TO = RTrim("2")
     DA_TO = ""
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 27 Monto aportaciones voluntarias aplicadas al calculo del impuesto *******************************
     DA_TO = RTrim("2")
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 28 Aportaciones voluntarias deducibles *******************************
     DA_TO = ""
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 29 aportaciones voluntarias deducibles *****************
     DA_TO = ""
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem campos 30 Ingresos por separacion *************************
     Rem For I = 29 To 46
        Rem DA_TO = ""
        DA_TO = RTrim("2")
        DatoTe = DatoTe + DA_TO + "|"
     Rem Next I
     Rem campos 31 Asimilados a salarios *************************
        Rem DA_TO = ""
        DA_TO = RTrim("2")
        DatoTe = DatoTe + DA_TO + "|"
     Rem campos 32 PAGOS DEL PATRON EFECTUADOS A SUS TRABAJADORES ******************
        Rem DA_TO = ""
        DA_TO = RTrim("1")
        DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPOS 33 AL 56  ASIMILADOS A SALARIOS **********************************
     Rem DA_TO = ""
     Rem DatoTe = DatoTe + DA_TO + "|"
     Rem campo 48 IMPUESTO RETENDIO A ASIMILADOS A SALARIOS *************
     Rem DA_TO = ""
     Rem DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 58 Sueldos, salarios gravados **** SE CONVIERTE EN 54 AL NO TOMAR LA OPCION DE ASIMILADOS
     DA_TO = LTrim(Str(CLng(ArAcum.Pnormal)))
     If ArAcum.Pnormal = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     
     Rem CAMPO 59 Sueldos, salarios EXENTOS SE CONVIERTE EN 55 AL NO TOMAR LA OPCION DE ASIMILADOS
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 60 Gratificacion Anual Gravado SE CONVIERTE EN 56
     If ArAcum.Pagui > (empresa.sm * 30) Then
           AguinaldoGravado = ArAcum.Pagui - (empresa.sm * 30)
           AguinaldoExento = (empresa.sm * 30)
           DA_TO = LTrim(Str(CLng(AguinaldoGravado)))
           If AguinaldoGravado = 0 Then DA_TO = ""
           DatoTe = DatoTe + DA_TO + "|"
           Else
           AguinaldoGravado = 0
           AguinaldoExento = ArAcum.Pagui
           DA_TO = LTrim(Str(CLng(AguinaldoGravado)))
           If AguinaldoGravado = 0 Then DA_TO = ""
           DatoTe = DatoTe + DA_TO + "|"
     End If
     Rem campo 61 Gratificacion anual Exento SE CONVIERTE EN 57
     DA_TO = LTrim(Str(CLng(AguinaldoExento)))
     If AguinaldoExento = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 62 Viaticos y gastos de viaje gravado SE CONVIERTE EN 58
     DA_TO = LTrim(Str(CLng(ArAcum.Pviaticos)))
     If ArAcum.Pviaticos = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 63 Viaticos y Gastos de Viaje Exento SE CONVIERTE EN 59
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 64 Tiempo extraordinario Gravado SE CONVIERTE EN 60
     Rem DA_TO = LTrim(Str(CLng(ArAcum.Pextra)))
     Rem If ArAcum.Pextra = 0 Then DA_TO = ""
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 65 Tiempo extraordinario Exento SE CONVIERTE EN 61
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 66 Prima Vacacional Gravada SE CONVIERTE EN 62
     If ArAcum.Pvaca > (empresa.sm * 15) Then
           PrimaVacacionalGravada = ArAcum.Pvaca - (empresa.sm * 15)
           PrimaVacacioNalExenta = (empresa.sm * 15)
           DA_TO = LTrim(Str(CLng(PrimaVacacionalGravada)))
           If PrimaVacacionalGravada = 0 Then DA_TO = ""
           DatoTe = DatoTe + DA_TO + "|"
           Else
           PrimaVacacionalGravada = 0
           PrimaVacacioNalExenta = ArAcum.Pvaca
           DA_TO = LTrim(Str(CLng(PrimaVacacionalGravada)))
           If PrimaVacacionalGravada = 0 Then DA_TO = ""
           DatoTe = DatoTe + DA_TO + "|"
     End If
     Rem campo 67 Prima Vacacional Exenta SE CONVIERTE EN 63
     DA_TO = LTrim(Str(CLng(PrimaVacacioNalExenta)))
     If PrimaVacacioNalExenta = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 68 Prima dominical gravada SE CONVIERTE EN 64
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 69 Prima dominical exenta SE CONVIERTE EN 65
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 70 PTU Gravada SE CONVIERTE EN 66
     If ArAcum.PPTU > (empresa.sm * 15) Then
           PTUGravada = ArAcum.PPTU - (empresa.sm * 15)
           PTUExenta = (empresa.sm * 15)
           DA_TO = LTrim(Str(CLng(PTUGravada)))
           If PTUGravada = 0 Then DA_TO = ""
           DatoTe = DatoTe + DA_TO + "|"
           Else
           PTUGravada = 0
           PTUExenta = ArAcum.PPTU
           DA_TO = LTrim(Str(CLng(PTUGravada)))
           If PTUGravada = 0 Then DA_TO = ""
           DatoTe = DatoTe + DA_TO + "|"
     End If
     Rem campo 71 PTU Exenta SE CONVIERTE EN 67
     DA_TO = LTrim(Str(CLng(PTUExenta)))
     If PTUExenta = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 72 Reembolso de gtos. medicos Gravados SE CONVIERTE EN 68
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 73 Reembolso de gtos. medicos exentos SE CONVIERTE EN 69
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 74 Fondo de ahorro Gravados SE CONVIERTE EN 70
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 75 Fondo de ahorro Exento SE CONVIERTE EN 71
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 76 Caja de ahorro Gravados SE CONVIERTE EN 72
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 77 Caja de ahorro Exento SE CONVIERTE EN 73
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 78 Vales de despensa Gravado SE CONVIERTE EN 74
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 79 Vales de despensa Exento SE CONVIERTE EN 75
     DA_TO = LTrim(Str(CLng(ArAcum.Pexenta)))
     If ArAcum.Pexenta = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 80 Ayuda para gtos de funeral Gravado SE CONVIERTE EN 76
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 81 Ayuda para gtos de funeral exento SE CONVIERTE EN 77
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 82 Contribuciones a cargo del trabajador pag.x el patron Gravado SE CONVIERTE EN 78
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 83 Contribuciones a cargo del trabajador pag.x el patron Exento SE CONVIERTE EN 79
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 84 Premios por puntualidad Gravado SE CONVIERTE EN 80
     DA_TO = LTrim(Str(CLng(ArAcum.Pextra)))
     If ArAcum.Pextra = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 85 Premios por puntualidad Exento SE CONVIERTE EN 81
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 86 to 109 SE CONVIERTE EN 82 AL 105
     For i = 86 To 109
        DA_TO = "0"
        DatoTe = DatoTe + DA_TO + "|"
     Next i
     Rem CAMPO 110 PAGOS EFECTUADOS X OTROS EMPLEADORES GRAVADO SE CONVIERTE EN 106
     If FAcum1 > 0 Then
      'If Ot_Acum.PPTU < (empresa.sm * 15) Then
            'AcumSup_CorEx = 0
            'AcumSup_CorEx = CLng(Ot_Acum.PPTU)
            'Else
            'AcumSup_CorEx = 0
            'AcumSup_CorEx = CLng(empresa.sm * 15)
      'End If
      
            AcumSup_CorEx = CLng(Ot_Acum.PPTU)
            If AcumSup_Cor > 0 Then
                 DA_TO = LTrim(Str(AcumSup_Cor))
                 Else
                 DA_TO = "0"
            End If
            Else
            AcumSup_Cor = 0
            DA_TO = "0"
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 111 PAGOS EFECTUADOS X OTROS EMPLEADORES EXENTO SE CONVIERTE EN 107
     If (FAcum1 > 0) And (AcumSup_CorEx > 0) Then
           
           DA_TO = LTrim(Str(AcumSup_CorEx))
           Else
           DA_TO = "0"
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 112 OTROS INGRESOS POR SALARIOS GRAVADOS SE CONVIERTE EN 108
     DA_TO = LTrim(Str(CLng(ArAcum.Potras)))
     If ArAcum.Potras = 0 Then DA_TO = "0"
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 113 OTROS INGRESOS POR SALARIOS EXENTOS SE CONVIERTE EN 109
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 114 SUMA INGRESO GRAVADO SE CONVIERTE EN 110
     SumaGravado = ArAcum.Pnormal + ArAcum.Pextra + ArAcum.Pviaticos + ArAcum.Potras _
                   + AguinaldoGravado + PrimaVacacionalGravada + PTUGravada
     DA_TO = LTrim(Str(CLng(SumaGravado)))
     If SumaGravado = 0 Then DA_TO = "0"
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 115 SUMA INGRESO EXENTO SE CONVIERTE EN 111
     SumaExento = AguinaldoExento + PrimaVacacioNalExenta + PTUExenta
     DA_TO = LTrim(Str(CLng(SumaExento)))
     If SumaExento = 0 Then DA_TO = "0"
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 116 Impuesto retenido SE CONVIERTE EN 112
     If ArAcum.DImpret < 0 Then ArAcum.DImpret = 0
     TOT_ret = 0: TOT_ret = ArAcum.DImpret
     DA_TO = LTrim(Str(CLng(TOT_ret)))
     If TOT_ret = 0 Then DA_TO = "0"
     DatoTe = DatoTe + DA_TO + "|"
     
     Rem CAMPO 117 Impuesto retenido por otros patrones SE CONVIERTE EN 113
     If (FAcum1 > 0) Then
           DA_TO = LTrim(Str(CLng(Ot_Acum.DImpret)))
           Else
           DA_TO = "0"
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 118 SALDO A FAVOR X COMPENSAR SE CONVIERTE EN 114
     DA_TO = "0"
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 119 SALDO A FAVOR DEL EJERCICIO ANTERIOR NO COMPENSADO DURANTE... SE CONVIERTE EN 115
     DA_TO = "0"
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 120 SUMA CREDITO AL SALARIO CALCULADO SE CONVIERTE EN 1116
     If ArAcum.DCrApl < 0 Then ArAcum.DCrApl = 0
     DA_TO = LTrim(Str(CLng(ArAcum.DCrApl)))
     If ArAcum.DCrApl = 0 Then DA_TO = "0"
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 121 SUMA CREDITO AL SALARIO pagado SE CONVIERTE EN 117
     Rem DA_TO = LTrim(Str(CLng(ArAcum.DCrPag * -1)))
     If ArAcum.DCrPag < 0 Then ArAcum.DCrPag = ArAcum.DCrPag * -1
     DA_TO = LTrim(Str(CLng(ArAcum.DCrPag)))
     If Int(ArAcum.DCrPag) = 0 Then DA_TO = "0"
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 122 Ingresos de prestaciones de prevision social gravada  SE CONVIERTE EN 118
     If ArAcum.Pexenta > 0 Then
        DA_TO = LTrim(Str(CLng(ArAcum.Pexenta)))
        Else
        DA_TO = "0"
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 123 Ingresos de prestaciones de prevision social exentas SE CONVIERTE EN 119
     If ArAcum.Pexenta > 0 Then
        DA_TO = LTrim(Str(CLng(ArAcum.Pexenta)))
        Else
        DA_TO = "0"
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 124 SUMA INGRESOS POR SUELDOS Y SALARIOS ***** GRAVADO SE CONVIERTE EN 120
     Rem     ************************************************
     Rem    **************************************************
     Rem   ****************************************************
     Rem  ******************************************************
     Rem  ******************************************************
     Rem   ****************************************************
     Rem    **************************************************
     Rem     ************************************************
     If AcumSup_Cor > 0 Then
         SumaGravado = SumaGravado + AcumSup_Cor
         DA_TO = LTrim(Str(CLng(SumaGravado)))
         Else
         DA_TO = ""
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 116 SUMA INGRESO Exento
     'If AcumSup_CorEx > 0 Then
         'SumaExento = SumaExento + AcumSup_CorEx
         'DA_TO = LTrim(Str(CLng(SumaExento)))
         'Else
         'DA_TO = ""
     'End If
     'DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 125 AUMENTADO EN 2005 MONTO DEL IMPTO.LOCAL A LOS ING.POR SDOS.Y SAL.
        DA_TO = "0"
        DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 126 AUMENTADO EN 2005 MONTO DEL SUBSDIDIO AL EMPLEO ENTREGADO EN EFECTIVO AL TRAB.
        DA_TO = LTrim(Str(CLng(ArAcum.DCrPag)))
        Rem DA_TO = "0"
        DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 127 TOTAL DE LAS APORTACIONES VOLUNTARIAS.
        DA_TO = "0"
        DatoTe = DatoTe + DA_TO + "|"
     Rem Campo &&& AUMENTADO EN 2005 IMPTO.LOCAL A LOS ING.POR SDOS.Y SAL.
        Rem DA_TO = "0"
        Rem DatoTe = DatoTe + DA_TO + "|"
        
     Rem CAMPO 128 ISR CONFORME TARIFA ANUAL
     If CALCULOANUAL = "1" Then
            impto = 0
            If AcumSup_Cor > 0 Then
                Subidio_doble1 = (ArAcum.DSubioAp + Ot_Acum.DSubioAp)
                Subidio_doble2 = (ArAcum.DSubioAp + Ot_Acum.DSubioAp + ArAcum.DSubNoap + Ot_Acum.DSubNoap)
                If Subidio_doble2 > 0 Then
                        Subidio_doble = Subidio_doble1 / Subidio_doble2
                        Else
                        Subidio_doble = empresa.psub
                        
                End If
                calc_anual CLng(SumaGravado), impto, Subidio_doble
                Else
                calc_anual CLng(SumaGravado), impto, empresa.psub
            End If
            
            DA_TO = LTrim(Str(CLng(imptotal - subdio)))
            Else
            DA_TO = "0"
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 129 SUBSIDIO ACREDITABLE 116 SE CONVIERTE EN 126
     
     DA_TO = LTrim(Str(CLng(subdio)))
     If subdio = 0 Then DA_TO = "0"
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 130 SUBSIDIO NO ACREDITABLE 117 SE CONVIERTE EN 127
     DA_TO = LTrim(Str(CLng(SUBDIONOACREDITABLE)))
     If SUBDIONOACREDITABLE = 0 Then DA_TO = "0"
        DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 120 Subsidio acreditable fraccion III en 2001
     'DA_TO = ""
     'DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 121 Subsidio acreditable fraccion IV 2001
     'DA_TO = ""
     'DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 131 SUMA impuesto sobre ingresos acumulables 118 SE CONVIERTE EN 128
     
     If CALCULOANUAL = "1" Then
            DA_TO = LTrim(Str(CLng(imptotal - subdio)))
            If (imptotal - JCre.Aplicado) < 0 Then DA_TO = "0"
            Else
            DA_TO = "0"
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 132 Impuesto sobre ingresos no acumulables ******* 119 SE CONVIERTE EN 129
     DA_TO = "0"
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 133 Impuesto local a los ingresos por sueldos y salarios ******* 119 SE CONVIERTE EN 129
     DA_TO = "0"
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 134 MONTO DEL SUBSIDIO PARA EL EMPLEO QUE LE CORRESPONDIO DURANTE EL EJERCICIO ******* 119 SE CONVIERTE EN 129
     DA_TO = LTrim(Str(CLng(ArAcum.DSubioAp)))
     Rem DatoTe = DatoTe + DA_TO + "|"
     DatoTe = DatoTe + DA_TO
     Rem DatoTe = DatoTe + DA_TO + "|||||||||"
     Rem CAMPO 124 Total de sueldos, salarios y conceptos asimilados
     'DA_TO = LTrim(Str(CLng(SumaGravado + SumaExento)))
     'If CLng(SumaGravado + SumaExento) = 0 Then DA_TO = ""
     'DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 125 Ingresos exentos
     'DA_TO = LTrim(Str(CLng(SumaExento)))
     'If CLng(SumaExento) = 0 Then DA_TO = ""
     'DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 126 Ingresos no acumulables
     'DA_TO = ""
     'DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 127 Ingresos acumulables
     'DA_TO = LTrim(Str(CLng(SumaGravado)))
     'If CLng(SumaGravado) = 0 Then DA_TO = ""
     'DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 128 Impuesto sobre la renta causado en ejercicio que declara
     
     'If CALCULOANUAL = "1" Then
            'If AcumSup_Cor > 0 Then
               'DA_TO = LTrim(Str(CLng(impto - Ot_Acum.DCrApl)))
               'If (CLng(impto - Ot_Acum.DCrApl)) < 0 Then DA_TO = ""
               'Else
               'DA_TO = LTrim(Str(CLng(impto)))
            'End If
            'Else
            'DA_TO = ""
           
     'End If
     'DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 129 Impuesto retenido en el ejercicio que declara
     'If AcumSup_Cor > 0 Then
            'TOT_retenido = ArAcum.DImpret + Ot_Acum.DImpret
            'DA_TO = LTrim(Str(CLng(TOT_retenido)))
            'Else
            'TOT_retenido = ArAcum.DImpret
            'DA_TO = LTrim(Str(CLng(TOT_retenido)))
     'End If
     'If TOT_retenido = 0 Then DA_TO = ""
     'DatoTe = DatoTe + DA_TO + "|"
     Print #12, DatoTe
     impto = 0: imptotal = 0: subdio = 0: SUBDIONOACREDITABLE = 0
     AcumSup_Cor = 0: AcumSup_CorEx = 0
    Next r
    Close 3, 5, 12
    mensaje = "Archivo Generado Ultimo empleado: " + Chr(13) + RTrim(personal.nom) + _
              " " + RTrim(personal.ape1) + " "
    MsgBox mensaje
End Sub
Sub EliminaGuion(DA_TO, Empleado)
   RF_C = ""
    For i = 1 To Len(DA_TO)
        If (Mid(DA_TO, i, 1) = "-") Or (Mid(DA_TO, i, 1) = " ") Then
           Rem nada
           Else
           RF_C = RF_C + Mid(DA_TO, i, 1)
        End If
    Next i
    
    DA_TO = RF_C
    
End Sub

Private Sub InfoGene_Click()
    ArchivoInformativo
End Sub

Private Sub InfoTring_Click()
   OtrosIngr.Show
End Sub

Private Sub mesFin_Click(Index As Integer)
Dim acIngresos
Dim nombre
Dim acSueldo As Currency
acSueldo = 0
busquedaFinal = Index
AxNom1.tipo = True


If (busquedaInicial <> "" And busquedaFinal >= busquedaInicial) Then
    Dacu1.Rows = 1
    'For i = 1 To 1
        'For g = 0 To Dacu1.Cols - 1
        'Dacu1.TextMatrix(i, g) = ""
        'Next
    'Next
   'MsgBox ("Listo")
   If Dir("Auxiliar\AxNa*.*", vbArchive) <> "" Then
    Close 7
    Kill ("Auxiliar\AxNa*.*")
   End If
   
   
   On Error GoTo manejo1
   lectura
   Open "AcuTemp" For Random As 10 Len = Len(temporal)
   Ftem = LOF(10) / Len(temporal)
   Close 5
   Close 7
   Open "personal.dno" For Random As 13 Len = Len(personal)
   Open "AcuNom2.dno" For Random As 5 Len = Len(ArAcum)
   facum = LOF(5) / Len(ArAcum)
   
    For r = 1 To FPer
       ceros
       Put 5, r, ArAcum
   Next r
    
   For r = 1 To Ftem: Get 10, r, temporal
    Rem On Error GoTo manejo1
    extras = Left(Trim(temporal.miarchivo), 4)
        For l = busquedaInicial To busquedaFinal
            nomMesL = Left(Trim(mesInicial.Item(l).Caption), 4)
            If extras = nomMesL Then
                temporal.miarchivo = Trim(temporal.miarchivo)
                Open temporal.miarchivo For Random As 4 Len = Len(nomina)
                fnom = LOF(4) / Len(nomina)
                Arch1 = Trim(temporal.miarchivo)
                Arch1 = Mid(Arch1, 1, Len(Arch1) - 4) + ".cmp"
                Open Arch1 For Random As 6 Len = Len(Nom_Com)
                'MsgBox (temporal.miarchivo + "  " + Arch1)
                FNomcom = LOF(6) / Len(Nom_Com)
                Debug.Print Arch1
                reconocedornomina
                If EXTRA = 3 Then
                    Rem GoTo manejo1
                    Close 4, 6
                Else
                    For i = 1 To fnom
                        Get 4, i, nomina
                        Get 6, i, Nom_Com
                        Get 13, i, personal
                        InGresos = nomina.sueldo + nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri + nomina.aguin + nomina.viaticos + nomina.pvac + nomina.otras + nomina.exentos + nomina.ptu
                        If InGresos <> 0 Then
                    
                            If EXTRA = 1 Then
                                Rem If (Val(Mid(LTrim(personal.fab), 7, 4)) > 0) And (Val(Mid(LTrim(personal.fab), 7, 4)) <= empresa.ao) Then
                                Nom_Com.CredNe = 0: Nom_Com.CreTot = 0
                                Nom_Com.ImpTot = 0: Nom_Com.subapl = 0
                                Nom_Com.subdio = 0: Nom_Com.subNap = 0
                                Nom_Com.ImpTot = nomina.ispt
                                Rem Put 6, I, Nom_Com
                            Rem End If
                            Rem If (Val(Mid(LTrim(personal.fal), 7, 4)) > 0) And (Val(Mid(LTrim(personal.fal), 7, 4)) = empresa.ao) Then
                                Rem Nom_Com.CredNe = 0: Nom_Com.CreTot = 0
                                Rem Nom_Com.ImpTot = 0: Nom_Com.subapl = 0
                                Rem Nom_Com.subdio = 0: Nom_Com.subNap = 0
                                Rem Nom_Com.ImpTot = nomina.ispt
                                Rem Put 6, I, Nom_Com
                            Rem End If
                            Rem If UCase(Left(LTrim(Arch1), 3)) = "PTU" Then
                                Rem Nom_Com.CredNe = 0: Nom_Com.CreTot = 0
                                Rem Nom_Com.ImpTot = 0: Nom_Com.subapl = 0
                                Rem Nom_Com.subdio = 0: Nom_Com.subNap = 0
                                Rem Nom_Com.ImpTot = nomina.ispt
                                Rem Put 6, I, Nom_Com
                            Rem End If
                            End If
                            Get 5, i, ArAcum
                            NoAux = "Auxiliar\AxNa" + LTrim(Str(i))
                            Open NoAux For Random As 7 Len = Len(AxNom)
                            FiNax = LOF(7) / Len(AxNom)
                            FiNax = FiNax + 1
                            ArAcum.Pdias = ArAcum.Pdias + nomina.dias
                            ArAcum.Pnormal = ArAcum.Pnormal + nomina.sueldo
                            If RTrim(temporal.miarchivo) = UCase("PREM2005.NOM") Then
                                nomina.hs_tri = nomina.otras: nomina.otras = 0
                            End If
                            ArAcum.Pextra = ArAcum.Pextra + (nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri)
                            ArAcum.Pviaticos = ArAcum.Pviaticos + nomina.viaticos
                            ArAcum.Pagui = ArAcum.Pagui + nomina.aguin
                            ArAcum.Pvaca = ArAcum.Pvaca + nomina.pvac
                            ArAcum.Potras = ArAcum.Potras + nomina.otras
                            ArAcum.PPTU = ArAcum.PPTU + nomina.ptu
                            ArAcum.Pexenta = ArAcum.Pexenta + nomina.exentos
                            ArAcum.DImpto = ArAcum.DImpto + Nom_Com.ImpTot
                            ArAcum.DSubioAp = ArAcum.DSubioAp + Nom_Com.subapl
                            ArAcum.DCrApl = ArAcum.DCrApl + Nom_Com.CreTot
                            ArAcum.DImpret = ArAcum.DImpret + nomina.ispt
                            ArAcum.DCrPag = ArAcum.DCrPag + nomina.crdsal
                            ArAcum.DSubNoap = ArAcum.DSubNoap + Nom_Com.subNap
                            ArAcum.DImss = ArAcum.DImss + nomina.imss
                            ArAcum.DPrestamos = ArAcum.DPrestamos + nomina.prestamos
                            ArAcum.DTonacot = ArAcum.DTonacot + nomina.fonacot
                            ArAcum.DTelefono = ArAcum.DTelefono + nomina.telefono
                            ArAcum.DOtrasded = ArAcum.DOtrasded + nomina.otraded
                            Put 5, i, ArAcum
                            Auxliar
                            Put 7, FiNax, AxNom
                    
                            Close 7
                            AxPer.Pdias = nomina.dias
                            AxPer.Pnormal = nomina.sueldo
                            AxPer.Pextra = (nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri)
                            AxPer.Pviaticos = nomina.viaticos
                            AxPer.Pvaca = nomina.pvac
                            AxPer.Pagui = nomina.aguin
                            AxPer.Potras = nomina.otras
                            AxPer.PPTU = nomina.ptu
                            AxPer.Pexenta = AxPer.Pexenta + nomina.exentos
                            AxPer.DImpto = AxPer.DImpto + Nom_Com.ImpTot
                            AxPer.DSubioAp = AxPer.DSubioAp + Nom_Com.subapl
                            AxPer.DCrApl = AxPer.DCrApl + Nom_Com.CreTot
                            AxPer.DImpret = AxPer.DImpret + nomina.ispt
                            AxPer.DCrPag = AxPer.DCrPag + nomina.crdsal
                            AxPer.DSubNoap = AxPer.DSubNoap + Nom_Com.subNap
                            AxPer.DImss = AxPer.DImss + nomina.imss
                            AxPer.DPrestamos = AxPer.DPrestamos + nomina.prestamos
                            AxPer.DTonacot = AxPer.DTonacot + nomina.fonacot
                            AxPer.DTelefono = AxPer.DTelefono + nomina.telefono
                            AxPer.DOtrasded = AxPer.DOtrasded + nomina.otraded
                    
                        End If
                    Next i
                    Close 4, 6
                End If
            End If
        Next
    Next r
    facum = LOF(5) / Len(ArAcum)

    For r = 1 To facum: Get 5, r, ArAcum
            Get 13, r, personal
            InGresos = ArAcum.Pnormal + ArAcum.Pextra + ArAcum.Pviaticos + ArAcum.Pvaca + ArAcum.Potras + ArAcum.Pagui + ArAcum.PPTU + ArAcum.Pexenta
            If InGresos <> 0 Then
            nombre = RTrim(personal.ape1) + " " + RTrim(personal.ape2) + " " + RTrim(personal.nom)
            entrada = Format(r, "###0") & Chr(9) & nombre & Chr(9) & Format(ArAcum.Pdias, "###,###,##0.00") & Chr(9) & Format(ArAcum.Pnormal, "###,###,##0.00") & Chr(9) & _
                      Format(ArAcum.Pextra, "###,###,##0.00") & Chr(9) & _
                      Format(ArAcum.Pviaticos, "###,###,##0.00") & Chr(9) & Format(ArAcum.Pvaca, "###,###,##0.00") & Chr(9) & Format(ArAcum.Potras, "###,###,##0.00") & Chr(9) & _
                      Format(ArAcum.Pagui, "###,###,##0.00") & Chr(9) & _
                      Format(ArAcum.PPTU, "###,###,##0.00") & Chr(9) & Format(ArAcum.Pexenta, "###,###,##0.00") & Chr(9) & _
                      Format(InGresos, "###,###,##0.00") & Chr(9) & Format(ArAcum.DImpto, "###,###,##0.00") _
                      & Chr(9) & Format(ArAcum.DSubioAp, "###,###,##0.00") & Chr(9) & Format(ArAcum.DCrApl, "###,###,##0.00") _
                      & Chr(9) & Format(ArAcum.DImpret, "###,###,##0.00") & Chr(9) & Format(ArAcum.DCrPag, "###,###,##0.00") _
                      & Chr(9) & Format(ArAcum.DSubNoap, "###,###,##0.00") & Chr(9) & Format(ArAcum.DImss, "###,###,##0.00") _
                      & Chr(9) & Format(ArAcum.DPrestamos, "###,###,##0.00") & Chr(9) & Format(ArAcum.DTelefono, "###,###,##0.00") _
                      & Chr(9) & Format(ArAcum.DTonacot, "###,###,##0.00") & Chr(9) & Format(ArAcum.DOtrasded, "###,###,##0.00") _
                      & Chr(9) & personal.rfc & Chr(9) & (" " + personal.imss) & Chr(9) & (" " + personal.fal) & Chr(9) & (" " + personal.fab)
            
            Dacu1.AddItem entrada
            
            End If
    Next r
    
    Dacu1.Rows = Dacu1.Rows + 1
    sumando
    Close 10
    Close 13
    Close 5
    
    Close: Exit Sub
    
    
manejo1:
   'MsgBox "Falla al Acumular " & Chr(13) & "revisa archivos  *.Nom", vbCritical
    MsgBox (Err.Description & ": " & Err.Description)
Close
    

Else
    MsgBox ("Error al hacer la busqueda")
End If

busquedaInicial = ""
busquedaFinal = ""

End Sub

Private Sub mesInicial_Click(Index As Integer)
    busquedaInicial = Index
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Open "personal.dno" For Random As 3 Len = Len(personal)
        Numpersona = Dacu1.TextMatrix(Dacu1.Row, 0)
        Get 3, Numpersona, personal
        Text1.Text = UCase(Text1.Text)
        Dacu1.TextMatrix(Dacu1.Row, 23) = Text1.Text
        personal.rfc = Text1.Text
        Put 3, Numpersona, personal
        Close 3
        Dacu1.SetFocus
    End If
End Sub

Private Sub VerAr_Click()
    Verifica.Show
End Sub

Private Sub versionActual_Click()
informacion.Show
End Sub

' comentario
