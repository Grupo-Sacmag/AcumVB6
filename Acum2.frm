VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Acu2 
   Caption         =   "Acumulado de sueldos"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6975
   Icon            =   "Acum2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
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
      Top             =   360
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   1
      Cols            =   25
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
   Begin VB.Menu Ver 
      Caption         =   "&Ver"
      Begin VB.Menu VerAr 
         Caption         =   "&Archivos Acumulados"
         Shortcut        =   ^V
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
End
Attribute VB_Name = "acu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hoja As Integer
Dim Femp As Long, Femco As Long, FPer As Long, sum As Currency
Dim FiNax As Long, EXTRA As Integer, Arch1
Dim FAcum1 As Long, Fin_Otreg As Long
Sub inicio()
Rem On Error GoTo saltalo
    Open "C:\Archivos de programa\sccontr.soc" For Random As 3 Len = Len(SCont)
    Get 3, 1, SCont
    If SCont.guarda <= " " Then
        ChDir "C:\"
        Else
        ChDir SCont.guarda
    End If
saltalo:
    Close 3
End Sub

Sub Apertura()
    miarchivo = dir("Auxiliar", vbDirectory)
    Rem ******   Verifica que Exista el archivo de auxiliares *****
    If miarchivo = "" Then
       respuesta = MsgBox("No existe el Subdirectorio de Auxiliares desea crearlo ", vbYesNo, "Auxiliares ")
       If respuesta = vbYes Then
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
    Dacu1.Col = 2: Dacu1.CellFontBold = True: Dacu1.ColWidth(2) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Dias"
    Dacu1.Col = 3: Dacu1.CellFontBold = True: Dacu1.ColWidth(3) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Sueldos"
    Dacu1.Col = 4: Dacu1.CellFontBold = True: Dacu1.ColWidth(4) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Prem.Punt"
    Dacu1.Col = 5: Dacu1.CellFontBold = True: Dacu1.ColWidth(5) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Viaticos"
    Dacu1.Col = 6: Dacu1.CellFontBold = True: Dacu1.ColWidth(6) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "P.Vacac."
    Dacu1.Col = 7: Dacu1.CellFontBold = True: Dacu1.ColWidth(7) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Otras"
    Dacu1.Col = 8: Dacu1.CellFontBold = True: Dacu1.ColWidth(8) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Aguinaldo"
    Dacu1.Col = 9: Dacu1.CellFontBold = True: Dacu1.ColWidth(9) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "P.T.U."
    Dacu1.Col = 10: Dacu1.CellFontBold = True: Dacu1.ColWidth(10) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Exentos"
    Dacu1.Col = 11: Dacu1.CellFontBold = True: Dacu1.ColWidth(11) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Total"
    Dacu1.Col = 12: Dacu1.CellFontBold = True: Dacu1.ColWidth(12) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Impto."
    Dacu1.Col = 13: Dacu1.CellFontBold = True: Dacu1.ColWidth(13) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Subdio.Apl."
    Dacu1.Col = 14: Dacu1.CellFontBold = True: Dacu1.ColWidth(14) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Cr.Apl."
    Dacu1.Col = 15: Dacu1.CellFontBold = True: Dacu1.ColWidth(15) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Impto.Ret."
    Dacu1.Col = 16: Dacu1.CellFontBold = True: Dacu1.ColWidth(16) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Cr.Pag."
    Dacu1.Col = 17: Dacu1.CellFontBold = True: Dacu1.ColWidth(17) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Subdio.No apl."
    Dacu1.Col = 18: Dacu1.CellFontBold = True: Dacu1.ColWidth(18) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "IMSS"
    Dacu1.Col = 19: Dacu1.CellFontBold = True: Dacu1.ColWidth(19) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Prestamos"
    Dacu1.Col = 20: Dacu1.CellFontBold = True: Dacu1.ColWidth(20) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Fonacot"
    Dacu1.Col = 21: Dacu1.CellFontBold = True: Dacu1.ColWidth(21) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Telefonos"
    Dacu1.Col = 22: Dacu1.CellFontBold = True: Dacu1.ColWidth(22) = 1300:  Dacu1.CellAlignment = 3: Dacu1.Text = "Otras"
    Dacu1.Col = 23: Dacu1.CellFontBold = True: Dacu1.ColWidth(23) = 1500:  Dacu1.CellAlignment = 3: Dacu1.Text = "Rfc"
    Dacu1.Col = 24: Dacu1.CellFontBold = True: Dacu1.ColWidth(24) = 1500:  Dacu1.CellAlignment = 3: Dacu1.Text = "Imss"
End Sub

Private Sub ArCamb_Click()
    Load Camdir
    acu2.Caption = "Acumulado de sueldos"
    Camdir.Show 1
    Dacu1.Clear
    Dacu1.Rows = 1
    
    Form_Load
End Sub
Sub lectura()
    Apertura
    miarchivo = dir("*.nom")
    r = 1
    Close 10
    Open "AcuTemp" For Random As 10 Len = Len(temporal)
    Ftem = LOF(10) / Len(temporal)
    If Ftem > 0 Then
        Close 10
        Ftem = 0
        Kill "AcuTemp"
        Open "AcuTemp" For Random As 10 Len = Len(temporal)
        Kill "Auxiliar\*.*"
    End If
    Ftem = Ftem + 1
    temporal.miarchivo = miarchivo: Put 10, Ftem, temporal
    Do Until miarchivo = ""
        miarchivo = dir
        If miarchivo <> "" Then
            Ftem = Ftem + 1
            temporal.miarchivo = miarchivo: Put 10, Ftem, temporal
        End If
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
  Rem veridir
  Open "Empresa.Dno" For Random As 1 Len = Len(empresa)
  Femp = LOF(1) / Len(empresa)
  
  If Femp < 1 Then
            MsgBox "No existen archivos de nomina" & Chr(13) & _
                   "cambie el subdirectorio"
                   Close
            Exit Sub
   End If
   Get 1, Femp, empresa
   acu2.Caption = acu2.Caption + " " + RTrim(empresa.name) + " " + RTrim(empresa.ao)
  Close 2
  Open "EmpComp.Dno" For Random As 2 Len = Len(Dat_ide)
  Femco = LOF(2) / Len(Dat_ide)
  Close 3
  Open "Personal.Dno" For Random As 3 Len = Len(personal)
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
End Sub
Sub Acumula_todo()
   lectura
   Open "AcuTemp" For Random As 10 Len = Len(temporal)
   Ftem = LOF(10) / Len(temporal)
   Close 5
   Close 7
   Open "personal.dno" For Random As 13 Len = Len(personal)
   Open "AcuNom2.Dno" For Random As 5 Len = Len(ArAcum)
   facum = LOF(5) / Len(ArAcum)
      For r = 1 To FPer
       ceros
       Put 5, r, ArAcum
   Next r
    
   For r = 1 To Ftem: Get 10, r, temporal
    On Error GoTo manejo1
       Open temporal.miarchivo For Random As 4 Len = Len(nomina)
       
       Fnom = LOF(4) / Len(nomina)
       Arch1 = RTrim(temporal.miarchivo)
       Arch1 = Mid(Arch1, 1, Len(Arch1) - 4) + ".cmp"
       Open Arch1 For Random As 6 Len = Len(Nom_Com)
       FNomcom = LOF(6) / Len(Nom_Com)
       reconocedornomina
         For i = 1 To Fnom
                Get 4, i, nomina
                Get 6, i, Nom_Com
                Get 13, i, personal
                InGresos = nomina.sueldo + nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri + nomina.aguin + nomina.viaticos + nomina.pvac + nomina.otras + nomina.exentos + nomina.ptu
                If InGresos > 0 Then
                    
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
                    If RTrim(temporal.miarchivo) = UCase("PREM2004.NOM") Then
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
          
    Next r
    facum = LOF(5) / Len(ArAcum)
    
    For r = 1 To facum: Get 5, r, ArAcum
            Get 3, r, personal
            InGresos = ArAcum.Pnormal + ArAcum.Pextra + ArAcum.Pviaticos + ArAcum.Pvaca + ArAcum.Potras + ArAcum.Pagui + ArAcum.PPTU + ArAcum.Pexenta
            If InGresos > 0 Then
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
                      & Chr(9) & personal.rfc & Chr(9) & (" " + personal.imss)
            Dacu1.AddItem entrada
            End If
    Next r
    Dacu1.Rows = Dacu1.Rows + 1
manejo1:
 Close
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
  If Dacu1.Row > 0 And Dacu1.Col > 1 Then
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
    If Dacu1.Row > 1 And Dacu1.Col > 1 Then
        Dacu1.CellBackColor = vbWhite
    End If
End Sub

Private Sub EdAct_Click()
    Dacu1.FixedRows = 1
    Dacu1.FixedCols = 2

End Sub

Private Sub EdCop_Click()
   Clipboard.Clear
   For i = Dacu1.Row To Dacu1.RowSel
      For f = Dacu1.Col To Dacu1.ColSel
            Clipboard.SetText Clipboard.GetText + Dacu1.TextMatrix(i, f) & Chr(9)
      Next f
      Clipboard.SetText Clipboard.GetText + Chr(13)
   Next i
   difer = Dacu1.RowSel - Dacu1.Row

End Sub

Private Sub EdInm_Click()

End Sub


Private Sub EDesc_Click()
    Dacu1.FixedRows = 0
    Dacu1.FixedCols = 0

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
End Sub

Private Sub Form_Resize()
    Dacu1.Width = ScaleWidth
    Dacu1.Height = ScaleHeight
End Sub
Sub ArchivoInformativo()
 Dim DatoTe, A_o, Empleado As Integer, DA_TO As String
 Dim CALCULOANUAL As String, SubsidioAplicable As Currency
 Close 5
 Open "AcuNom2.Dno" For Random As 5 Len = Len(ArAcum)
 facum = LOF(5) / Len(ArAcum)
 Close 3
 Open "PerOtre.dno" For Random As 4 Len = Len(Otros_Rgtros)
 Fin_Otreg = LOF(4) / Len(Otros_Rgtros)
 Open "Personal.Dno" For Random As 3 Len = Len(personal)
 FPer = LOF(3) / Len(personal)
 Close 16
 Rem ******** Otros_Rgtros
 Open "AcuNom1.Dno" For Random As 16 Len = Len(Ot_Acum)
 FAcum1 = LOF(16) / Len(Ot_Acum)
 If facum = 0 Then Close 16: Kill "AcuNom1.Dno"
 Open "Infm04.txt" For Output As 12
 For r = 1 To Dacu1.Rows - 2
    If Dacu1.TextMatrix(r, 0) > "" Then
       Empleado = Dacu1.TextMatrix(r, 0)
       Get 3, Empleado, personal
       Get 5, Empleado, ArAcum
       Get 4, Empleado, Otros_Rgtros
       If FAcum1 > 0 Then
            Get 16, Empleado, Ot_Acum
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
     
     Rem CAMPO 6 APELLIDO PATERNO ******************************
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
     SubsidioAplicable = 1 - ((1 - empresa.psub) / 2)
     SubsidioAplicable = Format(1 - ((1 - empresa.psub) / 2), "#.0000")
     
     subdioforma = Format(SubsidioAplicable, "#.0000")
     If subdioforma < 1 Then
            DA_TO = SubsidioAplicable
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
        DA_TO = "STE750109F31"
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
     Rem CAMPO 26 Pagos por separacion *******************************
     DA_TO = RTrim("2")
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 27 Asimilado a salarios *******************************
     DA_TO = RTrim("2")
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 28 Pagos del patron a sus trabajadores *****************
     DA_TO = RTrim("1")
     Rem If Len(DA_TO) < 16 Then DA_TO = DA_TO + String(16 - Len(DA_TO), " ")
     DatoTe = DatoTe + DA_TO + "|"
     Rem campos 29 a 46 Ingresos por separacion *************************
     Rem For I = 29 To 46
        Rem DA_TO = ""
        Rem DatoTe = DatoTe + DA_TO + "|"
     Rem Next I
     
     Rem campo 47 ASIMILADOS A SALARIOS **********************************
     Rem DA_TO = ""
     Rem DatoTe = DatoTe + DA_TO + "|"
     Rem campo 48 IMPUESTO RETENDIO A ASIMILADOS A SALARIOS *************
     Rem DA_TO = ""
     Rem DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 49 Sueldos, salarios gravados
     DA_TO = LTrim(Str(CLng(ArAcum.Pnormal)))
     If ArAcum.Pnormal = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     
     Rem CAMPO 50 Sueldos, salarios EXENTOS
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 51 Gratificacion Anual Gravado
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
     Rem campo 52 Gratificacion anual Exento
     DA_TO = LTrim(Str(CLng(AguinaldoExento)))
     If AguinaldoExento = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 53 Viaticos y gastos de viaje gravado
     DA_TO = LTrim(Str(CLng(ArAcum.Pviaticos)))
     If ArAcum.Pviaticos = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 54 Viaticos y Gastos de Viaje Exento
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 55 Tiempo extraordinario Gravado
     Rem DA_TO = LTrim(Str(CLng(ArAcum.Pextra)))
     Rem If ArAcum.Pextra = 0 Then DA_TO = ""
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 56 Tiempo extraordinario Exento
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 57 Prima Vacacional Gravada
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
     Rem campo 58 Prima Vacacional Exenta
     DA_TO = LTrim(Str(CLng(PrimaVacacioNalExenta)))
     If PrimaVacacioNalExenta = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 59 Prima dominical gravada
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 60 Prima dominical exenta
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 61 PTU Gravada
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
     Rem campo 62 PTU Exenta
     DA_TO = LTrim(Str(CLng(PTUExenta)))
     If PTUExenta = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 63 Reembolso de gtos. medicos Gravados
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 64 Reembolso de gtos. medicos exentos
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 65 Fondo de ahorro Gravados
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 66 Fondo de ahorro Exento
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 67 Caja de ahorro Gravados
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 68 Caja de ahorro Exento
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 69 Vales de despensa Gravado
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 70 Vales de despensa Exento
     DA_TO = LTrim(Str(CLng(ArAcum.Pexenta)))
     If ArAcum.Pexenta = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 71 Ayuda para gtos de funeral Gravado
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 72 Ayuda para gtos de funeral exento
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 73 Contribuciones a cargo del trabajador pag.x el patron Gravado
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 74 Contribuciones a cargo del trabajador pag.x el patron Exento
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 75 Premios por puntualidad Gravado
     DA_TO = LTrim(Str(CLng(ArAcum.Pextra)))
     If ArAcum.Pextra = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 76 Premios por puntualidad Exento
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem Campo 77 to 100
     For i = 77 To 100
        DA_TO = ""
        DatoTe = DatoTe + DA_TO + "|"
     Next i
     Rem CAMPO 101 PAGOS EFECTUADOS X OTROS EMPLEADORES GRAVADO
     If FAcum1 > 0 Then
      If Ot_Acum.PPTU < (empresa.sm * 15) Then
            AcumSup_CorEx = 0
            
            AcumSup_CorEx = CLng(Ot_Acum.PPTU)
            Else
            AcumSup_CorEx = 0
            
            AcumSup_CorEx = CLng(empresa.sm * 15)
      End If
      AcumSup_Cor = 0
      AcumSup_Cor = CLng(Ot_Acum.Pnormal + Ot_Acum.Pagui + Ot_Acum.Pextra _
            + Ot_Acum.Potras + Ot_Acum.PPTU + Ot_Acum.Pvaca + Ot_Acum.Pviaticos) - AcumSup_CorEx
           If AcumSup_Cor > 0 Then
                 DA_TO = LTrim(Str(AcumSup_Cor))
                 Else
                 DA_TO = ""
            End If
           Else
           AcumSup_Cor = 0
           DA_TO = ""
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 102 PAGOS EFECTUADOS X OTROS EMPLEADORES EXENTO
     If (FAcum1 > 0) And (AcumSup_CorEx > 0) Then
           
           DA_TO = LTrim(Str(AcumSup_CorEx))
           Else
           DA_TO = ""
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 103 OTROS INGRESOS POR SALARIOS GRAVADOS
     DA_TO = LTrim(Str(CLng(ArAcum.Potras)))
     If ArAcum.Potras = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 104 OTROS INGRESOS POR SALARIOS EXENTOS
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 105 SUMA INGRESO GRAVADO
     SumaGravado = ArAcum.Pnormal + ArAcum.Pextra + ArAcum.Pviaticos + ArAcum.Potras _
                   + AguinaldoGravado + PrimaVacacionalGravada + PTUGravada
     DA_TO = LTrim(Str(CLng(SumaGravado)))
     If SumaGravado = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 106 SUMA INGRESO EXENTO
     SumaExento = AguinaldoExento + PrimaVacacioNalExenta + PTUExenta
     DA_TO = LTrim(Str(CLng(SumaExento)))
     If SumaExento = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 107 Impuesto retenido
     If ArAcum.DImpret < 0 Then ArAcum.DImpret = 0
     TOT_ret = 0: TOT_ret = ArAcum.DImpret
     DA_TO = LTrim(Str(CLng(TOT_ret)))
     If TOT_ret = 0 Then DA_TO = "0"
     DatoTe = DatoTe + DA_TO + "|"
     
     Rem CAMPO 108 Impuesto retenido por otros patrones
     If (FAcum1 > 0) Then
           DA_TO = LTrim(Str(CLng(Ot_Acum.DImpret)))
           Else
           DA_TO = ""
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 109 SALDO A FAVOR X COMPENSAR
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 110 SALDO A FAVOR DEL EJERCICIO ANTERIOR NO COMPENSADO DURANTE...
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 111 SUMA CREDITO AL SALARIO CALCULADO
     DA_TO = LTrim(Str(CLng(ArAcum.DCrApl)))
     If ArAcum.DCrApl = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 112 SUMA CREDITO AL SALARIO pagado
     DA_TO = LTrim(Str(CLng(ArAcum.DCrPag * -1)))
     If ArAcum.DCrPag = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 113 Ingresos de prestaciones de prevision social gravada
     If ArAcum.Pexenta > 0 Then
        DA_TO = LTrim(Str(CLng(ArAcum.Pexenta)))
        Else
        DA_TO = ""
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 114 Ingresos de prestaciones de prevision social exentas
     If ArAcum.Pexenta > 0 Then
        DA_TO = LTrim(Str(CLng(ArAcum.Pexenta)))
        Else
        DA_TO = ""
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 115 SUMA INGRESO GRAVADO
     'If AcumSup_Cor > 0 Then
         'SumaGravado = SumaGravado + AcumSup_Cor
         'DA_TO = LTrim(Str(CLng(SumaGravado)))
         'Else
         'DA_TO = ""
     'End If
     'DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 116 SUMA INGRESO Exento
     'If AcumSup_CorEx > 0 Then
         'SumaExento = SumaExento + AcumSup_CorEx
         'DA_TO = LTrim(Str(CLng(SumaExento)))
         'Else
         'DA_TO = ""
     'End If
     'DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 117 SUMA impuesto tarifa anual 115
     If CALCULOANUAL = "1" Then
            impto = 0
            If AcumSup_Cor > 0 Then
                
                Subidio_doble = (ArAcum.DSubioAp + Ot_Acum.DSubioAp) / (ArAcum.DSubioAp + Ot_Acum.DSubioAp + ArAcum.DSubNoap + Ot_Acum.DSubNoap)
                calc_anual CLng(SumaGravado), impto, Subidio_doble
                Else
                calc_anual CLng(SumaGravado), impto, empresa.psub
            End If
            DA_TO = LTrim(Str(CLng(imptotal)))
            Else
            DA_TO = ""
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 118 SUBSIDIO ACREDITABLE 116
     DA_TO = LTrim(Str(CLng(subdio)))
     If subdio = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 119 SUBSIDIO ACREDITABLE 117
     DA_TO = LTrim(Str(CLng(SUBDIONOACREDITABLE)))
     If SUBDIONOACREDITABLE = 0 Then DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 120 Subsidio acreditable fraccion III en 2001
     'DA_TO = ""
     'DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 121 Subsidio acreditable fraccion IV 2001
     'DA_TO = ""
     'DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 122 SUMA impuesto sobre ingresos acumulables 118
     If CALCULOANUAL = "1" Then
            DA_TO = LTrim(Str(CLng(imptotal - subdio)))
            Else
            DA_TO = ""
     End If
     DatoTe = DatoTe + DA_TO + "|"
     Rem CAMPO 123 Impuesto sobre ingresos no acumulables 119
     DA_TO = ""
     DatoTe = DatoTe + DA_TO + "|"
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

Private Sub Text1_GotFocus()
    SendKeys "{end}"
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
