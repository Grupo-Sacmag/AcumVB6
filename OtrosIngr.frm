VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form OtrosIngr 
   Caption         =   "Otros ingresos"
   ClientHeight    =   5565
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid IngTot 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   19
      FixedCols       =   4
      FocusRect       =   2
      FillStyle       =   1
   End
   Begin VB.Menu OIAr 
      Caption         =   "&Archivo"
      Begin VB.Menu OIArGu 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu OISep1 
         Caption         =   "-"
      End
      Begin VB.Menu OIArSal 
         Caption         =   "&Salida"
      End
   End
   Begin VB.Menu OIEd 
      Caption         =   "&Edicion"
      Begin VB.Menu OIEdSelT 
         Caption         =   "&Seleccionar todo"
      End
      Begin VB.Menu OIEdSep1 
         Caption         =   "-"
      End
      Begin VB.Menu OIEdCop 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "OtrosIngr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Exe_nto As Currency, Deotra, psub1 As Currency, r As Integer
Dim Ingreso_Aqui As Currency, Ingreso_Gravable1 As Currency, Mas_ingresos As Currency, OTpt_ret As Currency
Dim Ingreso_Gravable2 As Currency, Sub_Apl2 As Currency, Sub_NoAp2 As Currency, tot_diast As Currency
Sub IngExen()
   Exe_nto = 0
  For g = 6 To 10
   Select Case g
      Case 6, 9
      If Deotra.TextMatrix(r, g) <> "" Then
          If (empresa.sm * 15) > Deotra.TextMatrix(r, g) Then
              Exe_nto = Exe_nto + Deotra.TextMatrix(r, g)
              
              Else
              Exe_nto = Exe_nto + (empresa.sm * 15)
              
          End If
      End If
      Case 8
      If Deotra.TextMatrix(r, g) <> "" Then
          If (empresa.sm * 30) > Deotra.TextMatrix(r, g) Then
             Exe_nto = Exe_nto + Deotra.TextMatrix(r, g)
             
             Else
             Exe_nto = Exe_nto + (empresa.sm * 30)
             
          End If
      End If
      Case 10
      If Deotra.TextMatrix(r, g) <> "" Then
         Exe_nto = Exe_nto + Deotra.TextMatrix(r, g)
      End If
   End Select
  Next g
  
End Sub
Sub Enca_bzdo1()
    IngTot.TextMatrix(0, 0) = Deotra.TextMatrix(0, 0)
    IngTot.TextMatrix(0, 1) = Deotra.TextMatrix(0, 1)
    IngTot.TextMatrix(0, 2) = "Ingreso Total"
    IngTot.TextMatrix(0, 3) = "Exento"
    IngTot.TextMatrix(0, 4) = "Gravado"
    IngTot.TextMatrix(0, 5) = Deotra.TextMatrix(0, 13)
    IngTot.TextMatrix(0, 6) = Deotra.TextMatrix(0, 17)
    IngTot.TextMatrix(0, 7) = Deotra.TextMatrix(0, 14)
    IngTot.TextMatrix(0, 8) = Deotra.TextMatrix(0, 16)
    IngTot.TextMatrix(0, 9) = "Dias T."
    IngTot.TextMatrix(0, 10) = "Impto.Tot"
    IngTot.TextMatrix(0, 11) = "Sub-Acred."
    IngTot.TextMatrix(0, 12) = "Cred.al Sal."
    IngTot.TextMatrix(0, 13) = "Impto.Causado"
    For i = 0 To 13:: IngTot.Col = i: IngTot.CellFontBold = True: IngTot.CellAlignment = 4: IngTot.CellForeColor = vbBlue: Next
    IngTot.ColAlignment(9) = 4
End Sub
Sub Enca_bzdo()
    IngTot.TextMatrix(0, 0) = Deotra.TextMatrix(0, 0)
    IngTot.TextMatrix(0, 1) = Deotra.TextMatrix(0, 1)
    IngTot.TextMatrix(0, 2) = "Ingreso Total"
    IngTot.TextMatrix(0, 3) = "Exento"
    IngTot.TextMatrix(0, 4) = "Gravado"
    IngTot.TextMatrix(0, 5) = Deotra.TextMatrix(0, 13)
    IngTot.TextMatrix(0, 6) = Deotra.TextMatrix(0, 17)
    IngTot.TextMatrix(0, 7) = Deotra.TextMatrix(0, 14)
    IngTot.TextMatrix(0, 8) = Deotra.TextMatrix(0, 16)
    IngTot.TextMatrix(0, 9) = "Otr.Patr."
    IngTot.TextMatrix(0, 10) = "Otr.Exentos"
    IngTot.TextMatrix(0, 11) = "Otr.Ing.Grav."
    IngTot.TextMatrix(0, 12) = "Otr.Sub.Aplic."
    IngTot.TextMatrix(0, 13) = "Otr.Sub.No A."
    IngTot.TextMatrix(0, 14) = "Dias Tot."
    IngTot.TextMatrix(0, 15) = "Ing.Tot."
    IngTot.TextMatrix(0, 16) = "Sub.Apl."
    IngTot.TextMatrix(0, 17) = "I.Causado"
    IngTot.TextMatrix(0, 18) = "Dif.ret."
    For i = 0 To 18: IngTot.Col = i: IngTot.CellFontBold = True: IngTot.CellAlignment = 4: IngTot.CellForeColor = vbBlue: Next
    
End Sub
Private Sub Form_Load()
    Dim Aplic_sub As Currency, NoAplic_sub As Currency
    Set Deotra = Acu21.Dacu1
    Close 1
    Open "Empresa.dno" For Random As 1 Len = Len(empresa)
    cm = LOF(1) / Len(empresa)
    Get 1, cm, empresa
    Close 2
    Open "AcuNom2.dno" For Random As 2 Len = Len(ArAcum)
    dm = LOF(2) / Len(ArAcum)
    If dm > 0 Then
        psub1 = empresa.psub
        IngTot.Rows = 1
        IngTot.ColWidth(0) = 800: IngTot.ColWidth(1) = 3300
        For i = 2 To IngTot.Cols - 1: IngTot.ColWidth(i) = 1300: Next i
        Enca_bzdo
        For r = 1 To Acu21.Dacu1.Rows - 2
             IngExen
             rgtro = Deotra.TextMatrix(r, 0)
             Adi_cional
             Ingreso_Aqui = Deotra.TextMatrix(r, 11)
             Ingreso_Gravable1 = Ingreso_Aqui - Exe_nto
             Ingreso_Total = Ingreso_Gravable1 + Ingreso_Gravable2
             If tot_diast > 344 Then
                    If Deotra.TextMatrix(r, 13) <> "" Then Aplic_sub = Deotra.TextMatrix(r, 13) Else Aplic_sub = 0
                    If Deotra.TextMatrix(r, 17) <> "" Then NoAplic_sub = Deotra.TextMatrix(r, 17) Else NoAplic_sub = 0
                    psub = (Aplic_sub + Sub_Apl2) / (Aplic_sub + NoAplic_sub + Sub_Apl2 + Sub_NoAp2)
                    calc_anual Ingreso_Total, impto, psub
                    If Deotra.TextMatrix(r, 15) <> "" Then ImptoNomina = Deotra.TextMatrix(r, 15) Else ImptoNomina = 0
                    If Deotra.TextMatrix(r, 16) <> "" Then Crenomina = Deotra.TextMatrix(r, 16) Else Crenomina = 0
                    Dif_imp = ImptoRes.Apagar - ImptoNomina - Crenomina - OTpt_ret

             End If
             If Deotra.TextMatrix(r, 0) > 0 Then
             IngTot.AddItem Deotra.TextMatrix(r, 0) & Chr(9) & Deotra.TextMatrix(r, 1) _
                            & Chr(9) & Format(Ingreso_Aqui, "#,##0.00;(#,##0.00)") & Chr(9) & Format(Exe_nto, "#,##0.00;(#,##0.00)") _
                            & Chr(9) & Format(Ingreso_Gravable1, "#,##0.00;(#,##0.00)") & Chr(9) & _
                            Deotra.TextMatrix(r, 13) & Chr(9) & Deotra.TextMatrix(r, 17) & Chr(9) & _
                            Deotra.TextMatrix(r, 14) & Chr(9) & Deotra.TextMatrix(r, 16) & Chr(9) & _
                            Format(Mas_ingresos, "#,##0.00;(#,##0.00)") & Chr(9) & Format(Mon_Exento, "#,##0.00;(#,##0.00)") _
                            & Chr(9) & Format(Ingreso_Gravable2, "#,##0.00;(#,##0.00)") & Chr(9) & _
                            Format(Sub_Apl2, "#,##0.00;(#,##0.00)") & Chr(9) & Format(Sub_NoAp2, "#,##0.00;(#,##0.00)") _
                            & Chr(9) & Format(tot_diast, "#,##0.00;(#,##0.00)") & Chr(9) & Format(Ingreso_Total, "#,##0.00;(#,##0.00)") _
                            & Chr(9) & Format(psub, "#0.0000") & Chr(9) & Format(ImptoRes.Apagar, "#,##0.00;(#,##0.00)") _
                            & Chr(9) & Format(Dif_imp, "#,##0.00;(#,##0.00)")
              ImptoRes.Apagar = 0: ImptoRes.Calc = 0: ImptoRes.Cdto = 0: ImptoRes.Subdo = 0: Dif_imp = 0
             End If
    Next r
    For i = 0 To IngTot.Rows - 1: IngTot.Col = 4: IngTot.Row = i: IngTot.CellBackColor = vbYellow: Next i
    IngTot.Rows = IngTot.Rows + 1
    Suma_r
    IngTot.Row = IngTot.Rows - 1
    For i = 0 To 18:: IngTot.Col = i: IngTot.CellFontBold = True: IngTot.CellForeColor = vbBlue: Next i
    Else
      psub = empresa.psub
      IngTot.Rows = 1
      IngTot.ColWidth(0) = 800: IngTot.ColWidth(1) = 3300
      For i = 2 To IngTot.Cols - 1: IngTot.ColWidth(i) = 1300: Next i
      Enca_bzdo1
      For r = 1 To Acu21.Dacu1.Rows - 2
             IngExen
             Ingreso_Aqui = Deotra.TextMatrix(r, 11)
             Ingreso_Gravable1 = Ingreso_Aqui - Exe_nto
             If (Deotra.TextMatrix(r, 2) <> "") And (Deotra.TextMatrix(r, 2) >= 300) Then
                    calc_anual Ingreso_Gravable1, impto, psub
                    If Deotra.TextMatrix(r, 15) <> "" Then ImptoNomina = Deotra.TextMatrix(r, 15) Else ImptoNomina = 0
                    If Deotra.TextMatrix(r, 16) <> "" Then Crenomina = Deotra.TextMatrix(r, 16) Else Crenomina = 0
                    Dif_imp = ImptoRes.Apagar - ImptoNomina - Crenomina
                    
                    Else
                    Rem nada
                    
             End If
             If Deotra.TextMatrix(r, 0) > 0 Then
             IngTot.AddItem Deotra.TextMatrix(r, 0) & Chr(9) & Deotra.TextMatrix(r, 1) _
                    & Chr(9) & Format(Ingreso_Aqui, "#,##0.00;(#,##0.00)") & Chr(9) & _
                    Format(Exe_nto, "#,##0.00;(#,##0.00)") & Chr(9) & Format(Ingreso_Gravable1, "#,##0.00;(#,##0.00)") _
                    & Chr(9) & Deotra.TextMatrix(r, 13) & Chr(9) & Deotra.TextMatrix(r, 17) & Chr(9) & _
                    Deotra.TextMatrix(r, 14) & Chr(9) & Deotra.TextMatrix(r, 16) & Chr(9) & _
                    Format(Deotra.TextMatrix(r, 2), "#,##0") & Chr(9) & _
                    Format(ImptoRes.Calc, "#,##0.00;(#,##0.00)") & Chr(9) & Format(ImptoRes.Subdo, "#,##0.00;(#,##0.00)") _
                    & Chr(9) & Format(ImptoRes.Cdto, "#,##0.00;(#,##0.00)") & Chr(9) & Format(ImptoRes.Apagar, "#,##0.00;(#,##0.00)") _
                    & Chr(9) & Format(Dif_imp, "#,##0.00;(#,##0.00)")

             End If
             ImptoRes.Apagar = 0: ImptoRes.Calc = 0: ImptoRes.Cdto = 0: ImptoRes.Subdo = 0
       Next r
       For i = 0 To IngTot.Rows - 1: IngTot.Col = 4: IngTot.Row = i: IngTot.CellBackColor = vbYellow: Next i
       IngTot.Rows = IngTot.Rows + 1
       Suma_r
       IngTot.Row = IngTot.Rows - 1
       For i = 0 To 18:: IngTot.Col = i: IngTot.CellFontBold = True: IngTot.CellForeColor = vbBlue: Next i
    End If
End Sub
Sub Suma_r()
  Dim Suma_ndo As Currency
   For r = 2 To IngTot.Cols - 1
       Suma_ndo = 0
       For i = 1 To IngTot.Rows - 2
          If IngTot.TextMatrix(i, r) <> "" Then
                Suma_ndo = Suma_ndo + IngTot.TextMatrix(i, r)
          End If
       Next i
       IngTot.TextMatrix(IngTot.Rows - 1, r) = Format(Suma_ndo, "#,##0.00;(#,##0.00)")
   Next r
   
End Sub
Sub Adi_cional()
   tot_diast = 0
   If rgtro <= (LOF(2) / Len(ArAcum)) Then
        Get 2, rgtro, ArAcum
        Mas_ingresos = ArAcum.Pagui + ArAcum.Pexenta + ArAcum.Pextra + ArAcum.Pnormal + ArAcum.Potras + ArAcum.Pvaca + ArAcum.Pviaticos + ArAcum.PPTU
        Exencion2
        Ingreso_Gravable2 = Mas_ingresos - Mon_Exento
        Sub_Apl2 = ArAcum.DSubioAp
        Sub_NoAp2 = ArAcum.DSubNoap
        tot_diast = Deotra.TextMatrix(r, 2) + ArAcum.Pdias
        OTpt_ret = ArAcum.DImpret + ArAcum.DCrPag
   End If
End Sub
Private Sub Form_Resize()
    IngTot.Width = OtrosIngr.ScaleWidth
    IngTot.Height = OtrosIngr.ScaleHeight - 200
End Sub

Private Sub IngTot_KeyPress(KeyAscii As Integer)
     Select Case (KeyAscii)
     Case 13
     Rem nada
     DatosPersonales.Show
     End Select
End Sub

Private Sub OIArSal_Click()
    Unload Me
End Sub

Private Sub OIEdCop_Click()
    Clipboard.Clear
   For i = IngTot.Row To IngTot.RowSel
      For f = IngTot.Col To IngTot.ColSel
            Clipboard.SetText Clipboard.GetText + IngTot.TextMatrix(i, f) & Chr(9)
      Next f
      Clipboard.SetText Clipboard.GetText + Chr(13)
   Next i
   difer = IngTot.RowSel - IngTot.Row

End Sub

Private Sub OIEdSelT_Click()
   IngTot.Col = 0
   IngTot.Row = 0
   IngTot.ColSel = IngTot.Cols - 1
   IngTot.RowSel = IngTot.Rows - 1
End Sub
