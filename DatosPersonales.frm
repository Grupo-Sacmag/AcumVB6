VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form DatosPersonales 
   Caption         =   "Datos Personales y Adicionales"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Dp_hj 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9128
      _Version        =   393216
   End
End
Attribute VB_Name = "DatosPersonales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Inicia_DP()
    Dim Suma_Ing As Currency, Suma_exe As Currency, Suma_IT As Currency
    Dim Suma_gra As Currency, Dif_Dias As Currency
    Set MU = OtrosIngr.IngTot
    Rem Set MO = acu21.Dacu1
    DatosPersonales.Caption = MU.TextMatrix(MU.Row, 1)
    Suma_Ing = MU.TextMatrix(OtrosIngr.IngTot.Row, 2)
    Suma_Ing = Suma_Ing + MU.TextMatrix(OtrosIngr.IngTot.Row, 9)
    Suma_exe = MU.TextMatrix(OtrosIngr.IngTot.Row, 3)
    Suma_exe = Suma_exe + MU.TextMatrix(OtrosIngr.IngTot.Row, 10)
    Suma_gra = MU.TextMatrix(OtrosIngr.IngTot.Row, 4)
    Suma_gra = Suma_exe + MU.TextMatrix(OtrosIngr.IngTot.Row, 11)
    Suma_IT = MU.TextMatrix(OtrosIngr.IngTot.Row, 4)
    Suma_IT = Suma_IT + MU.TextMatrix(OtrosIngr.IngTot.Row, 11)
    Dif_Dias = MU.TextMatrix(MU.Row, 14)
    Dif_Dias = Dif_Dias + Acu21.Dacu1.TextMatrix(MU.Row, 2)
    Dp_hj.AddItem "¿Se efectuo Cálculo? " & Chr(9) & "" & Chr(9) & ""
    Dp_hj.AddItem "Fecha de Alta " & Chr(9) & "" & Chr(9) & ""
    Dp_hj.AddItem "Fecha de Baja" & Chr(9) & "" & Chr(9) & ""
    Dp_hj.AddItem "Dias Trabajados " & Chr(9) & Acu21.Dacu1.TextMatrix(MU.Row, 2) & Chr(9) & MU.TextMatrix(MU.Row, 14) _
                    & Chr(9) & Format(Dif_Dias, "#,##0")
    Dp_hj.AddItem "Ingreso Total " & Chr(9) & MU.TextMatrix(MU.Row, 2) & Chr(9) & _
                   MU.TextMatrix(MU.Row, 9) & Chr(9) & Format(Suma_Ing, "#,##0.00;(#,##0.00)")
    Dp_hj.AddItem "Exento " & Chr(9) & MU.TextMatrix(OtrosIngr.IngTot.Row, 3) _
                   & Chr(9) & MU.TextMatrix(OtrosIngr.IngTot.Row, 10) & Chr(9) & Format(Suma_exe, "#,##0.00;(#,##0.00)")
    Dp_hj.AddItem "Gravado " & Chr(9) & MU.TextMatrix(OtrosIngr.IngTot.Row, 4) _
                   & Chr(9) & MU.TextMatrix(OtrosIngr.IngTot.Row, 11) & Chr(9) & Format(Suma_gra, "#,##0.00;(#,##0.00)")
    Dp_hj.AddItem "Impuesto Total " & Chr(9) & "" & Chr(9) & ""
    Dp_hj.AddItem "Subsidio Aplicado " & Chr(9) & "" & Chr(9) & ""
    Dp_hj.AddItem "Subsidio no Aplicado " & Chr(9) & "" & Chr(9) & ""
    Dp_hj.AddItem "Credito Aplicado  " & Chr(9) & "" & Chr(9) & ""
    Dp_hj.AddItem "Crédito Pagado " & Chr(9) & "" & Chr(9) & ""
    Dp_hj.AddItem "Porcentaje de Subsidio " & Chr(9) & "" & Chr(9) & ""
    Dp_hj.AddItem "Impuesto Retenido " & Chr(9) & "" & Chr(9) & ""
End Sub

Private Sub Form_Load()
   Dp_hj.Cols = 4: Dp_hj.Rows = 1
   Dp_hj.ColWidth(0) = 2500: For i = 1 To 3: Dp_hj.ColWidth(i) = 1300: Next i
   Inicia_DP
    'DEFINECOL
    'CAP1.Clear
    Open "C:\2015\CONT\NOMINA\PERSONAL.dno" For Random As #2 Len = Len(personal)
    Open "C:\2015\CONT\NOMINA\ENE22015.NOM" For Random As #1 Len = Len(nomina)
    cm = LOF(1) / Len(nomina)
    ArchivoTexto = "C:\2015\CONT\NOMINA\MEZCLA3.TXT"
    Open ArchivoTexto For Output As #3
    For r = 1 To cm: Get 2, r, personal: Get 1, r, nomina
    If (nomina.sueldo > 0) Then
        'CAP1.AddItem r & Chr(9) & personal.nom & Chr(9) & personal.ape1 & Chr(9) & nomina.sueldo & Chr(9) & personal.imss & Chr(9) & personal.rfc
        Write #3, personal.nom, personal.ape1, nomina.sueldo, personal.imss, personal.rfc
    End If
    Next r
    Close
End Sub
