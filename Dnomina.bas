Attribute VB_Name = "Dnomina"
Type per
    nom As String * 20
    ape1 As String * 20
    ape2 As String * 20
    rfc As String * 18
    imss As String * 18
    fal As String * 12
    fab As String * 12
    ingr As Currency
    viat As Currency
    otras As Currency
    integrado As Currency
 End Type
 Type nom
     dias As Currency
     hsnor As Currency
     hs_no As Currency
     hsdbl As Currency
     hs_db As Currency
     hstri As Currency
     hs_tr As Currency
     ispt As Currency
     crdsal As Currency
     imss As Currency
     sueldo As Currency
     hs_nor As Currency
     hs_dbl As Currency
     hs_tri As Currency
     viaticos As Currency
     pvac As Currency
     otras As Currency
     aguin As Currency
     ptu As Currency
     exentos As Currency
     prestamos As Currency
     fonacot As Currency
     telefono As Currency
     otraded As Currency
  End Type
  
  Type empre
       name As String * 60
       ao As Integer
       sm As Currency
       psub As Currency
       fecha As String * 14
  End Type
 Type da_id
       Emp_Rfc As String * 25
       Emp_Dom As String * 70
       Rep_Legapp As String * 20
       Rep_Legapm As String * 20
       Rep_Legapn As String * 20
       Rep_Rfc As String * 25
       Rep_Curp As String * 25
       suc As String * 4
       cta As String * 12
       dias As Integer
  End Type
 Type ult
     ColIni As Long
     ColFin As Long
     RenIni As Long
     RenFin As Long
     num As Long
     ubi As Integer
     renglon As Long
     texto As String * 30
     texto1 As String * 30
     Poliza As Integer
     impresion As Integer
 End Type
 Type nomco
     ArchImp As String * 50
     PSubDi As Currency
     subdio As Currency
     subapl As Currency
     subNap As Currency
     CreTot As Currency
     CredNe As Currency
     ImpTot As Currency
End Type
Type art
     liminf As Currency
     limsup As Currency
     cuotaf As Currency
     porcsl As Currency
  End Type
  Type subs
     liminfs As Currency
     limsups As Currency
     cuotafs As Currency
     porcsls As Currency
  End Type
   Type cred
     crede As Currency
     crea As Currency
     cresam As Currency
  End Type
  Type Impre
        Calc As Currency
        Subdo As Currency
        Cdto As Currency
        Apagar As Currency
  End Type
  Type JS
       Acredit As Currency
       NoAcred As Currency
       Total As Currency
  End Type
  Type CRD
       Aplicado As Currency
       Pagado As Currency
  End Type
Type basini
     datoarch As String * 64
 End Type
 Public Basico As basini, ImptoRes As Impre, JSub As JS
 Public JCre As CRD
 Public Nom_Com As nomco
 Public ultimo As ult, Mon_Exento As Currency
 Public Dat_ide As da_id
 Public empresa As empre, InGresos As Currency
 Public nomina As nom, DeDucciones As Currency
 Public personal As per, neto As Currency
 Public rgtro As Integer
 Public mm(20), dd(20) As Integer, CREDITOSAL As Currency
 Public SUBDIONOACREDITABLE As Currency, subdio As Currency, imptotal As Currency
 Public impto As Currency
 Public articulo As art, subsidio As subs, credito As cred
 Sub calc_anual2(base, impto, psub)
    Dim ISR_1 As Currency, SUB_1 As Currency, CRED_1 As Currency
    Rem calculo base, impto, psub
    ISR_1 = Nom_Com.ImpTot: SUB_1 = Nom_Com.subapl: CRED_1 = Nom_Com.CredNe
    Close 3: Close 4: Close 5
    Open (Dir_imptos + "ISR177.03") For Random As #3 Len = Len(articulo)
    Dem = LOF(3) / Len(articulo)
    Open (Dir_imptos + "SUB178.03") For Random As #4 Len = Len(subsidio)
    EM = LOF(4) / Len(subsidio)
    Open (Dir_imptos + "CRE116.03") For Random As #5 Len = Len(credito)
    eem = LOF(5) / Len(credito)
    Rem  **** CALCULAR IMPUESTO ****
    
    baseor = base
    baseanual = 0
    
    Rem detbase
    
    Rem If regtro = 587 Then Stop
    Rem base = baseanual + Base_anual1
    
    For i = 1 To Dem: Get 3, i, articulo
     If base > (articulo.liminf) And base < (articulo.limsup) Then
           marginal = ((base - articulo.liminf) * (articulo.porcsl / 100))
           impto = marginal + articulo.cuotaf
           Nom_Com.ImpTot = impto - ArAcum.DImpto - ImtoTo_otra: Rem   *********************
           Imp_mag = (articulo.porcsl / 100)
           i = Dem
     End If
    Next i
    mientras = psub
GoTo SuBsidio_Salto
    For i = 1 To EM: Get 4, i, subsidio
     If P_sub1 > 0 Then psub = P_sub1
     If base > (subsidio.liminfs) And base < (subsidio.limsups) Then
           marginal2 = ((base - subsidio.liminfs) * Imp_mag)
           subdio = (marginal2 * subsidio.porcsls / 100) + (subsidio.cuotafs)
           Rem subdio = (marginal * subsidio.porcsls / 100) + (subsidio.cuotafs)
           subdiono = subdio
           subdio = (subdio * psub)
           subdiono = subdiono - subdio
           Nom_Com.subapl = subdio - ArAcum.DSubioAp - Sub_Aplic1
           Nom_Com.subNap = subdiono - ArAcum.DSubNoap - (Psub_Extra - Sub_Aplic1)
           Nom_Com.subdio = Nom_Com.subapl + Nom_Com.subNap
           Nom_Com.PSubDi = psub
           i = EM
     End If
    Next i
SuBsidio_Salto:
     subdio = ArAcum.DSubioAp + SUB_1
    psub = mientras
Rem ************************** ELIMINA CREDITO ANUAL *********************************
    'For i = 1 To eem: Get 5, i, credito
     'If base > (credito.crede) And base < credito.crea Then
         'creere = (credito.cresam)
         'nom_com.CreTot = creere - ArAcum.DCrApl - Crd_deotra
         Rem If nom_com.CreTot < 0 Then nom_com.CreTot = 0
         
     'End If
    'Next i
Rem *************************  ELIMINA CREDITO ANUAL *********************************

    creere = SUMA_CREDITO_MES + CREDITO_PROV
    
    Rem impto = impto - subdio - creere
    
    impto = impto - subdio
    Rem ***********************************************************************************************
    Rem SI EL  IMPUESTO ANUAL ES MAYOR AL SUBSIDIO ENTONCES SE CALCULA EL IMPUESTO DE LA ULTIMA NOMINA*
    Rem ***********************************************************************************************
    
    If impto > 0 Then
      Rem impto = impto - ArAcum.DImpret - ArAcum.DCrPag - Imto_deotra - Crpag_deotra
      
      impto = impto - ArAcum.DImpret - Imto_deotra
      Else
      impto = CRED_1
      Rem AQUI VA EL OTRO
    End If
    
    Rem If (impto < 0) And (ArAcum.DSubioAp > 1) Then impto = 0
    Rem impto = impto - ArAcum.DImpret + ArAcum.DCrPag - Imto_deotra + Crpag_deotra
    Rem If impto < 0 Then
        Rem If ArAcum.DCrApl = 0 Then
                Rem nom_com.CredNe = 0
                Rem Else
            Rem Stop
            Rem impto = CRED_1
            
     Rem End If
        Rem Else
        Rem nom_com.CredNe = 0
    Rem End If
    Rem Put 14, regtro, Nom_Com
    Rem impto = impto - subdio
    base = baseor
    Close 10
End Sub
  Sub calc_anual(base, impto, psub)
    Close 13: Close 14: Close 15
    Open "C:\TARIFA10\ISR177.03" For Random As #13 Len = Len(articulo)
    Dem = LOF(13) / Len(articulo)
    Open "C:\TARIFA10\SUB178.03" For Random As #14 Len = Len(subsidio)
    EM = LOF(14) / Len(subsidio)
    Open "C:\TARIFA10\CRE116.03" For Random As #15 Len = Len(credito)
    eem = LOF(15) / Len(credito)
    Rem  **** CALCULAR IMPUESTO ****
    baseor = base
    Rem detbase
    base = base
    For i = 1 To Dem: Get 13, i, articulo
     If base > (articulo.liminf) And base < (articulo.limsup) Then
           marginal = ((base - articulo.liminf) * (articulo.porcsl / 100))
           impto = marginal + articulo.cuotaf
           Imp_mag = (articulo.porcsl / 100)
           i = Dem
     End If
    Next i
    ImptoRes.Calc = impto
    imptotal = ImptoRes.Calc
 GoTo SALTASUB
    For i = 1 To EM: Get 14, i, subsidio
     If base > (subsidio.liminfs) And base < (subsidio.limsups) Then
           marginal2 = ((base - subsidio.liminfs) * Imp_mag)
           subdio = (marginal2 * subsidio.porcsls / 100) + (subsidio.cuotafs)
           Rem subdio = (marginal * subsidio.porcsls / 100) + (subsidio.cuotafs)
           SUBDIONOACREDITABLE = subdio
           JSub.Total = SUBDIONOACREDITABLE
           subdio = (subdio * psub)
           JSub.Acredit = subdio
           SUBDIONOACREDITABLE = SUBDIONOACREDITABLE - subdio
           JSub.NoAcred = SUBDIONOACREDITABLE
           i = EM
     End If
     
    Next i
    Rem ImptoRes.Subdo = JSub.Acredit
SALTASUB:
GoTo SALTACRE:
    ImptoRes.Subdo = 0
    For i = 1 To eem: Get 15, i, credito
     If base > (credito.crede) And base < credito.crea Then
         creere = (credito.cresam)
         JCre.Aplicado = creere
     End If
    Next i
    ImptoRes.Cdto = JCre.Aplicado
    creere = 0
    impto = impto - subdio - JCre.Aplicado
    ImptoRes.Apagar = impto
    base = baseor
    If impto < 0 Then JCre.Pagado = impto: impto = 0
SALTACRE:
    
    Close 13, 14, 15
  End Sub

 Sub veridir()
 Rem On Error GoTo corrdire
    
    Open "C:\GconTa\perma.dno" For Random As #7 Len = Len(Basico)
    fin_basico = LOF(7) / Len(Basico)
    
      If fin_basico > 0 Then
         Get 7, 1, Basico
         Direc_torio = RTrim(Basico.datoarch)
         
         If Direc_torio <> "" Then
                Rem Acu2.Drive1.Drive = Left(Direc_torio, 2)
                ChDir Direc_torio
                Rem Acu2.Dir1 = Direc_torio
                Close 7
         End If
         Open "C:\GconTa\perma.dno" For Random As #7 Len = Len(Basico)
         fin_basico = LOF(7) / Len(Basico)
         If fin_basico > 1 Then
           Get 7, 2, Basico
           dir_obras = RTrim(Basico.datoarch)
         End If
      End If
      Close 7
      
      GoTo saltoerror
corrdire:
   Close 7
   ChDir "C:\GconTa"
   Exit Sub
saltoerror:
   
 End Sub

 Sub colocar(ancho2, valor$, us_o As String)
     ancho2 = 0
     ancho = Printer.TextWidth(valor$)
     ancho1 = Printer.TextWidth(us_o)
     ancho2 = ancho1 - ancho
     Rem Printer.CurrentX = Printer.currex + ancho2
 End Sub
Sub centrar(ancho2, micadena As String, anchototal As Long)
     ancho2 = 0
     ancho = Printer.TextWidth(micadena) / 2
     ancho1 = anchototal / 2
     ancho2 = ancho1 - ancho
End Sub

Sub Exencion2()
   Mon_Exento = ArAcum.Pexenta
   If ArAcum.Pvaca < (empresa.sm * 15) Then
                Mon_Exento = Mon_Exento + ArAcum.Pvaca
                Else
                Mon_Exento = Mon_Exento + (empresa.sm * 15)
   End If
   If ArAcum.PPTU < (empresa.sm * 15) Then
                Mon_Exento = Mon_Exento + ArAcum.PPTU
                Else
                Mon_Exento = Mon_Exento + (empresa.sm * 15)
   End If
   If ArAcum.Pagui < (empresa.sm * 30) Then
                Mon_Exento = Mon_Exento + ArAcum.Pagui
                Else
                Mon_Exento = Mon_Exento + (empresa.sm * 30)
    
   End If
            
   
End Sub
