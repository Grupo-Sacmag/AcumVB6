Attribute VB_Name = "ModAcu"
Type sc
    guarda As String * 64
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

Public AxNom As AxN, Aumento1 As Long, Kincenal As Integer
Public AxPer As Nacu, Ftem As Long, r As Long, a_opago As Integer
Public ArAcum As Nacu, SUMA_CREDITO_MES As Currency
Public Ot_Acum As Nacu, ta_r As Integer
Public temporal As te, DIR_CALC As String
Public SCont As sc, AcumSup_Cor As Currency, AcumSup_CorEx As Currency
Public Subidio_doble As Currency, DirecT_arifas As String
Sub derecha(ancho2, ltotal, cadena As String)
    ancho2 = 0
    ancho2 = (ltotal - Printer.TextWidth(cadena))
End Sub

