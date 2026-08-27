VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DatosConta 
   Caption         =   "Ubicacion Archivo de Datos para captura contable"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Examinar 
      Caption         =   "Examinar..."
      Height          =   495
      Index           =   0
      Left            =   5160
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "DatosConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Persona As Integer
Sub Determina()
   Rem
End Sub

Private Sub Examinar_Click(Index As Integer)
   Select Case Index
      Case 0
        personal.nom = UCase(Text1(1).Text)
        personal.ape1 = UCase(Text1(2).Text)
        personal.ape2 = UCase(Text1(3).Text)
        personal.rfc = UCase(Text1(4).Text)
        Put 2, Persona, personal
        Unload DatosConta
        
      Case 1
        Unload DatosConta
   End Select
End Sub

Private Sub Form_Load()
  If Acu21.Dacu1.Row < 1 Then
      MsgBox "No existe personal seleccionado"
      Examinar_Click 1
     Else
     Persona = Acu21.Dacu1.TextMatrix(Acu21.Dacu1.Row, 0)
     Dim W
     Label1(0).Move 200, 700, 2700, 700
     Label1(0).Alignment = 1
     Label1(0).FontBold = True
     Text1(0).Move 3200, 700, 2700, 300
     Text1(0).Text = ""
     Examinar(0).Move 6000, 700, 1700, 300
     For W = 1 To 1
         Load Examinar(W)
         Examinar(W).Move 6000, (W + 1) * 700, 1700, 300
         Examinar(W).Visible = True
     Next W
     Examinar(0).Caption = "Archivar"
     Examinar(1).Caption = "Salir"
     For W = 1 To 4
        Load Label1(W)
        Load Text1(W)
        Label1(W).Move 200, (W + 1) * 700, 2700, 700
        Text1(W).Move 3200, (W + 1) * 700, 2700, 300
        Label1(W).Visible = True
        Text1(W).Visible = True
        Text1(W).Text = ""
        Label1(W).Alignment = 1
        Label1(W).FontBold = True
     Next W
    Label1(0).Caption = "Numero : "
    Label1(1).Caption = "Nombre: "
    Label1(2).Caption = "Apellido Paterno : "
    Label1(3).Caption = "Apellido Materno : "
    Label1(4).Caption = "R. F. C. : "
    regis_tro
  End If
    Rem Label1(5).Caption = "SubCuenta de Abono : "
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Rem nada a ver que pasa
End Sub
Sub regis_tro()
   Open "Personal.dno" For Random As 2 Len = Len(personal)
   cm = LOF(2) / Len(personal)
   Persona = Acu21.Dacu1.TextMatrix(Acu21.Dacu1.Row, 0)
   Get 2, Persona, personal
   Text1(0).Text = Persona
   Text1(1).Text = UCase(personal.nom)
   Text1(2).Text = UCase(personal.ape1)
   Text1(3).Text = UCase(personal.ape2)
   Text1(4).Text = UCase(personal.rfc)
  End Sub

