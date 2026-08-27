VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Verifica 
   Caption         =   "Archivos Incluidos en el Acumulado."
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4335
   Icon            =   "Verifica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Veri2 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7011
      _Version        =   393216
      BackColorBkg    =   -2147483633
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "Verifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Veri2.Rows = 1
    Veri2.Col = 0: Veri2.ColWidth(0) = 800: Veri2.CellFontBold = True: Veri2.CellAlignment = 3: Veri2.Text = "Orden"
    Veri2.Col = 1: Veri2.ColWidth(1) = 3000: Veri2.CellFontBold = True: Veri2.CellAlignment = 3: Veri2.Text = " Archivo "
    Open "AcuTemp" For Random As 10 Len = Len(temporal)
    Ftem = LOF(10) / Len(temporal)

    For r = 1 To Ftem: Get 10, r, temporal
    
      Veri2.AddItem Format(r, "####0") & Chr(9) & temporal.miarchivo
    Next r
    
    Close 10
End Sub
