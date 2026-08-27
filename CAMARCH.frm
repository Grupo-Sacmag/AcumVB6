VERSION 5.00
Begin VB.Form Camdir 
   Caption         =   "Cambio de Subdirectorio"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   Icon            =   "CAMARCH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   2895
   End
   Begin VB.FileListBox File1 
      Height          =   4770
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
   Begin VB.DirListBox Dir1 
      Height          =   4140
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   5160
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Directorio :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "Camdir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()

    File1_Click
    File1_DblClick

End Sub


Private Sub Dir1_KeyPress(KeyAscii As Integer)

    File1_Click
    File1_DblClick

End Sub


Private Sub Drive1_Change()

    On Error GoTo manejodrive

    ChDrive Drive1.Drive

    Dir1.Path = Drive1.Drive
    Dir1 = Dir1.Path

    Dir1_Change

    GoTo saledriv


manejodrive:

    Drive1.Drive = "C:"
    Dir1 = "C:\"


saledriv:

End Sub


Private Sub File1_Click()

    If Dir1.Path <> Dir1.List(Dir1.ListIndex) Then

        Dir1.Path = Dir1.List(Dir1.ListIndex)

        File1 = Dir1.Path

        Exit Sub

    End If

    File1 = Dir1.Path

End Sub


Private Sub File1_DblClick()

    On Error GoTo ErrorDirectorio

    File1 = Dir1.Path

    ChDir CurDir(Dir1)

    Label2.Caption = Dir1.Path

    Exit Sub


ErrorDirectorio:

    Err.Clear

    MsgBox "No fue posible acceder al directorio seleccionado.", _
           vbExclamation, _
           "Directorio"

End Sub


Private Sub Form_Load()

    Label2.Caption = Dir1

End Sub


Private Sub GuardarDirectorio()

    On Error GoTo ErrorGuardar

    Close 3

    Open "C:\GconTa\sccontr.soc" For Random As 3 Len = Len(SCont)

    SCont.guarda = Dir1

    Put 3, 1, SCont

    Close 3

    Exit Sub


ErrorGuardar:

    Close 3

    MsgBox "No fue posible guardar el directorio seleccionado." & _
           vbCrLf & _
           Err.Description, _
           vbExclamation, _
           "Directorio"

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim Respuesta As Integer

    Respuesta = MsgBox( _
        "¿Desea guardar el directorio seleccionado?" & vbCrLf & vbCrLf & _
        "SÍ = Guardar el directorio y continuar." & vbCrLf & _
        "NO = Cerrar la aplicación sin guardar." & vbCrLf & _
        "CANCELAR = Regresar al selector.", _
        vbYesNoCancel + vbQuestion, _
        "Cambio de directorio")

    Select Case Respuesta


        '==================================================
        ' SÍ
        ' Guardar la ruta exactamente como lo hacía
        ' originalmente Camdir.
        '==================================================
        Case vbYes

            On Error GoTo ErrorDirectorio

            ' Intentar colocarnos en el directorio
            ' seleccionado.
            ChDir Dir1.Path

            ' NO verificamos archivos.
            GuardarDirectorio

            Exit Sub


        '==================================================
        ' NO
        ' Cerrar TODA la aplicación sin cambiar
        ' sccontr.soc.
        '==================================================
        Case vbNo

            Close
            End


        '==================================================
        ' CANCELAR
        ' Regresar al selector de directorios.
        '==================================================
        Case vbCancel

            Cancel = True
            Exit Sub

    End Select

    Exit Sub


ErrorDirectorio:

    Err.Clear

    MsgBox "No fue posible acceder al directorio seleccionado." & _
           vbCrLf & vbCrLf & _
           "Seleccione otro directorio.", _
           vbExclamation, _
           "Directorio no disponible"

    Cancel = True

End Sub

' comentario
