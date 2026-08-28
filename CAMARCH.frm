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

    On Error GoTo ErrorDirectorio

    File1.Path = Dir1.Path
    Label2.Caption = Dir1.Path

    Exit Sub

ErrorDirectorio:

    Err.Clear

End Sub


Private Sub Dir1_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    File1.Path = Dir1.Path
    Label2.Caption = Dir1.Path

    On Error GoTo 0

End Sub


Private Sub Drive1_Change()

    On Error GoTo manejodrive

    ChDrive Left$(Drive1.Drive, 1)

    Dir1.Path = Drive1.Drive
    File1.Path = Dir1.Path

    Label2.Caption = Dir1.Path

    Exit Sub


manejodrive:

    Err.Clear

    On Error Resume Next

    Drive1.Drive = "C:"
    ChDrive "C"
    Dir1.Path = "C:\"
    File1.Path = "C:\"
    Label2.Caption = "C:\"

    On Error GoTo 0

End Sub


Private Sub File1_Click()

    ' File1 solamente debe mostrar los archivos
    ' del directorio seleccionado.
    '
    ' NO debe cambiar Dir1.Path.

    On Error Resume Next

    File1.Path = Dir1.Path
    Label2.Caption = Dir1.Path

    On Error GoTo 0

End Sub


Private Sub File1_DblClick()

    Dim RutaSeleccionada As String

    RutaSeleccionada = Trim$(Dir1.Path)

    If RutaSeleccionada = "" Then Exit Sub

    If CambiarADirectorio(RutaSeleccionada) Then

        File1.Path = RutaSeleccionada
        Label2.Caption = RutaSeleccionada

    Else

        MsgBox "No fue posible acceder al directorio seleccionado." & _
               vbCrLf & vbCrLf & _
               RutaSeleccionada, _
               vbExclamation, _
               "Directorio"

    End If

End Sub


Private Sub Form_Load()

    On Error Resume Next

    Label2.Caption = Dir1.Path
    File1.Path = Dir1.Path

    On Error GoTo 0

End Sub


Private Function CambiarADirectorio(ByVal Ruta As String) As Boolean

    Dim Unidad As String

    CambiarADirectorio = False

    Ruta = Trim$(Ruta)

    If Ruta = "" Then Exit Function

    On Error GoTo ErrorDirectorio


    '==========================================================
    ' CAMBIAR UNIDAD
    '==========================================================
    If Len(Ruta) >= 2 Then

        If Mid$(Ruta, 2, 1) = ":" Then

            Unidad = Left$(Ruta, 1)
            ChDrive Unidad

        End If

    End If


    '==========================================================
    ' CAMBIAR DIRECTORIO
    '==========================================================
    ChDir Ruta


    '==========================================================
    ' COMPROBAR QUE REALMENTE QUEDAMOS EN ESA CARPETA
    '==========================================================
    If StrComp( _
            QuitarDiagonalFinal(CurDir$), _
            QuitarDiagonalFinal(Ruta), _
            vbTextCompare) <> 0 Then

        Exit Function

    End If


    CambiarADirectorio = True

    Exit Function


ErrorDirectorio:

    Err.Clear
    CambiarADirectorio = False

End Function


Private Function QuitarDiagonalFinal(ByVal Ruta As String) As String

    Ruta = Trim$(Ruta)

    If Len(Ruta) > 3 Then

        Do While Right$(Ruta, 1) = "\"

            Ruta = Left$(Ruta, Len(Ruta) - 1)

        Loop

    End If

    QuitarDiagonalFinal = Ruta

End Function


Private Sub GuardarDirectorio()

    Dim RutaSeleccionada As String

    On Error GoTo ErrorGuardar

    RutaSeleccionada = Trim$(Dir1.Path)

    If RutaSeleccionada = "" Then

        MsgBox "No hay un directorio seleccionado.", _
               vbExclamation, _
               "Directorio"

        Exit Sub

    End If


    '==========================================================
    ' COMPROBAR PRIMERO QUE LA RUTA SEA REAL
    '==========================================================
    If CambiarADirectorio(RutaSeleccionada) = False Then

        MsgBox "No fue posible acceder al directorio seleccionado." & _
               vbCrLf & vbCrLf & _
               RutaSeleccionada, _
               vbExclamation, _
               "Directorio"

        Exit Sub

    End If


    '==========================================================
    ' RECREAR SCCONTR.SOC
    '
    ' IMPORTANTE:
    ' Antes usaba un registro de 64 caracteres.
    ' Ahora utiliza 260.
    '
    ' Por eso eliminamos el archivo anterior antes de guardar.
    '==========================================================
    Close 3

    If Dir$("C:\GconTa\sccontr.soc") <> "" Then

        Kill "C:\GconTa\sccontr.soc"

    End If


    Open "C:\GconTa\sccontr.soc" For Random As 3 Len = Len(SCont)

    SCont.guarda = RutaSeleccionada

    Put 3, 1, SCont

    Close 3

    Exit Sub


ErrorGuardar:

    Dim NumeroError As Long
    Dim DescripcionError As String

    NumeroError = Err.Number
    DescripcionError = Err.Description

    On Error Resume Next
    Close 3
    On Error GoTo 0

    MsgBox "No fue posible guardar el directorio seleccionado." & _
           vbCrLf & vbCrLf & _
           "Error: " & CStr(NumeroError) & _
           vbCrLf & _
           DescripcionError, _
           vbExclamation, _
           "Directorio"

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim Respuesta As Integer
    Dim RutaSeleccionada As String

    RutaSeleccionada = Trim$(Dir1.Path)


    Respuesta = MsgBox( _
        "¿Desea guardar el directorio seleccionado?" & vbCrLf & vbCrLf & _
        "SÍ = Guardar el directorio y continuar." & vbCrLf & _
        "NO = Cerrar la aplicación sin guardar." & vbCrLf & _
        "CANCELAR = Regresar al selector.", _
        vbYesNoCancel + vbQuestion, _
        "Cambio de directorio")


    Select Case Respuesta


        '==========================================================
        ' SÍ
        '==========================================================
        Case vbYes

            If RutaSeleccionada = "" Then

                MsgBox "No hay un directorio seleccionado.", _
                       vbExclamation, _
                       "Directorio"

                Cancel = True
                Exit Sub

            End If


            ' Comprobar físicamente que la carpeta se puede usar.
            If CambiarADirectorio(RutaSeleccionada) = False Then

                MsgBox "No fue posible acceder al directorio seleccionado." & _
                       vbCrLf & vbCrLf & _
                       RutaSeleccionada, _
                       vbExclamation, _
                       "Directorio no disponible"

                Cancel = True
                Exit Sub

            End If


            GuardarDirectorio

            Exit Sub


        '==========================================================
        ' NO
        '==========================================================
        Case vbNo

            Close
            End


        '==========================================================
        ' CANCELAR
        '==========================================================
        Case vbCancel

            Cancel = True
            Exit Sub

    End Select

End Sub

