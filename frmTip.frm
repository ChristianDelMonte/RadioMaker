VERSION 5.00
Begin VB.Form frmTip 
   Caption         =   "Sugerencia del día"
   ClientHeight    =   3315
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5385
   ForeColor       =   &H8000000F&
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Mostrar sugerencias al iniciar"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2940
      Width           =   3720
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Siguiente sugerencia"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmTip.frx":08CA
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sabía que..."
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1995
         Left            =   135
         TabIndex        =   4
         Top             =   585
         Width           =   3435
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' La base de datos en memoria de sugerencias.
Dim Tips As New Collection

' Nombre del archivo de sugerencias
Const TIP_FILE = "Data\TipOfDay.Dat"

' Índice en la colección de la sugerencia actualmente mostrada.
Dim CurrentTip As Long

Private Sub DoNextTip()

    ' Seleccionar una sugerencia aleatoriamente.
    CurrentTip = Int((Tips.count * Rnd) + 1)
    
    ' O recorrer secuencialmente las sugerencias

'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' Mostrar.
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Leer cada sugerencia desde archivo.
    Dim InFile As Integer   ' Descriptor para archivo.
    
    ' Obtener el siguiente descriptor de archivo libre.
    InFile = FreeFile
    
    ' Asegurarse de que se especifica un archivo.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Asegurarse de que el archivo existe antes de intentar abrirlo.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Leer la colección desde un archivo de texto.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Mostrar una sugerencia aleatoriamente.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub chkLoadTipsAtStartup_Click()
    ' guardar si este formulario debe mostrarse o no al iniciar
    SaveSetting App.EXEName, "Opciones", "Mostrar sugerencias al iniciar", chkLoadTipsAtStartup.value
End Sub

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim ShowAtStartup As Long
    
    ' Ver si debemos mostrar al iniciar
    ShowAtStartup = GetSetting(App.EXEName, "Opciones", "Mostrar sugerencias al iniciar", 1)
    If ShowAtStartup = 0 Then
        Unload Me
        Exit Sub
    End If
        
    ' Establecer la casilla de verificación, que obligará a que el valor se vuelva a escribir en el Registro
    Me.chkLoadTipsAtStartup.value = vbChecked
    
    ' Semilla aleatoria
    Randomize
    
    ' Leer el archivo de sugerencias y mostrar una sugerencia aleatoriamente.
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = "no se ha encontrado el archivo " & TIP_FILE & vbCrLf & vbCrLf & _
           "Cree un archivo de texto llamado " & TIP_FILE & " con el Bloc de notas, con una sugerencia por línea. " & _
           "A continuación, colóquelo en el mismo directorio que la aplicación."
    End If

    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub
