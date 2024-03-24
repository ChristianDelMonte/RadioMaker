VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form XPlorer 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Explorador de archivos"
   ClientHeight    =   10290
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   5790
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Iformación de Archivo:"
      Height          =   1995
      Left            =   150
      TabIndex        =   9
      Top             =   7710
      Width           =   4545
      Begin VB.CommandButton Cmdplus 
         Caption         =   "+"
         Height          =   255
         Left            =   4260
         TabIndex        =   22
         Top             =   0
         Width           =   285
      End
      Begin VB.Label LblGenero 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   21
         Top             =   1710
         Width           =   3255
      End
      Begin VB.Label Lblano 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   20
         Top             =   1470
         Width           =   3255
      End
      Begin VB.Label LblComentario 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   19
         Top             =   1020
         Width           =   3255
      End
      Begin VB.Label LblAlbum 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   18
         Top             =   780
         Width           =   3255
      End
      Begin VB.Label LblArtista 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         Top             =   540
         Width           =   3255
      End
      Begin VB.Label LblTitulo 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   16
         Top             =   300
         Width           =   3255
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Genero:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1710
         Width           =   1000
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Año:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1470
         Width           =   1000
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Comentario:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1020
         Width           =   1000
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Album:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   1000
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Artista:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   540
         Width           =   1000
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Título:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   1000
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   150
      TabIndex        =   7
      Top             =   450
      Width           =   4545
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   150
      TabIndex        =   6
      Top             =   870
      Width           =   4545
   End
   Begin MSComctlLib.TreeView tvwDirTree 
      Height          =   1455
      Left            =   5970
      TabIndex        =   5
      Top             =   5910
      Visible         =   0   'False
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   2566
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   7
      Appearance      =   1
   End
   Begin VB.ComboBox ExCombo 
      DragIcon        =   "Explorer.frx":0000
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   7260
      Width           =   4545
   End
   Begin VB.CommandButton Cmdclose 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   9810
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      DragIcon        =   "Explorer.frx":030A
      Height          =   3600
      Left            =   150
      TabIndex        =   0
      Top             =   3360
      Width           =   4545
   End
   Begin RM100.TitelBar TitelBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   661
      BackColor       =   8421504
      BackColorCover  =   3
      BackColorV2Begin=   4210752
      BackColorV2End  =   0
      BackColorV1Begin=   4210752
      BackColorV1End  =   0
      ForeColor       =   16777215
      ShowMinimized   =   0   'False
      ShowMaximizedEnabled=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " Explorador de Archivos"
      CaptionPosX     =   1
      BorderNormal    =   2
      BorderColorHighLight=   0
      BorderColorDarkLight=   4210752
   End
   Begin VB.Label lblPath 
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5970
      TabIndex        =   4
      Top             =   5400
      Width           =   2940
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Archivos:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   7020
      Width           =   1815
   End
End
Attribute VB_Name = "XPlorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmdclose_Click()

Unload Me

End Sub

Private Sub Dir1_Change()

File1.path = Dir1.path

End Sub

Private Sub Drive1_Change()

On Error Resume Next
Dir1.path = Drive1.Drive

End Sub

Private Sub File1_Click()

Dim Completo As String

lblPath.Caption = File1.path

'extraemos la informacion TAG del archivo si existe

Completo = Trim(lblPath.Caption) & "\" & File1.filename

If GetMP3Tag(Completo) = True Then
    LblTitulo.Caption = Replace(Trim(MP3Info.sTitle), Chr(0), "")
    LblArtista.Caption = Replace(Trim(MP3Info.sArtist), Chr(0), "")
    LblAlbum.Caption = Replace(Trim(MP3Info.sAlbum), Chr(0), "")
    LblComentario.Caption = Replace(Trim(MP3Info.sComment), Chr(0), "")
    Lblano.Caption = Replace(Trim(MP3Info.sYear), Chr(0), "")
    LblGenero.Caption = Replace(Trim(MP3Info.sGenre), Chr(0), "")
Else
    LblTitulo.Caption = "sin información"
    LblArtista.Caption = "sin información"
    LblAlbum.Caption = "sin información"
    LblComentario.Caption = "sin información"
    Lblano.Caption = "sin información"
    LblGenero.Caption = "sin información"
End If

End Sub

Private Sub Form_Load()

'*** load some pictures *****
Me.Picture = LoadPicture(App.path & "\Imagenes\FND_COMPLETO.jpg")

Dim CmbItem(0 To 5)
Dim i

CmbItem(0) = "Todos los Archivos (*.wav;*.mp3;*.mp2;*.mp1;*.xm;*.mod;*.s3m;*.it;*.mtm;*.mo3;*.umx;*.ogg;*.tnd)"
CmbItem(1) = "Audio de Ondas (*.wav)"
CmbItem(2) = "Audio de Ondas Comprimido (*.mp3;*.mp2;*.mp1;*.ogg)"
CmbItem(3) = "Archivos de Modulo (*.xm;*.mod;*.s3m;*.it;*.mtm;*.mo3;*.umx)"
CmbItem(4) = "Archivos de Tandas (*.tnd)"
CmbItem(5) = "Todos los Archivos (*.*)"

For i = 0 To 5
    ExCombo.AddItem CmbItem(i)
Next i

ExCombo.ListIndex = 0

Dir1.path = App.path
Drive1.Drive = Left$(App.path, 2)

File1.path = Dir1.path
File1.Pattern = "*.wav;*.Wav;*.WAV;*.mp3;*.Mp3;*.MP3;*.mp2;*.mp1;*.xm;*.mod;*.s3m;*.it;*.mtm;*.mo3;*.umx;*.ogg;*.tnd"

'KeepOnTop XPlorer

End Sub

Private Sub ExCombo_Change()

Select Case ExCombo.ListIndex
    Case 0
        File1.path = Dir1.path
        File1.Pattern = "*.wav;*.Wav;*.WAV;*.mp3;*.Mp3;*.MP3;*.mp2;*.mp1;*.xm;*.mod;*.s3m;*.it;*.mtm;*.mo3;*.umx;*.ogg;*.tnd"
    Case 1
        File1.path = Dir1.path
        File1.Pattern = "*.wav;*.Wav;*.WAV"
    Case 2
        File1.path = Dir1.path
        File1.Pattern = "*.mp3;*.Mp3;*.MP3;*.mp2;*.Mp2;*.MP2;*.mp1;*.Mp1;*.MP1;*.ogg"
    Case 3
        File1.path = Dir1.path
        File1.Pattern = "*.xm;*.mod;*.s3m;*.it;*.mtm;*.mo3;*.umx"
    Case 4
        File1.path = Dir1.path
        File1.Pattern = "*.tnd"
    Case 5
        File1.path = Dir1.path
        File1.Pattern = "*.*"
End Select

End Sub

Private Sub ExCombo_Click()

On Error Resume Next
Select Case ExCombo.ListIndex
    Case 0
        File1.path = Dir1.path
        File1.Pattern = "*.wav;*.Wav;*.WAV;*.mp3;*.Mp3;*.MP3;*.mp2;*.mp1;*.xm;*.mod;*.s3m;*.it;*.mtm;*.mo3;*.umx;*.ogg;*.tnd"
    Case 1
        File1.path = Dir1.path
        File1.Pattern = "*.wav;*.Wav;*.WAV"
    Case 2
        File1.path = Dir1.path
        File1.Pattern = "*.mp3;*.Mp3;*.MP3;*.mp2;*.Mp2;*.MP2;*.mp1;*.Mp1;*.MP1;*.ogg"
    Case 3
        File1.path = Dir1.path
        File1.Pattern = "*.xm;*.mod;*.s3m;*.it;*.mtm;*.mo3;*.umx"
    Case 4
        File1.path = Dir1.path
        File1.Pattern = "*.tnd"
    Case 5
        File1.path = Dir1.path
        File1.Pattern = "*.*"
End Select

End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

File1.DragIcon = tvwDirTree.DragIcon
File1.Drag

End Sub

