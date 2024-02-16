VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form XPlorer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Explorador de archivos"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8595
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   235
      Left            =   1800
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   6
      Top             =   7680
      Visible         =   0   'False
      Width           =   235
   End
   Begin VB.ComboBox ExCombo 
      DragIcon        =   "Explorer.frx":0000
      Height          =   315
      Left            =   3330
      TabIndex        =   2
      Top             =   3480
      Width           =   3165
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Información de archivo..."
      Height          =   375
      Left            =   90
      TabIndex        =   3
      Top             =   3990
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5370
      TabIndex        =   5
      Top             =   3990
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      DragIcon        =   "Explorer.frx":030A
      Height          =   2235
      Left            =   3345
      TabIndex        =   1
      Top             =   720
      Width           =   3075
   End
   Begin ComctlLib.TreeView tvwDirTree 
      DragIcon        =   "Explorer.frx":0614
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4154
      _Version        =   327682
      Indentation     =   353
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   6450
      X2              =   90
      Y1              =   3885
      Y2              =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   6450
      X2              =   75
      Y1              =   3870
      Y2              =   3870
   End
   Begin ComctlLib.ImageList imgDirTree 
      Left            =   1080
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explorer.frx":091E
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explorer.frx":0A30
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explorer.frx":0B42
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explorer.frx":0C54
            Key             =   "floppy"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explorer.frx":0D66
            Key             =   "hard"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explorer.frx":0E78
            Key             =   "cdrom"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explorer.frx":0F8A
            Key             =   "net"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Explorer.frx":109C
            Key             =   "desk"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPath 
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6570
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Archivos:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3345
      TabIndex        =   4
      Top             =   3165
      Width           =   1815
   End
End
Attribute VB_Name = "XPlorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RResult As String
Dim sMainDrives() As String

Private Sub Form_Load()

Dim k As Integer, sSelectDrive As String
Dim sTemp2 As String, sTemp As String, iSlashSpot As Integer

Dim CmbItem(0 To 5)
Dim i

Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2

ReDim sMainDrives(0)
sSelectDrive = ListDrives(Me, XPlorer.tvwDirTree, XPlorer.imgDirTree, XPlorer.picIcon)

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

'Drive1.Drive = Left$(App.Path, 2)
'Dir1.Path = App.Path
File1.path = lblPath.Caption
File1.Pattern = "*.wav;*.Wav;*.WAV;*.mp3;*.Mp3;*.MP3;*.mp2;*.mp1;*.xm;*.mod;*.s3m;*.it;*.mtm;*.mo3;*.umx;*.ogg;*.tnd"

KeepOnTop XPlorer

End Sub

Private Sub tvwDirTree_Collapse(ByVal Node As ComctlLib.Node)

tvwDirTree_NodeClick Node

End Sub

Private Sub tvwDirTree_Expand(ByVal Node As ComctlLib.Node)

Dim lRet As Long, k As Integer
On Error Resume Next
If Node.Key = "root" Or Node.Key = "desk" Then
    Exit Sub
End If

'check to see if the drive or dir has been read
For k = 0 To UBound(sMainDrives)
    If sMainDrives(k) = Node.Key Then
        'We've already listed the sub dirs once so get out
        Exit Sub
    End If
Next
'check to see if there are sub dirs. If not get out
If Node.Children = 0 Then Exit Sub

Node.Sorted = True

Screen.MousePointer = 13
'true off the redraw for the tree view
lRet = SendMessage(tvwDirTree.Hwnd, WM_SETREDRAW, False, 0&)
'list all of the sub directories
ListDirs XPlorer.tvwDirTree, (Node.Key)
'redraw the control
lRet = SendMessage(tvwDirTree.Hwnd, WM_SETREDRAW, True, 0&)

'iAdd the dir so we don't read it twice
lRet = UBound(sMainDrives)
ReDim Preserve sMainDrives(lRet + 1)
sMainDrives(lRet + 1) = Node.Key
Screen.MousePointer = 0

End Sub

Private Sub Command1_Click()

HideWindow "Explor01"

End Sub

Private Sub ExCombo_Change()

Select Case ExCombo.ListIndex
    Case 0
        File1.path = Right$(lblPath, Len(lblPath) - 2)
        File1.Pattern = "*.wav;*.Wav;*.WAV;*.mp3;*.Mp3;*.MP3;*.mp2;*.mp1;*.xm;*.mod;*.s3m;*.it;*.mtm;*.mo3;*.umx;*.ogg;*.tnd"
    Case 1
        File1.path = Right$(lblPath, Len(lblPath) - 2)
        File1.Pattern = "*.wav;*.Wav;*.WAV"
    Case 2
        File1.path = Right$(lblPath, Len(lblPath) - 2)
        File1.Pattern = "*.mp3;*.Mp3;*.MP3;*.mp2;*.Mp2;*.MP2;*.mp1;*.Mp1;*.MP1;*.ogg"
    Case 3
        File1.path = Right$(lblPath, Len(lblPath) - 2)
        File1.Pattern = "*.xm;*.mod;*.s3m;*.it;*.mtm;*.mo3;*.umx"
    Case 4
        File1.path = Right$(lblPath, Len(lblPath) - 2)
        File1.Pattern = "*.tnd"
    Case 5
        File1.path = Right$(lblPath, Len(lblPath) - 2)
        File1.Pattern = "*.*"
End Select

End Sub

Private Sub ExCombo_Click()

On Error Resume Next
Select Case ExCombo.ListIndex
    Case 0
        File1.path = Right$(lblPath, Len(lblPath) - 2)
        File1.Pattern = "*.wav;*.Wav;*.WAV;*.mp3;*.Mp3;*.MP3;*.mp2;*.mp1;*.xm;*.mod;*.s3m;*.it;*.mtm;*.mo3;*.umx;*.ogg;*.tnd"
    Case 1
        File1.path = Right$(lblPath, Len(lblPath) - 2)
        File1.Pattern = "*.wav;*.Wav;*.WAV"
    Case 2
        File1.path = Right$(lblPath, Len(lblPath) - 2)
        File1.Pattern = "*.mp3;*.Mp3;*.MP3;*.mp2;*.Mp2;*.MP2;*.mp1;*.Mp1;*.MP1;*.ogg"
    Case 3
        File1.path = Right$(lblPath, Len(lblPath) - 2)
        File1.Pattern = "*.xm;*.mod;*.s3m;*.it;*.mtm;*.mo3;*.umx"
    Case 4
        File1.path = Right$(lblPath, Len(lblPath) - 2)
        File1.Pattern = "*.tnd"
    Case 5
        File1.path = Right$(lblPath, Len(lblPath) - 2)
        File1.Pattern = "*.*"
End Select

End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

File1.DragIcon = tvwDirTree.DragIcon
File1.Drag

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Set XPlorer = Nothing
HideWindow "Explor01"

End Sub

Private Sub Form_Terminate()

HideWindow "Explor01"

End Sub

Private Sub Form_Unload(Cancel As Integer)

HideWindow "Explor01"

End Sub

Private Sub tvwDirTree_NodeClick(ByVal Node As ComctlLib.Node)

On Error Resume Next
'if it is a floppy then recheck
If Node.Key = "A:" Or Node.Key = "B:" Then
    Screen.MousePointer = 13
    ListDirs XPlorer.tvwDirTree, (Node.Key)
    Screen.MousePointer = 0
End If

If Node.Key = "root" Or Node.Key = "desk" Then
    lblPath = ""
    File1.path = ""
    Exit Sub
End If

If TextWidth(Node.Key) > lblPath.Width Then
    lblPath = Node.Key
    File1.path = Right$(lblPath, Len(lblPath) - 2)
Else
    ' if not put it at the bottom
    lblPath = vbCrLf & Node.Key
    File1.path = Right$(lblPath, Len(lblPath) - 2)
End If

End Sub


