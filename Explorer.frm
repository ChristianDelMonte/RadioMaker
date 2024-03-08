VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form Old_XPlorer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Explorador de archivos"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9930
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgDirTree 
      Left            =   1110
      Top             =   7470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":0000
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":0112
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":0224
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":0336
            Key             =   "floppy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":0448
            Key             =   "hard"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":055A
            Key             =   "cdroom"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":066C
            Key             =   "net"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":077E
            Key             =   "desk"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwDirTree 
      Height          =   4785
      Left            =   210
      TabIndex        =   7
      Top             =   720
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   8440
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   7
      Appearance      =   1
   End
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
      TabIndex        =   5
      Top             =   7680
      Visible         =   0   'False
      Width           =   235
   End
   Begin VB.ComboBox ExCombo 
      DragIcon        =   "Explorer.frx":0890
      Height          =   315
      Left            =   4470
      TabIndex        =   1
      Top             =   6030
      Width           =   3165
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Información de archivo..."
      Height          =   375
      Left            =   1230
      TabIndex        =   2
      Top             =   6540
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   6510
      TabIndex        =   4
      Top             =   6540
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      DragIcon        =   "Explorer.frx":0B9A
      Height          =   4770
      Left            =   4710
      TabIndex        =   0
      Top             =   720
      Width           =   5175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   7590
      X2              =   1230
      Y1              =   6435
      Y2              =   6435
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   7590
      X2              =   1215
      Y1              =   6420
      Y2              =   6420
   End
   Begin VB.Label lblPath 
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6570
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Archivos:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4485
      TabIndex        =   3
      Top             =   5715
      Width           =   1815
   End
End
Attribute VB_Name = "Old_XPlorer"
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

Private Sub tvwDirTree_Collapse(ByVal Node As MSComctlLib.Node)

tvwDirTree_NodeClick Node

End Sub

Private Sub tvwDirTree_Expand(ByVal Node As MSComctlLib.Node)

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
lRet = SendMessage(tvwDirTree.hWnd, WM_SETREDRAW, False, 0&)
'list all of the sub directories
ListDirs XPlorer.tvwDirTree, (Node.Key)
'redraw the control
lRet = SendMessage(tvwDirTree.hWnd, WM_SETREDRAW, True, 0&)

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

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

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

Private Sub tvwDirTree_NodeClick(ByVal Node As MSComctlLib.Node)

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


