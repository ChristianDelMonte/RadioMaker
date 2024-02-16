VERSION 5.00
Begin VB.Form frmtest 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   4485
      ScaleHeight     =   240
      ScaleWidth      =   1230
      TabIndex        =   28
      Top             =   1380
      Width           =   1290
      Begin VB.PictureBox tr6 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   990
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   34
         ToolTipText     =   "Tiempo transcurrido"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox tr1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   45
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   33
         ToolTipText     =   "Tiempo transcurrido"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox tr2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   240
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   32
         ToolTipText     =   "Tiempo transcurrido"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox tr3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   420
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   31
         ToolTipText     =   "Tiempo transcurrido"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox tr4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   615
         ScaleHeight     =   180
         ScaleWidth      =   195
         TabIndex        =   30
         ToolTipText     =   "Tiempo transcurrido"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox tr5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   810
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   29
         ToolTipText     =   "Tiempo transcurrido"
         Top             =   15
         Width           =   190
      End
   End
   Begin VB.Timer TmrClock 
      Left            =   1185
      Top             =   5265
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   1695
      ScaleHeight     =   240
      ScaleWidth      =   1575
      TabIndex        =   19
      Top             =   1965
      Width           =   1635
      Begin VB.PictureBox Cp5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   780
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   27
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   585
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   26
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   390
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   25
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   210
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   24
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   15
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   23
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp6 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   960
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   22
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp7 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1155
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   21
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Cp8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1350
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   20
         ToolTipText     =   "Hora actual"
         Top             =   15
         Width           =   190
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   3705
      ScaleHeight     =   240
      ScaleWidth      =   1575
      TabIndex        =   10
      Top             =   1965
      Width           =   1635
      Begin VB.PictureBox Tp8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1350
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   18
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp7 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1155
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   17
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp6 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   960
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   16
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   15
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   15
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   210
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   14
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   390
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   13
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   585
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   12
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
      Begin VB.PictureBox Tp5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   780
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   11
         ToolTipText     =   "Hora de lanzamiento del próximo tema"
         Top             =   15
         Width           =   190
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00000000&
      Height          =   315
      Left            =   3735
      ScaleHeight     =   255
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   795
      Width           =   2055
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   15
         TabIndex        =   9
         ToolTipText     =   "Nombre de dispositivo en funcionamiento"
         Top             =   15
         Width           =   1965
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   1710
      ScaleHeight     =   240
      ScaleWidth      =   2535
      TabIndex        =   6
      Top             =   1380
      Width           =   2595
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   15
         TabIndex        =   7
         ToolTipText     =   "Nombre del tema actualmente en reproducción"
         Top             =   15
         Width           =   2475
      End
   End
   Begin VB.Timer TmrPos1 
      Left            =   675
      Top             =   5280
   End
   Begin VB.Timer TmrScope 
      Left            =   210
      Top             =   5235
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   165
      Left            =   1635
      ScaleHeight     =   105
      ScaleWidth      =   3060
      TabIndex        =   4
      Top             =   3345
      Width           =   3120
      Begin VB.PictureBox Ll 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   0
         ScaleHeight     =   150
         ScaleMode       =   0  'User
         ScaleWidth      =   105
         TabIndex        =   5
         Top             =   0
         Width           =   100
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   165
      Left            =   1635
      ScaleHeight     =   105
      ScaleMode       =   0  'User
      ScaleWidth      =   3060
      TabIndex        =   2
      Top             =   3495
      Width           =   3120
      Begin VB.PictureBox Lr 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   0
         ScaleHeight     =   150
         ScaleWidth      =   105
         TabIndex        =   3
         Top             =   0
         Width           =   100
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "V"
      Height          =   300
      Left            =   5460
      TabIndex        =   1
      ToolTipText     =   "Restaurar Radio Maker"
      Top             =   1980
      Width           =   330
   End
   Begin VB.PictureBox picMainSkin 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2490
      Left            =   0
      ScaleHeight     =   2490
      ScaleWidth      =   6165
      TabIndex        =   0
      Top             =   0
      Width           =   6165
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00000000&
         Height          =   840
         Left            =   1680
         ScaleHeight     =   52
         ScaleMode       =   0  'User
         ScaleWidth      =   101
         TabIndex        =   36
         Top             =   240
         Width           =   1575
         Begin VB.PictureBox Picfft1 
            BackColor       =   &H00404000&
            BorderStyle     =   0  'None
            Height          =   690
            Left            =   30
            ScaleHeight     =   46
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   96
            TabIndex        =   37
            Top             =   45
            Width           =   1440
         End
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         Height          =   180
         Left            =   3435
         TabIndex        =   35
         Top             =   2025
         Width           =   225
      End
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub SetBuffLevel(WLeft, WRight)

'SUB para visualizar un gráfico de 15 cuadros
'al compas de la musica.

Dim l, Lft As Integer
Dim r, Rgt As Integer

Lft = WLeft / 9
Rgt = WRight / 9

If Lft >= 14 Then
    Lft = 14
    Pic1.Picture = Clip1.GraphicCell(Lft)
Else
    Pic1.Picture = Clip1.GraphicCell(Lft)
End If

End Sub

Sub GetNextStreamTime()

'cual es el proximo tema a reproducir?
'only for Tanda streams

Dim NIndex As Integer
Dim NextLanz As String

On Error GoTo Err
NIndex = Tanda01.T1View.SelectedItem.Index   'numero de index
NIndex = NIndex + 1
Tanda01.T1View.ListItems.Item(NIndex).Selected = True

'extraemos los datos del del tema
NextLanz = Tanda01.T1View.SelectedItem.ListSubItems(4).Text     'hora de lanz

'volvemos a la posicion donde se encontraba el cursor
NIndex = NIndex - 1
Tanda01.T1View.ListItems.Item(NIndex).Selected = True

'actualizamos los datos de los demas temas
Call SetNextTime(Trim(NextLanz))
Exit Sub

Err:
'xxxxxxx
'no hay proximo tema, deshabilitar el display
Call RestoreDisplay(10)

End Sub


Public Sub SetAudioLevel(WLeft, WRight)

'level scope meter sub

Dim l, Lft As Integer
Dim r, Rgt As Integer
Dim i As Integer
Static ZMax%, RMax%

On Error Resume Next
'right level meter
If WRight > 180 Then
    RMax = (WRight * 24) + 100 'clip
Else
    RMax = (WRight * 24)
End If

'left level meter
If WLeft > 180 Then
    ZMax = (WLeft * 24) + 100  'clip
Else
    ZMax = (WLeft * 24)
End If

Lr.Width = RMax
Ll.Width = ZMax

End Sub


Private Sub Command1_Click()

TmrScope.Enabled = False
TmrPos1.Enabled = False

TopMenu.WindowState = 0

End Sub

Private Sub Form_Load()
    
    Dim WindowRegion As Long
    
    'load led1
    Picture1.Picture = LoadResPicture("BACK_LED", 0)
    Ll.Picture = LoadResPicture("FRONT_LED", 0)
    Ll.Width = 1
    'load led2
    Picture2.Picture = LoadResPicture("BACK_LED", 0)
    Lr.Picture = LoadResPicture("FRONT_LED", 0)
    Lr.Width = 1
    'paint clocks
    Call RestoreDisplay(10)
    'tiempo restante
    tr1.Picture = TopMenu.SmallClip.GraphicCell(10)
    tr2.Picture = TopMenu.SmallClip.GraphicCell(10)
    tr3.Picture = TopMenu.SmallClip.GraphicCell(10)
    tr4.Picture = TopMenu.SmallClip.GraphicCell(10)
    tr5.Picture = TopMenu.SmallClip.GraphicCell(10)
    tr6.Picture = TopMenu.SmallClip.GraphicCell(10)

    ' I set all these settings here so you won't forget
    ' them and have a non-working demo... Set them in
    ' design time
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
    Set picMainSkin.Picture = LoadResPicture("RM_MIN_SPC", 0)
    
    Me.Width = picMainSkin.Width
    Me.Height = picMainSkin.Height
    
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hWnd, WindowRegion, True
    
    'scope timer
    TmrScope.Enabled = True
    TmrScope.Interval = 10
    'stream position timer
    TmrPos1.Enabled = True
    TmrPos1.Interval = 10
    'real clock timer
    TmrClock.Enabled = True
    TmrClock.Interval = 10
    
    Me.Left = Screen.Width - Me.Width
    Me.top = (Screen.Height / 2) - (Me.Height / 2)
    
End Sub

Private Sub Label2_Change()

If Label2.Caption = "Tanda - Dev: 1" Then
    Call GetNextStreamTime
Else
    If Label2.Caption = "Tanda - Dev: 2" Then
        Call GetNextStreamTime
    Else
        Call RestoreDisplay(10)
    End If
End If

End Sub

Private Sub picMainSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
      ' Pass the handling of the mouse down message to
      ' the (non-existing really) form caption, so that
      ' the form itself will be dragged when the picture is dragged.
      '
      ' If you have Win 98, Make sure that the "Show window
      ' contents while dragging" display setting is on for nice results.
      
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub


Private Sub TmrClock_Timer()

Dim RealClock As String

RealClock = Time$

'update the clock
Call MinClock(RealClock)

End Sub

Private Sub TmrScope_Timer()

Dim LLft
Dim RRgt
Dim SType As String
Dim Result As Boolean
Dim Numstr As Long

If Stream01IsPlaying = True Then
    If Est12Control.Origen1.Caption = "E1" Or Est12Control.Origen1.Caption = "T1" Then
        'LLft = Stream01GetLEFTLevel
        'RRgt = Stream01GetRIGHTLevel
        'frmtest.SetAudioLevel LLft, RRgt
        'frmtest.SetBuffLevel LLft, RRgt
        Numstr = 1
        SType = "Stream"
        Call DrawMiniFFT(Numstr, SType, 6) 'fft spectrum display
    Else
        'xxx
    End If
Else
    If Stream02IsPlaying = True Then
        If Est12Control.Origen2.Caption = "E2" Or Est12Control.Origen2.Caption = "T2" Then
            'LLft = Stream02GetLEFTLevel
            'RRgt = Stream02GetRIGHTLevel
            'frmtest.SetAudioLevel LLft, RRgt
            'frmtest.SetBuffLevel LLft, RRgt
            Numstr = 2
            SType = "Stream"
            Call DrawMiniFFT(Numstr, SType, 6) 'fft spectrum display
        Else
            'xxx
        End If
    End If
End If

End Sub
