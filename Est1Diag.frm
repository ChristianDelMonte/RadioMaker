VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Est1Diag 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ESTACION 01"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6795
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   1860
      Index           =   10
      Left            =   3735
      TabIndex        =   50
      Top             =   6300
      Width           =   285
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1860
      Index           =   9
      Left            =   3375
      TabIndex        =   49
      Top             =   6300
      Width           =   240
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1860
      Index           =   8
      Left            =   3015
      TabIndex        =   48
      Top             =   6300
      Width           =   240
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1860
      Index           =   7
      Left            =   2655
      TabIndex        =   47
      Top             =   6300
      Width           =   240
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1860
      Index           =   6
      Left            =   2295
      TabIndex        =   46
      Top             =   6300
      Width           =   240
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1860
      Index           =   5
      Left            =   1935
      TabIndex        =   45
      Top             =   6300
      Width           =   240
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1860
      Index           =   4
      Left            =   1575
      TabIndex        =   44
      Top             =   6300
      Width           =   240
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1860
      Index           =   3
      Left            =   1215
      TabIndex        =   43
      Top             =   6300
      Width           =   240
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1860
      Index           =   2
      Left            =   855
      TabIndex        =   42
      Top             =   6300
      Width           =   240
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1860
      Index           =   1
      Left            =   495
      TabIndex        =   41
      Top             =   6300
      Width           =   240
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1860
      Index           =   0
      Left            =   135
      TabIndex        =   40
      Top             =   6300
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1365
      Left            =   90
      Picture         =   "Est1Diag.frx":0000
      ScaleHeight     =   1305
      ScaleWidth      =   3780
      TabIndex        =   38
      Top             =   4815
      Width           =   3840
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   315
         Picture         =   "Est1Diag.frx":1311
         ScaleHeight     =   1020
         ScaleWidth      =   120
         TabIndex        =   39
         Top             =   180
         Width           =   120
      End
   End
   Begin VB.CommandButton E1Save 
      Height          =   375
      Left            =   4680
      Picture         =   "Est1Diag.frx":292D
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Guardar archivo CUE"
      Top             =   3660
      Width           =   495
   End
   Begin VB.CommandButton E1Open 
      Height          =   375
      Left            =   4080
      Picture         =   "Est1Diag.frx":2A2F
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Abrir archivo CUE"
      Top             =   3660
      Width           =   495
   End
   Begin VB.CommandButton E1Import 
      Caption         =   "&Importar"
      Height          =   360
      Left            =   5370
      TabIndex        =   36
      ToolTipText     =   "Importar archivo CUE"
      Top             =   3660
      Width           =   915
   End
   Begin VB.CommandButton E1Stop 
      Height          =   375
      Left            =   2340
      Picture         =   "Est1Diag.frx":2B31
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Detener"
      Top             =   3660
      Width           =   855
   End
   Begin VB.CommandButton E1Pause 
      Height          =   375
      Left            =   1500
      Picture         =   "Est1Diag.frx":356D
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Pausar"
      Top             =   3660
      Width           =   855
   End
   Begin VB.CommandButton E1Play 
      Height          =   375
      Left            =   660
      Picture         =   "Est1Diag.frx":3FA9
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Reproducir"
      Top             =   3660
      Width           =   855
   End
   Begin VB.CommandButton E1Prev 
      Height          =   375
      Left            =   180
      Picture         =   "Est1Diag.frx":49E5
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Detener"
      Top             =   3660
      Width           =   495
   End
   Begin VB.CommandButton E1New 
      Height          =   375
      Left            =   3480
      Picture         =   "Est1Diag.frx":5421
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Nuevo archivo CUE"
      Top             =   3660
      Width           =   495
   End
   Begin VB.Frame Frame4 
      Caption         =   "CUE"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1620
      TabIndex        =   22
      Top             =   2475
      Width           =   3195
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "00:00"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "00:00"
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<M"
         Height          =   300
         Left            =   1140
         TabIndex        =   24
         ToolTipText     =   "Marcar la posicion de inicio de CUE"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "<M"
         Height          =   300
         Left            =   2700
         TabIndex        =   23
         ToolTipText     =   "Marcar la posicion de fin de CUE"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CUE Inicio"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CUE Final"
         Height          =   255
         Left            =   1680
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Posicionamiento "
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   6255
      Begin MSComctlLib.Slider E1Pos 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   10
         Max             =   100
         SelectRange     =   -1  'True
      End
      Begin VB.Label Label9 
         Caption         =   "En proceso:"
         Height          =   255
         Left            =   210
         TabIndex        =   18
         Top             =   630
         Width           =   915
      End
      Begin VB.Label LblCurrent 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "00:00,00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1140
         TabIndex        =   17
         Top             =   585
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Finalización:"
         Height          =   255
         Left            =   3990
         TabIndex        =   16
         Top             =   630
         Width           =   915
      End
      Begin VB.Label LblEnd 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "00:00,00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4905
         TabIndex        =   15
         Top             =   570
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Paneo"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   765
      Width           =   6255
      Begin VB.CommandButton CmdAutoPan 
         Caption         =   "AUTO PANEO"
         Height          =   255
         Left            =   4920
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin MSComctlLib.Slider E1Slide 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   10
         Min             =   -100
         Max             =   100
         TickFrequency   =   10
      End
      Begin VB.Label Label2 
         Caption         =   "Iz"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "Dr"
         Height          =   255
         Left            =   4500
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Volumen"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton CmdFOut 
         Caption         =   "FADE OUT"
         Height          =   255
         Left            =   5160
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdFIN 
         Caption         =   "FADE IN"
         Height          =   255
         Left            =   4200
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin MSComctlLib.Slider E1Vol 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   10
         Max             =   100
         SelStart        =   100
         TickFrequency   =   3
         Value           =   100
      End
   End
   Begin VB.CommandButton CmdRestore 
      Caption         =   "&Restablecer"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Timer TmrCUE 
      Left            =   3600
      Top             =   8580
   End
   Begin MSComDlg.CommonDialog CmdImport 
      Left            =   4140
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Timer TmrUpdt 
      Left            =   3180
      Top             =   8580
   End
   Begin VB.CommandButton E1Cue 
      Caption         =   "Activar CUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Activar / Desactivar CUE"
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   ">"
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Ocultar propiedades de audio"
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   240
      Left            =   180
      TabIndex        =   37
      Top             =   2745
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000018&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4455
      Left            =   6480
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   120
      Top             =   3600
      Width           =   6255
   End
   Begin VB.Label LblCurrByte 
      Caption         =   "0"
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   8805
      Width           =   1095
   End
   Begin VB.Label LblEndCue 
      Caption         =   "0"
      Height          =   255
      Left            =   1500
      TabIndex        =   5
      Top             =   8565
      Width           =   1215
   End
   Begin VB.Label LblStartCUE 
      Caption         =   "0"
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   8565
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6360
      Y1              =   4200
      Y2              =   4200
   End
End
Attribute VB_Name = "Est1Diag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RResult As String

Sub UpdatePos()

Dim ByteLen As String
Dim TimeLen As String
Dim FTime As String
Dim Convt1 As Long

If Est12Control.StopLabel1.Caption = "Stream" Then
    TimeLen = Stream01GetLen(1) 'get len of file in time=1
    FTime = GetSegsFromTime(TimeLen) 'formateamos el tiempo
    E1Pos.min = 0
    If FTime = 0 Then
        E1Pos.max = FTime + 1
    Else
        E1Pos.max = FTime
    End If
    If FTime <= 100 Then
        E1Pos.TickFrequency = 1
    Else
        If FTime > 100 And FTime < 200 Then
            E1Pos.TickFrequency = 2
        Else
            If FTime > 200 And FTime < 300 Then
                E1Pos.TickFrequency = 3
            Else
                If FTime > 300 And FTime < 400 Then
                    E1Pos.TickFrequency = 4
                Else
                    E1Pos.TickFrequency = 5
                End If
            End If
        End If
    End If
    E1Pos.value = 0
    E1Vol.value = 100
    E1Slide.value = 0
    LblEnd.Caption = GetTimeFromSegs(FTime)
    Est1Diag.Caption = "ESTACION 01 - " & Est01.Label1.Caption
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        'ByteLen = Music01GetLen(2)
        'Convt1 = CLng(ByteLen)
        'Convt1 = Convt1
        'E1Pos.Min = 0
        'If Convt1 = 0 Then
        '    E1Pos.max = Convt1 + 1
        'Else
        '    E1Pos.max = Convt1
        'End If
        'E1Pos.TickFrequency = 1
        'E1Pos.Value = 0
        'E1Pos.ToolTipText = Str$(E1Pos.Value)
        'E1Vol.Value = 100
        'E1Slide.Value = 0
        'LblEnd.Caption = Convt1
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub CmdAutoPan_Click()

Dim PanOrigen As Long
Dim PanRight As Long
Dim PanLeft As Long
Dim ActualPan As Long

PanOrigen = 0
PanLeft = -100
PanRight = 100
ActualPan = E1Slide.value

While ActualPan < PanRight
    ActualPan = E1Slide.value + 5   'de o a 100
    E1Slide.value = ActualPan
Wend
While ActualPan > PanOrigen
    ActualPan = E1Slide.value - 5   'de 100 a 0
    E1Slide.value = ActualPan
Wend
While ActualPan > PanLeft
    ActualPan = E1Slide.value - 5   'de 0 a -100
    E1Slide.value = ActualPan
Wend
While ActualPan < PanOrigen
    ActualPan = E1Slide.value + 5   'de -100 a 0
    E1Slide.value = ActualPan
Wend

End Sub

Sub CmdFIN_Click()

'FADE IN CLICK

Dim VolOrigen As Long
Dim VolDestino As Long
Dim ActualVol As Long

E1Vol.value = 0
VolOrigen = 0
VolDestino = 100
ActualVol = E1Vol.value

While ActualVol < VolDestino
    E1Vol.value = E1Vol.value + 5
    ActualVol = E1Vol.value
Wend

End Sub

Sub CmdFOut_Click()

'FADE OUT CLICK

Dim VolOrigen As Long
Dim VolDestino As Long
Dim ActualVol As Long

E1Vol.value = 100
VolOrigen = E1Vol.value
VolDestino = 0
ActualVol = E1Vol.value

While ActualVol > VolDestino
    E1Vol.value = E1Vol.value - 5
    ActualVol = E1Vol.value
Wend

End Sub

Sub CmdRestore_Click()

E1Vol.value = 100
E1Slide.value = 0
E1Pos.value = 0
E1Pos.SelStart = 0
E1Pos.SelLength = 0
Text1.Text = "00:00"
Text2.Text = "00:00"
LblStartCUE.Caption = 0
LblEndCue.Caption = 0

End Sub

Private Sub Command1_Click()

TmrUpdt.Interval = 0
TmrUpdt.Enabled = False
TmrCUE.Interval = 0
TmrCUE.Enabled = False

Unload Me
'RResult = HideWindow("Diag01")

End Sub

Private Sub Command3_Click()

TmrUpdt.Interval = 0
TmrUpdt.Enabled = False

UpdatePos

TmrUpdt.Enabled = True
TmrUpdt.Interval = 100

End Sub

Private Sub Command4_Click()

Text1.Text = LblCurrent.Caption
LblStartCUE.Caption = LblCurrByte.Caption
Command3.Enabled = False
Command4.Enabled = False
CmdRestore.Enabled = False
E1Prev.Enabled = False
E1Play.Enabled = False
E1Pause.Enabled = False
E1Stop.Enabled = False
E1New.Enabled = False
E1Open.Enabled = False
E1Save.Enabled = False
E1Import.Enabled = False
E1Cue.Enabled = False
E1Pos.SelStart = E1Pos.value
End Sub

Private Sub Command5_Click()

Text2.Text = LblCurrent.Caption
LblEndCue.Caption = LblCurrByte.Caption
Command3.Enabled = True
Command4.Enabled = True
CmdRestore.Enabled = True
E1Prev.Enabled = True
E1Play.Enabled = True
E1Pause.Enabled = True
E1Stop.Enabled = True
E1New.Enabled = True
E1Open.Enabled = True
E1Save.Enabled = True
E1Import.Enabled = True
E1Cue.Enabled = True
E1Pos.SelLength = E1Pos.value - E1Pos.SelStart
Call E1Cue_Click

End Sub

Sub E1Cue_Click()

If E1Cue.Caption = "Activar CUE" Then
    E1Cue.Caption = "Desactivar CUE"
    E1Cue.BackColor = &HFFFF&   'amarillo
    Text1.Enabled = False
    Text2.Enabled = False
    Command3.Enabled = False
    Command5.Enabled = False
    Command4.Enabled = False
    CmdRestore.Enabled = False
    E1Prev.Enabled = False
    E1Play.Enabled = False
    E1Pause.Enabled = False
    E1Stop.Enabled = False
    E1New.Enabled = False
    E1Open.Enabled = False
    E1Save.Enabled = False
    E1Import.Enabled = False
    TmrCUE.Enabled = True
    TmrCUE.Interval = 1
Else
    E1Cue.Caption = "Activar CUE"
    E1Cue.BackColor = &H8000000F    'gris
    Text1.Enabled = True
    Text2.Enabled = True
    Command3.Enabled = True
    Command5.Enabled = True
    Command4.Enabled = True
    CmdRestore.Enabled = True
    E1Prev.Enabled = True
    E1Play.Enabled = True
    E1Pause.Enabled = True
    E1Stop.Enabled = True
    E1New.Enabled = True
    E1Open.Enabled = True
    E1Save.Enabled = True
    E1Import.Enabled = True
    TmrCUE.Interval = 0
    TmrCUE.Enabled = False
End If

End Sub

Private Sub E1New_Click()

TmrUpdt.Interval = 0
TmrUpdt.Enabled = False

UpdatePos

TmrUpdt.Enabled = True
TmrUpdt.Interval = 100

CmdRestore_Click

End Sub

Private Sub E1Open_Click()

TmrUpdt.Interval = 0
TmrUpdt.Enabled = False

UpdatePos

TmrUpdt.Enabled = True
TmrUpdt.Interval = 100

Dim ContNum As Integer
Dim FileName As String
Dim LenFN As Long
Dim FileNTest As String
Dim NameFile As String

ContNum = CInt(Est01.Fi.Caption)    'extraemos el index del control
FileName = Trim(Est12Data.N1(ContNum).Caption)    'extraemos el path y el archivo de audio

LenFN = Len(FileName)
FileNTest = Mid$(FileName, LenFN - 2, 1)
If FileNTest = "." Then
    LenFN = LenFN - 3
Else
    LenFN = LenFN - 4
End If

NameFile = Left$(FileName, LenFN)
FileName = Trim(NameFile) & AppCUEFileExt

'guardamos la informacion CUe
OpenCUEFile 1, FileName

End Sub

Private Sub E1Pause_Click()

If Est12Control.StopLabel1.Caption = "Stream" Then
    Stream01Stop   'stream stop
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        Music01Stop    'music stop
    End If
End If

Est01.Caption = "ESTACION 01 - Pausado"
Est12Control.TmrPos1.Interval = 0
Est12Control.TmrPos1.Enabled = False

End Sub

Private Sub E1Play_Click()

If Est12Control.StopLabel1.Caption = "Stream" Then
    Stream01Play 0   'Stream play
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        Music01Play    'Music play
    End If
End If

Est01.Caption = "ESTACION 01 - Reproduciendo"
DefaultESTDisplay 1
Est01.Label1.ForeColor = &HFFFF00
Est12Control.TmrPos1.Enabled = True
Est12Control.TmrPos1.Interval = 100
UpdatePos
CmdFIN_Click

End Sub

Private Sub E1Pos_Scroll()

Dim Cnv1 As Long

If Est12Control.StopLabel1.Caption = "Stream" Then
    Cnv1 = E1Pos.value
    'change the stream position
    Stream01SetPosition Cnv1, 1
    E1Pos.ToolTipText = GetTimeFromSegs(E1Pos.value)
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        Cnv1 = E1Pos.value
        'change the music position
        Music01SetPosition Cnv1, 0
        E1Pos.ToolTipText = Str$(E1Pos.value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub E1Prev_Click()

If Est12Control.StopLabel1.Caption = "Stream" Then
    Stream01Restart    'stream restart
    Stream01Play 0       'stream play
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        Music01Restart     'music restart
        Music01Play         'music play
    End If
End If

Est01.Caption = "ESTACION 01 - Reproduciendo"
DefaultESTDisplay 1
Est01.Label1.ForeColor = &HFFFF00
Est12Control.TmrPos1.Enabled = True
Est12Control.TmrPos1.Interval = 100

End Sub

Private Sub E1Save_Click()

Dim ContNum As Integer
Dim FileName As String
Dim LenFN As Long
Dim FileNTest As String
Dim NameFile As String

ContNum = CInt(Est01.Fi.Caption)    'extraemos el index del control
FileName = Trim(Est12Data.N1(ContNum).Caption)    'extraemos el path y el archivo de audio

LenFN = Len(FileName)
FileNTest = Mid$(FileName, LenFN - 2, 1)
If FileNTest = "." Then
    LenFN = LenFN - 3
Else
    LenFN = LenFN - 4
End If

NameFile = Left$(FileName, LenFN)
FileName = Trim(NameFile) & AppCUEFileExt

'guardamos la informacion CUe
SaveCUEFile 1, FileName

End Sub

Private Sub E1Slide_Change()

If Est12Control.StopLabel1.Caption = "Stream" Then
    'change the stream pan position
    Stream01SetPan (E1Slide.value)
    E1Slide.ToolTipText = GetTimeFromSegs(E1Slide.value)
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        'change the music pan position
        Music01SetPan (E1Slide.value)
        E1Slide.ToolTipText = GetTimeFromSegs(E1Slide.value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub E1Stop_Click()

CmdFOut_Click

If Est12Control.StopLabel1.Caption = "Stream" Then
    Stream01Restart    'stream restart
    Stream01Stop       'stream stop
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        Music01Restart     'music restart
        Music01Stop         'music stop
    End If
End If

Est01.Caption = "ESTACION 01 - Detenido"
DefaultESTDisplay 1
Est01.Label1.ForeColor = &H808000     'celeste oscuro(desactivado)
Est12Control.TmrPos1.Interval = 0
Est12Control.TmrPos1.Enabled = False

End Sub

Private Sub E1Vol_Change()

If Est12Control.StopLabel1.Caption = "Stream" Then
    'change the stream volume
    Stream01SetVolume (E1Vol.value)
    E1Vol.ToolTipText = GetTimeFromSegs(E1Vol.value)
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        'change the music volume
        Music01SetVolume (E1Vol.value)
        E1Vol.ToolTipText = GetTimeFromSegs(E1Vol.value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub Form_Load()

UpdatePos

TmrUpdt.Enabled = True
TmrUpdt.Interval = 100

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

TmrUpdt.Interval = 0
TmrUpdt.Enabled = False
TmrCUE.Interval = 0
TmrCUE.Enabled = False

'RResult = HideWindow("Diag01")

End Sub

Private Sub Form_Terminate()

TmrUpdt.Interval = 0
TmrUpdt.Enabled = False
TmrCUE.Interval = 0
TmrCUE.Enabled = False

'RResult = HideWindow("Diag01")

End Sub

Private Sub Form_Unload(Cancel As Integer)

TmrUpdt.Interval = 0
TmrUpdt.Enabled = False
TmrCUE.Interval = 0
TmrCUE.Enabled = False

'RResult = HideWindow("Diag01")

End Sub

Private Sub TmrCUE_Timer()

    Dim StartByte As Long
    Dim EndByte As Long
    Dim ActualByte As Long
    
    StartByte = CLng(LblStartCUE.Caption)
    EndByte = CLng(LblEndCue.Caption)
    ActualByte = CLng(LblCurrByte.Caption)
    
    'calculations
    EndByte = (EndByte / 6000) / 3
    ActualByte = (ActualByte / 6000) / 3
    
    'change the stream position for a cue start
If Est12Control.StopLabel1.Caption = "Stream" Then
    Do While ActualByte >= EndByte
        Stream01SetPosition StartByte, 2
        Exit Do
    Loop
    E1Pos.ToolTipText = GetTimeFromSegs(E1Pos.value)
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
'        'Cnv1 = E1Pos.Value
'        'change the music position
'        'Music01SetPosition Cnv1, 0
'        'E1Pos.ToolTipText = Str$(E1Pos.Value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub TmrUpdt_Timer()

'UpdatePos

Dim BytePos As String
Dim TimePos As String
Dim Convt1 As Long

If Est12Control.StopLabel1.Caption = "Stream" Then
    TimePos = Stream01GetPosition(1) 'get position in time
    BytePos = Stream01GetPosition(2) 'get position in bytes
    TimePos = GetSegsFromTime(TimePos)
    E1Pos.value = TimePos
    E1Pos.ToolTipText = GetTimeFromSegs(E1Pos.value)
    LblCurrent.Caption = GetTimeFromSegs(TimePos)
    LblCurrByte.Caption = BytePos
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        BytePos = Music01GetPosition(1)
        BytePos = Left$(BytePos, 2)
        Convt1 = CLng(BytePos)
        E1Pos.value = Convt1
        LblCurrent.Caption = Convt1
        E1Pos.ToolTipText = Str$(E1Pos.value)
    Else
        Exit Sub
    End If
End If

End Sub
