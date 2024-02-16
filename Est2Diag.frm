VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Est2Diag 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ESTACION 02"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6780
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton E2Stop 
      Height          =   375
      Left            =   2310
      Picture         =   "Est2Diag.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Detener"
      Top             =   3630
      Width           =   855
   End
   Begin VB.CommandButton E2Pause 
      Height          =   375
      Left            =   1470
      Picture         =   "Est2Diag.frx":0A3C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Pausar"
      Top             =   3630
      Width           =   855
   End
   Begin VB.CommandButton E2Play 
      Height          =   375
      Left            =   630
      Picture         =   "Est2Diag.frx":1478
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Reproducir"
      Top             =   3630
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   ">"
      Height          =   495
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Ocultar propiedades de audio"
      Top             =   210
      Width           =   255
   End
   Begin VB.CommandButton E2Cue 
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
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Activar / Desactivar CUE"
      Top             =   4290
      Width           =   1815
   End
   Begin VB.Timer TmrUpdt 
      Left            =   4770
      Top             =   5130
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   90
      TabIndex        =   31
      Top             =   4290
      Width           =   1095
   End
   Begin VB.Timer TmrCUE 
      Left            =   5190
      Top             =   5130
   End
   Begin VB.CommandButton CmdRestore 
      Caption         =   "&Restablecer"
      Height          =   375
      Left            =   1290
      TabIndex        =   30
      Top             =   4290
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Volumen"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   90
      TabIndex        =   26
      Top             =   90
      Width           =   6255
      Begin VB.CommandButton CmdFIN 
         Caption         =   "FADE IN"
         Height          =   255
         Left            =   4185
         TabIndex        =   28
         Top             =   225
         Width           =   855
      End
      Begin VB.CommandButton CmdFOut 
         Caption         =   "FADE OUT"
         Height          =   255
         Left            =   5160
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin MSComctlLib.Slider E2Vol 
         Height          =   255
         Left            =   120
         TabIndex        =   29
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
   Begin VB.Frame Frame2 
      Caption         =   "Paneo"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   90
      TabIndex        =   21
      Top             =   810
      Width           =   6255
      Begin VB.CommandButton CmdAutoPan 
         Caption         =   "AUTO PANEO"
         Height          =   255
         Left            =   4920
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin MSComctlLib.Slider E2Slide 
         Height          =   255
         Left            =   240
         TabIndex        =   23
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
      Begin VB.Label Label3 
         Caption         =   "Dr"
         Height          =   255
         Left            =   4500
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Iz"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Posicionamiento "
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   90
      TabIndex        =   15
      Top             =   1530
      Width           =   6255
      Begin MSComctlLib.Slider E2Pos 
         Height          =   315
         Left            =   120
         TabIndex        =   16
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
         TabIndex        =   20
         Top             =   570
         Width           =   1170
      End
      Begin VB.Label Label11 
         Caption         =   "Finalización:"
         Height          =   255
         Left            =   3990
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   585
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "En proceso:"
         Height          =   255
         Left            =   210
         TabIndex        =   17
         Top             =   630
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "CUE"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1590
      TabIndex        =   8
      Top             =   2610
      Width           =   3195
      Begin VB.CommandButton Command5 
         Caption         =   "<M"
         Height          =   300
         Left            =   2700
         TabIndex        =   12
         ToolTipText     =   "Marcar la posicion de fin de CUE"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<M"
         Height          =   300
         Left            =   1140
         TabIndex        =   11
         ToolTipText     =   "Marcar la posicion de inicio de CUE"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "00:00"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "00:00"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CUE Final"
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CUE Inicio"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton E2New 
      Height          =   375
      Left            =   3450
      Picture         =   "Est2Diag.frx":1EB4
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Nuevo archivo CUE"
      Top             =   3630
      Width           =   495
   End
   Begin VB.CommandButton E2Prev 
      Height          =   375
      Left            =   150
      Picture         =   "Est2Diag.frx":23E6
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Detener"
      Top             =   3630
      Width           =   495
   End
   Begin VB.CommandButton E2Import 
      Caption         =   "&Importar"
      Height          =   360
      Left            =   5340
      TabIndex        =   2
      ToolTipText     =   "Importar archivo CUE"
      Top             =   3630
      Width           =   915
   End
   Begin VB.CommandButton E2Open 
      Height          =   375
      Left            =   4050
      Picture         =   "Est2Diag.frx":2E22
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abrir archivo CUE"
      Top             =   3630
      Width           =   495
   End
   Begin VB.CommandButton E2Save 
      Height          =   375
      Left            =   4650
      Picture         =   "Est2Diag.frx":2F24
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Guardar archivo CUE"
      Top             =   3630
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CmdImport 
      Left            =   5730
      Top             =   5070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line1 
      X1              =   90
      X2              =   6330
      Y1              =   4170
      Y2              =   4170
   End
   Begin VB.Label LblStartCUE 
      Caption         =   "0"
      Height          =   255
      Left            =   90
      TabIndex        =   37
      Top             =   5070
      Width           =   1095
   End
   Begin VB.Label LblEndCue 
      Caption         =   "0"
      Height          =   255
      Left            =   1290
      TabIndex        =   36
      Top             =   5070
      Width           =   1215
   End
   Begin VB.Label LblCurrByte 
      Caption         =   "0"
      Height          =   255
      Left            =   90
      TabIndex        =   35
      Top             =   5310
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   90
      Top             =   3570
      Width           =   6255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000018&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4455
      Left            =   6450
      Top             =   210
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   240
      Left            =   150
      TabIndex        =   34
      Top             =   2625
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "Est2Diag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub UpdatePos()

Dim ByteLen As String
Dim TimeLen As String
Dim FTime As String
Dim Convt2 As Long

If Est12Control.StopLabel2.Caption = "Stream" Then
    TimeLen = Stream02GetLen(1) 'get len of file in time=1
    FTime = GetSegsFromTime(TimeLen) 'formateamos el tiempo
    E2Pos.min = 0
    If FTime = 0 Then
        E2Pos.max = FTime + 1
    Else
        E2Pos.max = FTime
    End If
    If FTime <= 100 Then
        E2Pos.TickFrequency = 1
    Else
        If FTime > 100 And FTime < 200 Then
            E2Pos.TickFrequency = 2
        Else
            If FTime > 200 And FTime < 300 Then
                E2Pos.TickFrequency = 3
            Else
                If FTime > 300 And FTime < 400 Then
                    E2Pos.TickFrequency = 4
                Else
                    E2Pos.TickFrequency = 5
                End If
            End If
        End If
    End If
    E2Pos.value = 0
    E2Vol.value = 100
    E2Slide.value = 0
    LblEnd.Caption = GetTimeFromSegs(FTime)
    Est2Diag.Caption = "ESTACION 02 - " & Est02.Label1.Caption
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        'ByteLen = Music02GetLen(2)
        'Convt2 = CLng(ByteLen)
        'Convt2 = Convt2
        'E2Pos.Min = 0
        'If Convt2 = 0 Then
        '    E2Pos.max = Convt2 + 1
        'Else
        '    E2Pos.max = Convt2
        'End If
        'E2Pos.TickFrequency = 1
        'E2Pos.Value = 0
        'E2Pos.ToolTipText = Str$(E2Pos.Value)
        'E2Vol.Value = 100
        'E2Slide.Value = 0
        'LblEnd.Caption = Convt2
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
ActualPan = E2Slide.value

While ActualPan < PanRight
    ActualPan = E2Slide.value + 5   'de o a 100
    E2Slide.value = ActualPan
Wend
While ActualPan > PanOrigen
    ActualPan = E2Slide.value - 5   'de 100 a 0
    E2Slide.value = ActualPan
Wend
While ActualPan > PanLeft
    ActualPan = E2Slide.value - 5   'de 0 a -100
    E2Slide.value = ActualPan
Wend
While ActualPan < PanOrigen
    ActualPan = E2Slide.value + 5   'de -100 a 0
    E2Slide.value = ActualPan
Wend

End Sub

Private Sub CmdFIN_Click()

'FADE IN CLICK

Dim VolOrigen As Long
Dim VolDestino As Long
Dim ActualVol As Long

E2Vol.value = 0
VolOrigen = 0
VolDestino = 100
ActualVol = E2Vol.value

While ActualVol < VolDestino
    E2Vol.value = E2Vol.value + 5
    ActualVol = E2Vol.value
Wend

End Sub

Private Sub CmdFOut_Click()

'FADE OUT CLICK

Dim VolOrigen As Long
Dim VolDestino As Long
Dim ActualVol As Long

E2Vol.value = 100
VolOrigen = E2Vol.value
VolDestino = 0
ActualVol = E2Vol.value

While ActualVol > VolDestino
    E2Vol.value = E2Vol.value - 5
    ActualVol = E2Vol.value
Wend

End Sub

Private Sub CmdRestore_Click()

E2Vol.value = 100
E2Slide.value = 0
E2Pos.value = 0
E2Pos.SelStart = 0
E2Pos.SelLength = 0
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
E2Prev.Enabled = False
E2Play.Enabled = False
E2Pause.Enabled = False
E2Stop.Enabled = False
E2New.Enabled = False
E2Open.Enabled = False
E2Save.Enabled = False
E2Import.Enabled = False
E2Cue.Enabled = False
E2Pos.SelStart = E2Pos.value

End Sub

Private Sub Command5_Click()

Text2.Text = LblCurrent.Caption
LblEndCue.Caption = LblCurrByte.Caption
Command3.Enabled = True
Command4.Enabled = True
CmdRestore.Enabled = True
E2Prev.Enabled = True
E2Play.Enabled = True
E2Pause.Enabled = True
E2Stop.Enabled = True
E2New.Enabled = True
E2Open.Enabled = True
E2Save.Enabled = True
E2Import.Enabled = True
E2Cue.Enabled = True
E2Pos.SelLength = E2Pos.value - E2Pos.SelStart
Call E2Cue_Click

End Sub

Private Sub E2Cue_Click()

If E2Cue.Caption = "Activar CUE" Then
    E2Cue.Caption = "Desactivar CUE"
    E2Cue.BackColor = &HFFFF&   'amarillo
    Text1.Enabled = False
    Text2.Enabled = False
    Command3.Enabled = False
    Command5.Enabled = False
    Command4.Enabled = False
    CmdRestore.Enabled = False
    E2Prev.Enabled = False
    E2Play.Enabled = False
    E2Pause.Enabled = False
    E2Stop.Enabled = False
    E2New.Enabled = False
    E2Open.Enabled = False
    E2Save.Enabled = False
    E2Import.Enabled = False
    TmrCUE.Enabled = True
    TmrCUE.Interval = 1
Else
    E2Cue.Caption = "Activar CUE"
    E2Cue.BackColor = &H8000000F    'gris
    Text1.Enabled = True
    Text2.Enabled = True
    Command3.Enabled = True
    Command5.Enabled = True
    Command4.Enabled = True
    CmdRestore.Enabled = True
    E2Prev.Enabled = True
    E2Play.Enabled = True
    E2Pause.Enabled = True
    E2Stop.Enabled = True
    E2New.Enabled = True
    E2Open.Enabled = True
    E2Save.Enabled = True
    E2Import.Enabled = True
    TmrCUE.Interval = 0
    TmrCUE.Enabled = False
End If

End Sub

Private Sub E2New_Click()

TmrUpdt.Interval = 0
TmrUpdt.Enabled = False

UpdatePos

TmrUpdt.Enabled = True
TmrUpdt.Interval = 100

CmdRestore_Click

End Sub

Private Sub E2Open_Click()

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
FileName = Trim(Est12Data.N2(ContNum).Caption)    'extraemos el path y el archivo de audio

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
OpenCUEFile 2, FileName

End Sub

Private Sub E2Pause_Click()

If Est12Control.StopLabel2.Caption = "Stream" Then
    Stream02Stop   'stream stop
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        Music02Stop    'music stop
    End If
End If

Est02.Caption = "ESTACION 02 - Pausado"
Est12Control.TmrPos2.Interval = 0
Est12Control.TmrPos2.Enabled = False

End Sub

Private Sub E2Play_Click()

If Est12Control.StopLabel2.Caption = "Stream" Then
    Stream02Play 0   'Stream play
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        Music02Play    'Music play
    End If
End If

Est02.Caption = "ESTACION 02 - Reproduciendo"
DefaultESTDisplay 2
Est02.Label1.ForeColor = &HFFFF00
Est12Control.TmrPos2.Enabled = True
Est12Control.TmrPos2.Interval = 100
UpdatePos
CmdFIN_Click

End Sub

Private Sub E2Prev_Click()

If Est12Control.StopLabel2.Caption = "Stream" Then
    Stream02Restart    'stream restart
    Stream02Play 0       'stream play
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        Music02Restart     'music restart
        Music02Play         'music play
    End If
End If

Est02.Caption = "ESTACION 02 - Reproduciendo"
DefaultESTDisplay 2
Est02.Label1.ForeColor = &HFFFF00
Est12Control.TmrPos2.Enabled = True
Est12Control.TmrPos2.Interval = 100

End Sub

Private Sub E2Save_Click()

Dim ContNum As Integer
Dim FileName As String
Dim LenFN As Long
Dim FileNTest As String
Dim NameFile As String

ContNum = CInt(Est02.Fi.Caption)    'extraemos el index del control
FileName = Trim(Est12Data.N2(ContNum).Caption)    'extraemos el path y el archivo de audio

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
SaveCUEFile 2, FileName

End Sub

Private Sub E2Stop_Click()

CmdFOut_Click

If Est12Control.StopLabel2.Caption = "Stream" Then
    Stream02Restart    'stream restart
    Stream02Stop       'stream stop
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        Music02Restart     'music restart
        Music02Stop         'music stop
    End If
End If

Est02.Caption = "ESTACION 02 - Detenido"
DefaultESTDisplay 2
Est02.Label1.ForeColor = &H808000     'celeste oscuro(desactivado)
Est12Control.TmrPos2.Interval = 0
Est12Control.TmrPos2.Enabled = False

End Sub

Private Sub E2Pos_Scroll()

Dim Cnv2 As Long

If Est12Control.StopLabel2.Caption = "Stream" Then
    Cnv2 = E2Pos.value
    'change the stream position
    Stream02SetPosition Cnv2, 1
    E2Pos.ToolTipText = GetTimeFromSegs(E2Pos.value)
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        Cnv2 = E2Pos.value
        'change the music position
        Music02SetPosition Cnv2, 0
        E2Pos.ToolTipText = Str$(E2Pos.value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub E2Slide_Change()

If Est12Control.StopLabel2.Caption = "Stream" Then
    'change the stream pan position
    Stream02SetPan (E2Slide.value)
    E2Slide.ToolTipText = GetTimeFromSegs(E2Slide.value)
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        'change the music pan position
        Music02SetPan (E2Slide.value)
        E2Slide.ToolTipText = GetTimeFromSegs(E2Slide.value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub E2Vol_Change()

If Est12Control.StopLabel2.Caption = "Stream" Then
    'change the stream volume
    Stream02SetVolume (E2Vol.value)
    E2Vol.ToolTipText = GetTimeFromSegs(E2Vol.value)
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        'change the music volume
        Music02SetVolume (E2Vol.value)
        E2Vol.ToolTipText = GetTimeFromSegs(E2Vol.value)
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

'RResult = HideWindow("Diag02")

End Sub

Private Sub Form_Terminate()

TmrUpdt.Interval = 0
TmrUpdt.Enabled = False
TmrCUE.Interval = 0
TmrCUE.Enabled = False

'RResult = HideWindow("Diag02")

End Sub

Private Sub Form_Unload(Cancel As Integer)

TmrUpdt.Interval = 0
TmrUpdt.Enabled = False
TmrCUE.Interval = 0
TmrCUE.Enabled = False

'RResult = HideWindow("Diag02")

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
If Est12Control.StopLabel2.Caption = "Stream" Then
    Do While ActualByte >= EndByte
        Stream02SetPosition StartByte, 2
        Exit Do
    Loop
    E2Pos.ToolTipText = GetTimeFromSegs(E2Pos.value)
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
'        'Cnv2 = E2Pos.Value
'        'change the music position
'        'Music02SetPosition Cnv2, 0
'        'E2Pos.ToolTipText = Str$(E2Pos.Value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub TmrUpdt_Timer()

'UpdatePos

Dim BytePos As String
Dim TimePos As String
Dim Convt2 As Long

If Est12Control.StopLabel2.Caption = "Stream" Then
    TimePos = Stream02GetPosition(1) 'get position in time
    BytePos = Stream02GetPosition(2) 'get position in bytes
    TimePos = GetSegsFromTime(TimePos)
    E2Pos.value = TimePos
    E2Pos.ToolTipText = GetTimeFromSegs(E2Pos.value)
    LblCurrent.Caption = GetTimeFromSegs(TimePos)
    LblCurrByte.Caption = BytePos
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        BytePos = Music02GetPosition(1)
        BytePos = Left$(BytePos, 2)
        Convt2 = CLng(BytePos)
        E2Pos.value = Convt2
        LblCurrent.Caption = Convt2
        E2Pos.ToolTipText = Str$(E2Pos.value)
    Else
        Exit Sub
    End If
End If

End Sub
