VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form Est01 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "ESTACION 01 - Detenido"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   15360
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Est01.frx":0000
   ScaleHeight     =   5205
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin RM100.DC_Control_Bt E1New 
      Height          =   465
      Left            =   6240
      TabIndex        =   135
      Top             =   4620
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   820
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   4210752
      PicDown         =   "Est01.frx":2458F
      PicHot          =   "Est01.frx":7EB41
      PicNormal       =   "Est01.frx":D90F3
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin RM100.DC_Control_Bt E1Play 
      Height          =   465
      Left            =   3930
      TabIndex        =   132
      Top             =   4620
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   820
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   4210752
      PicDown         =   "Est01.frx":1336A5
      PicHot          =   "Est01.frx":18DC57
      PicNormal       =   "Est01.frx":1E8209
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
   End
   Begin RM100.DC_Control_Bt P11 
      Height          =   285
      Index           =   0
      Left            =   4440
      TabIndex        =   121
      Top             =   1080
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   "1"
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   0
      Left            =   150
      TabIndex        =   97
      Top             =   1500
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin VB.VScrollBar fxsc 
      Height          =   1050
      Index           =   10
      Left            =   11640
      Max             =   20
      TabIndex        =   91
      Top             =   2100
      Value           =   18
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   9
      Left            =   11265
      Max             =   20
      TabIndex        =   90
      Top             =   1860
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   8
      Left            =   10950
      Max             =   20
      TabIndex        =   89
      Top             =   1860
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   7
      Left            =   10635
      Max             =   20
      TabIndex        =   88
      Top             =   1860
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   6
      Left            =   10320
      Max             =   20
      TabIndex        =   87
      Top             =   1860
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   5
      Left            =   10005
      Max             =   20
      TabIndex        =   86
      Top             =   1860
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   4
      Left            =   9690
      Max             =   20
      TabIndex        =   85
      Top             =   1860
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   3
      Left            =   9375
      Max             =   20
      TabIndex        =   84
      Top             =   1860
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   2
      Left            =   9060
      Max             =   20
      TabIndex        =   83
      Top             =   1860
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   1
      Left            =   8745
      Max             =   20
      TabIndex        =   82
      Top             =   1860
      Value           =   10
      Width           =   240
   End
   Begin VB.VScrollBar fxsc 
      Height          =   1290
      Index           =   0
      Left            =   8430
      Max             =   20
      TabIndex        =   81
      Top             =   1860
      Value           =   10
      Width           =   240
   End
   Begin VB.PictureBox Picfft1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   8310
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   135
      TabIndex        =   59
      Top             =   975
      Width           =   2025
   End
   Begin VB.PictureBox pcontdw 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   11430
      Picture         =   "Est01.frx":2427BB
      ScaleHeight     =   255
      ScaleWidth      =   405
      TabIndex        =   58
      ToolTipText     =   "Modo CONTINUO desactivado"
      Top             =   570
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.CommandButton CmdFOut 
      Caption         =   "FADE OUT"
      Enabled         =   0   'False
      Height          =   495
      Left            =   12810
      TabIndex        =   57
      ToolTipText     =   "Fade-Out automatico"
      Top             =   1890
      Width           =   645
   End
   Begin VB.CommandButton CmdFIN 
      Caption         =   "FADE IN"
      Enabled         =   0   'False
      Height          =   495
      Left            =   12150
      TabIndex        =   56
      ToolTipText     =   "Fade-In automatico"
      Top             =   1890
      Width           =   555
   End
   Begin VB.CommandButton CmdAutoPan 
      Caption         =   "AUTO PANEO"
      Height          =   255
      Left            =   13800
      TabIndex        =   55
      ToolTipText     =   "Paneo Izq>Der - Der>Izq - automatico"
      Top             =   1950
      Width           =   1215
   End
   Begin VB.CommandButton E1Cue 
      Caption         =   "AC"
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
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Activar / Desactivar CUE"
      Top             =   4620
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton CmdActualiz 
      Caption         =   "A"
      Height          =   375
      Left            =   8130
      TabIndex        =   48
      ToolTipText     =   "Actualizar datos"
      Top             =   4635
      UseMaskColor    =   -1  'True
      Width           =   915
   End
   Begin VB.CommandButton CmdRestore 
      Caption         =   "R"
      Height          =   375
      Left            =   9045
      TabIndex        =   47
      ToolTipText     =   "Restablecer valores a predeterminado"
      Top             =   4635
      UseMaskColor    =   -1  'True
      Width           =   1005
   End
   Begin VB.CommandButton NewCUE 
      Height          =   375
      Left            =   10455
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Restaurar Preferencias"
      Top             =   4635
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton E1Import 
      Caption         =   "I"
      Height          =   375
      Left            =   11760
      TabIndex        =   45
      ToolTipText     =   "Importar archivo CUE"
      Top             =   4635
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.CommandButton OpenCUE 
      Height          =   375
      Left            =   10830
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Abrir Preferencias"
      Top             =   4635
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton SaveCUE 
      Height          =   375
      Left            =   11205
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Guardar Preferencias"
      Top             =   4635
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.TextBox TxtName 
      Height          =   465
      Left            =   5670
      TabIndex        =   33
      Text            =   "Text3"
      Top             =   5610
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.PictureBox E1p6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2520
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   38
      Top             =   975
      Width           =   190
   End
   Begin VB.PictureBox E1p0 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2520
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   37
      Top             =   630
      Width           =   190
   End
   Begin VB.Timer TmoutAuto 
      Left            =   12390
      Top             =   5580
   End
   Begin VB.Timer Tmout 
      Left            =   11895
      Top             =   5580
   End
   Begin VB.Timer TMin 
      Left            =   11400
      Top             =   5580
   End
   Begin VB.Timer TmrScopeLite 
      Left            =   105
      Top             =   5625
   End
   Begin VB.Timer TmrCUE 
      Left            =   14730
      Top             =   5535
   End
   Begin VB.Frame Frame3 
      Caption         =   "Posicionamiento "
      ForeColor       =   &H000000FF&
      Height          =   1020
      Left            =   3600
      TabIndex        =   28
      Top             =   7050
      Width           =   7155
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Finalización:"
         Height          =   255
         Left            =   4950
         TabIndex        =   30
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "En proceso:"
         Height          =   255
         Left            =   210
         TabIndex        =   29
         Top             =   720
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "CUE"
      ForeColor       =   &H000000FF&
      Height          =   870
      Left            =   12015
      TabIndex        =   23
      Top             =   2520
      Width           =   3195
      Begin VB.CommandButton Command10 
         Caption         =   "<M"
         Height          =   300
         Left            =   2655
         TabIndex        =   51
         ToolTipText     =   "Marcar la posicion de fin de CUE"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "<M"
         Height          =   300
         Left            =   1170
         TabIndex        =   50
         ToolTipText     =   "Marcar la posicion de inicio de CUE"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "00:00:00"
         Top             =   495
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "00:00:00"
         Top             =   495
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CUE Final"
         Height          =   255
         Left            =   1665
         TabIndex        =   27
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CUE Inicio"
         Height          =   255
         Left            =   165
         TabIndex        =   26
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3930
      Picture         =   "Est01.frx":242D93
      ScaleHeight     =   300
      ScaleWidth      =   3630
      TabIndex        =   19
      Top             =   450
      Width           =   3630
      Begin VB.PictureBox Ll 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   60
         Picture         =   "Est01.frx":2453A7
         ScaleHeight     =   240
         ScaleMode       =   0  'User
         ScaleWidth      =   3510
         TabIndex        =   20
         Top             =   30
         Width           =   3510
      End
   End
   Begin VB.CommandButton E1Savea 
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Archivo de Estacion"
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton E1Opena 
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Abrir archivo de Estacion"
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton E1Newa 
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Nuevo archivo de estacion"
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton E1Stopa 
      Height          =   375
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Detener"
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.PictureBox E1p5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3495
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   630
      Width           =   190
   End
   Begin VB.PictureBox E1p4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3300
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   630
      Width           =   190
   End
   Begin VB.PictureBox E1p3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3105
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   630
      Width           =   190
   End
   Begin VB.PictureBox E1p2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2925
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   630
      Width           =   190
   End
   Begin VB.PictureBox E1p1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2730
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   630
      Width           =   190
   End
   Begin VB.PictureBox E1p11 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3495
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   975
      Width           =   190
   End
   Begin VB.PictureBox E1p10 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3300
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   975
      Width           =   190
   End
   Begin VB.PictureBox E1p9 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3105
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   975
      Width           =   190
   End
   Begin VB.PictureBox E1p8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2925
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   975
      Width           =   190
   End
   Begin VB.PictureBox E1p7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2730
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   975
      Width           =   190
   End
   Begin VB.CommandButton E1Pausea 
      Height          =   375
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Pausar"
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton E1Playa 
      Height          =   375
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Reproducir"
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.PictureBox Picture4 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   53
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Picture5 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   54
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox pcontup 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   11430
      Picture         =   "Est01.frx":24A6ED
      ScaleHeight     =   255
      ScaleWidth      =   405
      TabIndex        =   75
      ToolTipText     =   "Modo CONTINUO activado"
      Top             =   570
      Width           =   405
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3930
      Picture         =   "Est01.frx":24ACC5
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   3630
      TabIndex        =   21
      Top             =   750
      Width           =   3630
      Begin VB.PictureBox Lr 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   60
         Picture         =   "Est01.frx":24D2D9
         ScaleHeight     =   240
         ScaleWidth      =   3510
         TabIndex        =   22
         Top             =   30
         Width           =   3510
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   2370
      ScaleHeight     =   825
      ScaleWidth      =   1395
      TabIndex        =   76
      Top             =   510
      Width           =   1395
   End
   Begin ComctlLib.Slider E1Pos 
      Height          =   225
      Left            =   8280
      TabIndex        =   79
      Top             =   3780
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   397
      _Version        =   327682
      Max             =   100
      TickFrequency   =   5
   End
   Begin RM100.TitelBar TitelBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   661
      BackColor       =   8421504
      BackColorCover  =   3
      BackColorV2Begin=   4210752
      BackColorV2End  =   0
      BackColorV1Begin=   4210752
      BackColorV1End  =   0
      ForeColor       =   16777215
      ShowClose       =   0   'False
      ShowCloseEnabled=   0   'False
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
      Caption         =   " ESTACION 01 - Detenido"
      CaptionPosX     =   1
      BorderNormal    =   2
      BorderColorDarkLight=   12632256
   End
   Begin VB.PictureBox Picture3 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   94
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Picture7 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   95
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Picture8 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   96
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Picture9 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   98
      Top             =   0
      Width           =   0
   End
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   1
      Left            =   1980
      TabIndex        =   99
      Top             =   1500
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin VB.PictureBox Picture10 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   100
      Top             =   0
      Width           =   0
   End
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   2
      Left            =   3810
      TabIndex        =   101
      Top             =   1500
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   3
      Left            =   5640
      TabIndex        =   102
      Top             =   1500
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   4
      Left            =   150
      TabIndex        =   103
      Top             =   2100
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   5
      Left            =   1980
      TabIndex        =   104
      Top             =   2100
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   6
      Left            =   3810
      TabIndex        =   105
      Top             =   2100
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   7
      Left            =   5640
      TabIndex        =   106
      Top             =   2100
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   8
      Left            =   150
      TabIndex        =   107
      Top             =   2700
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   9
      Left            =   1980
      TabIndex        =   108
      Top             =   2700
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   10
      Left            =   3810
      TabIndex        =   109
      Top             =   2700
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   11
      Left            =   5640
      TabIndex        =   110
      Top             =   2700
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   12
      Left            =   150
      TabIndex        =   111
      Top             =   3300
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   13
      Left            =   1980
      TabIndex        =   112
      Top             =   3300
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   14
      Left            =   3810
      TabIndex        =   113
      Top             =   3300
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   15
      Left            =   5640
      TabIndex        =   114
      Top             =   3300
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   16
      Left            =   150
      TabIndex        =   115
      Top             =   3900
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   17
      Left            =   1980
      TabIndex        =   116
      Top             =   3900
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   18
      Left            =   3810
      TabIndex        =   117
      Top             =   3900
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   19
      Left            =   5640
      TabIndex        =   118
      Top             =   3900
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   20
      Left            =   150
      TabIndex        =   119
      Top             =   4500
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin RM100.DC_Control_Bt E11 
      Height          =   585
      Index           =   21
      Left            =   1980
      TabIndex        =   120
      Top             =   4500
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1032
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
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
   Begin VB.PictureBox Picture11 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   122
      Top             =   0
      Width           =   0
   End
   Begin RM100.DC_Control_Bt P11 
      Height          =   285
      Index           =   1
      Left            =   4770
      TabIndex        =   123
      Top             =   1080
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   "2"
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
   Begin VB.PictureBox Picture12 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   124
      Top             =   0
      Width           =   0
   End
   Begin RM100.DC_Control_Bt P11 
      Height          =   285
      Index           =   2
      Left            =   5100
      TabIndex        =   125
      Top             =   1080
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   "3"
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
   Begin RM100.DC_Control_Bt P11 
      Height          =   285
      Index           =   3
      Left            =   5430
      TabIndex        =   126
      Top             =   1080
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   "4"
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
   Begin RM100.DC_Control_Bt P11 
      Height          =   285
      Index           =   4
      Left            =   5760
      TabIndex        =   127
      Top             =   1080
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   "5"
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
   Begin RM100.DC_Control_Bt P11 
      Height          =   285
      Index           =   5
      Left            =   6090
      TabIndex        =   128
      Top             =   1080
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   "6"
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
   Begin RM100.DC_Control_Bt P11 
      Height          =   285
      Index           =   6
      Left            =   6420
      TabIndex        =   129
      Top             =   1080
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   "7"
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
   Begin RM100.DC_Control_Bt P11 
      Height          =   285
      Index           =   7
      Left            =   6750
      TabIndex        =   130
      Top             =   1080
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   "8"
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
   Begin RM100.DC_Control_Bt P11 
      Height          =   285
      Index           =   8
      Left            =   7080
      TabIndex        =   131
      Top             =   1080
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   "9"
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
   Begin RM100.DC_Control_Bt E1Pause 
      Height          =   465
      Left            =   4650
      TabIndex        =   133
      Top             =   4620
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   820
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      MaskColor       =   4210752
      PicDown         =   "Est01.frx":25261F
      PicHot          =   "Est01.frx":2ACBD1
      PicNormal       =   "Est01.frx":307183
      PicSizeH        =   32
      PicSizeW        =   32
   End
   Begin RM100.DC_Control_Bt E1Stop 
      Height          =   465
      Left            =   5370
      TabIndex        =   134
      Top             =   4620
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   820
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   4210752
      PicDown         =   "Est01.frx":361735
      PicHot          =   "Est01.frx":3BBCE7
      PicNormal       =   "Est01.frx":416299
      PicSizeH        =   32
      PicSizeW        =   32
   End
   Begin RM100.DC_Control_Bt E1Open 
      Height          =   465
      Left            =   6660
      TabIndex        =   136
      Top             =   4620
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   820
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   4210752
      PicDown         =   "Est01.frx":47084B
      PicHot          =   "Est01.frx":4CADFD
      PicNormal       =   "Est01.frx":5253AF
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin RM100.DC_Control_Bt E1Save 
      Height          =   465
      Left            =   7080
      TabIndex        =   137
      Top             =   4620
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   820
      BackColor       =   4210752
      ButtonStyle     =   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   4210752
      PicDown         =   "Est01.frx":57F961
      PicHot          =   "Est01.frx":5D9F13
      PicNormal       =   "Est01.frx":6344C5
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin RM100.ucKnob E1Slide 
      Height          =   1305
      Left            =   13710
      TabIndex        =   93
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2302
      Value           =   50
      ForeColor       =   4210752
      TicksLongFrequency=   50
      TicksSmallHiden =   -1  'True
      TicksStyleCircle=   -1  'True
      TickForeColor   =   16776960
   End
   Begin RM100.ucKnob E1Vol 
      Height          =   1305
      Left            =   12120
      TabIndex        =   92
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2302
      Value           =   50
      ForeColor       =   4210752
      TickForeColor   =   16776960
   End
   Begin VB.Label LblEnd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00:00:00"
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
      Left            =   13950
      TabIndex        =   78
      Top             =   4050
      Width           =   1170
   End
   Begin VB.Label LblCurrent 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00:00:00"
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
      Left            =   9240
      TabIndex        =   77
      Top             =   4050
      Width           =   1170
   End
   Begin VB.Label Lfft 
      BackStyle       =   0  'Transparent
      Caption         =   "FFT"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   8310
      TabIndex        =   74
      ToolTipText     =   "Modo FFT"
      Top             =   570
      Width           =   375
   End
   Begin VB.Label Lspc 
      BackStyle       =   0  'Transparent
      Caption         =   "SPC"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   8730
      TabIndex        =   73
      ToolTipText     =   "Modo Espectro"
      Top             =   570
      Width           =   375
   End
   Begin VB.Label Lspcz 
      BackStyle       =   0  'Transparent
      Caption         =   "Izq"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   9345
      TabIndex        =   72
      ToolTipText     =   "Espectro izquierdo"
      Top             =   570
      Width           =   285
   End
   Begin VB.Label Lspcd 
      BackStyle       =   0  'Transparent
      Caption         =   "Der"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   9660
      TabIndex        =   71
      ToolTipText     =   "Espectro derecho"
      Top             =   570
      Width           =   285
   End
   Begin VB.Label Lspcb 
      BackStyle       =   0  'Transparent
      Caption         =   "Amb"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   10020
      TabIndex        =   70
      ToolTipText     =   "Espectro combinado"
      Top             =   570
      Width           =   375
   End
   Begin VB.Label LAplay 
      BackStyle       =   0  'Transparent
      Caption         =   "Autoplay"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   10620
      TabIndex        =   69
      ToolTipText     =   "Autoreproducción al hacer click"
      Top             =   570
      Width           =   645
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F-In/Out:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   10755
      TabIndex        =   68
      Top             =   930
      Width           =   690
   End
   Begin VB.Label LFin 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   11520
      TabIndex        =   67
      Top             =   930
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CUE:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   10800
      TabIndex        =   66
      Top             =   1185
      Width           =   645
   End
   Begin VB.Label LCue 
      BackStyle       =   0  'Transparent
      Caption         =   "Man"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   11520
      TabIndex        =   65
      Top             =   1185
      Width           =   375
   End
   Begin VB.Label fft2 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   10800
      TabIndex        =   64
      Top             =   1485
      Width           =   105
   End
   Begin VB.Label fft4 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   10980
      TabIndex        =   63
      Top             =   1485
      Width           =   150
   End
   Begin VB.Label fft6 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   11205
      TabIndex        =   62
      Top             =   1485
      Width           =   150
   End
   Begin VB.Label fft8 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   11430
      TabIndex        =   61
      Top             =   1485
      Width           =   150
   End
   Begin VB.Label fft10 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   11655
      TabIndex        =   60
      Top             =   1485
      Width           =   195
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "FX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   11640
      TabIndex        =   52
      ToolTipText     =   "Numero de Página"
      Top             =   1830
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   8040
      Top             =   4530
      Width           =   7185
   End
   Begin VB.Label LblOutvol 
      Caption         =   "0"
      Height          =   255
      Left            =   9960
      TabIndex        =   41
      Top             =   6030
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LblInvol 
      Caption         =   "100"
      Height          =   255
      Left            =   9510
      TabIndex        =   40
      Top             =   6030
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LblCurrVol 
      Caption         =   "100"
      Height          =   255
      Left            =   9510
      TabIndex        =   39
      Top             =   5820
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label LblStartCUE 
      Caption         =   "0"
      Height          =   255
      Left            =   8520
      TabIndex        =   36
      Top             =   5580
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label LblEndCue 
      Caption         =   "0"
      Height          =   255
      Left            =   8520
      TabIndex        =   35
      Top             =   5805
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label LblCurrByte 
      Caption         =   "0"
      Height          =   255
      Left            =   9510
      TabIndex        =   34
      Top             =   5580
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image E1Pic 
      Height          =   4560
      Left            =   7575
      Top             =   495
      Width           =   390
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Rvb"
      Height          =   195
      Left            =   11610
      TabIndex        =   32
      Top             =   3165
      Width           =   330
   End
   Begin VB.Label Lindex 
      BackColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   4290
      TabIndex        =   31
      Top             =   5985
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Fi 
      BackColor       =   &H0080FF80&
      Height          =   255
      Left            =   4290
      TabIndex        =   18
      Top             =   5625
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00808000&
      Height          =   660
      Left            =   195
      TabIndex        =   4
      Top             =   600
      Width           =   2130
   End
   Begin VB.Label Pn 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      ToolTipText     =   "Numero de Página"
      Top             =   1065
      Width           =   495
   End
   Begin VB.Label Fn 
      BackColor       =   &H00808000&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1170
      TabIndex        =   1
      Top             =   5625
      Visible         =   0   'False
      Width           =   3075
   End
End
Attribute VB_Name = "Est01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'dimensiones de conversion
Dim ConvertNm           'numeros
Dim ConvertNNm As Long
Dim ConvertTx As String   'textos
Dim ConvertTxT As String
Dim EstNum As Long

Dim FileExt As String

'Dimensiones de Archivos
Dim FileN As String
Dim FileNPath As String
Dim Completo As String
Dim SSTitle As String
Dim FileTP As String

'dimensiones de resultado
Dim Result As String
Dim RResult As String
Dim ResultInfo As Boolean

'Dimensiones de tiempo
Dim TimeNcv As String
Dim PosNcv As String

'dimensiones de datos de configuracion
Dim ConfigData As ConfigRecord

Private Sub UpdatePos()

Dim ByteLen As String
Dim TimeLen As String
Dim FTime As String
Dim Convt1 As Long

On Error Resume Next
If Est12Control.StopLabel1.Caption = "Stream" Then
    TimeLen = Stream01GetLen(1) 'get len of file in time=1
    FTime = FormatSegs(TimeLen) 'formateamos el tiempo
    E1Pos.Min = 0
    If FTime = 0 Or FTime = "" Then
        E1Pos.Max = FTime + 1
    Else
        E1Pos.Max = FTime
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
    E1Pos.Value = 0
    E1Vol.Value = 100
    E1Slide.Value = 0
    E1Pos.SmallChange = 10
    E1Pos.LargeChange = 10
    LblEnd.Caption = ConvSecToMin(CInt(FTime))
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        ByteLen = Music01GetLen(1)  'row/pos
        Convt1 = CLng(ByteLen)
        Convt1 = Convt1
        E1Pos.Min = 0
        If Convt1 = 0 Then
            E1Pos.Max = Convt1 + 1
        Else
            E1Pos.Max = Convt1
        End If
        E1Pos.TickFrequency = 1
        E1Pos.Value = 0
        E1Pos.ToolTipText = Str$(E1Pos.Value)
        E1Vol.Value = 100
        E1Slide.Value = 0
        LblEnd.Caption = Convt1
        E1Pos.SmallChange = 1
        E1Pos.LargeChange = 1
    Else
        Exit Sub
    End If
End If

End Sub

Public Sub SetAudioLevel(WLeft, WRight)

'Dim l, Lft As Integer
'Dim r, Rgt As Integer
'Dim i As Integer
'Static ZMax%, RMax%

'On Error Resume Next
WLeft = WLeft / 7
WRight = WRight / 7

'right level meter
'If WRight > 180 Then
'    RMax = (WRight * 24) + 100 'clip
'Else
'    RMax = (WRight * 24)
'End If

'left level meter
'If WLeft > 180 Then
'    ZMax = (WLeft * 24) + 100  'clip
'Else
'    ZMax = (WLeft * 24)
'End If

Lr.Width = WRight 'RMax
Ll.Width = WLeft ' ZMax

'Debug.Print "R:" & WRight & "L:" & WLeft

End Sub

Private Sub OpenE1PageFile()

'extraemos los datos necesarios para realizar la operacion
ConvertNm = Pn.Caption
Select Case ConvertNm
    Case "1"
        ConvertNNm = 1
    Case "2"
        ConvertNNm = 2
    Case "3"
        ConvertNNm = 3
    Case "4"
        ConvertNNm = 4
    Case "5"
        ConvertNNm = 5
    Case "6"
        ConvertNNm = 6
    Case "7"
        ConvertNNm = 7
    Case "8"
        ConvertNNm = 8
    Case "9"
        ConvertNNm = 9
End Select

ConvertTx = Trim(Fn.Caption)
If ConvertTx = "" Or ConvertTx = " " Then
    TopMenu.EstCmd.InitDir = App.path & AppEstDir
    TopMenu.EstCmd.Filter = "Archivo de Estacion (*.est)|*.est|Archivos de Estacion"
    TopMenu.EstCmd.DialogTitle = "ESTACION 01 - Abrir archivo"
    TopMenu.EstCmd.FilterIndex = 1
    TopMenu.EstCmd.ShowOpen
    ConvertTx = TopMenu.EstCmd.filename
    EstNum = 1
    Result = OpenEstFile(EstNum, ConvertNNm, ConvertTx)
    If Result = "NotOK" Then
        Exit Sub
    End If
    Fn.Caption = ConvertTx
Else
    EstNum = 1
    Result = OpenEstFile(EstNum, ConvertNNm, ConvertTx)
    If Result = "NotOK" Then
        Exit Sub
    End If
End If

End Sub

Private Sub DeployAudioFile(WConNum As Integer)

If XPlorer.File1.filename = "" Or XPlorer.File1.filename = " " Then
    MsgBox LoadResString(137), vbCritical
    Exit Sub
End If

'.wav, .mp3, .it, .xm
FileExt = StripExtFromFile(XPlorer.File1.filename)
FileN = StripFileFromExt(XPlorer.File1.filename)
FileNPath = Right$(XPlorer.lblPath, Len(XPlorer.lblPath) - 2)
Completo = Right$(XPlorer.lblPath, Len(XPlorer.lblPath) - 2) & "\" & XPlorer.File1.filename

'seleccion de formato de archivo y extraccion de informacion header
Select Case Trim(UCase(FileExt))
    
   'STREAM TYPE WAV-MP1-MP2-MP3-OGG
    Case "WAV", "MP1", "MP2", "MP3", "OGG"
        Est12Data.N1(WConNum).Caption = Completo                  'nombre y path
        Est12Data.c1(WConNum).Caption = FileN                     'nombre solo
        'gets the file len and convert into time
        ConvertTx = FileLoadLen(Completo, "Stream")
        TimeNcv = FormatSegs(ConvertTx)
        Result = ConvSecToMin(CInt(TimeNcv))
        'put the file time into est01
        Est12Data.D1(WConNum).Caption = Result
        E11(WConNum).Caption = FileN    'nombre del archivo
        E11(WConNum).BackColor = &H404040
        E11(WConNum).ToolTipText = "Duración: " & Result
        Est12Data.V1(WConNum).Caption = "Stream"
                 
    'MUSIC TYPE XM-MOD-S3M-IT-MTM-MO3-UMX
    Case "XM", "MOD", "S3M", "IT", "MTM", "MO3", "UMX"
        Est12Data.N1(WConNum).Caption = Completo                  'nombre y path
        Est12Data.c1(WConNum).Caption = FileN                     'nombre solo
        Est12Data.D1(WConNum).Caption = ""
        E11(WConNum).Caption = FileN    'nombre del archivo
        E11(WConNum).BackColor = &H404040
        E11(WConNum).ToolTipText = ""
        Est12Data.V1(WConNum).Caption = "Music"
        
    'TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TNDTND
    Case "TND"
        MsgBox LoadResString(191), vbInformation, "Radio Maker"
        E11(WConNum).BackColor = &H404040
        
    Case Else
        MsgBox LoadResString(191), vbInformation, "Radio Maker"
        E11(WConNum).BackColor = &H404040

End Select

End Sub

Private Sub CmdAutoPan_Click()

Dim PanOrigen As Long
Dim PanRight As Long
Dim PanLeft As Long
Dim ActualPan As Long

PanOrigen = 0
PanLeft = -100
PanRight = 100
ActualPan = E1Slide.Value

While ActualPan < PanRight
    ActualPan = E1Slide.Value + 5   'de o a 100
    E1Slide.Value = ActualPan
Wend
While ActualPan > PanOrigen
    ActualPan = E1Slide.Value - 5   'de 100 a 0
    E1Slide.Value = ActualPan
Wend
While ActualPan > PanLeft
    ActualPan = E1Slide.Value - 5   'de 0 a -100
    E1Slide.Value = ActualPan
Wend
While ActualPan < PanOrigen
    ActualPan = E1Slide.Value + 5   'de -100 a 0
    E1Slide.Value = ActualPan
Wend

End Sub

Private Sub CmdFIN_Click()

TMin.Enabled = True
TMin.Interval = 30

End Sub

Private Sub CmdFOut_Click()

Tmout.Enabled = True
Tmout.Interval = 30

End Sub

Private Sub CmdRestore_Click()

'E1Vol.value = 100
'E1Slide.value = 0
'E1Pos.value = 0
E1Pos.SelStart = 0
E1Pos.SelLength = 0
Text1.text = "00:00:00"
Text2.text = "00:00:00"
LblStartCUE.Caption = 0
LblEndCue.Caption = 0

Dim i As Integer
'restore all EQ presets and reverb
For i = 0 To 10
    fxsc(i).Value = 10
Next i
fxsc(10).Value = 18

End Sub

Private Sub Command10_Click()

Text2.text = LblCurrent.Caption
LblEndCue.Caption = LblCurrByte.Caption
E1Pos.SelLength = E1Pos.Value - E1Pos.SelStart

If LCue.Caption = "Auto" Then
    Call E1Cue_Click
End If

End Sub

Private Sub Command9_Click()

Text1.text = LblCurrent.Caption
LblStartCUE.Caption = LblCurrByte.Caption
E1Pos.SelStart = E1Pos.Value

Text2.SetFocus

End Sub

Private Sub E11_Click(index As Integer)

'desactivamos los CUE que esten activados
If E1Cue.Caption = "Desactivar CUE" Then
    E1Cue.Caption = "Activar CUE"
    E1Cue.BackColor = &H8000000F    'gris
    TmrCUE.Interval = 0
    TmrCUE.Enabled = False
End If
If Est02.E2Cue.Caption = "Desactivar CUE" Then
    Est02.E2Cue.Caption = "Activar CUE"
    Est02.E2Cue.BackColor = &H8000000F    'gris
    Est02.TmrCUE.Interval = 0
    Est02.TmrCUE.Enabled = False
End If

Dim X As Integer

X = index
If E11(X).Caption = "" Or E11(X).Caption = " " Then Exit Sub

'load and play the selected file
Est12Control.Origen1.Caption = "E1"
Result = Estacion01Play(index)
If Result = "NotOk" Then Exit Sub

RestoreDisplay 1     'sets the default display
RestoreAllActiveColor 1 'deactivate all controls
ChangeActiveColor index, 1  'activate the current control

Fi.Caption = index

'gets the config device data
ConfigData = OpenConfigFile

If ConfigData.Aud_Show_FTT = 1 Or ConfigData.Aud_Show_SCOPE = 1 Then
    'activate the level meter
    TmrScopeLite.Enabled = True
    TmrScopeLite.Interval = 15
Else
    'deactivate the level meter
    TmrScopeLite.Interval = 0
    TmrScopeLite.Enabled = False
End If

'activate the clock timer
TopMenu.ProcTimer.Enabled = True
TopMenu.ProcTimer.Interval = 1
'actualizamos los controles
UpdatePos

'***********************************************************
'Automatic Open the Preset file for the selected file stream
Dim ContNum As Integer
Dim filename As String
Dim LenFN As Long
Dim FileNTest As String
Dim NameFile As String

If Est01.Fi.Caption = "" Then MsgBox LoadResString(154): Exit Sub

ContNum = CInt(Est01.Fi.Caption)    'extraemos el index del control
filename = Trim(Est12Data.N1(ContNum).Caption)    'extraemos el path y el archivo de audio
NameFile = StripFileFromExt(filename)
filename = Trim(NameFile) & AppCUEFileExt

'abrimos la informacion CUe
OpenCUEFile 1, filename

'starts the fade in/out
If LFin.Caption = "Auto" Then
    If Stream02IsPlaying = True Or Music02IsPlaying = True Then
        If Est12Control.Origen2.Caption = "E2" Then
            Est02.TmoutAuto.Enabled = True
            Est02.TmoutAuto.Interval = 30
        End If
    End If
    If Stream01IsPlaying = True Or Music01IsPlaying = True Then
        If Est12Control.Origen1.Caption = "E1" Then
            E1Vol.Value = 0
            TMin.Enabled = True
            TMin.Interval = 30
        End If
    End If
End If

'chequeamos por el cue auto
If LCue.Caption = "Auto" Then
    Call E1Cue_Click
End If

End Sub

Private Sub E11_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)

DeployAudioFile index   'drag & drop the selected file in xplorer

End Sub

Private Sub E11_DragOver(index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

Select Case State
    Case 0  'drag not finished
        XPlorer.File1.DragIcon = XPlorer.ExCombo.DragIcon
        E11(index).BackColor = &H80FF80    'verde (modificacion)
    Case 1  'finished drag
        XPlorer.File1.DragIcon = XPlorer.tvwDirTree.DragIcon
        E11(index).BackColor = &H404040        'gris (normal)
End Select

End Sub


Private Sub E11_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'button 1=left button
'button 2=right button
'button 4=mid button

Debug.Print Button

If Button = 2 Then
    If E11(index).Caption = "" Or E11(index).Caption = " " Then
        Exit Sub
    End If
    'deploy options menu
    TxtName.Visible = True
    TxtName.Top = E11(index).Top
    TxtName.Left = E11(index).Left
    TxtName.Height = E11(index).Height
    TxtName.Width = E11(index).Width
    TxtName.text = E11(index).Caption
    'TxtName.top
    TxtName.SetFocus
    'seteamos el label para saber de que control se trata
    Lindex.Caption = index
Else
    If Button = 4 Then
    '    'mark control to delete content
        E11(index).BackColor = &HFF&              'rojo
    End If
End If

End Sub


Private Sub E1Cue_Click()

Dim Len1 As Long
Dim Len2 As Long

Len1 = CLng(LblStartCUE.Caption)
Len2 = CLng(LblEndCue.Caption)
If Len1 <= 0 Then
    Exit Sub
Else
    If Len2 <= 0 Then
        Exit Sub
    End If
End If

If E1Cue.Caption = "Activar CUE" Then
    E1Cue.Caption = "Desactivar CUE"
    E1Cue.BackColor = &HFFFF&   'amarillo
    TmrCUE.Enabled = True
    TmrCUE.Interval = 100
Else
    E1Cue.Caption = "Activar CUE"
    E1Cue.BackColor = &H8000000F    'gris
    TmrCUE.Interval = 0
    TmrCUE.Enabled = False
End If

End Sub

Private Sub E1Import_Click()

On Error Resume Next
TopMenu.NTSCmd.InitDir = App.path & AppEstDir
TopMenu.NTSCmd.Filter = "NetShow region (*.txt)|*.txt|NetShow region"
TopMenu.NTSCmd.DialogTitle = "ESTACION 01 - importar archivo"
TopMenu.NTSCmd.CancelError = True
TopMenu.NTSCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

ConvertTx = TopMenu.NTSCmd.filename

Result = GetNetShowAudioRegion(ConvertTx, 1)
If Result = "NotOk" Then
    MsgBox LoadResString(157), vbCritical
    Exit Sub
End If

End Sub

Private Sub E1New_Click()

SetDefControl 1   'set the default control´s caption
Pn.Caption = "1"
Fn.Caption = ""

End Sub

Private Sub E1Open_Click()

'extraemos los datos necesarios para realizar la operacion
ConvertNm = Pn.Caption
Select Case ConvertNm
    Case "1"
        ConvertNNm = 1
    Case "2"
        ConvertNNm = 2
    Case "3"
        ConvertNNm = 3
    Case "4"
        ConvertNNm = 4
    Case "5"
        ConvertNNm = 5
    Case "6"
        ConvertNNm = 6
    Case "7"
        ConvertNNm = 7
    Case "8"
        ConvertNNm = 8
    Case "9"
        ConvertNNm = 9
End Select

On Error Resume Next
TopMenu.EstCmd.InitDir = App.path & AppEstDir
TopMenu.EstCmd.Filter = "Archivo de Estación (*.est)|*.est|Archivos de Estación"
TopMenu.EstCmd.DialogTitle = "ESTACION 01 - Abrir archivo de estación"
TopMenu.EstCmd.CancelError = True
TopMenu.EstCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

ConvertTx = TopMenu.EstCmd.filename

EstNum = 1
Result = OpenEstFile(EstNum, ConvertNNm, ConvertTx)
If Result = "NotOK" Then
    Exit Sub
End If
Fn.Caption = ConvertTx

End Sub

Private Sub E1Pause_Click()

If Est12Control.StopLabel1.Caption = "Stream" And Est12Control.Origen1.Caption = "E1" Then
    Stream01Stop   'stream stop
Else
    If Est12Control.StopLabel1.Caption = "Music" And Est12Control.Origen1.Caption = "E1" Then
        Music01Stop    'music stop
    Else
        Exit Sub
    End If
End If

TitelBar1.Caption = "ESTACION 01 - Pausado"

End Sub

Private Sub E1Play_Click()

'desactivamos los CUE que se encuentren activados
If E1Cue.Caption = "Desactivar CUE" Then
    E1Cue.Caption = "Activar CUE"
    E1Cue.BackColor = &H8000000F    'gris
    TmrCUE.Interval = 0
    TmrCUE.Enabled = False
End If
If Est02.E2Cue.Caption = "Desactivar CUE" Then
    Est02.E2Cue.Caption = "Activar CUE"
    Est02.E2Cue.BackColor = &H8000000F    'gris
    Est02.TmrCUE.Interval = 0
    Est02.TmrCUE.Enabled = False
End If

If Est12Control.StopLabel1.Caption = "Stream" And Est12Control.Origen1.Caption = "E1" Then
    If Est01.pcontup.Visible = True Then    'loop enabled?
        Stream01Play (BASS_SAMPLE_LOOP)
    Else
        Stream01Play (0)
    End If
Else
    If Est12Control.StopLabel1.Caption = "Music" And Est12Control.Origen1.Caption = "E1" Then
        Music01Play    'Music play
    Else
        Exit Sub
    End If
End If

TitelBar1.Caption = "ESTACION 01 - Reproduciendo"
RestoreDisplay 1
Est12Control.Origen1.Caption = "E1"
Label1.ForeColor = &HFFFF00

'gets the config device data
ConfigData = OpenConfigFile

If ConfigData.Aud_Show_FTT = 1 Or ConfigData.Aud_Show_SCOPE = 1 Then
    'activate the level meter
    TmrScopeLite.Enabled = True
    TmrScopeLite.Interval = 25
Else
    'deactivate the level meter
    TmrScopeLite.Interval = 0
    TmrScopeLite.Enabled = False
End If

'activamos el timer de posicion
TopMenu.ProcTimer.Enabled = True
TopMenu.ProcTimer.Interval = 1

'actualizamos los controles
UpdatePos

'starts the fade in/out
If LFin.Caption = "Auto" Then
    If Stream02IsPlaying = True Or Music02IsPlaying = True Then
        If Est12Control.Origen2.Caption = "E2" Then
            Est02.TmoutAuto.Enabled = True
            Est02.TmoutAuto.Interval = 30
        End If
    End If
    If Stream01IsPlaying = True Or Music01IsPlaying = True Then
        If Est12Control.Origen1.Caption = "E1" Then
            E1Vol.Value = 0
            TMin.Enabled = True
            TMin.Interval = 30
        End If
    End If
End If

'checks for cue auto
If LCue.Caption = "Auto" Then
    Call E1Cue_Click
End If

End Sub

Private Sub E1Pos_Scroll()

Dim Cnv1 As Long

If Est12Control.StopLabel1.Caption = "Stream" And Est12Control.Origen1.Caption = "E1" Then
    Cnv1 = E1Pos.Value
    'change the stream position
    Stream01SetPosition Cnv1, 1
    E1Pos.ToolTipText = ConvSecToMin(CInt(E1Pos.Value))
Else
    If Est12Control.StopLabel1.Caption = "Music" And Est12Control.Origen1.Caption = "E1" Then
        Cnv1 = E1Pos.Value
        'change the music position
        Music01SetPosition Cnv1, 0
        E1Pos.ToolTipText = Str$(E1Pos.Value)
    Else
        Exit Sub
    End If
End If

End Sub

Public Sub E1Save_Click()

'extraemos los datos necesarios para realizar la operacion
ConvertNm = Pn.Caption
Select Case ConvertNm
    Case "1"
        ConvertNNm = 1
    Case "2"
        ConvertNNm = 2
    Case "3"
        ConvertNNm = 3
    Case "4"
        ConvertNNm = 4
    Case "5"
        ConvertNNm = 5
    Case "6"
        ConvertNNm = 6
    Case "7"
        ConvertNNm = 7
    Case "8"
        ConvertNNm = 8
    Case "9"
        ConvertNNm = 9
End Select

ConvertTxT = Trim(Fn.Caption)

On Error Resume Next
If ConvertTxT = "" Or ConvertTxT = " " Then
    TopMenu.EstCmd.InitDir = App.path & AppEstDir
    TopMenu.EstCmd.Filter = "Archivo de Estación (*.est)|*.est|Archivos de Estación"
    TopMenu.EstCmd.DialogTitle = "ESTACION 01 - Guardar archivo de estación"
    TopMenu.EstCmd.FilterIndex = 1
    TopMenu.EstCmd.CancelError = True
    TopMenu.EstCmd.ShowSave

    If err.Number = 32755 Then Exit Sub
    
    ConvertTx = TopMenu.EstCmd.filename
    
    Fn.Caption = ConvertTx
    EstNum = 1
    Result = SaveEstFile(EstNum, ConvertNNm, ConvertTx)
    If Result = "NotOK" Then
        Exit Sub
    End If
Else
    ConvertTx = Trim(Fn.Caption)
    EstNum = 1
    Result = SaveEstFile(EstNum, ConvertNNm, ConvertTx)
    If Result = "NotOK" Then
        Exit Sub
    End If
End If

End Sub

Private Sub E1Slide_Change()

If Est12Control.StopLabel1.Caption = "Stream" And Est12Control.Origen1.Caption = "E1" Then
    'change the stream pan position
    Stream01SetPan (E1Slide.Value)
    E1Slide.ToolTipText = E1Slide.Value
Else
    If Est12Control.StopLabel1.Caption = "Music" And Est12Control.Origen1.Caption = "E1" Then
        'change the music pan position
        Music01SetPan (E1Slide.Value)
        E1Slide.ToolTipText = E1Slide.Value
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub E1Stop_Click()

'chequeamos si el fade-out esta en automatico
If LFin.Caption = "Auto" Then
    TmoutAuto.Enabled = True
    TmoutAuto.Interval = 30
    Exit Sub
End If

If Est12Control.StopLabel1.Caption = "Stream" And Est12Control.Origen1.Caption = "E1" Then
    Stream01Restart    'stream restart
    Stream01Stop       'stream stop
Else
    If Est12Control.StopLabel1.Caption = "Music" And Est12Control.Origen1.Caption = "E1" Then
        Music01Restart     'music restart
        Music01Stop         'music stop
    Else
        GoSub Cont
    End If
End If

Cont:
'desactivamos el scope
TmrScopeLite.Interval = 0
TmrScopeLite.Enabled = False
'reseteamos los displays
Lr.Width = 0
Ll.Width = 0
Picfft1.Cls

'chequeamos el cue auto
If E1Cue.Caption = "Desactivar CUE" Then
    E1Cue.Caption = "Activar CUE"
    E1Cue.BackColor = &H8000000F    'gris
    TmrCUE.Interval = 0
    TmrCUE.Enabled = False
End If

End Sub

Private Sub E1Vol_Change()

If Est12Control.StopLabel1.Caption = "Stream" And Est12Control.Origen1.Caption = "E1" Then
    'change the stream volume
    Stream01SetVolume (E1Vol.Value)
    E1Vol.ToolTipText = E1Vol.Value
    LblCurrVol.Caption = E1Vol.Value
Else
    If Est12Control.StopLabel1.Caption = "Music" And Est12Control.Origen1.Caption = "E1" Then
        'change the music volume
        Music01SetVolume (E1Vol.Value)
        E1Vol.ToolTipText = E1Vol.Value
        LblCurrVol.Caption = E1Vol.Value
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub fft10_Click()

If fft10.ForeColor = &H808000 Then   'verde oscuro
    fft10.ForeColor = &HFFFF00   'verde claro
    fft4.ForeColor = &H808000
    fft6.ForeColor = &H808000
    fft8.ForeColor = &H808000
    fft2.ForeColor = &H808000
End If

End Sub

Private Sub fft2_Click()

If fft2.ForeColor = &H808000 Then   'verde oscuro
    fft2.ForeColor = &HFFFF00   'verde claro
    fft4.ForeColor = &H808000
    fft6.ForeColor = &H808000
    fft8.ForeColor = &H808000
    fft10.ForeColor = &H808000
End If

End Sub

Private Sub fft4_Click()

If fft4.ForeColor = &H808000 Then   'verde oscuro
    fft4.ForeColor = &HFFFF00   'verde claro
    fft2.ForeColor = &H808000
    fft6.ForeColor = &H808000
    fft8.ForeColor = &H808000
    fft10.ForeColor = &H808000
End If

End Sub


Private Sub fft6_Click()

If fft6.ForeColor = &H808000 Then   'verde oscuro
    fft6.ForeColor = &HFFFF00   'verde claro
    fft4.ForeColor = &H808000
    fft2.ForeColor = &H808000
    fft8.ForeColor = &H808000
    fft10.ForeColor = &H808000
End If

End Sub


Private Sub fft8_Click()

If fft8.ForeColor = &H808000 Then   'verde oscuro
    fft8.ForeColor = &HFFFF00   'verde claro
    fft4.ForeColor = &H808000
    fft6.ForeColor = &H808000
    fft2.ForeColor = &H808000
    fft10.ForeColor = &H808000
End If

End Sub

Private Sub Form_Load()

'*** load the commands strings
E1Cue.Caption = LoadResString(2007)
E1Import.Caption = LoadResString(2006)
CmdRestore.Caption = LoadResString(2005)
CmdActualiz.Caption = LoadResString(2004)

'*** load commands pictures
    E1Pic.Picture = LoadResPicture("EST_01", 0)
    'load led1
    Picture1.Picture = LoadResPicture("BACK_LED", 0)
    Ll.Picture = LoadResPicture("FRONT_LED", 0)
    'load led2
    Picture2.Picture = LoadResPicture("BACK_LED", 0)
    Lr.Picture = LoadResPicture("FRONT_LED", 0)
    'Load control pictures
    'E1Play.PictureNormal = LoadResPicture("R_PLAY", 0)
    'E1Pause.PictureNormal = LoadResPicture("R_PAUSE", 0)
    'E1Stop.PictureNormal = LoadResPicture("R_STOP", 0)
    '--- more...
    'E1New.PictureNormal = LoadResPicture("ICO_NEW", 0)
    'E1Open.PictureNormal = LoadResPicture("ICO_OPEN", 0)
    'E1Save.PictureNormal = LoadResPicture("ICO_SAVE", 0)
    '--- and more...
    NewCUE.Picture = LoadResPicture("ICO_NEW", 0)
    OpenCUE.Picture = LoadResPicture("ICO_OPEN", 0)
    SaveCUE.Picture = LoadResPicture("ICO_SAVE", 0)
    '--- and much mooooore....
    'Image1.Picture = LoadResPicture("EST_PANEL", 0)
    'Picture3.Picture = LoadResPicture("EST_PANEL_FFT", 0)
    
    'Reset the size
    Lr.Width = 0
    Ll.Width = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

HideWindow "Est01"

End Sub

Private Sub Form_Terminate()

HideWindow "Est01"

End Sub

Private Sub Form_Unload(Cancel As Integer)

HideWindow "Est01"

End Sub

Private Sub fxsc_Change(index As Integer)

UpdateFX01 (index)

End Sub

Private Sub fxsc_Scroll(index As Integer)

UpdateFX01 (index)

End Sub

Private Sub LAplay_Click()

If LAplay.ForeColor = &H808000 Then
    LAplay.ForeColor = &HFFFF00 'claro
Else
    LAplay.ForeColor = &H808000 'oscuro
End If

End Sub

Private Sub LCue_Click()

If LCue.Caption = "Auto" Then
    LCue.Caption = "Man"
Else
    LCue.Caption = "Auto"
End If

End Sub

Private Sub Lfft_Click()

If Lfft.ForeColor = &H808000 Then 'verde oscuro
    Lfft.ForeColor = &HFFFF00   'verde claro
    Lspc.ForeColor = &H808000   'verde oscuro
    Lspcz.ForeColor = &H808000   'verde oscuro
    Lspcd.ForeColor = &H808000   'verde oscuro
    Lspcb.ForeColor = &H808000   'verde oscuro
End If

End Sub

Private Sub LFin_Click()

If LFin.Caption = "Man" Then
    LFin.Caption = "Auto"
    CmdFIN.Enabled = False
    CmdFOut.Enabled = False
Else
    LFin.Caption = "Man"
    CmdFIN.Enabled = True
    CmdFOut.Enabled = True
End If

End Sub

Private Sub Lspc_Click()

If Lspc.ForeColor = &H808000 Then 'verde oscuro
    Lspc.ForeColor = &HFFFF00   'verde claro
    Lspcb.ForeColor = &HFFFF00   'verde claro
    Lspcz.ForeColor = &H808000   'verde oscuro
    Lspcd.ForeColor = &H808000   'verde oscuro
    Lfft.ForeColor = &H808000   'verde oscuro
End If

End Sub


Private Sub Lspcb_Click()

If Lspc.ForeColor = &HFFFF00 Then 'verde claro
    Lspc.ForeColor = &HFFFF00   'verde claro
    Lspcb.ForeColor = &HFFFF00   'verde claro
    Lspcz.ForeColor = &H808000   'verde oscuro
    Lspcd.ForeColor = &H808000   'verde oscuro
    Lfft.ForeColor = &H808000   'verde oscuro
End If

End Sub

Private Sub Lspcd_Click()

If Lspc.ForeColor = &HFFFF00 Then 'verde claro
    Lspc.ForeColor = &HFFFF00   'verde claro
    Lspcd.ForeColor = &HFFFF00   'verde claro
    Lspcz.ForeColor = &H808000   'verde oscuro
    Lspcb.ForeColor = &H808000   'verde oscuro
    Lfft.ForeColor = &H808000   'verde oscuro
End If

End Sub

Private Sub Lspcz_Click()

If Lspc.ForeColor = &HFFFF00 Then 'verde claro
    Lspc.ForeColor = &HFFFF00   'verde claro
    Lspcz.ForeColor = &HFFFF00   'verde claro
    Lspcb.ForeColor = &H808000   'verde oscuro
    Lspcd.ForeColor = &H808000   'verde oscuro
    Lfft.ForeColor = &H808000   'verde oscuro
End If

End Sub

Private Sub NewCUE_Click()

CmdRestore_Click

End Sub

Private Sub OpenCUE_Click()

Dim ContNum As Integer
Dim filename As String
Dim LenFN As Long
Dim FileNTest As String
Dim NameFile As String

If Est01.Fi.Caption = "" Then MsgBox LoadResString(154): Exit Sub

ContNum = CInt(Est01.Fi.Caption)    'extraemos el index del control
filename = Trim(Est12Data.N1(ContNum).Caption)    'extraemos el path y el archivo de audio
NameFile = StripFileFromExt(filename)
filename = Trim(NameFile) & AppCUEFileExt

'guardamos la informacion CUe
OpenCUEFile 1, filename

End Sub

Private Sub P11_Click(index As Integer)

If TxtName.Visible = True Then
    TxtName.Visible = False
End If

ConvertTxT = Trim(Fn.Caption)
If ConvertTxT = "" Or ConvertTxT = " " Then Exit Sub

Select Case index
    Case 0
        Call E1Save_Click   'save the old page file
        SetDefControl 1   'set the default control´s caption
        Pn.Caption = "1"
        Call OpenE1PageFile 'Open the new page file
    Case 1
        Call E1Save_Click
        SetDefControl 1
        Pn.Caption = "2"
        Call OpenE1PageFile
    Case 2
        Call E1Save_Click
        SetDefControl 1
        Pn.Caption = "3"
        Call OpenE1PageFile
    Case 3
        Call E1Save_Click
        SetDefControl 1
        Pn.Caption = "4"
        Call OpenE1PageFile
    Case 4
        Call E1Save_Click
        SetDefControl 1
        Pn.Caption = "5"
        Call OpenE1PageFile
    Case 5
        Call E1Save_Click
        SetDefControl 1
        Pn.Caption = "6"
        Call OpenE1PageFile
    Case 6
        Call E1Save_Click
        SetDefControl 1
        Pn.Caption = "7"
        Call OpenE1PageFile
    Case 7
        Call E1Save_Click
        SetDefControl 1
        Pn.Caption = "8"
        Call OpenE1PageFile
    Case 8
        Call E1Save_Click
        SetDefControl 1
        Pn.Caption = "9"
        Call OpenE1PageFile
End Select

End Sub

Private Sub pcontdw_Click()

pcontdw.Visible = False
pcontup.Visible = True

End Sub

Private Sub pcontup_Click()

pcontdw.Visible = True
pcontup.Visible = False

End Sub

Private Sub SaveCUE_Click()

Dim ContNum As Integer
Dim filename As String
Dim LenFN As Long
Dim FileNTest As String
Dim NameFile As String

If Est01.Fi.Caption = "" Then MsgBox LoadResString(154): Exit Sub

ContNum = CInt(Est01.Fi.Caption)    'extraemos el index del control
filename = Trim(Est12Data.N1(ContNum).Caption)    'extraemos el path y el archivo de audio
NameFile = StripFileFromExt(filename)
filename = Trim(NameFile) & AppCUEFileExt

'guardamos la informacion CUe
SaveCUEFile 1, filename

End Sub

Private Sub TMin_Timer()

If E1Vol.Value = 100 Or E1Vol.Value = CLng(LblInvol.Caption) Then
    TMin.Interval = 0
    TMin.Enabled = False
Else
    E1Vol.Value = E1Vol.Value + 1
End If

End Sub

Private Sub Tmout_Timer()

If E1Vol.Value = 0 Or E1Vol.Value = CLng(LblOutvol.Caption) Then
    Tmout.Interval = 0
    Tmout.Enabled = False
Else
    E1Vol.Value = E1Vol.Value - 1
End If

End Sub

Private Sub TmOutAuto_Timer()

If E1Vol.Value = 0 Then
    If Est12Control.StopLabel1.Caption = "Stream" And Est12Control.Origen1.Caption = "E1" Then
        Stream01Restart    'stream restart
        Stream01Stop       'stream stop
    Else
        If Est12Control.StopLabel1.Caption = "Music" And Est12Control.Origen1.Caption = "E1" Then
            Music01Restart     'music restart
            Music01Stop         'music stop
        Else
            'desactivamos el scope
            TmrScopeLite.Interval = 0
            TmrScopeLite.Enabled = False
            'reseteamos los displays
            Lr.Width = 0
            Ll.Width = 0
            Picfft1.Cls
            TmoutAuto.Interval = 0
            TmoutAuto.Enabled = False
        End If
    End If
    'desactivamos el scope
    TmrScopeLite.Interval = 0
    TmrScopeLite.Enabled = False
    'reseteamos los displays
    Lr.Width = 0
    Ll.Width = 0
    Picfft1.Cls
    TmoutAuto.Interval = 0
    TmoutAuto.Enabled = False
Else
    E1Vol.Value = E1Vol.Value - 1
End If

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
    E1Pos.ToolTipText = ConvSecToMin(CInt(E1Pos.Value))
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

Private Sub TmrScopeLite_Timer()

'funciona tanto como para la Estacion01 como para la Tanda01
Dim LLft
Dim RRgt
Dim SType As String

If Est12Control.StopLabel1 = "Stream" Then
    If Est12Control.Origen1.Caption = "E1" Then
        LLft = Stream01GetLEFTLevel
        RRgt = Stream01GetRIGHTLevel
        Est01.SetAudioLevel LLft, RRgt
        SType = "Stream"
    End If
End If

If Est12Control.StopLabel1 = "Music" Then
    If Est12Control.Origen1.Caption = "E1" Then
        LLft = Music01GetLEFTLevel
        RRgt = Music01GetRIGHTLevel
        Est01.SetAudioLevel LLft, RRgt
        SType = "Music"
    End If
End If

'chequeamos por el tipo de display en est01
If Lfft.ForeColor = &HFFFF00 Then 'verde claro
    If fft2.ForeColor = &HFFFF00 Then   'verde claro
        Call DrawFFT(1, SType, 2) 'fft spectrum display
    Else
        If fft4.ForeColor = &HFFFF00 Then   'verde claro
            Call DrawFFT(1, SType, 4) 'fft spectrum display
        Else
            If fft6.ForeColor = &HFFFF00 Then   'verde claro
                Call DrawFFT(1, SType, 6) 'fft spectrum display
            Else
                If fft8.ForeColor = &HFFFF00 Then   'verde claro
                    Call DrawFFT(1, SType, 8) 'fft spectrum display
                Else
                    If fft10.ForeColor = &HFFFF00 Then   'verde claro
                        Call DrawFFT(1, SType, 10) 'fft spectrum display
                    Else
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
Else
    If Lspc.ForeColor = &HFFFF00 Then
        If Lspcz.ForeColor = &HFFFF00 Then  'scope izquiero
            Call DrawScope(&HFFFF00, &H808000, 5, 0, 130, 50, 1, SType, ScopeSideBySide)
        End If
        If Lspcd.ForeColor = &HFFFF00 Then  'scope derecho
            Call DrawScope(&H808000, &HFFFF00, 5, 0, 130, 50, 1, SType, ScopeSideBySide)
        End If
        If Lspcb.ForeColor = &HFFFF00 Then  'scope dual
            Call DrawScope(&HFFFF00, &HFFFF00, 5, 0, 130, 50, 1, SType, ScopeDouble)
        End If
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)

Dim IDX As Integer

If KeyAscii = 13 Then   'ENTER
    IDX = CInt(Lindex.Caption)
    E11(IDX).Caption = TxtName.text
    Est12Data.c1(IDX).Caption = TxtName.text
    TxtName.Visible = False
End If
If KeyAscii = 27 Or KeyAscii = 13 Then 'ESCAPE or ENTER
    TxtName.Visible = False
End If

End Sub
