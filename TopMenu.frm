VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form TopMenu 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Only RadioMaker"
   ClientHeight    =   9435
   ClientLeft      =   330
   ClientTop       =   375
   ClientWidth     =   15165
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "TopMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   15165
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox c0 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   14790
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   55
      Top             =   900
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox c5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   13440
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   54
      Top             =   900
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox c9 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   14520
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   53
      Top             =   900
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox c8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   14250
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   52
      Top             =   900
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox c7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   13980
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   51
      Top             =   900
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox c6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   13710
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   50
      Top             =   900
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox c4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   13170
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   49
      Top             =   900
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox c3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   12900
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   48
      Top             =   900
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox c2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   12630
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   47
      Top             =   900
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox c1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   12360
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   46
      Top             =   900
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer TimerRMVoiceCheck 
      Left            =   10020
      Top             =   870
   End
   Begin VB.Timer ProcTimer 
      Left            =   7560
      Top             =   4890
   End
   Begin VB.Timer TmrSendPos 
      Left            =   10530
      Top             =   870
   End
   Begin MSComDlg.CommonDialog NTSCmd 
      Left            =   2430
      Top             =   5895
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox PrgT8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   9450
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   41
      Top             =   435
      Width           =   190
   End
   Begin VB.PictureBox PrgT7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   9270
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   40
      Top             =   435
      Width           =   190
   End
   Begin VB.PictureBox PrgT6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   9090
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   39
      Top             =   435
      Width           =   190
   End
   Begin VB.PictureBox PrgT5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   8910
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   38
      Top             =   435
      Width           =   190
   End
   Begin VB.PictureBox PrgT4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   8715
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   37
      Top             =   435
      Width           =   190
   End
   Begin VB.PictureBox PrgT3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   8520
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   36
      Top             =   435
      Width           =   190
   End
   Begin VB.PictureBox PrgT2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   8340
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   35
      Top             =   435
      Width           =   190
   End
   Begin VB.PictureBox PrgT1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   8145
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   34
      Top             =   435
      Width           =   190
   End
   Begin VB.PictureBox Pht1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   8145
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   33
      Top             =   165
      Width           =   190
   End
   Begin VB.PictureBox Pht2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   8340
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   32
      Top             =   165
      Width           =   190
   End
   Begin VB.PictureBox Pht3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   8520
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   31
      Top             =   165
      Width           =   190
   End
   Begin VB.PictureBox Pht4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   8715
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   30
      Top             =   165
      Width           =   190
   End
   Begin VB.PictureBox Pht5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   8910
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   29
      Top             =   165
      Width           =   190
   End
   Begin VB.PictureBox Pht6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   9090
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   28
      Top             =   165
      Width           =   190
   End
   Begin VB.PictureBox Pht7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   9270
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   27
      Top             =   165
      Width           =   190
   End
   Begin VB.PictureBox Pht8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   9450
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   26
      Top             =   165
      Width           =   190
   End
   Begin MSComDlg.CommonDialog WaveCmd 
      Left            =   3060
      Top             =   5910
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog PHCmd 
      Left            =   4185
      Top             =   5265
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog ProgCmd 
      Left            =   3600
      Top             =   5265
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog TandaCmd 
      Left            =   3015
      Top             =   5265
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox t8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   11745
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   255
      Width           =   255
   End
   Begin VB.PictureBox t7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   11505
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   255
      Width           =   255
   End
   Begin VB.PictureBox t6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   11265
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   16
      Top             =   255
      Width           =   255
   End
   Begin VB.PictureBox t5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   11025
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   255
      Width           =   255
   End
   Begin VB.PictureBox t4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   10785
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   255
      Width           =   255
   End
   Begin VB.PictureBox t3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   10545
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   255
      Width           =   255
   End
   Begin VB.PictureBox t2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   10305
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   255
      Width           =   255
   End
   Begin VB.PictureBox t1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   10065
      ScaleHeight     =   300
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   255
      Width           =   243
   End
   Begin VB.PictureBox f1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   12480
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   10
      Top             =   270
      Width           =   255
   End
   Begin VB.PictureBox f2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   12750
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   9
      Top             =   270
      Width           =   255
   End
   Begin VB.PictureBox f3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   12990
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   8
      Top             =   270
      Width           =   255
   End
   Begin VB.PictureBox f4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   13230
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   270
      Width           =   255
   End
   Begin VB.PictureBox f5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   13470
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   270
      Width           =   255
   End
   Begin VB.PictureBox f6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   13710
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   270
      Width           =   255
   End
   Begin VB.PictureBox f7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   13950
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   270
      Width           =   255
   End
   Begin VB.PictureBox f8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   14190
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   270
      Width           =   255
   End
   Begin VB.PictureBox f9 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   14430
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   270
      Width           =   255
   End
   Begin VB.PictureBox f10 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   14670
      ScaleHeight     =   300
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   270
      Width           =   255
   End
   Begin MSComDlg.CommonDialog EstCmd 
      Left            =   2430
      Top             =   5265
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer ClockTimer 
      Left            =   7050
      Top             =   4890
   End
   Begin PicClip.PictureClip BigClip 
      Left            =   5850
      Top             =   5790
      _ExtentX        =   5186
      _ExtentY        =   476
      _Version        =   393216
      Cols            =   14
      Picture         =   "TopMenu.frx":08CA
   End
   Begin PicClip.PictureClip SmallClip 
      Left            =   6270
      Top             =   5550
      _ExtentX        =   3731
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   14
      Picture         =   "TopMenu.frx":3274
   End
   Begin VB.PictureBox PicSmall 
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   30
      ScaleHeight     =   600
      ScaleWidth      =   2400
      TabIndex        =   0
      Top             =   90
      Width           =   2460
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   2970
      Top             =   4410
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog BlockCmd 
      Left            =   3630
      Top             =   5910
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin PicClip.PictureClip TempClip 
      Left            =   5850
      Top             =   6150
      _ExtentX        =   6297
      _ExtentY        =   476
      _Version        =   393216
      Cols            =   17
      Picture         =   "TopMenu.frx":484E
   End
   Begin VB.Label SType 
      BackColor       =   &H000080FF&
      Caption         =   "Normal"
      Height          =   240
      Left            =   630
      TabIndex        =   45
      Top             =   1980
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LType 
      BackColor       =   &H000080FF&
      Caption         =   "Normal"
      Height          =   240
      Left            =   630
      TabIndex        =   44
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label OType 
      BackColor       =   &H000080FF&
      Caption         =   "Normal"
      Height          =   240
      Left            =   630
      TabIndex        =   43
      Top             =   1710
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label PRGName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   5490
      TabIndex        =   42
      Top             =   390
      Width           =   2625
   End
   Begin VB.Label NumberIdx 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8730
      TabIndex        =   25
      Top             =   1020
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label PHName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   5490
      TabIndex        =   24
      Top             =   165
      Width           =   2625
   End
   Begin VB.Label PHTime 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9045
      TabIndex        =   23
      Top             =   1020
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   5400
      Picture         =   "TopMenu.frx":7AF8
      Stretch         =   -1  'True
      Top             =   30
      Width           =   4380
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Desactivada"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   4320
      TabIndex        =   22
      Top             =   390
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Program. de Tandas:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   2655
      TabIndex        =   21
      Top             =   390
      Width           =   1590
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Desactivada"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   4320
      TabIndex        =   20
      Top             =   165
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Program. Horaria:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   2835
      TabIndex        =   19
      Top             =   165
      Width           =   1410
   End
   Begin VB.Image TopPic 
      Height          =   735
      Left            =   2490
      Picture         =   "TopMenu.frx":9368
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2985
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   9810
      Picture         =   "TopMenu.frx":ABD8
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   12240
      Picture         =   "TopMenu.frx":C448
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2895
   End
   Begin VB.Menu RmMenu 
      Caption         =   "&RadioMaker"
      Begin VB.Menu RMConfig 
         Caption         =   "&Configuración..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu RMSep 
         Caption         =   "-"
      End
      Begin VB.Menu RMExit 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Empty 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu EstacionMnu 
      Caption         =   "&Estaciones"
      Begin VB.Menu EstacionMnuEst1 
         Caption         =   "Estación 01"
         Begin VB.Menu Est1New 
            Caption         =   "Nueva compaginación"
         End
         Begin VB.Menu Est1Sep0 
            Caption         =   "-"
         End
         Begin VB.Menu Est1Open 
            Caption         =   "Abrir..."
         End
         Begin VB.Menu Est1Sep1 
            Caption         =   "-"
         End
         Begin VB.Menu Est1Save 
            Caption         =   "Guardar"
         End
         Begin VB.Menu Est1SaveAs 
            Caption         =   "Guardar &como..."
         End
      End
      Begin VB.Menu EstacionMnuEst2 
         Caption         =   "Estación 02"
         Begin VB.Menu Est2New 
            Caption         =   "Nueva compaginación"
         End
         Begin VB.Menu Est2Sep0 
            Caption         =   "-"
         End
         Begin VB.Menu Est2Open 
            Caption         =   "Abrir..."
         End
         Begin VB.Menu Est2Sep1 
            Caption         =   "-"
         End
         Begin VB.Menu Est2Save 
            Caption         =   "Guardar"
         End
         Begin VB.Menu Est2SaveAs 
            Caption         =   "Guardar como..."
         End
      End
   End
   Begin VB.Menu Empty0 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu TandasMnu 
      Caption         =   "&Tandas"
      Begin VB.Menu TndNueva 
         Caption         =   "&Nueva Tanda"
         Shortcut        =   ^N
      End
      Begin VB.Menu TndSep1 
         Caption         =   "-"
      End
      Begin VB.Menu TndAbrir 
         Caption         =   "&Abrir..."
         Shortcut        =   ^A
      End
      Begin VB.Menu TndSep2 
         Caption         =   "-"
      End
      Begin VB.Menu TndGuardar 
         Caption         =   "&Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu TndGuardarComo 
         Caption         =   "Guardar &como..."
      End
      Begin VB.Menu TndSep3 
         Caption         =   "-"
      End
      Begin VB.Menu TndAudioFileSubMenu 
         Caption         =   "Ar&chivos de Audio"
         Begin VB.Menu TndAudioProp 
            Caption         =   "&Propiedades..."
            Shortcut        =   ^{INSERT}
         End
         Begin VB.Menu SepAudio1 
            Caption         =   "-"
         End
         Begin VB.Menu TndAudioDelete 
            Caption         =   "&Eliminar"
            Shortcut        =   {DEL}
         End
      End
      Begin VB.Menu TndBlockFileSubMenu 
         Caption         =   "&Bloques publicitarios"
         Begin VB.Menu TndBlkConfig 
            Caption         =   "Configurar..."
         End
      End
   End
   Begin VB.Menu Empty1 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu TndReporteSubMenu 
      Caption         =   "&Reportes"
      Begin VB.Menu TndReporteEditar 
         Caption         =   "&Visualizar / Editar..."
         Shortcut        =   ^R
      End
      Begin VB.Menu TndReportePrint 
         Caption         =   "&Imprimir..."
         Shortcut        =   ^P
      End
      Begin VB.Menu TndReportSep1 
         Caption         =   "-"
      End
      Begin VB.Menu TndReportDel 
         Caption         =   "&Eliminar..."
      End
   End
   Begin VB.Menu Empty2 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu PlugInsmnu 
      Caption         =   "&Plug-Ins"
      Begin VB.Menu PlugRMPlug 
         Caption         =   "RMPlug-Ins"
         Begin VB.Menu PluginExplorer 
            Caption         =   "RM Explorer..."
         End
         Begin VB.Menu PlugInMiniPlayer 
            Caption         =   "RM Mini player..."
         End
         Begin VB.Menu PlugInCDRipper 
            Caption         =   "RM CD-Ripper..."
         End
         Begin VB.Menu PlugInNataly 
            Caption         =   "RM Voice..."
         End
         Begin VB.Menu PlugInController 
            Caption         =   "RM Controlador..."
         End
         Begin VB.Menu PlugInFXModule 
            Caption         =   "RM modulo FX..."
         End
         Begin VB.Menu PlugInFilter 
            Caption         =   "RM modulo de Filtros..."
         End
         Begin VB.Menu PlugInEditec 
            Caption         =   "RM Edición de audio..."
         End
         Begin VB.Menu PlugInDatabase 
            Caption         =   "RM Base de Datos..."
         End
      End
   End
   Begin VB.Menu Empty3 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu VentanasMnu 
      Caption         =   "&Ventanas"
      Begin VB.Menu SbOrderAuto 
         Caption         =   "&Ordenar Automaticamente"
      End
      Begin VB.Menu Sep01 
         Caption         =   "-"
      End
      Begin VB.Menu SbEst01 
         Caption         =   "Estacion 0&1"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu SbEst02 
         Caption         =   "Estacion 0&2"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu Sep02 
         Caption         =   "-"
      End
      Begin VB.Menu SbTnd01 
         Caption         =   "Creacion de &Tandas"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu SbPrg01 
         Caption         =   "&Programacion de Tandas"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu Sep04 
         Caption         =   "-"
      End
      Begin VB.Menu SbExplor 
         Caption         =   "E&xplotador de Archivos"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu Sep05 
         Caption         =   "-"
      End
      Begin VB.Menu SbHerram 
         Caption         =   "&Herramientas"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu Sep06 
         Caption         =   "-"
      End
      Begin VB.Menu SbView 
         Caption         =   "&Vistas"
         Begin VB.Menu View3x3 
            Caption         =   "3x3"
         End
         Begin VB.Menu View4x4v 
            Caption         =   "4x4 Vertical"
         End
         Begin VB.Menu View4x4h 
            Caption         =   "4x4 Horizontal"
         End
         Begin VB.Menu ViewSep1 
            Caption         =   "-"
         End
         Begin VB.Menu ViewDefault 
            Caption         =   "Por Defecto"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu Empty4 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu AyudaMnu 
      Caption         =   "&Ayuda"
      Begin VB.Menu Ayuda_Cmd 
         Caption         =   "&Contenido..."
      End
      Begin VB.Menu Indice_cmd 
         Caption         =   "&Indice..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu AyudaSep 
         Caption         =   "-"
      End
      Begin VB.Menu RMWeb 
         Caption         =   "RadioMaker en la WEB"
         Begin VB.Menu RMWeb_Help 
            Caption         =   "Ayuda en línea..."
         End
         Begin VB.Menu RMWeb_Soporte 
            Caption         =   "Soporte técnico..."
         End
         Begin VB.Menu RMWeb_sep 
            Caption         =   "-"
         End
         Begin VB.Menu RMWeb_Novedades 
            Caption         =   "Novedades..."
         End
         Begin VB.Menu RMWeb_Actualiza 
            Caption         =   "Actualizaciones..."
         End
      End
      Begin VB.Menu Rm100Condiciones 
         Caption         =   "Condiciones de &licencia"
      End
      Begin VB.Menu AyudaSep2 
         Caption         =   "-"
      End
      Begin VB.Menu AcercaRM100_cmd 
         Caption         =   "&Acerca RadioMaker..."
      End
   End
End
Attribute VB_Name = "TopMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'************************************************
'ONLY Radiomaker TopMenu main proccess code
'Copyright (c) 1987-2008 ONLY Development inc.
'************************************************

'------------------------------------------------
'ultimos cambios realizados 07-12-2008
'------------------------------------------------

'-------- OTHER DIM
'Dim EndingFlag As Boolean
'Dimensiones de resultado
'Dim RResult As String

Dim RealClock As String
Dim RealDate As String
'Dim RealTemp As String

'Dimensiones de resultado de sonido digital
'Dim ResultDev As String
Dim Result As String
Dim NewResult As Boolean

'dimensiones para el manejo de contadores de fecha y temperatura
Dim a As Integer, B As Integer

'PlugIn DLL object dim
Public RMPlugIn As Object

Public Function LoadPlugIn(WFPlugInName As String, WAction As String) As String

Nuevamente:
On Error GoTo ErrorInShell

    Set RMPlugIn = CreateObject(WFPlugInName & ".AddonClass")
    If RMPlugIn.LoadControl("P•@`wœÔ6©>©™ø=Ø/1Ð0pâ:²", WAction) = 0 Then
        MsgBox "Error al intentar cargar el AddOn. Consulte a su proveedor de software", vbCritical, WFPlugInName & " - LoadControl ERROR!"
        Set RMPlugIn = Nothing
        LoadPlugIn = "NotOk"
        Exit Function
    Else
        LoadPlugIn = "Ok"
    End If
    
Exit Function

ErrorInShell:
'err 429=objeto no registrado
Set RMPlugIn = Nothing
If err.Number = 429 Then
    NewResult = RegUNRegLib(WFPlugInName, 1)
    If NewResult = True Then
        'Resume Nuevamente
    Else
        MsgBox "PlugIn Loader Error N429 - " & WFPlugInName
        LoadPlugIn = "NotOk"
    End If
Else
    MsgBox "PlugIn Loader Error N" & err.Number & " - " & WFPlugInName
    LoadPlugIn = "NotOk"
End If

End Function

Public Function UnloadPlugIn(WFPlugInName As String) As String

'-------------------
'revisar 22-12-2007
'-------------------

On Error GoTo ErrorInShell

Select Case WFPlugInName
    Case "RMXplorer.dll"
        RMPlugIn.unloadcontrol
        Set RMPlugIn = Nothing
        
    Case "RMPlayer.dll"
        RMPlugIn.unloadcontrol
        Set RMPlugIn = Nothing
        
    Case "RMRipper.dll"
        RMPlugIn.unloadcontrol
        Set RMPlugIn = Nothing
        
    Case "RMVoice.dll"
        RMPlugIn.unloadcontrol
        Set RMPlugIn = Nothing
        
    Case "RMController.dll"
        RMPlugIn.unloadcontrol
        Set RMPlugIn = Nothing
        
    Case "RMXModule.dll"
        RMPlugIn.unloadcontrol
        Set RMPlugIn = Nothing
        
    Case "RMFilter.dll"
        RMPlugIn.unloadcontrol
        Set RMPlugIn = Nothing
        
    Case "RMEditec.dll"
        RMPlugIn.unloadcontrol
        Set RMPlugIn = Nothing
        
    Case "RMDatabase.dll"
        RMPlugIn.unloadcontrol
        Set RMPlugIn = Nothing

    Case Else
        RMPlugIn.unloadcontrol
        Set RMPlugIn = Nothing
End Select

    UnloadPlugIn = "Ok"
Exit Function

ErrorInShell:
    UnloadPlugIn = "NotOk"
End Function

Public Function GetPlugInList(WPlugInName As String) As String

'-------------------
'revisar 22-12-2007
'-------------------

Dim Directory As String
Dim TxtChK As String

'chequeos
TxtChK = Trim(WPlugInName)
If TxtChK = "" Or TxtChK = " " Then
    GetPlugInList = "NotOk"
End If

'seteos necesarios
Directory = App.Path & AppPlugInDir & "\" & WPlugInName

If FileExist(Directory) = True Then
    Select Case WPlugInName
        Case "RMXplorer.dll"
            TopMenu.PluginExplorer.Visible = True
            TopMenu.PluginExplorer.Enabled = True
        Case "RMPlayer.dll"
            TopMenu.PlugInMiniPlayer.Visible = True
            TopMenu.PlugInMiniPlayer.Enabled = True
        Case "RMRipper.dll"
            TopMenu.PlugInCDRipper.Visible = True
            TopMenu.PlugInCDRipper.Enabled = True
        Case "RMVoice.dll"
            TopMenu.PlugInNataly.Visible = True
            TopMenu.PlugInNataly.Enabled = True
        Case "RMController.dll"
            TopMenu.PlugInController.Visible = True
            TopMenu.PlugInController.Enabled = True
        Case "RMXModule.dll"
            TopMenu.PlugInFXModule.Visible = True
            TopMenu.PlugInFXModule.Enabled = True
        Case "RMFilter.dll"
            TopMenu.PlugInFilter.Visible = True
            TopMenu.PlugInFilter.Enabled = True
        Case "RMEditec.dll"
            TopMenu.PlugInEditec.Visible = True
            TopMenu.PlugInEditec.Enabled = True
        Case "RMDatabase.dll"
            TopMenu.PlugInDatabase.Visible = True
            TopMenu.PlugInDatabase.Enabled = True
        Case Else
            'xxx not implemented yet...
    End Select
Else
    Select Case WPlugInName
        Case "RMXplorer.dll"
            TopMenu.PluginExplorer.Visible = False
            TopMenu.PluginExplorer.Enabled = False
        Case "RMPlayer.dll"
            TopMenu.PlugInMiniPlayer.Visible = False
            TopMenu.PlugInMiniPlayer.Enabled = False
        Case "RMRipper.dll"
            TopMenu.PlugInCDRipper.Visible = False
            TopMenu.PlugInCDRipper.Enabled = False
        Case "RMVoice.dll"
            TopMenu.PlugInNataly.Visible = False
            TopMenu.PlugInNataly.Enabled = False
        Case "RMController.dll"
            TopMenu.PlugInController.Visible = False
            TopMenu.PlugInController.Enabled = False
        Case "RMXModule.dll"
            TopMenu.PlugInFXModule.Visible = False
            TopMenu.PlugInFXModule.Enabled = False
        Case "RMFilter.dll"
            TopMenu.PlugInFilter.Visible = False
            TopMenu.PlugInFilter.Enabled = False
        Case "RMEditec.dll"
            TopMenu.PlugInEditec.Visible = False
            TopMenu.PlugInEditec.Enabled = False
        Case "RMDatabase.dll"
            TopMenu.PlugInDatabase.Visible = False
            TopMenu.PlugInDatabase.Enabled = False
        Case Else
            'xxx not implemented yet...
    End Select
End If

GetPlugInList = "Ok"

End Function

Sub GetNextStreamTime()

'This module is intented only for RMPlayer.dll plug-in
'cual es el proximo tema a reproducir?
'only for Tanda streams

Dim nIndex As Integer
Dim NextLanz As String

On Error GoTo err
nIndex = Tanda01.T1View.SelectedItem.Index   'numero de index
nIndex = nIndex + 1
Tanda01.T1View.ListItems.Item(nIndex).Selected = True

'extraemos los datos del del tema
'NextLanz = Tanda01.T1View.SelectedItem.ListSubItems(4).Text     'hora de lanz

'volvemos a la posicion donde se encontraba el cursor
nIndex = nIndex - 1
Tanda01.T1View.ListItems.Item(nIndex).Selected = True

'actualizamos los datos de los demas temas
RMPlugIn.SetNextTime Trim(NextLanz)
Exit Sub

err:
RMPlugIn.SetNextTime "00:00:00"

End Sub

Sub EndApp()

Dim Cnt1 As String
Dim Cnt2 As String
Dim I As Integer

'--------------------------------------------------------
'--------------------------------------------------------
'chequeamos los cambios de datos en el programa
'*** ESTACION 01
If Est01.Fn.Caption = "" Or Est01.Fn.Caption = " " Then
    For I = 0 To 21
        If Not Est01.E11(I).Caption = "" Then
            Call SaveChanges("EST1")
            Exit For
        End If
    Next I
End If
'*** ESTACION 02
If Est02.Fn.Caption = "" Or Est02.Fn.Caption = " " Then
    For I = 0 To 21
        If Not Est02.E21(I).Caption = "" Then
            Call SaveChanges("EST2")
            Exit For
        End If
    Next I
End If
'*** TANDAS
If Tanda01.Fn.Caption = "" Or Tanda01.Caption = " " Then
    If Tanda01.T1View.ListItems.count >= 1 Then
        Call SaveChanges("TANDA")
    End If
End If
'*** PROGRAMACION DE TANDAS
If Prg01.Fn.Caption = "" Or Prg01.Fn.Caption = " " Then
    For I = 0 To 23
        If Not Prg01.Prg1(I).Caption = "" Then
            Call SaveChanges("PROGTANDA")
            Exit For
        End If
    Next I
End If

'--------------------------------------------------------
'--------------------------------------------------------
'desabilitamos el clock del topmenu
ClockTimer.Interval = 0
ClockTimer.Enabled = False

'--------------------------------------------------------
'--------------------------------------------------------
'chequeos varios
If Est12Control.StopLabel1.Caption = "Stream" Then
    Cnt1 = "Stream"
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        Cnt1 = "Music"
    Else
        Cnt1 = "None"
    End If
End If
If Est12Control.StopLabel2.Caption = "Stream" Then
    Cnt2 = "Stream"
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        Cnt2 = "Music"
    Else
        Cnt2 = "None"
    End If
End If

'--------------------------------------------------------
'--------------------------------------------------------
'Desactivamos el sistema de sonido
CloseDevice Cnt1, Cnt2

'--------------------------------------------------------
'--------------------------------------------------------
'guardamos los datos del programa antes de finalizar
PutState

'--------------------------------------------------------
'--------------------------------------------------------
'finalizamos la aplicacion
End

End Sub

Private Sub AcercaRM100_cmd_Click()

Acerca.Show 1, Me

End Sub

Private Sub c0_Click()

Call c1_Click

End Sub

Private Sub c1_Click()

Dim Tvol As Long

Result = LoadPlugIn("RMVoice", "LoadSilent")
If Result = "NotOk" Then
    MsgBox "Not loaded. rmvoice"
Else
    If Stream01IsPlaying = True Then
        Tvol = CLng(Est01.LblCurrVol.Caption) / 4
        Est01.LblOutvol.Caption = Tvol
        Est01.LblInvol.Caption = Est01.LblCurrVol.Caption
        Est01.Tmout.Enabled = True
        Est01.Tmout.Interval = 5
    Else
        If Stream02IsPlaying = True Then
            Tvol = CLng(Est02.LblCurrVol.Caption) / 4
            Est02.LblOutvol.Caption = Tvol
            Est02.LblInvol.Caption = Est02.LblCurrVol.Caption
            Est02.Tmout.Enabled = True
            Est02.Tmout.Interval = 5
        End If
    End If
    Result = LoadPlugIn("RMVoice", "SayTemperatura")
    'RMPlugIn.ExecuteCommand ("SayHora")
    TimerRMVoiceCheck.Enabled = True
    TimerRMVoiceCheck.Interval = 1000
End If

End Sub

Private Sub c2_Click()

Call c1_Click

End Sub

Private Sub c3_Click()

Call c1_Click

End Sub

Private Sub c4_Click()

Call c1_Click

End Sub

Private Sub c5_Click()

Call c1_Click

End Sub

Private Sub c6_Click()

Call c1_Click

End Sub

Private Sub c7_Click()

Call c1_Click

End Sub

Private Sub c8_Click()

Call c1_Click

End Sub

Private Sub c9_Click()

Call c1_Click

End Sub

Private Sub ClockTimer_Timer()

'---------------------
'revisado 07-12-2008
'---------------------

'chequeamos la hora del sistema y visualizamos
RealClock = time$
Call TopClock(RealClock)

'chequeamos el contador para mostrar la temperatura y humedad o la fecha del sistema
If a = 10 Then
    If B = 10 Then
        'desabilitamos el tiempo y mostramos la fecha nuevamente
        ClimaDisplay -1
        DateDisplay 1
        RealDate = Date$
        Call TopDate(RealDate)
        a = 0: B = 0
    Else
    'cargamos el plugin para evitar problemas
        'Result = LoadPlugIn("RMVoice", "LoadSilent")
        'If Result = "NotOk" Then
            'xxx nada
        'Else
            'mostramos los datos del tiempo
            'If RMPlugIn.wtemperature = "N/A" Then
                'Call TopClima("000°C", "00%")
                'nada error
            'Else
                'ClimaDisplay 1
                'DateDisplay -1
                'Call TopClima(RMPlugIn.wtemperature & "°C", RMPlugIn.whumedad & "%")
            'End If
            B = B + 1
        'End If
    End If
Else
    'desabilitamos el tiempo y mostramos la fecha nuevamente
    ClimaDisplay -1
    DateDisplay 1
    RealDate = Date$
    Call TopDate(RealDate)
    a = a + 1
End If

'chequeamos por el estado de los Plug-Ins cargados
On Error GoTo Continue
Select Case RMPlugIn.DLLName
    '/////////////////////////////////////// RmPlayer.dll
    Case "RMPlayer"
        If RMPlugIn.GetPlugInState = 0 Then
            TmrSendPos.Interval = 0
            TmrSendPos.Enabled = False
            Result = UnloadPlugIn("RMPlayer.dll")
            Set RMPlugIn = Nothing
        End If
    '/////////////////////////////////////// RmXplorer.dll
    Case "RMXplorer"
        If RMPlugIn.GetPlugInState = 0 Then
            Result = UnloadPlugIn("RMXplorer.dll")
            Set RMPlugIn = Nothing
        End If
    '/////////////////////////////////////// RmRipper.dll
    Case "RMRipper"
        If RMPlugIn.GetPlugInState = 0 Then
            Result = UnloadPlugIn("RMRipper.dll")
            Set RMPlugIn = Nothing
        End If
    '/////////////////////////////////////// RmNataly.dll
    Case "RMVoice"
        If RMPlugIn.GetPlugInState = 0 Then
            Result = UnloadPlugIn("RMVoice.dll")
            Set RMPlugIn = Nothing
        End If
    '/////////////////////////////////////// RmController.dll
    Case "RMController"
        If RMPlugIn.GetPlugInState = 0 Then
            Result = UnloadPlugIn("RMController.dll")
            Set RMPlugIn = Nothing
        End If
    '/////////////////////////////////////// RmXModule.dll
    Case "RMXModule"
        If RMPlugIn.GetPlugInState = 0 Then
            Result = UnloadPlugIn("RMXModule.dll")
            Set RMPlugIn = Nothing
        End If
    '/////////////////////////////////////// RmFilter.dll
    Case "RMFilter"
        If RMPlugIn.GetPlugInState = 0 Then
            Result = UnloadPlugIn("RMFilter.dll")
            Set RMPlugIn = Nothing
        End If
    '/////////////////////////////////////// RmEditec.dll
    Case "RMEditec"
        If RMPlugIn.GetPlugInState = 0 Then
            Result = UnloadPlugIn("RMEditec.dll")
            Set RMPlugIn = Nothing
        End If
    '/////////////////////////////////////// RmDatabase.dll
    Case "RMDatabase"
        If RMPlugIn.GetPlugInState = 0 Then
            Result = UnloadPlugIn("RMDatabase.dll")
            Set RMPlugIn = Nothing
        End If
    '/////////////////////////////////////// another plug in .dll
    Case Else
        If RMPlugIn.GetPlugInState = 0 Then
            Result = UnloadPlugIn(RMPlugIn.PlugInName)
            Set RMPlugIn = Nothing
        End If
End Select
Exit Sub

Continue:
End Sub

Private Sub f1_Click()

Dim Tvol As Long

Result = LoadPlugIn("RMVoice", "LoadSilent")
If Result = "NotOk" Then
    MsgBox "Not loaded. rmvoice"
Else
    If Stream01IsPlaying = True Then
        Tvol = CLng(Est01.LblCurrVol.Caption) / 4
        Est01.LblOutvol.Caption = Tvol
        Est01.LblInvol.Caption = Est01.LblCurrVol.Caption
        Est01.Tmout.Enabled = True
        Est01.Tmout.Interval = 5
    Else
        If Stream02IsPlaying = True Then
            Tvol = CLng(Est02.LblCurrVol.Caption) / 4
            Est02.LblOutvol.Caption = Tvol
            Est02.LblInvol.Caption = Est02.LblCurrVol.Caption
            Est02.Tmout.Enabled = True
            Est02.Tmout.Interval = 5
        End If
    End If
    Result = LoadPlugIn("RMVoice", "SayTemperatura")
    'RMPlugIn.ExecuteCommand ("SayHora")
    TimerRMVoiceCheck.Enabled = True
    TimerRMVoiceCheck.Interval = 1000
End If

End Sub

Private Sub f10_Click()

Call f1_Click

End Sub

Private Sub f2_Click()

Call f1_Click

End Sub

Private Sub f3_Click()

Call f1_Click

End Sub

Private Sub f4_Click()

Call f1_Click

End Sub

Private Sub f5_Click()

Call f1_Click

End Sub

Private Sub f6_Click()

Call f1_Click

End Sub

Private Sub f7_Click()

Call f1_Click

End Sub

Private Sub f8_Click()

Call f1_Click

End Sub

Private Sub f9_Click()

Call f1_Click

End Sub

Private Sub Form_Load()

PicSmall.Picture = LoadResPicture("RM_SMALL", 0)

    On Error Resume Next
    '///Seteamos la prioridad del programa para evitar deshabilitaciones
    'this hides our program from the XP,NT,2k endtask list
    'App.Title = ""
    
    'this remove our program from the windows 9x endtask list
    'Dim process As Long
    'process = GetCurrentProcessId()
    'Call RegisterServiceProcess(process, RSP_SIMPLE_SERVICE)
    
    'set it to realtime priority so windows endtask won't allow closing it
    Call SetPriorityClass(GetCurrentProcess(), HIGH_PRIORITY_CLASS)

Me.Caption = "Only RadioMaker v." & App.Major & "." & App.Minor & " - Rev." & App.Revision

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call EndApp

End Sub

Private Sub Form_Resize()

'gets the config device data
ConfigData = OpenConfigFile

On Error Resume Next

Select Case ConfigData.Aud_Show_MiniRM      'mostrar MINI-RM? (RmPlayer.dll)
    Case 1  'si
        If TopMenu.WindowState = 0 Then
            ShowWindow "All"
            Call CheckForTimers(1)
            TmrSendPos.Interval = 0
            TmrSendPos.Enabled = False
            Result = UnloadPlugIn("RMPlayer.dll")
        Else
            Call CheckForTimers(0)
            HideWindow "All"
            Result = LoadPlugIn("RMPlayer", "ShowMain")
            If Result = "NotOk" Then
                'xxxxx
                Exit Sub
            Else
                TmrSendPos.Enabled = True
                TmrSendPos.Interval = 25
            End If
        End If
        
    Case Else   'no
        If TopMenu.WindowState = 0 Then
            ShowWindow "All"
            Call CheckForTimers(1)
        Else
            Call CheckForTimers(0)
            HideWindow "All"
        End If
End Select

End Sub

Private Sub Form_Terminate()

Call EndApp

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call EndApp

End Sub

Private Sub Label2_Click()

MsgBox "Opción No Implementada.", vbInformation
Exit Sub

If Label2.Caption = "Desactivada" Then
    'activamos el formulario de ph
    FrmTime.Show
    'activar la programacion horaria
    Call FrmTime.PHActive_Click
Else
    'desactivar la programacion horaria
    Call FrmTime.PHDesactive_Click
End If

End Sub

Private Sub Label4_Click()

If Label4.Caption = "Desactivada" Then
    Label4.Caption = "Activada"
    Label4.ForeColor = &HFFFF00
    'activar la programacion de Tandas
Else
    Label4.Caption = "Desactivada"
    Label4.ForeColor = &H808000
    'desactivar la programacion de Tandas
End If

End Sub

Private Sub PicSmall_Click()

Acerca.Show 1, Me

End Sub

Private Sub PlugInCDRipper_Click()

Result = LoadPlugIn("RMRipper", "ShowMain")
If Result = "NotOk" Then
    'xxx
End If

End Sub

Private Sub PlugInController_Click()

Result = LoadPlugIn("RMController", "ShowMain")
If Result = "NotOk" Then
    'xxx
End If

End Sub

Private Sub PlugInDatabase_Click()

Result = LoadPlugIn("RMDatabase", "ShowMain")
If Result = "NotOk" Then
    'xxx
End If

End Sub

Private Sub PlugInEditec_Click()

Result = LoadPlugIn("RMEditec", "ShowMain")
If Result = "NotOk" Then
    'xxx
End If

End Sub

Private Sub PluginExplorer_Click()

Result = LoadPlugIn("RMXplorer", "ShowMain")
If Result = "NotOk" Then
    'xxx
End If

End Sub

Private Sub PlugInFilter_Click()

Result = LoadPlugIn("RMFilter", "ShowMain")
If Result = "NotOk" Then
    'xxx
End If

End Sub

Private Sub PlugInFXModule_Click()

Result = LoadPlugIn("RMXModule", "ShowMain")
If Result = "NotOk" Then
    'xxx
End If

End Sub

Private Sub PlugInMiniPlayer_Click()

Result = LoadPlugIn("RMPlayer", "ShowMain")
If Result = "NotOk" Then
    'xxx
End If

End Sub

Private Sub PlugInNataly_Click()

Result = LoadPlugIn("RMVoice", "ShowMain")
If Result = "NotOk" Then
    'xxx
End If

End Sub

Private Sub ProcTimer_Timer()

Dim PosByte As Long, PosTime As Long
Dim LenByte As Long, LenTime As Long
Dim RestTime As Long
Dim Rst1 As String, Rst2 As String

If Stream01IsPlaying = True Or Music01IsPlaying = True Then
    Select Case Est12Control.Origen1.Caption
        Case "E1"   '////////////////////////////////////////////////////////////
            PosTime = Stream01GetPosition(1)  'position in time (seconds)
            LenTime = Stream01GetLen(1)       'lenght in time (seconds)
            PosByte = Stream01GetPosition(2)  'position in bytes
            LenByte = Stream01GetLen(2)       'lengh in bytes
            'chequeamos por el tipo de visualizacion (normal o restante)
            If TopMenu.LType.Caption = "Normal" Then
                Rst1 = ConvSecToMin(PosTime)
                Call SetDigClock(Rst1, 1, "Normal")
            Else
                RestTime = LenTime - PosTime
                Rst1 = ConvSecToMin(RestTime)
                Call SetDigClock(Rst1, 1, "Restante")
            End If
            If TopMenu.OType.Caption = "Normal" Then
                Rst2 = PosByte
                Call SetDigNum(Trim(Rst2), 1, "Normal")
            Else
                RestTime = LenByte - PosByte
                Rst2 = Trim(Str$(RestTime))
                Call SetDigNum(Trim(Rst2), 1, "Restante")
            End If
            'positions from advanced mode
            Est01.E1Pos.Value = PosTime
            Est01.E1Pos.ToolTipText = ConvSecToMin(CDbl(Est01.E1Pos.Value))
            Est01.LblCurrent.Caption = ConvSecToMin(PosTime)
            Est01.LblCurrByte.Caption = PosByte
        Case "T1"   '////////////////////////////////////////////////////////////
            PosTime = Stream01GetPosition(1)  'position in time
            LenTime = Stream01GetLen(1)   'lenght in time
            'chequeamos por el tipo de visualizacion (normal o restante)
            If TopMenu.LType.Caption = "Normal" Then
                Rst1 = ConvSecToMin(PosTime)
                Call SetDigClock(Rst1, 3, "Normal")
            Else
                RestTime = LenTime - PosTime
                Rst1 = ConvSecToMin(RestTime)
                Call SetDigClock(Rst1, 3, "Restante")
            End If
    End Select
End If

If Stream02IsPlaying = True Or Music02IsPlaying = True Then
    Select Case Est12Control.Origen2.Caption
        Case "E2"   '////////////////////////////////////////////////////////////
            PosTime = Stream02GetPosition(1)  'position in time (seconds)
            LenTime = Stream02GetLen(1)       'lenght in time (seconds)
            PosByte = Stream02GetPosition(2)  'position in bytes
            LenByte = Stream02GetLen(2)       'lengh in bytes
            'chequeamos por el tipo de visualizacion (normal o restante)
            If TopMenu.LType.Caption = "Normal" Then
                Rst1 = ConvSecToMin(PosTime)
                Call SetDigClock(Rst1, 2, "Normal")
            Else
                RestTime = LenTime - PosTime
                Rst1 = ConvSecToMin(RestTime)
                Call SetDigClock(Rst1, 2, "Restante")
            End If
            If TopMenu.OType.Caption = "Normal" Then
                Rst2 = PosByte
                Call SetDigNum(Trim(Rst2), 2, "Normal")
            Else
                RestTime = LenByte - PosByte
                Rst2 = Trim(Str$(RestTime))
                Call SetDigNum(Trim(Rst2), 2, "Restante")
            End If
            'positions from advanced mode
            Est02.E2Pos.Value = PosTime
            Est02.E2Pos.ToolTipText = ConvSecToMin(CDbl(Est02.E2Pos.Value))
            Est02.LblCurrent.Caption = ConvSecToMin(PosTime)
            Est02.LblCurrByte.Caption = PosByte
        Case "T2"   '////////////////////////////////////////////////////////////
            PosTime = Stream02GetPosition(1)  'position in time
            LenTime = Stream02GetLen(1)   'lenght in time
            'chequeamos por el tipo de visualizacion (normal o restante)
            If TopMenu.LType.Caption = "Normal" Then
                Rst2 = ConvSecToMin(PosTime)
                Call SetDigClock(Rst2, 4, "Normal")
            Else
                RestTime = LenTime - PosTime
                Rst2 = ConvSecToMin(RestTime)
                Call SetDigClock(Rst2, 4, "Restante")
            End If
    End Select
End If

'//////////////////////////////////////////// DESACTIVACIONES
If Stream01IsPlaying = False Then
    Select Case Est12Control.Origen1.Caption
        Case "E1"
            If Music01IsPlaying = False Then
                Est01.Caption = "ESTACION 01 - Detenido"
                RestoreDisplay 1
                Est01.Label1.ForeColor = &H808000     'celeste oscuro(desactivado)
            End If
        Case "T1"
            'desactivamos el control
            RestoreDisplay 3
            Tanda01.T1Name.ForeColor = &H808000     'celeste oscuro(desactivado)
    End Select
End If
If Stream02IsPlaying = False Then
    Select Case Est12Control.Origen2.Caption
        Case "E2"
            If Music02IsPlaying = False Then
                Est02.Caption = "ESTACION 02 - Detenido"
                RestoreDisplay 2
                Est02.Label1.ForeColor = &H808000     'celeste oscuro(desactivado)
            End If
        Case "T2"
            'desactivamos el control
            RestoreDisplay 4
            Tanda01.T2Name.ForeColor = &H808000     'celeste oscuro(desactivado)
    End Select
End If

'////////////////////////////////////////// MAS DESACTIVACIONES
If Stream01IsPlaying = False And Music01IsPlaying = False And Stream02IsPlaying = False And Music02IsPlaying = False Then
    ProcTimer.Interval = 0
    ProcTimer.Enabled = False
End If

End Sub

Private Sub RMConfig_Click()

Config.Show 1, Me

End Sub

Private Sub RMExit_Click()

Call EndApp

End Sub

Private Sub RMWeb_Actualiza_Click()

'// go to the Creaciones Digitales web page...
ShellExecute 0, "open", "http://www.liveupdate.interbandas.com.ar", vbNullString, vbNullString, SW_MAXIMIZE

End Sub

Private Sub RMWeb_Help_Click()

'// go to the Creaciones Digitales web page...
ShellExecute 0, "open", "http://www.interbandas.com.ar/Productos.htm", vbNullString, vbNullString, SW_MAXIMIZE

End Sub

Private Sub RMWeb_Novedades_Click()

'// go to the Creaciones Digitales web page...
ShellExecute 0, "open", "http://www.interbandas.com.ar", vbNullString, vbNullString, SW_MAXIMIZE

End Sub

Private Sub RMWeb_Soporte_Click()

'// go to the Creaciones Digitales web page...
ShellExecute 0, "open", "http://www.interbandas.com.ar/Servicios.htm", vbNullString, vbNullString, SW_MAXIMIZE

End Sub

Private Sub SbEst01_Click()

    'If TopMenu.SbEst01.Checked = False Then
    '    ShowWindow "Est01"
    '    If Est01.Command1.Caption = ">" Then
    '        Est01.Width = 15360
    '        Est01.Left = 0
    '    End If
    '    If Tanda01.WindowState = 0 Then
    '        ShowWindow "Tnd01"
    '    End If
    'Else
    '    HideWindow "Est01"
    '    If Tanda01.WindowState = 0 Then
    '        ShowWindow "Tnd01"
    '    End If
    'End If

End Sub

Private Sub SbEst02_Click()

    If TopMenu.SbEst02.Checked = False Then
        ShowWindow "Est02"
        If Est02.Command1.Caption = ">" Then
            Est02.Width = 15360
            Est02.Left = 0
        End If
        If Prg01.WindowState = 0 Then
            ShowWindow "Prg01"
        End If
    Else
        HideWindow "Est02"
        If Prg01.WindowState = 0 Then
            ShowWindow "Prg01"
        End If
    End If

End Sub

Private Sub SbExplor_Click()

    If TopMenu.SbExplor.Checked = False Then
        ShowWindow "Explor01"
    Else
        HideWindow "Explor01"
    End If

End Sub

Private Sub SbHerram_Click()

    If TopMenu.SbHerram.Checked = False Then
        ShowWindow "DwMenu"
    Else
        HideWindow "DwMenu"
    End If

End Sub

Private Sub SbOrderAuto_Click()

'ordenamos las ventanas determinadas por defecto
If TopMenu.View3x3.Checked = True Or TopMenu.ViewDefault.Checked = True Then
    OrderWindow "TopMenu", "Default"
    OrderWindow "DwMenu", "Default"
    OrderWindow "Est01", "Default"
    OrderWindow "Est02", "Default"
    OrderWindow "Prg01", "Default"
    OrderWindow "Tnd01", "Default"
End If
If TopMenu.View4x4h.Checked = True Then
    OrderWindow "TopMenu", "4x4h"
    OrderWindow "DwMenu", "4x4h"
    OrderWindow "Est01", "4x4h"
    OrderWindow "Est02", "4x4h"
    OrderWindow "Tnd01", "4x4h"
    OrderWindow "Prg01", "4x4h"
End If
If TopMenu.View4x4v.Checked = True Then
    OrderWindow "TopMenu", "4x4v"
    OrderWindow "DwMenu", "4x4v"
    OrderWindow "Est01", "4x4v"
    OrderWindow "Est02", "4x4v"
    OrderWindow "Tnd01", "4x4v"
    OrderWindow "Prg01", "4x4v"
End If

End Sub

Private Sub SbPrg01_Click()

    If TopMenu.SbPrg01.Checked = False Then
        ShowWindow "Prg01"
        If Tanda01.WindowState = 0 Then
            ShowWindow "Tnd01"
        End If
        ShowWindow "Prg01"
    Else
        HideWindow "Prg01"
        If Tanda01.WindowState = 0 Then
            ShowWindow "Tnd01"
        End If
    End If

End Sub

Private Sub SbTnd01_Click()

    If TopMenu.SbTnd01.Checked = False Then
        ShowWindow "Tnd01"
        If Prg01.WindowState = 0 Then
            ShowWindow "Prg01"
        End If
    Else
        HideWindow "Tnd01"
        If Prg01.WindowState = 0 Then
            ShowWindow "Prg01"
        End If
    End If

End Sub

Private Sub t1_Click()

Dim Tvol As Long

Result = LoadPlugIn("RMVoice", "LoadSilent")
If Result = "NotOk" Then
    MsgBox "Not loaded. rmvoice"
Else
    If Stream01IsPlaying = True Then
        Tvol = CLng(Est01.LblCurrVol.Caption) / 4
        Est01.LblOutvol.Caption = Tvol
        Est01.LblInvol.Caption = Est01.LblCurrVol.Caption
        Est01.Tmout.Enabled = True
        Est01.Tmout.Interval = 5
    Else
        If Stream02IsPlaying = True Then
            Tvol = CLng(Est02.LblCurrVol.Caption) / 4
            Est02.LblOutvol.Caption = Tvol
            Est02.LblInvol.Caption = Est02.LblCurrVol.Caption
            Est02.Tmout.Enabled = True
            Est02.Tmout.Interval = 5
        End If
    End If
    Result = LoadPlugIn("RMVoice", "SayHora")
    'RMPlugIn.ExecuteCommand ("SayHora")
    TimerRMVoiceCheck.Enabled = True
    TimerRMVoiceCheck.Interval = 1000
End If

End Sub

Private Sub t2_Click()

Call t1_Click

End Sub

Private Sub t3_Click()

Call t1_Click

End Sub

Private Sub t4_Click()

Call t1_Click

End Sub

Private Sub t5_Click()

Call t1_Click

End Sub

Private Sub t6_Click()

Call t1_Click

End Sub

Private Sub t7_Click()

Call t1_Click

End Sub

Private Sub t8_Click()

Call t1_Click

End Sub

Private Sub TimerRMVoiceCheck_Timer()

'//////////////////////////////////////////////////
'* This timer is intented only for use with
'* RMVoice.dll PlugIn.-
'//////////////////////////////////////////////////

Dim Tvol As Long

If RMPlugIn.PlugIsRunning = False Then
    If Stream01IsPlaying = True Then
        If Est12Control.Origen1.Caption = "E1" Then
            Est01.TMin.Enabled = True
            Est01.TMin.Interval = 5
        End If
    Else
        If Stream02IsPlaying = True Then
            If Est12Control.Origen2.Caption = "E2" Then
                Est02.TMin.Enabled = True
                Est02.TMin.Interval = 5
            End If
        Else
            TimerRMVoiceCheck.Interval = 0
            TimerRMVoiceCheck.Enabled = False
        End If
    End If
End If

End Sub

Private Sub TmrSendPos_Timer()

'//////////////////////////////////////////////////
'* This timer is intented only for use with
'* RMplayer.dll PlugIn.-
'//////////////////////////////////////////////////

Dim Rpos1 As String
Dim Rpos2 As String
Dim Rlen1 As String
Dim Rlen2 As String

Dim TimePos As String
Dim TimePosLen As String
Dim BytePos As String
Dim BytePosLen As String

Dim Convt1 As Long

Dim Test As Long
Dim Test2 As Long

On Error Resume Next

If Est12Control.StopLabel1.Caption = "Stream" Then
    If Stream01IsPlaying = True Then
        Rpos1 = Stream01GetPosition(1)  'position in time
        Rpos2 = Stream01GetPosition(2)  'position in bytes
        Rlen1 = Stream01GetLen(1)   'lenght in time
        Rlen2 = Stream01GetLen(2)   'lengh in bytes
        TimePos = FormatSegs(Rpos1)
        TimePosLen = FormatSegs(Rlen1)
        BytePos = Rpos2
        BytePosLen = Rlen2
        'chequeamos por el tipo de visualizacion (normal o restante)
        If Est12Control.Origen1.Caption = "E1" Then
            RMPlugIn.SetStreamName Est01.Label1.Caption
            RMPlugIn.SetStatusText ("Estación 01")
            If TopMenu.LType.Caption = "Normal" Then
                Test = CLng(Trim(TimePos))
                Test2 = CLng(Trim(BytePos))
                RMPlugIn.SetTime ConvSecToMin(CDbl(Test)), "Normal"
                Call SendMiniFFT(1, "Stream", 6) 'fft spectrum display
                Call SendMiniScope(1, "Stream")
            Else
                Test = CLng(Trim(TimePosLen)) - CLng(Trim(TimePos))
                Test2 = CLng(Trim(BytePosLen)) - CLng(Trim(BytePos))
                RMPlugIn.SetTime ConvSecToMin(CDbl(Test)), "Restante"
                Call SendMiniFFT(1, "Stream", 6) 'fft spectrum display
                Call SendMiniScope(1, "Stream")
            End If
        Else
            If Est12Control.Origen1.Caption = "T1" Then
                RMPlugIn.SetStreamName Tanda01.T1Name.Caption
                RMPlugIn.SetStatusText "Tanda - Dev: 1"
                Call GetNextStreamTime
                If TopMenu.LType.Caption = "Normal" Then
                    Test = CLng(Trim(TimePos))
                    Test2 = CLng(Trim(BytePos))
                    RMPlugIn.SetTime ConvSecToMin(CDbl(Test)), "Normal"
                    Call SendMiniFFT(1, "Stream", 6) 'fft spectrum display
                    Call SendMiniScope(1, "Stream")
                Else
                    Test = CLng(Trim(TimePosLen)) - CLng(Trim(TimePos))
                    Test2 = CLng(Trim(BytePosLen)) - CLng(Trim(BytePos))
                    RMPlugIn.SetTime ConvSecToMin(CDbl(Test)), "Restante"
                    Call SendMiniFFT(1, "Stream", 6) 'fft spectrum display
                    Call SendMiniScope(1, "Stream")
                End If
            End If
        End If
    Else
        If Stream02IsPlaying = True Then
            Rpos1 = Stream02GetPosition(1)  'position in time
            Rpos2 = Stream02GetPosition(2)  'position in bytes
            Rlen1 = Stream02GetLen(1)   'lenght in time
            Rlen2 = Stream02GetLen(2)   'lengh in bytes
            TimePos = FormatSegs(Rpos1)
            TimePosLen = FormatSegs(Rlen1)
            BytePos = Rpos2
            BytePosLen = Rlen2
            'chequeamos por el tipo de visualizacion (normal o restante)
            If Est12Control.Origen2.Caption = "E2" Then
                RMPlugIn.SetStreamName Est02.Label1.Caption
                RMPlugIn.SetStatusText "Estación 02"
                If TopMenu.LType.Caption = "Normal" Then
                    Test = CLng(Trim(TimePos))
                    Test2 = CLng(Trim(BytePos))
                    RMPlugIn.SetTime ConvSecToMin(CDbl(Test)), "Normal"
                    Call SendMiniFFT(2, "Stream", 6) 'fft spectrum display
                    Call SendMiniScope(2, "Stream")
                Else
                    Test = CLng(Trim(TimePosLen)) - CLng(Trim(TimePos))
                    Test2 = CLng(Trim(BytePosLen)) - CLng(Trim(BytePos))
                    RMPlugIn.SetTime ConvSecToMin(CDbl(Test)), "Restante"
                    Call SendMiniFFT(2, "Stream", 6) 'fft spectrum display
                    Call SendMiniScope(2, "Stream")
                End If
            Else
                If Est12Control.Origen2.Caption = "T2" Then
                    RMPlugIn.SetStreamName Tanda01.T2Name.Caption
                    RMPlugIn.SetStatusText "Tanda - Dev: 2"
                    Call GetNextStreamTime
                    If TopMenu.LType.Caption = "Normal" Then
                        Test = CLng(Trim(TimePos))
                        Test2 = CLng(Trim(BytePos))
                        RMPlugIn.SetTime ConvSecToMin(CDbl(Test)), "Normal"
                        Call SendMiniFFT(2, "Stream", 6) 'fft spectrum display
                        Call SendMiniScope(2, "Stream")
                    Else
                        Test = CLng(Trim(TimePosLen)) - CLng(Trim(TimePos))
                        Test2 = CLng(Trim(BytePosLen)) - CLng(Trim(BytePos))
                        RMPlugIn.SetTime ConvSecToMin(CDbl(Test)), "Restante"
                        Call SendMiniFFT(2, "Stream", 6) 'fft spectrum display
                        Call SendMiniScope(2, "Stream")
                    End If
                End If
            End If
        Else
            RMPlugIn.SetStreamName "---"
            RMPlugIn.SetStatusText "- D E T E N I D O -"
        End If
    End If
End If

End Sub

Private Sub View3x3_Click()

If TopMenu.View3x3.Checked = False Then
    TopMenu.View3x3.Checked = True
    TopMenu.ViewDefault.Checked = False
    TopMenu.View4x4h.Checked = False
    TopMenu.View4x4v.Checked = False
End If

If TopMenu.View3x3.Checked = True Or TopMenu.ViewDefault.Checked = True Then
    OrderWindow "TopMenu", "Default"
    OrderWindow "DwMenu", "Default"
    OrderWindow "Est01", "Default"
    OrderWindow "Est02", "Default"
    OrderWindow "Prg01", "Default"
    OrderWindow "Tnd01", "Default"
End If
If TopMenu.View4x4h.Checked = True Then
    OrderWindow "TopMenu", "4x4h"
    OrderWindow "DwMenu", "4x4h"
    OrderWindow "Est01", "4x4h"
    OrderWindow "Est02", "4x4h"
    OrderWindow "Tnd01", "4x4h"
    OrderWindow "Prg01", "4x4h"
End If
If TopMenu.View4x4v.Checked = True Then
    OrderWindow "TopMenu", "4x4v"
    OrderWindow "DwMenu", "4x4v"
    OrderWindow "Est01", "4x4v"
    OrderWindow "Est02", "4x4v"
    OrderWindow "Tnd01", "4x4v"
    OrderWindow "Prg01", "4x4v"
End If

End Sub

Private Sub View4x4h_Click()

If TopMenu.View4x4h.Checked = False Then
    TopMenu.View4x4h.Checked = True
    TopMenu.ViewDefault.Checked = False
    TopMenu.View4x4v.Checked = False
    TopMenu.View3x3.Checked = False
End If

If TopMenu.View3x3.Checked = True Or TopMenu.ViewDefault.Checked = True Then
    OrderWindow "TopMenu", "Default"
    OrderWindow "DwMenu", "Default"
    OrderWindow "Est01", "Default"
    OrderWindow "Est02", "Default"
    OrderWindow "Prg01", "Default"
    OrderWindow "Tnd01", "Default"
End If
If TopMenu.View4x4h.Checked = True Then
    OrderWindow "TopMenu", "4x4h"
    OrderWindow "DwMenu", "4x4h"
    OrderWindow "Est01", "4x4h"
    OrderWindow "Est02", "4x4h"
    OrderWindow "Tnd01", "4x4h"
    OrderWindow "Prg01", "4x4h"
End If
If TopMenu.View4x4v.Checked = True Then
    OrderWindow "TopMenu", "4x4v"
    OrderWindow "DwMenu", "4x4v"
    OrderWindow "Est01", "4x4v"
    OrderWindow "Est02", "4x4v"
    OrderWindow "Tnd01", "4x4v"
    OrderWindow "Prg01", "4x4v"
End If


End Sub

Private Sub View4x4v_Click()

If TopMenu.View4x4v.Checked = False Then
    TopMenu.View4x4v.Checked = True
    TopMenu.ViewDefault.Checked = False
    TopMenu.View4x4h.Checked = False
    TopMenu.View3x3.Checked = False
End If

If TopMenu.View3x3.Checked = True Or TopMenu.ViewDefault.Checked = True Then
    OrderWindow "TopMenu", "Default"
    OrderWindow "DwMenu", "Default"
    OrderWindow "Est01", "Default"
    OrderWindow "Est02", "Default"
    OrderWindow "Prg01", "Default"
    OrderWindow "Tnd01", "Default"
End If
If TopMenu.View4x4h.Checked = True Then
    OrderWindow "TopMenu", "4x4h"
    OrderWindow "DwMenu", "4x4h"
    OrderWindow "Est01", "4x4h"
    OrderWindow "Est02", "4x4h"
    OrderWindow "Tnd01", "4x4h"
    OrderWindow "Prg01", "4x4h"
End If
If TopMenu.View4x4v.Checked = True Then
    OrderWindow "TopMenu", "4x4v"
    OrderWindow "DwMenu", "4x4v"
    OrderWindow "Est01", "4x4v"
    OrderWindow "Est02", "4x4v"
    OrderWindow "Tnd01", "4x4v"
    OrderWindow "Prg01", "4x4v"
End If

End Sub

Private Sub ViewDefault_Click()

If TopMenu.ViewDefault.Checked = False Then
    TopMenu.ViewDefault.Checked = True
    TopMenu.View4x4v.Checked = False
    TopMenu.View4x4h.Checked = False
    TopMenu.View3x3.Checked = False
End If

If TopMenu.View3x3.Checked = True Or TopMenu.ViewDefault.Checked = True Then
    OrderWindow "TopMenu", "Default"
    OrderWindow "DwMenu", "Default"
    OrderWindow "Est01", "Default"
    OrderWindow "Est02", "Default"
    OrderWindow "Prg01", "Default"
    OrderWindow "Tnd01", "Default"
End If
If TopMenu.View4x4h.Checked = True Then
    OrderWindow "TopMenu", "4x4h"
    OrderWindow "DwMenu", "4x4h"
    OrderWindow "Est01", "4x4h"
    OrderWindow "Est02", "4x4h"
    OrderWindow "Tnd01", "4x4h"
    OrderWindow "Prg01", "4x4h"
End If
If TopMenu.View4x4v.Checked = True Then
    OrderWindow "TopMenu", "4x4v"
    OrderWindow "DwMenu", "4x4v"
    OrderWindow "Est01", "4x4v"
    OrderWindow "Est02", "4x4v"
    OrderWindow "Tnd01", "4x4v"
    OrderWindow "Prg01", "4x4v"
End If

End Sub
