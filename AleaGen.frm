VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AleatorGen 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7470
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9615
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      Caption         =   "A&yuda"
      Height          =   375
      Left            =   120
      TabIndex        =   115
      ToolTipText     =   "Abre una Configuración guardada anteriormente"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "&Abrir"
      Height          =   375
      Left            =   3000
      TabIndex        =   113
      ToolTipText     =   "Abre una Configuración guardada anteriormente"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   4200
      TabIndex        =   112
      ToolTipText     =   "Guarda la Configuración actual"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton RestoreAll 
      Caption         =   "&Restaurar"
      Height          =   375
      Left            =   5400
      TabIndex        =   111
      ToolTipText     =   "Restaura la Configuración a sus valores originales"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   8400
      TabIndex        =   110
      ToolTipText     =   "Cancela la operación y regresa a DAF"
      Top             =   6960
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Configuracion"
      TabPicture(0)   =   "AleaGen.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label15"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "PanelComerciales1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "PanelRadio1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "PanelComerciales2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "PanelRadio2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "PanelTemas"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CmD1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Temas"
      TabPicture(1)   =   "AleaGen.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "Command4"
      Tab(1).Control(2)=   "Command2"
      Tab(1).Control(3)=   "List1"
      Tab(1).Control(4)=   "Dir1"
      Tab(1).Control(5)=   "File1"
      Tab(1).Control(6)=   "Drive1"
      Tab(1).Control(7)=   "Frame3"
      Tab(1).Control(8)=   "Label6"
      Tab(1).Control(9)=   "L04"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Comerciales"
      TabPicture(2)   =   "AleaGen.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command3"
      Tab(2).Control(1)=   "Command8"
      Tab(2).Control(2)=   "Command6"
      Tab(2).Control(3)=   "List2"
      Tab(2).Control(4)=   "Dir2"
      Tab(2).Control(5)=   "File2"
      Tab(2).Control(6)=   "Drive2"
      Tab(2).Control(7)=   "Frame2"
      Tab(2).Control(8)=   "Label8"
      Tab(2).Control(9)=   "Label7"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Institucionales"
      TabPicture(3)   =   "AleaGen.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command5"
      Tab(3).Control(1)=   "Command12"
      Tab(3).Control(2)=   "Command10"
      Tab(3).Control(3)=   "List3"
      Tab(3).Control(4)=   "Dir3"
      Tab(3).Control(5)=   "File3"
      Tab(3).Control(6)=   "Drive3"
      Tab(3).Control(7)=   "Frame4"
      Tab(3).Control(8)=   "Label10"
      Tab(3).Control(9)=   "Label9"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Generador de Tandas"
      TabPicture(4)   =   "AleaGen.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "EspHora"
      Tab(4).Control(1)=   "Command7"
      Tab(4).Control(2)=   "Frame9"
      Tab(4).Control(3)=   "Frame8"
      Tab(4).Control(4)=   "Frame7"
      Tab(4).Control(5)=   "Frame6"
      Tab(4).Control(6)=   "Command15"
      Tab(4).Control(7)=   "Frame5"
      Tab(4).Control(8)=   "Label41"
      Tab(4).Control(9)=   "Label34"
      Tab(4).Control(10)=   "Label35"
      Tab(4).Control(11)=   "Label17"
      Tab(4).Control(12)=   "Label16"
      Tab(4).ControlCount=   13
      TabCaption(5)   =   "Acerca"
      TabPicture(5)   =   "AleaGen.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Picture1"
      Tab(5).Control(1)=   "Label37"
      Tab(5).Control(2)=   "Label31"
      Tab(5).ControlCount=   3
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   -74280
         Picture         =   "AleaGen.frx":00A8
         ScaleHeight     =   600
         ScaleWidth      =   2130
         TabIndex        =   116
         Top             =   1680
         Width           =   2130
      End
      Begin VB.TextBox EspHora 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   -67200
         MaxLength       =   2
         TabIndex        =   108
         Text            =   "00"
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<<< GENERAR TANDA >>>"
         Height          =   495
         Left            =   -69120
         Picture         =   "AleaGen.frx":14C2
         TabIndex        =   95
         ToolTipText     =   " GENERAR LA TANDA "
         Top             =   3720
         Width           =   3255
      End
      Begin VB.Frame Frame9 
         Height          =   1695
         Left            =   -74760
         TabIndex        =   85
         Top             =   3000
         Width           =   5415
         Begin VB.Label Label39 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   615
            Left            =   4440
            TabIndex        =   101
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   615
            Left            =   3600
            TabIndex        =   100
            Top             =   960
            Width           =   135
         End
         Begin VB.Label TndS 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   615
            Left            =   4560
            TabIndex        =   99
            Top             =   960
            Width           =   735
         End
         Begin VB.Label TndM 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   615
            Left            =   3720
            TabIndex        =   98
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "PROCESO DE CREACION - STATUS"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   240
            Width           =   5175
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "TIEMPO TOTAL DE TANDA"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2880
            TabIndex        =   93
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label TndH 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   615
            Left            =   2880
            TabIndex        =   92
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "INSTITUCIONALES"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2040
            TabIndex        =   90
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "COMERCIALES"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2040
            TabIndex        =   88
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "TEMAS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2040
            TabIndex        =   86
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Institucionales"
         Height          =   1095
         Left            =   -68760
         TabIndex        =   73
         Top             =   1320
         Width           =   2895
         Begin VB.Label Label28 
            Caption         =   "tema / as"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   84
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Radio2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "-"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   83
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label26 
            Caption         =   "Cada"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   82
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label25 
            Caption         =   "institucional / es"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   81
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Radio1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "-"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   960
            TabIndex        =   80
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label23 
            Caption         =   "Incluir"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   79
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Comerciales"
         Height          =   1095
         Left            =   -71760
         TabIndex        =   71
         Top             =   1320
         Width           =   2895
         Begin VB.Label Label22 
            Caption         =   "tema / as"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   78
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Comer2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "-"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   77
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label20 
            Caption         =   "Cada"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   76
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label19 
            Caption         =   "comercial / es"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   75
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Comer1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "-"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   74
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label111 
            Caption         =   "Incluir"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   72
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Temas"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   69
         Top             =   1320
         Width           =   2895
         Begin VB.Label TemasDsc 
            Caption         =   "No ha seleccionado ningun tema"
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   120
            TabIndex        =   70
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "B"
         Height          =   495
         Left            =   -71640
         TabIndex        =   66
         ToolTipText     =   "Eliminar todo el contenido del Listado de Temas"
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "B"
         Height          =   495
         Left            =   -71640
         TabIndex        =   65
         ToolTipText     =   "Eliminar todo el contenido del Listado de Temas"
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "B"
         Height          =   495
         Left            =   -71640
         TabIndex        =   64
         ToolTipText     =   "Eliminar todo el contenido del Listado de Temas"
         Top             =   3960
         Width           =   495
      End
      Begin MSComDlg.CommonDialog CmD1 
         Left            =   8640
         Top             =   4440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00000000&
         Caption         =   "&Finalizar"
         Height          =   375
         Left            =   -69120
         TabIndex        =   96
         ToolTipText     =   "Finalizar y regresar a DAF"
         Top             =   4320
         Width           =   3255
      End
      Begin VB.Frame PanelTemas 
         Height          =   1335
         Left            =   240
         TabIndex        =   60
         Top             =   960
         Width           =   2775
         Begin VB.CheckBox TemasOrAle 
            Caption         =   "Incluir Temas Aleatorios"
            Height          =   255
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox IncludComer 
            Caption         =   "Incluir Comerciales"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   1695
         End
         Begin VB.CheckBox IncludRadio 
            Caption         =   "Incluir Institucionales"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   960
            Width           =   2415
         End
      End
      Begin VB.Frame PanelRadio2 
         Caption         =   "Incluir cada..."
         Height          =   855
         Left            =   6240
         TabIndex        =   59
         Top             =   2880
         Width           =   2895
         Begin VB.CheckBox Rad1Tema 
            Caption         =   "1 Tema"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox Rad2Temas 
            Caption         =   "2 Temas"
            Height          =   255
            Left            =   1680
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox Rad3Temas 
            Caption         =   "3 Temas"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox Rad4Temas 
            Caption         =   "4 Temas"
            Height          =   255
            Left            =   1680
            TabIndex        =   18
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame PanelComerciales2 
         Caption         =   "Incluir cada..."
         Height          =   855
         Left            =   3240
         TabIndex        =   57
         Top             =   2880
         Width           =   2775
         Begin VB.CheckBox Cad4Temas 
            Caption         =   "4 Temas"
            Height          =   255
            Left            =   1560
            TabIndex        =   10
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox Cad3Temas 
            Caption         =   "3 Temas"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox Cad2Temas 
            Caption         =   "2 Temas"
            Height          =   255
            Left            =   1560
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox Cad1Tema 
            Caption         =   "1 Tema"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame PanelRadio1 
         Height          =   1815
         Left            =   6240
         TabIndex        =   55
         Top             =   960
         Width           =   2895
         Begin VB.CheckBox RadMixInter 
            Caption         =   "Incluir como Mixado Intermedio"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   2535
         End
         Begin VB.CheckBox Rad2Cad 
            Caption         =   "2 Institucionales cada..."
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1440
            Width           =   2415
         End
         Begin VB.CheckBox Rad1Cad 
            Caption         =   "1 Institucional cada..."
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CheckBox RadAle 
            Caption         =   "Ordenar Aleatoriamente"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Incluir"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   600
            Width           =   2655
         End
      End
      Begin VB.Frame PanelComerciales1 
         Height          =   1815
         Left            =   3240
         TabIndex        =   54
         Top             =   960
         Width           =   2775
         Begin VB.CheckBox Com3Cad 
            Caption         =   "3 Comerciales cada..."
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   2415
         End
         Begin VB.CheckBox Com2Cad 
            Caption         =   "2 Comerciales cada..."
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   2415
         End
         Begin VB.CheckBox Com1Cad 
            Caption         =   "1 Comercial cada..."
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Width           =   2415
         End
         Begin VB.CheckBox ComAle 
            Caption         =   "Ordenar Aleatoriamente"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Incluir"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.CommandButton Command12 
         Caption         =   ">"
         Height          =   495
         Left            =   -71640
         TabIndex        =   53
         ToolTipText     =   "Quitar todos los institucionales seleccionados"
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "<<"
         Height          =   1695
         Left            =   -71640
         TabIndex        =   52
         ToolTipText     =   "Transpasar todos los institucionales que se encuentran en el directorio seleccionado"
         Top             =   1560
         Width           =   495
      End
      Begin VB.ListBox List3 
         Height          =   3570
         Left            =   -74760
         TabIndex        =   50
         Top             =   960
         Width           =   3015
      End
      Begin VB.DirListBox Dir3 
         Height          =   3015
         Left            =   -68400
         TabIndex        =   49
         Top             =   1440
         Width           =   2535
      End
      Begin VB.FileListBox File3 
         Height          =   3015
         Left            =   -71040
         Pattern         =   "*.wav"
         TabIndex        =   48
         Top             =   1440
         Width           =   2535
      End
      Begin VB.DriveListBox Drive3 
         Height          =   315
         Left            =   -71040
         TabIndex        =   47
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton Command8 
         Caption         =   ">"
         Height          =   495
         Left            =   -71640
         TabIndex        =   45
         ToolTipText     =   "Quitar todos los comerciales seleccionados"
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "<<"
         Height          =   1695
         Left            =   -71640
         TabIndex        =   44
         ToolTipText     =   "Transpasar todos los comerciales que se encuentran en el directorio seleccionado"
         Top             =   1560
         Width           =   495
      End
      Begin VB.ListBox List2 
         Height          =   3570
         Left            =   -74760
         TabIndex        =   42
         Top             =   960
         Width           =   3015
      End
      Begin VB.DirListBox Dir2 
         Height          =   3015
         Left            =   -68400
         TabIndex        =   41
         Top             =   1440
         Width           =   2535
      End
      Begin VB.FileListBox File2 
         Height          =   3015
         Left            =   -71040
         Pattern         =   "*.wav"
         TabIndex        =   40
         Top             =   1440
         Width           =   2535
      End
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   -71040
         TabIndex        =   39
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         Caption         =   ">"
         Height          =   495
         Left            =   -71640
         TabIndex        =   37
         ToolTipText     =   "Eliminar el tema seleccionado del Listado de Temas"
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<<"
         Height          =   1695
         Left            =   -71640
         TabIndex        =   36
         ToolTipText     =   "Transpasar todos los temas del directorio seleccionado"
         Top             =   1560
         Width           =   495
      End
      Begin VB.ListBox List1 
         Height          =   3570
         Left            =   -74760
         TabIndex        =   34
         Top             =   960
         Width           =   3015
      End
      Begin VB.DirListBox Dir1 
         Height          =   3015
         Left            =   -68400
         TabIndex        =   33
         Top             =   1440
         Width           =   2535
      End
      Begin VB.FileListBox File1 
         Height          =   3015
         Left            =   -71040
         Pattern         =   "*.wav"
         TabIndex        =   32
         Top             =   1440
         Width           =   2535
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   -71040
         TabIndex        =   31
         Top             =   960
         Width           =   2535
      End
      Begin VB.Frame Frame5 
         Caption         =   "IMPORTANTE"
         ForeColor       =   &H000000C0&
         Height          =   975
         Left            =   -74760
         TabIndex        =   28
         Top             =   5160
         Width           =   8895
         Begin VB.Label Label5 
            Caption         =   $"AleaGen.frx":17CC
            ForeColor       =   &H000000C0&
            Height          =   615
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   8655
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "IMPORTANTE"
         ForeColor       =   &H000000C0&
         Height          =   1215
         Left            =   -74760
         TabIndex        =   26
         Top             =   4920
         Width           =   8895
         Begin VB.Label Label4 
            Caption         =   $"AleaGen.frx":18F6
            ForeColor       =   &H000000C0&
            Height          =   855
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   8655
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "IMPORTANTE"
         ForeColor       =   &H000000C0&
         Height          =   1215
         Left            =   -74760
         TabIndex        =   24
         Top             =   4920
         Width           =   8895
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   $"AleaGen.frx":1A8D
            ForeColor       =   &H000000C0&
            Height          =   855
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   8655
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "IMPORTANTE"
         ForeColor       =   &H000000C0&
         Height          =   1215
         Left            =   240
         TabIndex        =   22
         Top             =   4920
         Width           =   8895
         Begin VB.Label Label2 
            Caption         =   $"AleaGen.frx":1BFA
            ForeColor       =   &H000000C0&
            Height          =   855
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   8655
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "IMPORTANTE"
         ForeColor       =   &H000000C0&
         Height          =   1215
         Left            =   -74760
         TabIndex        =   20
         Top             =   4920
         Width           =   8895
         Begin VB.Label Label1 
            Caption         =   $"AleaGen.frx":1DAA
            ForeColor       =   &H000000C0&
            Height          =   855
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   8655
         End
      End
      Begin VB.Label Label41 
         Caption         =   "horas"
         Height          =   255
         Left            =   -66480
         TabIndex        =   109
         Top             =   3300
         Width           =   495
      End
      Begin VB.Label Label34 
         Caption         =   "Especifique la duración de la Tanda a Generar."
         Height          =   495
         Left            =   -69120
         TabIndex        =   107
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Versión 1.0 - 32bits"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -73920
         TabIndex        =   106
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "GENERADOR ALEATORIO AUTOMATICO DE TANDAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   -74280
         TabIndex        =   105
         Top             =   960
         Width           =   8055
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "GENERADOR ALEATORIO DE TANDAS"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74760
         TabIndex        =   97
         Top             =   2640
         Width           =   8895
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ud. ha especificado que desea crear una tanda con las siguientes características:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   68
         Top             =   960
         Width           =   8895
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "RESUMEN"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74760
         TabIndex        =   67
         Top             =   600
         Width           =   8895
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "INSTITUCIONALES"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   63
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "COMERCIALES"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   62
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "TEMAS"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Examinador de archivos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -71040
         TabIndex        =   51
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Listado de Institucionales"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74760
         TabIndex        =   46
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Examinador de archivos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -71040
         TabIndex        =   43
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Listado de Comerciales"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74760
         TabIndex        =   38
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Examinador de archivos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -71040
         TabIndex        =   35
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label L04 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Listado de Temas"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "GENERADOR AUTOMATICO DE TANDAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   114
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label Lss 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "--"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   104
      Top             =   8235
      Width           =   615
   End
   Begin VB.Label Lmm 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "--"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   103
      Top             =   7995
      Width           =   615
   End
   Begin VB.Label Lhh 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "--"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   102
      Top             =   7755
      Width           =   615
   End
End
Attribute VB_Name = "AleatorGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Mtf As Temas

Dim NuevoNombre As String
Dim NuevoNumero As Integer
Dim Longitud As Integer
Dim Position As Integer
Dim Position1 As Integer
Dim Position2 As Integer
Dim NumReg As Integer
Dim Total As String
Dim TndHDuracion As Integer
Dim TndCDuracion As Integer
Dim TemasPorCom
Dim CantCom
Dim TemasPorRad
Dim CantRad
Dim ChkFl

Dim i, a, b, c
Dim MyNum

Sub GetTime(Tiempo As String)
' AQUI VA EL CODIGO PARA CALCULAR EL TIEMPO DE LAS TANDAS
Dim TotalMin
Dim TotalSeg
Dim TotalH
Dim ViejaHoraMin As Long
Dim ViejaHoraSeg As Long
Dim ViejaHora As Long
Dim NuevaHoraMin As Long
Dim NuevaHoraSeg As Long
Dim NuevaHora As Long
Dim HoraLen
Dim MinLen
Dim SegLen
Dim Horita
Dim Minuto
Dim Segundo

'COMENZAMOS A SUMAR EL TIEMPO teniendo en cuenta el tiempo que ya
'tenemos de los temas anteriores
Tiempo = LTrim(RTrim(Tiempo))
If Tiempo = "" Or Tiempo = " " Or Tiempo = "--:--" Then
    ViejaHoraMin = 0
    ViejaHoraSeg = 0
    ViejaHora = 0
Else
    ViejaHoraMin = Left$(Tiempo, 2)
    ViejaHoraSeg = Right$(Tiempo, 2)
    ViejaHora = 0
End If

TotalMin = Lmm.Caption
TotalSeg = Lss.Caption
TotalH = Lhh.Caption

If TotalH = "--" Or TotalH = "" Then
    TotalH = 0
Else
    TotalH = TotalH
End If
If TotalMin = "--" Or TotalMin = "" Then
    TotalMin = 0
Else
    TotalMin = TotalMin
End If
If TotalSeg = "--" Or TotalSeg = "" Then
    TotalSeg = 0
Else
    TotalSeg = TotalSeg
End If

'Sumamos los tiempos viejos con los nuevos y calculamos el total
NuevaHoraSeg = ViejaHoraSeg + TotalSeg
NuevaHoraMin = ViejaHoraMin + TotalMin
NuevaHora = ViejaHora + TotalH

If NuevaHoraSeg > 59 And NuevaHoraSeg < 119 Then
    NuevaHoraSeg = NuevaHoraSeg - 60: NuevaHoraMin = NuevaHoraMin + 1
Else
    NuevaHoraSeg = NuevaHoraSeg
End If
If NuevaHoraMin > 59 And NuevaHoraMin < 119 Then
    NuevaHoraMin = NuevaHoraMin - 60: NuevaHora = NuevaHora + 1
Else
    NuevaHoraMin = NuevaHoraMin
End If
If NuevaHora > 24 Or NuevaHora = 24 Then
    'MsgBox "No se pueden maniular TANDAS mayores a 24HS."
    'MsgBox "Por favor corrija el tiempo de las tandas para evitar inconvenientes. Creacion Abortada."
    NuevaHora = NuevaHora
    'Exit Sub
Else
    NuevaHora = NuevaHora
End If

Finalizacion:
'Ponemos el nuevo tiempo en los labels
Lhh.Caption = NuevaHora
Lmm.Caption = NuevaHoraMin
Lss.Caption = NuevaHoraSeg

'extraemos la longitud de los tiempos y si estan mal puestos los arreglamos
'Los tiempo por ejemplo no pueden aparecer asi:
'5:5:5
'tendrian que aparecer asi:
'05:05:05
'asi que basicamente lo que hacemos es corregir eso.
HoraLen = Len(Lhh.Caption)
MinLen = Len(Lmm.Caption)
SegLen = Len(Lss.Caption)

If HoraLen = 1 Then
    Horita = Lhh.Caption
    Lhh.Caption = "0" & Horita
Else
    Lhh.Caption = NuevaHora
End If
If MinLen = 1 Then
    Minuto = Lmm.Caption
    Lmm.Caption = "0" & Minuto
Else
    Lmm.Caption = NuevaHoraMin
End If
If SegLen = 1 Then
    Segundo = Lss.Caption
    Lss.Caption = "0" & Segundo
Else
    Lss.Caption = NuevaHoraSeg
End If
GoSub Continuar
End

Continuar:
End Sub

Sub AbreConfiguracion(Fnam As String)
Dim data As String
Dim t1, t2, t3
Dim C1, C2, c3, c4, c5, c6, c7, c8
Dim r1, r2, r3, r4, r5, r6, r7, r8

'abrimos el archivo
On Error GoTo er
Open Fnam For Input As #33
Input #33, data
Close #33

'extraemos los datos definidos por el usuario
'temas
t1 = Left$(data, 1)
'Comerciales
t2 = Mid$(data, 2, 1)
    C1 = Mid$(data, 4, 1)
    C2 = Mid$(data, 5, 1)
    c3 = Mid$(data, 6, 1)
    c4 = Mid$(data, 7, 1)
    c5 = Mid$(data, 8, 1)
    c6 = Mid$(data, 9, 1)
    c7 = Mid$(data, 10, 1)
    c8 = Mid$(data, 11, 1)
'ImagenRadio
t3 = Mid$(data, 3, 1)
    r1 = Mid$(data, 12, 1)
    r2 = Mid$(data, 13, 1)
    r3 = Mid$(data, 14, 1)
    r4 = Mid$(data, 15, 1)
    r5 = Mid$(data, 16, 1)
    r6 = Mid$(data, 17, 1)
    r7 = Mid$(data, 18, 1)
    r8 = Right$(data, 1)

'Actualizamos los Controles
'Temas
TemasOrAle.value = t1
'Comerciales
IncludComer.value = t2
    ComAle.value = C1
    Com1Cad.value = C2
    Com2Cad.value = c3
    Com3Cad.value = c4
    Cad1Tema.value = c5
    Cad2Temas.value = c6
    Cad3Temas.value = c7
    Cad4Temas.value = c8
'ImagenRadio
IncludRadio.value = t3
    RadAle.value = r1
    RadMixInter.value = r2
    Rad1Cad.value = r3
    Rad2Cad.value = r4
    Rad1Tema.value = r5
    Rad2Temas.value = r6
    Rad3Temas.value = r7
    Rad4Temas.value = r8
Exit Sub

er:
MsgBox "No se puede abrir el archivo de configuración del generador. Consulte a su proveedor de Software.", vbCritical
ErrorReporte "No se puede abrir la configuracion del generados. Modulo AleatorGen - AbreConfiguracion"
Exit Sub

End Sub


Sub GuardaConfiguracion(Fn As String, Dt As String)

On Error GoTo er
Open Fn For Output As #23
Write #23, Dt
Close #23
Exit Sub

er:
MsgBox "Ha ocurrido un error al intentar guardar la configuración del sistema. Consulte a su proveedor de Software.", vbCritical
ErrorReporte "No se pudo guardar la configuración del generador aleatorio. Modulo AleatorGen - GuardaConfiguracion."
Exit Sub

End Sub


Sub SoloTemas()

'SOLAMENTE TEMAS DE MANERA ALEATORIA
If EspHora.Text = "" Or EspHora.Text = " " Or EspHora.Text = "00" Or Left$(EspHora.Text, 1) = "-" Then
    MsgBox "La duración de la Tanda que desea generar es incorrecta. Por favor corrija la duración de la misma (en Hs) e intente nuevamente.", vbCritical
    Exit Sub
Else
    TndHDuracion = CInt(EspHora.Text)
End If

'ABRIMOS EL ARCHIVO
'para guardar los datos...
    CmD1.DialogTitle = "RM Generador de Tandas - Guardar Tanda Generada."
    CmD1.InitDir = App.Path & AppTandaDir
    CmD1.Filter = "Archivo de Tanda (*.*)|*.*"
    CmD1.FilterIndex = 1
    CmD1.ShowSave
    NuevoNombre = CmD1.FileName

If NuevoNombre = "" Or NuevoNombre = " " Then
    MsgBox "Debe especificar el nombre del archivo a generar. Escriba el nombre del archivo e intente nuevamente.", vbCritical
    Exit Sub
End If

NuevoNumero = FreeFile
Longitud = Len(Mtf)

On Error GoTo oups
Open NuevoNombre For Random As NuevoNumero Len = Longitud
Position = 0

Restarting:
If List1.ListCount < 1 Then
    MsgBox "No se han seleccionado los temas. Seleccione los temas e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If

For i = 1 To 2
    List1.Selected(Int((List1.ListCount * Rnd) + 1) - 1) = True
    ChkFl = Right$(List1.Text, 3)
    Select Case ChkFl
        Case "wav", "WAV", "Wav", "wAv", "waV"
            Position = Position + 1
            Mtf.id = Position
            Mtf.Name = List1.Text
            Label18.Caption = Position
            Mtf.Direccion = File1.Path
            Mtf.Hora = "00:00"
            Total = File1.Path & "\" & List1.Text
            ''wHeadInfo (Total)    'extraemos la informacion del WAV
            ''Mtf.Duracion = wInfo.wPlaytime
            ''GetTime Mtf.Duracion
            'Mtf.NameMix = "Sin Mix Intermedio"
            'Mtf.DireccionMix = "---"
            'Mtf.HoraMix = "00:00"
            'Mtf.DuracionMix = "00:00"
            Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
            Put #NuevoNumero, Position, Mtf
        Case "mp3", "MP3", "Mp3", "mP3"
            Position = Position + 1
            Mtf.id = Position
            Mtf.Name = List1.Text
            Label18.Caption = Position
            Mtf.Direccion = File1.Path
            Mtf.Hora = "00:00"
            Total = File1.Path & "\" & List1.Text
            ReadMP3Header (Total)        'extraemos la informacion del Mp3
            'Mtf.Duracion = MP3HInfo.FPlayTime
            'GetTime 'Mtf.Duracion
            'Mtf.NameMix = "Sin Mix Intermedio"
            'Mtf.DireccionMix = "---"
            'Mtf.HoraMix = "00:00"
            'Mtf.DuracionMix = "00:00"
            Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
            Put #NuevoNumero, Position, Mtf
        Case Else
            'xxxx           'dejamos en blanco para que no agregue nada mas
    End Select
Next i

TndCDuracion = CInt(Lhh.Caption)
If TndCDuracion >= TndHDuracion Then
    GoSub Finalizar
Else
    GoSub Restarting
End If
Exit Sub

'----------------------------------------------
Finalizar:
Close NuevoNumero
TndH.Caption = Lhh.Caption
TndM.Caption = Lmm.Caption
TndS.Caption = Lss.Caption
MsgBox "La Tanda " & UCase(NuevoNombre) & " ha sido generada satisfactoriamente.", vbInformation
Command15_Click
Exit Sub

'----------------------------------------------
oups:
MsgBox "Ha Ocurrido un error inesperado al intentar generar la Tanda. Por favor consulte con su proveedor de software.", vbCritical
Close
Resume Continue
Exit Sub

Continue:
End Sub
Sub TemasyComAuto()

'TEMAS Y COMERCIALES AUTOMATICOS ALEATORIOS
If EspHora.Text = "" Or EspHora.Text = " " Or EspHora.Text = "00" Or Left$(EspHora.Text, 1) = "-" Then
    MsgBox "La duración de la Tanda que desea generar es incorrecta. Por favor corrija la duración de la misma (en Hs) e intente nuevamente.", vbCritical
    Exit Sub
Else
    TndHDuracion = CInt(EspHora.Text)
End If

'ABRIMOS EL ARCHIVO
'para guardar los datos...
    CmD1.DialogTitle = "RM Generador de Tandas - Guardar Tanda Generada."
    CmD1.InitDir = App.Path & AppTandaDir
    CmD1.Filter = "Archivo de Tanda (*.*)|*.*"
    CmD1.FilterIndex = 1
    CmD1.ShowSave
    NuevoNombre = CmD1.FileName

If NuevoNombre = "" Or NuevoNombre = " " Then
    MsgBox "Debe especificar el nombre del archivo a generar. Escriba el nombre del archivo e intente nuevamente.", vbCritical
    Exit Sub
End If

NuevoNumero = FreeFile
Longitud = Len(Mtf)

On Error GoTo oups
Open NuevoNombre For Random As NuevoNumero Len = Longitud
Position = 0

Restarting:
If List1.ListCount < 1 Then
    MsgBox "No se han seleccionado los temas. Seleccione los temas e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If
If List2.ListCount < 1 Then
    MsgBox "No se han seleccionado los comerciales. realice la selección e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If

For i = 1 To 3
    MyNum = (Int(2 * Rnd) + 1)
    If MyNum = 1 Then
        List1.Selected(Int((List1.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List1.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List1.Text
                Label18.Caption = Position
                Mtf.Direccion = File1.Path
                Mtf.Hora = "00:00"
                Total = File1.Path & "\" & List1.Text
                ''wHeadInfo (Total)    'extraemos la informacion del WAV
                ''Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List1.Text
                Label18.Caption = Position
                Mtf.Direccion = File1.Path
                Mtf.Hora = "00:00"
                Total = File1.Path & "\" & List1.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    Else
        List2.Selected(Int((List2.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List2.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List2.Text
                Label24.Caption = Position
                Mtf.Direccion = File2.Path
                Mtf.Hora = "00:00"
                Total = File2.Path & "\" & List2.Text
                ''wHeadInfo (Total)    'extraemos la informacion del WAV
                ''Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List2.Text
                Label24.Caption = Position
                Mtf.Direccion = File2.Path
                Mtf.Hora = "00:00"
                Total = File2.Path & "\" & List2.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    End If
Next i

TndCDuracion = CInt(Lhh.Caption)
If TndCDuracion >= TndHDuracion Then
    GoSub Finalizar
Else
    GoSub Restarting
End If
Exit Sub

'----------------------------------------------
Finalizar:
Close NuevoNumero
TndH.Caption = Lhh.Caption
TndM.Caption = Lmm.Caption
TndS.Caption = Lss.Caption
MsgBox "La Tanda " & UCase(NuevoNombre) & " ha sido generada satisfactoriamente.", vbInformation
Command15_Click
Exit Sub

'----------------------------------------------
oups:
MsgBox "Ha Ocurrido un error inesperado al intentar generar la Tanda. Por favor consulte con su proveedor de software.", vbCritical
Close
Exit Sub

End Sub

Sub TemasyComAutoyRadAuto()

'TEMAS, COMERCIALES E INSTITUCIONALES AUTOMATICOS ALEATORIOS
If EspHora.Text = "" Or EspHora.Text = " " Or EspHora.Text = "00" Or Left$(EspHora.Text, 1) = "-" Then
    MsgBox "La duración de la Tanda que desea generar es incorrecta. Por favor corrija la duración de la misma (en Hs) e intente nuevamente.", vbCritical
    Exit Sub
Else
    TndHDuracion = CInt(EspHora.Text)
End If

'ABRIMOS EL ARCHIVO
'para guardar los datos...
    CmD1.DialogTitle = "DAF Generador de Tandas - Guardar Tanda Generada."
    CmD1.InitDir = App.Path & AppTandaDir
    CmD1.Filter = "Archivo de Tanda (*.*)|*.*"
    CmD1.FilterIndex = 1
    CmD1.ShowSave
    NuevoNombre = CmD1.FileName

If NuevoNombre = "" Or NuevoNombre = " " Then
    MsgBox "Debe especificar el nombre del archivo a generar. Escriba el nombre del archivo e intente nuevamente.", vbCritical
    Exit Sub
End If

NuevoNumero = FreeFile
Longitud = Len(Mtf)

On Error GoTo oups
Open NuevoNombre For Random As NuevoNumero Len = Longitud
Position = 0

Restarting:
If List1.ListCount < 1 Then
    MsgBox "No se han seleccionado los temas. Seleccione los temas e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If
If List2.ListCount < 1 Then
    MsgBox "No se han seleccionado los comerciales. realice la selección e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If
If List3.ListCount < 1 Then
    MsgBox "No se han seleccionado los comerciales de radio. realice la selección e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If

For i = 1 To 3
    MyNum = (Int(3 * Rnd) + 1)
    If MyNum = 1 Then
        List1.Selected(Int((List1.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List1.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List1.Text
                Label18.Caption = Position
                Mtf.Direccion = File1.Path
                Mtf.Hora = "00:00"
                Total = File1.Path & "\" & List1.Text
                ''wHeadInfo (Total)    'extraemos la informacion del WAV
                'Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List1.Text
                Label18.Caption = Position
                Mtf.Direccion = File1.Path
                Mtf.Hora = "00:00"
                Total = File1.Path & "\" & List1.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    Else
        If MyNum = 2 Then
            List2.Selected(Int((List2.ListCount * Rnd) + 1) - 1) = True
            ChkFl = Right$(List2.Text, 3)
            Select Case ChkFl
                Case "wav", "WAV", "Wav", "wAv", "waV"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List2.Text
                    Label24.Caption = Position
                    Mtf.Direccion = File2.Path
                    Mtf.Hora = "00:00"
                    Total = File2.Path & "\" & List2.Text
                    ''wHeadInfo (Total)    'extraemos la informacion del WAV
                    'Mtf.Duracion = wInfo.wPlaytime
                    'GetTime 'Mtf.Duracion
                    'Mtf.NameMix = "Sin Mix Intermedio"
                    'Mtf.DireccionMix = "---"
                    'Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case "mp3", "MP3", "Mp3", "mP3"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List2.Text
                    Label24.Caption = Position
                    Mtf.Direccion = File2.Path
                    Mtf.Hora = "00:00"
                    Total = File2.Path & "\" & List2.Text
                    ReadMP3Header (Total)        'extraemos la informacion del Mp3
                    'Mtf.Duracion = MP3HInfo.FPlayTime
                    'GetTime 'Mtf.Duracion
                    'Mtf.NameMix = "Sin Mix Intermedio"
                    'Mtf.DireccionMix = "---"
                    'Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case Else
                    'xxxx           'dejamos en blanco para que no agregue nada mas
            End Select
        Else
            List3.Selected(Int((List3.ListCount * Rnd) + 1) - 1) = True
            ChkFl = Right$(List3.Text, 3)
            Select Case ChkFl
                Case "wav", "WAV", "Wav", "wAv", "waV"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List3.Text
                    Label29.Caption = Position
                    Mtf.Direccion = File3.Path
                    Mtf.Hora = "00:00"
                    Total = File3.Path & "\" & List3.Text
                    ''wHeadInfo (Total)    'extraemos la informacion del WAV
                    'Mtf.Duracion = wInfo.wPlaytime
                    'GetTime 'Mtf.Duracion
                    'Mtf.NameMix = "Sin Mix Intermedio"
                    'Mtf.DireccionMix = "---"
                    'Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case "mp3", "MP3", "Mp3", "mP3"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List3.Text
                    Label29.Caption = Position
                    Mtf.Direccion = File3.Path
                    Mtf.Hora = "00:00"
                    Total = File3.Path & "\" & List3.Text
                    ReadMP3Header (Total)        'extraemos la informacion del Mp3
                    'Mtf.Duracion = MP3HInfo.FPlayTime
                    'GetTime 'Mtf.Duracion
                    'Mtf.NameMix = "Sin Mix Intermedio"
                    'Mtf.DireccionMix = "---"
                    'Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case Else
                    'xxxx           'dejamos en blanco para que no agregue nada mas
            End Select
        End If
    End If
Next i

TndCDuracion = CInt(Lhh.Caption)
If TndCDuracion >= TndHDuracion Then
    GoSub Finalizar
Else
    GoSub Restarting
End If
Exit Sub

'----------------------------------------------
Finalizar:
Close NuevoNumero
TndH.Caption = Lhh.Caption
TndM.Caption = Lmm.Caption
TndS.Caption = Lss.Caption
MsgBox "La Tanda " & UCase(NuevoNombre) & " ha sido generada satisfactoriamente.", vbInformation
Command15_Click
Exit Sub

'----------------------------------------------
oups:
MsgBox "Ha Ocurrido un error inesperado al intentar generar la Tanda. Por favor consulte con su proveedor de software.", vbCritical
Close
Exit Sub

End Sub

Sub TemasyComAutoyRadCustom()

'TEMAS Y COMERCIALES AUTOMATICOS ALEATORIOS E INSTITUCIONALES CUSTOM
If EspHora.Text = "" Or EspHora.Text = " " Or EspHora.Text = "00" Or Left$(EspHora.Text, 1) = "-" Then
    MsgBox "La duración de la Tanda que desea generar es incorrecta. Por favor corrija la duración de la misma (en Hs) e intente nuevamente.", vbCritical
    Exit Sub
Else
    TndHDuracion = CInt(EspHora.Text)
End If

'ABRIMOS EL ARCHIVO
'para guardar los datos...
    CmD1.DialogTitle = "DAF Generador de Tandas - Guardar Tanda Generada."
    CmD1.InitDir = App.Path & AppTandaDir
    CmD1.Filter = "Archivo de Tanda (*.*)|*.*"
    CmD1.FilterIndex = 1
    CmD1.ShowSave
    NuevoNombre = CmD1.FileName

If NuevoNombre = "" Or NuevoNombre = " " Then
    MsgBox "Debe especificar el nombre del archivo a generar. Escriba el nombre del archivo e intente nuevamente.", vbCritical
    Exit Sub
End If

NuevoNumero = FreeFile
Longitud = Len(Mtf)

On Error GoTo oups
Open NuevoNombre For Random As NuevoNumero Len = Longitud
Position = 0

Restarting:
If List1.ListCount < 1 Then
    MsgBox "No se han seleccionado los temas. Seleccione los temas e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If
If List2.ListCount < 1 Then
    MsgBox "No se han seleccionado los comerciales. realice la selección e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If
If List3.ListCount < 1 Then
    MsgBox "No se han seleccionado los comerciales de radio. realice la selección e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If

TemasPorRad = Radio2.Caption    'cada cuantos temas?
CantRad = Radio1.Caption        'cuantos comerciales?

For i = 1 To 2
    For a = 1 To TemasPorRad
        MyNum = (Int(2 * Rnd) + 1)
        If MyNum = 1 Then
            List1.Selected(Int((List1.ListCount * Rnd) + 1) - 1) = True
            ChkFl = Right$(List1.Text, 3)
            Select Case ChkFl
                Case "wav", "WAV", "Wav", "wAv", "waV"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List1.Text
                    Label18.Caption = Position
                    Mtf.Direccion = File1.Path
                    Mtf.Hora = "00:00"
                    Total = File1.Path & "\" & List1.Text
                    ''wHeadInfo (Total)    'extraemos la informacion del WAV
                    'Mtf.Duracion = wInfo.wPlaytime
                    'GetTime 'Mtf.Duracion
                    'Mtf.NameMix = "Sin Mix Intermedio"
                    'Mtf.DireccionMix = "---"
                    'Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case "mp3", "MP3", "Mp3", "mP3"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List1.Text
                    Label18.Caption = Position
                    Mtf.Direccion = File1.Path
                    Mtf.Hora = "00:00"
                    Total = File1.Path & "\" & List1.Text
                    ReadMP3Header (Total)        'extraemos la informacion del Mp3
                    'Mtf.Duracion = MP3HInfo.FPlayTime
                    'GetTime 'Mtf.Duracion
                    'Mtf.NameMix = "Sin Mix Intermedio"
                    'Mtf.DireccionMix = "---"
                    'Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case Else
                    'xxxx           'dejamos en blanco para que no agregue nada mas
            End Select
        Else
            List2.Selected(Int((List2.ListCount * Rnd) + 1) - 1) = True
            ChkFl = Right$(List2.Text, 3)
            Select Case ChkFl
                Case "wav", "WAV", "Wav", "wAv", "waV"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List2.Text
                    Label24.Caption = Position
                    Mtf.Direccion = File2.Path
                    Mtf.Hora = "00:00"
                    Total = File2.Path & "\" & List2.Text
                    ''wHeadInfo (Total)    'extraemos la informacion del WAV
                    'Mtf.Duracion = wInfo.wPlaytime
                    'GetTime 'Mtf.Duracion
                    'Mtf.NameMix = "Sin Mix Intermedio"
                    'Mtf.DireccionMix = "---"
                    'Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case "mp3", "MP3", "Mp3", "mP3"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List2.Text
                    Label24.Caption = Position
                    Mtf.Direccion = File2.Path
                    Mtf.Hora = "00:00"
                    Total = File2.Path & "\" & List2.Text
                    ReadMP3Header (Total)        'extraemos la informacion del Mp3
                    'Mtf.Duracion = MP3HInfo.FPlayTime
                    'GetTime 'Mtf.Duracion
                    'Mtf.NameMix = "Sin Mix Intermedio"
                    'Mtf.DireccionMix = "---"
                    'Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case Else
                    'xxxx           'dejamos en blanco para que no agregue nada mas
            End Select
        End If
    Next a
    For b = 1 To CantRad
        List3.Selected(Int((List3.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List3.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List3.Text
                Label29.Caption = Position
                Mtf.Direccion = File3.Path
                Mtf.Hora = "00:00"
                Total = File3.Path & "\" & List3.Text
                ''wHeadInfo (Total)    'extraemos la informacion del WAV
                'Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List3.Text
                Label29.Caption = Position
                Mtf.Direccion = File3.Path
                Mtf.Hora = "00:00"
                Total = File3.Path & "\" & List3.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    Next b
Next i

TndCDuracion = CInt(Lhh.Caption)
If TndCDuracion >= TndHDuracion Then
    GoSub Finalizar
Else
    GoSub Restarting
End If
Exit Sub

'----------------------------------------------
Finalizar:
Close NuevoNumero
TndH.Caption = Lhh.Caption
TndM.Caption = Lmm.Caption
TndS.Caption = Lss.Caption
MsgBox "La Tanda " & UCase(NuevoNombre) & " ha sido generada satisfactoriamente.", vbInformation
Command15_Click
Exit Sub

'----------------------------------------------
oups:
MsgBox "Ha Ocurrido un error inesperado al intentar generar la Tanda. Por favor consulte con su proveedor de software.", vbCritical
Close
Exit Sub

End Sub

Sub TemasyComCustom()

'TEMAS Y COMERCIALES CUSTOM
If EspHora.Text = "" Or EspHora.Text = " " Or EspHora.Text = "00" Or Left$(EspHora.Text, 1) = "-" Then
    MsgBox "La duración de la Tanda que desea generar es incorrecta. Por favor corrija la duración de la misma (en Hs) e intente nuevamente.", vbCritical
    Exit Sub
Else
    TndHDuracion = CInt(EspHora.Text)
End If

'ABRIMOS EL ARCHIVO
'para guardar los datos...
    CmD1.DialogTitle = "DAF Generador de Tandas - Guardar Tanda Generada."
    CmD1.InitDir = App.Path & AppTandaDir
    CmD1.Filter = "Archivo de Tanda (*.*)|*.*"
    CmD1.FilterIndex = 1
    CmD1.ShowSave
    NuevoNombre = CmD1.FileName

If NuevoNombre = "" Or NuevoNombre = " " Then
    MsgBox "Debe especificar el nombre del archivo a generar. Escriba el nombre del archivo e intente nuevamente.", vbCritical
    Exit Sub
End If

NuevoNumero = FreeFile
Longitud = Len(Mtf)

On Error GoTo oups
Open NuevoNombre For Random As NuevoNumero Len = Longitud
Position = 0

Restarting:
If List1.ListCount < 1 Then
    MsgBox "No se han seleccionado los temas. Seleccione los temas e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If
If List2.ListCount < 1 Then
    MsgBox "No se han seleccionado los comerciales. realice la selección e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If

TemasPorCom = Comer2.Caption    'cada cuantos temas?
CantCom = Comer1.Caption        'cuantos comerciales?
For i = 1 To 2
    For a = 1 To TemasPorCom
        List1.Selected(Int((List1.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List1.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List1.Text
                Label18.Caption = Position
                Mtf.Direccion = File1.Path
                Mtf.Hora = "00:00"
                Total = File1.Path & "\" & List1.Text
                ''wHeadInfo (Total)    'extraemos la informacion del WAV
                'Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List1.Text
                Label18.Caption = Position
                Mtf.Direccion = File1.Path
                Mtf.Hora = "00:00"
                Total = File1.Path & "\" & List1.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    Next a
    For b = 1 To CantCom
        List2.Selected(Int((List2.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List2.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List2.Text
                Label24.Caption = Position
                Mtf.Direccion = File2.Path
                Mtf.Hora = "00:00"
                Total = File2.Path & "\" & List2.Text
                ''wHeadInfo (Total)    'extraemos la informacion del WAV
                'Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List2.Text
                Label24.Caption = Position
                Mtf.Direccion = File2.Path
                Mtf.Hora = "00:00"
                Total = File2.Path & "\" & List2.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    Next b
Next i

TndCDuracion = CInt(Lhh.Caption)
If TndCDuracion >= TndHDuracion Then
    GoSub Finalizar
Else
    GoSub Restarting
End If
Exit Sub

'----------------------------------------------
Finalizar:
Close NuevoNumero
TndH.Caption = Lhh.Caption
TndM.Caption = Lmm.Caption
TndS.Caption = Lss.Caption
MsgBox "La Tanda " & UCase(NuevoNombre) & " ha sido generada satisfactoriamente.", vbInformation
Command15_Click
Exit Sub

'----------------------------------------------
oups:
MsgBox "Ha Ocurrido un error inesperado al intentar generar la Tanda. Por favor consulte con su proveedor de software.", vbCritical
Close
Exit Sub

End Sub

Sub TemasyComCustomyRadAuto()

'TEMAS Y COMERCIALES CUSTOM E INSTITUCIONALES AUTO
If EspHora.Text = "" Or EspHora.Text = " " Or EspHora.Text = "00" Or Left$(EspHora.Text, 1) = "-" Then
    MsgBox "La duración de la Tanda que desea generar es incorrecta. Por favor corrija la duración de la misma (en Hs) e intente nuevamente.", vbCritical
    Exit Sub
Else
    TndHDuracion = CInt(EspHora.Text)
End If

'ABRIMOS EL ARCHIVO
'para guardar los datos...
    CmD1.DialogTitle = "DAF Generador de Tandas - Guardar Tanda Generada."
    CmD1.InitDir = App.Path & AppTandaDir
    CmD1.Filter = "Archivo de Tanda (*.*)|*.*"
    CmD1.FilterIndex = 1
    CmD1.ShowSave
    NuevoNombre = CmD1.FileName

If NuevoNombre = "" Or NuevoNombre = " " Then
    MsgBox "Debe especificar el nombre del archivo a generar. Escriba el nombre del archivo e intente nuevamente.", vbCritical
    Exit Sub
End If

NuevoNumero = FreeFile
Longitud = Len(Mtf)

On Error GoTo oups
Open NuevoNombre For Random As NuevoNumero Len = Longitud
Position = 0

Restarting:
If List1.ListCount < 1 Then
    MsgBox "No se han seleccionado los temas. Seleccione los temas e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If
If List2.ListCount < 1 Then
    MsgBox "No se han seleccionado los comerciales. realice la selección e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If
If List3.ListCount < 1 Then
    MsgBox "No se han seleccionado los comerciales de radio. realice la selección e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If

TemasPorCom = Comer2.Caption    'cada cuantos temas? 1,2,3,4
CantCom = Comer1.Caption        'cuantos comerciales?

For i = 1 To 2
    For a = 1 To TemasPorCom
        MyNum = (Int(2 * Rnd) + 1)
        If MyNum = 1 Then
            List1.Selected(Int((List1.ListCount * Rnd) + 1) - 1) = True
            ChkFl = Right$(List1.Text, 3)
            Select Case ChkFl
                Case "wav", "WAV", "Wav", "wAv", "waV"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List1.Text
                    Label18.Caption = Position
                    Mtf.Direccion = File1.Path
                    Mtf.Hora = "00:00"
                    Total = File1.Path & "\" & List1.Text
                    ''wHeadInfo (Total)    'extraemos la informacion del WAV
                    'Mtf.Duracion = wInfo.wPlaytime
                    'GetTime 'Mtf.Duracion
                    'Mtf.NameMix = "Sin Mix Intermedio"
                    'Mtf.DireccionMix = "---"
                    'Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case "mp3", "MP3", "Mp3", "mP3"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List1.Text
                    Label18.Caption = Position
                    Mtf.Direccion = File1.Path
                    Mtf.Hora = "00:00"
                    Total = File1.Path & "\" & List1.Text
                    ReadMP3Header (Total)        'extraemos la informacion del Mp3
                    'Mtf.Duracion = MP3HInfo.FPlayTime
                    'GetTime 'Mtf.Duracion
                    'Mtf.NameMix = "Sin Mix Intermedio"
                    'Mtf.DireccionMix = "---"
                    'Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case Else
                    'xxxx           'dejamos en blanco para que no agregue nada mas
            End Select
        Else
            List3.Selected(Int((List3.ListCount * Rnd) + 1) - 1) = True
            ChkFl = Right$(List3.Text, 3)
            Select Case ChkFl
                Case "wav", "WAV", "Wav", "wAv", "waV"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List3.Text
                    Label29.Caption = Position
                    Mtf.Direccion = File3.Path
                    Mtf.Hora = "00:00"
                    Total = File3.Path & "\" & List3.Text
                    'wHeadInfo (Total)    'extraemos la informacion del WAV
                    'Mtf.Duracion = wInfo.wPlaytime
                    'GetTime 'Mtf.Duracion
                    'Mtf.NameMix = "Sin Mix Intermedio"
                    'Mtf.DireccionMix = "---"
                    'Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case "mp3", "MP3", "Mp3", "mP3"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List3.Text
                    Label29.Caption = Position
                    Mtf.Direccion = File3.Path
                    Mtf.Hora = "00:00"
                    Total = File3.Path & "\" & List3.Text
                    ReadMP3Header (Total)        'extraemos la informacion del Mp3
                    'Mtf.Duracion = MP3HInfo.FPlayTime
                    'GetTime 'Mtf.Duracion
                    'Mtf.NameMix = "Sin Mix Intermedio"
                    'Mtf.DireccionMix = "---"
                    'Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case Else
                    'xxxx           'dejamos en blanco para que no agregue nada mas
            End Select
        End If
    Next a
    For b = 1 To CantCom
        List2.Selected(Int((List2.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List2.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List2.Text
                Label24.Caption = Position
                Mtf.Direccion = File2.Path
                Mtf.Hora = "00:00"
                Total = File2.Path & "\" & List2.Text
                'wHeadInfo (Total)    'extraemos la informacion del WAV
                'Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List2.Text
                Label24.Caption = Position
                Mtf.Direccion = File2.Path
                Mtf.Hora = "00:00"
                Total = File2.Path & "\" & List2.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                'Mtf.NameMix = "Sin Mix Intermedio"
                'Mtf.DireccionMix = "---"
                'Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    Next b
Next i

TndCDuracion = CInt(Lhh.Caption)
If TndCDuracion >= TndHDuracion Then
    GoSub Finalizar
Else
    GoSub Restarting
End If
Exit Sub

'----------------------------------------------
Finalizar:
Close NuevoNumero
TndH.Caption = Lhh.Caption
TndM.Caption = Lmm.Caption
TndS.Caption = Lss.Caption
MsgBox "La Tanda " & UCase(NuevoNombre) & " ha sido generada satisfactoriamente.", vbInformation
Command15_Click
Exit Sub

'----------------------------------------------
oups:
MsgBox "Ha Ocurrido un error inesperado al intentar generar la Tanda. Por favor consulte con su proveedor de software.", vbCritical
Close
Exit Sub

End Sub

Sub TemasyComCustomyRadCustom()

'TEMAS, COMERCIALES E INSTITUCIONALES CUSTOM
If EspHora.Text = "" Or EspHora.Text = " " Or EspHora.Text = "00" Or Left$(EspHora.Text, 1) = "-" Then
    MsgBox "La duración de la Tanda que desea generar es incorrecta. Por favor corrija la duración de la misma (en Hs) e intente nuevamente.", vbCritical
    Exit Sub
Else
    TndHDuracion = CInt(EspHora.Text)
End If

'ABRIMOS EL ARCHIVO
'para guardar los datos...
    CmD1.DialogTitle = "DAF Generador de Tandas - Guardar Tanda Generada."
    CmD1.InitDir = App.Path & AppTandaDir
    CmD1.Filter = "Archivo de Tanda (*.*)|*.*"
    CmD1.FilterIndex = 1
    CmD1.ShowSave
    NuevoNombre = CmD1.FileName

If NuevoNombre = "" Or NuevoNombre = " " Then
    MsgBox "Debe especificar el nombre del archivo a generar. Escriba el nombre del archivo e intente nuevamente.", vbCritical
    Exit Sub
End If

NuevoNumero = FreeFile
Longitud = Len(Mtf)

On Error GoTo oups
Open NuevoNombre For Random As NuevoNumero Len = Longitud
Position = 0
NumReg = 0

Restarting:
If List1.ListCount < 1 Then
    MsgBox "No se han seleccionado los temas. Seleccione los temas e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If
If List2.ListCount < 1 Then
    MsgBox "No se han seleccionado los comerciales. realice la selección e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If
If List3.ListCount < 1 Then
    MsgBox "No se han seleccionado los comerciales de radio. realice la selección e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If

CantCom = Comer1.Caption        'cuantos? 1 comer cada, 2 comer cada...
TemasPorCom = Comer2.Caption    'cada cuantos temas...? 1, 2, 3, 4
CantRad = Radio1.Caption        'cuantos? 1 insti cada, 2 insti cada...
TemasPorRad = Radio2.Caption    'cada cuantos temas? 1, 2 , 3 , 4

If TemasPorCom = TemasPorRad Then
    For i = 1 To TemasPorCom
        List1.Selected(Int((List1.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List1.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List1.Text
                Label18.Caption = Position
                Mtf.Direccion = File1.Path
                Mtf.Hora = "00:00"
                Total = File1.Path & "\" & List1.Text
                'wHeadInfo (Total)    'extraemos la informacion del WAV
                'Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List1.Text
                Label18.Caption = Position
                Mtf.Direccion = File1.Path
                Mtf.Hora = "00:00"
                Total = File1.Path & "\" & List1.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    Next i
    For a = 1 To CantCom
        List2.Selected(Int((List2.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List2.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List2.Text
                Label24.Caption = Position
                Mtf.Direccion = File2.Path
                Mtf.Hora = "00:00"
                Total = File2.Path & "\" & List2.Text
                'wHeadInfo (Total)    'extraemos la informacion del WAV
                'Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List2.Text
                Label24.Caption = Position
                Mtf.Direccion = File2.Path
                Mtf.Hora = "00:00"
                Total = File2.Path & "\" & List2.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    Next a
    For b = 1 To CantRad
        List3.Selected(Int((List3.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List3.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List3.Text
                Label29.Caption = Position
                Mtf.Direccion = File3.Path
                Mtf.Hora = "00:00"
                Total = File3.Path & "\" & List3.Text
                'wHeadInfo (Total)    'extraemos la informacion del WAV
                'Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List3.Text
                Label29.Caption = Position
                Mtf.Direccion = File3.Path
                Mtf.Hora = "00:00"
                Total = File3.Path & "\" & List3.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    Next b
Else
    If TemasPorCom > TemasPorRad Then
        For i = 1 To TemasPorCom
            List1.Selected(Int((List1.ListCount * Rnd) + 1) - 1) = True
            ChkFl = Right$(List1.Text, 3)
            Select Case ChkFl
                Case "wav", "WAV", "Wav", "wAv", "waV"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List1.Text
                    Label18.Caption = Position
                    Mtf.Direccion = File1.Path
                    Mtf.Hora = "00:00"
                    Total = File1.Path & "\" & List1.Text
                    'wHeadInfo (Total)    'extraemos la informacion del WAV
                    'Mtf.Duracion = wInfo.wPlaytime
                    'GetTime 'Mtf.Duracion
                    Mtf.NameMix = "Sin Mix Intermedio"
                    Mtf.DireccionMix = "---"
                    Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case "mp3", "MP3", "Mp3", "mP3"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List1.Text
                    Label18.Caption = Position
                    Mtf.Direccion = File1.Path
                    Mtf.Hora = "00:00"
                    Total = File1.Path & "\" & List1.Text
                    ReadMP3Header (Total)        'extraemos la informacion del Mp3
                    'Mtf.Duracion = MP3HInfo.FPlayTime
                    'GetTime 'Mtf.Duracion
                    Mtf.NameMix = "Sin Mix Intermedio"
                    Mtf.DireccionMix = "---"
                    Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case Else
                    'xxxx           'dejamos en blanco para que no agregue nada mas
            End Select
            If i = TemasPorRad Then
                For b = 1 To CantRad
                    List3.Selected(Int((List3.ListCount * Rnd) + 1) - 1) = True
                    ChkFl = Right$(List3.Text, 3)
                    Select Case ChkFl
                        Case "wav", "WAV", "Wav", "wAv", "waV"
                            Position = Position + 1
                            Mtf.id = Position
                            Mtf.Name = List3.Text
                            Label29.Caption = Position
                            Mtf.Direccion = File3.Path
                            Mtf.Hora = "00:00"
                            Total = File3.Path & "\" & List3.Text
                            'wHeadInfo (Total)    'extraemos la informacion del WAV
                            'Mtf.Duracion = wInfo.wPlaytime
                            'GetTime 'Mtf.Duracion
                            Mtf.NameMix = "Sin Mix Intermedio"
                            Mtf.DireccionMix = "---"
                            Mtf.HoraMix = "00:00"
                            'Mtf.DuracionMix = "00:00"
                            Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                            Put #NuevoNumero, Position, Mtf
                        Case "mp3", "MP3", "Mp3", "mP3"
                            Position = Position + 1
                            Mtf.id = Position
                            Mtf.Name = List3.Text
                            Label29.Caption = Position
                            Mtf.Direccion = File3.Path
                            Mtf.Hora = "00:00"
                            Total = File3.Path & "\" & List3.Text
                            ReadMP3Header (Total)        'extraemos la informacion del Mp3
                            'Mtf.Duracion = MP3HInfo.FPlayTime
                            'GetTime 'Mtf.Duracion
                            Mtf.NameMix = "Sin Mix Intermedio"
                            Mtf.DireccionMix = "---"
                            Mtf.HoraMix = "00:00"
                            'Mtf.DuracionMix = "00:00"
                            Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                            Put #NuevoNumero, Position, Mtf
                        Case Else
                            'xxxx           'dejamos en blanco para que no agregue nada mas
                    End Select
                Next b
            End If
        Next i
        For a = 1 To CantCom
            List2.Selected(Int((List2.ListCount * Rnd) + 1) - 1) = True
            ChkFl = Right$(List2.Text, 3)
            Select Case ChkFl
                Case "wav", "WAV", "Wav", "wAv", "waV"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List2.Text
                    Label24.Caption = Position
                    Mtf.Direccion = File2.Path
                    Mtf.Hora = "00:00"
                    Total = File2.Path & "\" & List2.Text
                    'wHeadInfo (Total)    'extraemos la informacion del WAV
                    'Mtf.Duracion = wInfo.wPlaytime
                    'GetTime 'Mtf.Duracion
                    Mtf.NameMix = "Sin Mix Intermedio"
                    Mtf.DireccionMix = "---"
                    Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case "mp3", "MP3", "Mp3", "mP3"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List2.Text
                    Label24.Caption = Position
                    Mtf.Direccion = File2.Path
                    Mtf.Hora = "00:00"
                    Total = File2.Path & "\" & List2.Text
                    ReadMP3Header (Total)        'extraemos la informacion del Mp3
                    'Mtf.Duracion = MP3HInfo.FPlayTime
                    'GetTime 'Mtf.Duracion
                    Mtf.NameMix = "Sin Mix Intermedio"
                    Mtf.DireccionMix = "---"
                    Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case Else
                    'xxxx           'dejamos en blanco para que no agregue nada mas
            End Select
        Next a
    Else
        For i = 1 To TemasPorRad
            List1.Selected(Int((List1.ListCount * Rnd) + 1) - 1) = True
            ChkFl = Right$(List1.Text, 3)
            Select Case ChkFl
                Case "wav", "WAV", "Wav", "wAv", "waV"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List1.Text
                    Label18.Caption = Position
                    Mtf.Direccion = File1.Path
                    Mtf.Hora = "00:00"
                    Total = File1.Path & "\" & List1.Text
                    'wHeadInfo (Total)    'extraemos la informacion del WAV
                    'Mtf.Duracion = wInfo.wPlaytime
                    'GetTime 'Mtf.Duracion
                    Mtf.NameMix = "Sin Mix Intermedio"
                    Mtf.DireccionMix = "---"
                    Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case "mp3", "MP3", "Mp3", "mP3"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List1.Text
                    Label18.Caption = Position
                    Mtf.Direccion = File1.Path
                    Mtf.Hora = "00:00"
                    Total = File1.Path & "\" & List1.Text
                    ReadMP3Header (Total)        'extraemos la informacion del Mp3
                    'Mtf.Duracion = MP3HInfo.FPlayTime
                    'GetTime 'Mtf.Duracion
                    Mtf.NameMix = "Sin Mix Intermedio"
                    Mtf.DireccionMix = "---"
                    Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case Else
                    'xxxx           'dejamos en blanco para que no agregue nada mas
            End Select
            If i = TemasPorCom Then
                For b = 1 To CantCom
                    List2.Selected(Int((List2.ListCount * Rnd) + 1) - 1) = True
                    ChkFl = Right$(List2.Text, 3)
                    Select Case ChkFl
                        Case "wav", "WAV", "Wav", "wAv", "waV"
                            Position = Position + 1
                            Mtf.id = Position
                            Mtf.Name = List2.Text
                            Label24.Caption = Position
                            Mtf.Direccion = File2.Path
                            Mtf.Hora = "00:00"
                            Total = File2.Path & "\" & List2.Text
                            'wHeadInfo (Total)    'extraemos la informacion del WAV
                            'Mtf.Duracion = wInfo.wPlaytime
                            'GetTime 'Mtf.Duracion
                            Mtf.NameMix = "Sin Mix Intermedio"
                            Mtf.DireccionMix = "---"
                            Mtf.HoraMix = "00:00"
                            'Mtf.DuracionMix = "00:00"
                            Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                            Put #NuevoNumero, Position, Mtf
                        Case "mp3", "MP3", "Mp3", "mP3"
                            Position = Position + 1
                            Mtf.id = Position
                            Mtf.Name = List2.Text
                            Label24.Caption = Position
                            Mtf.Direccion = File2.Path
                            Mtf.Hora = "00:00"
                            Total = File2.Path & "\" & List2.Text
                            ReadMP3Header (Total)        'extraemos la informacion del Mp3
                            'Mtf.Duracion = MP3HInfo.FPlayTime
                            'GetTime 'Mtf.Duracion
                            Mtf.NameMix = "Sin Mix Intermedio"
                            Mtf.DireccionMix = "---"
                            Mtf.HoraMix = "00:00"
                            'Mtf.DuracionMix = "00:00"
                            Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                            Put #NuevoNumero, Position, Mtf
                        Case Else
                            'xxxx           'dejamos en blanco para que no agregue nada mas
                    End Select
                Next b
            End If
        Next i
        For a = 1 To CantRad
            List3.Selected(Int((List3.ListCount * Rnd) + 1) - 1) = True
            ChkFl = Right$(List3.Text, 3)
            Select Case ChkFl
                Case "wav", "WAV", "Wav", "wAv", "waV"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List3.Text
                    Label29.Caption = Position
                    Mtf.Direccion = File3.Path
                    Mtf.Hora = "00:00"
                    Total = File3.Path & "\" & List3.Text
                    'wHeadInfo (Total)    'extraemos la informacion del WAV
                    'Mtf.Duracion = wInfo.wPlaytime
                    'GetTime 'Mtf.Duracion
                    Mtf.NameMix = "Sin Mix Intermedio"
                    Mtf.DireccionMix = "---"
                    Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case "mp3", "MP3", "Mp3", "mP3"
                    Position = Position + 1
                    Mtf.id = Position
                    Mtf.Name = List3.Text
                    Label29.Caption = Position
                    Mtf.Direccion = File3.Path
                    Mtf.Hora = "00:00"
                    Total = File3.Path & "\" & List3.Text
                    ReadMP3Header (Total)        'extraemos la informacion del Mp3
                    'Mtf.Duracion = MP3HInfo.FPlayTime
                    'GetTime 'Mtf.Duracion
                    Mtf.NameMix = "Sin Mix Intermedio"
                    Mtf.DireccionMix = "---"
                    Mtf.HoraMix = "00:00"
                    'Mtf.DuracionMix = "00:00"
                    Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                    Put #NuevoNumero, Position, Mtf
                Case Else
                    'xxxx           'dejamos en blanco para que no agregue nada mas
            End Select
        Next a
    End If
End If
TndCDuracion = CInt(Lhh.Caption)
If TndCDuracion >= TndHDuracion Then
    GoSub Finalizar
Else
    GoSub Restarting
End If
Exit Sub

'----------------------------------------------
Finalizar:
Close NuevoNumero
TndH.Caption = Lhh.Caption
TndM.Caption = Lmm.Caption
TndS.Caption = Lss.Caption
MsgBox "La Tanda " & UCase(NuevoNombre) & " ha sido generada satisfactoriamente.", vbInformation
Command15_Click
Exit Sub

'----------------------------------------------
oups:
MsgBox "Ha Ocurrido un error inesperado al intentar generar la Tanda. Por favor consulte con su proveedor de software.", vbCritical
Close
Exit Sub

End Sub

Sub TemasyRadAuto()

'TEMAS E INSTITUCIONALES ALEATORIOS AUTOMATICOS
If EspHora.Text = "" Or EspHora.Text = " " Or EspHora.Text = "00" Or Left$(EspHora.Text, 1) = "-" Then
    MsgBox "La duración de la Tanda que desea generar es incorrecta. Por favor corrija la duración de la misma (en Hs) e intente nuevamente.", vbCritical
    Exit Sub
Else
    TndHDuracion = CInt(EspHora.Text)
End If

'ABRIMOS EL ARCHIVO
'para guardar los datos...
    CmD1.DialogTitle = "DAF Generador de Tandas - Guardar Tanda Generada."
    CmD1.InitDir = App.Path & AppTandaDir
    CmD1.Filter = "Archivo de Tanda (*.*)|*.*"
    CmD1.FilterIndex = 1
    CmD1.ShowSave
    NuevoNombre = CmD1.FileName

If NuevoNombre = "" Or NuevoNombre = " " Then
    MsgBox "Debe especificar el nombre del archivo a generar. Escriba el nombre del archivo e intente nuevamente.", vbCritical
    Exit Sub
End If

NuevoNumero = FreeFile
Longitud = Len(Mtf)

On Error GoTo oups
Open NuevoNombre For Random As NuevoNumero Len = Longitud
Position = 0
NumReg = 0

Restarting:
If List1.ListCount < 1 Then
    MsgBox "No se han seleccionado los temas. Seleccione los temas e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If
If List3.ListCount < 1 Then
    MsgBox "No se han seleccionado los comerciales de radio. realice la selección e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If

For i = 1 To 3
    MyNum = (Int(2 * Rnd) + 1)
    If MyNum = 1 Then
        List1.Selected(Int((List1.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List1.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List1.Text
                Label18.Caption = Position
                Mtf.Direccion = File1.Path
                Mtf.Hora = "00:00"
                Total = File1.Path & "\" & List1.Text
                'wHeadInfo (Total)    'extraemos la informacion del WAV
                'Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List1.Text
                Label18.Caption = Position
                Mtf.Direccion = File1.Path
                Mtf.Hora = "00:00"
                Total = File1.Path & "\" & List1.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    Else
        List3.Selected(Int((List3.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List3.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List3.Text
                Label29.Caption = Position
                Mtf.Direccion = File3.Path
                Mtf.Hora = "00:00"
                Total = File3.Path & "\" & List3.Text
                'wHeadInfo (Total)    'extraemos la informacion del WAV
                'Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List3.Text
                Label29.Caption = Position
                Mtf.Direccion = File3.Path
                Mtf.Hora = "00:00"
                Total = File3.Path & "\" & List3.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    End If
Next i

TndCDuracion = CInt(Lhh.Caption)
If TndCDuracion >= TndHDuracion Then
    GoSub Finalizar
Else
    GoSub Restarting
End If
Exit Sub

'----------------------------------------------
Finalizar:
Close NuevoNumero
TndH.Caption = Lhh.Caption
TndM.Caption = Lmm.Caption
TndS.Caption = Lss.Caption
MsgBox "La Tanda " & UCase(NuevoNombre) & " ha sido generada satisfactoriamente.", vbInformation
Command15_Click
Exit Sub

'----------------------------------------------
oups:
MsgBox "Ha Ocurrido un error inesperado al intentar generar la Tanda. Por favor consulte con su proveedor de software.", vbCritical
Close
Exit Sub

End Sub

Sub TemasyRadCustom()

'TEMAS E INSTITUCIONALES CUSTOM
If EspHora.Text = "" Or EspHora.Text = " " Or EspHora.Text = "00" Or Left$(EspHora.Text, 1) = "-" Then
    MsgBox "La duración de la Tanda que desea generar es incorrecta. Por favor corrija la duración de la misma (en Hs) e intente nuevamente.", vbCritical
    Exit Sub
Else
    TndHDuracion = CInt(EspHora.Text)
End If

'ABRIMOS EL ARCHIVO
'para guardar los datos...
    CmD1.DialogTitle = "DAF Generador de Tandas - Guardar Tanda Generada."
    CmD1.InitDir = App.Path & AppTandaDir
    CmD1.Filter = "Archivo de Tanda (*.*)|*.*"
    CmD1.FilterIndex = 1
    CmD1.ShowSave
    NuevoNombre = CmD1.FileName

If NuevoNombre = "" Or NuevoNombre = " " Then
    MsgBox "Debe especificar el nombre del archivo a generar. Escriba el nombre del archivo e intente nuevamente.", vbCritical
    Exit Sub
End If

NuevoNumero = FreeFile
Longitud = Len(Mtf)

On Error GoTo oups
Open NuevoNombre For Random As NuevoNumero Len = Longitud
Position = 0
NumReg = 0

Restarting:
If List1.ListCount < 1 Then
    MsgBox "No se han seleccionado los temas. Seleccione los temas e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If
If List3.ListCount < 1 Then
    MsgBox "No se han seleccionado los comerciales de radio. realice la selección e intente nuevamente", vbCritical
    Close NuevoNumero
    Exit Sub
End If

TemasPorRad = Radio2.Caption    'cada cuantos temas?
CantRad = Radio1.Caption        'cuantos comerciales?
For i = 1 To 2
    For a = 1 To TemasPorRad
        List1.Selected(Int((List1.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List1.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List1.Text
                Label18.Caption = Position
                Mtf.Direccion = File1.Path
                Mtf.Hora = "00:00"
                Total = File1.Path & "\" & List1.Text
                'wHeadInfo (Total)    'extraemos la informacion del WAV
                'Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List1.Text
                Label18.Caption = Position
                Mtf.Direccion = File1.Path
                Mtf.Hora = "00:00"
                Total = File1.Path & "\" & List1.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    Next a
    For b = 1 To CantRad
        List3.Selected(Int((List3.ListCount * Rnd) + 1) - 1) = True
        ChkFl = Right$(List3.Text, 3)
        Select Case ChkFl
            Case "wav", "WAV", "Wav", "wAv", "waV"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List3.Text
                Label29.Caption = Position
                Mtf.Direccion = File3.Path
                Mtf.Hora = "00:00"
                Total = File3.Path & "\" & List3.Text
                'wHeadInfo (Total)    'extraemos la informacion del WAV
                'Mtf.Duracion = wInfo.wPlaytime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case "mp3", "MP3", "Mp3", "mP3"
                Position = Position + 1
                Mtf.id = Position
                Mtf.Name = List3.Text
                Label29.Caption = Position
                Mtf.Direccion = File3.Path
                Mtf.Hora = "00:00"
                Total = File3.Path & "\" & List3.Text
                ReadMP3Header (Total)        'extraemos la informacion del Mp3
                'Mtf.Duracion = MP3HInfo.FPlayTime
                'GetTime 'Mtf.Duracion
                Mtf.NameMix = "Sin Mix Intermedio"
                Mtf.DireccionMix = "---"
                Mtf.HoraMix = "00:00"
                'Mtf.DuracionMix = "00:00"
                Mtf.TotalDur = Lhh.Caption & ":" & Lmm.Caption & ":" & Lss.Caption
                Put #NuevoNumero, Position, Mtf
            Case Else
                'xxxx           'dejamos en blanco para que no agregue nada mas
        End Select
    Next b
Next i

TndCDuracion = CInt(Lhh.Caption)
If TndCDuracion >= TndHDuracion Then
    GoSub Finalizar
Else
    GoSub Restarting
End If
Exit Sub

'----------------------------------------------
Finalizar:
Close NuevoNumero
TndH.Caption = Lhh.Caption
TndM.Caption = Lmm.Caption
TndS.Caption = Lss.Caption
MsgBox "La Tanda " & UCase(NuevoNombre) & " ha sido generada satisfactoriamente.", vbInformation
Command15_Click
Exit Sub

'----------------------------------------------
oups:
MsgBox "Ha Ocurrido un error inesperado al intentar generar la Tanda. Por favor consulte con su proveedor de software.", vbCritical
Close
Exit Sub

End Sub

Private Sub Cad1Tema_Click()
If Cad1Tema.value = 1 Then
    Cad2Temas.value = 0
    Cad3Temas.value = 0
    Cad4Temas.value = 0
    ComAle.value = 0
    Comer2.Caption = "1"
End If

End Sub

Private Sub Cad2Temas_Click()
If Cad2Temas.value = 1 Then
    Cad1Tema.value = 0
    Cad3Temas.value = 0
    Cad4Temas.value = 0
    ComAle.value = 0
    Comer2.Caption = "2"
End If
End Sub


Private Sub Cad3Temas_Click()
If Cad3Temas.value = 1 Then
    Cad1Tema.value = 0
    Cad2Temas.value = 0
    Cad4Temas.value = 0
    ComAle.value = 0
    Comer2.Caption = "3"
End If

End Sub


Private Sub Cad4Temas_Click()
If Cad4Temas.value = 1 Then
    Cad1Tema.value = 0
    Cad2Temas.value = 0
    Cad3Temas.value = 0
    ComAle.value = 0
    Comer2.Caption = "4"
End If
End Sub


Private Sub Com1Cad_Click()
If Com1Cad.value = 1 Then
    Com2Cad.value = 0
    Com3Cad.value = 0
    ComAle.value = 0
    Comer1.Caption = "1"
End If

End Sub

Private Sub Com2Cad_Click()
If Com2Cad.value = 1 Then
    Com1Cad.value = 0
    Com3Cad.value = 0
    ComAle.value = 0
    Comer1.Caption = "2"
    End If
End Sub


Private Sub Com3Cad_Click()
If Com3Cad.value = 1 Then
    Com1Cad.value = 0
    Com2Cad.value = 0
    ComAle.value = 0
    Comer1.Caption = "3"
End If
End Sub


Private Sub ComAle_Click()
If ComAle.value = 1 Then
    Com1Cad.value = 0
    Com2Cad.value = 0
    Com3Cad.value = 0
    Cad1Tema.value = 0
    Cad2Temas.value = 0
    Cad3Temas.value = 0
    Cad4Temas.value = 0
    Com1Cad.Enabled = False
    Com2Cad.Enabled = False
    Com3Cad.Enabled = False
    Cad1Tema.Enabled = False
    Cad2Temas.Enabled = False
    Cad3Temas.Enabled = False
    Cad4Temas.Enabled = False
    Comer1.Caption = "0"
    Comer2.Caption = "0"
Else
    Com1Cad.Enabled = True
    Com2Cad.Enabled = True
    Com3Cad.Enabled = True
    Cad1Tema.Enabled = True
    Cad2Temas.Enabled = True
    Cad3Temas.Enabled = True
    Cad4Temas.Enabled = True
End If

End Sub

Private Sub Command1_Click()
List1.Clear
End Sub

Private Sub Command10_Click()
Dim MyFile
Dim Archivo
Dim Resultado
Dim CheckFile
Dim CheckFirst
Dim X

List3.Clear
Archivo = File3.Path & "\*.*"

MyFile = Dir(Archivo)

If MyFile = "" Or MyFile = " " Then
    Resultado = "No se econtró ningun archivo."
    List3.AddItem Resultado
    Exit Sub
Else
    CheckFirst = Right$(MyFile, 3)
    Select Case CheckFirst
        Case "", " "
            'xxx
        Case "wav", "Wav", "WAV", "wAv", "waV"
            Resultado = MyFile
            List3.AddItem Resultado
        Case "mp3", "Mp3", "MP3", "mP3"
            Resultado = MyFile
            List3.AddItem Resultado
        Case Else
            'xxx
    End Select
End If

For X = 1 To 999
    MyFile = Dir
    CheckFile = Right$(MyFile, 3)
    Select Case CheckFile
        Case "", " "
            GoSub Finish
        Case "wav", "Wav", "WAV", "wAv", "waV"
            Resultado = MyFile
            List3.AddItem Resultado
        Case "mp3", "Mp3", "MP3", "mP3"
            Resultado = MyFile
            List3.AddItem Resultado
        Case Else
            'xxx
    End Select
Next X
GoSub Finish
Exit Sub

Finish:
MsgBox "Ahora verifique los datos y genere la Tanda con el Generador de Tandas.", vbInformation
SSTab1.Tab = 4

Exit Sub
End Sub

Private Sub Command11_Click()

Dim ProgHelp

ProgHelp = App.Path + "\Ayuda.exe"
On Error GoTo ErrorInShell
Shell (ProgHelp), vbNormalFocus
Exit Sub

ErrorInShell:
MsgBox "Consulte el manual de usuario que se encuentra en el CD-ROM de Digital Audio Forge.", vbCritical
ErrorReporte "Error al intentar realizar el Shell del programa de Ayuda. Modulo Startup - Command7_Click"
Exit Sub


End Sub

Private Sub Command12_Click()
On Error GoTo er
List3.RemoveItem List3.ListIndex
Exit Sub

er:
MsgBox "Para eliminar un archivo de la lista primero deberá seleccionar el mismo.", vbCritical
End Sub

Private Sub Command13_Click()
Dim Nombre As String

'Abrimos...
CmD1.DialogTitle = "RM Generador de Tandas - Abrir configuración"
CmD1.InitDir = App.Path & AppConfigDir
CmD1.Filter = "Archivo de Generación (*.gen)|*.gen"
CmD1.FilterIndex = 1
CmD1.ShowOpen
Nombre = CmD1.FileName

'Abrimos la configuracion
On Error GoTo er
If Nombre = "" Then GoSub er
If Nombre = " " Then GoSub er
AbreConfiguracion Nombre
Exit Sub

er:
Resume Continue
Exit Sub

Continue:
End Sub

Private Sub Command14_Click()
Dim Nombre As String
Dim Datos As String
Dim Temas
Dim Comerciales
Dim ImagenR
Dim t1, t2, t3
Dim C1, C2, c3, c4, c5, c6, c7, c8
Dim r1, r2, r3, r4, r5, r6, r7, r8

'extraemos los datos definidos por el usuario
'Temas
If TemasOrAle.value = 1 Then
    t1 = "1"
Else
    t1 = "0"
End If

'Comerciales
If IncludComer.value = 1 Then
    t2 = "1"
Else
    t2 = "0"
End If
    If ComAle.value = 1 Then
        C1 = "1"
    Else
        C1 = "0"
    End If
    If Com1Cad.value = 1 Then
        C2 = "1"
    Else
        C2 = "0"
    End If
    If Com2Cad.value = 1 Then
        c3 = "1"
    Else
        c3 = "0"
    End If
    If Com3Cad.value = 1 Then
        c4 = "1"
    Else
        c4 = "0"
    End If
    If Cad1Tema.value = 1 Then
        c5 = "1"
    Else
        c5 = "0"
    End If
    If Cad2Temas.value = 1 Then
        c6 = "1"
    Else
        c6 = "0"
    End If
    If Cad3Temas.value = 1 Then
        c7 = "1"
    Else
        c7 = "0"
    End If
    If Cad4Temas.value = 1 Then
        c8 = "1"
    Else
        c8 = "0"
    End If

'ImagenRadio
If IncludRadio.value = 1 Then
    t3 = "1"
Else
    t3 = "0"
End If
    If RadAle.value = 1 Then
        r1 = "1"
    Else
        r1 = "0"
    End If
    If RadMixInter.value = 1 Then
        r2 = "1"
    Else
        r2 = "0"
    End If
    If Rad1Cad.value = 1 Then
        r3 = "1"
    Else
        r3 = "0"
    End If
    If Rad2Cad.value = 1 Then
        r4 = "1"
    Else
        r4 = "0"
    End If
    If Rad1Tema.value = 1 Then
        r5 = "1"
    Else
        r5 = "0"
    End If
    If Rad2Temas.value = 1 Then
        r6 = "1"
    Else
        r6 = "0"
    End If
    If Rad3Temas.value = 1 Then
        r7 = "1"
    Else
        r7 = "0"
    End If
    If Rad4Temas.value = 1 Then
        r8 = "1"
    Else
        r8 = "0"
    End If

'Formateamos los datos
Temas = t1 & t2 & t3
Comerciales = C1 & C2 & c3 & c4 & c5 & c6 & c7 & c8
ImagenR = r1 & r2 & r3 & r4 & r5 & r6 & r7 & r8
Datos = Temas & Comerciales & ImagenR

'Guardamos...
CmD1.DialogTitle = "RM Generador de Tandas - Guardar configuración"
CmD1.InitDir = App.Path & AppConfigDir
CmD1.Filter = "Archivo de Generación (*.gen)|*.gen"
CmD1.FilterIndex = 1
CmD1.ShowSave
Nombre = CmD1.FileName

'Guardamos la configuracion
On Error GoTo er
If Nombre = "" Then GoSub er
If Nombre = " " Then GoSub er
GuardaConfiguracion Nombre, Datos
Exit Sub

er:
Resume Continue
Exit Sub

Continue:
End Sub

Sub Command15_Click()
'Startup.Enabled = True
Unload Me
End Sub

Private Sub Command2_Click()
Dim MyFile
Dim Archivo
Dim Resultado
Dim CheckFile
Dim CheckFirst
Dim X

List1.Clear
Archivo = File1.Path & "\*.*"

MyFile = Dir(Archivo)

If MyFile = "" Or MyFile = " " Then
    Resultado = "No se econtró ningun archivo."
    List1.AddItem Resultado
    Exit Sub
Else
    CheckFirst = Right$(MyFile, 3)
    Select Case CheckFirst
        Case "", " "
            'xxx
        Case "wav", "Wav", "WAV", "wAv", "waV"
            Resultado = MyFile
            List1.AddItem Resultado
        Case "mp3", "Mp3", "MP3", "mP3"
            Resultado = MyFile
            List1.AddItem Resultado
        Case Else
            'xxx
    End Select
End If

For X = 1 To 999
    MyFile = Dir
    CheckFile = Right$(MyFile, 3)
    Select Case CheckFile
        Case "", " "
            GoSub Finish
        Case "wav", "Wav", "WAV", "wAv", "waV"
            Resultado = MyFile
            List1.AddItem Resultado
        Case "mp3", "Mp3", "MP3", "mP3"
            Resultado = MyFile
            List1.AddItem Resultado
        Case Else
            'xxx
    End Select
Next X
GoSub Finish
Exit Sub

Finish:
If IncludComer.value = 1 Then
    MsgBox "Ahora deberá seleccionar los Comerciales que desea incluir en la Tanda.", vbInformation
    SSTab1.Tab = 2
Else
    If IncludRadio.value = 1 Then
        MsgBox "Ahora deberá seleccionar los Institucionales que desee incluir en la Tanda.", vbInformation
        SSTab1.Tab = 3
    Else
        MsgBox "Ahora verifique los datos y genere la Tanda con el Generador de Tandas.", vbInformation
        SSTab1.Tab = 4
    End If
End If

Exit Sub
End Sub

Private Sub Command3_Click()
List2.Clear
End Sub

Private Sub Command4_Click()
On Error GoTo er
List1.RemoveItem List1.ListIndex
Exit Sub

er:
MsgBox "Para eliminar un archivo de la lista primero deberá seleccionar el mismo.", vbCritical
End Sub


Private Sub Command5_Click()
List3.Clear
End Sub

Private Sub Command6_Click()
Dim MyFile
Dim Archivo
Dim Resultado
Dim CheckFile
Dim CheckFirst
Dim X

List2.Clear
Archivo = File2.Path & "\*.*"

MyFile = Dir(Archivo)

If MyFile = "" Or MyFile = " " Then
    Resultado = "No se econtró ningun archivo."
    List2.AddItem Resultado
    Exit Sub
Else
    CheckFirst = Right$(MyFile, 3)
    Select Case CheckFirst
        Case "", " "
            'xxx
        Case "wav", "Wav", "WAV", "wAv", "waV"
            Resultado = MyFile
            List2.AddItem Resultado
        Case "mp3", "Mp3", "MP3", "mP3"
            Resultado = MyFile
            List2.AddItem Resultado
        Case Else
            'xxx
    End Select
End If

For X = 1 To 999
    MyFile = Dir
    CheckFile = Right$(MyFile, 3)
    Select Case CheckFile
        Case "", " "
            GoSub Finish
        Case "wav", "Wav", "WAV", "wAv", "waV"
            Resultado = MyFile
            List2.AddItem Resultado
        Case "mp3", "Mp3", "MP3", "mP3"
            Resultado = MyFile
            List2.AddItem Resultado
        Case Else
            'xxx
    End Select
Next X
GoSub Finish
Exit Sub

Finish:
If IncludRadio.value = 1 Then
    MsgBox "Ahora deberá seleccionar los Institucionales que desee incluir en la Tanda.", vbInformation
    SSTab1.Tab = 3
Else
    MsgBox "Ahora verifique los datos y genere la Tanda con el Generador de Tandas.", vbInformation
    SSTab1.Tab = 4
End If
Exit Sub

End Sub

Private Sub Command7_Click()

'Chequeamos las condiciones expuestas por el usuario
'y de acuerdo a ello...

TemasPorCom = Trim(Comer2.Caption)    'cada cuantos temas?
TemasPorRad = Trim(Radio2.Caption)    'cada cuantos intitucionales

If TemasOrAle.value = 1 Then
    If TemasPorCom = "-" Then
        If TemasPorRad = "-" Then
            Call SoloTemas
        Else
            If TemasPorRad = "0" Then
                Call TemasyRadAuto
            Else
                If TemasPorRad = "1" Then
                    Call TemasyRadCustom
                Else
                    Call TemasyRadCustom
                End If
            End If
        End If
    Else
        If TemasPorCom = "0" Then
            If TemasPorRad = "-" Then
                Call TemasyComAuto
            Else
                If TemasPorRad = "0" Then
                    Call TemasyComAutoyRadAuto
                Else
                    If TemasPorRad = "1" Then
                        Call TemasyComAutoyRadCustom
                    Else
                        Call TemasyComAutoyRadCustom
                    End If
                End If
            End If
        Else
            If TemasPorCom = "1" Then
                If TemasPorRad = "-" Then
                    Call TemasyComCustom
                Else
                    If TemasPorRad = "0" Then
                        Call TemasyComCustomyRadAuto
                    Else
                        If TemasPorRad = "1" Then
                            Call TemasyComCustomyRadCustom
                        Else
                            Call TemasyComCustomyRadCustom
                        End If
                    End If
                End If
            Else
                If TemasPorRad = "-" Then
                    Call TemasyComCustom
                Else
                    If TemasPorRad = "0" Then
                        Call TemasyComCustomyRadAuto
                    Else
                        If TemasPorRad = "1" Then
                            Call TemasyComCustomyRadCustom
                        Else
                            Call TemasyComCustomyRadCustom
                        End If
                    End If
                End If
            End If
        End If
    End If
Else
    MsgBox "No se puede iniciar la creación de Tandas sin antes haber seleccionado los temas que se incluirán en la misma.", vbCritical
    Exit Sub
End If

End Sub

Private Sub Command8_Click()
On Error GoTo er
List2.RemoveItem List2.ListIndex
Exit Sub

er:
MsgBox "Para eliminar un archivo de la lista primero deberá seleccionar el mismo.", vbCritical
End Sub

Private Sub Command9_Click()
'Startup.Enabled = True
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Pattern = "*.wav;*.Wav;*.WAV;*.mp3;*.Mp3;*.MP3"
End Sub

Private Sub Dir2_Change()
File2.Path = Dir2.Path
File2.Pattern = "*.wav;*.Wav;*.WAV;*.mp3;*.Mp3;*.MP3"
End Sub

Private Sub Dir3_Change()
File3.Path = Dir3.Path
File3.Pattern = "*.wav;*.Wav;*.WAV;*.mp3;*.Mp3;*.MP3"
End Sub

Private Sub Drive1_Change()
On Error GoTo er
Dir1.Path = Drive1.drive
Exit Sub

er:
End Sub
Private Sub Drive2_Change()
On Error GoTo er
Dir2.Path = Drive2.drive
Exit Sub

er:
End Sub

Private Sub Drive3_Change()
On Error GoTo er
Dir3.Path = Drive3.drive
Exit Sub

er:
End Sub

Private Sub EspHora_Change()
Dim Inthor
Dim Nl
Dim msg0
Inthor = 0
On Error GoTo IntError
Inthor = CInt(EspHora.Text)

'verificamos que el tiempo (en segundos) no sobrepase los 60
If Inthor > 24 Then
    Nl = Chr$(10) + Chr$(13)
    msg0 = "Las Tandas no pueden sobrepasar las 24Hs."
    MsgBox Nl + msg0 + Nl
    EspHora.Text = "00"
End If
Exit Sub

IntError:
Exit Sub

End Sub

Private Sub EspHora_GotFocus()
    EspHora.SelStart = 0
    EspHora.SelLength = Len(EspHora.Text)
End Sub

Private Sub Form_Load()

Dim AppMusicPath
Dim AppDrive

'Startup.Enabled = False
    
'Desabilitamos el panel de comerciales
    PanelComerciales1.Enabled = False
    PanelComerciales2.Enabled = False
    ComAle.Enabled = False
    Com1Cad.Enabled = False
    Com2Cad.Enabled = False
    Com3Cad.Enabled = False
    Cad1Tema.Enabled = False
    Cad2Temas.Enabled = False
    Cad3Temas.Enabled = False
    Cad4Temas.Enabled = False
    
'Desabilitamos el panel de imagen radio
    PanelRadio1.Enabled = False
    PanelRadio2.Enabled = False
    RadAle.Enabled = False
    RadMixInter.Enabled = False
    Rad1Cad.Enabled = False
    Rad2Cad.Enabled = False
    Rad1Tema.Enabled = False
    Rad2Temas.Enabled = False
    Rad3Temas.Enabled = False
    Rad4Temas.Enabled = False

    IncludComer.Enabled = False
    IncludRadio.Enabled = False

TemasDsc.Caption = ""

'extraemos el directorio definido por el usuario como Maestro
'AppMusicPath = DataForm.Dr1.Caption

'verificamos su validez
If AppMusicPath = "" Or AppMusicPath = " " Then
    MsgBox "No se ha configurado un directorio Maestro para el manejo de los temas.", vbCritical
    MsgBox "Se utilizará el directorio por defecto de la aplicación.", vbCritical
    AppMusicPath = Left$(App.Path, 2) & AppDefaultMusicPath
    AppDrive = Left$(AppMusicPath, 2) & "\"
Else
    'AppMusicPath = DataForm.Dr1.Caption
    AppDrive = Left$(AppMusicPath, 2) & "\"
End If

File1.Pattern = "*.wav;*.Wav;*.WAV;*.mp3;*.Mp3;*.MP3"
Drive1.drive = AppDrive
Dir1.Path = AppMusicPath

File2.Pattern = "*.wav;*.Wav;*.WAV;*.mp3;*.Mp3;*.MP3"
Drive2.drive = AppDrive
Dir2.Path = AppMusicPath

File3.Pattern = "*.wav;*.Wav;*.WAV;*.mp3;*.Mp3;*.MP3"
Drive3.drive = AppDrive
Dir3.Path = AppMusicPath

End Sub

Private Sub IncludComer_Click()
If IncludComer.value = 1 Then
    PanelComerciales1.Enabled = True
    PanelComerciales2.Enabled = True
    ComAle.Enabled = True
    Com1Cad.Enabled = True
    Com2Cad.Enabled = True
    Com3Cad.Enabled = True
    Cad1Tema.Enabled = True
    Cad2Temas.Enabled = True
    Cad3Temas.Enabled = True
    Cad4Temas.Enabled = True
    Comer1.Caption = "0"
    Comer2.Caption = "0"
Else
    PanelComerciales1.Enabled = False
    PanelComerciales2.Enabled = False
    
    ComAle.value = 0
    Com1Cad.value = 0
    Com2Cad.value = 0
    Com3Cad.value = 0
    Cad1Tema.value = 0
    Cad2Temas.value = 0
    Cad3Temas.value = 0
    Cad4Temas.value = 0

    ComAle.Enabled = False
    Com1Cad.Enabled = False
    Com2Cad.Enabled = False
    Com3Cad.Enabled = False
    Cad1Tema.Enabled = False
    Cad2Temas.Enabled = False
    Cad3Temas.Enabled = False
    Cad4Temas.Enabled = False
    Comer1.Caption = "-"
    Comer2.Caption = "-"
End If

End Sub

Private Sub IncludRadio_Click()
If IncludRadio.value = 1 Then
    PanelRadio1.Enabled = True
    PanelRadio2.Enabled = True
    RadAle.Enabled = True
    RadMixInter.Enabled = True
    Rad1Cad.Enabled = True
    Rad2Cad.Enabled = True
    Rad1Tema.Enabled = True
    Rad2Temas.Enabled = True
    Rad3Temas.Enabled = True
    Rad4Temas.Enabled = True
    Radio1.Caption = "0"
    Radio2.Caption = "0"
Else
    PanelRadio1.Enabled = False
    PanelRadio2.Enabled = False
    
    RadAle.value = 0
    RadMixInter.value = 0
    Rad1Cad.value = 0
    Rad2Cad.value = 0
    Rad1Tema.value = 0
    Rad2Temas.value = 0
    Rad3Temas.value = 0
    Rad4Temas.value = 0

    RadAle.Enabled = False
    RadMixInter.Enabled = False
    Rad1Cad.Enabled = False
    Rad2Cad.Enabled = False
    Rad1Tema.Enabled = False
    Rad2Temas.Enabled = False
    Rad3Temas.Enabled = False
    Rad4Temas.Enabled = False
    Radio1.Caption = "-"
    Radio2.Caption = "-"
End If

End Sub

Private Sub Rad1Cad_Click()
If Rad1Cad.value = 1 Then
    Rad2Cad.value = 0
    RadAle.value = 0
    Radio1.Caption = "1"
End If
End Sub

Private Sub Rad1Tema_Click()
If Rad1Tema.value = 1 Then
    Rad2Temas.value = 0
    Rad3Temas.value = 0
    Rad4Temas.value = 0
    RadAle.value = 0
    Radio2.Caption = "1"
End If

End Sub

Private Sub Rad2Cad_Click()
If Rad2Cad.value = 1 Then
    Rad1Cad.value = 0
    RadAle.value = 0
    Radio1.Caption = "2"
End If

End Sub

Private Sub Rad2Temas_Click()
If Rad2Temas.value = 1 Then
    Rad1Tema.value = 0
    Rad3Temas.value = 0
    Rad4Temas.value = 0
    RadAle.value = 0
    Radio2.Caption = "2"

End If
End Sub

Private Sub Rad3Temas_Click()
If Rad3Temas.value = 1 Then
    Rad1Tema.value = 0
    Rad2Temas.value = 0
    Rad4Temas.value = 0
    RadAle.value = 0
    Radio2.Caption = "3"

End If

End Sub

Private Sub Rad4Temas_Click()
If Rad4Temas.value = 1 Then
    Rad1Tema.value = 0
    Rad2Temas.value = 0
    Rad3Temas.value = 0
    RadAle.value = 0
    Radio2.Caption = "4"

End If

End Sub

Private Sub RadAle_Click()
If RadAle.value = 1 Then
    RadMixInter.value = 0
    Rad1Cad.value = 0
    Rad2Cad.value = 0
    Rad1Tema.value = 0
    Rad2Temas.value = 0
    Rad3Temas.value = 0
    Rad4Temas.value = 0
    RadMixInter.Enabled = False
    Rad1Cad.Enabled = False
    Rad2Cad.Enabled = False
    Rad1Tema.Enabled = False
    Rad2Temas.Enabled = False
    Rad3Temas.Enabled = False
    Rad4Temas.Enabled = False
    Radio1.Caption = "0"
    Radio2.Caption = "0"
Else
    RadMixInter.Enabled = True
    Rad1Cad.Enabled = True
    Rad2Cad.Enabled = True
    Rad1Tema.Enabled = True
    Rad2Temas.Enabled = True
    Rad3Temas.Enabled = True
    Rad4Temas.Enabled = True
End If
    
End Sub

Private Sub RadMixInter_Click()
If RadMixInter.value = 1 Then
    Rad2Cad.value = 0
    Rad2Cad.Enabled = False
    RadAle.value = 0
Else
    Rad2Cad.Enabled = True
End If

End Sub


Private Sub RestoreAll_Click()
    ComAle.value = 0
    Com1Cad.value = 0
    Com2Cad.value = 0
    Com3Cad.value = 0
    Cad1Tema.value = 0
    Cad2Temas.value = 0
    Cad3Temas.value = 0
    Cad4Temas.value = 0
    
    RadAle.value = 0
    RadMixInter.value = 0
    Rad1Cad.value = 0
    Rad2Cad.value = 0
    Rad1Tema.value = 0
    Rad2Temas.value = 0
    Rad3Temas.value = 0
    Rad4Temas.value = 0
    IncludRadio.value = 0
    IncludComer.value = 0
    TemasOrAle.value = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

If SSTab1.Tab = 0 Then
    Command13.Enabled = True
    Command14.Enabled = True
    RestoreAll.Enabled = True
Else
    Command13.Enabled = False
    Command14.Enabled = False
    RestoreAll.Enabled = False
End If

'Incluir Temas (obligatorio)
If TemasOrAle.value = 0 Then
    MsgBox "No ha seleccionado si desea incluir Temas en la tanda. Esta opción no puede ser pasada por alto. Por favor verifique la configuracion.", vbCritical
    SSTab1.Tab = 0
End If

'Incluir comerciales
If SSTab1.Tab = 2 And IncludComer.value = 0 Then
    MsgBox "No puede acceder a Comerciales porque en configuración Ud. especificó que no desea incluir comerciales.", vbCritical
    SSTab1.Tab = 0
End If

'Incluir Intitucionales
If SSTab1.Tab = 3 And IncludRadio.value = 0 Then
    MsgBox "No puede acceder a Institucionales porque en configuración Ud. especificó que no desea incluir los Institucionales.", vbCritical
    SSTab1.Tab = 0
End If

'Comerciales
If IncludComer.value = 1 And ComAle.value = 0 And Com1Cad.value = 0 And Com2Cad.value = 0 And Com3Cad.value = 0 Then
    MsgBox "Antes de continuar deberá especificar el tipo de órden que desea darle a los comerciales. Por favor verifique la configuración.", vbCritical
    SSTab1.Tab = 0
End If
If Com1Cad.value = 1 And Cad1Tema.value = 0 And Cad2Temas.value = 0 And Cad3Temas.value = 0 And Cad4Temas.value = 0 Then
    MsgBox "Antes de continuar deberá especificar cada cuántos temas desea Ud. incluir los comerciales. Cada 1 tema?, Cada 2 temas?... Por favor verifique la configuración.", vbCritical
    SSTab1.Tab = 0
End If
If Com2Cad.value = 1 And Cad1Tema.value = 0 And Cad2Temas.value = 0 And Cad3Temas.value = 0 And Cad4Temas.value = 0 Then
    MsgBox "Antes de continuar deberá especificar cada cuántos temas desea Ud. incluir los comerciales. Cada 1 tema?, Cada 2 temas?... Por favor verifique la configuración.", vbCritical
    SSTab1.Tab = 0
End If
If Com3Cad.value = 1 And Cad1Tema.value = 0 And Cad2Temas.value = 0 And Cad3Temas.value = 0 And Cad4Temas.value = 0 Then
    MsgBox "Antes de continuar deberá especificar cada cuántos temas desea Ud. incluir los comerciales. Cada 1 tema?, Cada 2 temas?... Por favor verifique la configuración.", vbCritical
    SSTab1.Tab = 0
End If

'Institucionales
If IncludRadio.value = 1 And RadAle.value = 0 And RadMixInter.value = 0 And Rad1Cad.value = 0 And Rad2Cad.value = 0 Then
    MsgBox "Antes de continuar deberá especificar que tipo de órden desea darle a los Institucionales. Por favor verifique la configuración.", vbCritical
    SSTab1.Tab = 0
End If

If Rad1Cad.value = 1 And Rad1Tema.value = 0 And Rad2Temas.value = 0 And Rad3Temas.value = 0 And Rad4Temas.value = 0 Then
    MsgBox "Antes de continuar deberá especificar cada cuántos temas desea incluir los Institucionales. Cada 1 tema?, cada 2 temas?... Por favor verifique la configuración.", vbCritical
    SSTab1.Tab = 0
End If
If Rad2Cad.value = 1 And Rad1Tema.value = 0 And Rad2Temas.value = 0 And Rad3Temas.value = 0 And Rad4Temas.value = 0 Then
    MsgBox "Antes de continuar deberá especificar cada cuántos temas desea incluir los Institucionales. Cada 1 tema?, cada 2 temas?... Por favor verifique la configuración.", vbCritical
    SSTab1.Tab = 0
End If
If RadMixInter.value = 1 And Rad1Cad.value = 0 And Rad1Tema.value = 0 And Rad2Temas.value = 0 And Rad3Temas.value = 0 And Rad4Temas.value = 0 Then
    MsgBox "Antes de continuar deberá especificar cada cuántos temas desea incluir los Institucionales. Cada 1 tema?, cada 2 temas?... Por favor verifique la configuración.", vbCritical
    SSTab1.Tab = 0
End If

End Sub

Private Sub TemasOrAle_Click()
If TemasOrAle.value = 1 Then
    IncludComer.Enabled = True
    IncludRadio.Enabled = True
    TemasDsc.Caption = "Los temas seleccionados se ordenarán aleatoriamente para componer una tanda."
Else
    TemasDsc.Caption = ""
    IncludComer.value = 0
    IncludComer.Enabled = False
    IncludRadio.value = 0
    IncludRadio.Enabled = False
End If
End Sub


