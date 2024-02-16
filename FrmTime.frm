VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form FrmTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programación Horaria"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9840
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton PHNew 
      Height          =   375
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   136
      ToolTipText     =   "Nuevo archivo Horario"
      Top             =   5730
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton PHOpen 
      Height          =   375
      Left            =   540
      Style           =   1  'Graphical
      TabIndex        =   135
      ToolTipText     =   "Abrir archivo de programación horaria"
      Top             =   5730
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton PHSave 
      Height          =   375
      Left            =   945
      Style           =   1  'Graphical
      TabIndex        =   134
      ToolTipText     =   "Guardar archivo de programación horaria"
      Top             =   5730
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   23
      Left            =   5400
      TabIndex        =   133
      ToolTipText     =   "Remover"
      Top             =   5175
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   22
      Left            =   5400
      TabIndex        =   132
      ToolTipText     =   "Remover"
      Top             =   4815
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   21
      Left            =   5400
      TabIndex        =   131
      ToolTipText     =   "Remover"
      Top             =   4455
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   20
      Left            =   5400
      TabIndex        =   130
      ToolTipText     =   "Remover"
      Top             =   4095
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   19
      Left            =   5400
      TabIndex        =   129
      ToolTipText     =   "Remover"
      Top             =   3735
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   18
      Left            =   5400
      TabIndex        =   128
      ToolTipText     =   "Remover"
      Top             =   3375
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   17
      Left            =   5400
      TabIndex        =   127
      ToolTipText     =   "Remover"
      Top             =   3015
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   16
      Left            =   5400
      TabIndex        =   126
      ToolTipText     =   "Remover"
      Top             =   2655
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   15
      Left            =   5400
      TabIndex        =   125
      ToolTipText     =   "Remover"
      Top             =   2295
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   14
      Left            =   5400
      TabIndex        =   124
      ToolTipText     =   "Remover"
      Top             =   1935
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   13
      Left            =   5400
      TabIndex        =   123
      ToolTipText     =   "Remover"
      Top             =   1575
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   12
      Left            =   5400
      TabIndex        =   122
      ToolTipText     =   "Remover"
      Top             =   1215
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   11
      Left            =   405
      TabIndex        =   121
      ToolTipText     =   "Remover"
      Top             =   5175
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   10
      Left            =   405
      TabIndex        =   120
      ToolTipText     =   "Remover"
      Top             =   4815
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   9
      Left            =   405
      TabIndex        =   119
      ToolTipText     =   "Remover"
      Top             =   4455
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   8
      Left            =   405
      TabIndex        =   118
      ToolTipText     =   "Remover"
      Top             =   4095
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   7
      Left            =   405
      TabIndex        =   117
      ToolTipText     =   "Remover"
      Top             =   3735
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   6
      Left            =   405
      TabIndex        =   116
      ToolTipText     =   "Remover"
      Top             =   3375
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   5
      Left            =   405
      TabIndex        =   115
      ToolTipText     =   "Remover"
      Top             =   3015
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   4
      Left            =   405
      TabIndex        =   114
      ToolTipText     =   "Remover"
      Top             =   2655
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   3
      Left            =   405
      TabIndex        =   113
      ToolTipText     =   "Remover"
      Top             =   2295
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   2
      Left            =   405
      TabIndex        =   112
      ToolTipText     =   "Remover"
      Top             =   1935
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   1
      Left            =   405
      TabIndex        =   111
      ToolTipText     =   "Remover"
      Top             =   1575
      Width           =   285
   End
   Begin VB.CommandButton RmvItm 
      Caption         =   "R"
      Height          =   330
      Index           =   0
      Left            =   405
      TabIndex        =   110
      ToolTipText     =   "Remover"
      Top             =   1215
      Width           =   285
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   645
      Left            =   105
      TabIndex        =   103
      Top             =   150
      Width           =   9645
      Begin ComctlLib.Slider SldVel 
         Height          =   195
         Left            =   7830
         TabIndex        =   138
         Top             =   330
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   344
         _Version        =   327682
      End
      Begin ComctlLib.Slider SldVol 
         Height          =   255
         Left            =   3960
         TabIndex        =   137
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   450
         _Version        =   327682
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         Height          =   240
         Left            =   7470
         TabIndex        =   109
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         Height          =   240
         Left            =   9135
         TabIndex        =   108
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         Height          =   240
         Left            =   3600
         TabIndex        =   107
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         Height          =   240
         Left            =   5265
         TabIndex        =   106
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Al Activar, bajar el volumen de la musica en un:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   135
         TabIndex        =   105
         Top             =   270
         Width           =   3390
      End
      Begin VB.Label Label9 
         Caption         =   "A una velocidad del:"
         Height          =   195
         Left            =   5940
         TabIndex        =   104
         Top             =   270
         Width           =   1500
      End
   End
   Begin VB.Timer T2VOut 
      Left            =   6480
      Top             =   7560
   End
   Begin VB.Timer T2VIn 
      Left            =   5985
      Top             =   7560
   End
   Begin VB.Timer T1VOut 
      Left            =   6480
      Top             =   7065
   End
   Begin VB.Timer T1VIn 
      Left            =   5985
      Top             =   7065
   End
   Begin VB.Timer E2Vout 
      Left            =   5175
      Top             =   7560
   End
   Begin VB.Timer E2VIn 
      Left            =   4680
      Top             =   7560
   End
   Begin VB.Timer E1VOut 
      Left            =   5175
      Top             =   7065
   End
   Begin VB.Timer E1Vin 
      Left            =   4680
      Top             =   7065
   End
   Begin VB.Timer PHTimer 
      Left            =   1800
      Top             =   7110
   End
   Begin VB.CommandButton PHCancel 
      Caption         =   "Cc"
      Height          =   375
      Left            =   8445
      TabIndex        =   74
      ToolTipText     =   "Cancelar"
      Top             =   5730
      Width           =   1275
   End
   Begin VB.CommandButton PHDesactive 
      Caption         =   "D"
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
      Left            =   2925
      TabIndex        =   73
      ToolTipText     =   "Desactivar programación horaria"
      Top             =   5730
      Width           =   1365
   End
   Begin VB.CommandButton PHActive 
      Caption         =   "A"
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
      Left            =   1440
      TabIndex        =   72
      ToolTipText     =   "Activar programación horaria"
      Top             =   5730
      Width           =   1365
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   23
      Left            =   9135
      MaxLength       =   5
      TabIndex        =   71
      Text            =   "00:00"
      Top             =   5175
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   22
      Left            =   9135
      MaxLength       =   5
      TabIndex        =   68
      Text            =   "00:00"
      Top             =   4815
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   21
      Left            =   9135
      MaxLength       =   5
      TabIndex        =   65
      Text            =   "00:00"
      Top             =   4455
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   20
      Left            =   9135
      MaxLength       =   5
      TabIndex        =   62
      Text            =   "00:00"
      Top             =   4095
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   23
      Left            =   8190
      TabIndex        =   70
      Top             =   5175
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   22
      Left            =   8190
      TabIndex        =   67
      Top             =   4815
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   21
      Left            =   8190
      TabIndex        =   64
      Top             =   4455
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   20
      Left            =   8190
      TabIndex        =   61
      Top             =   4095
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   69
      Top             =   5175
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   4815
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   21
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   4455
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   4095
      Width           =   2400
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   19
      Left            =   9135
      MaxLength       =   5
      TabIndex        =   59
      Text            =   "00:00"
      Top             =   3735
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   18
      Left            =   9135
      MaxLength       =   5
      TabIndex        =   56
      Text            =   "00:00"
      Top             =   3375
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   17
      Left            =   9135
      MaxLength       =   5
      TabIndex        =   53
      Text            =   "00:00"
      Top             =   3015
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   16
      Left            =   9135
      MaxLength       =   5
      TabIndex        =   50
      Text            =   "00:00"
      Top             =   2655
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   15
      Left            =   9135
      MaxLength       =   5
      TabIndex        =   47
      Text            =   "00:00"
      Top             =   2295
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   14
      Left            =   9135
      MaxLength       =   5
      TabIndex        =   44
      Text            =   "00:00"
      Top             =   1935
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   13
      Left            =   9135
      MaxLength       =   5
      TabIndex        =   41
      Text            =   "00:00"
      Top             =   1575
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   12
      Left            =   9135
      MaxLength       =   5
      TabIndex        =   38
      Text            =   "00:00"
      Top             =   1215
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   11
      Left            =   4140
      MaxLength       =   5
      TabIndex        =   35
      Text            =   "00:00"
      Top             =   5175
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   10
      Left            =   4140
      MaxLength       =   5
      TabIndex        =   32
      Text            =   "00:00"
      Top             =   4815
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   9
      Left            =   4140
      MaxLength       =   5
      TabIndex        =   29
      Text            =   "00:00"
      Top             =   4455
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   8
      Left            =   4140
      MaxLength       =   5
      TabIndex        =   26
      Text            =   "00:00"
      Top             =   4095
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   7
      Left            =   4140
      MaxLength       =   5
      TabIndex        =   23
      Text            =   "00:00"
      Top             =   3735
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   6
      Left            =   4140
      MaxLength       =   5
      TabIndex        =   20
      Text            =   "00:00"
      Top             =   3375
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   4140
      MaxLength       =   5
      TabIndex        =   17
      Text            =   "00:00"
      Top             =   3015
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   4140
      MaxLength       =   5
      TabIndex        =   14
      Text            =   "00:00"
      Top             =   2655
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   4140
      MaxLength       =   5
      TabIndex        =   11
      Text            =   "00:00"
      Top             =   2295
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   4140
      MaxLength       =   5
      TabIndex        =   8
      Text            =   "00:00"
      Top             =   1935
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   4140
      MaxLength       =   5
      TabIndex        =   5
      Text            =   "00:00"
      Top             =   1575
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   4140
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "00:00"
      Top             =   1215
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   19
      Left            =   8190
      TabIndex        =   58
      Top             =   3735
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   18
      Left            =   8190
      TabIndex        =   55
      Top             =   3375
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   17
      Left            =   8190
      TabIndex        =   52
      Top             =   3015
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   16
      Left            =   8190
      TabIndex        =   49
      Top             =   2655
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   15
      Left            =   8190
      TabIndex        =   46
      Top             =   2295
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   14
      Left            =   8190
      TabIndex        =   43
      Top             =   1935
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   13
      Left            =   8190
      TabIndex        =   40
      Top             =   1575
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   12
      Left            =   8190
      TabIndex        =   37
      Top             =   1215
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   11
      Left            =   3195
      TabIndex        =   34
      Top             =   5175
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   10
      Left            =   3195
      TabIndex        =   31
      Top             =   4815
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   9
      Left            =   3195
      TabIndex        =   28
      Top             =   4455
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   8
      Left            =   3195
      TabIndex        =   25
      Top             =   4095
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   7
      Left            =   3195
      TabIndex        =   22
      Top             =   3735
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   6
      Left            =   3195
      TabIndex        =   19
      Top             =   3375
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   5
      Left            =   3195
      TabIndex        =   16
      Top             =   3015
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   4
      Left            =   3195
      TabIndex        =   13
      Top             =   2655
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   3
      Left            =   3195
      TabIndex        =   10
      Top             =   2295
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   2
      Left            =   3195
      TabIndex        =   7
      Top             =   1935
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   1
      Left            =   3195
      TabIndex        =   4
      Top             =   1575
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "e"
      Height          =   330
      Index           =   0
      Left            =   3195
      TabIndex        =   1
      Top             =   1215
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   3735
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   3375
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   3015
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   2655
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   2295
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   1935
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   1575
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   1215
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   5175
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   4815
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   4455
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   4095
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3735
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3375
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3015
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2655
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2295
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1935
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1575
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1215
      Width           =   2400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   9690
      X2              =   105
      Y1              =   5625
      Y2              =   5625
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   9690
      X2              =   105
      Y1              =   5610
      Y2              =   5610
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   23
      Left            =   5130
      TabIndex        =   102
      Top             =   5220
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   22
      Left            =   5130
      TabIndex        =   101
      Top             =   4860
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   21
      Left            =   5130
      TabIndex        =   100
      Top             =   4500
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   20
      Left            =   5130
      TabIndex        =   99
      Top             =   4140
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   19
      Left            =   5130
      TabIndex        =   98
      Top             =   3780
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   18
      Left            =   5130
      TabIndex        =   97
      Top             =   3420
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   5130
      TabIndex        =   96
      Top             =   3060
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   16
      Left            =   5130
      TabIndex        =   95
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   15
      Left            =   5130
      TabIndex        =   94
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   14
      Left            =   5130
      TabIndex        =   93
      Top             =   1980
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   13
      Left            =   5130
      TabIndex        =   92
      Top             =   1620
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   5130
      TabIndex        =   91
      Top             =   1260
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   135
      TabIndex        =   90
      Top             =   5220
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   135
      TabIndex        =   89
      Top             =   4860
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   135
      TabIndex        =   88
      Top             =   4500
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   135
      TabIndex        =   87
      Top             =   4140
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   135
      TabIndex        =   86
      Top             =   3780
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   135
      TabIndex        =   85
      Top             =   3420
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   135
      TabIndex        =   84
      Top             =   3060
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   135
      TabIndex        =   83
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   135
      TabIndex        =   82
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   135
      TabIndex        =   81
      Top             =   1980
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   135
      TabIndex        =   80
      Top             =   1620
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   135
      TabIndex        =   79
      Top             =   1260
      Width           =   240
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Lanz."
      Height          =   195
      Left            =   9135
      TabIndex        =   78
      Top             =   945
      Width           =   555
   End
   Begin VB.Label Label3 
      Caption         =   "Archivos horarios:"
      Height          =   195
      Left            =   5715
      TabIndex        =   77
      Top             =   945
      Width           =   2445
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Lanz."
      Height          =   195
      Left            =   4140
      TabIndex        =   76
      Top             =   945
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Archivos horarios:"
      Height          =   195
      Left            =   720
      TabIndex        =   75
      Top             =   945
      Width           =   2400
   End
End
Attribute VB_Name = "FrmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

On Error Resume Next
TopMenu.WaveCmd.InitDir = App.Path & AppDefaultMusicPath
TopMenu.WaveCmd.Filter = "Archivos de Audio (*.wav; *.mp1; *.mp2; *.mp3)|*.wav; *.mp1; *.mp2; *.mp3|Todos los archivos de Audio"
TopMenu.WaveCmd.DialogTitle = "Programacion Horaria - Abrir archivo de Audio"
TopMenu.WaveCmd.CancelError = True
TopMenu.WaveCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

Text1(Index).Text = TopMenu.WaveCmd.filename
Text1(Index).BackColor = &HC0FFFF
Text2(Index).SetFocus
Text2(Index).BackColor = &HC0FFFF

End Sub

Private Sub E1Vin_Timer()

'FADEIN FOR EST01 ONLY

If Est01.E1Vol.Value = 100 Then
    E1Vin.Interval = 0
    E1Vin.Enabled = False
    Exit Sub
Else
    Est01.E1Vol.Value = Est01.E1Vol.Value + 1
End If

End Sub

Private Sub E1VOut_Timer()

'FADE OUT FOR EST01 ONLY

Dim VVal As Integer
VVal = SldVol.Value

If Est01.E1Vol.Value = VVal Then
    E1VOut.Interval = 0
    E1VOut.Enabled = False
Else
    Est01.E1Vol.Value = Est01.E1Vol.Value - 1
End If

End Sub

Private Sub E2VIn_Timer()

'FADEIN FOR EST02 ONLY

If Est02.E2Vol.Value = 100 Then
    E2VIn.Interval = 0
    E2VIn.Enabled = False
    Exit Sub
Else
    Est02.E2Vol.Value = Est02.E2Vol.Value + 1
End If

End Sub

Private Sub E2Vout_Timer()

'FADE OUT FOR EST02 ONLY

Dim VVal As Integer
VVal = SldVol.Value

If Est02.E2Vol.Value = VVal Then
    E2Vout.Interval = 0
    E2Vout.Enabled = False
Else
    Est02.E2Vol.Value = Est02.E2Vol.Value - 1
End If

End Sub

Private Sub Form_Load()

Dim i As Integer

    '--- strings to load
    PHCancel.Caption = LoadResString(2010)
    PHActive.Caption = LoadResString(2011)
    PHDesactive.Caption = LoadResString(2012)
    For i = 0 To 23
        Command1(i).Caption = LoadResString(2002)
    Next i
    
    '--- more icons to load...
    PHNew.Picture = LoadResPicture("ICO_NEW", 0)
    PHOpen.Picture = LoadResPicture("ICO_OPEN", 0)
    PHSave.Picture = LoadResPicture("ICO_SAVE", 0)

'activar el AutoOpen
'para abrir el ultimo archivo PH utilizado

End Sub

Sub PHActive_Click()

'chequeos de displays
If TopMenu.Label2.Caption = "Desactivada" Then
    TopMenu.Label2.Caption = "Activada"
    TopMenu.Label2.ForeColor = &HFFFF00
End If

'codigo para activacion de ph
Dim NFile   'nombre de archivo
Dim Ltime   'hora de lanzamiento
Dim Atime   'hora actual
Dim Lh, Lm  'hora y minuto de lanzamiento
Dim Ah, Am  'hora y minuto actual
Dim Ni As Integer   'numero de index del ultimo textbox usado
Dim i As Integer    'numero de index para el conteo de controles

Ni = CInt(TopMenu.NumberIdx.Caption)
If Ni >= 23 Then
    'se termino el proceso del ph
    If TopMenu.WindowState = 1 Then
        TopMenu.WindowState = 0
        If TopMenu.Label2.Caption = "Activada" Then
            TopMenu.Label2.Caption = "Desactivada"
            TopMenu.Label2.ForeColor = &H808000
        End If
    Else
        If TopMenu.Label2.Caption = "Activada" Then
            TopMenu.Label2.Caption = "Desactivada"
            TopMenu.Label2.ForeColor = &H808000
        End If
    End If
    'se procede a la desactivacion del mismo
    PHTimer.Interval = 0
    PHTimer.Enabled = False
    'se resetean todos los datos
    If FrmTime.WindowState = 1 Then
        FrmTime.Visible = True
        FrmTime.WindowState = 0 'normal
    End If
    TopMenu.PHName.Caption = "---"
    TopMenu.PHName.ForeColor = &H808000 'desactivado
    TopMenu.PHTime = ""
    Call RestoreDisplay(6)   'default display TIME in PHTIMER
    TopMenu.NumberIdx.Caption = "0"
    'y se cierra la ventana del ph
    Unload Me
End If

For i = Ni To 23
    NFile = Text1(i).Text
    Ltime = Text2(i).Text & ":00"
    Atime = time$
    Lh = Left$(Ltime, 2): Lm = Mid$(Ltime, 4, 2)
    Ah = Left$(Atime, 2): Am = Mid$(Atime, 4, 2)
    If NFile = "" Or NFile = " " Then
        'nothing to do, go to the next textbox
    Else
        If Ah > Lh Then
            'nothing to do, go to the next textbox
        Else
            If Ah = Lh Then
                If Am >= Lm Then
                    'nothing to do, go to the next textbox
                Else
                    TopMenu.PHName.Caption = Text1(i).Text
                    TopMenu.PHName.ForeColor = &HFFFF00 'activada
                    TopMenu.PHTime.Caption = Text2(i).Text & ":00"
                    SetTOPTime (Text2(i).Text & ":00")
                    TopMenu.NumberIdx.Caption = i
                    'activamos el reloj de control de PH
                    PHTimer.Enabled = True
                    PHTimer.Interval = 1000
                    'terminamos la activacion
                    GoSub Active
                    Exit For
                End If
            Else
                TopMenu.PHName.Caption = Text1(i).Text
                TopMenu.PHName.ForeColor = &HFFFF00 'activada
                TopMenu.PHTime.Caption = Text2(i).Text & ":00"
                SetTOPTime (Text2(i).Text & ":00")
                TopMenu.NumberIdx.Caption = i
                'activamos el reloj de control de PH
                PHTimer.Enabled = True
                PHTimer.Interval = 1000
                'terminamos la activacion
                GoSub Active
                Exit For
            End If
        End If
    End If
Next i

Call PHDesactive_Click
Exit Sub

Active:
'continue con la activacion
If FrmTime.WindowState = 0 Then
    FrmTime.WindowState = 1 'minimizada
    FrmTime.Visible = False
End If

End Sub

Private Sub PHCancel_Click()

Unload Me

End Sub

Sub PHDesactive_Click()

If TopMenu.WindowState = 1 Then
    TopMenu.WindowState = 0
End If

'chequeos de displays
If TopMenu.Label2.Caption = "Activada" Then
    TopMenu.Label2.Caption = "Desactivada"
    TopMenu.Label2.ForeColor = &H808000
End If

'se resetean todos los datos
If FrmTime.WindowState = 1 Then
    FrmTime.Visible = True
    FrmTime.WindowState = 0 'normal
End If

TopMenu.PHName.Caption = "---"
TopMenu.PHName.ForeColor = &H808000 'desactivado
TopMenu.PHTime = ""
Call RestoreDisplay(6)   'default display TIME in PHTIMER
TopMenu.NumberIdx.Caption = "0"

End Sub

Private Sub PHNew_Click()

Dim Cont As Integer
'reset the form
For Cont = 0 To 23
    Text1(Cont).Text = ""
    Text1(Cont).BackColor = &HFFFFFF
    Text2(Cont).Text = "00:00"
    Text2(Cont).BackColor = &HFFFFFF
Next Cont

End Sub

Private Sub PHOpen_Click()

On Error Resume Next
TopMenu.PHCmd.InitDir = App.Path & AppPHDir
TopMenu.PHCmd.Filter = "Archivo PH (*.ph1)|*.ph1|Archivos de Programación Horaria"
TopMenu.PHCmd.DialogTitle = "Programación Horaria - Abrir archivo"
TopMenu.PHCmd.CancelError = True
TopMenu.PHCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

Result = OpenPHFile(TopMenu.PHCmd.filename)
If Result = "NotOK" Then
    MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
    Exit Sub
End If

End Sub

Private Sub PHSave_Click()

On Error Resume Next
TopMenu.PHCmd.InitDir = App.Path & AppPHDir
TopMenu.PHCmd.Filter = "Archivo PH (*.ph1)|*.ph1|Archivos de Programación horaria"
TopMenu.PHCmd.DialogTitle = "Programacion Horaria - Guardar archivo"
TopMenu.PHCmd.FilterIndex = 1
TopMenu.PHCmd.CancelError = True
TopMenu.PHCmd.ShowSave

If err.Number = 32755 Then Exit Sub

Result = SavePHFile(TopMenu.PHCmd.filename)
If Result = "NotOK" Then
    MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
    Exit Sub
End If

End Sub

Private Sub PHTimer_Timer()

Dim Oldi As Integer
Dim Newi As Integer

Dim TMClock As String
Dim Result As String

TMClock = time$

If TopMenu.Label2.Caption = "Activada" And TopMenu.Label2.ForeColor = &HFFFF00 Then
    'activar la programacion horaria
    If TopMenu.PHTime.Caption = TMClock Then
        Call PHPlay
    Else
        'nothing to do... just wait for time = clock
    End If
Else
    PHTimer.Interval = 0
    PHTimer.Enabled = False
End If

End Sub

Private Sub RmvItm_Click(Index As Integer)

Text1(Index).Text = ""
Text1(Index).BackColor = &HFFFFFF
Text2(Index).Text = "00:00"
Text2(Index).BackColor = &HFFFFFF

End Sub

Private Sub SldVel_Change()

SldVel.ToolTipText = SldVel.Value & "%"

End Sub

Private Sub SldVel_Scroll()

SldVel.ToolTipText = SldVel.Value & "%"

End Sub

Private Sub SldVol_Change()

SldVol.ToolTipText = SldVol.Value & "%"

End Sub

Private Sub SldVol_Scroll()

SldVol.ToolTipText = SldVol.Value & "%"

End Sub

Private Sub T1VIn_Timer()

'FADEIN FOR TANDA01 ONLY

If Tanda01.T1Vol.Value = 100 Then
    T1VIn.Interval = 0
    T1VIn.Enabled = False
    Exit Sub
Else
    Tanda01.T1Vol.Value = Tanda01.T1Vol.Value + 1
End If

End Sub

Private Sub T1VOut_Timer()

'FADE OUT FOR TANDA01 ONLY

Dim VVol As Integer
VVol = SldVol.Value

If Tanda01.T1Vol.Value = VVol Then
    T1VOut.Interval = 0
    T1VOut.Enabled = False
Else
    Tanda01.T1Vol.Value = Tanda01.T1Vol.Value - 1
End If

End Sub

Private Sub T2VIn_Timer()

'FADEIN FOR TANDA02 ONLY

If Tanda01.T2Vol.Value = 100 Then
    T2VIn.Interval = 0
    T2VIn.Enabled = False
    Exit Sub
Else
    Tanda01.T2Vol.Value = Tanda01.T2Vol.Value + 1
End If

End Sub

Private Sub T2VOut_Timer()

'FADE OUT FOR TANDA02 ONLY

Dim VVol As Integer
VVol = SldVol.Value

If Tanda01.T2Vol.Value = VVol Then
    T2VOut.Interval = 0
    T2VOut.Enabled = False
Else
    Tanda01.T2Vol.Value = Tanda01.T2Vol.Value - 1
End If

End Sub

Private Sub Text2_Change(Index As Integer)

If Text1(Index).Text = "" Or Text1(Index).Text = " " Then
    Text2(Index).Text = "00:00"
End If

End Sub

Private Sub Text2_GotFocus(Index As Integer)

Text2(Index).SelStart = 0
Text2(Index).SelLength = Len(Text2(Index).Text)

End Sub

Private Sub Text2_LostFocus(Index As Integer)

Dim LenCheck
LenCheck = Len(Text2(Index).Text)

'check the len for validations
If LenCheck < 5 Then
    MsgBox "Hora de Lanzamiento no válida. Utilice un formato de 24hs. (nó AM/PM).", vbInformation
    Text2(Index).SetFocus
    Exit Sub
End If
If LenCheck > 5 Then
    MsgBox "Hora de Lanzamiento no válida. Utilice un formato de 24hs. (nó AM/PM).", vbInformation
    Text2(Index).SetFocus
    Exit Sub
End If

Dim Hora As String
Dim Minutos As String

'extraemos los datos de hora especificados para el lanzamiento
Hora = Left$(Text2(Index).Text, 2)
Minutos = Right$(Text2(Index).Text, 2)

'Procedemos al chequeo de la misma
On Error Resume Next
If LenCheck = 5 Then
    If Hora > 23 Or Hora < 0 Then
        If Hora = "00" Then
            'xxx
        Else
            MsgBox "La hora de Lanzamiento especificada: " & Hora & " no es válida. Utilice un formato de 24hs. (nó AM/PM).", vbInformation
            Text2(Index).SetFocus
            Exit Sub
        End If
    Else
        'la hora esta bien. chequeamos los minutos.
        If Minutos > 59 Then
            If Minutos = "00" Then
                'xxx
            Else
                MsgBox "Los minutos de Lanzamiento especificados: " & Minutos & " no son válidos. Utilice un formato de 24hs. (nó AM/PM).", vbInformation
                Text2(Index).SetFocus
                Exit Sub
            End If
        Else
            'xxx
        End If
    End If
Else
    MsgBox "Hora de Lanzamiento no válida. Utilice un formato de 24hs. (nó AM/PM).", vbInformation
    Text2(Index).SetFocus
    Exit Sub
End If

End Sub
