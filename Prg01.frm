VERSION 5.00
Begin VB.Form Prg01 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "PROGRAMACION DE TANDAS - Detenido"
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   7875
   ControlBox      =   0   'False
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox p1t1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6075
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   525
      Width           =   190
   End
   Begin VB.PictureBox p1t2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6270
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   525
      Width           =   190
   End
   Begin VB.PictureBox p1t3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6450
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   525
      Width           =   190
   End
   Begin VB.PictureBox p1t4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6645
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   525
      Width           =   190
   End
   Begin VB.PictureBox p1t5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6840
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   525
      Width           =   190
   End
   Begin VB.PictureBox p1t6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   7020
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   525
      Width           =   190
   End
   Begin VB.PictureBox p1t7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   7200
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   525
      Width           =   190
   End
   Begin VB.PictureBox p1t8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   7380
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   525
      Width           =   190
   End
   Begin VB.TextBox TxtRename 
      Height          =   435
      Left            =   1155
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   6210
      Visible         =   0   'False
      Width           =   1170
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   0
      Left            =   450
      TabIndex        =   38
      Top             =   870
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   1
      Left            =   4230
      TabIndex        =   39
      Top             =   870
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   2
      Left            =   450
      TabIndex        =   40
      Top             =   1140
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   3
      Left            =   4230
      TabIndex        =   41
      Top             =   1140
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   4
      Left            =   450
      TabIndex        =   42
      Top             =   1410
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   5
      Left            =   4230
      TabIndex        =   43
      Top             =   1410
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   6
      Left            =   450
      TabIndex        =   44
      Top             =   1680
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   7
      Left            =   4230
      TabIndex        =   45
      Top             =   1680
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   8
      Left            =   450
      TabIndex        =   46
      Top             =   1950
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   9
      Left            =   4230
      TabIndex        =   47
      Top             =   1950
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   10
      Left            =   450
      TabIndex        =   48
      Top             =   2220
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   11
      Left            =   4230
      TabIndex        =   49
      Top             =   2220
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   12
      Left            =   450
      TabIndex        =   50
      Top             =   2490
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   13
      Left            =   4230
      TabIndex        =   51
      Top             =   2490
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   14
      Left            =   450
      TabIndex        =   52
      Top             =   2760
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   15
      Left            =   4230
      TabIndex        =   53
      Top             =   2760
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   16
      Left            =   450
      TabIndex        =   54
      Top             =   3030
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   17
      Left            =   4230
      TabIndex        =   55
      Top             =   3030
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   18
      Left            =   450
      TabIndex        =   56
      Top             =   3300
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   19
      Left            =   4230
      TabIndex        =   57
      Top             =   3300
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   20
      Left            =   450
      TabIndex        =   58
      Top             =   3570
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   21
      Left            =   4230
      TabIndex        =   59
      Top             =   3570
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   22
      Left            =   450
      TabIndex        =   60
      Top             =   3840
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.DC_Control_Bt Prg1 
      Height          =   255
      Index           =   23
      Left            =   4230
      TabIndex        =   61
      Top             =   3840
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   450
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
      ForeColor       =   -2147483633
   End
   Begin RM100.TitelBar TitelBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   661
      BackColor       =   8421504
      BackColorCover  =   5
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
      Caption         =   " PROGRAMACION DE TANDAS - Detenido"
      CaptionPosX     =   1
      BorderNormal    =   2
      BorderColorHighLight=   0
      BorderColorDarkLight=   12632256
   End
   Begin RM100.DC_Control_Bt P1New 
      Height          =   585
      Left            =   5970
      TabIndex        =   63
      ToolTipText     =   "Nueva paginación"
      Top             =   4230
      Width           =   525
      _ExtentX        =   926
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
      ForeColor       =   -2147483633
      MaskColor       =   4210752
      PicDown         =   "Prg01.frx":0000
      PicHot          =   "Prg01.frx":5A5B2
      PicNormal       =   "Prg01.frx":B4B64
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin RM100.DC_Control_Bt P1Play 
      Height          =   585
      Left            =   480
      TabIndex        =   64
      Top             =   4230
      Width           =   735
      _ExtentX        =   1296
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
      ForeColor       =   -2147483633
      MaskColor       =   4210752
      PicDown         =   "Prg01.frx":10F116
      PicHot          =   "Prg01.frx":1696C8
      PicNormal       =   "Prg01.frx":1C3C7A
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
   End
   Begin RM100.DC_Control_Bt P1Pause 
      Height          =   585
      Left            =   1290
      TabIndex        =   65
      Top             =   4230
      Width           =   735
      _ExtentX        =   1296
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
      MaskColor       =   4210752
      PicDown         =   "Prg01.frx":21E22C
      PicHot          =   "Prg01.frx":2787DE
      PicNormal       =   "Prg01.frx":2D2D90
      PicSizeH        =   32
      PicSizeW        =   32
   End
   Begin RM100.DC_Control_Bt P1Stop 
      Height          =   585
      Left            =   2100
      TabIndex        =   66
      Top             =   4230
      Width           =   735
      _ExtentX        =   1296
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
      ForeColor       =   -2147483633
      MaskColor       =   4210752
      PicDown         =   "Prg01.frx":32D342
      PicHot          =   "Prg01.frx":3878F4
      PicNormal       =   "Prg01.frx":3E1EA6
      PicSizeH        =   32
      PicSizeW        =   32
   End
   Begin RM100.DC_Control_Bt P1Open 
      Height          =   585
      Left            =   6570
      TabIndex        =   67
      ToolTipText     =   "Abrir Paginación"
      Top             =   4230
      Width           =   525
      _ExtentX        =   926
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
      ForeColor       =   -2147483633
      MaskColor       =   4210752
      PicDown         =   "Prg01.frx":43C458
      PicHot          =   "Prg01.frx":496A0A
      PicNormal       =   "Prg01.frx":4F0FBC
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin RM100.DC_Control_Bt P1Save 
      Height          =   585
      Left            =   7170
      TabIndex        =   68
      ToolTipText     =   "Guardar paginación"
      Top             =   4230
      Width           =   525
      _ExtentX        =   926
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
      ForeColor       =   -2147483633
      MaskColor       =   4210752
      PicDown         =   "Prg01.frx":54B56E
      PicHot          =   "Prg01.frx":5A5B20
      PicNormal       =   "Prg01.frx":6000D2
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
   End
   Begin VB.Label Fn 
      BackColor       =   &H00808000&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3420
      TabIndex        =   37
      Top             =   6240
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL Dur:"
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   5145
      TabIndex        =   35
      Top             =   525
      Width           =   915
   End
   Begin VB.Label LblName 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Programación 1 - Sin Nombre.prg"
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   240
      TabIndex        =   24
      ToolTipText     =   "Nombre de archivo"
      Top             =   510
      Width           =   4770
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   150
      TabIndex        =   36
      Top             =   450
      Width           =   7560
   End
   Begin VB.Label Lindex 
      BackColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   2460
      TabIndex        =   26
      Top             =   6225
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "24"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   23
      Left            =   3945
      TabIndex        =   23
      Top             =   3870
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "22"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   22
      Left            =   3945
      TabIndex        =   22
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "23"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   21
      Left            =   165
      TabIndex        =   21
      Top             =   3870
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "21"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   20
      Left            =   165
      TabIndex        =   20
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "20"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   19
      Left            =   3945
      TabIndex        =   19
      Top             =   3330
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "18"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   18
      Left            =   3945
      TabIndex        =   18
      Top             =   3060
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "16"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   17
      Left            =   3945
      TabIndex        =   17
      Top             =   2790
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "14"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   16
      Left            =   3945
      TabIndex        =   16
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "12"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   15
      Left            =   3945
      TabIndex        =   15
      Top             =   2250
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "10"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   14
      Left            =   3945
      TabIndex        =   14
      Top             =   1980
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "8"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   13
      Left            =   3945
      TabIndex        =   13
      Top             =   1710
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "6"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   12
      Left            =   3945
      TabIndex        =   12
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "4"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   11
      Left            =   3945
      TabIndex        =   11
      Top             =   1170
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "2"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   10
      Left            =   3945
      TabIndex        =   10
      Top             =   900
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "19"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   9
      Left            =   165
      TabIndex        =   9
      Top             =   3330
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "17"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   8
      Left            =   165
      TabIndex        =   8
      Top             =   3060
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "15"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   7
      Left            =   165
      TabIndex        =   7
      Top             =   2790
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "13"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   6
      Left            =   165
      TabIndex        =   6
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "11"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   5
      Left            =   165
      TabIndex        =   5
      Top             =   2250
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "9"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   4
      Left            =   165
      TabIndex        =   4
      Top             =   1980
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "7"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   3
      Left            =   165
      TabIndex        =   3
      Top             =   1710
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "5"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   2
      Left            =   165
      TabIndex        =   2
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "3"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   1
      Left            =   165
      TabIndex        =   1
      Top             =   1170
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   900
      Width           =   255
   End
End
Attribute VB_Name = "Prg01"
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

Sub DeployPRGFile(WContNum As Integer)

'PF=tanda filename
'PC=tanda name or caption
'PD=tanda duracion

Dim IDX As Integer
IDX = WContNum

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
        MsgBox LoadResString(190), vbInformation, "Radio Maker"
        
    'MUSIC TYPE XM-MOD-S3M-IT-MTM-MO3-UMX
    Case "XM", "MOD", "S3M", "IT", "MTM", "MO3", "UMX"
        MsgBox LoadResString(190), vbInformation, "Radio Maker"
        
    'TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TNDTND
    Case "TND"
        Est12Data.PF(IDX).Caption = Completo
        Est12Data.PC(IDX).Caption = FileN
        'Est12Data.PD(IdX).Caption = xxxxx duracion va aqui
        Prg1(IDX).Caption = FileN
        Prg1(IDX).BackColor = &H8000000F
        'Prg1(IdX).ToolTipText = "Duración: " & xxxxx duracion va aqui
        
    Case Else
        MsgBox LoadResString(190), vbInformation, "Radio Maker"

End Select

End Sub

Private Sub E1Play_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

HideWindow "Prg01"

End Sub

Private Sub Form_Terminate()

HideWindow "Prg01"

End Sub

Private Sub Form_Unload(Cancel As Integer)

HideWindow "Prg01"

End Sub

Private Sub P1New_Click()

Dim i As Integer

For i = 0 To 23
    Fn.Caption = ""
    Prg1(i).Caption = ""
    Prg1(i).ToolTipText = ""
    Est12Data.PF(i).Caption = ""
    Est12Data.PC(i).Caption = ""
    Est12Data.PD(i).Caption = ""
    LblName.Caption = "Programación 1 - Sin Nombre.prg"
    LblName.ForeColor = &H808000
Next i

End Sub

Private Sub P1Open_Click()

Dim lnText As Long
Dim NewName As String

On Error Resume Next
TopMenu.ProgCmd.InitDir = App.path & AppProgDir & "\"
TopMenu.ProgCmd.Filter = "Programación de Tandas (*.prg)|*.prg|Programación de Tandas"
TopMenu.ProgCmd.DialogTitle = "Programación de Tandas - Abrir archivo"
TopMenu.ProgCmd.CancelError = True
TopMenu.ProgCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

    'restauramos los valores a 0
    Fn.Caption = ""
    Call RestoreDisplay(10)

ConvertTx = TopMenu.ProgCmd.filename

Result = OpenPrgFile(ConvertTx)
If Result = "NotOK" Then
    'MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
    Exit Sub
End If

Fn.Caption = ConvertTx
lnText = Len(ConvertTx)
If lnText > 60 Then
    NewName = Left$(ConvertTx, 3) & " ... " & Right$(ConvertTx, 50)
Else
    NewName = ConvertTx
End If
LblName.Caption = NewName
LblName.ForeColor = &HFFFF00    'verde claro

End Sub

Sub P1Save_Click()

Dim lnText As Long
Dim NewName As String

ConvertTxT = Trim(Fn.Caption)

On Error Resume Next
If ConvertTxT = "" Or ConvertTxT = " " Then
    TopMenu.ProgCmd.InitDir = App.path & AppProgDir & "\"
    TopMenu.ProgCmd.Filter = "Programacion de Tandas (*.prg)|*.prg|Programación de Tandas"
    TopMenu.ProgCmd.DialogTitle = "Programacion de Tandas - Guardar archivo"
    TopMenu.ProgCmd.FilterIndex = 1
    TopMenu.ProgCmd.CancelError = True
    TopMenu.ProgCmd.ShowSave

    If err.Number = 32755 Then Exit Sub
    
    ConvertTx = TopMenu.ProgCmd.filename

    Fn.Caption = ConvertTx
    Result = SavePrgFile(ConvertTx)
    If Result = "NotOK" Then
        MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
        Exit Sub
    End If
Else
    ConvertTx = Trim(Fn.Caption)
    Kill ConvertTx
    Result = SavePrgFile(ConvertTx)
    If Result = "NotOK" Then
        'MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
        Exit Sub
    End If
End If

Fn.Caption = ConvertTx
lnText = Len(ConvertTx)
If lnText > 60 Then
    NewName = Left$(ConvertTx, 3) & " ... " & Right$(ConvertTx, 50)
Else
    NewName = ConvertTx
End If

LblName.Caption = NewName
LblName.ForeColor = &HFFFF00    'verde claro

End Sub

Private Sub Prg1_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)

DeployPRGFile index   'drag & drop the selected file in xplorer

End Sub

Private Sub Prg1_DragOver(index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

Select Case State
    Case 0  'drag not finished
        XPlorer.File1.DragIcon = XPlorer.ExCombo.DragIcon
        Prg1(index).BackColor = &H80FF80    'verde (modificacion)
    Case 1  'finished drag
        XPlorer.File1.DragIcon = XPlorer.tvwDirTree.DragIcon
        Prg1(index).BackColor = &H8000000F  'gris (normal)
End Select

End Sub

Private Sub Prg1_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'button 1=left button
'button 2=right button
'button 4=mid button

If Button = 2 Then
    If Prg1(index).Caption = "" Or Prg1(index).Caption = " " Then
        Exit Sub
    End If
    'deploy options menu
    TxtRename.Visible = True
    TxtRename.Top = Prg1(index).Top
    TxtRename.Left = Prg1(index).Left
    TxtRename.Height = Prg1(index).Height
    TxtRename.Width = Prg1(index).Width
    TxtRename.text = Prg1(index).Caption
    TxtRename.SetFocus
    'seteamos el label para saber de que control se trata
    Lindex.Caption = index
Else
    If Button = 4 Then
        'mark control to delete content
        Prg1(index).BackColor = &HFF&        'rojo
    Else
        'nothing to do
    End If
End If

End Sub

Private Sub TxtRename_KeyPress(KeyAscii As Integer)

Dim IDX As Integer

If KeyAscii = 13 Then   'ENTER
    IDX = CInt(Lindex.Caption)
    Prg1(IDX).Caption = TxtRename.text
    Est12Data.PC(IDX).Caption = TxtRename.text
    TxtRename.Visible = False
End If
If KeyAscii = 27 Or KeyAscii = 13 Then 'ESCAPE or ENTER
    TxtRename.Visible = False
End If

End Sub
