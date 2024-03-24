VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form Tanda01 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   7440
   ClientLeft      =   15
   ClientTop       =   -30
   ClientWidth     =   7635
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7440
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar Prbar1 
      Height          =   285
      Left            =   150
      TabIndex        =   76
      Top             =   5910
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   8790
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Slider T1Vol 
      Height          =   255
      Left            =   1860
      TabIndex        =   74
      Top             =   8730
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   10
      Max             =   100
      SelStart        =   100
      TickFrequency   =   10
      Value           =   100
   End
   Begin MSComctlLib.ListView T1View 
      Height          =   4485
      Left            =   120
      TabIndex        =   73
      Top             =   1380
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   7911
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a1"
         Text            =   "1"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "a2"
         Text            =   "2"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "a3"
         Text            =   "3"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "a4"
         Text            =   "4"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "a5"
         Text            =   "5"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "a6"
         Text            =   "6"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "a7"
         Text            =   "7"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "a8"
         Text            =   "8"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "a9"
         Text            =   "9"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "a10"
         Text            =   "10"
         Object.Width           =   706
      EndProperty
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   6900
      Max             =   0
      Min             =   10
      TabIndex        =   67
      Top             =   570
      Value           =   3
      Width           =   135
   End
   Begin VB.CommandButton CmdBlock 
      Height          =   375
      Left            =   2250
      Picture         =   "Tanda01.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Agregar / Eliminar / modificar bloques"
      Top             =   6945
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.PictureBox T1F8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   7260
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   63
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1F7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   7080
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1F6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6900
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   61
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1F5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6720
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1F4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6525
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   59
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1F3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6330
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   58
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1F2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6150
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   57
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1F1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   5955
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   56
      TabStop         =   0   'False
      ToolTipText     =   "Hora de FINALIZACION de la Tanda en reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1I8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4650
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1I7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4470
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1I6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4290
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1I5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4110
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   51
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1I4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3915
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1I3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3720
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   49
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1I2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3540
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1I1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3345
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   47
      TabStop         =   0   'False
      ToolTipText     =   "Hora de INICIO de reproducción"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1t1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   825
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "Duración TOTAL de la Tanda"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1t2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1020
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "Duración TOTAL de la Tanda"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1t3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1200
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   "Duración TOTAL de la Tanda"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1t4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1395
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   "Duración TOTAL de la Tanda"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1t5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1590
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Duración TOTAL de la Tanda"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1t6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1770
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Duración TOTAL de la Tanda"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1t7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1950
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "Duración TOTAL de la Tanda"
      Top             =   6360
      Width           =   190
   End
   Begin VB.PictureBox T1t8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2130
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   "Duración TOTAL de la Tanda"
      Top             =   6360
      Width           =   190
   End
   Begin VB.CommandButton T1OrderA 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5490
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Reordenar / actualizar tiempo desde el tema seleccionado hacia el final de la lista"
      Top             =   6945
      UseMaskColor    =   -1  'True
      Width           =   800
   End
   Begin VB.CommandButton T1Order 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Reordenar / actualizar tiempo desde el comienzo de la Lista"
      Top             =   6945
      UseMaskColor    =   -1  'True
      Width           =   800
   End
   Begin VB.CommandButton T1Stop 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1650
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Detener"
      Top             =   6945
      UseMaskColor    =   -1  'True
      Width           =   500
   End
   Begin VB.CommandButton T1Play 
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Reproducir seleccionado"
      Top             =   6945
      UseMaskColor    =   -1  'True
      Width           =   700
   End
   Begin VB.CommandButton T1Next 
      Enabled         =   0   'False
      Height          =   375
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Reproducir continuo"
      Top             =   6945
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.PictureBox T1p6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3360
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   870
      Width           =   190
   End
   Begin VB.PictureBox T1p0 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3360
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   585
      Width           =   190
   End
   Begin VB.Timer SyncTimer 
      Left            =   2220
      Top             =   9555
   End
   Begin VB.TextBox Intr 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   6630
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   26
      Text            =   "2"
      Top             =   600
      Width           =   405
   End
   Begin VB.CommandButton T1Del 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Eliminar"
      Top             =   6945
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton T1Down 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3435
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Bajar"
      Top             =   6945
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton T1Up 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3075
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Subir"
      Top             =   6945
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Timer TmOut2 
      Left            =   5955
      Top             =   8925
   End
   Begin VB.Timer TmIn2 
      Left            =   5460
      Top             =   8925
   End
   Begin VB.Timer TmOut1 
      Left            =   5955
      Top             =   8430
   End
   Begin VB.Timer TmIn1 
      Left            =   5460
      Top             =   8430
   End
   Begin VB.CommandButton T1Save 
      Height          =   375
      Left            =   7095
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Guardar Tanda"
      Top             =   6945
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton T1Open 
      Height          =   375
      Left            =   6735
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Abrir Tanda"
      Top             =   6945
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton T1New 
      Height          =   375
      Left            =   6375
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Nueva Tanda"
      Top             =   6945
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton T1Prop 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Propiedades de audio"
      Top             =   6945
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.PictureBox T1p7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3570
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   870
      Width           =   190
   End
   Begin VB.PictureBox T1p8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3765
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   870
      Width           =   190
   End
   Begin VB.PictureBox T1p9 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3945
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   870
      Width           =   190
   End
   Begin VB.PictureBox T1p10 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4140
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   870
      Width           =   190
   End
   Begin VB.PictureBox T1p11 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4335
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   870
      Width           =   190
   End
   Begin VB.PictureBox T1p1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3570
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   585
      Width           =   190
   End
   Begin VB.PictureBox T1p2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3765
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   585
      Width           =   190
   End
   Begin VB.PictureBox T1p3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3945
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   585
      Width           =   190
   End
   Begin VB.PictureBox T1p4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4140
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   585
      Width           =   190
   End
   Begin VB.PictureBox T1p5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   4335
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   585
      Width           =   190
   End
   Begin RM100.TitelBar TitelBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   661
      BackColor       =   8421504
      BackColorCover  =   3
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
      Caption         =   "  TANDAS - Detenido"
      CaptionPosX     =   1
      BorderNormal    =   2
      BorderColorHighLight=   0
      BorderColorDarkLight=   4210752
   End
   Begin MSComctlLib.Slider T2Vol 
      Height          =   255
      Left            =   1830
      TabIndex        =   75
      Top             =   9180
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   10
      Max             =   100
      SelStart        =   100
      TickFrequency   =   10
      Value           =   100
   End
   Begin VB.Label BlkFn 
      BackColor       =   &H00FFFF00&
      Height          =   195
      Left            =   2220
      TabIndex        =   71
      Top             =   8310
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "no definido.blk"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   5550
      TabIndex        =   70
      Top             =   900
      Width           =   1275
   End
   Begin VB.Label LBlk 
      BackStyle       =   0  'Transparent
      Caption         =   "/ Man"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   6900
      TabIndex        =   69
      Top             =   900
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bloque:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   4890
      TabIndex        =   68
      Top             =   900
      Width           =   585
   End
   Begin VB.Label FTime 
      BackColor       =   &H00FF8080&
      Caption         =   "---"
      Height          =   240
      Left            =   3120
      TabIndex        =   65
      Top             =   7800
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "FIN:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   5610
      TabIndex        =   64
      Top             =   6360
      Width           =   315
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "INICIO:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2775
      TabIndex        =   55
      Top             =   6360
      Width           =   555
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   180
      TabIndex        =   46
      Top             =   6360
      Width           =   585
   End
   Begin VB.Label SyncStream 
      BackColor       =   &H000080FF&
      Height          =   240
      Left            =   2715
      TabIndex        =   33
      Top             =   9645
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label SyncLabel 
      BackColor       =   &H000040C0&
      Caption         =   "0"
      Height          =   240
      Left            =   4650
      TabIndex        =   32
      Top             =   9645
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Ltime 
      BackColor       =   &H00FF8080&
      Caption         =   "00:00:00"
      Height          =   240
      Left            =   2220
      TabIndex        =   31
      Top             =   7800
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Dev-2:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   210
      TabIndex        =   30
      Top             =   870
      Width           =   510
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dev-1:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   210
      TabIndex        =   29
      Top             =   600
      Width           =   510
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "segs."
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   7065
      TabIndex        =   28
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inter:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   6165
      TabIndex        =   27
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F-In/Out:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   4890
      TabIndex        =   25
      Top             =   600
      Width           =   690
   End
   Begin VB.Label LFin 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   5655
      TabIndex        =   24
      Top             =   600
      Width           =   375
   End
   Begin VB.Label LKey 
      BackColor       =   &H0080FF80&
      Caption         =   "1"
      Height          =   240
      Left            =   3900
      TabIndex        =   23
      Top             =   8070
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Fn 
      BackColor       =   &H00008000&
      Height          =   210
      Left            =   2220
      TabIndex        =   22
      Top             =   8070
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Shape T1Shape 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   120
      Top             =   6885
      Width           =   7410
   End
   Begin VB.Label T2Name 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   705
      TabIndex        =   21
      Top             =   870
      Width           =   2580
   End
   Begin VB.Label T1Name 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   705
      TabIndex        =   20
      Top             =   600
      Width           =   2580
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   4830
      Stretch         =   -1  'True
      Top             =   420
      Width           =   2700
   End
   Begin VB.Image Image2 
      Height          =   885
      Left            =   75
      Stretch         =   -1  'True
      Top             =   420
      Width           =   4635
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   60
      Stretch         =   -1  'True
      Top             =   6270
      Width           =   7530
   End
   Begin VB.Menu BlockMnu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu Blockmenu_Comandos 
         Caption         =   "&Insertar comandos"
         Begin VB.Menu Blockmenu_Comandos_HM 
            Caption         =   "Reproducir Hora y minutos"
            Enabled         =   0   'False
         End
         Begin VB.Menu Blockmenu_Comandos_TH 
            Caption         =   "Reproducir Temperatura y humedad"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu BlockMnu_sep0 
         Caption         =   "-"
      End
      Begin VB.Menu BlockMnu_define 
         Caption         =   "&Definir bloque de utilización..."
         Enabled         =   0   'False
      End
      Begin VB.Menu BlockMnu_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu BlockMnu_insert 
         Caption         =   "&Insertar bloque..."
         Enabled         =   0   'False
      End
      Begin VB.Menu BlockMnu_delete 
         Caption         =   "&Eliminar bloque"
         Enabled         =   0   'False
      End
      Begin VB.Menu BlockMnu_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu BlockMnu_config 
         Caption         =   "&Configurar bloques..."
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "Tanda01"
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

'dimensiones de listitem
Dim ItmZ As ListItem
Dim ItmX As ListItem

Dim TxtKey As String
Dim NewKey As String
Dim ANum As Integer
Dim ONum As Long
Dim NNum As Long
Dim nIndex As Integer

'dimensiones Tanda time check
Dim LzTime As String
Dim AcTime As String
Dim LzTh1, LzTm1, LzTs1 As Integer
Dim AcTh1, AcTm1, AcTs1 As Integer

Sub DeployTandaFile()

Dim OTime As String
Dim NTime As String
Dim RTime As String
Dim Time1 As Double
Dim Time2 As Double
Dim TMint As Double
Dim Resultado As Double
Dim FileNTag As String
Dim StrVal As String, StrVal2 As String

If XPlorer.File1.filename = "" Or XPlorer.File1.filename = " " Then
    MsgBox LoadResString(137), vbCritical
    Exit Sub
End If

'.wav, .mp3, .it, .xm
'.wav, .mp3, .it, .xm
FileExt = StripExtFromFile(XPlorer.File1.filename)
FileN = StripFileFromExt(XPlorer.File1.filename)
FileNPath = Trim(XPlorer.lblPath)
Completo = Trim(XPlorer.lblPath) & "\" & XPlorer.File1.filename

If GetMP3Tag(Completo) = True Then
    StrVal = Replace(Trim(MP3Info.sArtist), Chr(0), "")
    StrVal2 = Replace(Trim(MP3Info.sTitle), Chr(0), "")
    FileNTag = StrVal & " - " & StrVal2
    'Debug.Print FileNTag
    'Dim i As Integer
    'For i = 1 To Len(MP3Info.sArtist)
    '    Debug.Print Mid(MP3Info.sArtist, i, 1) & "   Ascii =  " & Asc(Mid(MP3Info.sArtist, i, 1))
    'Next
Else
    FileNTag = FileN
End If

'seleccion de formato de archivo y extraccion de informacion header
Select Case Trim(UCase(FileExt))
    
    'STREAM TYPE WAV-MP1-MP2-MP3-OGG
    Case "WAV", "MP1", "MP2", "MP3", "OGG"
        ONum = T1View.ListItems.count
        NNum = ONum + 1
        TxtKey = "r"
        NewKey = TxtKey & NNum
        
        Set ItmZ = T1View.ListItems.Add(NNum, NewKey, Completo) 'path & file
        ItmZ.SubItems(1) = "Stream"     'file type
        ItmZ.SubItems(2) = FileNTag        'file name
        'gets the file len and convert into time
        ConvertTx = FileLoadLen(Completo, "Stream")
        TimeNcv = FormatSegs(ConvertTx)
        Result = ConvSecToMin(CInt(TimeNcv))
        'refresh the time display
            OTime = Trim(Tanda01.Ltime.Caption)
            NTime = Trim(Result)
            Time1 = ConvMinToSec(OTime)
            Time2 = ConvMinToSec(NTime)
            'tiempo de mixado intermedio
            TMint = CDbl(Trim(Tanda01.Intr.Text))
            Resultado = Time1 + Time2
            Resultado = (Resultado - TMint) + 1
            RTime = ConvSecToMin(Resultado)
            SetSumTime RTime, 1
            Tanda01.Ltime.Caption = RTime
        'put the rest of info
        ItmZ.SubItems(3) = Result      'duracion del tema
        ItmZ.SubItems(4) = "00:00:00"  'poner aqui la hora de lanzamiento
        ItmZ.SubItems(5) = "-----"     'poner aqui el path & file del mixado
        ItmZ.SubItems(6) = "-----"     'poner aqui el type del mixado
        ItmZ.SubItems(7) = "-----"     'poner aqui el nombre del mixado interm.
        ItmZ.SubItems(8) = "00:00:00"     'poner aqui la duracion del mixado
        ItmZ.SubItems(9) = "00:00"  'poner aqui la hora de lanzam. del mixado
        'Completo                  'nombre y path
        'FileN                     'nombre solo
            
    'MUSIC TYPE XM-MOD-S3M-IT-MTM-MO3-UMX
    Case "XM", "MOD", "S3M", "IT", "MTM", "MO3", "UMX"
        MsgBox LoadResString(184), vbInformation, "Radio Maker"
        
    'TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TNDTND
    Case "tnd", "Tnd", "tNd", "tnD", "TNd", "TND", "tND"
        MsgBox LoadResString(184), vbInformation, "Radio Maker"
        
    Case Else
        MsgBox LoadResString(184), vbInformation, "Radio Maker"

End Select

End Sub

Private Sub Blockmenu_Comandos_HM_Click()

    ONum = T1View.ListItems.count
    NNum = ONum + 1
    TxtKey = "r"
    NewKey = TxtKey & NNum
    
    Completo = "-----"
    Set ItmZ = T1View.ListItems.Add(NNum, NewKey, Completo) 'path & file

    'Set ItmX = T1View.ListItems.Item.ForeColor = &H8000000D
    ItmZ.SubItems(1) = "Command"     'file type
    ItmZ.SubItems(2) = ">>>>>> Reproducir Hora"        'file name
    'gets the file len and convert into time
    'put the rest of info
    ItmZ.SubItems(3) = "00:00:00"  'duracion del tema
    ItmZ.SubItems(4) = "00:00:00"  'poner aqui la hora de lanzamiento
    ItmZ.SubItems(5) = "-----"     'poner aqui el path & file del mixado
    ItmZ.SubItems(6) = "-----"     'poner aqui el type del mixado
    ItmZ.SubItems(7) = "-----"     'poner aqui el nombre del mixado interm.
    ItmZ.SubItems(8) = "00:00:00"     'poner aqui la duracion del mixado
    ItmZ.SubItems(9) = "00:00"  'poner aqui la hora de lanzam. del mixado
    'T1View.SelectedItem.ForeColor = &H8000000D

End Sub


Private Sub BlockMnu_config_Click()

'/// Configurar los bloques publicitarios.

FrmBlock.Show

End Sub

Private Sub BlockMnu_define_Click()

'/// display the Open dialog box
On Error Resume Next
TopMenu.BlockCmd.InitDir = App.path & AppBlockDir
TopMenu.BlockCmd.Filter = "Archivo de Bloque (*.blk)|*.blk|Archivo de Bloque"
TopMenu.BlockCmd.DialogTitle = "Bloques de publicidad - Abrir archivo de bloque."
TopMenu.BlockCmd.CancelError = True
TopMenu.BlockCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

BlkFn.Caption = TopMenu.BlockCmd.filename
Label10.Caption = StripFileFromDir(BlkFn.Caption)

End Sub

Private Sub BlockMnu_delete_Click()

'/// eliminar un bloque publicitario.

End Sub

Private Sub BlockMnu_insert_Click()

'/// Insertar un bloque publicitario.

End Sub

Private Sub ClockTimer_Timer()

End Sub

Private Sub CmdBlock_Click()

'display the block menu
PopupMenu BlockMnu

End Sub

Private Sub Form_Load()

'*** load some pictures *****
Me.Picture = LoadPicture(App.path & "\Imagenes\FND_COMPLETO.jpg")
Me.Image1 = LoadPicture(App.path & "\Imagenes\FND_PANEL_NEW.jpg")
Me.Image2 = LoadPicture(App.path & "\Imagenes\FND_PANEL_NEW.jpg")
Me.Image3 = LoadPicture(App.path & "\Imagenes\FND_PANEL_NEW.jpg")

'*** load commands pictures
    T1Next.Picture = LoadResPicture("R_NEXT", 0)
    T1Play.Picture = LoadResPicture("R_PLAY", 0)
    T1Stop.Picture = LoadResPicture("R_STOP", 0)
    T1Up.Picture = LoadResPicture("ICO_UP", 0)
    T1Down.Picture = LoadResPicture("ICO_DOWN", 0)
    T1Del.Picture = LoadResPicture("ICO_DELETE", 0)
    T1Prop.Picture = LoadResPicture("ICO_PROP", 0)
    T1Order.Picture = LoadResPicture("R_SYNC_ALL", 0)
    T1OrderA.Picture = LoadResPicture("R_SYNC_SELECTED", 0)
    T1New.Picture = LoadResPicture("ICO_NEW", 0)
    T1Open.Picture = LoadResPicture("ICO_OPEN", 0)
    T1Save.Picture = LoadResPicture("ICO_SAVE", 0)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'HideWindow "Tnd01"

End Sub

Private Sub Form_Terminate()

'HideWindow "Tnd01"

End Sub

Private Sub Form_Unload(Cancel As Integer)

'HideWindow "Tnd01"

End Sub

Private Sub Label10_Click()

MsgBox "Opción no implementada.", vbInformation
Exit Sub

Call BlockMnu_define_Click

End Sub

Private Sub LBlk_Click()

MsgBox "Opción no implementada.", vbInformation
Exit Sub

If LBlk.Caption = "/ Man" Then
    LBlk.Caption = "/ Auto"
    BlockMnu_insert.Enabled = False
    BlockMnu_delete.Enabled = False
    If Trim(Label10.Caption) = "no definido.blk" Then
        Call Label10_Click
    End If
Else
    LBlk.Caption = "/ Man"
    BlockMnu_insert.Enabled = True
    BlockMnu_delete.Enabled = True
End If

End Sub

Private Sub LFin_Click()

If LFin.Caption = "Man" Then
    LFin.Caption = "Auto"
Else
    LFin.Caption = "Man"
End If

End Sub

Private Sub SyncTimer_Timer()

Dim SStream As String
Dim SyncTime As String

Dim a1 As Long
Dim a2 As Long

SStream = Trim(SyncStream.Caption)
SyncTime = Trim(SyncLabel.Caption)  'syn time (in seconds)
a1 = CLng(SyncTime)

Select Case SStream
    Case "Stream01"
        'a2 = Stream01GetPosition(1) 'position in seconds
        a2 = GStreamGetPosition(1, 1)
        'MsgBox "Actual Pos: " & a2 & " - Synctime: " & a1
        If a2 >= a1 Then
            SyncStream.Caption = ""
            SyncLabel.Caption = ""
            SyncTimer.Interval = 0
            SyncTimer.Enabled = False
            Call Tanda01.T1Next_Click
        End If
        
    Case "Stream02"
        'a2 = Stream02GetPosition(1) 'position in seconds
        a2 = GStreamGetPosition(2, 1)
        'MsgBox "Actual Pos: " & a2 & " - Synctime: " & a1
        If a2 >= a1 Then
            SyncStream.Caption = ""
            SyncLabel.Caption = ""
            SyncTimer.Interval = 0
            SyncTimer.Enabled = False
            Call Tanda01.T1Next_Click
        End If
        
    Case Else
        SyncStream.Caption = ""
        SyncLabel.Caption = ""
        SyncTimer.Interval = 0
        SyncTimer.Enabled = False
End Select

End Sub

Private Sub T1Del_Click()

Dim TTime As String
Dim RTime As String
Dim Time1 As Long
Dim Time2 As Long
Dim Resultado As Long
Dim TMint As Double

On Error Resume Next
'primero extraemos el tiempo del item seleccionado
TTime = Trim(Ltime.Caption)
'RTime = Trim(T1View.SelectedItem.SubItems(3).Text)
'restamos el tiempo del item al del total
Time1 = ConvMinToSec(TTime)
Time2 = ConvMinToSec(RTime)
'tiempo de mixado intermedio
TMint = CDbl(Trim(Tanda01.Intr.Text))
Resultado = Time1 - Time2
Resultado = (Resultado + TMint) - 1
Result = ConvSecToMin(Resultado)
'restauramos el display
SetSumTime Result, 1
Ltime.Caption = Result

'eliminamos el item seleccionado
nIndex = T1View.SelectedItem.index
T1View.ListItems.Remove (nIndex)

'seleccionamos el item anterior al mismo
T1View.ListItems.Item(nIndex - 1).Selected = True
If T1View.ListItems.count < 1 Then
    T1View.ListItems.Clear
    Fn.Caption = ""
    Ltime.Caption = "00:00:00"
    Call RestoreDisplay(5)
    'deactivate al controls
    T1Next.Enabled = False
    T1Play.Enabled = False
    T1Stop.Enabled = False
    T1Up.Enabled = False
    T1Down.Enabled = False
    T1Del.Enabled = False
    T1Prop.Enabled = False
    T1Order.Enabled = False
    T1OrderA.Enabled = False
End If
Exit Sub

er:
End Sub

Private Sub T1Down_Click()

Dim DataA(0 To 9) As String
Dim DataKa As String

Dim DataB(0 To 9) As String
Dim DataKb As String

Dim ONum As Integer
Dim nCount As Integer
Dim NNum As Integer

On Error GoTo Continue
'chequeos necesarios
nCount = T1View.ListItems.count
ONum = T1View.SelectedItem.index
NNum = T1View.SelectedItem.index + 1

If NNum > nCount Or nCount = ONum Then Exit Sub

'extraemos los datos del item
DataA(0) = T1View.SelectedItem.Text    'file & path
DataA(1) = T1View.SelectedItem.SubItems(1) 'filetype
DataA(2) = T1View.SelectedItem.SubItems(2) 'filename
DataA(3) = T1View.SelectedItem.SubItems(3)
DataA(4) = T1View.SelectedItem.SubItems(4)
DataA(5) = T1View.SelectedItem.SubItems(5)
DataA(6) = T1View.SelectedItem.SubItems(6)
DataA(7) = T1View.SelectedItem.SubItems(7)
DataA(8) = T1View.SelectedItem.SubItems(8)
DataA(9) = T1View.SelectedItem.SubItems(9)
DataKa = T1View.SelectedItem.Key

'seleccionamos el siguiente item hacia abajo
nIndex = NNum
T1View.ListItems.Item(nIndex).Selected = True

'extraemos los datos del item
DataB(0) = T1View.SelectedItem.Text    'file & path
DataB(1) = T1View.SelectedItem.SubItems(1)   'filetype
DataB(2) = T1View.SelectedItem.SubItems(2)  'filename
DataB(3) = T1View.SelectedItem.SubItems(3)
DataB(4) = T1View.SelectedItem.SubItems(4)
DataB(5) = T1View.SelectedItem.SubItems(5)
DataB(6) = T1View.SelectedItem.SubItems(6)
DataB(7) = T1View.SelectedItem.SubItems(7)
DataB(8) = T1View.SelectedItem.SubItems(8)
DataB(9) = T1View.SelectedItem.SubItems(9)
DataKb = T1View.SelectedItem.Key

'ponemos los nuevos datos
T1View.ListItems.Remove (nIndex)
Set ItmX = T1View.ListItems.Add(nIndex, DataKb, DataA(0)) 'path & file
ItmX.SubItems(1) = DataA(1)
ItmX.SubItems(2) = DataA(2)
ItmX.SubItems(3) = DataA(3)
ItmX.SubItems(4) = DataA(4)
ItmX.SubItems(5) = DataA(5)
ItmX.SubItems(6) = DataA(6)
ItmX.SubItems(7) = DataA(7)
ItmX.SubItems(8) = DataA(8)
ItmX.SubItems(9) = DataA(9)

'seleccionamos el index anterior
nIndex = nIndex - 1
T1View.ListItems.Item(nIndex).Selected = True

'ponemos los nuevos datos
T1View.ListItems.Remove (nIndex)
Set ItmX = T1View.ListItems.Add(nIndex, DataKa, DataB(0)) 'path & file
ItmX.SubItems(1) = DataB(1)
ItmX.SubItems(2) = DataB(2)
ItmX.SubItems(3) = DataB(3)
ItmX.SubItems(4) = DataB(4)
ItmX.SubItems(5) = DataB(5)
ItmX.SubItems(6) = DataB(6)
ItmX.SubItems(7) = DataB(7)
ItmX.SubItems(8) = DataB(8)
ItmX.SubItems(9) = DataB(9)

'una vez finalizado. seleccionamos el item
nIndex = nIndex + 1
T1View.ListItems.Item(nIndex).Selected = True
Exit Sub

Continue:
    'nothing to do....
End Sub

Private Sub T1New_Click()

T1View.ListItems.Clear
Fn.Caption = ""
Ltime.Caption = "00:00:00"
Call RestoreDisplay(5)

'deactivate al controls
T1Next.Enabled = False
T1Play.Enabled = False
T1Stop.Enabled = False
T1Up.Enabled = False
T1Down.Enabled = False
T1Del.Enabled = False
T1Prop.Enabled = False
T1Order.Enabled = False
T1OrderA.Enabled = False

End Sub

Sub T1Next_Click()

'gets the count of item in list
If T1View.ListItems.count < 1 Then Exit Sub

'desabilitaciones
T1View.Enabled = False
T1Next.Enabled = False
T1Play.Enabled = False
T1Stop.Enabled = True
T1Up.Enabled = False
T1Down.Enabled = False
T1Del.Enabled = False
T1Prop.Enabled = False
T1Order.Enabled = False
T1OrderA.Enabled = False
T1New.Enabled = False
T1Open.Enabled = False
T1Save.Enabled = False

'deactivate all controls
RestoreAllActiveColor 1

'starts the fadeout
If Stream01IsPlaying = True Or Music01IsPlaying = True Then 'stream01 fade out
    If LFin.Caption = "Auto" Then
        If Est12Control.Origen1.Caption = "E1" Then
            Est01.TmoutAuto.Enabled = True
            Est01.TmoutAuto.Interval = 50
        Else
            TmOut1.Enabled = True
            TmOut1.Interval = 50
        End If
    End If
End If
If Stream02IsPlaying = True Or Music02IsPlaying = True Then 'stream02 fade out
    If LFin.Caption = "Auto" Then
        If Est12Control.Origen2.Caption = "E2" Then
            Est02.TmoutAuto.Enabled = True
            Est02.TmoutAuto.Interval = 50
        Else
            TmOut2.Enabled = True
            TmOut2.Interval = 50
        End If
    End If
End If

'//// gets the file to play
Dim nIndex As Integer

On Error GoTo Continue
nIndex = CInt(LKey.Caption) 'get the index of the file to play
T1View.ListItems.Item(nIndex).Selected = True   'select the item

'gets the file info
FileN = T1View.SelectedItem.Text    'file
FileTP = T1View.SelectedItem.SubItems(1)   'filetype
SSTitle = T1View.SelectedItem.SubItems(2)  'file title

'//// checks for file exists
If FileExist(FileN) = False Then
    nIndex = nIndex + 1
    If nIndex > T1View.ListItems.count Then
        nIndex = 1
    End If
    Tanda01.T1View.ListItems.Item(nIndex).Selected = True
    'gets the file info of new file
    FileN = T1View.SelectedItem.Text    'file
    FileTP = T1View.SelectedItem.SubItems(1)   'filetype
    SSTitle = T1View.SelectedItem.SubItems(2) 'file title
End If

'****************** FILE CUE & FX PRESETS load...
Dim filename As String
Dim NameFile As String

filename = Trim(FileN)    'extraemos el path y el archivo de audio
NameFile = StripFileFromExt(filename)
filename = Trim(NameFile) & AppCUEFileExt

'**************** COMENZAMOS LA REPRODUCCION DEL AUDIO

'Chequeamos los dispositivos en uso y decidimos cual usar (dev1 or dev2)
If Stream01IsPlaying = True Or Music01IsPlaying = True Then
    If Stream02IsPlaying = True Or Music02IsPlaying = True Then
        'activate the fade in
        TmIn1.Enabled = True
        TmIn1.Interval = 50
        'close stream1 and play the file
        GStreamStop 1
        'load and play the selected file
        Call Tanda01Play(FileN, SSTitle, FileTP, "Yes")  '//// USE DEV 1 ////
        'load CUE info & FX info
        OpenCUEFile 1, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    Else
        'activate the fade in
        TmIn2.Enabled = True
        TmIn2.Interval = 50
        'close stream2 and play the file
        GStreamStop 2
        'load and play the selected file
        Call Tanda02Play(FileN, SSTitle, FileTP, "Yes")  '//// USE DEV 2 ////
        'load CUE info & FX info
        OpenCUEFile 2, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    End If
Else
    If Stream02IsPlaying = True Or Music02IsPlaying = True Then
        'activate the fade in
        TmIn1.Enabled = True
        TmIn1.Interval = 50
        'close stream1 and play the file
        GStreamStop 1
        'load and play the selected file
        Call Tanda01Play(FileN, SSTitle, FileTP, "Yes")  '//// USE DEV 1 ////
        'load CUE info & FX info
        OpenCUEFile 1, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    Else
        'activate the fade in
        TmIn1.Enabled = True
        TmIn1.Interval = 50
        'close stream1 and play the file
        GStreamStop 1
        'load and play the selected file
        Call Tanda01Play(FileN, SSTitle, FileTP, "Yes")  '//// USE DEV 1 ////
        'load CUE info & FX info
        OpenCUEFile 1, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    End If
End If

'sets the next item to play
LKey.Caption = nIndex + 1

Tanda01.Caption = "TANDA - Reproduciendo"
'reseteamos la hora de lanzamiento a la hora actual
'del sistema y actualizamos los temas subsiguientes
'a su correspondiente hora de lanzamiento.
Call OrderTndTime("ResetSelected")

If FTime.Caption = "---" Then
    Call SetStartTime
End If
Exit Sub

Continue:
'habilitaciones
'desabilitaciones
T1View.Enabled = True
T1Next.Enabled = True
T1Play.Enabled = True
T1Stop.Enabled = True
T1Up.Enabled = True
T1Down.Enabled = True
T1Del.Enabled = True
T1Prop.Enabled = True
T1Order.Enabled = True
T1OrderA.Enabled = True
T1New.Enabled = True
T1Open.Enabled = True
T1Save.Enabled = True

'disable the sync timer
SyncTimer.Interval = 0
SyncTimer.Enabled = False

End Sub

Private Sub T1Open_Click()

On Error Resume Next
TopMenu.TandaCmd.InitDir = App.path & AppTandaDir & "\"
TopMenu.TandaCmd.Filter = "Archivo de Tanda (*.tnd)|*.tnd|Archivos de Tanda"
TopMenu.TandaCmd.DialogTitle = "TANDAS - Abrir archivo"
TopMenu.TandaCmd.CancelError = True
TopMenu.TandaCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

    'restauramos los valores a 0
    T1View.ListItems.Clear
    Fn.Caption = ""
    Ltime.Caption = "00:00:00"
    Call RestoreDisplay(5)

ConvertTx = TopMenu.TandaCmd.filename

Result = OpenTandaFile(ConvertTx)
If Result = "NotOK" Then
    'MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
    Exit Sub
End If

Fn.Caption = ConvertTx

End Sub

Private Sub T1Order_Click()

Call OrderTndTime("ResetAll")

End Sub

Sub T1OrderA_Click()

Call OrderTndTime("ResetSelected")

End Sub

Sub T1Play_Click()

'get the list item count
If T1View.ListItems.count < 1 Then Exit Sub

'deactivate all controls
RestoreAllActiveColor 1

'starts the fadeout an other device in use (if there is another in use)
If Stream01IsPlaying = True Or Music01IsPlaying = True Then
    If LFin.Caption = "Auto" Then
        If Est12Control.Origen1.Caption = "E1" Then
            Est01.TmoutAuto.Enabled = True
            Est01.TmoutAuto.Interval = 50
        Else
            TmOut1.Enabled = True
            TmOut1.Interval = 50
        End If
    End If
End If

If Stream02IsPlaying = True Or Music02IsPlaying = True Then
    If LFin.Caption = "Auto" Then
        If Est12Control.Origen2.Caption = "E2" Then
            Est02.TmoutAuto.Enabled = True
            Est02.TmoutAuto.Interval = 50
        Else
            TmOut2.Enabled = True
            TmOut2.Interval = 50
        End If
    End If
End If

Dim nIndex As Integer

On Error GoTo err
nIndex = CInt(LKey.Caption) 'get the index of the file to play
T1View.ListItems.Item(nIndex).Selected = True   'select the item

'//// gets the file info
FileN = T1View.SelectedItem.Text    'file
FileTP = T1View.SelectedItem.SubItems(1)   'filetype
SSTitle = T1View.SelectedItem.SubItems(2)  'file title

'//// checks for file exists
If FileExist(FileN) = False Then
    nIndex = nIndex + 1
    If nIndex > T1View.ListItems.count Then
        nIndex = 1
    End If
    Tanda01.T1View.ListItems.Item(nIndex).Selected = True
    'gets the file info of new file
    FileN = T1View.SelectedItem.Text    'file
    FileTP = T1View.SelectedItem.SubItems(1)  'filetype
    SSTitle = T1View.SelectedItem.SubItems(2) 'file title
End If

'****************** FILE CUE & FX PRESETS load...
Dim filename As String
Dim NameFile As String

filename = Trim(FileN)    'extraemos el path y el archivo de audio
NameFile = StripFileFromExt(filename)
filename = Trim(NameFile) & AppCUEFileExt

'**************** COMENZAMOS LA REPRODUCCION DEL ARCHIVO DE AUDIO
Tanda01.Caption = "TANDA - Reproduciendo"

'Chequeamos los dispositivos en uso y decidimos cual usar (dev1 or dev2)
If Stream01IsPlaying = True Or Music01IsPlaying = True Then
    If Stream02IsPlaying = True Or Music02IsPlaying = True Then
        'activate the fade in
        TmIn1.Enabled = True
        TmIn1.Interval = 50
        'close stream1 and play the file
        GStreamStop 1
        'load and play the selected file
        Call Tanda01Play(FileN, SSTitle, FileTP, "No")  '//// USE DEV 1 ////
        'load CUE info & FX info
        OpenCUEFile 1, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    Else
        'activate the fade in
        TmIn2.Enabled = True
        TmIn2.Interval = 50
        'close stream2 and play the file
        GStreamStop 2
        'load and play the selected file
        Call Tanda02Play(FileN, SSTitle, FileTP, "No")  '//// USE DEV 2 ////
        'load CUE info & FX info
        OpenCUEFile 2, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    End If
Else
    If Stream02IsPlaying = True Or Music02IsPlaying = True Then
        'activate the fade in
        TmIn1.Enabled = True
        TmIn1.Interval = 50
        'close stream1 and play the file
        GStreamStop 1
        'load and play the selected file
        Call Tanda01Play(FileN, SSTitle, FileTP, "No")  '//// USE DEV 1 ////
        'load CUE info & FX info
        OpenCUEFile 1, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    Else
        'activate the fade in
        TmIn1.Enabled = True
        TmIn1.Interval = 50
        'close stream1 and play the file
        GStreamStop 1
        'load and play the selected file
        Call Tanda01Play(FileN, SSTitle, FileTP, "No")  '//// USE DEV 1 ////
        'load CUE info & FX info
        OpenCUEFile 1, filename
        'activate the clock timer
        TopMenu.ProcTimer.Enabled = True
        TopMenu.ProcTimer.Interval = 1
    End If
End If

err:
'//// disable the sync timer
SyncTimer.Interval = 0
SyncTimer.Enabled = False

End Sub

Private Sub T1Prop_Click()

Dim filename As String

filename = Trim(Tanda01.T1View.SelectedItem.Text)    'file & path

'chequeamos por la validez de los datos
If FileExist(filename) = False Then
    MsgBox "El archivo de audio seleccionado no existe o fué eliminado.", vbCritical
    MsgBox "Remueva el item de la lista para evitar futuros inconvenientes.", vbInformation
    Exit Sub
Else
    'display de audio prop window
    AudioProp.Show
End If

End Sub

Sub T1Save_Click()

ConvertTxT = Trim(Fn.Caption)

On Error Resume Next
If ConvertTxT = "" Or ConvertTxT = " " Then
    TopMenu.TandaCmd.InitDir = App.path & AppTandaDir & "\"
    TopMenu.TandaCmd.Filter = "Archivo de Tandas (*.tnd)|*.tnd|Archivos de Tanda"
    TopMenu.TandaCmd.DialogTitle = "TANDAS - Guardar archivo"
    TopMenu.TandaCmd.FilterIndex = 1
    TopMenu.TandaCmd.CancelError = True
    TopMenu.TandaCmd.ShowSave

    If err.Number = 32755 Then Exit Sub
    
    ConvertTx = TopMenu.TandaCmd.filename

    Fn.Caption = ConvertTx
    Result = SaveTandaFile(ConvertTx)
    If Result = "NotOK" Then
        MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
        Exit Sub
    End If
Else
    ConvertTx = Trim(Fn.Caption)
    Kill ConvertTx
    Result = SaveTandaFile(ConvertTx)
    If Result = "NotOK" Then
        'MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
        Exit Sub
    End If
End If

End Sub

Private Sub T1Stop_Click()

If T1View.ListItems.count < 1 Then Exit Sub

'habilitaciones
T1View.Enabled = True
T1Next.Enabled = True
T1Play.Enabled = True
T1Stop.Enabled = True
T1Up.Enabled = True
T1Down.Enabled = True
T1Del.Enabled = True
T1Prop.Enabled = True
T1Order.Enabled = True
T1OrderA.Enabled = True
T1New.Enabled = True
T1Open.Enabled = True
T1Save.Enabled = True

If Est12Control.Origen1.Caption = "T1" Then
    If Stream01IsPlaying = True Then
        TmOut1.Enabled = True
        TmOut1.Interval = 50
        Tanda01.Caption = "TANDA - Detenido"
    Else
        If Music01IsPlaying = True Then
            'Music01Stop
            'Music01Restart
            'Tanda01.Caption = "TANDA - Detenido"
        Else
            'nothing to do
        End If
    End If
Else
    'nothing to do
End If

If Est12Control.Origen2.Caption = "T2" Then
    If Stream02IsPlaying = True Then
        TmOut2.Enabled = True
        TmOut2.Interval = 50
        Tanda01.Caption = "TANDA - Detenido"
    Else
        If Music02IsPlaying = True Then
            'Music02Stop
            'Music02Restart
            'Tanda01.Caption = "TANDA - Detenido"
        Else
            'nothing to do
        End If
    End If
Else
    'nothing to do
End If

'disable the sync timer
SyncTimer.Interval = 0
SyncTimer.Enabled = False

'set the time to nothing
FTime.Caption = "---"

End Sub

Private Sub T1Up_Click()

Dim DataA(0 To 9) As String
Dim DataKa As String

Dim DataB(0 To 9) As String
Dim DataKb As String

Dim ONum As Integer
Dim nCount As Integer
Dim NNum As Integer

On Error GoTo Continue
'chequeos necesarios
nCount = T1View.ListItems.count
ONum = T1View.SelectedItem.index
NNum = T1View.SelectedItem.index - 1

If NNum < 0 Or ONum = 1 Then Exit Sub

'extraemos los datos del item seleccionado
DataA(0) = T1View.SelectedItem.Text    'file & path
DataA(1) = T1View.SelectedItem.SubItems(1)   'filetype
DataA(2) = T1View.SelectedItem.SubItems(2)  'filename
DataA(3) = T1View.SelectedItem.SubItems(3)
DataA(4) = T1View.SelectedItem.SubItems(4)
DataA(5) = T1View.SelectedItem.SubItems(5)
DataA(6) = T1View.SelectedItem.SubItems(6)
DataA(7) = T1View.SelectedItem.SubItems(7)
DataA(8) = T1View.SelectedItem.SubItems(8)
DataA(9) = T1View.SelectedItem.SubItems(9)
DataKa = T1View.SelectedItem.Key

'seleccionamos el siguiente item hacia abajo
nIndex = NNum
T1View.ListItems.Item(nIndex).Selected = True

'extraemos los datos del item
DataB(0) = T1View.SelectedItem.Text    'file & path
DataB(1) = T1View.SelectedItem.SubItems(1)   'filetype
DataB(2) = T1View.SelectedItem.SubItems(2)  'filename
DataB(3) = T1View.SelectedItem.SubItems(3)
DataB(4) = T1View.SelectedItem.SubItems(4)
DataB(5) = T1View.SelectedItem.SubItems(5)
DataB(6) = T1View.SelectedItem.SubItems(6)
DataB(7) = T1View.SelectedItem.SubItems(7)
DataB(8) = T1View.SelectedItem.SubItems(8)
DataB(9) = T1View.SelectedItem.SubItems(9)
DataKb = T1View.SelectedItem.Key

'ponemos los nuevos datos
T1View.ListItems.Remove (nIndex)
Set ItmX = T1View.ListItems.Add(nIndex, DataKb, DataA(0)) 'path & file
ItmX.SubItems(1) = DataA(1)
ItmX.SubItems(2) = DataA(2)
ItmX.SubItems(3) = DataA(3)
ItmX.SubItems(4) = DataA(4)
ItmX.SubItems(5) = DataA(5)
ItmX.SubItems(6) = DataA(6)
ItmX.SubItems(7) = DataA(7)
ItmX.SubItems(8) = DataA(8)
ItmX.SubItems(9) = DataA(9)

'seleccionamos el index anterior
nIndex = nIndex + 1
T1View.ListItems.Item(nIndex).Selected = True

'ponemos los nuevos datos
T1View.ListItems.Remove (nIndex)
Set ItmX = T1View.ListItems.Add(nIndex, DataKa, DataB(0)) 'path & file
ItmX.SubItems(1) = DataB(1)
ItmX.SubItems(2) = DataB(2)
ItmX.SubItems(3) = DataB(3)
ItmX.SubItems(4) = DataB(4)
ItmX.SubItems(5) = DataB(5)
ItmX.SubItems(6) = DataB(6)
ItmX.SubItems(7) = DataB(7)
ItmX.SubItems(8) = DataB(8)
ItmX.SubItems(9) = DataB(9)

'una vez finalizado. seleccionamos el item
nIndex = nIndex - 1
T1View.ListItems.Item(nIndex).Selected = True
Exit Sub

Continue:
    'nothing to do....
End Sub

Private Sub T1View_Click()

On Error GoTo er
LKey.Caption = T1View.SelectedItem.index
T1Next.Enabled = True
T1Play.Enabled = True
T1Stop.Enabled = True
T1Up.Enabled = True
T1Down.Enabled = True
T1Del.Enabled = True
T1Prop.Enabled = True
T1Order.Enabled = True
T1OrderA.Enabled = True
Exit Sub

er:
T1Next.Enabled = False
T1Play.Enabled = False
T1Stop.Enabled = False
T1Up.Enabled = False
T1Down.Enabled = False
T1Del.Enabled = False
T1Prop.Enabled = False
T1Order.Enabled = False
T1OrderA.Enabled = False
End Sub

Private Sub T1View_DblClick()

Call T1Play_Click

End Sub

Private Sub T1View_DragDrop(Source As Control, X As Single, Y As Single)

DeployTandaFile 'drag & drop the selected file in xplorer

End Sub

Private Sub T1View_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

Select Case State
    Case 0  'drag not finished
        XPlorer.File1.DragIcon = XPlorer.ExCombo.DragIcon
        'E11(Index).BackColor = &H80FF80    'verde (modificacion)
    Case 1  'finished drag
        XPlorer.File1.DragIcon = XPlorer.tvwDirTree.DragIcon
        'E11(Index).BackColor = &H8000000F  'gris (normal)
End Select

End Sub

Private Sub T1View_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'button 1=left button
'button 2=right button
'button 4=mid button

If Button = 2 Then
    'xxxxxx
Else
    'xxxxxx
End If

End Sub

Private Sub T1Vol_Change()

If Est12Control.StopLabel1.Caption = "Stream" Then
    'change the stream volume
    'Stream01SetVolume (T1Vol.Value)
    GStreamSetVolume 1, T1Vol.Value
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        'change the music volume
        Music01SetVolume (T1Vol.Value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub T1Vol_Scroll()

If Est12Control.StopLabel1.Caption = "Stream" Then
    'change the stream volume
    'Stream01SetVolume (T1Vol.Value)
    GStreamSetVolume 1, T1Vol.Value
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        'change the music volume
        Music01SetVolume (T1Vol.Value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub T2Vol_Change()

If Est12Control.StopLabel2.Caption = "Stream" Then
    'change the stream volume
    'Stream02SetVolume (T2Vol.Value)
    GStreamSetVolume 2, T2Vol.Value
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        'change the music volume
        Music02SetVolume (T2Vol.Value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub T2Vol_Scroll()

If Est12Control.StopLabel2.Caption = "Stream" Then
    'change the stream volume
    'Stream02SetVolume (T2Vol.Value)
    GStreamSetVolume 2, T2Vol.Value
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        'change the music volume
        Music02SetVolume (T2Vol.Value)
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub TmIn1_Timer()

If T1Vol.Value = 100 Then
    TmIn1.Interval = 0
    TmIn1.Enabled = False
    Exit Sub
Else
    T1Vol.Value = T1Vol.Value + 5
End If

End Sub

Private Sub TmIn2_Timer()

If T2Vol.Value = 100 Then
    TmIn2.Interval = 0
    TmIn2.Enabled = False
    Exit Sub
Else
    T2Vol.Value = T2Vol.Value + 5
End If

End Sub

Private Sub TmOut1_Timer()

If T1Vol.Value = 0 Then
    If Est12Control.StopLabel1.Caption = "Stream" And Est12Control.Origen1.Caption = "T1" Then
        'Stream01Restart    'stream restart
        GStreamRestart 1
        GStreamStop 1
    Else
        If Est12Control.StopLabel1.Caption = "Music" And Est12Control.Origen1.Caption = "T1" Then
            Music01Restart     'music restart
            Music01Stop         'music stop
        Else
            Exit Sub
        End If
    End If
    TmOut1.Interval = 0
    TmOut1.Enabled = False
Else
    T1Vol.Value = T1Vol.Value - 5
End If

End Sub

Private Sub TmOut2_Timer()

If T2Vol.Value = 0 Then
    If Est12Control.StopLabel1.Caption = "Stream" And Est12Control.Origen2.Caption = "T2" Then
        'Stream02Restart    'stream restart
        GStreamRestart 2
        GStreamStop 2
    Else
        If Est12Control.StopLabel1.Caption = "Music" And Est12Control.Origen2.Caption = "T2" Then
            Music02Restart     'music restart
            Music02Stop         'music stop
        Else
            Exit Sub
        End If
    End If
    TmOut2.Interval = 0
    TmOut2.Enabled = False
Else
    T2Vol.Value = T2Vol.Value - 5
End If

End Sub

Private Sub VScroll1_Change()

Intr.Text = VScroll1.Value

End Sub
