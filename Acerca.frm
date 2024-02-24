VERSION 5.00
Begin VB.Form Acerca 
   BorderStyle     =   0  'None
   Caption         =   "Acerca RadioMaker"
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   6090
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAcp 
      Caption         =   "A"
      Height          =   405
      Left            =   4620
      TabIndex        =   5
      Top             =   7260
      Width           =   1275
   End
   Begin VB.PictureBox PicAbt 
      AutoSize        =   -1  'True
      Height          =   6810
      Left            =   0
      ScaleHeight     =   6750
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   375
      Width           =   6060
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agradecimientos especiales a:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   5190
         Width           =   2220
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"Acerca.frx":0000
         Height          =   420
         Left            =   120
         TabIndex        =   2
         Top             =   5445
         Width           =   5820
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Al Sr. Cándido Blas Iradis (FM Milenio) Prov. de Salta, Rep. Argentina."
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   5850
         Width           =   4965
      End
   End
   Begin RM100.TitelBar TitelBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6090
      _ExtentX        =   10742
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
      Caption         =   " ACERCA RadioMaker"
      CaptionPosX     =   1
      BorderNormal    =   2
      BorderColorDarkLight=   12632256
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      Height          =   240
      Left            =   150
      TabIndex        =   4
      Top             =   7260
      Width           =   4470
   End
End
Attribute VB_Name = "Acerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAcp_Click()

Unload Me

End Sub

Private Sub Form_Load()

PicAbt.Picture = LoadPicture(App.path & "\Imagenes\RM_ACERCA.bmp")
CmdAcp.Caption = LoadResString(2000)

Label5.Caption = "Versión: " & App.Major & "." & App.Minor & " - Revisión: " & App.Revision

'KeepOnTop Acerca

End Sub

