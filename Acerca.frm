VERSION 5.00
Begin VB.Form Acerca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca RadioMaker"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5985
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAcp 
      Caption         =   "A"
      Height          =   405
      Left            =   4590
      TabIndex        =   5
      Top             =   6870
      Width           =   1275
   End
   Begin VB.PictureBox PicAbt 
      AutoSize        =   -1  'True
      Height          =   6810
      Left            =   -30
      ScaleHeight     =   6750
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   -15
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   6870
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

PicAbt.Picture = LoadResPicture("RM_ABOUT", 0)
CmdAcp.Caption = LoadResString(2000)

Label5.Caption = "Versión: " & App.Major & "." & App.Minor & " - Revisión: " & App.Revision

'KeepOnTop Acerca

End Sub

Private Sub PicAbt_Click()

End Sub
