VERSION 5.00
Begin VB.Form Est12Control 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Estacion 01 y 02 Digital Output Control"
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Label LblFX 
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Origen2 
      Caption         =   "--"
      Height          =   240
      Left            =   4905
      TabIndex        =   4
      Top             =   2025
      Width           =   870
   End
   Begin VB.Label Origen1 
      Caption         =   "--"
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   2025
      Width           =   870
   End
   Begin VB.Label StopLabel2 
      Caption         =   "Stream"
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label StopLabel1 
      Caption         =   "Stream"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estacion 01 y 02 Digital Output Control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Est12Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
