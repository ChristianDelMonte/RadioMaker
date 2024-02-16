VERSION 5.00
Begin VB.Form FrmPrefs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preferencias"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8265
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCc 
      Caption         =   "Cc"
      Height          =   375
      Left            =   7095
      TabIndex        =   1
      Top             =   3060
      Width           =   1065
   End
   Begin VB.CommandButton CmdAc 
      Caption         =   "Ac"
      Height          =   375
      Left            =   5940
      TabIndex        =   0
      Top             =   3060
      Width           =   1065
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   8130
      X2              =   105
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   8160
      X2              =   90
      Y1              =   2910
      Y2              =   2910
   End
End
Attribute VB_Name = "FrmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
