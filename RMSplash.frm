VERSION 5.00
Begin VB.Form RMSplash 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6015
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6285
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Progress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   5985
      ScaleHeight     =   6000
      ScaleWidth      =   300
      TabIndex        =   3
      Top             =   0
      Width           =   300
   End
   Begin VB.PictureBox Status 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   5985
      ScaleHeight     =   6000
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   0
      Width           =   300
   End
   Begin VB.Timer Timer1 
      Left            =   2310
      Top             =   6450
   End
   Begin VB.PictureBox PicSplash 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   0
      ScaleHeight     =   6000
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin VB.Label LblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Versión X.Xx"
         Height          =   255
         Left            =   1395
         TabIndex        =   1
         Top             =   5010
         Width           =   1275
      End
   End
End
Attribute VB_Name = "RMSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (c) 2002 - 2022 ONLY development inc.

Option Explicit

Private Sub AppRunning()

Dim sMsg As String
        
If App.PrevInstance Then
    sMsg = App.EXEName & " esta actualmente en uso!. "
    MsgBox sMsg, vbCritical, "Alerta!!!"
    End
End If

End Sub

Private Sub Form_Load()

'---some strings to load
LblVersion.Caption = "Versión: " & App.Major & "." & App.Minor & "." & App.Revision

'abrimos el archivo de configuracion
On Error Resume Next
Open App.path & AppUpdateDir & AppUpdateVerFile For Output As #16
Write #16, App.Major & "." & App.Minor & "." & App.Revision, App.path
Close #16

'---some bitmaps to load
PicSplash.Picture = LoadResPicture("RM_INTRO", 0)
Status.Picture = LoadResPicture("PROGRESS_01", 0)
Progress.Picture = LoadResPicture("PROGRESS_02", 0)

'KeepOnTop RMSplash
Progress.Height = 6000

Timer1.Enabled = True
Timer1.Interval = 50

End Sub

Private Sub Timer1_Timer()

Dim PMax

If Progress.Height = 15 Then
    Progress.Height = 5
    Call Main
Else
    If Progress.Height = 3000 Then
        Call AppRunning
        Progress.Height = Progress.Height - 100
    Else
        Progress.Height = Progress.Height - 100
    End If
End If

End Sub
