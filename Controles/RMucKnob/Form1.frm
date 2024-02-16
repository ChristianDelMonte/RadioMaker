VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form2"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9855
   LinkTopic       =   "Form2"
   ScaleHeight     =   6105
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   2175
      Left            =   600
      ScaleHeight     =   2115
      ScaleWidth      =   5835
      TabIndex        =   10
      Top             =   3360
      Width           =   5895
      Begin Proyecto1.ucKnob ucKnob2 
         Height          =   975
         Index           =   2
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   1095
         _extentx        =   1931
         _extenty        =   1720
         max             =   50
         forecolor       =   3769344
         backcolor       =   27392
         lightintencity  =   70
         tickssize       =   1
         tickslongfrequency=   5
         tickforecolor   =   15724527
         tickbackcolor   =   35840
      End
      Begin Proyecto1.ucKnob ucKnob2 
         Height          =   975
         Index           =   1
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   1095
         _extentx        =   1931
         _extenty        =   1720
         max             =   50
         forecolor       =   3750201
         backcolor       =   4342338
         lightintencity  =   100
         tickssize       =   1
         tickslongfrequency=   5
         roundstyle      =   -1  'True
         tickforecolor   =   9764657
         tickbackcolor   =   6513507
      End
      Begin Proyecto1.ucKnob ucKnob2 
         Height          =   975
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   1095
         _extentx        =   1931
         _extenty        =   1720
         max             =   50
         forecolor       =   10263708
         backcolor       =   7566195
         lightintencity  =   100
         tickssize       =   1
         tickslongfrequency=   5
         tickforecolor   =   16734810
         tickbackcolor   =   27614
      End
   End
   Begin Proyecto1.ucKnob ucKnob1 
      Height          =   2055
      Index           =   3
      Left            =   4800
      TabIndex        =   9
      Top             =   480
      Width           =   1935
      _extentx        =   3413
      _extenty        =   3625
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MAX"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   195
      Index           =   1
      Left            =   3720
      TabIndex        =   8
      Top             =   2880
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MIN"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   195
      Index           =   0
      Left            =   3000
      TabIndex        =   7
      Top             =   2880
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1920
      TabIndex        =   6
      Top             =   2880
      Width           =   105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   1800
      Width           =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF3131&
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   2880
      Width           =   90
   End
   Begin Proyecto1.ucKnob ucKnob1 
      Height          =   1215
      Index           =   4
      Left            =   960
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
      _extentx        =   1931
      _extenty        =   2143
      max             =   2
      steps           =   2
      tickssize       =   0
      tickspenwidth   =   2
      tickslongfrequency=   1
   End
   Begin Proyecto1.ucKnob ucKnob1 
      Height          =   1035
      Index           =   0
      Left            =   3000
      TabIndex        =   2
      Top             =   1920
      Width           =   1275
      _extentx        =   2249
      _extenty        =   1826
      backcolor       =   15724527
      tickssize       =   1
      tickslongfrequency=   5
      tickssmallhiden =   -1  'True
      ticksstylecircle=   -1  'True
      tickbackcolor   =   -2147483638
   End
   Begin Proyecto1.ucKnob ucKnob1 
      Height          =   1095
      Index           =   2
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   1095
      _extentx        =   1931
      _extenty        =   1931
      steps           =   10
      backcolor       =   5963685
      tickssize       =   1
      tickssmallhiden =   -1  'True
      ticksstylecircle=   -1  'True
      roundstyle      =   -1  'True
   End
   Begin Proyecto1.ucKnob ucKnob1 
      Height          =   795
      Index           =   1
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   795
      _extentx        =   1402
      _extenty        =   1402
      backcolor       =   5374133
      tickssize       =   1
      tickslongfrequency=   5
      tickssmallhiden =   -1  'True
      tickforecolor   =   13500655
      tickbackcolor   =   -2147483638
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents cSubClass As clsSubClass
Attribute cSubClass.VB_VarHelpID = -1
Private Const WM_MOUSEWHEEL As Long = &H20A
Private mControlKnob As ucKnob

Private Sub cSubClass_WindowProc(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
    Dim lSteps As Long
    If uMsg = WM_MOUSEWHEEL Then
        If Not mControlKnob Is Nothing Then
    
            If mControlKnob.Steps = 0 Then
                lSteps = 1
            Else
                lSteps = (mControlKnob.Max - mControlKnob.Min) / mControlKnob.Steps
            End If

            If wParam < 0 Then
                If mControlKnob.Value < mControlKnob.Max Then
                    mControlKnob.Value = mControlKnob.Value + lSteps
                End If
            Else
                If mControlKnob.Value > mControlKnob.Min Then
                    mControlKnob.Value = mControlKnob.Value - lSteps
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Set cSubClass = New clsSubClass
    cSubClass.ssc_Subclass Me.hwnd
    cSubClass.ssc_AddMsg Me.hwnd, WM_MOUSEWHEEL, MSG_AFTER
    cSubClass.ssc_Subclass Picture1.hwnd
    cSubClass.ssc_AddMsg Picture1.hwnd, WM_MOUSEWHEEL, MSG_AFTER
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set cSubClass = Nothing
End Sub

Private Sub ucKnob1_Change(Index As Integer)

    If Index = 4 Then
        Dim i As Long
        For i = 0 To 2
            Label1(i).ForeColor = &H80000012
        Next
        Label1(ucKnob1(Index).Value).ForeColor = vbBlue
    End If
End Sub

Private Sub ucKnob1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set mControlKnob = ucKnob1(Index)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set mControlKnob = Nothing
End Sub

Private Sub ucKnob2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set mControlKnob = ucKnob2(Index)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set mControlKnob = Nothing
End Sub

