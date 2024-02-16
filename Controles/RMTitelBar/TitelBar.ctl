VERSION 5.00
Begin VB.UserControl TitelBar 
   Alignable       =   -1  'True
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   575
   ToolboxBitmap   =   "TitelBar.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   1800
      Top             =   180
   End
   Begin VB.PictureBox Image1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image TBARMax 
      Height          =   240
      Left            =   4005
      Top             =   90
      Width           =   240
   End
   Begin VB.Image imgMax 
      Height          =   210
      Index           =   23
      Left            =   6930
      Picture         =   "TitelBar.ctx":0312
      Top             =   1845
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgMax 
      Height          =   210
      Index           =   22
      Left            =   6930
      Picture         =   "TitelBar.ctx":06DC
      Top             =   1305
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgMax 
      Height          =   210
      Index           =   21
      Left            =   6930
      Picture         =   "TitelBar.ctx":0AA6
      Top             =   855
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgMax 
      Height          =   240
      Index           =   3
      Left            =   1845
      Picture         =   "TitelBar.ctx":0E70
      Top             =   1305
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgMax 
      Height          =   240
      Index           =   2
      Left            =   1845
      Picture         =   "TitelBar.ctx":11FA
      Top             =   1035
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgMax 
      Height          =   240
      Index           =   1
      Left            =   1845
      Picture         =   "TitelBar.ctx":1584
      Top             =   810
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgMax 
      Height          =   270
      Index           =   13
      Left            =   4500
      Picture         =   "TitelBar.ctx":190E
      Top             =   1845
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgMax 
      Height          =   270
      Index           =   12
      Left            =   4500
      Picture         =   "TitelBar.ctx":1F40
      Top             =   1305
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgMax 
      Height          =   270
      Index           =   11
      Left            =   4500
      Picture         =   "TitelBar.ctx":2572
      Top             =   855
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgClose 
      Height          =   210
      Index           =   21
      Left            =   6210
      Picture         =   "TitelBar.ctx":2BA4
      Top             =   855
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgClose 
      Height          =   210
      Index           =   22
      Left            =   6210
      Picture         =   "TitelBar.ctx":2F6E
      Top             =   1305
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgMin 
      Height          =   210
      Index           =   21
      Left            =   6570
      Picture         =   "TitelBar.ctx":3338
      Top             =   855
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgMin 
      Height          =   210
      Index           =   22
      Left            =   6570
      Picture         =   "TitelBar.ctx":3702
      Top             =   1305
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgClose 
      Height          =   210
      Index           =   23
      Left            =   6210
      Picture         =   "TitelBar.ctx":3ACC
      Top             =   1845
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgMin 
      Height          =   210
      Index           =   23
      Left            =   6570
      Picture         =   "TitelBar.ctx":3E96
      Top             =   1845
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgMin 
      Height          =   270
      Index           =   13
      Left            =   4005
      Picture         =   "TitelBar.ctx":4260
      Top             =   1845
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgClose 
      Height          =   270
      Index           =   13
      Left            =   3465
      Picture         =   "TitelBar.ctx":4892
      Top             =   1845
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgMin 
      Height          =   270
      Index           =   12
      Left            =   4005
      Picture         =   "TitelBar.ctx":4EC4
      Top             =   1305
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgMin 
      Height          =   270
      Index           =   11
      Left            =   4005
      Picture         =   "TitelBar.ctx":54F6
      Top             =   855
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgClose 
      Height          =   270
      Index           =   12
      Left            =   3465
      Picture         =   "TitelBar.ctx":5B28
      Top             =   1305
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgClose 
      Height          =   270
      Index           =   11
      Left            =   3465
      Picture         =   "TitelBar.ctx":615A
      Top             =   855
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgClose 
      Height          =   240
      Index           =   1
      Left            =   1365
      Picture         =   "TitelBar.ctx":678C
      Top             =   810
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgClose 
      Height          =   240
      Index           =   2
      Left            =   1365
      Picture         =   "TitelBar.ctx":6B16
      Top             =   1050
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgClose 
      Height          =   240
      Index           =   3
      Left            =   1365
      Picture         =   "TitelBar.ctx":6EA0
      Top             =   1290
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgMin 
      Height          =   240
      Index           =   3
      Left            =   1605
      Picture         =   "TitelBar.ctx":722A
      Top             =   1290
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgMin 
      Height          =   240
      Index           =   2
      Left            =   1605
      Picture         =   "TitelBar.ctx":75B4
      Top             =   1050
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgMin 
      Height          =   240
      Index           =   1
      Left            =   1605
      Picture         =   "TitelBar.ctx":793E
      Top             =   810
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image TBARX 
      Height          =   240
      Left            =   3750
      Top             =   90
      Width           =   240
   End
   Begin VB.Image TBAR_ 
      Height          =   240
      Left            =   3495
      Top             =   90
      Width           =   240
   End
End
Attribute VB_Name = "TitelBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
' Titelbar V4.00
' (C) Jörg Pfeffer alias Peppa
' 16.10.2009 - english Styles and Maximized Button added
'
'
' Feel free to use this Control for your Projects and have fun !
' If you like to mention me, it would be very nice
'
'------------------------------------------------------------------





Option Explicit

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal fuFlags As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal hPal As Long, ByRef RGBColorOut As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To 32) As Byte
End Type

Private Const TRANSPARENT = 1
Private Const ANTIALIASED_QUALITY = 4
Private Const NONANTIALIASED_QUALITY = 3

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize

Public Enum TbarStyles
 None = 0
 LeftSimple = 1
 LeftStripes = 2
 LeftRidge = 3
 LeftStripesRidge = 4
 RightSimple = 5
 RightStripes = 6
 RightRidge = 7
 RightStripesRidge = 8
 Sweep = 9
 SweepRidgeSmall = 10
 SweepRidgeMiddle = 11
 SweepRidgeBig = 12
 ModernRight = 13
 Vista = 14
End Enum

Public Enum TCPositionsX
 tbbsAuto
 tbbsLeft
 tbbsCenter
 tbbsRight
End Enum

Public Enum TCPositionsY
 tbsAuto
 tbbsTop
 tbbsMiddle
 tbbsBottom
End Enum

Public Enum TCPosMode
 pManual
 pCenterX
 pCenterY
 pCenterXY
End Enum

Dim Trans_Over As Boolean
Dim VBorder As Byte
Dim lX As Integer
Dim lY As Integer
Dim IsActive As Byte

'DrawSateTypes
Const DST_COMPLEX = &H0
Const DST_TEXT = &H1
Const DST_PREFIXTEXT = &H2
Const DST_ICON = &H3
Const DST_BITMAP = &H4
Const DSS_NORMAL = &H0
Const DSS_UNION = &H10 ' Dither
Const DSS_DISABLED = &H20
Const DSS_MONO = &H80 ' Draw in colour of brush specified in hBrush
Const DSS_RIGHT = &H8000

'Default Property Values:
Const m_def_AutoRedraw = False
Const m_def_ToolTipText = ""
Const m_def_PicturePosMode = pManual
Const m_def_PicturePosX = 6
Const m_def_PicturePosY = 4
Const m_def_CaptionPosX = tbbsAuto
Const m_def_CaptionPosY = tbbsCenter
Const m_def_CaptionBorder = False
Const m_def_CaptionBorderColor = vbBlack

Const m_def_Caption3DWidth = 2
Const m_def_Caption3DTop = False
Const m_def_Caption3DBottom = True
Const m_def_Caption3DLeft = False
Const m_def_Caption3DRight = True
Const m_def_Style = 8
Const m_def_IconStyle = 0

'Property Variables:
Dim m_def_BackColorCover As Single
Dim m_def_BackColorV2Begin As OLE_COLOR
Dim m_def_BackColorV2End As OLE_COLOR
Dim m_def_BackColorV1Begin As OLE_COLOR
Dim m_def_BackColorV1End As OLE_COLOR
Dim m_def_CaptionColor As OLE_COLOR
Dim m_def_CaptionColorBack As OLE_COLOR
Dim m_def_CaptionShadowColor As OLE_COLOR
Dim m_def_BorderColorHighLight As OLE_COLOR
Dim m_def_BorderColorDarkLight As OLE_COLOR

Dim m_IconStyle As Byte
Dim m_BackColorCover As Single
Dim m_AutoRedraw As Boolean
Dim m_Style As Integer

Dim m_ShowClose As Boolean
Dim m_ShowMinimized As Boolean
Dim m_ShowMaximized As Boolean
Dim m_ShowCloseEnabled As Boolean
Dim m_ShowMinimizedEnabled As Boolean
Dim m_ShowMaximizedEnabled As Boolean

Dim m_DashBack As Boolean
Dim m_ToolTipText As String
Dim m_Caption As String
Dim m_PicturePosMode As TCPosMode
Dim m_PicturePosX As Integer
Dim m_PicturePosY As Integer
Dim m_CaptionPosX As TCPositionsX
Dim m_CaptionPosY As TCPositionsY
Dim m_PictureButton As StdPicture
Dim m_BackColorV2Begin As OLE_COLOR
Dim m_BackColorV2End As OLE_COLOR
Dim m_CaptionColorBack As OLE_COLOR
Dim m_BackColorV1Begin As OLE_COLOR
Dim m_BackColorV1End As OLE_COLOR
Dim m_CaptionColor As OLE_COLOR
Dim m_BorderNormal As Byte
Dim m_BorderColorHighLight As OLE_COLOR
Dim m_BorderColorDarkLight As OLE_COLOR
Dim m_Caption3DWidth As Byte
Dim m_Caption3DTop As Boolean
Dim m_Caption3DBottom As Boolean
Dim m_Caption3DLeft As Boolean
Dim m_Caption3DRight As Boolean
Dim m_CaptionBorder As Boolean
Dim m_CaptionBorderColor As OLE_COLOR
Dim m_CaptionShadowColor As OLE_COLOR

Public Function GetRealColor(ByVal Color As OLE_COLOR) As Long
  Dim R As Long
  R = OleTranslateColor(Color, 0, GetRealColor)
  If R <> 0 Then 'raise an error
  DoEvents
  End If
End Function

Public Property Get AutoRedraw() As Boolean
    AutoRedraw = m_AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    m_AutoRedraw = New_AutoRedraw
    PropertyChanged "AutoRedraw"
 '''Refresh
End Property

Public Property Get DashBack() As Boolean
    DashBack = m_DashBack
End Property

Public Property Let DashBack(ByVal New_DashBack As Boolean)
    m_DashBack = New_DashBack
    PropertyChanged "DashBack"
 Refresh
End Property



Public Property Get Style() As TbarStyles
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As TbarStyles)
    m_Style = New_Style
    PropertyChanged "Style"
 Refresh
End Property


Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
 Refresh
End Property


Public Property Get BackColorCover() As Single
    BackColorCover = m_BackColorCover
End Property

Public Property Let BackColorCover(ByVal New_BackColorCover As Single)
    m_BackColorCover = New_BackColorCover
    If m_BackColorCover >= 10 Then m_BackColorCover = 10
    If m_BackColorCover <= 0 Then m_BackColorCover = 1
    PropertyChanged "BackColorCover"
 Refresh
End Property


Public Property Get BackColorV2Begin() As OLE_COLOR
    BackColorV2Begin = m_BackColorV2Begin
End Property

Public Property Let BackColorV2Begin(ByVal New_BackColorV2Begin As OLE_COLOR)
    m_BackColorV2Begin = New_BackColorV2Begin
    PropertyChanged "BackColorV2Begin"
 Refresh
End Property

Public Property Get BackColorV2End() As OLE_COLOR
    BackColorV2End = m_BackColorV2End
End Property

Public Property Let BackColorV2End(ByVal New_BackColorV2End As OLE_COLOR)
    m_BackColorV2End = New_BackColorV2End
    PropertyChanged "BackColorV2End"
 Refresh
End Property

Public Property Get BackColorV1Begin() As OLE_COLOR
    BackColorV1Begin = m_BackColorV1Begin
End Property

Public Property Let BackColorV1Begin(ByVal New_BackColorV1Begin As OLE_COLOR)
    m_BackColorV1Begin = New_BackColorV1Begin
    PropertyChanged "BackColorV1Begin"
 Refresh
End Property

Public Property Get BackColorV1End() As OLE_COLOR
    BackColorV1End = m_BackColorV1End
End Property

Public Property Let BackColorV1End(ByVal New_BackColorV1End As OLE_COLOR)
    m_BackColorV1End = New_BackColorV1End
    PropertyChanged "BackColorV1End"
 Refresh
End Property

Public Property Get BorderColorHighLight() As OLE_COLOR
    BorderColorHighLight = m_BorderColorHighLight
End Property

Public Property Let BorderColorHighLight(ByVal New_BorderColorHighLight As OLE_COLOR)
    m_BorderColorHighLight = New_BorderColorHighLight
    PropertyChanged "BorderColorHighLight"
 Refresh
End Property

Public Property Get BorderColorDarkLight() As OLE_COLOR
    BorderColorDarkLight = m_BorderColorDarkLight
End Property

Public Property Let BorderColorDarkLight(ByVal New_BorderColorDarkLight As OLE_COLOR)
    m_BorderColorDarkLight = New_BorderColorDarkLight
    PropertyChanged "BorderColorDarkLight"
 Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
 Refresh
End Property

Public Property Get CaptionColorBack() As OLE_COLOR
    CaptionColorBack = m_CaptionColorBack
End Property

Public Property Let CaptionColorBack(ByVal New_CaptionColorBack As OLE_COLOR)
    m_CaptionColorBack = New_CaptionColorBack
    PropertyChanged "CaptionColorBack"
 Refresh
End Property

Public Property Get CaptionColor() As OLE_COLOR
    CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
    m_CaptionColor = New_CaptionColor
    PropertyChanged "CaptionColor"
    UserControl.ForeColor = CaptionColor
 Refresh
End Property

Public Property Get CaptionBorderColor() As OLE_COLOR
    CaptionBorderColor = m_CaptionBorderColor
End Property

Public Property Let CaptionBorderColor(ByVal New_CaptionBorderColor As OLE_COLOR)
    m_CaptionBorderColor = New_CaptionBorderColor
    PropertyChanged "CaptionBorderColor"
 Refresh
End Property

Public Property Get CaptionShadowColor() As OLE_COLOR
    CaptionShadowColor = m_CaptionShadowColor
End Property

Public Property Let CaptionShadowColor(ByVal New_CaptionShadowColor As OLE_COLOR)
    m_CaptionShadowColor = New_CaptionShadowColor
    PropertyChanged "CaptionShadowColor"
 Refresh
End Property


Public Property Get CaptionBorder() As Boolean
    CaptionBorder = m_CaptionBorder
End Property

Public Property Let CaptionBorder(ByVal New_CaptionBorder As Boolean)
    m_CaptionBorder = New_CaptionBorder
    PropertyChanged "CaptionBorder"
 Refresh
End Property


Public Property Get Caption3DWidth() As Byte
    Caption3DWidth = m_Caption3DWidth
End Property

Public Property Let Caption3DWidth(ByVal New_Caption3DWidth As Byte)
    m_Caption3DWidth = New_Caption3DWidth
    PropertyChanged "Caption3DWidth"
 Refresh
End Property

Public Property Get Caption3DTop() As Boolean
    Caption3DTop = m_Caption3DTop
End Property

Public Property Let Caption3DTop(ByVal New_Caption3DTop As Boolean)
    m_Caption3DTop = New_Caption3DTop
    PropertyChanged "Caption3DappTop"
 Refresh
End Property


Public Property Get Caption3DBottom() As Boolean
    Caption3DBottom = m_Caption3DBottom
End Property

Public Property Let Caption3DBottom(ByVal New_Caption3DBottom As Boolean)
    m_Caption3DBottom = New_Caption3DBottom
    PropertyChanged "Caption3DappBottom"
 Refresh
End Property

Public Property Get Caption3DLeft() As Boolean
    Caption3DLeft = m_Caption3DLeft
End Property

Public Property Let Caption3DLeft(ByVal New_Caption3DLeft As Boolean)
    m_Caption3DLeft = New_Caption3DLeft
    PropertyChanged "Caption3DappLeft"
 Refresh
End Property

Public Property Get Caption3DRight() As Boolean
    Caption3DRight = m_Caption3DRight
End Property

Public Property Let Caption3DRight(ByVal New_Caption3DRight As Boolean)
    m_Caption3DRight = New_Caption3DRight
    PropertyChanged "Caption3DappRight"
 Refresh
End Property

Public Property Get ShowClose() As Boolean
    ShowClose = m_ShowClose
End Property

Public Property Let ShowClose(ByVal New_ShowClose As Boolean)
    m_ShowClose = New_ShowClose
    PropertyChanged "ShowClose"
 Call SetTBX_
End Property

Public Property Get ShowMinimized() As Boolean
    ShowMinimized = m_ShowMinimized
End Property

Public Property Let ShowMinimized(ByVal New_ShowMinimized As Boolean)
    m_ShowMinimized = New_ShowMinimized
    PropertyChanged "ShowMinimized"
 Call SetTBX_
End Property


Public Property Get ShowMaximized() As Boolean
    ShowMaximized = m_ShowMaximized
End Property

Public Property Let ShowMaximized(ByVal New_ShowMaximized As Boolean)
    m_ShowMaximized = New_ShowMaximized
    PropertyChanged "ShowMaximized"
 Call SetTBX_
End Property



Public Property Get ShowCloseEnabled() As Boolean
    ShowCloseEnabled = m_ShowCloseEnabled
End Property

Public Property Let ShowCloseEnabled(ByVal New_ShowCloseEnabled As Boolean)
    m_ShowCloseEnabled = New_ShowCloseEnabled
    PropertyChanged "ShowCloseEnabled"
 Call SetTBX_
End Property

Public Property Get ShowMinimizedEnabled() As Boolean
    ShowMinimizedEnabled = m_ShowMinimizedEnabled
End Property

Public Property Let ShowMinimizedEnabled(ByVal New_ShowMinimizedEnabled As Boolean)
    m_ShowMinimizedEnabled = New_ShowMinimizedEnabled
    PropertyChanged "ShowMinimizedEnabled"
 Call SetTBX_
End Property

Public Property Get ShowMaximizedEnabled() As Boolean
    ShowMaximizedEnabled = m_ShowMaximizedEnabled
End Property

Public Property Let ShowMaximizedEnabled(ByVal New_ShowMaximizedEnabled As Boolean)
    m_ShowMaximizedEnabled = New_ShowMaximizedEnabled
    PropertyChanged "ShowMaximizedEnabled"
 Call SetTBX_
End Property



Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
 Refresh
End Property

Public Property Get ToolTipText() As String
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
    '''Refresh
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
 Refresh
End Property


Public Property Get CaptionPosX() As TCPositionsX
    CaptionPosX = m_CaptionPosX
End Property

Public Property Let CaptionPosX(ByVal New_CaptionPosX As TCPositionsX)
    m_CaptionPosX = New_CaptionPosX
    PropertyChanged "CaptionPosX"
 Refresh
End Property

Public Property Get CaptionPosY() As TCPositionsY
    CaptionPosY = m_CaptionPosY
End Property

Public Property Let CaptionPosY(ByVal New_CaptionPosY As TCPositionsY)
    m_CaptionPosY = New_CaptionPosY
    PropertyChanged "CaptionPosY"
 Refresh
End Property

Public Property Get BorderNormal() As Byte
    BorderNormal = m_BorderNormal
End Property

Public Property Let BorderNormal(ByVal New_BorderNormal As Byte)
    m_BorderNormal = New_BorderNormal
    PropertyChanged "BorderNormal"
 Refresh
End Property


Public Property Get IconStyle() As Byte
    IconStyle = m_IconStyle
End Property

Public Property Let IconStyle(ByVal New_IconStyle As Byte)
    m_IconStyle = New_IconStyle
    If m_IconStyle >= 2 Then m_IconStyle = 2
    PropertyChanged "IconStyle"
 Refresh
End Property


Public Property Get Picture() As Picture
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
  Call SetImageI
End Property

Public Property Get PictureButton() As StdPicture
    Set PictureButton = m_PictureButton
End Property

Public Property Set PictureButton(ByVal New_PictureButton As StdPicture)
    Set m_PictureButton = New_PictureButton
    PropertyChanged "PictureButton"
' Call SetImageI
Refresh
End Property


Public Property Get PicturePosMode() As TCPosMode
    PicturePosMode = m_PicturePosMode
End Property

Public Property Let PicturePosMode(ByVal New_PicturePosMode As TCPosMode)
 m_PicturePosMode = New_PicturePosMode
 PropertyChanged "PicturePosMode"
'  Call SetImageI
Refresh
End Property


Public Property Get PicturePosX() As Integer
    PicturePosX = m_PicturePosX
End Property

Public Property Let PicturePosX(ByVal New_PicturePosX As Integer)
    m_PicturePosX = New_PicturePosX
    PropertyChanged "PicturePosX"
' Call SetImageI
Refresh
End Property

Public Property Get PicturePosY() As Integer
    PicturePosY = m_PicturePosY
End Property

Public Property Let PicturePosY(ByVal New_PicturePosY As Integer)
    m_PicturePosY = New_PicturePosY
    PropertyChanged "PicturePosY"
'  Call SetImageI
Refresh
End Property





Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get HasDC() As Boolean
    HasDC = UserControl.HasDC
End Property


Public Sub Cls()
    UserControl.Cls
End Sub

Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property



Public Sub TBAR__MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Call ReleaseCapture
 IsActive = 255
 Trans_Over = True
Extender.Parent.WindowState = 1
DoEvents
End Sub

Private Sub TBAR__MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Call ReleaseCapture
 IsActive = 255
 Trans_Over = True

If ShowMinimizedEnabled = True Then
 If TBAR_.Picture <> imgMin(1 + (m_IconStyle * 10)).Picture Then Set TBAR_.Picture = imgMin(1 + (m_IconStyle * 10)).Picture
End If

End Sub


Private Sub TBARMax_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Call ReleaseCapture
 IsActive = 255
 Trans_Over = True

If ShowMaximizedEnabled = True Then
 If TBARMax.Picture <> imgMax(1 + (m_IconStyle * 10)).Picture Then Set TBARMax.Picture = imgMax(1 + (m_IconStyle * 10)).Picture
End If

End Sub


Private Sub TBARMax_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Call ReleaseCapture
 IsActive = 255
 Trans_Over = True

 If Extender.Parent.WindowState = 0 Then
   Extender.Parent.WindowState = 2
 Else
   Extender.Parent.WindowState = 0
 End If

DoEvents
End Sub

Public Sub TBARX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Call ReleaseCapture
 IsActive = 255
 Trans_Over = True
Unload Extender.Parent
DoEvents
End Sub

Private Sub TBARX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Timer1.Enabled = False
Call ReleaseCapture
 IsActive = 255
 Trans_Over = True

If ShowCloseEnabled = True Then
 If TBARX.Picture <> imgClose(1 + (m_IconStyle * 10)).Picture Then Set TBARX.Picture = imgClose(1 + (m_IconStyle * 10)).Picture
End If


End Sub

Public Sub UserControl_Resize()
 RaiseEvent Resize
Refresh
End Sub

Public Sub UserControl_Click()
RaiseEvent Click
'''Refresh
End Sub


Public Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Public Sub UserControl_KeyPress(KeyAscii As Integer)
 RaiseEvent KeyPress(KeyAscii)
End Sub

Public Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If GetCapture <> UserControl.hwnd Then SetCapture UserControl.hwnd
 IsActive = 255
 Trans_Over = True
 '''Refresh
 RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 If (X > UserControl.TBARX.Left) And (X < UserControl.TBARX.Left + UserControl.TBARX.Width) And _
 (Y > UserControl.TBARX.Top) And (Y < UserControl.TBARX.Top + UserControl.TBARX.Height) Then ReleaseCapture: Exit Sub
 
 If (X > UserControl.TBAR_.Left) And (X < UserControl.TBAR_.Left + UserControl.TBAR_.Width) And _
 (Y > UserControl.TBAR_.Top) And (Y < UserControl.TBAR_.Top + UserControl.TBAR_.Height) Then ReleaseCapture: Exit Sub
 
 If (X > UserControl.TBARMax.Left) And (X < UserControl.TBARMax.Left + UserControl.TBARMax.Width) And _
 (Y > UserControl.TBARMax.Top) And (Y < UserControl.TBARMax.Top + UserControl.TBARMax.Height) Then ReleaseCapture: Exit Sub
 
 

If ShowCloseEnabled = True Then
 If TBARX.Picture <> imgClose(2 + (m_IconStyle * 10)).Picture Then Set TBARX.Picture = imgClose(2 + (m_IconStyle * 10)).Picture
End If

If ShowMaximizedEnabled = True Then
 If TBARMax.Picture <> imgMax(2 + (m_IconStyle * 10)).Picture Then Set TBARMax.Picture = imgMax(2 + (m_IconStyle * 10)).Picture
End If

If ShowMinimizedEnabled = True Then
 If TBAR_.Picture <> imgMin(2 + (m_IconStyle * 10)).Picture Then Set TBAR_.Picture = imgMin(2 + (m_IconStyle * 10)).Picture
End If

If IsActive = 255 Then
 If pCursorInWindow = True Then
  Trans_Over = True
  If GetCapture <> UserControl.hwnd Then SetCapture UserControl.hwnd
  IsActive = 1
'''  Refresh
  Timer1.Enabled = True
 End If
End If
    
 If Button = 1 Then
  Trans_Over = False
  IsActive = 255
  Timer1.Enabled = False
  Call ReleaseCapture
  Call SendMessage(Extender.Parent.hwnd, &HA1, 2, 0&)
 End If
    
    
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 IsActive = 255
 Trans_Over = False
 '''Refresh
 Call ReleaseCapture
 RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub Timer1_Timer()
If pCursorInWindow = False Then
  Timer1.Enabled = False
  Trans_Over = False
  IsActive = 255
  If GetCapture = UserControl.hwnd Then ReleaseCapture
  '''Refresh
End If

End Sub


Private Function pCursorInWindow() As Boolean
    Dim R As RECT
    Dim pt As POINTAPI
    Dim lRet As Long
    
    lRet = GetClientRect(UserControl.hwnd, R)
    lRet = GetCursorPos(pt)
    lRet = ScreenToClient(UserControl.hwnd, pt)
    
    'PtInRect is a bit flaky, so go for the manual method
    pCursorInWindow = Not (pt.X < 0 Or pt.X > UserControl.ScaleWidth Or pt.Y < 0 Or pt.Y > UserControl.ScaleHeight)
    
End Function




Public Sub UserControl_InitProperties()
'Set UserControl.Font = Ambient.Font
'm_AutoRedraw = m_def_AutoRedraw
'm_ToolTipText = m_def_ToolTipText
'm_BackColorCover = m_def_BackColorCover

Extender.Align = 1
UserControl.Height = 21
m_Caption = UserControl.name
m_PicturePosMode = m_def_PicturePosMode
m_PicturePosX = m_def_PicturePosX
m_PicturePosY = m_def_PicturePosY
m_CaptionPosX = m_def_CaptionPosX
m_CaptionPosY = m_def_CaptionPosY
m_BackColorV1Begin = m_def_BackColorV1Begin
m_BackColorV1End = m_def_BackColorV1End
m_CaptionColor = m_def_CaptionColor
m_BackColorV2Begin = m_def_BackColorV2Begin
m_BackColorV2End = m_def_BackColorV2End
m_CaptionColorBack = m_def_CaptionColorBack
UserControl.BackColor = m_def_BackColorV2End
UserControl.ForeColor = m_def_CaptionColor
m_BorderNormal = 2
m_BackColorCover = 3
m_BorderColorHighLight = m_def_BorderColorHighLight
m_BorderColorDarkLight = m_def_BorderColorDarkLight
m_Caption3DWidth = m_def_Caption3DWidth
m_Caption3DTop = m_def_Caption3DTop
m_Caption3DLeft = m_def_Caption3DLeft
m_Caption3DRight = m_def_Caption3DRight
m_Caption3DBottom = m_def_Caption3DBottom
m_CaptionBorder = m_def_CaptionBorder
m_CaptionBorderColor = m_def_CaptionBorderColor
m_CaptionShadowColor = m_def_CaptionShadowColor
m_ShowClose = True
m_ShowMinimized = True
m_ShowMaximized = False
m_ShowCloseEnabled = True
m_ShowMinimizedEnabled = True
m_ShowMaximizedEnabled = True
m_Style = m_def_Style
m_IconStyle = m_def_IconStyle
Set m_PictureButton = Nothing
 
 
IsActive = 255
Refresh
End Sub

Public Sub UserControl_Initialize()
IsActive = 255
'Set TBARX.Picture = imgClose(2).Picture
'Set TBAR_.Picture = imgMin(2).Picture

m_def_BackColorV2Begin = RGB(90, 110, 140)
m_def_BackColorV2End = RGB(38, 48, 79)
m_def_BackColorV1Begin = RGB(230, 230, 230)
m_def_BackColorV1End = RGB(90, 110, 140)
m_def_CaptionColor = RGB(255, 255, 255)
m_def_CaptionColorBack = RGB(60, 60, 60)
m_def_CaptionShadowColor = RGB(60, 60, 60)
m_def_BorderColorHighLight = RGB(255, 255, 255)
m_def_BorderColorDarkLight = RGB(0, 0, 0)

End Sub


Public Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", m_def_BackColorV1Begin)
    m_AutoRedraw = PropBag.ReadProperty("AutoRedraw", m_def_AutoRedraw)
    m_DashBack = PropBag.ReadProperty("DashBack", False)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_BackColorCover = PropBag.ReadProperty("BackColorCover", m_def_BackColorCover)
    m_BackColorV2Begin = PropBag.ReadProperty("BackColorV2Begin", m_def_BackColorV2Begin)
    m_BackColorV2End = PropBag.ReadProperty("BackColorV2End", m_def_BackColorV2End)
    m_BackColorV1Begin = PropBag.ReadProperty("BackColorV1Begin", m_def_BackColorV1Begin)
    m_BackColorV1End = PropBag.ReadProperty("BackColorV1End", m_def_BackColorV1End)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_CaptionColorBack = PropBag.ReadProperty("CaptionColorBack", m_def_CaptionColorBack)
    m_CaptionColor = PropBag.ReadProperty("CaptionColor", m_def_CaptionColor)
    m_ShowClose = PropBag.ReadProperty("ShowClose", True)
    m_ShowMinimized = PropBag.ReadProperty("ShowMinimized", True)
    m_ShowMaximized = PropBag.ReadProperty("ShowMaximized", False)
    m_ShowCloseEnabled = PropBag.ReadProperty("ShowCloseEnabled", True)
    m_ShowMinimizedEnabled = PropBag.ReadProperty("ShowMinimizedEnabled", True)
    m_ShowMaximizedEnabled = PropBag.ReadProperty("ShowMaximizedEnabled", False)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    m_Caption = PropBag.ReadProperty("Caption", UserControl.name)
    m_PicturePosMode = PropBag.ReadProperty("PicturePosMode", m_def_PicturePosMode)
    m_PicturePosX = PropBag.ReadProperty("PicturePosX", m_def_PicturePosX)
    m_PicturePosY = PropBag.ReadProperty("PicturePosY", m_def_PicturePosY)
    m_CaptionPosX = PropBag.ReadProperty("CaptionPosX", m_def_CaptionPosX)
    m_CaptionPosY = PropBag.ReadProperty("CaptionPosY", m_def_CaptionPosY)
    m_BorderNormal = PropBag.ReadProperty("BorderNormal", 0)
    m_IconStyle = PropBag.ReadProperty("IconStyle", 0)
    m_BorderColorHighLight = PropBag.ReadProperty("BorderColorHighLight", m_def_BorderColorHighLight)
    m_BorderColorDarkLight = PropBag.ReadProperty("BorderColorDarkLight", m_def_BorderColorDarkLight)
    m_Caption3DWidth = PropBag.ReadProperty("Caption3DWidth", m_def_Caption3DWidth)
    m_Caption3DTop = PropBag.ReadProperty("Caption3DTop", m_def_Caption3DTop)
    m_Caption3DLeft = PropBag.ReadProperty("Caption3DLeft", m_def_Caption3DLeft)
    m_Caption3DRight = PropBag.ReadProperty("Caption3DRight", m_def_Caption3DRight)
    m_Caption3DBottom = PropBag.ReadProperty("Caption3DBottom", m_def_Caption3DBottom)
    m_CaptionBorder = PropBag.ReadProperty("CaptionBorder", m_def_CaptionBorder)
    m_CaptionBorderColor = PropBag.ReadProperty("CaptionBorderColor", m_def_CaptionBorderColor)
    m_CaptionShadowColor = PropBag.ReadProperty("CaptionShadowColor", m_def_CaptionShadowColor)
    Set m_PictureButton = PropBag.ReadProperty("PictureButton", Nothing)
    
    
Refresh
End Sub

Public Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("AutoRedraw", m_AutoRedraw, m_def_AutoRedraw)
    Call PropBag.WriteProperty("DashBack", m_DashBack, False)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("BackColorCover", m_BackColorCover, m_def_BackColorCover)
    Call PropBag.WriteProperty("BackColorV2Begin", m_BackColorV2Begin, m_def_BackColorV2Begin)
    Call PropBag.WriteProperty("BackColorV2End", m_BackColorV2End, m_def_BackColorV2End)
    Call PropBag.WriteProperty("BackColorV1Begin", m_BackColorV1Begin, m_def_BackColorV1Begin)
    Call PropBag.WriteProperty("BackColorV1End", m_BackColorV1End, m_def_BackColorV1End)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("CaptionColorBack", m_CaptionColorBack, m_def_CaptionColorBack)
    Call PropBag.WriteProperty("CaptionColor", m_CaptionColor, m_def_CaptionColor)
    Call PropBag.WriteProperty("ShowClose", m_ShowClose, True)
    Call PropBag.WriteProperty("ShowMinimized", m_ShowMinimized, True)
    Call PropBag.WriteProperty("ShowMaximized", m_ShowMaximized, False)
    Call PropBag.WriteProperty("ShowCloseEnabled", m_ShowCloseEnabled, True)
    Call PropBag.WriteProperty("ShowMinimizedEnabled", m_ShowMinimizedEnabled, True)
    Call PropBag.WriteProperty("ShowMaximizedEnabled", m_ShowMaximizedEnabled, False)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("Caption", m_Caption, UserControl.name)
    Call PropBag.WriteProperty("PicturePosMode", m_PicturePosMode, m_def_PicturePosMode)
    Call PropBag.WriteProperty("PicturePosX", m_PicturePosX, m_def_PicturePosX)
    Call PropBag.WriteProperty("PicturePosY", m_PicturePosY, m_def_PicturePosY)
    Call PropBag.WriteProperty("CaptionPosX", m_CaptionPosX, m_def_CaptionPosX)
    Call PropBag.WriteProperty("CaptionPosY", m_CaptionPosY, m_def_CaptionPosY)
    Call PropBag.WriteProperty("PictureButton", m_PictureButton, Nothing)
    Call PropBag.WriteProperty("BorderNormal", m_BorderNormal, 0)
    Call PropBag.WriteProperty("IconStyle", m_IconStyle, 0)
    Call PropBag.WriteProperty("BorderColorHighLight", m_BorderColorHighLight, m_def_BorderColorHighLight)
    Call PropBag.WriteProperty("BorderColorDarkLight", m_BorderColorDarkLight, m_def_BorderColorDarkLight)
    Call PropBag.WriteProperty("Caption3DWidth", m_Caption3DWidth, m_def_Caption3DWidth)
    Call PropBag.WriteProperty("Caption3DTop", m_Caption3DTop, m_def_Caption3DTop)
    Call PropBag.WriteProperty("Caption3DLeft", m_Caption3DLeft, m_def_Caption3DLeft)
    Call PropBag.WriteProperty("Caption3DRight", m_Caption3DRight, m_def_Caption3DRight)
    Call PropBag.WriteProperty("Caption3DBottom", m_Caption3DBottom, m_def_Caption3DBottom)
    Call PropBag.WriteProperty("CaptionBorder", m_CaptionBorder, m_def_CaptionBorder)
    Call PropBag.WriteProperty("CaptionBorderColor", m_CaptionBorderColor, m_def_CaptionBorderColor)
    Call PropBag.WriteProperty("CaptionShadowColor", m_CaptionShadowColor, m_def_CaptionShadowColor)
   
'''Refresh

End Sub


Public Sub Refresh()

'MsgBox (Time)
UserControl.AutoRedraw = True
UserControl.Cls

UserControl.ForeColor = CaptionColor
Call SetTBX_
Call SetVerlauf(m_BackColorV1Begin, m_BackColorV1End, m_BackColorV2Begin, m_BackColorV2End)
Call SetImageI
Call Shadow
Call BSetBorder(m_BorderNormal)

UserControl.AutoRedraw = False

End Sub


Public Sub SetImageI()
Dim eDraw As Long
eDraw = DST_BITMAP
 
 Set Image1 = m_PictureButton
 If Image1 <> 0 Then
  If Image1.Picture.Type = vbPicTypeIcon Then eDraw = DST_ICON
  eDraw = eDraw Or DSS_NORMAL Or DSS_RIGHT
  If m_PicturePosMode <> pManual Then
   If m_PicturePosMode = pCenterX Or m_PicturePosMode = pCenterXY Then m_PicturePosX = (UserControl.ScaleWidth - Image1.Width) / 2
   If m_PicturePosMode = pCenterY Or m_PicturePosMode = pCenterXY Then m_PicturePosY = (UserControl.ScaleHeight - Image1.Height) / 2
  End If
  DrawState UserControl.hdc, 0, 0, Image1.Picture.handle, 0, m_PicturePosX, m_PicturePosY, Image1.Width, Image1.Height, eDraw
 End If

End Sub

Public Sub SetTBX_()
Dim BT As Single
Dim BT2 As Byte
Dim ICS As Byte
Dim FP As Integer


BT = 3: BT2 = 0
If m_Style = 13 Then BT = UserControl.TBARX.Width: BT2 = 2
If m_Style = 14 Then BT = 6: BT2 = 2
 
 UserControl.TBARX.Visible = m_ShowClose
 UserControl.TBAR_.Visible = m_ShowMinimized
 UserControl.TBARMax.Visible = m_ShowMaximized
 
 ICS = m_IconStyle * 10
 
 If ShowCloseEnabled = True Then
   If TBARX.Picture <> imgClose(2 + ICS).Picture Then Set TBARX.Picture = imgClose(2 + ICS).Picture
 Else
   If TBARX.Picture <> imgClose(3 + ICS).Picture Then Set TBARX.Picture = imgClose(3 + ICS).Picture
 End If
 
 If ShowMaximizedEnabled = True Then
   If TBARMax.Picture <> imgMax(2 + ICS).Picture Then Set TBARMax.Picture = imgMax(2 + ICS).Picture
 Else
   If TBARMax.Picture <> imgMax(3 + ICS).Picture Then Set TBARMax.Picture = imgMax(3 + ICS).Picture
 End If
 
 If ShowMinimizedEnabled = True Then
   If TBAR_.Picture <> imgMin(2 + ICS).Picture Then Set TBAR_.Picture = imgMin(2 + ICS).Picture
 Else
   If TBAR_.Picture <> imgMin(3 + ICS).Picture Then Set TBAR_.Picture = imgMin(3 + ICS).Picture
 End If
  
  
  
 FP = UserControl.ScaleWidth
 
 If ShowClose = True Then
   UserControl.TBARX.Left = FP - UserControl.TBARX.Width - 3 - BT
   UserControl.TBARX.Top = ((UserControl.ScaleHeight - UserControl.TBARX.Height) / 2) + BT2
   FP = UserControl.TBARX.Left: BT = 3
 End If
 
 If ShowMaximized = True Then
   UserControl.TBARMax.Left = FP - UserControl.TBARMax.Width - 3 - BT
   UserControl.TBARMax.Top = ((UserControl.ScaleHeight - UserControl.TBARMax.Height) / 2) + BT2
   FP = UserControl.TBARMax.Left
 End If
 
 If ShowMinimized = True Then
   UserControl.TBAR_.Left = FP - UserControl.TBAR_.Width - 6
   UserControl.TBAR_.Top = ((UserControl.ScaleHeight - UserControl.TBAR_.Height) / 2) + BT2
   FP = UserControl.TBAR_.Left
 End If

End Sub



Sub SetVerlauf(BegColor, EndColor, BegColor2, EndColor2)
Dim mH As Long
Dim mD As Long
Dim iRet As Long
Dim h1 As Long
Dim h2 As Long
Dim h3 As Long
Dim D1 As Long
Dim D2 As Long
Dim d3 As Long
Dim cd1 As Single
Dim cd2 As Single
Dim cd3 As Single
Dim hb1 As Long
Dim hb2 As Long
Dim hb3 As Long
Dim db1 As Long
Dim db2 As Long
Dim db3 As Long
Dim cb1 As Single
Dim cb2 As Single
Dim cb3 As Single
Dim hc1 As Long
Dim hc2 As Long
Dim hc3 As Long
Dim dc1 As Long
Dim dc2 As Long
Dim dc3 As Long
Dim cc1 As Single
Dim cc2 As Single
Dim cc3 As Single


Dim ucH As Integer
Dim T As Integer
Dim BT As Single

Dim Sc
Dim Z As Integer
Dim VT
Dim t2
Dim SZ
Dim iAdder As Byte

Dim Xw As Integer
Dim Xy As Integer
Dim Xh As Integer
Dim Xs As Integer

UserControl.DrawWidth = 1

iRet = BegColor
mH = GetRealColor(iRet)
h1 = mH Mod 256
h2 = (mH \ 256) Mod 256
h3 = mH \ 256 \ 256

iRet = EndColor
mD = GetRealColor(iRet)
D1 = mD Mod 256
D2 = (mD \ 256) Mod 256
d3 = mD \ 256 \ 256

ucH = UserControl.ScaleHeight
If m_Style = 14 Then ucH = ucH / 2
If ucH <= 0 Then ucH = 25

cd1 = ((h1 - D1) / ucH)
cd2 = ((h2 - D2) / ucH)
cd3 = ((h3 - d3) / ucH)

iRet = BegColor2
mH = GetRealColor(iRet)
hb1 = mH Mod 256
hb2 = (mH \ 256) Mod 256
hb3 = mH \ 256 \ 256

iRet = EndColor2
mD = GetRealColor(iRet)
db1 = mD Mod 256
db2 = (mD \ 256) Mod 256
db3 = mD \ 256 \ 256

cb1 = ((hb1 - db1) / ucH)
cb2 = ((hb2 - db2) / ucH)
cb3 = ((hb3 - db3) / ucH)

iRet = m_BorderColorHighLight
mH = GetRealColor(iRet)
hc1 = mH Mod 256
hc2 = (mH \ 256) Mod 256
hc3 = mH \ 256 \ 256

iRet = m_BorderColorDarkLight
mD = GetRealColor(iRet)
dc1 = mD Mod 256
dc2 = (mD \ 256) Mod 256
dc3 = mD \ 256 \ 256

cc1 = ((hc1 - dc1) / (ucH * 2))
cc2 = ((hc2 - dc2) / (ucH * 2))
cc3 = ((hc3 - dc3) / (ucH * 2))


BT = m_BackColorCover / 10

If m_Style = 3 Or m_Style = 4 Or m_Style = 7 Or m_Style = 8 Then ucH = ucH - (m_BorderNormal + 1)
If m_Style = 10 Then ucH = ucH - (m_BorderNormal + 2)
If m_Style = 11 Then ucH = ucH - (m_BorderNormal + 5)
If m_Style = 12 Then ucH = ucH - (m_BorderNormal + 7)

If m_Style = 9 Or m_Style = 10 Or m_Style = 11 Or m_Style = 12 Then
 Sc = (2 / ucH)
 Z = (UserControl.ScaleHeight - ucH) / 2
 SZ = Z
 
 If ucH <> UserControl.ScaleHeight Then
  For T = 0 To Z
   UserControl.Line (0, T)-(UserControl.ScaleWidth, T), RGB(hb1 - (cb1 * T), hb2 - (cb2 * T), hb3 - (cb3 * T))
  Next T
 End If

 For t2 = -1 To 1 Step Sc
  VT = Tan(t2) * (ucH / 1.3)
 
  UserControl.Line (0, Z)-(Fix(UserControl.ScaleWidth * BT) + VT, Z), RGB(h1 - (cd1 * Z), h2 - (cd2 * Z), h3 - (cd3 * Z))
  UserControl.Line (Fix(UserControl.ScaleWidth * BT) + VT, Z)-(UserControl.ScaleWidth, Z), RGB(hb1 - (cb1 * Z), hb2 - (cb2 * Z), hb3 - (cb3 * Z))

  Z = Z + 1
 Next t2
 
 If ucH <> UserControl.ScaleHeight Then
  For T = (UserControl.ScaleHeight - SZ) To UserControl.ScaleHeight
   UserControl.Line (0, T)-(UserControl.ScaleWidth, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
  Next T
 End If
 
 GoTo Xout:
End If

If m_Style = 0 Then
 For T = 0 To ucH
  UserControl.Line (0, T)-(UserControl.ScaleWidth, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
 Next T
 GoTo Xout:
End If

If m_Style = 1 Or m_Style = 2 Or m_Style = 3 Or m_Style = 4 Then
 For T = 0 To ucH
  UserControl.Line (0, T)-(Fix(UserControl.ScaleWidth * BT) - T, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
  UserControl.Line (Fix(UserControl.ScaleWidth * BT) - T, T)-(UserControl.ScaleWidth, T), RGB(hb1 - (cb1 * T), hb2 - (cb2 * T), hb3 - (cb3 * T))
  If m_Style = 2 Or m_Style = 4 Then
   UserControl.Line (Fix(UserControl.ScaleWidth * BT) - T + 2, T)-(Fix(UserControl.ScaleWidth * BT) - T + 1, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
   UserControl.Line (Fix(UserControl.ScaleWidth * BT) - T + 6, T)-(Fix(UserControl.ScaleWidth * BT) - T + 5, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
   UserControl.Line (Fix(UserControl.ScaleWidth * BT) - T + 11, T)-(Fix(UserControl.ScaleWidth * BT) - T + 10, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
  End If
 Next T

 If ucH <> UserControl.ScaleHeight Then
  For T = ucH To UserControl.ScaleHeight
   UserControl.Line (0, T)-(UserControl.ScaleWidth, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
  Next T
 End If

GoTo Xout:
End If

If m_Style = 5 Or m_Style = 6 Or m_Style = 7 Or m_Style = 8 Then
 For T = 0 To ucH
  UserControl.Line (0, T)-(Fix(UserControl.ScaleWidth * BT) + T, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
  UserControl.Line (Fix(UserControl.ScaleWidth * BT) + T, T)-(UserControl.ScaleWidth, T), RGB(hb1 - (cb1 * T), hb2 - (cb2 * T), hb3 - (cb3 * T))
 Next T

  If m_Style = 6 Or m_Style = 8 Then
   For T = 0 To ucH
    UserControl.Line (Fix(UserControl.ScaleWidth * BT) + T + 2, T)-(Fix(UserControl.ScaleWidth * BT) + T + 1, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
    UserControl.Line (Fix(UserControl.ScaleWidth * BT) + T + 6, T)-(Fix(UserControl.ScaleWidth * BT) + T + 5, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
    UserControl.Line (Fix(UserControl.ScaleWidth * BT) + T + 11, T)-(Fix(UserControl.ScaleWidth * BT) + T + 10, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
   Next T
  End If
  
 If ucH <> UserControl.ScaleHeight Then
  For T = ucH To UserControl.ScaleHeight
   UserControl.Line (0, T)-(UserControl.ScaleWidth, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
  Next T
 End If

GoTo Xout:
End If


iAdder = 4
If m_Style = 13 Then
If m_ShowClose = False Or m_ShowMinimized = False Then GoTo Xout:
BT = UserControl.TBAR_.Left - ucH * 1.2
 
 For T = 0 To ucH
  UserControl.Line (0, T)-(BT + T, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
  If T > iAdder Then UserControl.Line (BT + T, T)-(UserControl.ScaleWidth - Fix(ucH / 2) + T, T), RGB(hb1 - (cb1 * T), hb2 - (cb2 * T), hb3 - (cb3 * T))
 Next T

 UserControl.Line (BT + iAdder, iAdder)-(UserControl.ScaleWidth - Fix(ucH / 2) + iAdder, iAdder), m_BorderColorDarkLight
 UserControl.Line (UserControl.ScaleWidth - Fix(ucH / 2) + iAdder, iAdder)-(UserControl.ScaleWidth + Fix(ucH / 2) + iAdder, ucH + iAdder), m_BorderColorDarkLight

 UserControl.Line (BT, 0)-(BT + ucH, ucH), m_BorderColorDarkLight

SZ = Fix((UserControl.TBARX.Height - 2) / 4)
VT = 3
Z = 0

For T = 0 To 3
  For t2 = 0 To VT
   UserControl.PSet (UserControl.TBAR_.Left - 5 - (t2 * 3), UserControl.TBARX.Top + 2 + (T * SZ)), RGB(90, 90, 90)
   UserControl.PSet (UserControl.TBAR_.Left - 5 - 1 - (t2 * 3), UserControl.TBARX.Top + 2 + (T * SZ) + 1), RGB(255, 255, 255)
  Next t2
  VT = VT - 1
 
  For t2 = 0 To Z
   UserControl.PSet (UserControl.TBARX.Left + UserControl.TBARX.Width + 3 + (t2 * 3), UserControl.TBARX.Top + 2 + (T * SZ)), RGB(90, 90, 90)
   UserControl.PSet (UserControl.TBARX.Left + UserControl.TBARX.Width + 3 + 1 + (t2 * 3), UserControl.TBARX.Top + 2 + (T * SZ) + 1), RGB(255, 255, 255)
  Next t2
  Z = Z + 1
Next T

GoTo Xout:
End If



iAdder = 0
If m_Style = 14 Then
  For T = 0 To ucH
   UserControl.Line (0, T)-(UserControl.ScaleWidth, T), RGB(h1 - (cd1 * T), h2 - (cd2 * T), h3 - (cd3 * T))
  Next T

  For T = 0 To ucH
   UserControl.Line (0, ucH + T)-(UserControl.ScaleWidth, ucH + T), RGB(hb1 - (cb1 * T), hb2 - (cb2 * T), hb3 - (cb3 * T))
  Next T

 BT = UserControl.TBAR_.Left - ucH - ucH * 1.2

 For T = 0 To (ucH * 2) - 1
  UserControl.Line (BT + T, T)-(UserControl.ScaleWidth, T), RGB(hc1 - (cc1 * T), hc2 - (cc2 * T), hc3 - (cc3 * T))
 Next T
 
 If m_BorderNormal >= 1 Then
  UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), m_BorderColorDarkLight
  UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), m_BorderColorDarkLight

  UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), m_BorderColorHighLight
  UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), m_BorderColorHighLight
 End If

 GoTo Xout:
End If




Xout:


If m_DashBack = True Then
 Xs = 6
 Xw = UserControl.TextWidth(m_Caption) * 1.2
 Xy = m_BorderNormal + 2
 Xh = UserControl.ScaleHeight - Xy - Xy
 
 StretchBlt UserControl.hdc, Xs, Xy, Xw, Xh, UserControl.hdc, Xs - 3, Xh + 2, 1, -(Xh), vbSrcCopy

 StretchBlt UserControl.hdc, Xs - 1, Xy + 1, 1, Xh - 2, UserControl.hdc, Xs - 3, Xh + 2, 1, -(Xh - 2), vbSrcCopy
 StretchBlt UserControl.hdc, Xs - 2, Xy + 2, 1, Xh - 4, UserControl.hdc, Xs - 3, Xh + 2, 1, -(Xh - 4), vbSrcCopy

 StretchBlt UserControl.hdc, Xw + Xs, Xy + 1, 1, Xh - 2, UserControl.hdc, Xs - 3, Xh + 2, 1, -(Xh - 2), vbSrcCopy
 StretchBlt UserControl.hdc, Xw + Xs + 1, Xy + 2, 1, Xh - 4, UserControl.hdc, Xs - 3, Xh + 2, 1, -(Xh - 4), vbSrcCopy
End If

End Sub

Sub BSetBorder(BorderWidth As Byte)

UserControl.DrawWidth = 1
If BorderWidth <= 0 Then BorderWidth = 0: Exit Sub

UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), m_BorderColorDarkLight
UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), m_BorderColorDarkLight

UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), m_BorderColorHighLight
UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), m_BorderColorHighLight



End Sub



Sub Shadow()
Dim hMasterFont As Long
Dim hFntPrev As Long
Dim NColor As Long
Dim ORed As Integer
Dim OBlue As Integer
Dim OGreen As Integer
Dim OutputSize As POINTAPI
Dim RColor As Integer
Dim GColor As Integer
Dim BColor As Integer
Dim BC As Long
Dim Tx As String
Dim arS As Boolean
Dim SDepth As Byte
Dim DX As Byte
Dim DY As Byte
Dim UX As Byte
Dim UY As Byte
Dim iRet As Long
Dim fc As Long
Dim I As Long
Dim CX As Integer
Dim CY As Integer
Dim T As Integer
Dim t2 As Integer
Dim J As Integer

SDepth = m_Caption3DWidth
Tx = m_Caption

DX = 0: DY = 0: UX = 0: UY = 0
If m_Caption3DTop = True Then UY = 1
If m_Caption3DBottom = True Then DY = 1
If m_Caption3DLeft = True Then UX = 1
If m_Caption3DRight = True Then DX = 1

iRet = m_CaptionShadowColor
fc = GetRealColor(iRet)
RColor = fc Mod 256
GColor = (fc \ 256) Mod 256
BColor = fc \ 65535

BC = 0
If m_CaptionBorder = True Then BC = m_CaptionBorderColor Or RGB(0, 0, 1)

If SDepth <> 0 Then
 'SDepth = SDepth + 1
 If BC <> 0 Then SDepth = SDepth + 1
End If

lX = 6
If Image1 <> 0 Then lX = lX + m_PicturePosX + Image1.Width

fc = UserControl.ForeColor
'I = GetRealColor(UserControl.BackColor)
I = GetRealColor(m_CaptionColorBack)

ORed = I Mod 256
OGreen = (I \ 256) Mod 256
OBlue = I \ 65536

arS = UserControl.AutoRedraw
UserControl.AutoRedraw = True
 hMasterFont = SpecialPrint(Tx, UserControl.FontName, UserControl.FontSize, 0, False)
 hFntPrev = SelectObject(UserControl.hdc, hMasterFont)
 SetBkMode UserControl.hdc, TRANSPARENT
 GetTextExtentPoint32 UserControl.hdc, Tx, Len(Tx), OutputSize
 
 lY = 0 ' m_BorderNormal + 1
 If m_CaptionPosY = tbbsAuto Then lY = (UserControl.ScaleHeight - OutputSize.Y) / 2
 If m_CaptionPosY = tbbsMiddle Then lY = Fix((UserControl.ScaleHeight - OutputSize.Y - m_BorderNormal) / 2)
 If m_CaptionPosY = tbbsBottom Then lY = UserControl.ScaleHeight - OutputSize.Y - m_BorderNormal
 
 If m_CaptionPosX = tbbsAuto Then lX = lX
 If m_CaptionPosX = tbbsLeft Then lX = m_BorderNormal + 1
 If m_CaptionPosX = tbbsCenter Then lX = ((UserControl.ScaleWidth - OutputSize.X) / 2)
 If m_CaptionPosX = tbbsRight Then lX = (UserControl.ScaleWidth - OutputSize.X) - m_BorderNormal - 1 - SDepth
 CX = Fix(lX)
 CY = Fix(lY)
  
 For T = SDepth To 1 Step -1
  t2 = SDepth - T
  NColor = RGB(((ORed / SDepth * T) + (RColor / SDepth * t2)), ((OGreen / SDepth * T) + (GColor / SDepth * t2)), ((OBlue / SDepth * T) + (BColor / SDepth * t2)))
  SetTextColor UserControl.hdc, NColor
 If DX = 1 And DY = 1 And UX = 1 And UY = 1 Then
  For J = -T To T
   TextOut UserControl.hdc, CX + J, CY + T, Tx, Len(Tx)
   TextOut UserControl.hdc, CX + J, CY - T, Tx, Len(Tx)
   TextOut UserControl.hdc, CX + T, CY + J, Tx, Len(Tx)
   TextOut UserControl.hdc, CX - T, CY - J, Tx, Len(Tx)
  Next J
 ElseIf DX = 1 And DY = 1 Then
  TextOut UserControl.hdc, CX + T, CY + T, Tx, Len(Tx)
 ElseIf UX = 1 And UY = 1 Then
   TextOut UserControl.hdc, CX - T, CY - T, Tx, Len(Tx)
 Else
  If DY = 1 Then TextOut UserControl.hdc, CX, CY + T, Tx, Len(Tx)
  If UY = 1 Then TextOut UserControl.hdc, CX, CY - T, Tx, Len(Tx)
  If DX = 1 Then TextOut UserControl.hdc, CX + T, CY, Tx, Len(Tx)
  If UX = 1 Then TextOut UserControl.hdc, CX - T, CY, Tx, Len(Tx)
 End If
Next T

 SelectObject UserControl.hdc, hFntPrev
 DeleteObject hMasterFont
 DeleteObject hFntPrev
 hMasterFont = SpecialPrint(Tx, UserControl.FontName, UserControl.FontSize, 0, True, 0, UserControl.FontBold)
 hFntPrev = SelectObject(UserControl.hdc, hMasterFont)

If BC <> 0 Then
 SetTextColor UserControl.hdc, BC

 TextOut UserControl.hdc, CX + 1, CY, Tx, Len(Tx)
 TextOut UserControl.hdc, CX + 1, CY + 1, Tx, Len(Tx)
 TextOut UserControl.hdc, CX + 1, CY - 1, Tx, Len(Tx)
 
 TextOut UserControl.hdc, CX - 1, CY, Tx, Len(Tx)
 TextOut UserControl.hdc, CX - 1, CY - 1, Tx, Len(Tx)
 TextOut UserControl.hdc, CX - 1, CY + 1, Tx, Len(Tx)
 
 TextOut UserControl.hdc, CX, CY - 1, Tx, Len(Tx)
 TextOut UserControl.hdc, CX, CY + 1, Tx, Len(Tx)
End If


SetTextColor UserControl.hdc, fc
TextOut UserControl.hdc, CX, CY, Tx, Len(Tx)

SelectObject UserControl.hdc, hFntPrev
DeleteObject hMasterFont
DeleteObject hFntPrev
UserControl.AutoRedraw = arS
UserControl.CurrentY = UserControl.CurrentY + UserControl.FontSize + 6

End Sub


Private Function SpecialPrint(Tx As String, FontName As String, FontSize As Integer, Optional XWidth As Integer = 0, Optional AntiAliased As Boolean = True, Optional Rotation As Integer = 0, Optional FWBold As Boolean = False) As Long
Dim Nfnt As LOGFONT, I As Long
Dim hMasterFont As Long
Dim hFntPrev As Long


FontName = FontName + String(32 - Len(FontName), 0)
For I = 1 To 32
 Nfnt.lfFaceName(I) = Asc(Mid$(FontName, I, 1))
Next
    
Nfnt.lfHeight = CLng(FontSize * 1.5)
Nfnt.lfWidth = XWidth

If AntiAliased = True Then
 Nfnt.lfQuality = ANTIALIASED_QUALITY
Else
 Nfnt.lfQuality = NONANTIALIASED_QUALITY
End If

Nfnt.lfItalic = 0
Nfnt.lfStrikeOut = 0
Nfnt.lfUnderline = 0
Nfnt.lfWeight = 0

If FWBold = True Then Nfnt.lfWeight = 800
    
'If Rotation > 0 Then Nfnt.lfEscapement = CLng(Rotation) * 10
Nfnt.lfEscapement = 0.1

hFntPrev = CreateFontIndirect(Nfnt)
SelectObject UserControl.hdc, hFntPrev
DeleteObject hFntPrev

End Function

