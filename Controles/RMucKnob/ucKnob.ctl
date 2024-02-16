VERSION 5.00
Begin VB.UserControl ucKnob 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1995
   ClipBehavior    =   0  'None
   PropertyPages   =   "ucKnob.ctx":0000
   ScaleHeight     =   133
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   133
   ToolboxBitmap   =   "ucKnob.ctx":0019
   Windowless      =   -1  'True
End
Attribute VB_Name = "ucKnob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------
'Module Name: ucKnob
'Autor:  Leandro Ascierto
'Web: www.leandroascierto.com
'Date: 26/12/2021
'Version: 1.0.0
'-----------------------------------------------
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GDIPlus" (ByVal hDC As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GDIPlus" (ByVal graphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipFillEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByVal mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipSetPenEndCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mEndCap As Long) As Long
Private Declare Function GdipSetPenStartCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mStartCap As Long) As Long
Private Declare Function GdipDrawEllipse Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "GDIPlus" (ByVal argb As Long, ByRef brush As Long) As Long
Private Declare Function GdipCreateFont Lib "GdiPlus.dll" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "GdiPlus.dll" (ByVal mFont As Long) As Long
Private Declare Function GdipDrawString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RectF, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "GDIPlus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "GDIPlus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "GDIPlus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As StringFormatFlags) As Long
Private Declare Function GdipSetStringFormatAlign Lib "GDIPlus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As StringAlignment) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdiplusStartup Lib "GDIPlus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GDIPlus" (ByVal Token As Long)
Private Declare Function GdipDrawLine Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Single, ByVal mY1 As Single, ByVal mX2 As Single, ByVal mY2 As Single) As Long
Private Declare Function GdipGetPathWorldBoundsI Lib "GdiPlus.dll" (ByVal path As Long, ByRef bounds As RECTL, ByVal matrix As Long, ByVal pen As Long) As Long
Private Declare Function GdipCreateMatrix Lib "GDIPlus" (matrix As Long) As Long
Private Declare Function GdipTranslateMatrix Lib "GDIPlus" (ByVal matrix As Long, ByVal offsetX As Single, ByVal offsetY As Single, ByVal order As MatrixOrder) As Long
Private Declare Function GdipRotateMatrix Lib "GDIPlus" (ByVal matrix As Long, ByVal Angle As Single, ByVal order As MatrixOrder) As Long
Private Declare Function GdipTransformPath Lib "GDIPlus" (ByVal path As Long, ByVal matrix As Long) As Long
Private Declare Function GdipAddPathPolygonI Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPoints As POINTL, ByVal mCount As Long) As Long
Private Declare Function GdipAddPathClosedCurve2I Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPoints As POINTL, ByVal mCount As Long, ByVal mTension As Single) As Long

Private Type POINTL
    X As Long
    Y As Long
End Type


Private Enum MatrixOrder
    MatrixOrderPrepend = &H0
    MatrixOrderAppend = &H1
End Enum

Private Type RectF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum
  
Private Enum StringAlignment
    StringAlignmentNear = &H0
    StringAlignmentCenter = &H1
    StringAlignmentFar = &H2
End Enum

Private Enum StringFormatFlags
    StringFormatFlagsNone = &H0
    StringFormatFlagsDirectionRightToLeft = &H1
    StringFormatFlagsDirectionVertical = &H2
    StringFormatFlagsNoFitBlackBox = &H4
    StringFormatFlagsDisplayFormatControl = &H20
    StringFormatFlagsNoFontFallback = &H400
    StringFormatFlagsMeasureTrailingSpaces = &H800
    StringFormatFlagsNoWrap = &H1000
    StringFormatFlagsLineLimit = &H2000
    StringFormatFlagsNoClip = &H4000
End Enum


Private Const LineCapRound              As Long = &H2
Private Const UnitPixel                 As Long = &H2&
Private Const LOGPIXELSX                As Long = 88
Private Const LOGPIXELSY                As Long = 90
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const GDIP_OK                   As Long = &H0

Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Public Event PrePaint(hdc As Long)
'Public Event PostPaint(ByVal hdc As Long)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
'Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

Private Const PI = 3.14159265358979
Private Const PI180 = PI / 180

Dim nScale As Single

Dim m_Min As Single
Dim m_Max As Single
Dim m_Value As Single
Dim m_Angle As Single
Dim m_StartAngle As Single
Dim m_Steps As Long

Dim GdipToken As Long
Dim mPercent As Single
Dim m_MemorPercent As Single
Dim m_PointY As Single


Dim m_ForeColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_TickForeColor As OLE_COLOR
Dim m_TickBackColor As OLE_COLOR
Dim m_LightIntencity As Long
Dim m_TicksSize As Long
Dim m_TicksPenWidth As Long
Dim m_TicksLongFrequency As Long
Dim m_TicksSmallHiden  As Boolean
Dim m_TicksStyleCircle As Boolean
Dim m_RoundStyle As Boolean
Dim cN As ClsNeumorphism

Public Property Get Min() As Single
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Value As Single)
    m_Min = New_Value
    mPercent = (m_Value - m_Min) * 100 / (m_Max - m_Min)
    PropertyChanged "Min"
    Refresh
End Property

Public Property Get Max() As Single
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Value As Single)
    m_Max = New_Value
    mPercent = (m_Value - m_Min) * 100 / (m_Max - m_Min)
    PropertyChanged "Max"
    Refresh
End Property

Public Property Get Value() As Single
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Single)
    m_Value = New_Value
    mPercent = (m_Value - m_Min) * 100 / (m_Max - m_Min)
    PropertyChanged "Value"
    Refresh
    RaiseEvent Change
End Property

Public Property Get Angle() As Single
    Angle = m_Angle
End Property

Public Property Let Angle(ByVal New_Value As Single)
    m_Angle = New_Value
    PropertyChanged "Angle"
    'CleanPen
    Refresh
End Property

Public Property Get StartAngle() As Single
    StartAngle = m_StartAngle
End Property

Public Property Let StartAngle(ByVal New_Value As Single)
    m_StartAngle = New_Value
    PropertyChanged "StartAngle"
    Refresh
End Property

Public Property Let Steps(ByVal New_Value As Long)
    m_Steps = New_Value
    PropertyChanged "Steps"
    Refresh
End Property

Public Property Get Steps() As Long
    Steps = m_Steps
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_Value As OLE_COLOR)
    m_ForeColor = New_Value
    PropertyChanged "ForeColor"
    Refresh
End Property

Public Property Get LightIntencity() As Long
    LightIntencity = m_LightIntencity
End Property

Public Property Let LightIntencity(ByVal New_Value As Long)
    m_LightIntencity = New_Value
    PropertyChanged "LightIntencity"
    Refresh
End Property

Public Property Get TicksSize() As Long
    TicksSize = m_TicksSize
End Property

Public Property Let TicksSize(ByVal New_Value As Long)
    m_TicksSize = New_Value
    PropertyChanged "TicksSize"
    Refresh
End Property

Public Property Get TicksPenWidth() As Long
    TicksPenWidth = m_TicksPenWidth
End Property

Public Property Let TicksPenWidth(ByVal New_Value As Long)
    m_TicksPenWidth = New_Value
    PropertyChanged "TicksPenWidth"
    Refresh
End Property

Public Property Get TicksLongFrequency() As Long
    TicksLongFrequency = m_TicksLongFrequency
End Property

Public Property Let TicksLongFrequency(ByVal New_Value As Long)
    m_TicksLongFrequency = New_Value
    PropertyChanged "TicksLongFrequency"
    Refresh
End Property

Public Property Get TicksSmallHiden() As Boolean
    TicksSmallHiden = m_TicksSmallHiden
End Property

Public Property Let TicksSmallHiden(ByVal New_Value As Boolean)
    m_TicksSmallHiden = New_Value
    PropertyChanged "TicksSmallHiden"
    Refresh
End Property

Public Property Get TicksStyleCircle() As Boolean
    TicksStyleCircle = m_TicksStyleCircle
End Property

Public Property Let TicksStyleCircle(ByVal New_Value As Boolean)
    m_TicksStyleCircle = New_Value
    PropertyChanged "TicksStyleCircle"
    Refresh
End Property

Public Property Get RoundStyle() As Boolean
    RoundStyle = m_RoundStyle
End Property

Public Property Let RoundStyle(ByVal New_Value As Boolean)
    m_RoundStyle = New_Value
    PropertyChanged "RoundStyle"
    Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_Value As OLE_COLOR)
    m_BackColor = New_Value
    PropertyChanged "BackColor"
    Refresh
End Property

Public Property Get TickForeColor() As OLE_COLOR
    TickForeColor = m_TickForeColor
End Property

Public Property Let TickForeColor(ByVal New_Value As OLE_COLOR)
    m_TickForeColor = New_Value
    PropertyChanged "TickForeColor"
    Refresh
End Property

Public Property Get TickBackColor() As OLE_COLOR
    TickBackColor = m_TickBackColor
End Property

Public Property Let TickBackColor(ByVal New_Value As OLE_COLOR)
    m_TickBackColor = New_Value
    PropertyChanged "TickBackColor"
    Refresh
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then m_PointY = Y
    m_MemorPercent = mPercent
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Dist  As Single
    Dim newValue As Single
    
    If Button = 1 Then
        Dist = (m_PointY - Y) / nScale

        If m_Steps = 0 Then
            mPercent = m_MemorPercent + Dist
        Else
            mPercent = m_MemorPercent + ((Dist \ 10) * (100 / m_Steps))
        End If
        
        If mPercent < 0 Then mPercent = 0
        If mPercent > 100 Then mPercent = 100

        Me.Refresh
     
        newValue = m_Value
        m_Value = m_Min + ((m_Max - m_Min) * mPercent / 100)
        
        
        
        If m_Value <> newValue Then
            RaiseEvent Change
        End If
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        m_Min = .ReadProperty("Min", 0)
        m_Max = .ReadProperty("Max", 100)
        m_Value = .ReadProperty("Value", 0)
        m_Angle = .ReadProperty("Angle", 280)
        m_StartAngle = .ReadProperty("StartAngle", -140)
        m_Steps = .ReadProperty("Steps", 0)
        m_ForeColor = .ReadProperty("ForeColor", Ambient.BackColor)
        m_BackColor = .ReadProperty("Backcolor", vbButtonShadow)
        m_LightIntencity = .ReadProperty("LightIntencity", 50)
        m_TicksSize = .ReadProperty("TicksSize", 2)
        m_TicksPenWidth = .ReadProperty("TicksPenWidth", 1)
        m_TicksLongFrequency = .ReadProperty("TicksLongFrequency", 10)
        m_TicksSmallHiden = .ReadProperty("TicksSmallHiden", False)
        m_TicksStyleCircle = .ReadProperty("TicksStyleCircle", False)
        m_RoundStyle = .ReadProperty("RoundStyle", False)
        m_TickForeColor = .ReadProperty("TickForeColor", vbHighlight)
        m_TickBackColor = .ReadProperty("TickBackColor", vbButtonShadow)
        mPercent = (m_Value - m_Min) * 100 / (m_Max - m_Min)
    End With
End Sub

Private Sub UserControl_Terminate()
    Call GdiplusShutdown(GdipToken)
    Set cN = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    
    With PropBag
        .WriteProperty "Min", m_Min, 0
        .WriteProperty "Max", m_Max, 100
        .WriteProperty "Value", m_Value, 0
        .WriteProperty "Angle", m_Angle, 280
        .WriteProperty "StartAngle", m_StartAngle, -140
        .WriteProperty "Steps", m_Steps, 0
        .WriteProperty "ForeColor", m_ForeColor, Ambient.BackColor
        .WriteProperty "BackColor", m_BackColor, vbButtonShadow
        .WriteProperty "LightIntencity", m_LightIntencity, 50
        .WriteProperty "TicksSize", m_TicksSize, 2
        .WriteProperty "TicksPenWidth", m_TicksPenWidth, 1
        .WriteProperty "TicksLongFrequency", m_TicksLongFrequency, 10
        .WriteProperty "TicksSmallHiden", m_TicksSmallHiden, False
        .WriteProperty "TicksStyleCircle", m_TicksStyleCircle, False
        .WriteProperty "RoundStyle", m_RoundStyle, False
        .WriteProperty "TickForeColor", m_TickForeColor, vbHighlight
        .WriteProperty "TickBackColor", m_TickBackColor, vbButtonShadow
    End With
End Sub

Private Sub UserControl_InitProperties()
    m_Min = 0
    m_Max = 100
    m_Value = 0
    m_Angle = 280
    m_StartAngle = -140
    m_ForeColor = Ambient.BackColor
    m_BackColor = vbButtonShadow
    m_LightIntencity = 50
    m_TicksSize = 2
    m_TicksPenWidth = 1
    m_TicksLongFrequency = 10
    m_TicksSmallHiden = False
    m_TicksStyleCircle = False
    m_RoundStyle = False
    m_TickForeColor = vbHighlight
    m_TickBackColor = vbButtonShadow
    m_Steps = 0
End Sub


Private Function DrawText(ByVal hGraphics As Long, ByVal text As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal oFont As StdFont, ByVal ForeColor As Long, Optional HAlign As Long, Optional VAlign As Long, Optional bWordWrap As Boolean) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RectF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long
    Dim hDC As Long

  
    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) <> GDIP_OK Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) <> GDIP_OK Then Exit Function
        'If GdipGetGenericFontFamilySerif(hFontFamily) Then Exit Function
    End If
    
    If GdipCreateStringFormat(0, 0, hFormat) = GDIP_OK Then
        If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        'GdipSetStringFormatFlags hFormat, HotkeyPrefixShow
        GdipSetStringFormatAlign hFormat, HAlign
        GdipSetStringFormatLineAlign hFormat, VAlign
    End If
        
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        

    hDC = GetDC(0&)
    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hDC, LOGPIXELSY), 72)
    ReleaseDC 0&, hDC

    layoutRect.Left = X: layoutRect.Top = Y
    layoutRect.Width = Width: layoutRect.Height = Height

    If GdipCreateSolidFill(ForeColor, hBrush) = GDIP_OK Then
        If GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont) = GDIP_OK Then
            GdipDrawString hGraphics, StrPtr(text), -1, hFont, layoutRect, hFormat, hBrush
            GdipDeleteFont hFont
        End If
        GdipDeleteBrush hBrush
    End If
    
    If hFormat Then GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily


End Function

Public Function GetWindowsDPI() As Double
    Dim hDC As Long, LPX  As Double
    hDC = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hDC, LOGPIXELSX))
    ReleaseDC 0, hDC

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    If UserControl.Enabled Then
        HitResult = vbHitResultHit
    End If
End Sub

Private Sub UserControl_Initialize()
    Dim GdipStartupInput As GdiplusStartupInput
    GdipStartupInput.GdiplusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
    nScale = GetWindowsDPI

    Set cN = New ClsNeumorphism
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
     RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub
'
'Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
'End Sub
'
'Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
'    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
'End Sub

Private Sub UserControl_Paint()
    Dim lScale As Long
    lScale = IIf(UserControl.ScaleWidth < UserControl.ScaleHeight, UserControl.ScaleWidth, UserControl.ScaleHeight)

    Draw UserControl.hDC, lScale, lScale
End Sub

Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Sub UserControl_Show()
    Me.Refresh
End Sub

Public Function RGBtoARGB(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
    'By LaVople
    ' GDI+ color conversion routines. Most GDI+ functions require ARGB format vs standard RGB format
    ' This routine will return the passed RGBcolor to RGBA format
    ' Passing VB system color constants is allowed, i.e., vbButtonFace
    ' Pass Opacity as a value from 0 to 255

    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    RGBtoARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
    
End Function
'*-
Private Sub Draw(hDC As Long, ScaleWidth As Long, ScaleHeight As Long)

    Dim hGraphics   As Long
    Dim hPen As Long, hBrush As Long
    Dim i           As Long
    Dim hPath       As Long
    Dim SL As Single, ST As Single
    Dim S As Single, c As Single
    Dim MidSize As Single, Size As Single
    Dim lPW As Long
    Dim mTotalLines As Long
    Dim lTicksSize As Long
    Dim CircleXY
    Dim WheelSize As Long
    Dim bDrawTick As Boolean
    Dim lDif As Long
    Dim a As Single, P As Single
    
    
    lTicksSize = m_TicksSize * nScale
    mTotalLines = m_Max - m_Min


    lPW = m_TicksPenWidth * nScale
    

    If GdipCreateFromHDC(hDC, hGraphics) = 0 Then

        Call GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)


        MidSize = (lPW / 2)
        Size = (ScaleWidth / 2) - MidSize
        SL = Size + MidSize
        ST = Size + MidSize
        
        MidSize = Size - lTicksSize
        CircleXY = lTicksSize * 3 + 8 * nScale
        WheelSize = ScaleHeight - CircleXY * 2
        
   
        GdipCreateSolidFill RGBtoARGB(m_BackColor, 100), hBrush
        GdipFillEllipseI hGraphics, hBrush, CircleXY, CircleXY, ScaleWidth - CircleXY * 2, ScaleHeight - CircleXY * 2
        GdipDeleteBrush hBrush
        
        GdipCreatePen1 RGBtoARGB(m_TickBackColor, 100), lPW, UnitPixel, hPen
        GdipDrawEllipse hGraphics, hPen, CircleXY, CircleXY, ScaleWidth - CircleXY * 2, ScaleHeight - CircleXY * 2
        GdipDeletePen hPen

        For i = 0 To mTotalLines '- 1
            P = i * 100 / mTotalLines
            a = m_StartAngle + (m_Angle * P / 100)
            
            S = Sin(a * PI180)
            c = Cos(a * PI180)

            If i Mod m_TicksLongFrequency = 0 Then
                bDrawTick = True
                MidSize = Size - lTicksSize * 3
                lDif = (MidSize - Size) / 2
            Else
                If m_TicksSmallHiden Then bDrawTick = False
                MidSize = Size - lTicksSize
                lDif = 0
            End If

            If bDrawTick Then
                If m_TicksStyleCircle Then
                    If mPercent >= P Or mPercent >= 99 Then
                        GdipCreateSolidFill RGBtoARGB(m_TickForeColor, 100), hBrush
                    Else
                        GdipCreateSolidFill RGBtoARGB(m_TickBackColor, 100), hBrush
                    End If
                    GdipFillEllipseI hGraphics, hBrush, SL + (S * (MidSize - lDif)) - (MidSize - Size) / 2, ST - (c * (MidSize - lDif)) - (MidSize - Size) / 2, MidSize - Size, MidSize - Size
                    GdipDeleteBrush hBrush
                Else
            
                    If mPercent >= P Or mPercent >= 99 Then
                        GdipCreatePen1 RGBtoARGB(m_TickForeColor, 100), lPW, UnitPixel, hPen
                    Else
                        GdipCreatePen1 RGBtoARGB(m_TickBackColor, 100), lPW, UnitPixel, hPen
                    End If
                    GdipSetPenStartCap hPen, LineCapRound
                    GdipSetPenEndCap hPen, LineCapRound
                    Call GdipDrawLine(hGraphics, hPen, SL + (S * MidSize), ST - (c * MidSize), S * Size + SL, -c * Size + ST)
                    GdipDeletePen hPen
                End If
            End If
        Next i

        a = m_StartAngle + (m_Angle * mPercent / 100)
        
        
        S = Sin(a * PI180)
        c = Cos(a * PI180)
        
        With cN
            .CleanUp
            .Blur = 4
            .Distance = 2
            .LightDirection = TopLeft
            .Gradient = True
            .GradientFlip = True
            .StatePressed = False
            .BackColor = m_ForeColor
            .Intencity = LightIntencity
            .Radius = 1000
        End With
        
        If m_RoundStyle Then
            cN.Draw hDC, CircleXY + lPW, CircleXY + lPW, WheelSize - lPW * 2, WheelSize - lPW * 2, 0, hGraphics
        Else
            hPath = CreateToolPath(WheelSize, WheelSize, 8, 40, 140, a)
            cN.Draw hDC, CircleXY, CircleXY, WheelSize, WheelSize, hPath, hGraphics
            GdipDeletePath hPath
        End If
        
        With cN
            .CleanUp
            .Blur = 3
            .Distance = 1
            .Gradient = True
            .StatePressed = True
            .LightDirection = TopLeft
            .Gradient = True
            .GradientFlip = True
            .Radius = 1000
        End With

        cN.Draw hDC, ScaleWidth / 2 - WheelSize / 3.4, ScaleWidth / 2 - WheelSize / 3.4, WheelSize / 1.7, WheelSize / 1.7, 0, hGraphics

        If m_TicksStyleCircle Then
            Size = WheelSize / 6
            
            If m_RoundStyle Then
                MidSize = WheelSize / 2.6
            Else
                MidSize = WheelSize / 3.8
            End If
            
            With cN
                .CleanUp
                .Blur = 1
                .Distance = 2
                '.Intencity = 40
                .Gradient = True
                .StatePressed = True
                .LightDirection = TopLeft
                .Gradient = True
                .GradientFlip = True
                .Radius = 1000
                .BackColor = m_TickForeColor
            End With
    
            cN.Draw hDC, SL + (S * MidSize) - (Size / 2), ST - (c * MidSize) - (Size / 2), Size, Size, 0, hGraphics

'            GdipCreateSolidFill RGBtoARGB(m_TickForeColor, 100), hBrush
'            GdipFillEllipseI hGraphics, hBrush, SL + (S * (MidSize - lDif)) - (MidSize - Size) / 2, ST - (C * (MidSize - lDif)) - (MidSize - Size) / 2, MidSize - Size, MidSize - Size
'            GdipDeleteBrush hBrush
        Else
            Size = 5
            MidSize = WheelSize / 1.7 / 2
        
            GdipCreatePen1 RGBtoARGB(m_TickForeColor, 100), lPW * 2, UnitPixel, hPen
            GdipSetPenStartCap hPen, LineCapRound
            GdipSetPenEndCap hPen, LineCapRound

            Call GdipDrawLine(hGraphics, hPen, SL + (S * MidSize), ST - (c * MidSize), S * Size + SL, -c * Size + ST)
            GdipDeletePen hPen
        End If
        
        GdipDeleteGraphics hGraphics
    End If
    

End Sub


'COPY FROM EDUARDO SHAPES CONTROL
Private Function CreateToolPath(Width As Long, Height As Long, mVertices As Long, mShift As Long, mCurvingFactor As Single, mAngle As Single) As Long

    Dim iHeight As Long
   
    Dim iPts() As POINTL

    Dim c As Long
    Dim iPts2() As POINTL
    Dim iPts3() As POINTL
    Dim iShift As Long
    Dim hPath As Long


    If Width < Height Then
        iHeight = Width
    Else
        iHeight = Height
    End If
    
    Call GdipCreatePath(&H0, hPath)

    ReDim iPts(mVertices * 2 - 1)

    For c = 0 To mVertices * 2 - 1
        iPts(c).X = (iHeight / 2) * Cos(2 * PI * (c + 1) / (mVertices * 2)) + Width / 2
        iPts(c).Y = (iHeight / 2) * Sin(2 * PI * (c + 1) / (mVertices * 2)) + Height / 2
    Next c

    ReDim iPts2(mVertices - 1)
    iShift = (iHeight / 100 * mShift / 3) '+ 10

    For c = 0 To mVertices - 1
        iPts2(c).X = (iHeight / 2 - iShift) * Cos(2 * PI * (c + 1) / mVertices) + Width / 2
        iPts2(c).Y = (iHeight / 2 - iShift) * Sin(2 * PI * (c + 1) / mVertices) + Height / 2
    Next c

    ReDim iPts3(mVertices * 2 - 1)
    For c = 0 To mVertices * 2 - 1
        If c Mod 2 = 0 Then
            iPts3(c).X = iPts2(c / 2).X
            iPts3(c).Y = iPts2(c / 2).Y
        Else
            iPts3(c).X = iPts((c + 1) Mod (UBound(iPts) + 1)).X
            iPts3(c).Y = iPts((c + 1) Mod (UBound(iPts) + 1)).Y
        End If
    Next c

    AddPolygon hPath, iPts3, mCurvingFactor

    'Thanks SomeYguy!!!
    If mAngle <> 0 Then
        Dim tMatrix   As Long
        Dim tRect As RECTL
        Dim cX As Single, cY As Single
        GdipGetPathWorldBoundsI hPath, tRect, 0&, 0
        cX = Width / 2
        cY = Height / 2
        GdipCreateMatrix tMatrix
        GdipTranslateMatrix tMatrix, -cX, -cY, MatrixOrderAppend
        GdipRotateMatrix tMatrix, mAngle, MatrixOrderAppend
        GdipTranslateMatrix tMatrix, cX, cY, MatrixOrderAppend
        GdipTransformPath hPath, tMatrix
    End If
    Call GdipGetPathWorldBoundsI(hPath, tRect, 0&, 0&)
    'Width = tRect.Width
    'Height = tRect.Height
    CreateToolPath = hPath
End Function


Private Sub AddPolygon(hPath As Long, Points() As POINTL, mCurvingFactor As Single)
    Dim mCurvingFactor2 As Single
    If mCurvingFactor = 0 Then
        GdipAddPathPolygonI hPath, Points(0), UBound(Points) + 1
    Else
        If mCurvingFactor < 0 Then
            mCurvingFactor2 = mCurvingFactor / 100 * 0.5
        Else
            mCurvingFactor2 = mCurvingFactor / 100 * 1
        End If
        GdipAddPathClosedCurve2I hPath, Points(0), UBound(Points) + 1, mCurvingFactor2
    End If
End Sub

