VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsNeumorphism"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------
'Autor:  Leandro Ascierto
'Web: www.leandroascierto.com
'Date: 17/01/2021
Private Declare Sub GdiplusShutdown Lib "GDIPlus" (ByVal Token As Long)
Private Declare Function GdiplusStartup Lib "GDIPlus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GDIPlus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GDIPlus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GDIPlus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "GDIPlus" (ByVal argb As Long, ByRef brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GDIPlus" (ByVal brush As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
'Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathRectangleI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreateTexture Lib "GdiPlus.dll" (ByVal mImage As Long, ByVal mWrapMode As Long, ByRef mTexture As Long) As Long
'Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapLockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mRect As RECTL, ByVal mFlags As ImageLockMode, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "GDIPlus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "GDIPlus" (ByVal graphics As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipSetClipPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPath As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipTranslateClipI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mDx As Long, ByVal mDy As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipGetPathWorldBoundsI Lib "GdiPlus.dll" (ByVal path As Long, ByRef bounds As RECTL, ByVal matrix As Long, ByVal pen As Long) As Long
Private Declare Function GdipTranslateTextureTransform Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mDx As Single, ByVal mDy As Single, ByVal mOrder As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
'Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipRestoreGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mState As Long) As Long
Private Declare Function GdipSaveGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByRef mState As Long) As Long

Private Const UnitPixel                 As Long = &H2&
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const CombineModeExclude        As Long = &H4
Private Const WrapModeTileFlipXY        As Long = &H3

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type BitmapData
    Width                       As Long
    Height                      As Long
    stride                      As Long
    PixelFormat                 As Long
    Scan0Ptr                    As Long
    ReservedPtr                 As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type


Private Enum ImageLockMode
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum

Public Enum eLightDirection
    TopLeft = 0
    TopRight = 1
    BottomRight = 2
    BottomLeft = 3
End Enum

Dim GdipToken As Long
Dim m_Distance As Long
Dim m_Radius As Long
Dim m_Intencity As Long
Dim m_Blur As Long
Dim m_BackColor As OLE_COLOR
Dim m_BrushBtnNormal As Long
Dim m_BrushBtnPressed As Long
Dim m_LightDirection As eLightDirection
Dim m_StatePressed As Boolean
Dim m_Gradient As Boolean
Dim m_GradientFlip As Boolean
Dim m_ShadowColor As OLE_COLOR
Dim m_LightColor As OLE_COLOR

Public Property Get Distance() As Long
    Distance = m_Distance
End Property

Public Property Let Distance(ByVal New_Value As Long)
    m_Distance = New_Value
    CleanUp
End Property

Public Property Get Radius() As Long
    Radius = m_Radius
End Property

Public Property Let Radius(ByVal New_Value As Long)
    m_Radius = New_Value
    CleanUp
End Property

Public Property Get Intencity() As Long
    Intencity = m_Intencity
End Property

Public Property Let Intencity(ByVal New_Value As Long)
    m_Intencity = New_Value
    CleanUp
End Property

Public Property Get Blur() As Long
    Blur = m_Blur
End Property

Public Property Let Blur(ByVal New_Value As Long)
    m_Blur = New_Value
    CleanUp
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_Value As OLE_COLOR)
    m_BackColor = New_Value
    CleanUp
End Property

Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(ByVal New_Value As OLE_COLOR)
    m_ShadowColor = New_Value
    CleanUp
End Property

Public Property Get LightColor() As OLE_COLOR
    LightColor = m_LightColor
End Property

Public Property Let LightColor(ByVal New_Value As OLE_COLOR)
    m_LightColor = New_Value
    CleanUp
End Property
Public Property Get LightDirection() As eLightDirection
    LightDirection = m_LightDirection
End Property

Public Property Let LightDirection(ByVal New_Value As eLightDirection)
    m_LightDirection = New_Value
    CleanUp
End Property

Public Property Get StatePressed() As Boolean
    StatePressed = m_StatePressed
End Property

Public Property Let StatePressed(ByVal New_Value As Boolean)
    m_StatePressed = New_Value
End Property

Public Property Get Gradient() As Boolean
    Gradient = m_Gradient
End Property

Public Property Let Gradient(ByVal New_Value As Boolean)
    m_Gradient = New_Value
    CleanUp
End Property

Public Property Get GradientFlip() As Boolean
    GradientFlip = m_GradientFlip
End Property

Public Property Let GradientFlip(ByVal New_Value As Boolean)
    m_GradientFlip = New_Value
    CleanUp
End Property

Public Sub CleanUp()
    If m_BrushBtnNormal Then GdipDeleteBrush m_BrushBtnNormal: m_BrushBtnNormal = 0&
    If m_BrushBtnPressed Then GdipDeleteBrush m_BrushBtnPressed: m_BrushBtnPressed = 0&
End Sub

'*1
Public Sub Draw(ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional CustomPath As Long, Optional SrcGraphic As Long)
    Dim hGraphics As Long, hGraphics2 As Long
    Dim hPath As Long, hBrush As Long, hPen As Long
    Dim hImage As Long
    Dim IsDark As Boolean
    Dim DB2 As Long, B2 As Long
    Dim GradientAngle As Long
    Dim x As Long, y As Long
    Dim LC As Long
    Dim RECT As RECTL
    Dim pRect As RECTL
    Dim Color1 As Long, Color2 As Long
    Dim D As Long
    Dim hState As Long
    
    If SrcGraphic = 0 Then
        GdipCreateFromHDC hDC, hGraphics
        GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    Else
        hGraphics = SrcGraphic
        GdipSaveGraphics hGraphics, hState
    End If

    
    
    B2 = m_Blur * 2 + 1
    DB2 = (B2 + m_Distance) * 2
    
    If CustomPath Then
        hPath = CustomPath
        GdipGetPathWorldBoundsI hPath, pRect, 0&, 0&
        D = (Abs(pRect.Left) + Abs(pRect.Top)) / 2
        DB2 = DB2 + D * 3
    Else
        hPath = CreateRoundPath(0, 0, Width, Height, m_Radius)
    End If
    
    If m_BrushBtnNormal And (m_StatePressed = False) Then
        GdipTranslateWorldTransform hGraphics, Left - m_Distance - B2 - D, Top - m_Distance - B2 - D, 0&
        GdipFillRectangleI hGraphics, m_BrushBtnNormal, 0, 0, Width + DB2, Height + DB2
        GdipDeleteGraphics hGraphics
        Exit Sub
    End If
    
    If m_BrushBtnPressed And (m_StatePressed = True) Then
        GdipTranslateWorldTransform hGraphics, Left, Top, 0&
        GdipFillPath hGraphics, m_BrushBtnPressed, hPath
        GdipDeletePath hPath
        GdipDeleteGraphics hGraphics
        Exit Sub
    End If
    
    IsDark = IsDarkColor(m_BackColor)
    LC = LuminanceColor(m_BackColor)
    
    If m_StatePressed = False Then
        GdipCreateBitmapFromScan0 Width + DB2, Height + DB2, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
        GdipGetImageGraphicsContext hImage, hGraphics2
        GdipSetSmoothingMode hGraphics2, SmoothingModeAntiAlias
    
        x = B2 + D: y = B2 + D
        Select Case m_LightDirection
            Case TopRight: x = x + m_Distance * 2
            Case BottomRight: x = x + m_Distance * 2: y = x
            Case BottomLeft: y = y + m_Distance * 2
        End Select
        
        GdipTranslateWorldTransform hGraphics2, x, y, 0&
        GdipCreateSolidFill RGBtoARGB(m_LightColor, m_Intencity), hBrush
        GdipFillPath hGraphics2, hBrush, hPath
        GdipDeleteBrush hBrush

        x = B2 + D: y = B2 + D
        Select Case m_LightDirection
            Case TopLeft: x = x + m_Distance * 2: y = x
            Case TopRight: y = y + m_Distance * 2
            Case BottomLeft: x = x + m_Distance * 2
        End Select
        
        GdipResetWorldTransform hGraphics2
        GdipTranslateWorldTransform hGraphics2, x, y, 0&
        GdipCreateSolidFill RGBtoARGB(m_ShadowColor, m_Intencity), hBrush
        GdipFillPath hGraphics2, hBrush, hPath
        GdipDeleteBrush hBrush
    
        BlurImage hImage, m_Blur
    
        GdipResetWorldTransform hGraphics2
        x = m_Distance + B2 + D: y = m_Distance + B2 + D
        GdipTranslateWorldTransform hGraphics2, x, y, 0&
        
        If m_Gradient Then
            RECT.Width = Width
            RECT.Height = Height
            GradientAngle = 45 + 90 * m_LightDirection - 180 * m_GradientFlip
            Color1 = RGBtoARGB(ShiftColor(m_ShadowColor, m_BackColor, m_Intencity), 100)
            Color2 = RGBtoARGB(ShiftColor(m_LightColor, m_BackColor, m_Intencity), 100)
            GdipCreateLineBrushFromRectWithAngleI RECT, Color1, Color2, GradientAngle, 0, WrapModeTileFlipXY, hBrush
        Else
            GdipCreateSolidFill RGBtoARGB(m_BackColor, 100), hBrush
        End If
        
        GdipFillPath hGraphics2, hBrush, hPath
        GdipDeleteBrush hBrush
        
        GdipCreateTexture hImage, &H0, hBrush
        GdipTranslateWorldTransform hGraphics, Left - x, Top - y, 0&
        GdipFillRectangleI hGraphics, hBrush, 0, 0, Width + DB2, Height + DB2
        'GdipDeleteBrush hBrush
        m_BrushBtnNormal = hBrush
    Else
        GdipCreateBitmapFromScan0 Width + m_Blur * 2, Height + m_Blur * 2, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
        GdipGetImageGraphicsContext hImage, hGraphics2
        GdipSetSmoothingMode hGraphics2, SmoothingModeAntiAlias
            
        If m_Gradient Then
            RECT.Width = Width
            RECT.Height = Height
            GradientAngle = 45 + 90 * m_LightDirection - 180 * m_GradientFlip
            Color1 = RGBtoARGB(ShiftColor(m_ShadowColor, m_BackColor, m_Intencity), 100)
            Color2 = RGBtoARGB(ShiftColor(m_LightColor, m_BackColor, m_Intencity), 100)
            GdipCreateLineBrushFromRectWithAngleI RECT, Color1, Color2, GradientAngle, 0, WrapModeTileFlipXY, hBrush
        Else
            GdipCreateSolidFill RGBtoARGB(m_BackColor, 100), hBrush
        End If

        GdipTranslateWorldTransform hGraphics2, m_Blur + -pRect.Left, m_Blur + -pRect.Top, 0&
        GdipFillPath hGraphics2, hBrush, hPath
        GdipDeleteBrush hBrush

        x = -m_Distance: y = -m_Distance
        Select Case m_LightDirection
            Case TopRight: x = x + m_Distance * 2
            Case BottomRight: x = x + m_Distance * 2: y = x
            Case BottomLeft: y = y + m_Distance * 2
        End Select

        GdipSetClipPath hGraphics2, hPath, CombineModeExclude
        GdipTranslateClipI hGraphics2, x, y
        GdipCreateSolidFill RGBtoARGB(m_LightColor, m_Intencity), hBrush 'IIf(IsDark, m_Intencity, 50)
        GdipFillPath hGraphics2, hBrush, hPath
        GdipDeleteBrush hBrush

        GdipTranslateClipI hGraphics2, (x * -1) * 2, (y * -1) * 2
        GdipCreateSolidFill RGBtoARGB(m_ShadowColor, m_Intencity), hBrush
        GdipFillPath hGraphics2, hBrush, hPath
        GdipDeleteBrush hBrush

        If IsDark = False Then
            GdipResetClip hGraphics2
            GdipCreatePen1 RGBtoARGB(m_ShadowColor, m_Intencity / 2), m_Blur / 4, UnitPixel, hPen
            GdipDrawPath hGraphics2, hPen, hPath
            GdipDeletePen hPen
        End If

        BlurImage hImage, m_Blur
        GdipCreateTexture hImage, &H0, hBrush
        GdipTranslateTextureTransform hBrush, pRect.Left - m_Blur, pRect.Top - m_Blur, 0
        GdipTranslateWorldTransform hGraphics, Left, Top, 0&
        GdipFillPath hGraphics, hBrush, hPath
'        'GdipDeleteBrush hBrush
        m_BrushBtnPressed = hBrush
    End If
    
    If CustomPath = 0 Then GdipDeletePath hPath
    GdipDisposeImage hImage
    GdipDeleteGraphics hGraphics2
    
    If SrcGraphic = 0 Then
        GdipDeleteGraphics hGraphics
    Else
        GdipRestoreGraphics hGraphics, hState
    End If
End Sub

Private Function CreateRoundPath(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal Radius As Single) As Long
    Dim hPath As Long
    If GdipCreatePath(&H0, hPath) = 0& Then
    
        If Radius > Width / 2 Then Radius = Width / 2
        If Radius > Height / 2 Then Radius = Height / 2
    
        If Radius = 0 Then
            GdipAddPathRectangleI hPath, Left, Top, Width, Height
        Else
            Radius = Radius * 2
            GdipAddPathArcI hPath, Left, Top, Radius, Radius, 180, 90
            GdipAddPathArcI hPath, Left + Width - Radius, Top, Radius, Radius, 270, 90
            GdipAddPathArcI hPath, Left + Width - Radius, Top + Height - Radius, Radius, Radius, 0, 90
            GdipAddPathArcI hPath, Left, Top + Height - Radius, Radius, Radius, 90, 90
            GdipClosePathFigure hPath
        End If
        CreateRoundPath = hPath
    End If
    
End Function


Private Function BlurImage(ByVal hImage As Long, BlurDepth As Long, _
                        Optional ByVal Left As Long, Optional ByVal Top As Long, _
                        Optional ByVal Width As Long, Optional ByVal Height As Long) As Boolean
                                        
    Dim RECT As RECTL
    Dim bmpData1 As BitmapData
    Dim srcBytes() As Byte
    Dim kDiv As Long
    Dim MaxX As Long, MaxY As Long, Width4 As Long
    Dim x As Long, y As Long
    Dim X0 As Long, X1 As Long, X2 As Long
    Dim Y0 As Long, Y1 As Long, Y2 As Long
    Dim a As Long, r As Long, g As Long, b As Long
    Dim dX0 As Long, dX2 As Long, dY0 As Long, dY2 As Long
    Dim mOut() As Byte, n As Long

    If BlurDepth <= 0& Then Exit Function
    If hImage = 0& Then Exit Function
    If Width = 0& Then Call GdipGetImageWidth(hImage, Width)
    If Height = 0& Then Call GdipGetImageHeight(hImage, Height)
 
    With RECT
        .Left = Left
        .Top = Top
        .Width = Width
        .Height = Height
    End With

    ReDim srcBytes(RECT.Width * RECT.Height * 4 - 1&)

    With bmpData1
        .Scan0Ptr = VarPtr(srcBytes(0&))
        .stride = 4& * RECT.Width
    End With
   
    If GdipBitmapLockBits(hImage, RECT, ImageLockModeUserInputBuf Or ImageLockModeRead Or ImageLockModeWrite, PixelFormat32bppPARGB, bmpData1) = 0& Then

        MaxX = Width - 1
        MaxY = Height - 1
        Width4 = Width * 4
        ReDim mOut(Width4 * Height - 1)
        kDiv = BlurDepth * 2 + 1
    
        For n = 0 To 1
          For y = 0 To MaxY
            b = 0
            g = 0
            r = 0
            a = 0
            X0 = y * Width4
            X1 = X0
            For x = 2 To BlurDepth
              X0 = X0 + 4
              b = b + srcBytes(X0 + 0)
              g = g + srcBytes(X0 + 1)
              r = r + srcBytes(X0 + 2)
              a = a + srcBytes(X0 + 3)
            Next x
            X0 = X1 + BlurDepth * 4
            X2 = X0
            b = b + b + srcBytes(X1 + 0) + srcBytes(X0 + 0)
            g = g + g + srcBytes(X1 + 1) + srcBytes(X0 + 1)
            r = r + r + srcBytes(X1 + 2) + srcBytes(X0 + 2)
            a = a + a + srcBytes(X1 + 3) + srcBytes(X0 + 3)
            dX0 = -4
            dX2 = 4
            For x = 0 To MaxX
              b = b + srcBytes(X2 + 0)
              g = g + srcBytes(X2 + 1)
              r = r + srcBytes(X2 + 2)
              a = a + srcBytes(X2 + 3)
              mOut(X1 + 0) = b \ kDiv
              mOut(X1 + 1) = g \ kDiv
              mOut(X1 + 2) = r \ kDiv
              mOut(X1 + 3) = a \ kDiv
              b = b - srcBytes(X0 + 0)
              g = g - srcBytes(X0 + 1)
              r = r - srcBytes(X0 + 2)
              a = a - srcBytes(X0 + 3)
              If x = BlurDepth Then dX0 = 4
              X0 = X0 + dX0
              X1 = X1 + 4
              If x = MaxX - BlurDepth Then dX2 = -4
              X2 = X2 + dX2
            Next x
          Next y
          
          For x = 0 To MaxX
            b = 0
            g = 0
            r = 0
            a = 0
            Y0 = x * 4
            Y1 = Y0
            For y = 2 To BlurDepth
              Y0 = Y0 + Width4
              b = b + mOut(Y0 + 0)
              g = g + mOut(Y0 + 1)
              r = r + mOut(Y0 + 2)
              a = a + mOut(Y0 + 3)
            Next y
            Y0 = Y1 + BlurDepth * Width4
            Y2 = Y0
            b = b + b + mOut(Y1 + 0) + mOut(Y0 + 0)
            g = g + g + mOut(Y1 + 1) + mOut(Y0 + 1)
            r = r + r + mOut(Y1 + 2) + mOut(Y0 + 2)
            a = a + a + mOut(Y1 + 3) + mOut(Y0 + 3)
            dY0 = -Width4
            dY2 = Width4
            For y = 0 To MaxY
              b = b + mOut(Y2 + 0)
              g = g + mOut(Y2 + 1)
              r = r + mOut(Y2 + 2)
              a = a + mOut(Y2 + 3)
              srcBytes(Y1 + 0) = b \ kDiv
              srcBytes(Y1 + 1) = g \ kDiv
              srcBytes(Y1 + 2) = r \ kDiv
              srcBytes(Y1 + 3) = a \ kDiv
              b = b - mOut(Y0 + 0)
              g = g - mOut(Y0 + 1)
              r = r - mOut(Y0 + 2)
              a = a - mOut(Y0 + 3)
              If y = BlurDepth Then dY0 = Width4
              Y0 = Y0 + dY0
              Y1 = Y1 + Width4
              If y = MaxY - BlurDepth Then dY2 = -Width4
              Y2 = Y2 + dY2
            Next y
          Next x
    
        Next n
      
        BlurImage = GdipBitmapUnlockBits(hImage, bmpData1) = 0&
    End If
End Function

Public Function LuminanceColor(ByVal Color As Long) As Long
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
    CopyMemory BGRA(0), Color, 4&
    LuminanceColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) * 100 / 638
    'LuminanceColor = ((CLng(BGRA(0)) + (CLng(BGRA(1))) + CLng(BGRA(2))) / 2) * 100 / 382
End Function

Public Function IsDarkColor(ByVal Color As Long) As Boolean
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
    CopyMemory BGRA(0), Color, 4&
    IsDarkColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) < 382
End Function

Public Function RGBtoARGB(ByVal RGBColor As Long, Optional ByVal Opacity As Long = 100) As Long
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

'Funcion para combinar dos colores
Public Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
  
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
  
    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
  
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
  
    CopyMemory ShiftColor, clrFore(0), 4
  
End Function


'Inicia GDI+
Private Sub InitGDI()
    Dim GdipStartupInput As GdiplusStartupInput
    GdipStartupInput.GdiplusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub
  
'Termina GDI+
Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub

Private Sub Class_Initialize()
    InitGDI
    m_Distance = 10
    m_Radius = 10
    m_Intencity = 40
    m_Blur = 15
    m_ShadowColor = vbBlack
    m_LightColor = vbWhite
    m_BackColor = vbButtonFace
End Sub

Private Sub Class_Terminate()
    CleanUp
    TerminateGDI
End Sub
