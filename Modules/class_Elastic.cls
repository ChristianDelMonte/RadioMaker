VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "class_Elastic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : Mikhail Shmukler
' * Web Site         : www.geocities.com/ResearchTriangle/6311/
' * E-Mail           : waty.thierry@usa.net
' * Date             : 13/10/98
' * Time             : 10:24
' * Module Name      : class_Elastic
' * Module Filename  : Elastic.cls
' **********************************************************************
' * Comments         :
' * This class can change size and location of controls On your form
' * 1. Resize form
' * 2. Change screen resolution
' * Assumes:1. Add Elastic.cls
' *         2. Add declaration 'Private El as New class_Elastic'
' *         3. Insert string like 'El.init Me' (formload event)
' *         4. Insert string like 'El.FormResize Me' (Resize event)
' *         5. Press 'F5' and resize form ....

'
'NOTE FROM FREEVBCODE.COM'S TESTING
'This works best if you:
'1. Declare a form_level variable such as m_IsLoading as boolean
'2. in Form_Load, enter code such as the following:
  'm_IsLoading = true
  'El.Init Me
   'm_IsLoading = False

'3. In Form Resize, use:
    'if not m_IsLoading then e1.FormResize.Me
'****************************************************************

Option Explicit
Private nFormHeight      As Integer
Private nFormWidth       As Integer
Private nNumOfControls   As Integer
Private nTop()           As Integer
Private nLeft()          As Integer
Private nHeight()        As Integer
Private nWidth()         As Integer
Private nFontSize()      As Integer
Private nRightMargin()   As Integer
Private bFirstTime       As Boolean

Public Sub Init(frm As Form, Optional nWindState As Variant)
   
   Dim I          As Integer
   Dim bWinMax    As Boolean
   
   bWinMax = Not IsMissing(nWindState)
   
   nFormHeight = frm.Height
   nFormWidth = frm.Width
   nNumOfControls = frm.Controls.Count - 1
   bFirstTime = True
   ReDim nTop(nNumOfControls)
   ReDim nLeft(nNumOfControls)
   ReDim nHeight(nNumOfControls)
   ReDim nWidth(nNumOfControls)
   ReDim nFontSize(nNumOfControls)
   
   ReDim nRightMargin(nNumOfControls)
   On Error Resume Next
   For I = 0 To nNumOfControls
      If TypeOf frm.Controls(I) Is Line Then
         nTop(I) = frm.Controls(I).Y1
         nLeft(I) = frm.Controls(I).x1
         nHeight(I) = frm.Controls(I).Y2
         nWidth(I) = frm.Controls(I).X2
      Else
         nTop(I) = frm.Controls(I).Top
         nLeft(I) = frm.Controls(I).Left
         nHeight(I) = frm.Controls(I).Height
         nWidth(I) = frm.Controls(I).Width
         nFontSize(I) = frm.FontSize
         nRightMargin(I) = frm.Controls(I).RightMargin
      End If
   Next
   
   If bWinMax Or frm.WindowState = 2 Then ' maxim
      frm.Height = Screen.Height
      frm.Width = Screen.Width
   Else
      frm.Height = frm.Height * Screen.Height / 7290
      frm.Width = frm.Width * Screen.Width / 9690
   End If
   
   bFirstTime = True
   
End Sub

Public Sub FormResize(frm As Form)
   
   Dim I             As Integer
   Dim nCaptionSize  As Integer
   Dim dRatioX       As Double
   Dim dRatioY       As Double
   Dim nSaveRedraw   As Long
   
   On Error Resume Next
   nSaveRedraw = frm.AutoRedraw
   
   frm.AutoRedraw = True
   
   If bFirstTime Then
      bFirstTime = False
      Exit Sub
   End If
   
   If frm.Height < nFormHeight / 2 Then frm.Height = nFormHeight / 2
   
   If frm.Width < nFormWidth / 2 Then frm.Width = nFormWidth / 2
   nCaptionSize = 400
   dRatioY = 1# * (nFormHeight - nCaptionSize) / (frm.Height - nCaptionSize)
   dRatioX = 1# * nFormWidth / frm.Width
   On Error Resume Next ' for comboboxes, timeres and other nonsizible controls
   
   For I = 0 To nNumOfControls
      If TypeOf frm.Controls(I) Is Line Then
         frm.Controls(I).Y1 = Int(nTop(I) / dRatioY)
         frm.Controls(I).x1 = Int(nLeft(I) / dRatioX)
         frm.Controls(I).Y2 = Int(nHeight(I) / dRatioY)
         frm.Controls(I).X2 = Int(nWidth(I) / dRatioX)
      Else
         frm.Controls(I).Top = Int(nTop(I) / dRatioY)
         frm.Controls(I).Left = Int(nLeft(I) / dRatioX)
         frm.Controls(I).Height = Int(nHeight(I) / dRatioY)
         frm.Controls(I).Width = Int(nWidth(I) / dRatioX)
         'frm.Controls(I).FontSize = Int(nFontSize(I) / dRatioX) + Int(nFontSize(I) / dRatioX) Mod 2
         frm.Controls(I).RightMargin = Int(nRightMargin(I) / dRatioY)
      End If
   Next
   
   frm.AutoRedraw = nSaveRedraw
   
End Sub


