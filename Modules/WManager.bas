Attribute VB_Name = "WindowManager"
'********************* RM100 *********************
'     RADIO MAKER WINDOW CONTROLLER MODULE
'COPYRIGHT (C) 1987-2024 ONLY development inc.
'Christian A. Del Monte
'*************************************************
' ultima modificacion: 18-02-24
'*************************************************

Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'constantes KEEPOnTop
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_MAXIMIZE = 3

        '********** REFERENCIAS *****************
        'Contantes de ventanas en 1024 x 768
            'Private Const WidthMax = 15360
            'Private Const HeightMax = 11520
        
        'Contantes de ventanas en 1600 x 900
            'Private Const WidthMax = 24000
            'Private Const HeightMax = 13500
        
        'Contantes de ventanas en 1920 x 1080
            'Private Const WidthMax = 28800
            'Private Const HeightMax = 16200
        '********** REFERENCIAS *****************

'Contantes de ventanas en 1600x900
Private Const WidthMax = 24000  ' 28800 en 1920
Private Const HeightMax = 13500  ' 16200 en 1080

Public Sub KeepOnTop(ByVal frmIn As Form)

'Keep form on top. Note that this is switched off if form is
'minimised, so place in resize event as well.
Const wFlags = SWP_NOMOVE Or SWP_NOSIZE

    SetWindowPos frmIn.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, wFlags    'Window will stay on top
    DoEvents

End Sub

Public Sub HideWindow(WWindow As String)

Select Case WWindow
    Case "TopMenu"  '--------------------------------------------------
        TopMenu.WindowState = 1
    
    Case "DwMenu"   '--------------------------------------------------
        DownMenu.WindowState = 1
        DownMenu.Visible = False
        TopMenu.SbHerram.Checked = False
    
    Case "Est01"    '--------------------------------------------------
        Est01.WindowState = 1
        Est01.Visible = False
        TopMenu.SbEst01.Checked = False
    
    Case "Est02"    '--------------------------------------------------
        Est02.WindowState = 1
        Est02.Visible = False
        TopMenu.SbEst02.Checked = False
    
    Case "Tnd01"    '--------------------------------------------------
        Tanda01.WindowState = 1
        Tanda01.Visible = False
        TopMenu.SbTnd01.Checked = False
    
    Case "Prg01"    '--------------------------------------------------
        Prg01.WindowState = 1
        Prg01.Visible = False
        TopMenu.SbPrg01.Checked = False
    
    Case "Explor01" '--------------------------------------------------
        Unload XPlorer
        TopMenu.SbExplor.Checked = False
            
    Case "All"  '--------------------------------------------------
        DownMenu.WindowState = 1
        DownMenu.Visible = False
        TopMenu.SbHerram.Checked = False
        Est01.WindowState = 1
        Est01.Visible = False
        TopMenu.SbEst01.Checked = False
        Est02.WindowState = 1
        Est02.Visible = False
        TopMenu.SbEst02.Checked = False
        Tanda01.WindowState = 1
        Tanda01.Visible = False
        TopMenu.SbTnd01.Checked = False
        Prg01.WindowState = 1
        Prg01.Visible = False
        TopMenu.SbPrg01.Checked = False
        Unload XPlorer
        TopMenu.SbExplor.Checked = False
    
    Case Else   '--------------------------------------------------
        'xxxx nothing...
End Select

End Sub

Public Sub OrderWindow(WWindow As String, WWOrder As String)

Dim Result As String

If WWOrder = "Default" Or WWOrder = "3x3" Then
    If TopMenu.ViewDefault.Checked = False Then
        TopMenu.ViewDefault.Checked = True
        TopMenu.View4x4h.Checked = False
        TopMenu.View4x4v.Checked = False
        TopMenu.View3x3.Checked = False
    Else
        TopMenu.View4x4h.Checked = False
        TopMenu.View4x4v.Checked = False
        TopMenu.View3x3.Checked = False
    End If
Else
    If WWOrder = "4x4h" Then
        If TopMenu.View4x4h.Checked = False Then
            TopMenu.View4x4h.Checked = True
            TopMenu.View4x4v.Checked = False
            TopMenu.View3x3.Checked = False
            TopMenu.ViewDefault.Checked = False
        Else
            TopMenu.View4x4v.Checked = False
            TopMenu.View3x3.Checked = False
            TopMenu.ViewDefault.Checked = False
        End If
    Else
        If WWOrder = "4x4v" Then
            If TopMenu.View4x4v.Checked = False Then
                TopMenu.View4x4v.Checked = True
                TopMenu.View4x4h.Checked = False
                TopMenu.View3x3.Checked = False
                TopMenu.ViewDefault.Checked = False
            Else
                TopMenu.View4x4h.Checked = False
                TopMenu.View3x3.Checked = False
                TopMenu.ViewDefault.Checked = False
            End If
        Else
            If TopMenu.ViewDefault.Checked = False Then
                TopMenu.ViewDefault.Checked = True
                TopMenu.View4x4h.Checked = False
                TopMenu.View4x4v.Checked = False
                TopMenu.View3x3.Checked = False
                WWOrder = "Default"
            Else
                TopMenu.View4x4h.Checked = False
                TopMenu.View4x4v.Checked = False
                TopMenu.View3x3.Checked = False
                WWOrder = "Default"
            End If
        End If
    End If
End If
            
Select Case WWOrder
    Case "3x3", "Default"
        Select Case WWindow
            Case "TopMenu"  '--------------------------------------------------
                TopMenu.Top = 0
                TopMenu.Left = 0
                TopMenu.Width = TopMenu.SysInfo1.WorkAreaWidth 'largo
                TopMenu.Height = 2280   'alto
            Case "DwMenu"   '--------------------------------------------------
                If TopMenu.SbHerram.Checked = False Then
                    DownMenu.WindowState = 0
                    DownMenu.Visible = True
                    TopMenu.SbHerram.Checked = True
                End If
                DownMenu.Left = 0
                DownMenu.Height = 1400
                DownMenu.Top = TopMenu.SysInfo1.WorkAreaHeight - DownMenu.Height
                DownMenu.Width = TopMenu.SysInfo1.WorkAreaWidth
            Case "Tnd01"    '--------------------------------------------------
                If TopMenu.SbTnd01.Checked = False Then
                    Tanda01.WindowState = 0
                    Tanda01.Visible = True
                    TopMenu.SbTnd01.Checked = True
                End If
                Tanda01.Top = TopMenu.Height
                Tanda01.Left = 0
            Case "Prg01"    '--------------------------------------------------
                If TopMenu.SbPrg01.Checked = False Then
                    Prg01.WindowState = 0
                    Prg01.Visible = True
                    TopMenu.SbPrg01.Checked = True
                End If
                Prg01.Top = TopMenu.Height + Tanda01.Height
                Prg01.Left = 0
            Case "Est01"    '--------------------------------------------------
                If TopMenu.SbEst01.Checked = False Then
                    Est01.WindowState = 0
                    Est01.Visible = True
                    TopMenu.SbEst01.Checked = True
                End If
                Est01.Top = TopMenu.Height
                Est01.Left = Tanda01.Width
            Case "Est02"    '--------------------------------------------------
                If TopMenu.SbEst02.Checked = False Then
                    Est02.WindowState = 0
                    Est02.Visible = True
                    TopMenu.SbEst02.Checked = True
                End If
                Est02.Top = TopMenu.Height + Est01.Height
                Est02.Left = Tanda01.Width
                'ordenamos las ventanas
                'Result = GetWPos(2, "Default")
            Case "Explor01" '--------------------------------------------------
                If TopMenu.SbExplor.Checked = False Then
                    XPlorer.WindowState = 0
                    XPlorer.Visible = True
                    TopMenu.SbExplor.Checked = True
                End If
                XPlorer.Top = TopMenu.Height
                XPlorer.Left = Tanda01.Width + Est01.Width
            Case Else   '--------------------------------------------------
        
        End Select
    Case "4x4v" '**************************REVISAR
        'organizacion 4x4 vertical
        
    Case "4x4h" ' *************************REVISAR
        'organizacion 44x4 horizontal
        
    Case Else
        'xxx nothing...
End Select

End Sub
Private Sub SetWinPos(WWindow As String, WWOrder As String)

' ***** CREO QUE HAY QUE REMOVER. EN REVISION 18-02-24

Exit Sub

Select Case WWOrder
    Case "Default"  '********************************************************
        Select Case WWindow
            Case "Explor01"

            Case "Tnd01"
                'xxx
            Case "Prg01"
                'xxx
            Case Else
                'xxx
        End Select
    Case "4x4v"     '********************************************************
        Select Case WWindow
            Case "Explor01"

            Case "Tnd01"
                If Prg01.WindowState = 1 Then

                Else

                End If
            Case "Prg01"
                'xxx
            Case Else
                'xxx
        End Select
    Case "4x4h"     '********************************************************
        Select Case WWindow
            Case "Explor01"

            Case "Tnd01"
                If Est01.WindowState = 1 Then

                Else

                End If
            Case "Prg01"
                'xxx
            Case Else
                'xxx
        End Select
    Case Else       '********************************************************
        'xxx
End Select

End Sub

Public Sub ShowWindow(WWindow As String)

Dim WWOrder As String

If WWindow = "Startup1" Then
    WWindow = "All"
    WWOrder = "Default"
    GoSub SelectWin
Else
    If WWindow = "Startup2" Then
        WWindow = "All"
        WWOrder = "4x4h"
        GoSub SelectWin
    Else
        If WWindow = "Startup3" Then
            WWindow = "All"
            WWOrder = "4x4v"
            GoSub SelectWin
        Else
            'xxx nothing to do
        End If
    End If
End If

If TopMenu.View3x3.Checked = True Or TopMenu.ViewDefault.Checked = True Then
    WWOrder = "Default"
End If
If TopMenu.View4x4h.Checked = True Then
    WWOrder = "4x4h"
End If
If TopMenu.View4x4v.Checked = True Then
    WWOrder = "4x4v"
End If

SelectWin:
Select Case WWindow
    Case "TopMenu"  '--------------------------------------------------
        If TopMenu.WindowState = 1 Then
            TopMenu.WindowState = 0
            TopMenu.Visible = True
        Else
            TopMenu.Show
        End If
        OrderWindow "TopMenu", WWOrder
    Case "DwMenu"   '--------------------------------------------------
        If DownMenu.WindowState = 1 Then
            DownMenu.WindowState = 0
            DownMenu.Visible = True
        Else
            DownMenu.Show
        End If
        OrderWindow "DwMenu", WWOrder
    Case "Tnd01"    '--------------------------------------------------
        If Tanda01.WindowState = 1 Then
            Tanda01.WindowState = 0
            Tanda01.Visible = True
        Else
            Tanda01.Show
        End If
        OrderWindow "Tnd01", WWOrder
    Case "Prg01"    '--------------------------------------------------
        If Prg01.WindowState = 1 Then
            Prg01.WindowState = 0
            Prg01.Visible = True
        Else
            Prg01.Show
        End If
        OrderWindow "Prg01", WWOrder
    Case "Est01"    '--------------------------------------------------
        If Est01.WindowState = 1 Then
            Est01.WindowState = 0
            Est01.Visible = True
        Else
            Est01.Show
        End If
        OrderWindow "Est01", WWOrder
    Case "Est02"    '--------------------------------------------------
        If Est02.WindowState = 1 Then
            Est02.WindowState = 0
            Est02.Visible = True
        Else
            Est02.Show
        End If
        OrderWindow "Est02", WWOrder
    Case "Explor01" '--------------------------------------------------
        If XPlorer.WindowState = 1 Then
            XPlorer.WindowState = 0
            XPlorer.Visible = True
        Else
            XPlorer.Show
        End If
        OrderWindow "Explor01", WWOrder
    Case "All"  '--------------------------------------------------
        'Activamos las ventanas determinadas por defecto en el programa
        If TopMenu.WindowState = 1 Then
            TopMenu.WindowState = 0
            TopMenu.Visible = True
        Else
            TopMenu.Show
        End If
        OrderWindow "TopMenu", WWOrder
        If DownMenu.WindowState = 1 Then
            DownMenu.WindowState = 0
            DownMenu.Visible = True
        Else
            DownMenu.Show
        End If
        OrderWindow "DwMenu", WWOrder
        If Tanda01.WindowState = 1 Then
            Tanda01.WindowState = 0
            Tanda01.Visible = True
        Else
            Tanda01.Show
        End If
        OrderWindow "Tnd01", WWOrder
        If Prg01.WindowState = 1 Then
            Prg01.WindowState = 0
            Prg01.Visible = True
        Else
            Prg01.Show
        End If
        OrderWindow "Prg01", WWOrder
        If Est01.WindowState = 1 Then
            Est01.WindowState = 0
            Est01.Visible = True
        Else
            Est01.Show
        End If
        OrderWindow "Est01", WWOrder
        If Est02.WindowState = 1 Then
            Est02.WindowState = 0
            Est02.Visible = True
        Else
            Est02.Show
        End If
        OrderWindow "Est02", WWOrder
        If XPlorer.WindowState = 1 Then
            XPlorer.WindowState = 0
            XPlorer.Visible = True
        Else
            XPlorer.Show
        End If
        OrderWindow "Explor01", WWOrder
    Case Else   '--------------------------------------------------
        'xxx nothing...
End Select

End Sub

