Attribute VB_Name = "WindowManager"
'********************* RM100 *********************
'     RADIO MAKER WINDOW CONTROLLER MODULE
'COPYRIGHT (C) 1987-2024 ONLY development inc.
'Christian A. Del Monte
'*************************************************
' ultima modificacion: 18-02-24
'*************************************************

Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

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
                TopMenu.Width = TopMenu.SysInfo1.WorkAreaWidth
                TopMenu.Height = 1530
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
            Case "Est01"    '--------------------------------------------------
                If TopMenu.SbEst01.Checked = False Then
                    Est01.WindowState = 0
                    Est01.Visible = True
                    TopMenu.SbEst01.Checked = True
                End If
                Est01.Top = TopMenu.Height
                Est01.Left = WidthMax / 3
                'ordenamos las ventanas
                'Result = GetWPos(1, "Default")
            Case "Est02"    '--------------------------------------------------
                If TopMenu.SbEst02.Checked = False Then
                    Est02.WindowState = 0
                    Est02.Visible = True
                    TopMenu.SbEst02.Checked = True
                End If
                Est02.Top = TopMenu.Height + Est01.Height
                Est02.Left = WidthMax / 3
                'ordenamos las ventanas
                'Result = GetWPos(2, "Default")
            Case "Tnd01"    '--------------------------------------------------
                If TopMenu.SbTnd01.Checked = False Then
                    Tanda01.WindowState = 0
                    Tanda01.Visible = True
                    TopMenu.SbTnd01.Checked = True
                End If
                Tanda01.Top = TopMenu.Height
                Tanda01.Left = 0
                'ordenamos los controles
                'SetWinPos "Tnd01", WWOrder
            Case "Prg01"    '--------------------------------------------------
                If TopMenu.SbPrg01.Checked = False Then
                    Prg01.WindowState = 0
                    Prg01.Visible = True
                    TopMenu.SbPrg01.Checked = True
                End If
                Prg01.Top = TopMenu.Height + Tanda01.Height
                Prg01.Left = 0
            Case "Explor01" '--------------------------------------------------
                If TopMenu.SbExplor.Checked = False Then
                    XPlorer.Show
                    TopMenu.SbExplor.Checked = True
                End If
                'ordenamos la ventana y sus controles
                'SetWinPos "Explor01", WWOrder
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
                XPlorer.tvwDirTree.Height = 4695: XPlorer.tvwDirTree.Left = 120
                XPlorer.tvwDirTree.Top = 720: XPlorer.tvwDirTree.Width = 3015
                XPlorer.File1.Height = 3795: XPlorer.File1.Left = 3240
                XPlorer.File1.Top = 720: XPlorer.File1.Width = 3495
                XPlorer.Label4.Height = 255: XPlorer.Label4.Left = 3240
                XPlorer.Label4.Top = 4800: XPlorer.Label4.Width = 1815
                XPlorer.ExCombo.Left = 3240
                XPlorer.ExCombo.Top = 5040: XPlorer.ExCombo.Width = 3495
                XPlorer.Line1.X1 = 120: XPlorer.Line1.X2 = 6720
                XPlorer.Line1.Y1 = 5520: XPlorer.Line1.Y2 = 5520
                XPlorer.Command1.Height = 375: XPlorer.Command1.Left = 5640
                XPlorer.Command1.Top = 5640: XPlorer.Command1.Width = 1095
                XPlorer.Command2.Height = 375: XPlorer.Command2.Left = 120
                XPlorer.Command2.Top = 5640: XPlorer.Command2.Width = 2055
            Case "Tnd01"
                Tanda01.T1View.Height = Tanda01.Height - 2200
                Tanda01.T1View.Width = Tanda01.Width - 300
                    Tanda01.Label6.Top = Tanda01.Height - 1210
                    Tanda01.Label5.Top = Tanda01.Label6.Top + 70
                    Tanda01.Label7.Top = Tanda01.Label6.Top + 70
                    Tanda01.Label8.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1T1.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1T2.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1T3.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1T4.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1T5.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1t6.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1t7.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1t8.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1I1.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1I2.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1I3.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1I4.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1I5.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1I6.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1I7.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1I8.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1F1.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1F2.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1F3.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1F4.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1F5.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1F6.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1F7.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1F8.Top = Tanda01.Label6.Top + 70
                Tanda01.T1Shape.Top = Tanda01.Height - 850
                Tanda01.CmdBlock.Top = Tanda01.T1Shape.Top + 70
                Tanda01.T1Next.Top = Tanda01.T1Shape.Top + 70
                Tanda01.T1Play.Top = Tanda01.T1Shape.Top + 70
                Tanda01.T1Stop.Top = Tanda01.T1Shape.Top + 70
                Tanda01.T1New.Top = Tanda01.T1Shape.Top + 70
                Tanda01.T1Open.Top = Tanda01.T1Shape.Top + 70
                Tanda01.T1Save.Top = Tanda01.T1Shape.Top + 70
                Tanda01.T1Prop.Top = Tanda01.T1Shape.Top + 70
                Tanda01.T1Up.Top = Tanda01.T1Shape.Top + 70
                Tanda01.T1Down.Top = Tanda01.T1Shape.Top + 70
                Tanda01.T1Del.Top = Tanda01.T1Shape.Top + 70
                Tanda01.T1Order.Top = Tanda01.T1Shape.Top + 70
                Tanda01.T1OrderA.Top = Tanda01.T1Shape.Top + 70
                Tanda01.Prbar1.Top = Tanda01.Height - 1530
                Tanda01.Prbar1.Width = Tanda01.T1View.Width
                'xxx
            Case "Prg01"
                'xxx
            Case Else
                'xxx
        End Select
    Case "4x4v"     '********************************************************
        Select Case WWindow
            Case "Explor01"
                XPlorer.tvwDirTree.Height = 4695: XPlorer.tvwDirTree.Left = 120
                XPlorer.tvwDirTree.Top = 720: XPlorer.tvwDirTree.Width = 3015
                XPlorer.File1.Height = 3795: XPlorer.File1.Left = 3240
                XPlorer.File1.Top = 720: XPlorer.File1.Width = 3495
                XPlorer.Label4.Height = 255: XPlorer.Label4.Left = 3240
                XPlorer.Label4.Top = 4800: XPlorer.Label4.Width = 1815
                XPlorer.ExCombo.Left = 3240
                XPlorer.ExCombo.Top = 5040: XPlorer.ExCombo.Width = 3495
                XPlorer.Line1.X1 = 120: XPlorer.Line1.X2 = 6720
                XPlorer.Line1.Y1 = 5520: XPlorer.Line1.Y2 = 5520
                XPlorer.Command1.Height = 375: XPlorer.Command1.Left = 5640
                XPlorer.Command1.Top = 5640: XPlorer.Command1.Width = 1095
                XPlorer.Command2.Height = 375: XPlorer.Command2.Left = 120
                XPlorer.Command2.Top = 5640: XPlorer.Command2.Width = 2055
            Case "Tnd01"
                If Prg01.WindowState = 1 Then
                    Tanda01.T1View.Height = Tanda01.Height - 2200
                    Tanda01.T1View.Width = Tanda01.Width - 300
                        Tanda01.Label6.Top = Tanda01.Height - 1210
                        Tanda01.Label5.Top = Tanda01.Label6.Top + 70
                        Tanda01.Label7.Top = Tanda01.Label6.Top + 70
                        Tanda01.Label8.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T1.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T2.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T3.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T4.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T5.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1t6.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1t7.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1t8.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I1.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I2.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I3.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I4.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I5.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I6.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I7.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I8.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F1.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F2.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F3.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F4.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F5.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F6.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F7.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F8.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1Shape.Top = Tanda01.Height - 850
                    Tanda01.CmdBlock.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Next.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Play.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Stop.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1New.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Open.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Save.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Prop.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Up.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Down.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Del.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Order.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1OrderA.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.Prbar1.Top = Tanda01.Height - 1530
                    Tanda01.Prbar1.Width = Tanda01.T1View.Width
                Else
                    Tanda01.T1View.Height = Tanda01.Height - 2200   '1850
                    Tanda01.T1View.Width = Tanda01.Width - 300
                        Tanda01.Label6.Top = Tanda01.Height - 1210
                        Tanda01.Label5.Top = Tanda01.Label6.Top + 70
                        Tanda01.Label7.Top = Tanda01.Label6.Top + 70
                        Tanda01.Label8.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T1.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T2.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T3.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T4.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T5.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1t6.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1t7.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1t8.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I1.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I2.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I3.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I4.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I5.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I6.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I7.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I8.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F1.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F2.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F3.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F4.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F5.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F6.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F7.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F8.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1Shape.Top = Tanda01.Height - 850
                    Tanda01.CmdBlock.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Next.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Play.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Stop.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1New.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Open.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Save.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Prop.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Up.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Down.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Del.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Order.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1OrderA.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.Prbar1.Top = Tanda01.Height - 1530
                    Tanda01.Prbar1.Width = Tanda01.T1View.Width
                End If
            Case "Prg01"
                'xxx
            Case Else
                'xxx
        End Select
    Case "4x4h"     '********************************************************
        Select Case WWindow
            Case "Explor01"
                XPlorer.tvwDirTree.Height = 4695: XPlorer.tvwDirTree.Left = 120
                XPlorer.tvwDirTree.Top = 720: XPlorer.tvwDirTree.Width = 3015
                XPlorer.File1.Height = 3795: XPlorer.File1.Left = 3240
                XPlorer.File1.Top = 720: XPlorer.File1.Width = 3495
                XPlorer.Label4.Height = 255: XPlorer.Label4.Left = 3240
                XPlorer.Label4.Top = 4800: XPlorer.Label4.Width = 1815
                XPlorer.ExCombo.Left = 3240
                XPlorer.ExCombo.Top = 5040: XPlorer.ExCombo.Width = 3495
                XPlorer.Line1.X1 = 120: XPlorer.Line1.X2 = 6720
                XPlorer.Line1.Y1 = 5520: XPlorer.Line1.Y2 = 5520
                XPlorer.Command1.Height = 375: XPlorer.Command1.Left = 5640
                XPlorer.Command1.Top = 5640: XPlorer.Command1.Width = 1095
                XPlorer.Command2.Height = 375: XPlorer.Command2.Left = 120
                XPlorer.Command2.Top = 5640: XPlorer.Command2.Width = 2055
            Case "Tnd01"
                If Est01.WindowState = 1 Then
                    Tanda01.T1View.Height = Tanda01.Height - 2200
                    Tanda01.T1View.Width = Tanda01.Width - 300
                        Tanda01.Label6.Top = Tanda01.Height - 1210
                        Tanda01.Label5.Top = Tanda01.Label6.Top + 70
                        Tanda01.Label7.Top = Tanda01.Label6.Top + 70
                        Tanda01.Label8.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T1.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T2.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T3.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T4.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T5.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1t6.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1t7.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1t8.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I1.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I2.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I3.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I4.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I5.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I6.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I7.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I8.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F1.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F2.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F3.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F4.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F5.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F6.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F7.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F8.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1Shape.Top = Tanda01.Height - 850
                    Tanda01.CmdBlock.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Next.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Play.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Stop.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1New.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Open.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Save.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Prop.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Up.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Down.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Del.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Order.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1OrderA.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.Prbar1.Top = Tanda01.Height - 1530
                    Tanda01.Prbar1.Width = Tanda01.T1View.Width
                Else
                    Tanda01.T1View.Height = Tanda01.Height - 2200
                    Tanda01.T1View.Width = Tanda01.Width - 300
                        Tanda01.Label6.Top = Tanda01.Height - 1210
                        Tanda01.Label5.Top = Tanda01.Label6.Top + 70
                        Tanda01.Label7.Top = Tanda01.Label6.Top + 70
                        Tanda01.Label8.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T1.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T2.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T3.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T4.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1T5.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1t6.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1t7.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1t8.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I1.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I2.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I3.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I4.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I5.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I6.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I7.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1I8.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F1.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F2.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F3.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F4.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F5.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F6.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F7.Top = Tanda01.Label6.Top + 70
                        Tanda01.T1F8.Top = Tanda01.Label6.Top + 70
                    Tanda01.T1Shape.Top = Tanda01.Height - 850
                    Tanda01.CmdBlock.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Next.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Play.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Stop.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1New.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Open.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Save.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Prop.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Up.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Down.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Del.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1Order.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.T1OrderA.Top = Tanda01.T1Shape.Top + 70
                    Tanda01.Prbar1.Top = Tanda01.Height - 1530
                    Tanda01.Prbar1.Width = Tanda01.T1View.Width
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
        XPlorer.Show
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
    Case Else   '--------------------------------------------------
        'xxx nothing...
End Select

End Sub

