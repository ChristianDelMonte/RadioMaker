VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{82B99A8B-A2F9-4E4F-B970-F5381DB68D7B}#4.0#0"; "DC_Control_Bt.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Explorador de archivos"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   13605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   13020
      TabIndex        =   8
      Top             =   2520
      Width           =   435
   End
   Begin ComctlLib.StatusBar StatusBar2 
      Height          =   315
      Left            =   2910
      TabIndex        =   7
      Top             =   60
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin DC_Control_Bt.dcButton dcButton3 
      Height          =   345
      Left            =   660
      TabIndex        =   5
      Top             =   30
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   609
      BackColor       =   15133676
      ButtonStyle     =   7
      Caption         =   "/"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DC_Control_Bt.dcButton dcButton2 
      Height          =   345
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   609
      BackColor       =   15133676
      ButtonStyle     =   7
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DC_Control_Bt.dcButton dcButton1 
      Height          =   345
      Index           =   0
      Left            =   12990
      TabIndex        =   2
      Top             =   390
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   609
      BackColor       =   15133676
      ButtonStyle     =   7
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   2895
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2415
      Left            =   4470
      TabIndex        =   3
      Top             =   420
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin DC_Control_Bt.dcButton dcButton4 
      Height          =   345
      Left            =   1020
      TabIndex        =   6
      Top             =   30
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   609
      BackColor       =   15133676
      ButtonStyle     =   7
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Sub Load_tbDrives()

Static i As Integer
Dim sDrive As String, strSave As String
Dim Ret As String
Dim keer As Integer

strSave = String(255, Chr$(0))
Ret = GetLogicalDriveStrings(255, strSave)
For keer = 1 To 100
    If Left$(strSave, InStr(1, strSave, Chr$(0))) = Chr$(0) Then Exit For
        sDrive = Left$(LCase(strSave), InStr(1, strSave, Chr$(0)) - 3)
        strSave = Right$(strSave, Len(strSave) - InStr(1, strSave, Chr$(0)))
    If Not i = 0 Then
Load dcButton2(i)
dcButton2(i).Caption = sDrive
dcButton2(i).Left = dcButton2(i - 1).Left + dcButton2(i - 1).Width + 70
dcButton2(i).Visible = True

End If

dcButton2(i).Caption = sDrive
     
Select Case GetDriveType(sDrive & ":\")
    Case 2
        Set dcButton2(i).PictureNormal = LoadPicture(App.Path & "\icons\2.ico")
        dcButton2(i).ToolTipText = "Disco Removible"
        'List1.AddItem "[-" & sDrive & "-]" & vbTab & "3½"""
        
    Case 3
        Set dcButton2(i).PictureNormal = LoadPicture(App.Path & "\icons\3.ico")
        dcButton2(i).ToolTipText = "Disco duro"
        'List1.AddItem "[-" & sDrive & "-]" & vbTab & VolumeName(sDrive)
    
    Case 4
        Set dcButton2(i).PictureNormal = LoadPicture(App.Path & "\icons\4.ico")
        dcButton2(i).ToolTipText = "Unidad remota"
        'List1.AddItem "[-" & sDrive & "-]" & vbTab & "Remote"
    
    Case 5
        Set dcButton2(i).PictureNormal = LoadPicture(App.Path & "\icons\5.ico")
        dcButton2(i).ToolTipText = "CD-ROM"
        'List1.AddItem "[-" & sDrive & "-]" & vbTab & "CD-ROM"
    
    Case 6
        Set dcButton2(i).PictureNormal = LoadPicture(App.Path & "\icons\6.ico")
        dcButton2(i).ToolTipText = "RAM Disk"
        'List1.AddItem "[-" & sDrive & "-]" & vbTab & "RAM Disk"
    
    Case Else
        Set dcButton2(i).PictureNormal = LoadPicture(App.Path & "\icons\7.ico")
        dcButton2(i).ToolTipText = "Desconocido"
        'List1.AddItem "[-" & sDrive & "-]" & vbTab & "Unknown"
End Select

If LCase(Mid(sDrive, 1, 1)) = LCase(Mid(Dir1.Path, 1, 1)) Then
    'Call GetDiskFreeSpaceEx(Mid(sDrive, 1, 1) & ":\", BytesFreeToCalller, TotalBytes, TotalFreeBytes)
    'StatusBar2.Panels(1).Text = Format$(((TotalFreeBytes * 10000) / 1024), "###,###,###,##0") & " Kb" & "  of  " & Format$(((TotalBytes * 10000) / 1024), "###,###,###,##0") & " Kb free.  (" & Format(100 - ((TotalFreeBytes * 100) / TotalBytes), "##") & "% used.)"
    status2 (Mid(sDrive, 1, 1))
End If
    
i = i + 1

Next keer

'List1.Height = List1.Height * List1.ListCount
dcButton3.Left = dcButton2(i - 1).Left + dcButton2(i - 1).Width + 70
dcButton4.Left = dcButton3.Left + dcButton3.Width + 70
StatusBar2.Left = dcButton4.Left + dcButton4.Width + 70
'StatusBar2.Width = ListView1.Width - StatusBar2.Left

End Sub

Public Function status2(uni As String)

Dim r As Long, BytesFreeToCalller As Currency, TotalBytes As Currency
Dim TotalFreeBytes As Currency, TotalBytesUsed As Currency

Call GetDiskFreeSpaceEx(uni & ":\", BytesFreeToCalller, TotalBytes, TotalFreeBytes)
Call LockWindowUpdate(StatusBar2.hwnd)

StatusBar2.Panels(1).Text = Format$(((TotalFreeBytes * 10000) / 1024), "###,###,###,##0") & " Kb" & "  de  " & Format$(((TotalBytes * 10000) / 1024), "###,###,###,##0") & " Kb libres.  (" & Format(100 - ((TotalFreeBytes * 100) / TotalBytes), "##") & "% usado.)"
Call LockWindowUpdate(0)

End Function

Private Sub Form_Load()

Load_tbDrives

End Sub
