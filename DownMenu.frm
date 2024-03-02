VERSION 5.00
Begin VB.Form DownMenu 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   15240
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSearch 
      Caption         =   "SRCH"
      Height          =   750
      Left            =   6030
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Explorador de archivos"
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton CmdBlock 
      Caption         =   "BLOCK"
      Enabled         =   0   'False
      Height          =   750
      Left            =   5265
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Creador de Bloques"
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton CmdPH 
      Caption         =   "PH"
      Enabled         =   0   'False
      Height          =   750
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Programación Horaria"
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton Est2Cmd 
      Caption         =   "E2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Estacion 02"
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton Est1Cmd 
      Caption         =   "E1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   2790
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Estacion 01"
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton PrgTndCmd 
      Caption         =   "PRG"
      Height          =   750
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Programacion de Tandas"
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton TndCmd 
      Caption         =   "TND"
      Height          =   750
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Creacion de Tandas"
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton CmdTrash 
      Caption         =   "DEL"
      Height          =   750
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Papelera"
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton CmdConfig 
      Caption         =   "CNFG"
      Height          =   750
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Configuracion"
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton XCmd 
      Caption         =   "X"
      Height          =   750
      Left            =   14310
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir del Sistema"
      Top             =   450
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin RM100.TitelBar TitelBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      BackColor       =   8421504
      BackColorCover  =   3
      BackColorV2Begin=   4210752
      BackColorV2End  =   0
      BackColorV1Begin=   4210752
      BackColorV1End  =   0
      ForeColor       =   16777215
      ShowMinimized   =   0   'False
      ShowMaximizedEnabled=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "  Barra de Herramientas"
      CaptionPosX     =   1
      BorderNormal    =   2
      BorderColorDarkLight=   12632256
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"DownMenu.frx":0000
      ForeColor       =   &H80000013&
      Height          =   615
      Left            =   8520
      TabIndex        =   10
      Top             =   540
      Width           =   5655
   End
End
Attribute VB_Name = "DownMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dimensiones de resultado
Dim RResult As String

Private Sub CmdBlock_Click()

FrmBlock.Show

End Sub

Private Sub CmdConfig_Click()

Config.Show 1, Me

End Sub

Private Sub CmdPH_Click()

If FrmTime.WindowState = 1 Then
    FrmTime.WindowState = 0
    FrmTime.Visible = True
Else
    FrmTime.Show
End If

End Sub

Private Sub CmdSearch_Click()

ShowWindow "Explor01"

End Sub

Private Sub CmdTrash_Click()

Dim i As Integer
Dim X As Integer

'BORRAR ITEMS DE LA ESTACION 01 Y ESTACION 02
For i = 0 To 21
    If Est01.E11(i).BackColor = &HFF& Then       'rojo
        Est12Data.N1(i).Caption = ""             'nombre y path
        Est12Data.c1(i).Caption = ""             'nombre solo
        Est12Data.D1(i).Caption = ""
        Est01.E11(i).Caption = ""                'nombre del archivo
        Est01.E11(i).BackColor = &H8000000F      'gris
        Est01.E11(i).ToolTipText = ""
        Est12Data.V1(i).Caption = ""             'stream or music?
    End If
    If Est02.E21(i).BackColor = &HFF& Then       'rojo
        Est12Data.N2(i).Caption = ""             'nombre y path
        Est12Data.c2(i).Caption = ""             'nombre solo
        Est12Data.D2(i).Caption = ""
        Est02.E21(i).Caption = ""                'nombre del archivo
        Est02.E21(i).BackColor = &H8000000F      'gris
        Est02.E21(i).ToolTipText = ""
        Est12Data.V2(i).Caption = ""             'stream or music?
    End If
Next i

'BORRAR ITEMS DE LA PROGRAMACION DE TANDAS
For X = 0 To 23
    If Prg01.Prg1(X).BackColor = &HFF& Then       'rojo
        Est12Data.PF(X).Caption = ""             'nombre y path
        Est12Data.PC(X).Caption = ""             'nombre solo
        Est12Data.PD(X).Caption = ""             'duracion
        Prg01.Prg1(X).Caption = ""
        Prg01.Prg1(X).BackColor = &H8000000F      'gris
        Prg01.Prg1(X).ToolTipText = ""
    End If
Next X

End Sub

Private Sub Est1Cmd_Click()

    If TopMenu.SbEst01.Checked = False Then
        ShowWindow "Est01"
        'If Est01.Command1.Caption = ">" Then
        '    Est01.Width = 15360
        '    Est01.Left = 0
        'End If
        If Tanda01.WindowState = 0 Then
            ShowWindow "Tnd01"
        End If
    Else
        HideWindow "Est01"
        If Tanda01.WindowState = 0 Then
            ShowWindow "Tnd01"
        End If
    End If

End Sub

Private Sub Est2Cmd_Click()

    If TopMenu.SbEst02.Checked = False Then
        ShowWindow "Est02"
        'If Est02.Command1.Caption = ">" Then
        '    Est02.Width = 15360
        '    Est02.Left = 0
        'End If
        If Prg01.WindowState = 0 Then
            ShowWindow "Prg01"
        End If
    Else
        HideWindow "Est02"
        If Prg01.WindowState = 0 Then
            ShowWindow "Prg01"
        End If
    End If
    
End Sub

Private Sub Form_Load()

'*** load some pictures *****
Me.Picture = LoadPicture(App.path & "\Imagenes\FND_COMPLETO.jpg")

TndCmd.Picture = LoadResPicture("ICO_TND", 0): TndCmd.Caption = ""
PrgTndCmd.Picture = LoadResPicture("ICO_PRG", 0): PrgTndCmd.Caption = ""
Est1Cmd.Picture = LoadResPicture("ICO_EST", 0): Est1Cmd.Caption = ""
Est2Cmd.Picture = LoadResPicture("ICO_EST", 0): Est2Cmd.Caption = ""
CmdSearch.Picture = LoadResPicture("ICO_SEARCH", 0): CmdSearch.Caption = ""
CmdTrash.Picture = LoadResPicture("ICO_TRASH", 0): CmdTrash.Caption = ""
XCmd.Picture = LoadResPicture("ICO_EXIT", 0): XCmd.Caption = ""
CmdConfig.Picture = LoadResPicture("ICO_CONFIG", 0): CmdConfig.Caption = ""
'CmdPH.Picture = LoadResPicture("ICO_PH", 0)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

HideWindow "DwMenu"

End Sub

Private Sub Form_Terminate()

HideWindow "DwMenu"

End Sub

Private Sub Form_Unload(Cancel As Integer)

HideWindow "DwMenu"

End Sub

Private Sub PrgTndCmd_Click()

    If TopMenu.SbPrg01.Checked = False Then
        ShowWindow "Prg01"
        If Tanda01.WindowState = 0 Then
            ShowWindow "Tnd01"
        End If
        ShowWindow "Prg01"
    Else
        HideWindow "Prg01"
        If Tanda01.WindowState = 0 Then
            ShowWindow "Tnd01"
        End If
    End If
    
End Sub

Private Sub TndCmd_Click()

    If TopMenu.SbTnd01.Checked = False Then
        ShowWindow "Tnd01"
        If Prg01.WindowState = 0 Then
            ShowWindow "Prg01"
        End If
    Else
        HideWindow "Tnd01"
        If Prg01.WindowState = 0 Then
            ShowWindow "Prg01"
        End If
    End If

End Sub

Private Sub XCmd_Click()

Call TopMenu.EndApp

End Sub
