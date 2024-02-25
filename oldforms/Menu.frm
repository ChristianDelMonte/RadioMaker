VERSION 5.00
Begin VB.Form EstMenu 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   2205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   0
      TabIndex        =   1
      Top             =   -90
      Width           =   2190
      Begin VB.PictureBox PcN 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   345
         ScaleHeight     =   285
         ScaleWidth      =   1815
         TabIndex        =   14
         Top             =   120
         Width           =   1815
         Begin VB.Image PcN1 
            Height          =   225
            Left            =   15
            Picture         =   "Menu.frx":0000
            Top             =   15
            Width           =   240
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nuevo"
            Height          =   195
            Left            =   315
            TabIndex        =   15
            Top             =   45
            Width           =   600
         End
      End
      Begin VB.PictureBox PcO 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   345
         ScaleHeight     =   285
         ScaleWidth      =   1815
         TabIndex        =   12
         Top             =   390
         Width           =   1815
         Begin VB.Image PcO1 
            Height          =   225
            Left            =   15
            Picture         =   "Menu.frx":0532
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Abrir..."
            Height          =   195
            Left            =   315
            TabIndex        =   13
            Top             =   45
            Width           =   555
         End
      End
      Begin VB.PictureBox PcG 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   345
         ScaleHeight     =   285
         ScaleWidth      =   1815
         TabIndex        =   10
         Top             =   660
         Width           =   1815
         Begin VB.Image PcG1 
            Height          =   225
            Left            =   15
            Picture         =   "Menu.frx":0A64
            Top             =   15
            Width           =   240
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Guardar"
            Height          =   195
            Left            =   315
            TabIndex        =   11
            Top             =   45
            Width           =   645
         End
      End
      Begin VB.PictureBox PcE 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   345
         ScaleHeight     =   285
         ScaleWidth      =   1815
         TabIndex        =   8
         Top             =   1920
         Width           =   1815
         Begin VB.Image PcE1 
            Height          =   240
            Left            =   15
            Picture         =   "Menu.frx":0F96
            Top             =   15
            Width           =   240
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Eliminar"
            Height          =   195
            Left            =   315
            TabIndex        =   9
            Top             =   45
            Width           =   600
         End
      End
      Begin VB.PictureBox PcP 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   345
         ScaleHeight     =   285
         ScaleWidth      =   1815
         TabIndex        =   6
         Top             =   1605
         Width           =   1815
         Begin VB.Image PcP1 
            Height          =   240
            Left            =   15
            Picture         =   "Menu.frx":1098
            Top             =   15
            Width           =   240
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Propiedades..."
            Height          =   195
            Left            =   315
            TabIndex        =   7
            Top             =   45
            Width           =   1140
         End
      End
      Begin VB.PictureBox PcGc 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   345
         ScaleHeight     =   285
         ScaleWidth      =   1815
         TabIndex        =   4
         Top             =   930
         Width           =   1815
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Guardar como..."
            Height          =   195
            Left            =   315
            TabIndex        =   5
            Top             =   45
            Width           =   1230
         End
      End
      Begin VB.PictureBox PcR 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   345
         ScaleHeight     =   285
         ScaleWidth      =   1815
         TabIndex        =   2
         Top             =   1290
         Width           =   1815
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Renombrar"
            Height          =   195
            Left            =   315
            TabIndex        =   3
            Top             =   45
            Width           =   825
         End
      End
      Begin VB.Image Image1 
         Height          =   1800
         Left            =   30
         Picture         =   "Menu.frx":119A
         Top             =   435
         Width           =   300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   315
         X2              =   315
         Y1              =   2220
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   300
         X2              =   2145
         Y1              =   1245
         Y2              =   1245
      End
   End
   Begin VB.Label LblEst 
      Caption         =   "Est01"
      Height          =   285
      Left            =   795
      TabIndex        =   0
      Top             =   2565
      Width           =   555
   End
End
Attribute VB_Name = "EstMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PcE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PcN.BackColor = &HC0C0C0
    PcN1.BorderStyle = 0
PcO.BackColor = &HC0C0C0
    PcO1.BorderStyle = 0
PcG.BackColor = &HC0C0C0
    PcG1.BorderStyle = 0
PcGc.BackColor = &HC0C0C0
PcR.BackColor = &HC0C0C0
PcP.BackColor = &HC0C0C0    'gris normal
    PcP1.BorderStyle = 0
PcE.BackColor = &HFFFFFF   'gris claro
    PcE1.BorderStyle = 1

End Sub

Private Sub PcG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PcN.BackColor = &HC0C0C0
    PcN1.BorderStyle = 0
PcO.BackColor = &HC0C0C0    'gris normal
    PcO1.BorderStyle = 0
PcG.BackColor = &HFFFFFF    'gris claro
    PcG1.BorderStyle = 1
PcGc.BackColor = &HC0C0C0
PcR.BackColor = &HC0C0C0
PcP.BackColor = &HC0C0C0
    PcP1.BorderStyle = 0
PcE.BackColor = &HC0C0C0
    PcE1.BorderStyle = 0

End Sub

Private Sub PcGc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PcN.BackColor = &HC0C0C0
    PcN1.BorderStyle = 0
PcO.BackColor = &HC0C0C0
    PcO1.BorderStyle = 0
PcG.BackColor = &HC0C0C0    'gris normal
    PcG1.BorderStyle = 0
PcGc.BackColor = &HFFFFFF   'gris claro
PcR.BackColor = &HC0C0C0
PcP.BackColor = &HC0C0C0
    PcP1.BorderStyle = 0
PcE.BackColor = &HC0C0C0
    PcE1.BorderStyle = 0

End Sub

Private Sub PcN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PcN.BackColor = &HFFFFFF    'gris claro
    PcN1.BorderStyle = 1
PcO.BackColor = &HC0C0C0    'gris normal
    PcO1.BorderStyle = 0
PcG.BackColor = &HC0C0C0
    PcG1.BorderStyle = 0
PcGc.BackColor = &HC0C0C0
PcR.BackColor = &HC0C0C0
PcP.BackColor = &HC0C0C0
    PcP1.BorderStyle = 0
PcE.BackColor = &HC0C0C0
    PcE1.BorderStyle = 0

End Sub

Private Sub PcO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PcN.BackColor = &HC0C0C0    'gris normal
    PcN1.BorderStyle = 0
PcO.BackColor = &HFFFFFF    'gris claro
    PcO1.BorderStyle = 1
PcG.BackColor = &HC0C0C0
    PcG1.BorderStyle = 0
PcGc.BackColor = &HC0C0C0
PcR.BackColor = &HC0C0C0
PcP.BackColor = &HC0C0C0
    PcP1.BorderStyle = 0
PcE.BackColor = &HC0C0C0
    PcE1.BorderStyle = 0

End Sub

Private Sub PcP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PcN.BackColor = &HC0C0C0
    PcN1.BorderStyle = 0
PcO.BackColor = &HC0C0C0
    PcO1.BorderStyle = 0
PcG.BackColor = &HC0C0C0
    PcG1.BorderStyle = 0
PcGc.BackColor = &HC0C0C0
PcR.BackColor = &HC0C0C0    'gris normal
PcP.BackColor = &HFFFFFF   'gris claro
    PcP1.BorderStyle = 1
PcE.BackColor = &HC0C0C0
    PcE1.BorderStyle = 0

End Sub

Private Sub PcR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

PcN.BackColor = &HC0C0C0
    PcN1.BorderStyle = 0
PcO.BackColor = &HC0C0C0
    PcO1.BorderStyle = 0
PcG.BackColor = &HC0C0C0
    PcG1.BorderStyle = 0
PcGc.BackColor = &HC0C0C0    'gris normal
PcR.BackColor = &HFFFFFF   'gris claro
PcP.BackColor = &HC0C0C0
    PcP1.BorderStyle = 0
PcE.BackColor = &HC0C0C0
    PcE1.BorderStyle = 0

End Sub
