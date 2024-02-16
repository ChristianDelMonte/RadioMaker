VERSION 5.00
Begin VB.Form Prg01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PROGRAMACION DE TANDAS - Detenido"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7590
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox p1t1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   5925
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   75
      Width           =   190
   End
   Begin VB.PictureBox p1t2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6120
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   75
      Width           =   190
   End
   Begin VB.PictureBox p1t3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6300
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   75
      Width           =   190
   End
   Begin VB.PictureBox p1t4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6495
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   75
      Width           =   190
   End
   Begin VB.PictureBox p1t5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6690
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   75
      Width           =   190
   End
   Begin VB.PictureBox p1t6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6870
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   75
      Width           =   190
   End
   Begin VB.PictureBox p1t7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   7050
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   75
      Width           =   190
   End
   Begin VB.PictureBox p1t8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   7230
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   75
      Width           =   190
   End
   Begin VB.TextBox TxtRename 
      Height          =   435
      Left            =   1215
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   4500
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton P1Stop 
      Height          =   375
      Left            =   1455
      Picture         =   "Prg01.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Detener"
      Top             =   3795
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton P1Pause 
      Height          =   375
      Left            =   855
      Picture         =   "Prg01.frx":0A3C
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Pausar"
      Top             =   3795
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton P1Save 
      Height          =   375
      Left            =   7065
      Picture         =   "Prg01.frx":1478
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Guardar Archivo de Programación"
      Top             =   3795
      UseMaskColor    =   -1  'True
      Width           =   420
   End
   Begin VB.CommandButton P1Open 
      Height          =   375
      Left            =   6660
      Picture         =   "Prg01.frx":157A
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Abrir archivo de Programación"
      Top             =   3795
      UseMaskColor    =   -1  'True
      Width           =   420
   End
   Begin VB.CommandButton P1Play 
      Height          =   375
      Left            =   135
      Picture         =   "Prg01.frx":167C
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Reproducir"
      Top             =   3795
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton P1New 
      Height          =   375
      Left            =   6255
      Picture         =   "Prg01.frx":20B8
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Nuevo archivo de Programación"
      Top             =   3795
      UseMaskColor    =   -1  'True
      Width           =   420
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   23
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3375
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   22
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3375
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   21
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3105
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   20
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3105
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   19
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2835
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   18
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2835
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   17
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2565
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   16
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2565
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   15
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2295
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   14
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2295
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   13
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2025
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   12
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2025
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   11
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1755
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   10
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1755
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   9
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1485
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   8
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1485
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   7
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1215
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   6
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1215
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   5
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   945
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   4
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   945
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   3
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   675
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   2
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   675
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   1
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   405
      Width           =   3450
   End
   Begin VB.CommandButton Prg1 
      Height          =   285
      Index           =   0
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   405
      Width           =   3450
   End
   Begin VB.Label Fn 
      BackColor       =   &H00808000&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   67
      Top             =   4530
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL Dur:"
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   4995
      TabIndex        =   65
      Top             =   75
      Width           =   915
   End
   Begin VB.Label LblName 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Programación 1 - Sin Nombre.prg"
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   90
      TabIndex        =   54
      ToolTipText     =   "Nombre de archivo"
      Top             =   60
      Width           =   4710
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   60
      TabIndex        =   66
      Top             =   30
      Width           =   7440
   End
   Begin VB.Label Lindex 
      BackColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   2520
      TabIndex        =   56
      Top             =   4515
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Shape E1Shape 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   45
      Top             =   3735
      Width           =   7500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "24"
      Height          =   240
      Index           =   23
      Left            =   3825
      TabIndex        =   53
      Top             =   3420
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "22"
      Height          =   240
      Index           =   22
      Left            =   3825
      TabIndex        =   52
      Top             =   3150
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "23"
      Height          =   240
      Index           =   21
      Left            =   45
      TabIndex        =   51
      Top             =   3420
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "21"
      Height          =   240
      Index           =   20
      Left            =   45
      TabIndex        =   50
      Top             =   3150
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "20"
      Height          =   240
      Index           =   19
      Left            =   3825
      TabIndex        =   49
      Top             =   2880
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "18"
      Height          =   240
      Index           =   18
      Left            =   3825
      TabIndex        =   48
      Top             =   2610
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "16"
      Height          =   240
      Index           =   17
      Left            =   3825
      TabIndex        =   47
      Top             =   2340
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "14"
      Height          =   240
      Index           =   16
      Left            =   3825
      TabIndex        =   46
      Top             =   2070
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "12"
      Height          =   240
      Index           =   15
      Left            =   3825
      TabIndex        =   45
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "10"
      Height          =   240
      Index           =   14
      Left            =   3825
      TabIndex        =   44
      Top             =   1530
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "8"
      Height          =   240
      Index           =   13
      Left            =   3825
      TabIndex        =   43
      Top             =   1260
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "6"
      Height          =   240
      Index           =   12
      Left            =   3825
      TabIndex        =   42
      Top             =   990
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "4"
      Height          =   240
      Index           =   11
      Left            =   3825
      TabIndex        =   41
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "2"
      Height          =   240
      Index           =   10
      Left            =   3825
      TabIndex        =   40
      Top             =   450
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "19"
      Height          =   240
      Index           =   9
      Left            =   45
      TabIndex        =   39
      Top             =   2880
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "17"
      Height          =   240
      Index           =   8
      Left            =   45
      TabIndex        =   38
      Top             =   2610
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "15"
      Height          =   240
      Index           =   7
      Left            =   45
      TabIndex        =   37
      Top             =   2340
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "13"
      Height          =   240
      Index           =   6
      Left            =   45
      TabIndex        =   36
      Top             =   2070
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "11"
      Height          =   240
      Index           =   5
      Left            =   45
      TabIndex        =   35
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "9"
      Height          =   240
      Index           =   4
      Left            =   45
      TabIndex        =   34
      Top             =   1530
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "7"
      Height          =   240
      Index           =   3
      Left            =   45
      TabIndex        =   33
      Top             =   1260
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "5"
      Height          =   240
      Index           =   2
      Left            =   45
      TabIndex        =   32
      Top             =   990
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "3"
      Height          =   240
      Index           =   1
      Left            =   45
      TabIndex        =   31
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   240
      Index           =   0
      Left            =   45
      TabIndex        =   30
      Top             =   450
      Width           =   195
   End
End
Attribute VB_Name = "Prg01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'dimensiones de conversion
Dim ConvertNm           'numeros
Dim ConvertNNm As Long
Dim ConvertTx As String   'textos
Dim ConvertTxT As String
Dim EstNum As Long

Dim FileExt As String

'Dimensiones de Archivos
Dim FileN As String
Dim FileNPath As String
Dim Completo As String
Dim SSTitle As String
Dim FileTP As String

'dimensiones de resultado
Dim Result As String
Dim RResult As String
Dim ResultInfo As Boolean

'Dimensiones de tiempo
Dim TimeNcv As String
Dim PosNcv As String

Sub DeployPRGFile(WContNum As Integer)

'PF=tanda filename
'PC=tanda name or caption
'PD=tanda duracion

Dim IDX As Integer
IDX = WContNum

If XPlorer.File1.filename = "" Or XPlorer.File1.filename = " " Then
    MsgBox LoadResString(137), vbCritical
    Exit Sub
End If

'.wav, .mp3, .it, .xm
FileExt = StripExtFromFile(XPlorer.File1.filename)
FileN = StripFileFromExt(XPlorer.File1.filename)
FileNPath = Right$(XPlorer.lblPath, Len(XPlorer.lblPath) - 2)
Completo = Right$(XPlorer.lblPath, Len(XPlorer.lblPath) - 2) & "\" & XPlorer.File1.filename

'seleccion de formato de archivo y extraccion de informacion header
Select Case Trim(UCase(FileExt))
    
    'STREAM TYPE WAV-MP1-MP2-MP3-OGG
    Case "WAV", "MP1", "MP2", "MP3", "OGG"
        MsgBox LoadResString(190), vbInformation, "Radio Maker"
        
    'MUSIC TYPE XM-MOD-S3M-IT-MTM-MO3-UMX
    Case "XM", "MOD", "S3M", "IT", "MTM", "MO3", "UMX"
        MsgBox LoadResString(190), vbInformation, "Radio Maker"
        
    'TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TND-TNDTND
    Case "TND"
        Est12Data.PF(IDX).Caption = Completo
        Est12Data.PC(IDX).Caption = FileN
        'Est12Data.PD(IdX).Caption = xxxxx duracion va aqui
        Prg1(IDX).Caption = FileN
        Prg1(IDX).BackColor = &H8000000F
        'Prg1(IdX).ToolTipText = "Duración: " & xxxxx duracion va aqui
        
    Case Else
        MsgBox LoadResString(190), vbInformation, "Radio Maker"

End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

HideWindow "Prg01"

End Sub

Private Sub Form_Terminate()

HideWindow "Prg01"

End Sub

Private Sub Form_Unload(Cancel As Integer)

HideWindow "Prg01"

End Sub

Private Sub P1New_Click()

Dim i As Integer

For i = 0 To 23
    Fn.Caption = ""
    Prg1(i).Caption = ""
    Prg1(i).ToolTipText = ""
    Est12Data.PF(i).Caption = ""
    Est12Data.PC(i).Caption = ""
    Est12Data.PD(i).Caption = ""
    LblName.Caption = "Programación 1 - Sin Nombre.prg"
    LblName.ForeColor = &H808000
Next i

End Sub

Private Sub P1Open_Click()

Dim lnText As Long
Dim NewName As String

On Error Resume Next
TopMenu.ProgCmd.InitDir = App.Path & AppProgDir & "\"
TopMenu.ProgCmd.Filter = "Programación de Tandas (*.prg)|*.prg|Programación de Tandas"
TopMenu.ProgCmd.DialogTitle = "Programación de Tandas - Abrir archivo"
TopMenu.ProgCmd.CancelError = True
TopMenu.ProgCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

    'restauramos los valores a 0
    Fn.Caption = ""
    Call RestoreDisplay(10)

ConvertTx = TopMenu.ProgCmd.filename

Result = OpenPrgFile(ConvertTx)
If Result = "NotOK" Then
    'MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
    Exit Sub
End If

Fn.Caption = ConvertTx
lnText = Len(ConvertTx)
If lnText > 60 Then
    NewName = Left$(ConvertTx, 3) & " ... " & Right$(ConvertTx, 50)
Else
    NewName = ConvertTx
End If
LblName.Caption = NewName
LblName.ForeColor = &HFFFF00    'verde claro

End Sub

Sub P1Save_Click()

Dim lnText As Long
Dim NewName As String

ConvertTxT = Trim(Fn.Caption)

On Error Resume Next
If ConvertTxT = "" Or ConvertTxT = " " Then
    TopMenu.ProgCmd.InitDir = App.Path & AppProgDir & "\"
    TopMenu.ProgCmd.Filter = "Programacion de Tandas (*.prg)|*.prg|Programación de Tandas"
    TopMenu.ProgCmd.DialogTitle = "Programacion de Tandas - Guardar archivo"
    TopMenu.ProgCmd.FilterIndex = 1
    TopMenu.ProgCmd.CancelError = True
    TopMenu.ProgCmd.ShowSave

    If err.Number = 32755 Then Exit Sub
    
    ConvertTx = TopMenu.ProgCmd.filename

    Fn.Caption = ConvertTx
    Result = SavePrgFile(ConvertTx)
    If Result = "NotOK" Then
        MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
        Exit Sub
    End If
Else
    ConvertTx = Trim(Fn.Caption)
    Kill ConvertTx
    Result = SavePrgFile(ConvertTx)
    If Result = "NotOK" Then
        'MsgBox "Ha ocurrido un Error. Operacion Abortada.", vbCritical
        Exit Sub
    End If
End If

Fn.Caption = ConvertTx
lnText = Len(ConvertTx)
If lnText > 60 Then
    NewName = Left$(ConvertTx, 3) & " ... " & Right$(ConvertTx, 50)
Else
    NewName = ConvertTx
End If

LblName.Caption = NewName
LblName.ForeColor = &HFFFF00    'verde claro

End Sub

Private Sub Prg1_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)

DeployPRGFile index   'drag & drop the selected file in xplorer

End Sub

Private Sub Prg1_DragOver(index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

Select Case State
    Case 0  'drag not finished
        XPlorer.File1.DragIcon = XPlorer.ExCombo.DragIcon
        Prg1(index).BackColor = &H80FF80    'verde (modificacion)
    Case 1  'finished drag
        XPlorer.File1.DragIcon = XPlorer.tvwDirTree.DragIcon
        Prg1(index).BackColor = &H8000000F  'gris (normal)
End Select

End Sub

Private Sub Prg1_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'button 1=left button
'button 2=right button
'button 4=mid button

If Button = 2 Then
    If Prg1(index).Caption = "" Or Prg1(index).Caption = " " Then
        Exit Sub
    End If
    'deploy options menu
    TxtRename.Visible = True
    TxtRename.top = Prg1(index).top
    TxtRename.Left = Prg1(index).Left
    TxtRename.Height = Prg1(index).Height
    TxtRename.Width = Prg1(index).Width
    TxtRename.Text = Prg1(index).Caption
    TxtRename.SetFocus
    'seteamos el label para saber de que control se trata
    Lindex.Caption = index
Else
    If Button = 4 Then
        'mark control to delete content
        Prg1(index).BackColor = &HFF&        'rojo
    Else
        'nothing to do
    End If
End If

End Sub

Private Sub TxtRename_KeyPress(KeyAscii As Integer)

Dim IDX As Integer

If KeyAscii = 13 Then   'ENTER
    IDX = CInt(Lindex.Caption)
    Prg1(IDX).Caption = TxtRename.Text
    Est12Data.PC(IDX).Caption = TxtRename.Text
    TxtRename.Visible = False
End If
If KeyAscii = 27 Or KeyAscii = 13 Then 'ESCAPE or ENTER
    TxtRename.Visible = False
End If

End Sub
