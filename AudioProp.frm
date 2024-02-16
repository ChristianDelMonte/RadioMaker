VERSION 5.00
Begin VB.Form AudioProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Propiedades de Audio - TANDA"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8010
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdRmv1 
      Caption         =   "R"
      Height          =   285
      Left            =   180
      TabIndex        =   44
      ToolTipText     =   "Remover"
      Top             =   720
      Width           =   555
   End
   Begin VB.CommandButton CmdActive 
      Caption         =   "AL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   11
      ToolTipText     =   "Activar hora de lanzamiento"
      Top             =   4380
      Width           =   2445
   End
   Begin VB.CommandButton CmdAcept 
      Caption         =   "Ac"
      Height          =   375
      Left            =   5355
      TabIndex        =   12
      ToolTipText     =   "Aceptar y guardar los cambios"
      Top             =   4380
      Width           =   1230
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cc"
      Height          =   375
      Left            =   6690
      TabIndex        =   13
      ToolTipText     =   "Cancelar"
      Top             =   4380
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Propiedades del archivo de audio"
      ForeColor       =   &H00000000&
      Height          =   3120
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   7845
      Begin VB.CommandButton CmdRmv2 
         Caption         =   "R"
         Height          =   285
         Left            =   90
         TabIndex        =   43
         ToolTipText     =   "Remover"
         Top             =   2025
         Width           =   555
      End
      Begin VB.TextBox TxtMLanz 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   5670
         MaxLength       =   5
         TabIndex        =   10
         Text            =   "00:00"
         Top             =   2610
         Width           =   735
      End
      Begin VB.TextBox TxtFlanz 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   5670
         MaxLength       =   8
         TabIndex        =   5
         Text            =   "00:00:00"
         Top             =   1170
         Width           =   960
      End
      Begin VB.PictureBox T1M5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   5310
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2655
         Width           =   190
      End
      Begin VB.TextBox TxtMName 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   9
         Top             =   2610
         Width           =   4245
      End
      Begin VB.PictureBox T1M4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   5115
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2655
         Width           =   190
      End
      Begin VB.PictureBox T1M3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4920
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2655
         Width           =   190
      End
      Begin VB.PictureBox T1M2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4740
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   27
         Top             =   2655
         Width           =   190
      End
      Begin VB.PictureBox T1M1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4545
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2655
         Width           =   190
      End
      Begin VB.PictureBox T1T5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   5310
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1215
         Width           =   190
      End
      Begin VB.CommandButton ExamM 
         Caption         =   "E"
         Height          =   330
         Left            =   5445
         TabIndex        =   7
         ToolTipText     =   "Examinar"
         Top             =   1980
         Width           =   870
      End
      Begin VB.TextBox TxtMFile 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2025
         Width           =   4695
      End
      Begin VB.TextBox TxtMType 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Stream"
         Top             =   2025
         Width           =   1140
      End
      Begin VB.CommandButton ExamF 
         Caption         =   "E"
         Height          =   330
         Left            =   5445
         TabIndex        =   2
         ToolTipText     =   "Examinar"
         Top             =   540
         Width           =   870
      End
      Begin VB.TextBox TxtFType 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Stream"
         Top             =   585
         Width           =   1140
      End
      Begin VB.PictureBox T1T1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4545
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1215
         Width           =   190
      End
      Begin VB.PictureBox T1T2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4740
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1215
         Width           =   190
      End
      Begin VB.PictureBox T1T3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4920
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1215
         Width           =   190
      End
      Begin VB.PictureBox T1T4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   5115
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1215
         Width           =   190
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4500
         ScaleHeight     =   285
         ScaleWidth      =   1050
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1050
      End
      Begin VB.TextBox TxtName 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   4
         Top             =   1170
         Width           =   4245
      End
      Begin VB.TextBox TxtFile 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   4695
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4500
         ScaleHeight     =   285
         ScaleWidth      =   1050
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2610
         Width           =   1050
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   135
         X2              =   7695
         Y1              =   1665
         Y2              =   1665
      End
      Begin VB.Label Label13 
         Caption         =   "mm:ss"
         Height          =   240
         Left            =   6480
         TabIndex        =   38
         Top             =   2655
         Width           =   510
      End
      Begin VB.Label Label12 
         Caption         =   "hh:mm:ss"
         Height          =   240
         Left            =   6705
         TabIndex        =   37
         Top             =   1215
         Width           =   780
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Hora de Lanz.:"
         Height          =   240
         Left            =   5670
         TabIndex        =   35
         Top             =   2385
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Hora de Lanz.:"
         Height          =   240
         Left            =   5670
         TabIndex        =   34
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del mixer:"
         Height          =   240
         Left            =   135
         TabIndex        =   32
         Top             =   2385
         Width           =   1320
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Duración:"
         Height          =   240
         Left            =   4500
         TabIndex        =   31
         Top             =   2385
         Width           =   780
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Path y nombre de archivo mixer:"
         Height          =   240
         Left            =   135
         TabIndex        =   24
         Top             =   1800
         Width           =   2310
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   6570
         TabIndex        =   23
         Top             =   1800
         Width           =   510
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   6570
         TabIndex        =   22
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Duración:"
         Height          =   240
         Left            =   4500
         TabIndex        =   16
         Top             =   945
         Width           =   780
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del tema:"
         Height          =   240
         Left            =   135
         TabIndex        =   15
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Path y nombre de archivo:"
         Height          =   240
         Left            =   135
         TabIndex        =   14
         Top             =   360
         Width           =   1995
      End
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   7905
      X2              =   75
      Y1              =   4260
      Y2              =   4260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   7905
      X2              =   75
      Y1              =   4245
      Y2              =   4245
   End
   Begin VB.Label TxtMDur 
      Caption         =   "00:00:00"
      Height          =   240
      Left            =   4230
      TabIndex        =   42
      Top             =   4590
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label TxtFdur 
      Caption         =   "00:00:00"
      Height          =   240
      Left            =   3375
      TabIndex        =   41
      Top             =   4590
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label LKey 
      Caption         =   "0"
      Height          =   240
      Left            =   3060
      TabIndex        =   40
      Top             =   4590
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Lidx 
      Caption         =   "0"
      Height          =   240
      Left            =   2790
      TabIndex        =   39
      Top             =   4590
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   $"AudioProp.frx":0000
      Height          =   780
      Left            =   90
      TabIndex        =   36
      Top             =   3375
      Width           =   7845
   End
End
Attribute VB_Name = "AudioProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAcept_Click()

Dim DataA(0 To 9) As String
Dim DataKa As String
Dim nIndex As Integer
Dim Response
Dim ItmX As ListItem

nIndex = CInt(Trim(Lidx.Caption))                   'numero de index
DataKa = Trim(Tanda01.T1View.SelectedItem.Key)      'key

'seleccionamos el item
Tanda01.T1View.ListItems.Item(nIndex).Selected = True

'seteamos los nuevos datos
'HOST FILE
DataA(0) = Trim(TxtFile.Text)    'file & path
DataA(1) = Trim(TxtFType.Text)     'filetype
DataA(2) = Trim(TxtName.Text)     'filename
DataA(3) = Trim(TxtFdur.Caption)     'duracion
DataA(4) = Trim(TxtFlanz.Text)     'hora de lanz
'MIXER FILE
DataA(5) = Trim(TxtMFile.Text)     'file & path
DataA(6) = Trim(TxtMType.Text)     'filetype
DataA(7) = Trim(TxtMName.Text)     'filename
DataA(8) = Trim(TxtMDur.Caption)     'duracion
DataA(9) = Trim(TxtMLanz.Text)     'hora de lanz

'chequeos necesarios
If DataA(0) = "-----" Then
    Response = MsgBox(LoadResString(185) & " " & LoadResString(132), vbYesNo, "RM100 Propiedad de Audio")
    If Response = vbYes Then
        'removemos los viejos datos
        Tanda01.T1View.ListItems.Remove (nIndex)
        Unload Me
        Exit Sub
    Else
        'nothing to do
        Exit Sub
    End If
Else
    'removemos los viejos datos
    Tanda01.T1View.ListItems.Remove (nIndex)
End If

'ponemos los nuevos datos
Set ItmX = Tanda01.T1View.ListItems.Add(nIndex, DataKa, DataA(0)) 'path & file
ItmX.SubItems(1) = DataA(1)
ItmX.SubItems(2) = DataA(2)
ItmX.SubItems(3) = DataA(3)
ItmX.SubItems(4) = DataA(4)
ItmX.SubItems(5) = DataA(5)
ItmX.SubItems(6) = DataA(6)
ItmX.SubItems(7) = DataA(7)
ItmX.SubItems(8) = DataA(8)
ItmX.SubItems(9) = DataA(9)

'una vez finalizado. seleccionamos el item
Tanda01.T1View.ListItems.Item(nIndex).Selected = True
'y... actualizamos las horas de lanzamiento
Call Tanda01.T1OrderA_Click
Unload Me
Exit Sub

Continue:
    'nothing to do....
End Sub

Private Sub cmdCancel_Click()

Unload Me

End Sub

Private Sub CmdRmv1_Click()

TxtFile.Text = "-----"
    TxtFile.BackColor = &H80000005
TxtName.Text = "-----"
    TxtName.BackColor = &H80000005
TxtFType.Text = "-----"
    TxtFType.BackColor = &H80000005
TxtFlanz.Text = "00:00"
    TxtFlanz.BackColor = &H80000005
Call RestoreDisplay(8)
TxtFdur.Caption = "00:00"

End Sub

Private Sub CmdRmv2_Click()

TxtMFile.Text = "-----"
    TxtMFile.BackColor = &H80000005
TxtMName.Text = "-----"
    TxtMName.BackColor = &H80000005
TxtMType.Text = "-----"
    TxtMType.BackColor = &H80000005
TxtMLanz.Text = "00:00"
    TxtMLanz.BackColor = &H80000005
Call RestoreDisplay(9)
TxtMDur.Caption = "00:00"

End Sub

Private Sub ExamF_Click()

Dim i As Integer, X As Integer, Z As Integer
Dim FileExt As String, FileN As String, Completo As String
Dim ConvertTx As String, TimeNcv As String, Result As String

On Error Resume Next
TopMenu.WaveCmd.InitDir = App.Path & AppDefaultMusicPath
TopMenu.WaveCmd.Filter = "Archivos de Audio (*.wav; *.mp1; *.mp2; *.mp3)|*.wav; *.mp1; *.mp2; *.mp3|Todos los archivos de Audio"
TopMenu.WaveCmd.DialogTitle = "Propiedades de Audio - Abrir archivo de Audio"
TopMenu.WaveCmd.CancelError = True
TopMenu.WaveCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

'extract the file type - .wav, .mp3, .it, .xm
FileExt = StripExtFromFile(TopMenu.WaveCmd.filename)

'extract the file name, with out path
FileN = StripFileFromDir(TopMenu.WaveCmd.filename)
FileN = StripFileFromExt(FileN)
Completo = TopMenu.WaveCmd.filename

'put the file info
TxtFile.Text = Completo
    TxtFile.BackColor = &HC0FFFF
TxtName = FileN
    TxtName.BackColor = &HC0FFFF
TxtFType.Text = "Stream"
    TxtFType.BackColor = &HC0FFFF
TxtFlanz.Text = "00:00:00"
    TxtFlanz.BackColor = &HC0FFFF
    
'put the file time
ConvertTx = FileLoadLen(Completo, "Stream")
TimeNcv = FormatSegs(ConvertTx)
Result = ConvSecToMin(CInt(TimeNcv))

'put the time display of file
TxtFdur.Caption = Result: SetAudioTime "1", Result

End Sub

Private Sub ExamM_Click()

Dim i As Integer, X As Integer
Dim Z As Integer, TxtS As String, TxtM As String
Dim IntS As Integer, IntM As Integer
Dim FileExt As String, FileN As String, Completo As String
Dim ConvertTx As String
Dim TimeNcv As String
Dim Result As String

On Error Resume Next
TopMenu.WaveCmd.InitDir = App.Path & AppDefaultMusicPath
TopMenu.WaveCmd.Filter = "Archivos de Audio (*.wav; *.mp1; *.mp2; *.mp3)|*.wav; *.mp1; *.mp2; *.mp3|Todos los archivos de Audio"
TopMenu.WaveCmd.DialogTitle = "Propiedades de Audio - Abrir archivo de Audio"
TopMenu.WaveCmd.CancelError = True
TopMenu.WaveCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

'extract the file type - .wav, .mp3, .it, .xm
FileExt = StripExtFromFile(TopMenu.WaveCmd.filename)

'extract the file name, with out path
FileN = StripFileFromDir(TopMenu.WaveCmd.filename)
FileN = StripFileFromExt(FileN)
Completo = TopMenu.WaveCmd.filename

'put the file time
ConvertTx = FileLoadLen(Completo, "Stream")
TimeNcv = FormatSegs(ConvertTx)
Result = ConvSecToMin(CInt(TimeNcv))

'check if the file is ok for deploy
TxtM = Left$(Result, 2)
TxtS = Right$(Result, 2)
IntM = CInt(TxtM)
IntS = CInt(TxtS)
If IntM >= 1 Then
    MsgBox LoadResString(186), vbInformation, "RM100 - Propiedades de Audio"
    Exit Sub
Else
    If IntS >= 60 Then
        MsgBox LoadResString(186), vbInformation, "RM100 - Propiedades de Audio"
        Exit Sub
    Else
        'nothing to do, all ok.
    End If
End If

'put the time display of file
TxtMDur.Caption = Result: SetAudioTime "2", Result
'put the file info
TxtMFile.Text = Completo
    TxtMFile.BackColor = &HC0FFFF
TxtMName.Text = FileN
    TxtMName.BackColor = &HC0FFFF
TxtMType.Text = "Stream"
    TxtMType.BackColor = &HC0FFFF
TxtMLanz.Text = "00:00"
    TxtMLanz.BackColor = &HC0FFFF
End Sub

Private Sub Form_Load()

'cargamos los strings
Me.Caption = LoadResString(2021)
Frame1.Caption = LoadResString(2013)
Label1.Caption = LoadResString(2014)
Label4.Caption = LoadResString(2015)
Label2.Caption = LoadResString(2016)
Label3.Caption = LoadResString(2017)
Label9.Caption = LoadResString(2018)
Label6.Caption = LoadResString(2019)
Label5.Caption = LoadResString(2015)
Label8.Caption = LoadResString(2020)
Label7.Caption = LoadResString(2017)
Label10.Caption = LoadResString(2018)
CmdAcept.Caption = LoadResString(2000)
cmdCancel.Caption = LoadResString(2001)
ExamF.Caption = LoadResString(2002)
ExamM.Caption = LoadResString(2002)
CmdActive.Caption = LoadResString(2008)
CmdRmv1.Caption = LoadResString(2009)
CmdRmv2.Caption = LoadResString(2009)

Dim DataA(0 To 9) As String

'extraemos los datos del item
'HOST FILE
DataA(0) = Tanda01.T1View.SelectedItem.Text    'file & path
'DataA(1) = Tanda01.T1View.SelectedItem.ListSubItems(1).Text     'filetype
'DataA(2) = Tanda01.T1View.SelectedItem.ListSubItems(2).Text     'filename
'DataA(3) = Tanda01.T1View.SelectedItem.ListSubItems(3).Text     'duracion
'DataA(4) = Tanda01.T1View.SelectedItem.ListSubItems(4).Text     'hora de lanz
'MIXER FILE
'DataA(5) = Tanda01.T1View.SelectedItem.ListSubItems(5).Text     'file & path
'DataA(6) = Tanda01.T1View.SelectedItem.ListSubItems(6).Text     'filetype
'DataA(7) = Tanda01.T1View.SelectedItem.ListSubItems(7).Text     'filename
'DataA(8) = Tanda01.T1View.SelectedItem.ListSubItems(8).Text     'duracion
'DataA(9) = Tanda01.T1View.SelectedItem.ListSubItems(9).Text     'hora de lanz

Lidx.Caption = Tanda01.T1View.SelectedItem.Index
LKey.Caption = Tanda01.T1View.SelectedItem.Key

'sets the dur display
Call RestoreDisplay(8)
Call RestoreDisplay(9)

'put the File data
TxtFile.Text = DataA(0)
TxtFType.Text = DataA(1)
TxtName.Text = DataA(2)
TxtFdur.Caption = Trim(DataA(3))
If TxtFdur.Caption = "00:00:00" Then
    'nothing to do
Else
    SetAudioTime "1", Trim(Right$(DataA(3), 5))
End If
TxtFlanz.Text = DataA(4)
    
'put the mix data
TxtMFile.Text = DataA(5)
TxtMType.Text = DataA(6)
TxtMName.Text = DataA(7)
TxtMDur.Caption = Trim(DataA(8))
If TxtMDur.Caption = "00:00:00" Then
    'nothing to do
Else
    SetAudioTime "2", Trim(Right$(DataA(8), 5))
End If
TxtMLanz.Text = DataA(9)

If TxtMFile.Text = "-----" Then
    TxtMFile.BackColor = &HFFFFFF
    TxtMType.BackColor = &HFFFFFF
    TxtMName.BackColor = &HFFFFFF
    TxtMLanz.BackColor = &HFFFFFF
End If

End Sub

Private Sub TxtFlanz_Change()

If TxtFlanz.Text = "" Or TxtFlanz.Text = " " Then
    TxtFlanz.Text = "00:00:00"
End If

End Sub

Private Sub TxtFlanz_GotFocus()

TxtFlanz.SelStart = 0
TxtFlanz.SelLength = Len(TxtFlanz.Text)

End Sub


Private Sub TxtFlanz_LostFocus()

Dim LenCheck
LenCheck = Len(TxtFlanz.Text)

'check the len for validations
If LenCheck < 8 Then
    MsgBox LoadResString(187), vbInformation
    TxtFlanz.SetFocus
    Exit Sub
End If
If LenCheck > 8 Then
    MsgBox LoadResString(187), vbInformation
    TxtFlanz.SetFocus
    Exit Sub
End If

Dim Hora As String, Minutos As String, Segundos As String

'extraemos los datos de hora especificados para el lanzamiento
Hora = Left$(TxtFlanz.Text, 2)
Minutos = Mid$(TxtFlanz.Text, 4, 2)
Segundos = Right$(TxtFlanz.Text, 2)

'Procedemos al chequeo de la misma
On Error Resume Next
If LenCheck = 8 Then
    If Hora > 23 Or Hora < 0 Then
        If Hora = "00" Then
            'xxxx
        Else
            MsgBox LoadResString(187), vbInformation
            TxtFlanz.SetFocus
            Exit Sub
        End If
    Else
        'la hora esta bien. chequeamos los minutos.
        If Minutos > 59 Then
            If Minutos = "00" Then
                'xx
            Else
                MsgBox LoadResString(188), vbInformation
                TxtFlanz.SetFocus
                Exit Sub
            End If
        Else
            If Segundos > 59 Then
                If Segundos = "00" Then
                    'xxx
                Else
                    MsgBox LoadResString(189), vbInformation
                    TxtFlanz.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
Else
    MsgBox LoadResString(187), vbInformation
    TxtFlanz.SetFocus
    Exit Sub
End If

End Sub

Private Sub TxtMLanz_Change()

If TxtMLanz.Text = "" Or TxtMLanz.Text = " " Then
    TxtMLanz.Text = "00:00"
End If

End Sub

Private Sub TxtMLanz_GotFocus()

TxtMLanz.SelStart = 0
TxtMLanz.SelLength = Len(TxtMLanz.Text)

End Sub

Private Sub TxtMLanz_LostFocus()

Dim LenCheck
LenCheck = Len(TxtMLanz.Text)

'check the len for validations
If LenCheck < 5 Then
    MsgBox LoadResString(187), vbInformation
    TxtMLanz.SetFocus
    Exit Sub
End If
If LenCheck > 5 Then
    MsgBox LoadResString(187), vbInformation
    TxtMLanz.SetFocus
    Exit Sub
End If

Dim Hora As String, Minutos As String

'extraemos los datos de hora especificados para el lanzamiento
Hora = Left$(TxtMLanz.Text, 2)
Minutos = Right$(TxtMLanz.Text, 2)

'Procedemos al chequeo de la misma
On Error Resume Next
If LenCheck = 5 Then
    If Hora > 23 Or Hora < 0 Then
        If Hora = "00" Then
            'xxxx
        Else
            MsgBox LoadResString(187), vbInformation
            TxtMLanz.SetFocus
            Exit Sub
        End If
    Else
        'la hora esta bien. chequeamos los minutos.
        If Minutos > 59 Then
            If Minutos = "00" Then
                'xx
            Else
                MsgBox LoadResString(188), vbInformation
                TxtMLanz.SetFocus
                Exit Sub
            End If
        Else
            'xxx
        End If
    End If
Else
    MsgBox LoadResString(187), vbInformation
    TxtMLanz.SetFocus
    Exit Sub
End If

End Sub
