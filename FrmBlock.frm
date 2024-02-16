VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form FrmBlock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bloques de Publicidad"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11340
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ViewBlock 
      Height          =   3975
      Left            =   2760
      TabIndex        =   46
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   7011
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ProgressBar PrgBar1 
      Height          =   285
      Left            =   4740
      TabIndex        =   45
      Top             =   5160
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton BLOpen 
      Height          =   375
      Left            =   570
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Abrir archivo de bloque"
      Top             =   5175
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton BLNew 
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Nuevo archivo de bloque"
      Top             =   5175
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones de AIRE"
      Height          =   4065
      Left            =   5970
      TabIndex        =   8
      Top             =   510
      Width           =   5265
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   225
         Index           =   2
         Left            =   4980
         Max             =   1
         Min             =   10
         TabIndex        =   31
         Top             =   3255
         Value           =   1
         Width           =   165
      End
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   4980
         Max             =   1
         Min             =   10
         TabIndex        =   30
         Top             =   2895
         Value           =   1
         Width           =   165
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   225
         Index           =   0
         Left            =   4980
         Max             =   1
         Min             =   10
         TabIndex        =   29
         Top             =   2535
         Value           =   1
         Width           =   165
      End
      Begin VB.CommandButton CmdH 
         Caption         =   "Set >>"
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   2745
         TabIndex        =   28
         Top             =   3225
         Width           =   630
      End
      Begin VB.CommandButton CmdH 
         Caption         =   "Set >>"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   2745
         TabIndex        =   27
         Top             =   2865
         Width           =   630
      End
      Begin VB.CommandButton CmdD 
         Caption         =   "Set >>"
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   2745
         TabIndex        =   26
         Top             =   1875
         Width           =   630
      End
      Begin VB.CommandButton CmdD 
         Caption         =   "Set >>"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   2745
         TabIndex        =   25
         Top             =   1515
         Width           =   630
      End
      Begin VB.TextBox PFH 
         Height          =   285
         Index           =   2
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   3225
         Width           =   1125
      End
      Begin VB.TextBox PFH 
         Height          =   285
         Index           =   1
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   2865
         Width           =   1125
      End
      Begin VB.TextBox PFD 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1875
         Width           =   1125
      End
      Begin VB.TextBox PFD 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1515
         Width           =   1125
      End
      Begin VB.CommandButton CmdD 
         Caption         =   "Set >>"
         Height          =   270
         Index           =   0
         Left            =   2745
         TabIndex        =   20
         Top             =   1155
         Width           =   630
      End
      Begin VB.CommandButton CmdH 
         Caption         =   "Set >>"
         Height          =   270
         Index           =   0
         Left            =   2745
         TabIndex        =   19
         Top             =   2505
         Width           =   630
      End
      Begin VB.TextBox PFD 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1155
         Width           =   1125
      End
      Begin VB.TextBox PFH 
         Height          =   285
         Index           =   0
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2505
         Width           =   1125
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   270
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   1140
         Width           =   2340
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   270
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   2490
         Width           =   2340
      End
      Begin VB.TextBox FInit 
         Height          =   285
         Left            =   780
         TabIndex        =   10
         Text            =   "08-01-2002"
         Top             =   480
         Width           =   1005
      End
      Begin VB.TextBox FEnd 
         Height          =   285
         Left            =   3480
         TabIndex        =   9
         Text            =   "08-01-2002"
         Top             =   480
         Width           =   1005
      End
      Begin VB.TextBox TxtCant 
         Height          =   285
         Index           =   0
         Left            =   4620
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "1"
         Top             =   2505
         Width           =   555
      End
      Begin VB.TextBox TxtCant 
         Height          =   285
         Index           =   2
         Left            =   4620
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "1"
         Top             =   3225
         Width           =   555
      End
      Begin VB.TextBox TxtCant 
         Height          =   285
         Index           =   1
         Left            =   4620
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "1"
         Top             =   2865
         Width           =   555
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "hh/mm/ss"
         Height          =   210
         Left            =   4290
         TabIndex        =   2
         Top             =   3660
         Width           =   750
      End
      Begin VB.Label LblDur 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "00:00:00"
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Left            =   3480
         TabIndex        =   44
         Top             =   3660
         Width           =   750
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Duración:"
         Height          =   210
         Left            =   2760
         TabIndex        =   43
         Top             =   3660
         Width           =   750
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Veces:"
         Height          =   195
         Left            =   4620
         TabIndex        =   37
         Top             =   2295
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hora preferencial de salida:"
         Height          =   210
         Left            =   285
         TabIndex        =   36
         Top             =   2280
         Width           =   2010
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Día preferencial de salida:"
         Height          =   225
         Left            =   285
         TabIndex        =   35
         Top             =   900
         Width           =   1980
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio:"
         Height          =   225
         Left            =   300
         TabIndex        =   14
         Top             =   510
         Width           =   525
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Finalización:"
         Height          =   225
         Left            =   2550
         TabIndex        =   13
         Top             =   510
         Width           =   945
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "dd/mm/aaaa"
         Height          =   225
         Left            =   780
         TabIndex        =   12
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "dd/mm/aaaa"
         Height          =   225
         Left            =   3480
         TabIndex        =   11
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.CommandButton BLSave 
      Height          =   375
      Left            =   1020
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Guardar cambios"
      Top             =   5175
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.DirListBox Dir1 
      Height          =   3915
      Left            =   135
      TabIndex        =   5
      Top             =   600
      Width           =   2550
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   135
      TabIndex        =   4
      Top             =   165
      Width           =   2565
   End
   Begin VB.CommandButton CmdAccept 
      Caption         =   "Ac"
      Height          =   390
      Left            =   8970
      TabIndex        =   1
      Top             =   5175
      Width           =   1065
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cc"
      Height          =   390
      Left            =   10155
      TabIndex        =   0
      Top             =   5175
      Width           =   1065
   End
   Begin VB.Label LblID 
      BackColor       =   &H0000FF00&
      Caption         =   "0"
      Height          =   195
      Left            =   3600
      TabIndex        =   42
      Top             =   5220
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LblName 
      BackColor       =   &H000080FF&
      Height          =   195
      Left            =   2070
      TabIndex        =   41
      Top             =   5220
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Archivo: // Sin nombre.blk \\"
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   5970
      TabIndex        =   38
      Top             =   180
      Width           =   5265
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "LISTA de archivos publicitarios:"
      Height          =   210
      Left            =   2820
      TabIndex        =   6
      Top             =   285
      Width           =   2325
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   11190
      X2              =   105
      Y1              =   5070
      Y2              =   5070
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   11190
      X2              =   105
      Y1              =   5055
      Y2              =   5055
   End
   Begin VB.Label LbLProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxx"
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   4650
      Width           =   2565
   End
End
Attribute VB_Name = "FrmBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ADDdata()

'////////////////////////////////////
'add the data to the lists
'***********************************

Combo1.text = "Cualquier horario"
Combo1.AddItem "Cualquier horario", 0
Combo1.AddItem "1 a 2 hs", 1
Combo1.AddItem "2 a 3 hs", 2
Combo1.AddItem "3 a 4 hs", 3
Combo1.AddItem "4 a 5 hs", 4
Combo1.AddItem "5 a 6 hs", 5
Combo1.AddItem "6 a 7 hs", 6
Combo1.AddItem "7 a 8 hs", 7
Combo1.AddItem "8 a 9 hs", 8
Combo1.AddItem "9 a 10 hs", 9
Combo1.AddItem "10 a 11 hs", 10
Combo1.AddItem "11 a 12 hs", 11
Combo1.AddItem "12 a 13 hs", 12
Combo1.AddItem "13 a 14 hs", 13
Combo1.AddItem "14 a 15 hs", 14
Combo1.AddItem "15 a 16 hs", 15
Combo1.AddItem "16 a 17 hs", 16
Combo1.AddItem "17 a 18 hs", 17
Combo1.AddItem "18 a 19 hs", 18
Combo1.AddItem "19 a 20 hs", 19
Combo1.AddItem "20 a 21 hs", 20
Combo1.AddItem "21 a 22 hs", 21
Combo1.AddItem "22 a 23 hs", 22
Combo1.AddItem "23 a 00 hs", 23
Combo1.AddItem "00 a 1 hs", 24

Combo2.text = "Todos los días"
Combo2.AddItem "Todos los días", 0
Combo2.AddItem "Domingo", 1
Combo2.AddItem "Lunes", 2
Combo2.AddItem "Martes", 3
Combo2.AddItem "Miercoles", 4
Combo2.AddItem "Jueves", 5
Combo2.AddItem "Viernes", 6
Combo2.AddItem "Sábado", 7

End Sub

Private Sub PutData(WData As BlockRecord)

Dim i As Integer

'/// get the correct data before save
For i = 0 To 2      '///// extract the user selected data (DAY)
    Select Case WData.FPrefD(i)
        Case BlockPrefD.Dom
            PFD(i).text = "Domingo"
            PFD(i).BackColor = &HC0FFFF
            CmdD(i).Enabled = True
        Case BlockPrefD.Lun
            PFD(i).text = "Lunes"
            PFD(i).BackColor = &HC0FFFF
            CmdD(i).Enabled = True
        Case BlockPrefD.Mar
            PFD(i).text = "Martes"
            PFD(i).BackColor = &HC0FFFF
            CmdD(i).Enabled = True
        Case BlockPrefD.Mie
            PFD(i).text = "Miercoles"
            PFD(i).BackColor = &HC0FFFF
            CmdD(i).Enabled = True
        Case BlockPrefD.Jue
            PFD(i).text = "Jueves"
            PFD(i).BackColor = &HC0FFFF
            CmdD(i).Enabled = True
        Case BlockPrefD.Vie
            PFD(i).text = "Viernes"
            PFD(i).BackColor = &HC0FFFF
            CmdD(i).Enabled = True
        Case BlockPrefD.Sab
            PFD(i).text = "Sábado"
            PFD(i).BackColor = &HC0FFFF
            CmdD(i).Enabled = True
        Case BlockPrefD.All
            PFD(i).text = "Todos los días"
            PFD(i).BackColor = &HC0FFFF
            CmdD(i).Enabled = True
        Case BlockPrefD.Vacio
            PFD(i).text = ""
            PFD(i).BackColor = &HFFFFFF
            CmdD(i).Enabled = False
    End Select
Next i

For i = 0 To 2      '///// extract the user selected data (HOUR)
    Select Case WData.FPrefH(i)
        Case BlockPrefH.d1a2
            PFH(i).text = "1 a 2 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d2a3
            PFH(i).text = "2 a 3 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d3a4
            PFH(i).text = "3 a 4 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d4a5
            PFH(i).text = "4 a 5 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d5a6
            PFH(i).text = "5 a 6 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d6a7
            PFH(i).text = "6 a 7 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d7a8
            PFH(i).text = "7 a 8 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d8a9
            PFH(i).text = "8 a 9 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d9a10
            PFH(i).text = "9 a 10 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d10a11
            PFH(i).text = "10 a 11 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d11a12
            PFH(i).text = "11 a 12 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d12a13
            PFH(i).text = "12 a 13 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d13a14
            PFH(i).text = "13 a 14 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d14a15
            PFH(i).text = "14 a 15 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d15a16
            PFH(i).text = "15 a 16 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d16a17
            PFH(i).text = "16 a 17 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d17a18
            PFH(i).text = "17 a 18 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d18a19
            PFH(i).text = "18 a 19 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d19a20
            PFH(i).text = "19 a 20 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d20a21
            PFH(i).text = "20 a 21 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d21a22
            PFH(i).text = "21 a 22 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d22a23
            PFH(i).text = "22 a 23 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d23a0
            PFH(i).text = "23 a 00 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.d0a1
            PFH(i).text = "00 a 1 hs"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.All
            PFH(i).text = "Cualquier horario"
            PFH(i).BackColor = &HC0FFFF
            CmdH(i).Enabled = True
        Case BlockPrefH.Vacio
            PFH(i).text = ""
            PFH(i).BackColor = &HFFFFFF
            CmdH(i).Enabled = False
    End Select
Next i

For i = 0 To 2      '///// extract the reproducer cant
    If WData.FCantV(i) = 0 Then
        TxtCant(i).text = WData.FCantV(i) + 1
        VScroll1(i).Value = WData.FCantV(i) + 1
    Else
        TxtCant(i).text = WData.FCantV(i)
        VScroll1(i).Value = WData.FCantV(i)
    End If
Next i

FInit.text = Trim(WData.FPubInit)
FEnd.text = Trim(WData.FPubFin)
LblDur.Caption = Trim(WData.FFileDur)

End Sub

Private Sub RestoreData()

'/////////////////////////////////////
'restore the data to default values
'************************************

PFH(0).text = "": PFH(0).BackColor = &HFFFFFF
PFH(1).text = "": PFH(1).BackColor = &HFFFFFF
PFH(2).text = "": PFH(2).BackColor = &HFFFFFF
CmdH(1).Enabled = False: CmdH(2).Enabled = False

PFD(0).text = "": PFD(0).BackColor = &HFFFFFF
PFD(1).text = "": PFD(1).BackColor = &HFFFFFF
PFD(2).text = "": PFD(2).BackColor = &HFFFFFF
CmdD(1).Enabled = False: CmdD(2).Enabled = False

TxtCant(0) = "1": TxtCant(1) = "1": TxtCant(2) = "1"
VScroll1(0).Value = 1: VScroll1(1).Value = 1: VScroll1(2).Value = 1
FInit.text = Date: FEnd.text = Date ': LblDur.Caption = "00:00:00"

End Sub

Private Function CheckDateVal(WDate As String) As Boolean

Dim LnDate As Long
Dim DD As String, MM As String, YY As String
Dim CurrYY As String

'/// get the current data for check validity
WDate = Trim(WDate)
LnDate = Len(WDate)

DD = Left$(WDate, 2)        'user day passed
MM = Mid$(WDate, 4, 2)      'user month passed
YY = Right$(WDate, 4)       'user year passed
CurrYY = Right$(Date, 4)    'current system year

'/// check the validity
If LnDate < 10 Then
    CheckDateVal = False
    Exit Function
Else
    If CLng(DD) > 31 Then
        CheckDateVal = False
        Exit Function
    Else
        If CLng(MM) > 12 Then
            CheckDateVal = False
            Exit Function
        Else
            If CLng(YY) < CLng(CurrYY) Then
                CheckDateVal = False
                Exit Function
            Else
                CheckDateVal = True
                Exit Function
            End If
        End If
    End If
End If

End Function

Private Function ReloadDir(WLocalPath As String, WLocalExt As String)

Dim FNum As Long, FTotalNum As Long
Dim i As Integer
Dim FKey As String
Dim FName As String
Dim ItmX As ListItem
Dim Result As INFODir

Set ItmX = Nothing
ViewBlock.ListItems.Clear

'cargamos los datos del directorio maestro
Result = GETInfoDir(WLocalPath, WLocalExt)
FTotalNum = Result.PFilesNum

If FTotalNum < 1 Then GoSub Finalizar

On Error GoTo Finish

FNum = 1
FName = Dir(WLocalPath & WLocalExt)
FKey = "f" & FNum
Set ItmX = ViewBlock.ListItems.Add(FNum, FKey, WLocalPath)
    ItmX.SubItems(1) = FName

PrgBar1.Visible = True
PrgBar1.Min = 1
PrgBar1.Max = FTotalNum + 1
LbLProgress.Caption = "Cargando...."

For i = 2 To FTotalNum
    FName = Dir
    FNum = i
    FKey = "f" & FNum
    Set ItmX = ViewBlock.ListItems.Add(FNum, FKey, WLocalPath)
        ItmX.SubItems(1) = FName
    PrgBar1.Value = i
    DoEvents
Next i

Finish:
PrgBar1.Value = 1
PrgBar1.Visible = False
LbLProgress.Caption = ""

Finalizar:
End Function

Private Sub BLNew_Click()

'/// set the new file data
LblName.Caption = ""
LblID.Caption = "0"
Label9.Caption = "Archivo: // Sin nombre.blk \\"
Call RestoreData

End Sub

Private Sub BLOpen_Click()

Dim ConvertTx As String, DataFile As String

'/// display the dialog box
On Error Resume Next
TopMenu.BlockCmd.InitDir = App.path & AppBlockDir
TopMenu.BlockCmd.Filter = "Archivo de Bloque (*.blk)|*.blk|Archivo de Bloque"
TopMenu.BlockCmd.DialogTitle = "Bloques de publicidad - Abrir archivo de bloque."
TopMenu.BlockCmd.CancelError = True
TopMenu.BlockCmd.ShowOpen

If err.Number = 32755 Then Exit Sub

ConvertTx = TopMenu.BlockCmd.filename

'load the file data in the list
DataFile = ViewBlock.SelectedItem.SubItems(1)

'/// lets open the file for read the data and check it for OK
BlockData = OpenBlockFile(ConvertTx, DataFile, 0)

If BlockData.id <= 0 Then
    LblName.Caption = ConvertTx
    Label9.Caption = "Archivo: // " & StripFileFromDir(ConvertTx) & " \\"
    Call RestoreData
Else
    LblName.Caption = ConvertTx
    Label9.Caption = "Archivo: // " & StripFileFromDir(ConvertTx) & " \\"
    Call PutData(BlockData)
End If

End Sub

Private Sub BLSave_Click()

Dim ConvertTx As String, DataFile As String
Dim Day(0 To 2) As Integer
Dim Hor(0 To 2) As Integer
Dim Cant(0 To 2) As Integer
Dim i As Integer, RgID As Integer
Dim Result As Boolean

'/// check the days validity--------------------
If PFD(0).text = "" Or PFD(0).text = " " Then
    MsgBox "Los días seleccionados no son correctos.", vbCritical, Me.Caption
    Exit Sub
End If
'/// check the hours validity-------------------
If PFH(0).text = "" Or PFH(0).text = " " Then
    MsgBox "La hora seleccionada no es correcta.", vbCritical, Me.Caption
    Exit Sub
End If
'/// check the date validity--------------------
Result = CheckDateVal(FInit.text)
If IsDate(Trim(FInit.text)) = False Or Result = False Then
    MsgBox "La fecha de inicio no es correcta.", vbCritical, Me.Caption
    Exit Sub
End If
Result = CheckDateVal(FEnd.text)
If IsDate(Trim(FEnd.text)) = False Or Result = False Then
    MsgBox "La fecha de finalización no es correcta.", vbCritical, Me.Caption
    Exit Sub
End If

'/// lets start the save of file
If Trim(LblName.Caption) = "" Then
    '/// display the save dialog box
    On Error Resume Next
    TopMenu.BlockCmd.InitDir = App.path & AppBlockDir
    TopMenu.BlockCmd.Filter = "Archivo de Bloque (*.blk)|*.blk|Archivo de Bloque"
    TopMenu.BlockCmd.DialogTitle = "Bloques de publicidad - Guardar archivo de bloque."
    TopMenu.BlockCmd.CancelError = True
    TopMenu.BlockCmd.ShowSave
    If err.Number = 32755 Then Exit Sub
    ConvertTx = TopMenu.BlockCmd.filename
Else
    ConvertTx = Trim(LblName.Caption)
End If

'/// get the correct data before save
For i = 0 To 2      '///// extract the user selected data (DAY)
    Select Case Trim(PFD(i).text)
        Case "Domingo"
            Day(i) = BlockPrefD.Dom
        Case "Lunes"
            Day(i) = BlockPrefD.Lun
        Case "Martes"
            Day(i) = BlockPrefD.Mar
        Case "Miercoles"
            Day(i) = BlockPrefD.Mie
        Case "Jueves"
            Day(i) = BlockPrefD.Jue
        Case "Viernes"
            Day(i) = BlockPrefD.Vie
        Case "Sábado"
            Day(i) = BlockPrefD.Sab
        Case "Todos los días"
            Day(i) = BlockPrefD.All
        Case Else
            Day(i) = BlockPrefD.Vacio
    End Select
Next i

For i = 0 To 2      '///// extract the user selected data (HOUR)
    Select Case Trim(PFH(i).text)
        Case "1 a 2 hs"
            Hor(i) = BlockPrefH.d1a2
        Case "2 a 3 hs"
            Hor(i) = BlockPrefH.d2a3
        Case "3 a 4 hs"
            Hor(i) = BlockPrefH.d3a4
        Case "4 a 5 hs"
            Hor(i) = BlockPrefH.d4a5
        Case "5 a 6 hs"
            Hor(i) = BlockPrefH.d5a6
        Case "6 a 7 hs"
            Hor(i) = BlockPrefH.d6a7
        Case "7 a 8 hs"
            Hor(i) = BlockPrefH.d7a8
        Case "8 a 9 hs"
            Hor(i) = BlockPrefH.d8a9
        Case "9 a 10 hs"
            Hor(i) = BlockPrefH.d9a10
        Case "10 a 11 hs"
            Hor(i) = BlockPrefH.d10a11
        Case "11 a 12 hs"
            Hor(i) = BlockPrefH.d11a12
        Case "12 a 13 hs"
            Hor(i) = BlockPrefH.d12a13
        Case "13 a 14 hs"
            Hor(i) = BlockPrefH.d13a14
        Case "14 a 15 hs"
            Hor(i) = BlockPrefH.d14a15
        Case "15 a 16 hs"
            Hor(i) = BlockPrefH.d15a16
        Case "16 a 17 hs"
            Hor(i) = BlockPrefH.d16a17
        Case "17 a 18 hs"
            Hor(i) = BlockPrefH.d17a18
        Case "18 a 19 hs"
            Hor(i) = BlockPrefH.d18a19
        Case "19 a 20 hs"
            Hor(i) = BlockPrefH.d19a20
        Case "20 a 21 hs"
            Hor(i) = BlockPrefH.d20a21
        Case "21 a 22 hs"
            Hor(i) = BlockPrefH.d21a22
        Case "22 a 23 hs"
            Hor(i) = BlockPrefH.d22a23
        Case "23 a 00 hs"
            Hor(i) = BlockPrefH.d23a0
        Case "00 a 1 hs"
            Hor(i) = BlockPrefH.d0a1
        Case "Cualquier horario"
            Hor(i) = BlockPrefH.All
        Case Else
            Hor(i) = BlockPrefH.Vacio
    End Select
Next i

For i = 0 To 2      '///// extract the reproducer cant
    Cant(i) = CInt(TxtCant(i).text)
Next i

'/// continue...
'/// set the data to be saved
BlockData.FFileName = ViewBlock.SelectedItem.SubItems(1)
BlockData.FFilePath = ViewBlock.SelectedItem.text
BlockData.FFileDur = LblDur.Caption
BlockData.FPrefD(0) = Day(0)
BlockData.FPrefD(1) = Day(1)
BlockData.FPrefD(2) = Day(2)
BlockData.FPrefH(0) = Hor(0)
BlockData.FPrefH(1) = Hor(1)
BlockData.FPrefH(2) = Hor(2)
BlockData.FCantV(0) = Cant(0)
BlockData.FCantV(1) = Cant(1)
BlockData.FCantV(2) = Cant(2)
BlockData.FPubInit = FInit.text
BlockData.FPubFin = FEnd.text

'/// SAVE THE DATA INTO THE FILE
RgID = CInt(LblID.Caption)

If SaveBlockFile(ConvertTx, BlockData, RgID) = True Then
    LblName.Caption = ConvertTx
    Label9.Caption = "Archivo: // " & StripFileFromDir(ConvertTx) & " \\"
Else
    LblName.Caption = ""
    Label9.Caption = "Archivo: // Sin nombre.blk \\"
End If

End Sub

Private Sub CmdAccept_Click()

Unload Me

End Sub

Private Sub cmdCancel_Click()

Unload Me

End Sub

Private Sub CmdD_Click(index As Integer)

If Combo2.text = "Todos los días" Then
    PFD(index).text = Combo2.text
    PFD(index).BackColor = &HC0FFFF
    Exit Sub
End If

PFD(index).text = Combo2.text
PFD(index).BackColor = &HC0FFFF

'/// enable the next control
On Error Resume Next
CmdD(index + 1).Enabled = True

End Sub

Private Sub CmdH_Click(index As Integer)

If Combo1.text = "Cualquier horario" Then
    PFH(index).text = Combo1.text
    PFH(index).BackColor = &HC0FFFF
    Exit Sub
End If

PFH(index).text = Combo1.text
PFH(index).BackColor = &HC0FFFF

'/// enable the next control
On Error Resume Next
CmdH(index + 1).Enabled = True
VScroll1(index + 1).Enabled = True

End Sub

Private Sub Dir1_Change()

'Dir1.Path = Drive1.drive

Call ReloadDir(Dir1.path, "*.mp3")

End Sub

Private Sub Drive1_Change()

On Error Resume Next
Dir1.path = Drive1.Drive

End Sub

Private Sub FEnd_GotFocus()

FEnd.SelStart = 0
FEnd.SelLength = Len(FEnd.text)

End Sub

Private Sub FInit_GotFocus()

FInit.SelStart = 0
FInit.SelLength = Len(FInit.text)

End Sub

Private Sub Form_Load()

LbLProgress.Caption = ""

'/// load some resources bitmaps
    BLNew.Picture = LoadResPicture("ICO_NEW", 0)
    BLOpen.Picture = LoadResPicture("ICO_OPEN", 0)
    BLSave.Picture = LoadResPicture("ICO_SAVE", 0)
    CmdAccept.Caption = LoadResString(2000)
    CmdCancel.Caption = LoadResString(2001)

Call RestoreData
Call ADDdata
FInit.text = Date
FEnd.text = Date

'cargamos los datos guardados en el archivo de configuracion
ConfigData = OpenConfigFile

On Error GoTo er
Dir1.path = GetCipherConfigData(Trim(ConfigData.Dir_Com))
Call ReloadDir(Dir1.path, "*.mp3")
Exit Sub

er:
Dir1.path = App.path
Call ReloadDir(Dir1.path, "*.mp3")

End Sub

Private Sub ViewBlock_Click()

Dim DataFile As String, ConvertTx As String
Dim fileLen As String, TimeNcv As String, Result As String
Dim FName As String

'load the block data in FrmBlock
DataFile = ViewBlock.SelectedItem.text
ConvertTx = Trim(LblName.Caption)

'set the file path & name
FName = DataFile & ViewBlock.SelectedItem.SubItems(1)

'load the file info (duracion
fileLen = FileLoadLen(FName, "Stream")
TimeNcv = FormatSegs(fileLen)
Result = ConvSecToMin(CInt(TimeNcv))
'put the time display of file
LblDur.Caption = Result

If Trim(DataFile) = "" Or Trim(DataFile) = " " Then
    Exit Sub
Else
    '// lets open the block file info (prefs)...
    If ConvertTx = "" Or ConvertTx = " " Then
        Exit Sub
    Else
        'load the file data in the list
        DataFile = ViewBlock.SelectedItem.SubItems(1)
        '/// lets open the file for read the data and check it for OK
        BlockData = OpenBlockFile(ConvertTx, DataFile, 0)
        If BlockData.id <= 0 Then
            LblName.Caption = ConvertTx
            Label9.Caption = "Archivo: // " & StripFileFromDir(ConvertTx) & " \\"
            Call RestoreData
            LblID.Caption = "0"
        Else
            LblName.Caption = ConvertTx
            Label9.Caption = "Archivo: // " & StripFileFromDir(ConvertTx) & " \\"
            Call PutData(BlockData)
            LblID.Caption = BlockData.id
        End If
    End If
End If

End Sub

Private Sub VScroll1_Change(index As Integer)

TxtCant(index).text = VScroll1(index).Value

End Sub
