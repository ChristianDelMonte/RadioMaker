VERSION 5.00
Begin VB.Form FrmRef 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ESTACION 01 - Formulario de Control"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      Caption         =   "Longitud Total (Seg/Mseg):"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      Caption         =   "Longitud Total (bytes):"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Caption         =   "Posicion Actual (Seg/Mseg):"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "Posicion Actual  (bytes):"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "Posicion Restante (Seg):"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "Posicion Restante (bytes):"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label20 
      Height          =   255
      Left            =   2280
      TabIndex        =   20
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label19 
      Height          =   255
      Left            =   2280
      TabIndex        =   19
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label18 
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label17 
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label16 
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label15 
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label14 
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label13 
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   3480
      X2              =   120
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label12 
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Posicion Restante (bytes):"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Posicion Restante (Seg):"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Posicion Actual  (bytes):"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Posicion Actual (Seg/Mseg):"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Longitud Total (bytes):"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Longitud Total (Seg/Mseg):"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "MUSIC 01"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "STREAM 01"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "FrmRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub
