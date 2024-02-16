VERSION 5.00
Begin VB.Form Est34Data 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   FillColor       =   &H00000080&
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   645
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estacion 03 y 04 Data Control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   176
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   21
      Left            =   8640
      TabIndex        =   175
      Top             =   9000
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   20
      Left            =   8640
      TabIndex        =   174
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   19
      Left            =   8640
      TabIndex        =   173
      Top             =   8280
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   18
      Left            =   8640
      TabIndex        =   172
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   17
      Left            =   8640
      TabIndex        =   171
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   16
      Left            =   8640
      TabIndex        =   170
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   15
      Left            =   8640
      TabIndex        =   169
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   14
      Left            =   8640
      TabIndex        =   168
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   13
      Left            =   8640
      TabIndex        =   167
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   12
      Left            =   8640
      TabIndex        =   166
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   11
      Left            =   8640
      TabIndex        =   165
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   10
      Left            =   8640
      TabIndex        =   164
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   9
      Left            =   8640
      TabIndex        =   163
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   8
      Left            =   8640
      TabIndex        =   162
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   7
      Left            =   8640
      TabIndex        =   161
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   6
      Left            =   8640
      TabIndex        =   160
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   5
      Left            =   8640
      TabIndex        =   159
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   4
      Left            =   8640
      TabIndex        =   158
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   3
      Left            =   8640
      TabIndex        =   157
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   2
      Left            =   8640
      TabIndex        =   156
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   1
      Left            =   8640
      TabIndex        =   155
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   0
      Left            =   8640
      TabIndex        =   154
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   21
      Left            =   7800
      TabIndex        =   153
      Top             =   9000
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   20
      Left            =   7800
      TabIndex        =   152
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   19
      Left            =   7800
      TabIndex        =   151
      Top             =   8280
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   18
      Left            =   7800
      TabIndex        =   150
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   17
      Left            =   7800
      TabIndex        =   149
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   16
      Left            =   7800
      TabIndex        =   148
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   15
      Left            =   7800
      TabIndex        =   147
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   14
      Left            =   7800
      TabIndex        =   146
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   13
      Left            =   7800
      TabIndex        =   145
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   12
      Left            =   7800
      TabIndex        =   144
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   11
      Left            =   7800
      TabIndex        =   143
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   10
      Left            =   7800
      TabIndex        =   142
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   9
      Left            =   7800
      TabIndex        =   141
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   8
      Left            =   7800
      TabIndex        =   140
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   7
      Left            =   7800
      TabIndex        =   139
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   6
      Left            =   7800
      TabIndex        =   138
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   5
      Left            =   7800
      TabIndex        =   137
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   4
      Left            =   7800
      TabIndex        =   136
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   3
      Left            =   7800
      TabIndex        =   135
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   2
      Left            =   7800
      TabIndex        =   134
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   1
      Left            =   7800
      TabIndex        =   133
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   0
      Left            =   7800
      TabIndex        =   132
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   21
      Left            =   6600
      TabIndex        =   131
      Top             =   9000
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   20
      Left            =   6600
      TabIndex        =   130
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   19
      Left            =   6600
      TabIndex        =   129
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   18
      Left            =   6600
      TabIndex        =   128
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   17
      Left            =   6600
      TabIndex        =   127
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   16
      Left            =   6600
      TabIndex        =   126
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   15
      Left            =   6600
      TabIndex        =   125
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   14
      Left            =   6600
      TabIndex        =   124
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   13
      Left            =   6600
      TabIndex        =   123
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   12
      Left            =   6600
      TabIndex        =   122
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   11
      Left            =   6600
      TabIndex        =   121
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   10
      Left            =   6600
      TabIndex        =   120
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   9
      Left            =   6600
      TabIndex        =   119
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   8
      Left            =   6600
      TabIndex        =   118
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   7
      Left            =   6600
      TabIndex        =   117
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   116
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   5
      Left            =   6600
      TabIndex        =   115
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   4
      Left            =   6600
      TabIndex        =   114
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   3
      Left            =   6600
      TabIndex        =   113
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   112
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   111
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   110
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   21
      Left            =   5400
      TabIndex        =   109
      Top             =   9000
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   20
      Left            =   5400
      TabIndex        =   108
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   19
      Left            =   5400
      TabIndex        =   107
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   18
      Left            =   5400
      TabIndex        =   106
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   17
      Left            =   5400
      TabIndex        =   105
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   16
      Left            =   5400
      TabIndex        =   104
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   15
      Left            =   5400
      TabIndex        =   103
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   14
      Left            =   5400
      TabIndex        =   102
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   13
      Left            =   5400
      TabIndex        =   101
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   12
      Left            =   5400
      TabIndex        =   100
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   11
      Left            =   5400
      TabIndex        =   99
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   10
      Left            =   5400
      TabIndex        =   98
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   9
      Left            =   5400
      TabIndex        =   97
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   8
      Left            =   5400
      TabIndex        =   96
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   7
      Left            =   5400
      TabIndex        =   95
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   6
      Left            =   5400
      TabIndex        =   94
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   93
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   4
      Left            =   5400
      TabIndex        =   92
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   91
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   90
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   89
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   88
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   21
      Left            =   3600
      TabIndex        =   87
      Top             =   9000
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   20
      Left            =   3600
      TabIndex        =   86
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   19
      Left            =   3600
      TabIndex        =   85
      Top             =   8280
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   18
      Left            =   3600
      TabIndex        =   84
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   17
      Left            =   3600
      TabIndex        =   83
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   16
      Left            =   3600
      TabIndex        =   82
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   15
      Left            =   3600
      TabIndex        =   81
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   14
      Left            =   3600
      TabIndex        =   80
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   13
      Left            =   3600
      TabIndex        =   79
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   12
      Left            =   3600
      TabIndex        =   78
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   11
      Left            =   3600
      TabIndex        =   77
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   10
      Left            =   3600
      TabIndex        =   76
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   75
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   74
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   7
      Left            =   3600
      TabIndex        =   73
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   72
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   71
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   70
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   69
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   68
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   67
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   66
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   21
      Left            =   2760
      TabIndex        =   65
      Top             =   9000
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   20
      Left            =   2760
      TabIndex        =   64
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   19
      Left            =   2760
      TabIndex        =   63
      Top             =   8280
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   18
      Left            =   2760
      TabIndex        =   62
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   17
      Left            =   2760
      TabIndex        =   61
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   16
      Left            =   2760
      TabIndex        =   60
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   15
      Left            =   2760
      TabIndex        =   59
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   14
      Left            =   2760
      TabIndex        =   58
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   13
      Left            =   2760
      TabIndex        =   57
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   12
      Left            =   2760
      TabIndex        =   56
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   11
      Left            =   2760
      TabIndex        =   55
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   54
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   9
      Left            =   2760
      TabIndex        =   53
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   52
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   51
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   50
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   49
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   48
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   47
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   46
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   45
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   44
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   21
      Left            =   1560
      TabIndex        =   43
      Top             =   9000
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   20
      Left            =   1560
      TabIndex        =   42
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   19
      Left            =   1560
      TabIndex        =   41
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   18
      Left            =   1560
      TabIndex        =   40
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   17
      Left            =   1560
      TabIndex        =   39
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   16
      Left            =   1560
      TabIndex        =   38
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   15
      Left            =   1560
      TabIndex        =   37
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   14
      Left            =   1560
      TabIndex        =   36
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   13
      Left            =   1560
      TabIndex        =   35
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   12
      Left            =   1560
      TabIndex        =   34
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   11
      Left            =   1560
      TabIndex        =   33
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   10
      Left            =   1560
      TabIndex        =   32
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   9
      Left            =   1560
      TabIndex        =   31
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   8
      Left            =   1560
      TabIndex        =   30
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   7
      Left            =   1560
      TabIndex        =   29
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   28
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   27
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   26
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   25
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   24
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   23
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   22
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   21
      Left            =   360
      TabIndex        =   21
      Top             =   9000
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   20
      Left            =   360
      TabIndex        =   20
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   19
      Left            =   360
      TabIndex        =   19
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   18
      Left            =   360
      TabIndex        =   18
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   17
      Left            =   360
      TabIndex        =   17
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   16
      Left            =   360
      TabIndex        =   16
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   15
      Left            =   360
      TabIndex        =   15
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   14
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   13
      Left            =   360
      TabIndex        =   13
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   12
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   11
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   10
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   9
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "Est34Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
