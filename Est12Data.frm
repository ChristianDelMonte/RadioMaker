VERSION 5.00
Begin VB.Form Est12Data 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Estacion 01 y 02 Data Control"
   ClientHeight    =   10110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   FillColor       =   &H00000080&
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10110
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Label PD 
      Height          =   285
      Index           =   23
      Left            =   8565
      TabIndex        =   248
      Top             =   9690
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   22
      Left            =   8565
      TabIndex        =   247
      Top             =   9300
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   21
      Left            =   8550
      TabIndex        =   246
      Top             =   8925
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   20
      Left            =   8550
      TabIndex        =   245
      Top             =   8550
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   19
      Left            =   8535
      TabIndex        =   244
      Top             =   8145
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   18
      Left            =   8535
      TabIndex        =   243
      Top             =   7785
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   17
      Left            =   8550
      TabIndex        =   242
      Top             =   7410
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   16
      Left            =   8550
      TabIndex        =   241
      Top             =   7050
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   15
      Left            =   8550
      TabIndex        =   240
      Top             =   6675
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   14
      Left            =   8535
      TabIndex        =   239
      Top             =   6270
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   13
      Left            =   8535
      TabIndex        =   238
      Top             =   5895
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   12
      Left            =   8535
      TabIndex        =   237
      Top             =   5520
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   11
      Left            =   8535
      TabIndex        =   236
      Top             =   5115
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   10
      Left            =   8535
      TabIndex        =   235
      Top             =   4755
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   9
      Left            =   8535
      TabIndex        =   234
      Top             =   4380
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   8
      Left            =   8520
      TabIndex        =   233
      Top             =   4005
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   7
      Left            =   8535
      TabIndex        =   232
      Top             =   3585
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   6
      Left            =   8520
      TabIndex        =   231
      Top             =   3195
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   5
      Left            =   8520
      TabIndex        =   230
      Top             =   2820
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   4
      Left            =   8520
      TabIndex        =   229
      Top             =   2445
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   3
      Left            =   8520
      TabIndex        =   228
      Top             =   2070
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   2
      Left            =   8505
      TabIndex        =   227
      Top             =   1695
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   1
      Left            =   8505
      TabIndex        =   226
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label PD 
      Height          =   285
      Index           =   0
      Left            =   8505
      TabIndex        =   225
      Top             =   945
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   23
      Left            =   7650
      TabIndex        =   224
      Top             =   9690
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   22
      Left            =   7650
      TabIndex        =   223
      Top             =   9300
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   21
      Left            =   7650
      TabIndex        =   222
      Top             =   8925
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   20
      Left            =   7650
      TabIndex        =   221
      Top             =   8565
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   19
      Left            =   7650
      TabIndex        =   220
      Top             =   8175
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   18
      Left            =   7650
      TabIndex        =   219
      Top             =   7800
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   17
      Left            =   7635
      TabIndex        =   218
      Top             =   7425
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   16
      Left            =   7635
      TabIndex        =   217
      Top             =   7035
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   15
      Left            =   7635
      TabIndex        =   216
      Top             =   6660
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   14
      Left            =   7635
      TabIndex        =   215
      Top             =   6270
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   13
      Left            =   7635
      TabIndex        =   214
      Top             =   5880
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   12
      Left            =   7635
      TabIndex        =   213
      Top             =   5505
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   11
      Left            =   7635
      TabIndex        =   212
      Top             =   5115
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   10
      Left            =   7635
      TabIndex        =   211
      Top             =   4725
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   9
      Left            =   7635
      TabIndex        =   210
      Top             =   4350
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   8
      Left            =   7635
      TabIndex        =   209
      Top             =   3975
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   7
      Left            =   7635
      TabIndex        =   208
      Top             =   3585
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   6
      Left            =   7635
      TabIndex        =   207
      Top             =   3210
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   5
      Left            =   7620
      TabIndex        =   206
      Top             =   2850
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   4
      Left            =   7620
      TabIndex        =   205
      Top             =   2460
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   3
      Left            =   7620
      TabIndex        =   204
      Top             =   2085
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   2
      Left            =   7620
      TabIndex        =   203
      Top             =   1710
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   1
      Left            =   7620
      TabIndex        =   202
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label PC 
      Height          =   285
      Index           =   0
      Left            =   7620
      TabIndex        =   201
      Top             =   945
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   23
      Left            =   6750
      TabIndex        =   200
      Top             =   9660
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   22
      Left            =   6750
      TabIndex        =   199
      Top             =   9270
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   21
      Left            =   6750
      TabIndex        =   198
      Top             =   8895
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   20
      Left            =   6750
      TabIndex        =   197
      Top             =   8535
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   19
      Left            =   6750
      TabIndex        =   196
      Top             =   8175
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   18
      Left            =   6750
      TabIndex        =   195
      Top             =   7785
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   17
      Left            =   6750
      TabIndex        =   194
      Top             =   7425
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   16
      Left            =   6750
      TabIndex        =   193
      Top             =   7020
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   15
      Left            =   6750
      TabIndex        =   192
      Top             =   6645
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   14
      Left            =   6750
      TabIndex        =   191
      Top             =   6240
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   13
      Left            =   6750
      TabIndex        =   190
      Top             =   5880
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   12
      Left            =   6750
      TabIndex        =   189
      Top             =   5505
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   11
      Left            =   6765
      TabIndex        =   188
      Top             =   5115
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   10
      Left            =   6750
      TabIndex        =   187
      Top             =   4725
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   9
      Left            =   6750
      TabIndex        =   186
      Top             =   4335
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   8
      Left            =   6750
      TabIndex        =   185
      Top             =   3945
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   7
      Left            =   6750
      TabIndex        =   184
      Top             =   3570
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   6
      Left            =   6750
      TabIndex        =   183
      Top             =   3180
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   5
      Left            =   6750
      TabIndex        =   182
      Top             =   2820
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   4
      Left            =   6750
      TabIndex        =   181
      Top             =   2460
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   3
      Left            =   6750
      TabIndex        =   180
      Top             =   2055
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   2
      Left            =   6750
      TabIndex        =   179
      Top             =   1680
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   1
      Left            =   6750
      TabIndex        =   178
      Top             =   1305
      Width           =   690
   End
   Begin VB.Label PF 
      Height          =   285
      Index           =   0
      Left            =   6750
      TabIndex        =   177
      Top             =   930
      Width           =   690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estacion 01 y 02 Data Control"
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
      Left            =   5610
      TabIndex        =   175
      Top             =   8520
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   20
      Left            =   5610
      TabIndex        =   174
      Top             =   8160
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   19
      Left            =   5610
      TabIndex        =   173
      Top             =   7800
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   18
      Left            =   5610
      TabIndex        =   172
      Top             =   7440
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   17
      Left            =   5610
      TabIndex        =   171
      Top             =   7080
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   16
      Left            =   5610
      TabIndex        =   170
      Top             =   6720
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   15
      Left            =   5610
      TabIndex        =   169
      Top             =   6360
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   14
      Left            =   5610
      TabIndex        =   168
      Top             =   6000
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   13
      Left            =   5610
      TabIndex        =   167
      Top             =   5640
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   12
      Left            =   5610
      TabIndex        =   166
      Top             =   5280
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   11
      Left            =   5610
      TabIndex        =   165
      Top             =   4920
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   10
      Left            =   5610
      TabIndex        =   164
      Top             =   4560
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   9
      Left            =   5610
      TabIndex        =   163
      Top             =   4200
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   8
      Left            =   5610
      TabIndex        =   162
      Top             =   3840
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   7
      Left            =   5610
      TabIndex        =   161
      Top             =   3480
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   6
      Left            =   5610
      TabIndex        =   160
      Top             =   3120
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   5
      Left            =   5610
      TabIndex        =   159
      Top             =   2760
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   4
      Left            =   5610
      TabIndex        =   158
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   3
      Left            =   5610
      TabIndex        =   157
      Top             =   2040
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   2
      Left            =   5610
      TabIndex        =   156
      Top             =   1680
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   1
      Left            =   5610
      TabIndex        =   155
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label V2 
      Height          =   255
      Index           =   0
      Left            =   5610
      TabIndex        =   154
      Top             =   960
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   21
      Left            =   4860
      TabIndex        =   153
      Top             =   8520
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   20
      Left            =   4860
      TabIndex        =   152
      Top             =   8160
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   19
      Left            =   4860
      TabIndex        =   151
      Top             =   7800
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   18
      Left            =   4860
      TabIndex        =   150
      Top             =   7440
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   17
      Left            =   4860
      TabIndex        =   149
      Top             =   7080
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   16
      Left            =   4860
      TabIndex        =   148
      Top             =   6720
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   15
      Left            =   4860
      TabIndex        =   147
      Top             =   6360
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   14
      Left            =   4860
      TabIndex        =   146
      Top             =   6000
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   13
      Left            =   4860
      TabIndex        =   145
      Top             =   5640
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   12
      Left            =   4860
      TabIndex        =   144
      Top             =   5280
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   11
      Left            =   4860
      TabIndex        =   143
      Top             =   4920
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   10
      Left            =   4860
      TabIndex        =   142
      Top             =   4560
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   9
      Left            =   4860
      TabIndex        =   141
      Top             =   4200
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   8
      Left            =   4860
      TabIndex        =   140
      Top             =   3840
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   7
      Left            =   4860
      TabIndex        =   139
      Top             =   3480
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   6
      Left            =   4860
      TabIndex        =   138
      Top             =   3120
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   5
      Left            =   4860
      TabIndex        =   137
      Top             =   2760
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   4
      Left            =   4860
      TabIndex        =   136
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   3
      Left            =   4860
      TabIndex        =   135
      Top             =   2040
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   2
      Left            =   4860
      TabIndex        =   134
      Top             =   1680
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   1
      Left            =   4860
      TabIndex        =   133
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label D2 
      Height          =   255
      Index           =   0
      Left            =   4860
      TabIndex        =   132
      Top             =   960
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   21
      Left            =   4110
      TabIndex        =   131
      Top             =   8520
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   20
      Left            =   4110
      TabIndex        =   130
      Top             =   8160
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   19
      Left            =   4110
      TabIndex        =   129
      Top             =   7800
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   18
      Left            =   4110
      TabIndex        =   128
      Top             =   7440
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   17
      Left            =   4110
      TabIndex        =   127
      Top             =   7080
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   16
      Left            =   4110
      TabIndex        =   126
      Top             =   6720
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   15
      Left            =   4110
      TabIndex        =   125
      Top             =   6360
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   14
      Left            =   4110
      TabIndex        =   124
      Top             =   6000
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   13
      Left            =   4110
      TabIndex        =   123
      Top             =   5640
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   12
      Left            =   4110
      TabIndex        =   122
      Top             =   5280
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   11
      Left            =   4110
      TabIndex        =   121
      Top             =   4920
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   10
      Left            =   4110
      TabIndex        =   120
      Top             =   4560
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   9
      Left            =   4110
      TabIndex        =   119
      Top             =   4200
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   8
      Left            =   4110
      TabIndex        =   118
      Top             =   3840
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   7
      Left            =   4110
      TabIndex        =   117
      Top             =   3480
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   6
      Left            =   4110
      TabIndex        =   116
      Top             =   3120
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   5
      Left            =   4110
      TabIndex        =   115
      Top             =   2760
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   4
      Left            =   4110
      TabIndex        =   114
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   3
      Left            =   4110
      TabIndex        =   113
      Top             =   2040
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   2
      Left            =   4110
      TabIndex        =   112
      Top             =   1680
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   1
      Left            =   4110
      TabIndex        =   111
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label N2 
      Height          =   255
      Index           =   0
      Left            =   4110
      TabIndex        =   110
      Top             =   960
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   21
      Left            =   3390
      TabIndex        =   109
      Top             =   8520
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   20
      Left            =   3390
      TabIndex        =   108
      Top             =   8160
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   19
      Left            =   3390
      TabIndex        =   107
      Top             =   7800
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   18
      Left            =   3390
      TabIndex        =   106
      Top             =   7440
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   17
      Left            =   3390
      TabIndex        =   105
      Top             =   7080
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   16
      Left            =   3390
      TabIndex        =   104
      Top             =   6720
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   15
      Left            =   3390
      TabIndex        =   103
      Top             =   6360
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   14
      Left            =   3390
      TabIndex        =   102
      Top             =   6000
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   13
      Left            =   3390
      TabIndex        =   101
      Top             =   5640
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   12
      Left            =   3390
      TabIndex        =   100
      Top             =   5280
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   11
      Left            =   3390
      TabIndex        =   99
      Top             =   4920
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   10
      Left            =   3390
      TabIndex        =   98
      Top             =   4560
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   9
      Left            =   3390
      TabIndex        =   97
      Top             =   4200
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   8
      Left            =   3390
      TabIndex        =   96
      Top             =   3840
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   7
      Left            =   3390
      TabIndex        =   95
      Top             =   3480
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   6
      Left            =   3390
      TabIndex        =   94
      Top             =   3120
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   5
      Left            =   3390
      TabIndex        =   93
      Top             =   2760
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   4
      Left            =   3390
      TabIndex        =   92
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   3
      Left            =   3390
      TabIndex        =   91
      Top             =   2040
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   2
      Left            =   3390
      TabIndex        =   90
      Top             =   1680
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   1
      Left            =   3390
      TabIndex        =   89
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label C2 
      Height          =   255
      Index           =   0
      Left            =   3390
      TabIndex        =   88
      Top             =   960
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   21
      Left            =   2220
      TabIndex        =   87
      Top             =   8535
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   20
      Left            =   2220
      TabIndex        =   86
      Top             =   8175
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   19
      Left            =   2220
      TabIndex        =   85
      Top             =   7815
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   18
      Left            =   2220
      TabIndex        =   84
      Top             =   7455
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   17
      Left            =   2220
      TabIndex        =   83
      Top             =   7095
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   16
      Left            =   2220
      TabIndex        =   82
      Top             =   6735
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   15
      Left            =   2220
      TabIndex        =   81
      Top             =   6375
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   14
      Left            =   2220
      TabIndex        =   80
      Top             =   6015
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   13
      Left            =   2220
      TabIndex        =   79
      Top             =   5655
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   12
      Left            =   2220
      TabIndex        =   78
      Top             =   5295
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   11
      Left            =   2220
      TabIndex        =   77
      Top             =   4935
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   10
      Left            =   2220
      TabIndex        =   76
      Top             =   4575
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   9
      Left            =   2220
      TabIndex        =   75
      Top             =   4215
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   8
      Left            =   2220
      TabIndex        =   74
      Top             =   3855
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   7
      Left            =   2220
      TabIndex        =   73
      Top             =   3495
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   6
      Left            =   2220
      TabIndex        =   72
      Top             =   3135
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   5
      Left            =   2220
      TabIndex        =   71
      Top             =   2775
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   4
      Left            =   2220
      TabIndex        =   70
      Top             =   2415
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   3
      Left            =   2220
      TabIndex        =   69
      Top             =   2055
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   2
      Left            =   2220
      TabIndex        =   68
      Top             =   1695
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   1
      Left            =   2220
      TabIndex        =   67
      Top             =   1335
      Width           =   675
   End
   Begin VB.Label V1 
      Height          =   255
      Index           =   0
      Left            =   2220
      TabIndex        =   66
      Top             =   975
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   21
      Left            =   1455
      TabIndex        =   65
      Top             =   8535
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   20
      Left            =   1455
      TabIndex        =   64
      Top             =   8175
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   19
      Left            =   1455
      TabIndex        =   63
      Top             =   7815
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   18
      Left            =   1455
      TabIndex        =   62
      Top             =   7455
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   17
      Left            =   1455
      TabIndex        =   61
      Top             =   7095
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   16
      Left            =   1455
      TabIndex        =   60
      Top             =   6735
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   15
      Left            =   1455
      TabIndex        =   59
      Top             =   6375
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   14
      Left            =   1455
      TabIndex        =   58
      Top             =   6015
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   13
      Left            =   1455
      TabIndex        =   57
      Top             =   5655
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   12
      Left            =   1455
      TabIndex        =   56
      Top             =   5295
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   11
      Left            =   1455
      TabIndex        =   55
      Top             =   4935
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   10
      Left            =   1455
      TabIndex        =   54
      Top             =   4575
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   9
      Left            =   1455
      TabIndex        =   53
      Top             =   4215
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   8
      Left            =   1455
      TabIndex        =   52
      Top             =   3855
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   7
      Left            =   1455
      TabIndex        =   51
      Top             =   3495
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   6
      Left            =   1455
      TabIndex        =   50
      Top             =   3135
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   5
      Left            =   1455
      TabIndex        =   49
      Top             =   2775
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   4
      Left            =   1455
      TabIndex        =   48
      Top             =   2415
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   3
      Left            =   1455
      TabIndex        =   47
      Top             =   2055
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   2
      Left            =   1455
      TabIndex        =   46
      Top             =   1695
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   1
      Left            =   1455
      TabIndex        =   45
      Top             =   1335
      Width           =   675
   End
   Begin VB.Label D1 
      Height          =   255
      Index           =   0
      Left            =   1455
      TabIndex        =   44
      Top             =   975
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   21
      Left            =   705
      TabIndex        =   43
      Top             =   8535
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   20
      Left            =   705
      TabIndex        =   42
      Top             =   8175
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   19
      Left            =   705
      TabIndex        =   41
      Top             =   7815
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   18
      Left            =   705
      TabIndex        =   40
      Top             =   7455
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   17
      Left            =   705
      TabIndex        =   39
      Top             =   7095
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   16
      Left            =   705
      TabIndex        =   38
      Top             =   6735
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   15
      Left            =   705
      TabIndex        =   37
      Top             =   6375
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   14
      Left            =   705
      TabIndex        =   36
      Top             =   6015
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   13
      Left            =   705
      TabIndex        =   35
      Top             =   5655
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   12
      Left            =   705
      TabIndex        =   34
      Top             =   5295
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   11
      Left            =   705
      TabIndex        =   33
      Top             =   4935
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   10
      Left            =   705
      TabIndex        =   32
      Top             =   4575
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   9
      Left            =   705
      TabIndex        =   31
      Top             =   4215
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   8
      Left            =   705
      TabIndex        =   30
      Top             =   3855
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   7
      Left            =   705
      TabIndex        =   29
      Top             =   3495
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   6
      Left            =   705
      TabIndex        =   28
      Top             =   3135
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   5
      Left            =   705
      TabIndex        =   27
      Top             =   2775
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   4
      Left            =   705
      TabIndex        =   26
      Top             =   2415
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   3
      Left            =   705
      TabIndex        =   25
      Top             =   2055
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   2
      Left            =   705
      TabIndex        =   24
      Top             =   1695
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   1
      Left            =   705
      TabIndex        =   23
      Top             =   1335
      Width           =   675
   End
   Begin VB.Label N1 
      Height          =   255
      Index           =   0
      Left            =   705
      TabIndex        =   22
      Top             =   975
      Width           =   675
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   21
      Left            =   180
      TabIndex        =   21
      Top             =   8535
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   20
      Left            =   180
      TabIndex        =   20
      Top             =   8175
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   19
      Left            =   180
      TabIndex        =   19
      Top             =   7815
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   18
      Left            =   180
      TabIndex        =   18
      Top             =   7455
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   17
      Left            =   180
      TabIndex        =   17
      Top             =   7095
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   16
      Left            =   180
      TabIndex        =   16
      Top             =   6735
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   15
      Left            =   180
      TabIndex        =   15
      Top             =   6375
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   14
      Left            =   180
      TabIndex        =   14
      Top             =   6015
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   13
      Left            =   180
      TabIndex        =   13
      Top             =   5655
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   12
      Left            =   180
      TabIndex        =   12
      Top             =   5295
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   11
      Left            =   180
      TabIndex        =   11
      Top             =   4935
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   10
      Left            =   180
      TabIndex        =   10
      Top             =   4575
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   9
      Left            =   180
      TabIndex        =   9
      Top             =   4215
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   8
      Left            =   180
      TabIndex        =   8
      Top             =   3855
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   7
      Left            =   180
      TabIndex        =   7
      Top             =   3495
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   6
      Top             =   3135
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   5
      Top             =   2775
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   4
      Top             =   2415
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   3
      Top             =   2055
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   1695
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   1335
      Width           =   495
   End
   Begin VB.Label C1 
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   975
      Width           =   495
   End
End
Attribute VB_Name = "Est12Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
