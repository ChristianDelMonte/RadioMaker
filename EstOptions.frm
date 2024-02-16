VERSION 5.00
Begin VB.Form EstOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ESTACIONES 01 y 02 - Opciones de Reproducción"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Opciones Generales"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Autoimportación de archivos NetShow PlayList"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Autoreproducción al hacer click"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "EstOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
