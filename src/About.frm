VERSION 5.00
Begin VB.Form About 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3630
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "About.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version 2.0.2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   3285
      TabIndex        =   14
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "for the project assignment problem"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   612
      Index           =   2
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   4932
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   0
      X1              =   240
      X2              =   255
      Y1              =   2400
      Y2              =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "UK.  SO17 1BJ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   252
      Index           =   3
      Left            =   4320
      TabIndex        =   11
      Top             =   3000
      Width           =   3132
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Highfield, Southampton"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   252
      Index           =   2
      Left            =   4320
      TabIndex        =   10
      Top             =   2760
      Width           =   3132
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "University of Southampton"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   252
      Index           =   1
      Left            =   4320
      TabIndex        =   9
      Top             =   2520
      Width           =   3132
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Faculty of Mathematical Studies"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   252
      Index           =   0
      Left            =   4320
      TabIndex        =   8
      Top             =   2280
      Width           =   3252
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   2
      X1              =   240
      X2              =   255
      Y1              =   3120
      Y2              =   3135
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   1
      X1              =   240
      X2              =   255
      Y1              =   2760
      Y2              =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Paul Harper"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   3000
      Width           =   1572
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Valter de Senna"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   2292
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Israel Vieira"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   1572
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed and Programmed by:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   4572
   End
   Begin VB.Image Image1 
      Height          =   1860
      Left            =   5400
      Picture         =   "About.frx":000C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2112
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "lgorithm"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   732
      Index           =   1
      Left            =   2810
      TabIndex        =   3
      Top             =   360
      Width           =   2292
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "enetic"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   612
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   1812
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   49.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   936
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   816
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   49.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1296
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   852
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   3372
      Left            =   120
      Top             =   120
      Width           =   7524
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

