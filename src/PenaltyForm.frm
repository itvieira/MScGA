VERSION 5.00
Begin VB.Form PenaltyForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penalty Function"
   ClientHeight    =   4530
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "PenaltyForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5895
      Begin VB.TextBox pen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   4800
         TabIndex        =   6
         Text            =   "360"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox pen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   4800
         TabIndex        =   7
         Text            =   "490"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox pen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   4800
         TabIndex        =   8
         Text            =   "640"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox pen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   4800
         TabIndex        =   9
         Text            =   "810"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox pen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   4800
         TabIndex        =   10
         Text            =   "1000"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox pen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   4800
         TabIndex        =   12
         Text            =   "3000"
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox pen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   1800
         TabIndex        =   11
         Text            =   "1000"
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox pen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1800
         TabIndex        =   5
         Text            =   "250"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox pen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1800
         TabIndex        =   4
         Text            =   "160"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox pen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1800
         TabIndex        =   3
         Text            =   "90"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox pen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1800
         TabIndex        =   2
         Text            =   "40"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox pen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1800
         TabIndex        =   1
         Text            =   "10"
         Top             =   240
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   1
         X1              =   120
         X2              =   5760
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   5760
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label1 
         Caption         =   "Ninth ---------------:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   3120
         TabIndex        =   25
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Sixth ---------------:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   3120
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Seventh ----------:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   3120
         TabIndex        =   23
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Eighth -------------:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   3120
         TabIndex        =   22
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tenth --------------:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   3120
         TabIndex        =   21
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Duplication-------:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   3120
         TabIndex        =   19
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Non Choice ------:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Fifth ----------------:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "Third ---------------:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "Second -----------:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "First ----------------:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Fourth -------------:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1572
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Please assign penalty weights for project allocations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "PenaltyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo Err_Form_Load
Dim i

For i = 0 To 11
   pen(i) = Penalty(i + 1)
Next i

Exit_Err_Form_Load:
Exit Sub

Err_Form_Load:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_Form_Load
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Form_Unload
Dim i, ChangedPenalties As Boolean
For i = 0 To 11
   If pen(i).DataChanged Then
   Penalty(i + 1) = pen(i)
   USavePenalty = False
   ChangedPenalties = True
   End If
Next i

If ChangedPenalties Then
    Screen.MousePointer = 11
    BCurrent = 999999999: BPos = 0: WCurrent = 0: WPos = 0
    For i = 1 To MainForm.FNChromo
        Fit = 0
        For j = 1 To MainForm.FNStudent
            'Fit calculation
            If MainForm.RGrid.TextMatrix(j, MyChromo(i, j) + 1) <> "" Then
                Fit = Fit + ((Penalty(CLng(MainForm.RGrid.TextMatrix(j, MyChromo(i, j) + 1)))) * (CSng(MainForm.RGrid.TextMatrix(j, 1))))
            Else
                Fit = Fit + (Penalty(11) * (CSng(MainForm.RGrid.TextMatrix(j, 1))))
            End If
            For k = 1 To j - 1
                If MyChromo(i, k) = MyChromo(i, j) Then
                    Fit = Fit + (Penalty(12) * (CSng(MainForm.RGrid.TextMatrix(j, 1))))
                    Exit For
                End If
            Next k
        Next j
    
        MyChromo(i, MainForm.FNStudent + 1) = Fit
    
        If Fit < BCurrent Then BCurrent = Fit: BPos = i
        If Fit > WCurrent Then WCurrent = Fit: WPos = i
    Next i
  
  MainForm.StartFit.Caption = BCurrent
  MainForm.WorstFit.Caption = WCurrent
  MainForm.CurrentFit.Caption = BCurrent
  MainForm.NRCount.Refresh
  MainForm.Results.Clear
  Call MainForm.WriteBestOnScreen(MainForm.FNStudent, BPos)
  Screen.MousePointer = 0
End If

Exit_Err_Form_Unload:
Exit Sub

Err_Form_Unload:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_Form_Unload
End Sub
