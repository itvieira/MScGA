VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   Caption         =   "Genetic Algorithm"
   ClientHeight    =   6180
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9630
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton StopRun 
      Height          =   375
      Left            =   3075
      Picture         =   "MainForm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Stop running."
      Top             =   1035
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   2895
      Begin VB.TextBox FMRate 
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
         Left            =   1800
         TabIndex        =   3
         Text            =   "0.02"
         Top             =   1530
         Width           =   975
      End
      Begin VB.TextBox FNStudent 
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
         Left            =   1800
         TabIndex        =   0
         Text            =   "2"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox FNProject 
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
         Left            =   1800
         TabIndex        =   1
         Text            =   "2"
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox FNChromo 
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
         Left            =   1800
         TabIndex        =   2
         Text            =   "50"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox FNRun 
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
         Left            =   1800
         TabIndex        =   4
         Text            =   "1000"
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Mutation rate --:"
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
         TabIndex        =   25
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Students ---------:"
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
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Projects ----------:"
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
         TabIndex        =   23
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Chromos ---------:"
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
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Run size ---------:"
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
         TabIndex        =   21
         Top             =   1980
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   3480
      TabIndex        =   9
      Top             =   0
      Width           =   8295
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Worst Fitness"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label WorstFit 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   255
         Left            =   5400
         TabIndex        =   26
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Start Fitness"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   4110
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label StartFit 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Best Fitness"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6720
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label CurrentFit 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   255
         Left            =   6720
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Current run size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2820
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label CurrentRun 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Running size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Start run size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label StartRun 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label NRCount 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog OSFile 
      Left            =   3600
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox LLegend 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "MainForm.frx":074C
      Top             =   1920
      Width           =   255
   End
   Begin VB.ListBox Results 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   2520
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid RGrid 
      Height          =   3495
      Left            =   3480
      TabIndex        =   6
      Top             =   1440
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   400
      BackColorBkg    =   -2147483633
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label RLegend 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P  R  O  J  E  C  T  S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnFileExport 
         Caption         =   "&Export"
         Begin VB.Menu mnFileExportTheBest 
            Caption         =   "&The best result"
         End
         Begin VB.Menu mnFileExportAll 
            Caption         =   "10 &best results"
         End
         Begin VB.Menu mnFileExportDatabase 
            Caption         =   "&All results"
         End
      End
      Begin VB.Menu mnFileS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnGrid 
      Caption         =   "&Grid"
      Begin VB.Menu mnGridRedraw 
         Caption         =   "&Draw"
      End
      Begin VB.Menu mnGridS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnGridFile 
         Caption         =   "&File"
         Begin VB.Menu mnGridFileOpen 
            Caption         =   "&Open"
         End
         Begin VB.Menu mnGridFileSave 
            Caption         =   "&Save"
         End
         Begin VB.Menu mnGridFileS1 
            Caption         =   "-"
         End
         Begin VB.Menu mnGridFileSaveAs 
            Caption         =   "Sav&e As"
         End
      End
   End
   Begin VB.Menu mnPopulation 
      Caption         =   "&Population"
      Begin VB.Menu mnPopulationNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnPopulationS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnPopulationFile 
         Caption         =   "&File"
         Begin VB.Menu mnPopulationFileOpen 
            Caption         =   "&Open"
         End
         Begin VB.Menu mnPopulationFileSave 
            Caption         =   "&Save"
         End
         Begin VB.Menu mnPopulationFileS1 
            Caption         =   "-"
         End
         Begin VB.Menu mnPopulationFileSaveAs 
            Caption         =   "Sav&e As"
         End
      End
   End
   Begin VB.Menu mnPenalty 
      Caption         =   "Pe&nalties"
      Begin VB.Menu mnPenDefault 
         Caption         =   "&Default"
      End
      Begin VB.Menu mnEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnPenaltyS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnPenaltyFile 
         Caption         =   "&File"
         Begin VB.Menu mnPenaltyFileOpen 
            Caption         =   "&Open"
         End
         Begin VB.Menu mnPenaltyFileSave 
            Caption         =   "&Save"
         End
         Begin VB.Menu mnPenaltyFileS1 
            Caption         =   "-"
         End
         Begin VB.Menu mnPenaltyFileSaveAs 
            Caption         =   "Sav&e As"
         End
      End
   End
   Begin VB.Menu mnAnalyses 
      Caption         =   "&Analysis"
      Begin VB.Menu mnAnalysesRun 
         Caption         =   "&Run"
      End
   End
   Begin VB.Menu mnAnalysesAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub FMRate_KeyPress(KeyAscii As Integer)
On Error GoTo Err_FMRate_KeyPress

If KeyAscii = 13 Or KeyAscii = 10 Then
   KeyAscii = 0
   SendKeys "{tab}"
ElseIf Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 44 Then
   KeyAscii = 0
End If

Exit_Err_FMRate_KeyPress:
Exit Sub

Err_FMRate_KeyPress:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_FMRate_KeyPress
End Sub

Private Sub FNChromo_Change()
On Error GoTo Err_FNChromo_Change

If IsNumeric(FNChromo) Then FMRate = Format(1 / FNChromo.Text, "#0.00")
mnPenalty.Enabled = False
mnAnalyses.Enabled = False

Exit_Err_FNChromo_Change:
Exit Sub

Err_FNChromo_Change:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_FNChromo_Change
End Sub

Private Sub FNChromo_KeyPress(KeyAscii As Integer)
On Error GoTo Err_FNChromo_KeyPress

If KeyAscii = 13 Or KeyAscii = 10 Then
   KeyAscii = 0
   SendKeys "{tab}"
ElseIf Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
   KeyAscii = 0
End If

Exit_Err_FNChromo_KeyPress:
Exit Sub

Err_FNChromo_KeyPress:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_FNChromo_KeyPress
End Sub

Private Sub FNProject_KeyPress(KeyAscii As Integer)
On Error GoTo Err_FNProject_KeyPress
    
    If KeyAscii = 13 Or KeyAscii = 10 Then
           KeyAscii = 0
           SendKeys "{tab}"
    ElseIf Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
           KeyAscii = 0
    End If
    
Exit_Err_FNProject_KeyPress:
Exit Sub

Err_FNProject_KeyPress:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_FNProject_KeyPress
End Sub

Private Sub FNRun_KeyPress(KeyAscii As Integer)
On Error GoTo Err_FNRun_KeyPress

If KeyAscii = 13 Or KeyAscii = 10 Then
   KeyAscii = 0
   SendKeys "{tab}"
ElseIf Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
   KeyAscii = 0
End If

Exit_Err_FNRun_KeyPress:
Exit Sub

Err_FNRun_KeyPress:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_FNRun_KeyPress
End Sub

Private Sub FNStudent_Change()
On Error GoTo Err_FNStudent_Change

mnPenalty.Enabled = False
mnAnalyses.Enabled = False

Exit_Err_FNStudent_Change:
Exit Sub

Err_FNStudent_Change:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_FNStudent_Change
End Sub

Private Sub FNStudent_KeyPress(KeyAscii As Integer)
On Error GoTo Err_FNStudent_KeyPress
    
    If KeyAscii = 13 Or KeyAscii = 10 Then
           KeyAscii = 0
           SendKeys "{tab}"
    ElseIf Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
           KeyAscii = 0
    End If
    
Exit_Err_FNStudent_KeyPress:
Exit Sub

Err_FNStudent_KeyPress:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_FNStudent_KeyPress
End Sub

Private Sub Form_Load()
On Error GoTo Err_Form_Load

    'Set default options
    mnFileExport.Enabled = False
    mnGridFileSave.Enabled = False
    mnGridFileSaveAs.Enabled = False
    mnPopulation.Enabled = False
    mnPopulationFileSave.Enabled = False
    mnPopulationFileSaveAs.Enabled = False
    mnPenalty.Enabled = False
    mnPenaltyFileSave.Enabled = False
    mnPenaltyFileSaveAs.Enabled = False
    mnAnalyses.Enabled = False
    PCurrentTime = 0
    BestFitTime = 0
    CGridFile = ""
    CPopulationFile = ""
    CPenaltyFile = ""
    USaveGrid = True
    USavePopulation = True
    USavePenalty = True
    CPenaltyFile = "Untitled"
        
    'Default penalty weights
    Penalty(1) = 1: Penalty(2) = 4: Penalty(3) = 9: Penalty(4) = 16: Penalty(5) = 25
    Penalty(6) = 36: Penalty(7) = 49: Penalty(8) = 64: Penalty(9) = 81: Penalty(10) = 100
    Penalty(11) = 100: Penalty(12) = 300
    RGrid.Cols = 0
    RGrid.Rows = 0

Exit_Err_Form_Load:
Exit Sub

Err_Form_Load:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_Form_Load
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Form_Resize

'Resize the grid
If Me.Height < 6585 Then Me.Height = 6585
If Me.Width < 6585 Then Me.Width = 6585

RGrid.Width = Me.ScaleWidth - 3620
RGrid.Height = Me.ScaleHeight - 1540
RGrid.Top = 1440
RGrid.Left = 3480
LLegend.Top = (RGrid.Height / 2) - (LLegend.Height / 4)
RLegend.Left = (RGrid.Left + (RGrid.Width / 2)) - (RLegend.Width / 4)
'Resize the listbox
Results.Height = Me.ScaleHeight - 2580

Exit_Err_Form_Resize:
Exit Sub

Err_Form_Resize:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_Form_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Form_Unload

If MsgBox("Are you sure want quit the system?", vbYesNo + vbQuestion, "The End.") = vbNo Then
   Cancel = True
Else
   If Not USaveGrid Then 'Verify if the user saves the grid
      If MsgBox("A change was made in the grid." & Chr(13) & "Do you want to save it?", vbYesNo + vbQuestion, "Saving Grid's Information.") = vbYes Then Call mnGridFileSave_Click
   End If
   
   If Not USavePopulation Then 'Verify if the user saves the population
      If MsgBox("A change was made in the population." & Chr(13) & "Do you want to save it?", vbYesNo + vbQuestion, "Saving Population's Information.") = vbYes Then Call mnPopulationFileSave_Click
   End If
   
   If Not USavePenalty Then 'Verify if the user saves the penalty
      If MsgBox("A change was made to the penalties." & Chr(13) & "Do you want to save it?", vbYesNo + vbQuestion, "Saving Penalty Information.") = vbYes Then Call mnPenaltyFileSave_Click
   End If
   
   Erase MyChromo
   Erase Penalty
   End
End If

Exit_Err_Form_Unload:
Screen.MousePointer = 0
Exit Sub

Err_Form_Unload:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_Form_Unload
End Sub

Private Sub mnAnalysesAbout_Click()
On Error GoTo Err_mnAnalysesAbout_Click

    About.Show 1

Exit_Err_mnAnalysesAbout_Click:
Exit Sub

Err_mnAnalysesAbout_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnAnalysesAbout_Click
End Sub
Private Sub mnAnalysesRun_Click()
On Error GoTo Err_mnAnalysesRun_Click
Dim i As Long, Fit As Single
Dim j As Long, TNStudent As Long, TNProject As Long, TNChromo As Long, TNRun As Long, TMRate As Single, MyUpdate As Long
'Verify the grid values
If (FNStudent = "" Or FNProject = "" Or FNChromo = "" Or FNRun = "" Or FMRate = "") Or (Not IsNumeric(FNStudent) Or Not IsNumeric(FNProject) Or Not IsNumeric(FNChromo) Or Not IsNumeric(FNRun) Or Not IsNumeric(FMRate)) Then
   MsgBox "Invalid number of parameters!" & Chr(13) & "Please inform a numeric value to: Students, Projects, Chromos, Mutation rate and Runs.", vbInformation
   Exit Sub
End If

If Not VerifyGrid Then Exit Sub

MyUpdate = 100
Screen.MousePointer = 11
Results.Clear
TNStudent = CLng(FNStudent)
TNProject = CLng(FNProject)
TNChromo = CLng(FNChromo)
TNRun = CLng(FNRun)
TMRate = CSng(FMRate)

'Run GA in the population
StartFit.Caption = BCurrent
StartRun.Caption = CurrentRun.Caption
StartFit.Refresh
RTime = Time()
StopRun.Visible = True
DoEvents

j = 0
Do While (j < TNRun) And (StopRun.Visible)
    Call BTour(TNStudent, TNProject, TNChromo, TMRate)
    If (j Mod MyUpdate) = 0 Then
        NRCount.Caption = j
        CurrentRun.Caption = CLng(StartRun.Caption) + j
        WorstFit.Caption = WCurrent
        CurrentFit.Caption = BCurrent
        DoEvents
        'NRCount.Refresh
    End If
    If StopRun.Visible Then j = j + 1
Loop
PCurrentTime = PCurrentTime + DateDiff("s", RTime, Time())
NRCount.Caption = ""
CurrentRun.Caption = CLng(StartRun.Caption) + j
USavePopulation = False

'Write the best result in the list box
Call WriteBestOnScreen(TNStudent, BPos)
StopRun.Visible = False

Exit_Err_mnAnalysesRun_Click:
Screen.MousePointer = 0
Exit Sub

Err_mnAnalysesRun_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnAnalysesRun_Click
End Sub

Private Sub mnEdit_Click()
On Error GoTo Err_mnEdit_Click

    PenaltyForm.Show 1

Exit_Err_mnEdit_Click:
Exit Sub

Err_mnEdit_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnEdit_Click
End Sub

Private Sub mnFileExit_Click()
On Error GoTo Err_mnFileExit_Click

  Unload Me

Exit_Err_mnFileExit_Click:
Exit Sub

Err_mnFileExit_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnFileExit_Click
End Sub

Private Sub mnFileExportAll_Click()
On Error GoTo Err_mnFileExportAll_Click
    
    If VerifyGrid Then Call ExportText(2)

Exit_Err_mnFileExportAll_Click:
Exit Sub

Err_mnFileExportAll_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnFileExportAll_Click
End Sub

Private Sub mnFileExportDatabase_Click()
On Error GoTo Err_mnFileExportDatabase_Click

    If VerifyGrid Then Call ExportText(3)

Exit_Err_mnFileExportDatabase_Click:
Exit Sub

Err_mnFileExportDatabase_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnFileExportDatabase_Click
End Sub

Private Sub mnFileExportTheBest_Click()
On Error GoTo Err_mnFileExportTheBest_Click

    If VerifyGrid Then Call ExportText(1)

Exit_Err_mnFileExportTheBest_Click:
Exit Sub

Err_mnFileExportTheBest_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnFileExportTheBest_Click
End Sub

Private Sub mnGridFileOpen_Click()
On Error GoTo Err_mnGridFileOpen_Click
    
    Call OpenSaveGridFile(2)
    mnGridFileSave.Enabled = True
    mnGridFileSaveAs.Enabled = True
    mnPopulation.Enabled = True
    Results.Clear
    
Exit_Err_mnGridFileOpen_Click:
Exit Sub

Err_mnGridFileOpen_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnGridFileOpen_Click
End Sub

Private Sub mnGridFileSave_Click()
On Error GoTo Err_mnGridFileSave_Click

   If VerifyGrid Then Call OpenSaveGridFile(1)
       
Exit_Err_mnGridFileSave_Click:
Exit Sub

Err_mnGridFileSave_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnGridFileSave_Click
End Sub

Private Sub mnGridFileSaveAs_Click()
On Error GoTo Err_mnGridFileSaveAs_Click

Dim Ntemp As String
If CGridFile <> "Untitled" Then
   Ntemp = CGridFile
   CGridFile = "Untitled"
End If
Call mnGridFileSave_Click
If CGridFile = "Untitled" Then CGridFile = Ntemp

Exit_Err_mnGridFileSaveAs_Click:
Exit Sub

Err_mnGridFileSaveAs_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnGridFileSaveAs_Click
End Sub

Private Sub mnGridRedraw_Click()
On Error GoTo Err_mnGridRedraw_Click
Dim i As Long, j As Long

If (FNStudent = "" Or FNProject = "") Or (Not IsNumeric(FNStudent) Or Not IsNumeric(FNProject)) Then
   MsgBox "Invalid number of parameters!" & Chr(13) & "Please inform a numeric value to: Students and Projects.", vbInformation
   Exit Sub
End If

If (FNStudent < 2 Or FNProject < 2) Then
   MsgBox "Invalid number of Students or Projects!" & Chr(13) & "The number of Students and Projects must be at least two to each.", vbInformation
   Exit Sub
End If

RGrid.Rows = CLng(FNStudent) + 1
RGrid.Cols = CLng(FNProject) + 2
RGrid.FixedRows = 1
RGrid.FixedCols = 1

RGrid.Clear
RGrid.Redraw = False

RGrid.RowHeight(0) = 450

For j = 0 To CLng(FNProject) + 1
    RGrid.Row = 0
    RGrid.Col = j
    If j = 0 Then
        RGrid.ColWidth(j) = 800
        RGrid.CellAlignment = 4
        RGrid.CellFontBold = True
        RGrid.CellForeColor = RGB(0, 0, 128)
    ElseIf j = 1 Then
        RGrid.Text = "Priority Weights"
        RGrid.ColWidth(j) = 720
        RGrid.WordWrap = True
        RGrid.CellAlignment = 4
    Else
        RGrid.Text = j - 1
        RGrid.ColWidth(j) = 500
        RGrid.CellAlignment = 4
    End If
Next j

For i = 1 To CLng(FNStudent)
    RGrid.Row = i
    RGrid.Col = 0
    RGrid.Text = i
    RGrid.CellAlignment = 4
    
    RGrid.Col = 1
    RGrid.Text = Format(1, "#0.00")
Next i

RGrid.Row = 1
RGrid.Redraw = True

mnGridFileSave.Enabled = True
mnGridFileSaveAs.Enabled = True
mnPopulation.Enabled = True
mnPenalty.Enabled = False
CGridFile = "Untitled"
Call WriteTittle

Exit_Err_mnGridRedraw_Click:
Exit Sub

Err_mnGridRedraw_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnGridRedraw_Click
End Sub

Private Sub mnPenaltyFileOpen_Click()
On Error GoTo Err_mnPenaltyFileOpen_Click
    
    Call OpenSavePenaltyFile(2)
    mnPenaltyFileSave.Enabled = True
    mnPenaltyFileSave.Enabled = True
    
Exit_Err_mnPenaltyFileOpen_Click:
Exit Sub

Err_mnPenaltyFileOpen_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnPenaltyFileOpen_Click
End Sub

Private Sub mnPenaltyFileSaveAs_Click()
On Error GoTo Err_mnPenaltyFileSaveAs_Click

Dim Ntemp As String
If CPenaltyFile <> "Untitled" Then
   Ntemp = CPenaltyFile
   CPenaltyFile = "Untitled"
End If
Call mnPenaltyFileSave_Click
If CPenaltyFile = "Untitled" Then CPenaltyFile = Ntemp

Exit_Err_mnPenaltyFileSaveAs_Click:
Exit Sub

Err_mnPenaltyFileSaveAs_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnPenaltyFileSaveAs_Click
End Sub

Private Sub mnPenaltyFileSave_Click()
On Error GoTo Err_mnPenaltyFileSave_Click
   
   If VerifyGrid Then Call OpenSavePenaltyFile(1)
   
Exit_Err_mnPenaltyFileSave_Click:
Exit Sub

Err_mnPenaltyFileSave_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnPenaltyFileSave_Click
End Sub

Private Sub mnPenDefault_Click()
On Error GoTo Err_mnPenDefault_Click
Dim i As Long, j As Long, k As Long, Fit As Long
Screen.MousePointer = 11
    'Default penalty weights
    Penalty(1) = 1: Penalty(2) = 4: Penalty(3) = 9: Penalty(4) = 16: Penalty(5) = 25
    Penalty(6) = 36: Penalty(7) = 49: Penalty(8) = 64: Penalty(9) = 81: Penalty(10) = 100
    Penalty(11) = 100: Penalty(12) = 300
    BCurrent = 999999999: BPos = 0: WCurrent = 0: WPos = 0
    For i = 1 To FNChromo
        Fit = 0
        For j = 1 To FNStudent
            'Fit calculation
            If RGrid.TextMatrix(j, MyChromo(i, j) + 1) <> "" Then
                Fit = Fit + ((Penalty(CLng(RGrid.TextMatrix(j, MyChromo(i, j) + 1)))) * (CSng(MainForm.RGrid.TextMatrix(j, 1))))
            Else
                Fit = Fit + (Penalty(11) * (CSng(RGrid.TextMatrix(j, 1))))
            End If
            For k = 1 To j - 1
                If MyChromo(i, k) = MyChromo(i, j) Then
                    Fit = Fit + (Penalty(12) * (CSng(RGrid.TextMatrix(j, 1))))
                    Exit For
                End If
            Next k
        Next j
    
        MyChromo(i, FNStudent + 1) = Fit
    
        If Fit < BCurrent Then BCurrent = Fit: BPos = i
        If Fit > WCurrent Then WCurrent = Fit: WPos = i
    Next i
  
  WorstFit.Caption = WCurrent
  CurrentFit.Caption = BCurrent
  StartFit.Caption = BCurrent
  NRCount.Refresh
  Results.Clear
  Call WriteBestOnScreen(FNStudent, BPos)
  Screen.MousePointer = 0

Exit_Err_mnPenDefault_Click:
Exit Sub

Err_mnPenDefault_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnPenDefault_Click
End Sub

Private Sub mnPopulationFileOpen_Click()
On Error GoTo Err_mnPopulationFileOpen_Click
   
   If VerifyGrid Then
      Results.Clear
      Call OpenSavePopulationFile(2)
      mnFileExport.Enabled = True
      mnPopulationFileSave.Enabled = True
      mnPopulationFileSaveAs.Enabled = True
      mnAnalyses.Enabled = True
      mnPenalty.Enabled = True
      mnPenaltyFileSaveAs.Enabled = True
   End If

Exit_Err_mnPopulationFileOpen_Click:
Exit Sub

Err_mnPopulationFileOpen_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnPopulationFileOpen_Click
End Sub

Private Sub mnPopulationFileSave_Click()
On Error GoTo Err_mnPopulationFileSave_Click
   
   If VerifyGrid Then Call OpenSavePopulationFile(1)
   
Exit_Err_mnPopulationFileSave_Click:
Exit Sub

Err_mnPopulationFileSave_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnPopulationFileSave_Click
End Sub

Private Sub mnPopulationFileSaveAs_Click()
On Error GoTo Err_mnPopulationFileSaveAs_Click

Dim Ntemp As String
If CPopulationFile <> "Untitled" Then
   Ntemp = CPopulationFile
   CPopulationFile = "Untitled"
End If
Call mnPopulationFileSave_Click
If CPopulationFile = "Untitled" Then CPopulationFile = Ntemp

Exit_Err_mnPopulationFileSaveAs_Click:
Exit Sub

Err_mnPopulationFileSaveAs_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnPopulationFileSaveAs_Click
End Sub

Private Sub mnPopulationNew_Click()
On Error GoTo Err_mnPopulationNew_Click
'Create the new population
If (FNStudent = "" Or FNProject = "" Or FNChromo = "") Or (Not IsNumeric(FNStudent) Or Not IsNumeric(FNProject) Or Not IsNumeric(FNChromo)) Then
   MsgBox "Invalid number of parameters!" & Chr(13) & "Please inform a numeric value to: Students, Projects and Chromos.", vbInformation
   Exit Sub
End If

If Not USavePopulation Then 'Verify if the user saves the population
   If MsgBox("A change was made in the population." & Chr(13) & "Do you want to save it?", vbYesNo + vbQuestion, "Saving Population's Information.") = vbYes Then Call mnPopulationFileSave_Click
End If

If VerifyGrid Then
   Results.Clear
   Call Chromo(CLng(FNStudent), CLng(FNProject), CLng(FNChromo))
   StartRun.Caption = 0
   CurrentRun.Caption = 0
   StartFit.Caption = BCurrent
   WorstFit.Caption = WCurrent
   CurrentFit.Caption = BCurrent
   mnFileExport.Enabled = True
   mnPopulationFileSave.Enabled = True
   mnPopulationFileSaveAs.Enabled = True
   mnAnalyses.Enabled = True
   mnPenalty.Enabled = True
   mnPenaltyFileSaveAs.Enabled = True
   USavePopulation = False
   CPopulationFile = "Untitled"
   Call WriteTittle
End If

Exit_Err_mnPopulationNew_Click:
Exit Sub

Err_mnPopulationNew_Click:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_mnPopulationNew_Click
End Sub

Private Sub Results_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Results_KeyPress
Dim T As Long
If KeyAscii = 3 Then
   For T = 0 To Results.ListCount - 1
      If Results.Selected(T) Then Clipboard.SetText (Clipboard.GetText & Chr(13) & Results.List(T))
   Next T
End If

Exit_Err_Results_KeyPress:
Exit Sub

Err_Results_KeyPress:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_Results_KeyPress
End Sub

Private Sub RGrid_EnterCell()
    If (RGrid.Col > 1) Then
        RGrid.TextMatrix(0, 0) = RGrid.TextMatrix(RGrid.Row, 0) & "," & RGrid.TextMatrix(0, RGrid.Col)
    Else
        RGrid.TextMatrix(0, 0) = RGrid.TextMatrix(RGrid.Row, 0) & ",0"
    End If
End Sub

Private Sub RGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_RGrid_KeyDown
Select Case KeyCode
    Case &H8 'BACKSPACE
        If Len(RGrid.Text) > 0 Then RGrid.Text = Left(RGrid.Text, (Len(RGrid.Text) - 1))
    Case &H2E 'DEL
        If Len(RGrid.Text) > 0 Then RGrid.Text = Right(RGrid.Text, (Len(RGrid.Text) - 1))
End Select

Exit_Err_RGrid_KeyDown:
Exit Sub

Err_RGrid_KeyDown:
MsgBox "Error Detected! Number: " & Err.Number & " Description: " & Err.Description, vbCritical
Resume Exit_Err_RGrid_KeyDown
End Sub

Private Sub RGrid_KeyPress(KeyAscii As Integer)
On Error GoTo Err_RGrid_KeyPress
If (KeyAscii = 10 Or KeyAscii = 13) Then
    If RGrid.Col < RGrid.Cols - 1 Then
        RGrid.Col = RGrid.Col + 1
    ElseIf RGrid.Row < RGrid.Rows - 1 Then
        RGrid.Row = RGrid.Row + 1: RGrid.Col = 1
    Else
        RGrid.Row = 1: RGrid.Col = 1
    End If
Else
    If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 46 Then
        If RGrid.Col <> 1 And KeyAscii = 46 Then
            KeyAscii = 0
        ElseIf KeyAscii = 46 And InStr(RGrid.Text, ".") <> 0 Then
            KeyAscii = 0
        ElseIf Len(RGrid.Text) <> "0" Then
            If Val(RGrid.Text & Chr(KeyAscii)) > 10 Then
                MsgBox "Value out side range: [1 - 10].", vbInformation
            Else
                RGrid.Text = RGrid.Text & Chr(KeyAscii)
                If Left(RGrid.Text, 1) = "." Then RGrid.Text = "0" & RGrid.Text
            End If
        Else
           RGrid.Text = Chr(KeyAscii)
        End If
    End If
End If

Exit_Err_RGrid_KeyPress:
Exit Sub

Err_RGrid_KeyPress:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_RGrid_KeyPress
End Sub

Public Function Chromo(NStudent As Long, NProject As Long, NChromo As Long)
'Create new generation of chromos
On Error GoTo Err_Chromo
Dim i As Long, j As Long, k As Long, Fit As Single
ReDim MyChromo(1 To NChromo, 1 To NStudent + 1)
Screen.MousePointer = 11
Randomize
BCurrent = 999999999: BPos = 0: WCurrent = 0: WPos = 0
For i = 1 To NChromo
    Fit = 0
    For j = 1 To NStudent
        MyChromo(i, j) = Int((NProject * Rnd) + 1)
        
        Do Until RGrid.TextMatrix(j, MyChromo(i, j) + 1) <> ""
              MyChromo(i, j) = Int((NProject * Rnd) + 1)
        Loop
        
        'Fit calculation
        Fit = Fit + ((Penalty(CLng(RGrid.TextMatrix(j, MyChromo(i, j) + 1)))) * (CSng(RGrid.TextMatrix(j, 1))))
        
        For k = 1 To j - 1
           If MyChromo(i, k) = MyChromo(i, j) Then
              Fit = Fit + (Penalty(12) * (CSng(RGrid.TextMatrix(j, 1))))
              Exit For
           End If
        Next k
    Next j
    MyChromo(i, NStudent + 1) = Fit
    
    If Fit < BCurrent Then BCurrent = Fit: BPos = i
    If Fit > WCurrent Then WCurrent = Fit: WPos = i
    
Next i

PCurrentTime = 0
BestFitTime = 0

Call WriteBestOnScreen(NStudent, BPos)

Exit_Err_Chromo:
Screen.MousePointer = 0
Exit Function

Err_Chromo:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_Chromo
End Function

Public Function BTour(NStudent As Long, NProject As Long, NChromo As Long, MRate As Single)
On Error GoTo Err_BTour
Dim Chromo1 As Long, Chromo2 As Long, Chromo3 As Long, Chromo4 As Long
Dim Win1 As Long, Win2 As Long, i As Long, k As Long, Fit As Single, MyFlag As Boolean
ReDim baby(1 To NStudent + 1) As Single

Randomize
Chromo1 = Int((NChromo * Rnd) + 1)
Chromo2 = Int((NChromo * Rnd) + 1)
Chromo3 = Int((NChromo * Rnd) + 1)
Chromo4 = Int((NChromo * Rnd) + 1)

If MyChromo(Chromo1, NStudent + 1) <= MyChromo(Chromo2, NStudent + 1) Then
   Win1 = Chromo1
Else
   Win1 = Chromo2
End If

If MyChromo(Chromo3, NStudent + 1) <= MyChromo(Chromo4, NStudent + 1) Then
   Win2 = Chromo3
Else
   Win2 = Chromo4
End If

'Create a new baby
For i = 1 To NStudent
    If Rnd < MyChromo(Win1, NStudent + 1) / (MyChromo(Win1, NStudent + 1) + MyChromo(Win2, NStudent + 1)) Then
       baby(i) = MyChromo(Win2, i)
    Else
       baby(i) = MyChromo(Win1, i)
    End If
Next i

'Baby's mutation
For i = 1 To NStudent
    If Rnd <= MRate Then baby(i) = Int(((NProject * Rnd) + 1))
Next i

'Baby's fitness calculation
Fit = 0
For i = 1 To NStudent
    If RGrid.TextMatrix(i, baby(i) + 1) = "" Then
       Fit = Fit + (Penalty(11) * (CSng(RGrid.TextMatrix(i, 1))))
    Else
       Fit = Fit + ((Penalty(CLng(RGrid.TextMatrix(i, baby(i) + 1)))) * (CLng(RGrid.TextMatrix(i, 1))))
    End If
       
    For k = 1 To i - 1
        If baby(k) = baby(i) Then
           Fit = Fit + (Penalty(12) * (CSng(RGrid.TextMatrix(i, 1))))
           Exit For
        End If
    Next k
Next i
baby(NStudent + 1) = Fit

'Replacement in the population
If baby(NStudent + 1) < WCurrent Then
   'Additional rotine 06/04/99
    For i = 1 To NChromo
       MyFlag = False
       If Format(baby(NStudent + 1), "#0.0000") = Format(MyChromo(i, NStudent + 1), "#0.0000") Then
          For k = 1 To NStudent
              If baby(k) <> MyChromo(i, k) Then
                 MyFlag = True
                 Exit For
              End If
          Next k
          If MyFlag = False Then Exit Function
       End If
    Next i
    '*************************************
    
    For i = 1 To NStudent + 1
      MyChromo(WPos, i) = baby(i)
    Next i
    
    'Loof for the best in the population
    If baby(NStudent + 1) < BCurrent Then
        BCurrent = baby(NStudent + 1)
        BPos = WPos
        BestFitTime = PCurrentTime + DateDiff("s", RTime, Time())
    End If
    
    'Look for the worst and best fitness
    WCurrent = 0: WPos = 0
    For i = 1 To NChromo
       If MyChromo(i, NStudent + 1) > WCurrent Then WCurrent = MyChromo(i, NStudent + 1): WPos = i
    Next i
End If

Exit_Err_BTour:
Exit Function

Err_BTour:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_BTour
End Function

Public Function OpenSaveGridFile(Opt As Integer)
'Save or open file with grid's information
On Error GoTo Err_OpenSaveGridFile
Dim c As Long, r As Long, LineText As String, Cline As Long, MPos As Long
Select Case Opt
       Case 1 ' Save grid's file
            If CGridFile = "Untitled" Then
                OSFile.DialogTitle = "Save grid file"
                OSFile.InitDir = App.Path
                OSFile.FileName = "Untitled"
                OSFile.Filter = "GA grid file (*.grd)|*.grd"
                OSFile.DefaultExt = "grd"
                OSFile.ShowSave
            
                If OSFile.FileName = "" Then
                   MsgBox "File without a name. Please give a name and try again.", vbCritical
                   Exit Function
                End If
            
                If Dir(OSFile.FileName) <> "" Then
                   If MsgBox("File already exist. Overwrite?", vbYesNo + vbQuestion) = vbNo Then Exit Function
                End If
                CGridFile = OSFile.FileName
            End If
            
            Screen.MousePointer = 11
            Open CGridFile For Output As #1 ' Open file for output.
            'File header
            Print #1, Format(RGrid.Rows - 1, "000") & " " & Format(RGrid.Cols - 1, "000") & " " & Date & " " & Time() & " ROWS|FIELDS|DATE|TIME| FIRST FIELD 6 POSITIONS AND 3 TO EACH OTHER"
            
            For r = 1 To RGrid.Rows - 1
                For c = 1 To RGrid.Cols - 1
                    If c = 1 Then ' Save the weight
                       Print #1, IIf(RGrid.TextMatrix(r, c) = "", "0", Format(RGrid.TextMatrix(r, c), "#000.00"));
                    Else
                       Print #1, IIf(RGrid.TextMatrix(r, c) = "", "000", Format(RGrid.TextMatrix(r, c), "000"));
                    End If
                Next c
                Print #1,
            Next r
            Close #1    ' Close file.
            USaveGrid = True
            Call WriteTittle

       Case 2 ' Open grid's file
            OSFile.DialogTitle = "Open grid file"
            OSFile.InitDir = App.Path
            OSFile.FileName = ""
            OSFile.Filter = "GA grid file (*.grd)|*.grd"
            OSFile.ShowOpen
            
            If Dir(OSFile.FileName) = "" Then
               MsgBox "File not found.", vbCritical
               Exit Function
            End If
            
            Screen.MousePointer = 11
            Open OSFile.FileName For Input As #1 ' Open file for input.
            Cline = 1
            Line Input #1, LineText
            If Val(Mid(LineText, 1, 3)) <> RGrid.Rows - 1 Or Val(Mid(LineText, 5, 3)) <> RGrid.Cols - 1 Then
               FNStudent = Val(Mid(LineText, 1, 3))
               FNProject = Val(Mid(LineText, 5, 3)) - 1
               Call mnGridRedraw_Click
            End If
            
            Do While Not EOF(1)
                 Line Input #1, LineText
                 MPos = 1
                 For c = 1 To RGrid.Cols - 1
                     If c = 1 Then ' Save the weight
                        RGrid.TextMatrix(Cline, c) = IIf(Mid(LineText, MPos, 6) = "000.00", 0, Format(Mid(LineText, MPos, 6), "#0.00"))
                        MPos = MPos + 6
                     Else
                        RGrid.TextMatrix(Cline, c) = IIf(Mid(LineText, MPos, 3) = "000", "", Val(Mid(LineText, MPos, 3)))
                        MPos = MPos + 3
                     End If
                 Next c
                 Cline = Cline + 1
            Loop
            Close #1    ' Close file.
            CGridFile = OSFile.FileName
            Call WriteTittle
End Select

Exit_Err_OpenSaveGridFile:
Screen.MousePointer = 0
Exit Function

Err_OpenSaveGridFile:
If Err.Number <> 32755 Then 'Cancel open or save
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
End If
Resume Exit_Err_OpenSaveGridFile
End Function

Public Function OpenSavePenaltyFile(Opt As Integer)
'Save or open file with penalty information
On Error GoTo Err_OpenSavePenaltyFile
Dim c As Long, r As Long, LineText As String, Cline As Long, MPos As Long
Select Case Opt
       Case 1 ' Save penalty file
            If CPenaltyFile = "Untitled" Then
                OSFile.DialogTitle = "Save penalty file"
                OSFile.InitDir = App.Path
                OSFile.FileName = "Untitled"
                OSFile.Filter = "GA penalty file (*.pen)|*.pen"
                OSFile.DefaultExt = "pen"
                OSFile.ShowSave
            
                If OSFile.FileName = "" Then
                   MsgBox "File without a name. Please give a name and try again.", vbCritical
                   Exit Function
                End If
            
                If Dir(OSFile.FileName) <> "" Then
                   If MsgBox("File already exist. Overwrite?", vbYesNo + vbQuestion) = vbNo Then Exit Function
                End If
                CPenaltyFile = OSFile.FileName
            End If
            
            Screen.MousePointer = 11
            Open CPenaltyFile For Output As #1 ' Open file for output.
            
            Print #1, Penalty(1)
            Print #1, Penalty(2)
            Print #1, Penalty(3)
            Print #1, Penalty(4)
            Print #1, Penalty(5)
            Print #1, Penalty(6)
            Print #1, Penalty(7)
            Print #1, Penalty(8)
            Print #1, Penalty(9)
            Print #1, Penalty(10)
            Print #1, Penalty(11)
            Print #1, Penalty(12)
            
            Close #1    ' Close file.
            USavePenalty = True
            Call WriteTittle

       Case 2 ' Open penalty file
            OSFile.DialogTitle = "Open penalty file"
            OSFile.InitDir = App.Path
            OSFile.FileName = ""
            OSFile.Filter = "GA penalty file (*.pen)|*.pen"
            OSFile.ShowOpen
            
            If Dir(OSFile.FileName) = "" Then
               MsgBox "File not found.", vbCritical
               Exit Function
            End If
            
            Screen.MousePointer = 11
            Open OSFile.FileName For Input As #1 ' Open file for input.
                                 
            For c = 1 To 12
                Line Input #1, LineText
                Penalty(c) = CLng(LineText)
            Next c
            
            Close #1    ' Close file.
            CPenaltyFile = OSFile.FileName
            Call WriteTittle
End Select

Exit_Err_OpenSavePenaltyFile:
Screen.MousePointer = 0
Exit Function

Err_OpenSavePenaltyFile:
If Err.Number <> 32755 Then 'Cancel open or save
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
End If
Resume Exit_Err_OpenSavePenaltyFile
End Function

Public Function VerifyGrid() As Boolean
On Error GoTo Err_VerifyGrid
Dim r As Long, c As Long, Nsec As Long
Screen.MousePointer = 11
'Verify if all student have a weight
For r = 1 To RGrid.Rows - 1
    If RGrid.TextMatrix(r, 1) = "" Then
       MsgBox "Invalid grid! All student must have a priority.", vbInformation
       VerifyGrid = False
       Screen.MousePointer = 0
       Exit Function
    End If
Next r

'Verify if all student had choosen a project
For r = 1 To RGrid.Rows - 1
    Nsec = 0
    For c = 2 To RGrid.Cols - 1
        If RGrid.TextMatrix(r, c) <> "" Then Nsec = Nsec + 1
    Next c
    If Nsec = 0 Then
       MsgBox "Invalid grid! All student must choose a project.", vbInformation
       VerifyGrid = False
       Screen.MousePointer = 0
       Exit Function
    End If
Next r
VerifyGrid = True

Exit_Err_VerifyGrid:
Screen.MousePointer = 0
Exit Function

Err_VerifyGrid:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_VerifyGrid
End Function

Public Function OpenSavePopulationFile(UOpt As Integer)
'Save or open file with current population's information
On Error GoTo Err_OpenSavePopulationFile
Dim c As Long, r As Long, LineText As String, MPos As Long
Dim NoChromo As Long, NoStudents As Long, Cline As Long

Select Case UOpt
       Case 1 ' Save population's file
            If CPopulationFile = "Untitled" Then
                OSFile.DialogTitle = "Save population file"
                OSFile.InitDir = App.Path
                OSFile.FileName = "Untitled"
                OSFile.Filter = "GA population file (*.pop)|*.pop"
                OSFile.DefaultExt = "pop"
                OSFile.ShowSave
               
                If OSFile.FileName = "" Then
                   MsgBox "File without a name. Please give a name and try again.", vbCritical
                   Exit Function
                 End If
            
                 If Dir(OSFile.FileName) <> "" Then
                    If MsgBox("File already exist. Overwrite?", vbYesNo + vbQuestion) = vbNo Then Exit Function
                 End If
                 CPopulationFile = OSFile.FileName
            End If
            Screen.MousePointer = 11
            Open CPopulationFile For Output As #1 ' Open file for output.
            'File header
            'ReDim MyChromo(1 To NChromo, 1 To NStudent + 1)
            NoChromo = UBound(MyChromo, 1)
            NoStudents = UBound(MyChromo, 2)
            Print #1, "|CHROMO|STUDENTS+1|PROJECTS|RUN TIME|BEST FITNESS|POSITION|WORST FITNESS|RUNNING TIME|"
            Print #1, Format(NoChromo, "000") & " " & Format(NoStudents, "000") & " " & Format(FNProject, "000") & " " & Format(CurrentRun.Caption, "0000000000") & " " & Format(BCurrent, "000000000.00") & " " & Format(BPos, "000") & " " & Format(WCurrent, "000000000.00") & " "; Format(WPos, "000") & " " & PCurrentTime
            For r = 1 To NoChromo
                For c = 1 To NoStudents
                       If c = NoStudents Then 'Fitness
                          Print #1, Trim(MyChromo(r, c));
                       Else
                          Print #1, Format(MyChromo(r, c), "000");
                       End If
                Next c
                Print #1,
            Next r
            Close #1    ' Close file.
            USavePopulation = True
            Call WriteTittle
            
       Case 2 'Open population's file
            OSFile.DialogTitle = "Open population file"
            OSFile.InitDir = App.Path
            OSFile.FileName = ""
            OSFile.Filter = "GA population file (*.pop)|*.pop"
            OSFile.ShowOpen
            
            If Dir(OSFile.FileName) = "" Then
               MsgBox "File not found.", vbCritical
               Exit Function
            End If
            
            Screen.MousePointer = 11
            Open OSFile.FileName For Input As #1 ' Open file for input.
            Line Input #1, LineText 'Load file's header
            Line Input #1, LineText 'Load file's header
            Cline = 1
            'Update screen and publics variables
            NoChromo = Val(Mid(LineText, 1, 3))
            NoStudents = Val(Mid(LineText, 5, 3))
            If (NoStudents <> RGrid.Rows) Or (Val(Mid(LineText, 9, 3)) <> RGrid.Cols - 2) Then
               MsgBox "The number of students or projects in the grid is different of the population in this file." & Chr(13) & "Population in this file:" & Chr(13) & "Students: " & NoStudents - 1 & Chr(13) & "Projects: " & Val(Mid(LineText, 9, 3)) & Chr(13) & "Chromos: " & NoChromo & "Please verify.", vbCritical
               Screen.MousePointer = 0
               Exit Function
            Else
               FNChromo = NoChromo
               CurrentRun.Caption = Val(Mid(LineText, 13, 10))
               BCurrent = Format(Mid(LineText, 24, 12), "#0.00")
               BPos = Val(Mid(LineText, 37, 3))
               WCurrent = Format(Mid(LineText, 41, 12), "#0.00")
               WPos = Val(Mid(LineText, 54, 3))
               PCurrentTime = Mid(LineText, 58, Len(LineText) - 57)
               CurrentFit.Caption = BCurrent
               WorstFit = WCurrent
               ReDim MyChromo(1 To NoChromo, 1 To NoStudents)
            End If
            
            Do While Not EOF(1)
                 Line Input #1, LineText
                 MPos = 1
                 For c = 1 To NoStudents
                     If c = NoStudents Then
                        MyChromo(Cline, c) = CLng(Trim(Mid(LineText, MPos, Len(LineText))))
                     Else
                        MyChromo(Cline, c) = CLng(Mid(LineText, MPos, 3))
                        MPos = MPos + 3
                     End If
                 Next c
                 Cline = Cline + 1
            Loop
            Call WriteBestOnScreen(NoStudents - 1, BPos)
            Close #1    ' Close file.
            CPopulationFile = OSFile.FileName
            Call WriteTittle
End Select

Exit_Err_OpenSavePopulationFile:
Screen.MousePointer = 0
Exit Function

Err_OpenSavePopulationFile:
If Err.Number <> 32755 Then 'Cancel open or save
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
End If
Resume Exit_Err_OpenSavePopulationFile
End Function

Public Function WriteBestOnScreen(TStudent As Long, Pos As Long)
On Error GoTo Err_WriteBestOnScreen
Dim j As Long
'Write the best result in the list box
Screen.MousePointer = 11
Results.AddItem "    THE BEST RESULT     "
Results.AddItem "========================"
Results.AddItem "Student |Project| Option"
For j = 1 To TStudent
    Results.AddItem "  " & Format(j, "000") & " ---- " & Format(MyChromo(Pos, j), "000") & " ---- " & IIf(RGrid.TextMatrix(j, MyChromo(Pos, j) + 1) = "", "*", RGrid.TextMatrix(j, MyChromo(Pos, j) + 1))
Next j
Results.AddItem "------------------------"
Results.AddItem "No. of runs  -: " & CurrentRun.Caption
Results.AddItem "Running time -: " & PCurrentTime & " s"
Results.AddItem "------------------------"
Results.AddItem "Worst Fitness : " & MyChromo(WPos, TStudent + 1)
Results.AddItem "------------------------"
Results.AddItem "Best Fitness -: " & MyChromo(Pos, TStudent + 1)
Results.AddItem "Fitness time -: " & BestFitTime & " s"
Results.AddItem "========================"

Exit_Err_WriteBestOnScreen:
Screen.MousePointer = 0
Exit Function

Err_WriteBestOnScreen:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_WriteBestOnScreen
End Function

Public Function WriteTittle()
On Error GoTo Err_WriteTittle
'Update screen's Ttittle
 Dim MyTiT As String
 MyTiT = "Genetic Algorithm"
 If CGridFile <> "" Then MyTiT = MyTiT & " - Grid[" & CGridFile & "]"
 If CPopulationFile <> "" Then MyTiT = MyTiT & " - Population[" & CPopulationFile & "]"
 If CPenaltyFile <> "" Then MyTiT = MyTiT & " - Penalty[" & CPenaltyFile & "]"
 Me.Caption = MyTiT

Exit_Err_WriteTittle:
Exit Function

Err_WriteTittle:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_WriteTittle
End Function

Public Function ExportText(UOpt As Integer)
On Error GoTo Err_ExportText
Dim j As Long, i As Long, TempChromo() As Single, CBestFit As Single, NPrint As Long
OSFile.DialogTitle = "Export Results"
OSFile.InitDir = App.Path
OSFile.FileName = "Untitled"
If UOpt = 3 Then
   OSFile.Filter = "GA Comma Delimited Text(*.txt)|*.txt"
Else
   OSFile.Filter = "GA Text File (*.txt)|*.txt"
End If
OSFile.DefaultExt = "txt"
OSFile.ShowSave

If OSFile.FileName = "" Then
   MsgBox "File without a name. Please give a name and try again.", vbCritical
   Exit Function
End If

If Dir(OSFile.FileName) <> "" Then
   If MsgBox("File already exist. Overwrite?", vbYesNo + vbQuestion) = vbNo Then Exit Function
End If
Screen.MousePointer = 11
Select Case UOpt
       Case 1 'The best result
            Open OSFile.FileName For Output As #1 ' Open file for output.
            'File header
            Print #1, "GENETIC ALGORITHM RESULTS - " & Date & " " & Time
            Print #1, "=============================================="
            Print #1,
            Print #1, "         THE BEST RESULT            "
            Print #1, "-----------------------------------"
            Print #1, "Student    |    Project   |  Option"
            For j = 1 To UBound(MyChromo, 2) - 1
                Print #1, "   " & Format(j, "00") & " ------------ " & Format(MyChromo(BPos, j), "00") & " --------- " & IIf(RGrid.TextMatrix(j, MyChromo(BPos, j) + 1) = "", "*", RGrid.TextMatrix(j, MyChromo(BPos, j) + 1))
            Next j
            Print #1, "   -----------------------------------"
            Print #1, "   Number of runs: " & CurrentRun.Caption
            Print #1, "   Running time--: " & PCurrentTime & " s"
            Print #1, "   -----------------------------------"
            Print #1, "   Worst Fitness-: " & MyChromo(WPos, UBound(MyChromo, 2))
            Print #1, "   -----------------------------------"
            Print #1, "   Best Fitness--: " & MyChromo(BPos, UBound(MyChromo, 2))
            Print #1, "   Fitness time--: " & BestFitTime & " s"
            Print #1, "=============================================="
            Close #1    ' Close file.
            Screen.MousePointer = 0
       Case 2 'The 10 best results
           'Write text file
           Open OSFile.FileName For Output As #1 ' Open file for output.
           CBestFit = BCurrent
           NPrint = 0
           'Print file's hearder
           Print #1, "GENETIC ALGORITHM RESULTS - " & Date & " " & Time
           Print #1, "=============================================="
           Print #1,
           Do Until NPrint >= 10
           For i = 1 To UBound(MyChromo, 1)
                If MyChromo(i, UBound(MyChromo, 2)) = CBestFit Then
                   Print #1, "Student    |    Project   |  Option  " & NPrint + 1
                   Print #1, "--------------------------------------"
                   For j = 1 To UBound(MyChromo, 2) - 1
                       Print #1, "   " & Format(j, "00") & " ------------ " & Format(MyChromo(i, j), "00") & " --------- " & IIf(RGrid.TextMatrix(j, MyChromo(i, j) + 1) = "", "*", RGrid.TextMatrix(j, MyChromo(i, j) + 1))
                   Next j
                   Print #1, "   -----------------------------------"
                   Print #1, "   Fitness--: " & MyChromo(i, UBound(MyChromo, 2))
                   Print #1,
                   NPrint = NPrint + 1
                End If
                If NPrint = 10 Then Exit For
           Next i
           
           CBestFit = CBestFit + 1
           Loop
           Print #1,
           Print #1, "***********************************************"
           Print #1, "   Number of runs: " & CurrentRun.Caption
           Print #1, "   -----------------------------------"
           Print #1, "   Running time--: " & PCurrentTime & " s"
           Print #1, "   -----------------------------------"
           Print #1, "   Best Fitness--: " & MyChromo(BPos, UBound(MyChromo, 2))
           Print #1,
           Print #1, "=============================================="
           Close #1    ' Close file.
       Case 3 'Comma delimited database
            Open OSFile.FileName For Output As #1 ' Open file for output.
            For i = 1 To UBound(MyChromo, 1)
                For j = 1 To UBound(MyChromo, 2)
                    If j = UBound(MyChromo, 2) Then
                         Print #1, Trim(MyChromo(i, j))
                    Else
                         Print #1, MyChromo(i, j) & ",";
                    End If
                Next j
            Next i
            Close #1    ' Close file.
End Select

Exit_Err_ExportText:
Screen.MousePointer = 0
Exit Function

Err_ExportText:
If Err.Number <> 32755 Then 'Cancel open or save
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
End If
Resume Exit_Err_ExportText
End Function

Private Sub RGrid_LeaveCell()
    If RGrid.Text = "0" Or RGrid.Text = "." Then
        RGrid.Text = ""
    ElseIf Left(RGrid.Text, 1) = "." Then
        RGrid.Text = "0" & RGrid.Text
    ElseIf Left(RGrid.Text, 1) = "0" And RGrid.Col <> 1 Then
        RGrid.Text = Val(RGrid.Text)
        If RGrid.Text = "0" Then RGrid.Text = ""
    ElseIf RGrid.Col = 1 Then
        If RGrid.Text = "" Or RGrid.Text = "." Then RGrid.Text = "0"
        RGrid.Text = Format(RGrid.Text, "#0.00")
    End If
End Sub

Private Sub StopRun_Click()
    Results.SetFocus
    StopRun.Visible = False
    DoEvents
End Sub
