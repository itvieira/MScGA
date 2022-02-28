Attribute VB_Name = "Utility"
Public MyChromo() As Single
Public PCurrentTime As Long 'Current running time to population
Public BestFitTime As Long  'Time taken to find the best fit
Public RTime As Date        'Current run statert time
Public USaveGrid As Boolean 'Verify if the user saves the grid
Public USavePopulation As Boolean 'Verify if the user saves the population
Public USavePenalty As Boolean 'Verify if the user saves the penalty
Public WCurrent As Single, WPos As Long, BCurrent As Single, BPos As Long
Public CGridFile As String
Public CPopulationFile As String
Public CPenaltyFile As String
Public Penalty(1 To 12) As Long 'Penalty weights given to choices

Sub Main()
On Error GoTo Err_Main
Dim i As Date
About.Show
About.Command1.Visible = False
About.Refresh
i = DateAdd("s", 3, Time())
Do Until i < Time()
Loop
MainForm.Show
MainForm.Refresh
Unload About

Exit_Err_Main:
Exit Sub

Err_Main:
MsgBox "Error detected!" & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbCritical, "Error detected during operation."
Resume Exit_Err_Main
End Sub
