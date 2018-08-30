VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IterativeSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'--------Commenting out Iterative Functionality for this version--------'

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    Dim currentShtStatus As sheetStatus
    ' If Save is selected run Save Function
    If Target.Range.Address = Range("SaveIteration").Address Then
        Call PreModify(IterativeSht, currentShtStatus)
        Call SaveXML
        Call PostModify(IterativeSht, currentShtStatus)
    ElseIf Target.Range.Address = Range("IterativeOutputBrowse").Address Then
        ' If Iterativeoutputbrowse link is selected run browse for iterative output file
        Call GetIterativeOutputFilePath
    End If
    
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    
    If Not Intersect(Target, Range("ParamName")) Is Nothing Then
        Range("ParamPath").Value = Range("AA" & Application.WorksheetFunction.Match(Range("ParamName").Value, Range("Z:Z"), 0)).Value
        'Range("ParamPath").Value = Range("AA" & Application.Match(Range("ParamName").Value, Range("Z:Z"), 0)).Value
    End If
    
End Sub


' IterativeOutputFilePath Function
'
' The purpose of this function is to open a dialog
' box to allow the user to change the output file path
Private Sub GetIterativeOutputFilePath()

    Dim FOpen As Variant
    
    ' Get Open File Name using Dialog box for output file
    ChDir Application.ThisWorkbook.path
    FOpen = Application.GetSaveAsFilename(Title:="Please choose an output file path", FileFilter:="CSV file (*.csv),*.csv")
    If FOpen <> False Then IterativeSht.Range("OutputFilePath").Value = FOpen
    
End Sub
