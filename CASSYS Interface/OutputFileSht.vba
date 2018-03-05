VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutputFileSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                                     OUTPUT FILE                                          '
'------------------------------------------------------------------------------------------'
' This worksheet is where the output file is defined.                                      '
' The output file is the excel workbook where the simulation data will be written.         '
' The user can select whether to Summarize, Detail, or not show certain output parameters. '
' The summary is done with a pivot table, showing sums and averages.                       '
' "Summarize": Shows data on Detail sheet and summarizes data on Summary sheet.            '
' "Detail": Shows simulated data on the Details sheet. Summary sheet remains hidden.       '
' "-": This parameter will not be shown as part of the simulated data.                     '

Option Explicit
Private Sub Worksheet_Activate()
    
    ' Resets the active sheet to the current sheet
    Me.Activate
    ' Selects the first editable cell upon activating this worksheet (the output file path)
    Range("OutputFilePath").Select
       
End Sub

' WorkSheet_Change Function
' This function is called when a cell is changed
'
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim outputFile As String ' The output file path
    Dim currentShtStatus As sheetStatus
    
    ' Events are disabled to prevent an infinite recursive loop
    Application.EnableEvents = False
    Call PreModify(OutputFileSht, currentShtStatus)

    ' Changes color of cells based on the selection of "Summarize"/"Detail"/"-"
    Call ChangeCellColour(Target)
    
    ' Calls the FormatOutputSheet subroutine in the event that rows are added or deleted
    ' It is assumed that the only reason column XFD is altered would be that a row is added or deleted
    If Not Intersect(Target, OutputFileSht.Range("XFD1:XFD1000")) Is Nothing Then
        ' Check if the affected row has a null value (this would mean that a row was added, not deleted)
        If (OutputFileSht.Range(Mid(OutputFileSht.Range("HeaderRow").Address, 2, 1) & Target.row)) = vbNullString Then
            ' If a new row was added then name it 'New Row'
            OutputFileSht.Range(Mid(OutputFileSht.Range("HeaderRow").Address, 2, 1) & Target.row).Value = "New Row " & Target.row
        End If
        If Application.ActiveSheet.Name = "OutputFile" Then
            Call FormatOutputSheet
        End If
    End If
    
    ' Checks the validity of the output file path
    If Not Intersect(Target, OutputFileSht.Range("OutputFilePath")) Is Nothing Then
        outputFile = OutputFileSht.Range("OutputFilePath").Value
        If outputFile <> vbNullString Then
            ' Check if the output file path corresponds to a .csv file
            If Not Right(outputFile, 4) = ".csv" Then
                If Not Right(outputFile, 4) = ".CSV" Then MsgBox "Output file must end in .csv."
            Else
                OutputFileSht.Range("OutputFilePath").Interior.Color = ColourWhite
            End If
        End If
     End If
     
    Call PostModify(OutputFileSht, currentShtStatus)
    
    ' Re-enable events after all actions have been completed
    Application.EnableEvents = True

End Sub

' WorkSheet_FollowHyperlink
'
' This function is called whenever a hyperlink is
' clicked in the Output File page
'
' The purpose of this function is to call the OutputFilePath
' function if the browse link is clicked
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    
    Dim currentShtStatus As sheetStatus
    If Target.Range.Address = Range("OutputBrowse").Address Then
        ' If outputbrowse link is selected run browse for output file
        Call GetOutputFilePath
    ElseIf Target.Range.Address = Range("SaveOutput").Address Then
        ' If the 'Save' link is selected then automatically save all changes so far
        Call PreModify(OutputFileSht, currentShtStatus)
        Call SaveXML
        Call PostModify(OutputFileSht, currentShtStatus)
    End If
    
End Sub

' OutputFilePath Function
'
' The purpose of this function is to open a dialog
' box to allow the user to change the output file path
Private Sub GetOutputFilePath()
    Dim currentShtStatus As sheetStatus
    Dim FOpen As Variant
    Dim FilePathLeft As String
    Dim FilePathLeft_csyx As String
    
    ' Get Open File Name using Dialog box for output file
    ChDir Application.ThisWorkbook.path
    FOpen = Application.GetSaveAsFilename(title:="Please choose an output file path", FileFilter:="CSV file (*.csv),*.csv")
    
    OutputFileSht.Range("FullOutputFile").Value = FOpen
    
    If FOpen <> False Then
        Call PreModify(OutputFileSht, currentShtStatus)
        Range("OutputFilePath").Value = FOpen
        
        ' Get the directory of the file minus the specific file name
        FilePathLeft = Left(FOpen, Len(ThisWorkbook.path))
        FilePathLeft = Replace(FilePathLeft, "/", "\")
        
        ' Get the directory of the .csyx file minus the .csyx file name
        FilePathLeft_csyx = Left(FOpen, Len(Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\"))))
        FilePathLeft_csyx = Replace(FilePathLeft_csyx, "/", "\")
        
        ' If the directory of the load file is the same as that of the csyx file
        If (Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\")) = FilePathLeft_csyx) And IntroSht.Range("LoadFilePath").Value <> "" Then
            OutputFileSht.Range("OutputFilePath").Value = Right(FOpen, Len(FOpen) - Len(FilePathLeft_csyx))
        ElseIf ThisWorkbook.path = FilePathLeft Then
            OutputFileSht.Range("OutputFilePath").Value = Right(FOpen, Len(FOpen) - Len(ThisWorkbook.path) - 1)
        Else:
            Range("OutputFilePath").Value = FOpen
        End If
        
        Call PostModify(OutputFileSht, currentShtStatus)
    End If
End Sub

' ChangeCellColour function
'
' This function is called by the WorkSheet_Change event when:
' "Summarize" is selected (Change colour to neon green)
' "Detail" is selected (Change colour to light green)
' "-" is selected (Change colour to white)

Sub ChangeCellColour(ByRef Target As Range)

    Dim cell As Range
    Dim isSectionHeader As Boolean
    Dim sectionBlock As Range

    For Each cell In Target
        ' Expression to denote the section header (the section title should not be changed)
        isSectionHeader = (cell.row = OutputFileSht.Range("Available_SectionStart").row Or cell.row = OutputFileSht.Range("Available_SectionStart").row + 1)
        
        ' The block of cells in the section that is being changed
        Set sectionBlock = Range(Cells(cell.row, OutputFileSht.Range("HeaderRow").Column), Cells(cell.row, OutputFileSht.Range("OutputParam").Column))
        
        If Not Intersect(cell, OutputFileSht.Range("OutputParam")) Is Nothing Then
            ' If the cell that was changed was a part of the 'OutputParam' box list
            If cell.Value = "Summarize" And Not isSectionHeader Then
                ' Change the colours of the cells to a neon green
                sectionBlock.Interior.Color = ColourBrightGreen
            ElseIf cell.Value = "Detail" And Not isSectionHeader Then
                ' Change the colours of the cells to a light green
                sectionBlock.Interior.Color = ColourMediumGreen
            ElseIf Not isSectionHeader Then
                ' Change the colours of the cells to white
                sectionBlock.Interior.Color = ColourWhite
            End If
        End If
    Next

End Sub
