Attribute VB_Name = "ButtonModule"
'            BUTTON MODULE             '
'--------------------------------------'
' The button module is a collection    '
' of the macros assigned to the action '
' buttons on the intro and iterative   '
' Mode sheet such as Load, Save, Save  '
' As, and Simulate.                    '

Option Explicit

Function NewButton() As Boolean
    Call ClearAll
    If Not BypassBeforeSave Then
        Sheets("Site").Activate
    End If
    
    ' For safety, make sure calculation is set to auto
    ' This should not be necessary but if for some reason calculate is not set back to auto in some other part of the program,
    ' at least opening a file or creating a new one will reset it
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Function
Function LoadButton() As Boolean
    
    Dim fileNameToLoad As String
    Dim introShtStatus As sheetStatus
    Dim errorShtStatus As sheetStatus
    
    ' Hide all irrelevant sheets
    ResultSht.Visible = xlSheetHidden
    SummarySht.Visible = xlSheetHidden
    ChartConfigSht.Visible = xlSheetHidden
    ReportSht.Visible = xlSheetHidden
    CompChart1.Visible = xlSheetHidden
    CompChart2.Visible = xlSheetHidden
    CompChart3.Visible = xlSheetHidden
    Inverter_DatabaseSht.Visible = xlSheetHidden
    PV_DatabaseSht.Visible = xlSheetHidden
    ErrorSht.Visible = xlSheetHidden
    MessageSht.Visible = xlSheetHidden
    'IterativeSht.Visible = xlSheetHidden
    LossDiagramSht.Visible = xlSheetHidden
    LossDiagramValueSht.Visible = xlSheetHidden
    
    ' Clears the error sheet and displays the message sheet
    ' to inform user that the file is loading
    Call PreModify(IntroSht, introShtStatus)
    Call PreModify(ErrorSht, errorShtStatus)
    
    ErrorSht.Rows("7:" & Rows.count).ClearContents
    ErrorSht.Columns("P:XFD").ClearContents
    
    fileNameToLoad = GetFileToLoad
    
    ' If the file could not be found then do not continue loading
    If fileNameToLoad = "" Then
        Call PostModify(ErrorSht, errorShtStatus)
        Call PostModify(IntroSht, introShtStatus)
        IntroSht.Activate
        Exit Function
    End If
    
    Call PrintMessage("Loading...", MessageSht.Range("A1"))
    
    ' Clear the data to start from a clean slate
    Call ClearAll
    
    ' Load the file
    Call Load(fileNameToLoad)
    
    
    ' Hides the "Loading" Message
    MessageSht.Visible = xlSheetHidden
    
    Call PostModify(ErrorSht, errorShtStatus)
    Call PostModify(IntroSht, introShtStatus)
    
   ' hide message sheet only if no errors were printed
    If ErrorSht.Range("ErrorsEncountered").Value = vbNullString Then
        IntroSht.Activate
        ErrorSht.Visible = xlSheetHidden
    Else
        MsgBox "Some events occurred during loading. You will be automatically redirected to the complete list of events."
        ErrorSht.Visible = xlSheetVisible
        ErrorSht.Activate
    End If
        
    ' For safety, make sure calculation is set to auto
    ' This should not be necessary but if for some reason calculate is not set back to auto in some other part of the program,
    ' at least opening a file or creating a new one will reset it
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Function
Function SaveButton() As Boolean
    
    Dim currentShtStatus As sheetStatus
    
    ' If Save link is selected save current file
    Call PreModify(IntroSht, currentShtStatus)
    Call SaveXML
    Call PostModify(IntroSht, currentShtStatus)
     
End Function
Function SaveAsButton() As Boolean

    Dim currentShtStatus As sheetStatus
        
    ' If Save As link is selected then the user is prompted to select file
    Call PreModify(IntroSht, currentShtStatus)
    IntroSht.Range("SaveFilePath").Value = vbNullString
    Call SaveXML
    Call PostModify(IntroSht, currentShtStatus)
    
End Function
' Simply ask the workbook to close
Function ExitButton() As Boolean
    
    ThisWorkbook.Close
     
End Function
Function SimulateButton() As Boolean

      ' Clears result sheet except for dates
        Dim currentShtStatus As sheetStatus
        
        Call PreModify(ResultSht, currentShtStatus)
        
        ResultSht.Rows("3:" & Rows.count).ClearContents
        ResultSht.Columns("D:XFD").ClearContents
        
        Call PostModify(ResultSht, currentShtStatus)
    
        ' Hides sheets
        ChartConfigSht.Visible = xlSheetHidden
        CompChart1.Visible = xlSheetHidden
        CompChart2.Visible = xlSheetHidden
        CompChart3.Visible = xlSheetHidden
        ErrorSht.Visible = xlSheetHidden
        ResultSht.Visible = xlSheetHidden
        SummarySht.Visible = xlSheetHidden
        LossDiagramSht.Visible = xlSheetHidden
        LossDiagramValueSht.Visible = xlSheetHidden
        
        ' If simulation link is selected run simulation
        Call Simulation
        
        ' For safety, make sure calculation is set to auto
        ' This should not be necessary but if for some reason calculate is not set back to auto in some other part of the program,
        ' at least opening a file or creating a new one will reset it
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
End Function



' When clicked on the output file page, shows information about each of the output options
Function HelpButton() As Boolean

    Call MsgBox("Export PDF Report of Site Definition: Create a PDF report containing all of the site information you have entered on each page." & _
    " This report can be used to reproduce the simulation in the future when the CSYX file is not avaiable to load." & vbNewLine & vbNewLine & "Click the drop down list next to an output parameter to choose from the following options:" & vbNewLine & vbNewLine & _
    "Summarize: Display this parameter on both the Results and Data Summary page after simulation. Note: Some parameters cannot be summarized." & vbNewLine & vbNewLine & _
    "Detail: Show simulation data for this parameter on the Results page, but do not provide a data summary." & vbNewLine & vbNewLine & _
    "'-': This parameter will not be displayed after simulation.", vbInformation, "CASSYS: Help")
    
End Function

Function ExportAsPdFButton() As Boolean
    Dim FOpen As Variant
    Application.DisplayAlerts = False
    ChDir Application.ThisWorkbook.path
    FOpen = Application.GetSaveAsFilename(Title:="Please specify the name and location of the exported PDF file.", FileFilter:="PDF file(*.pdf),*.pdf")
    ChDir Application.ThisWorkbook.path
    
    If FOpen <> False Then Call ReportSht.ExportReportToPDF(FOpen)
    Application.DisplayAlerts = True
End Function
' Adds a new output to the output file sheet upon clicking the button
'
Function InsertNewOutputButton() As Boolean
    Dim outputName As Variant
    Dim rowNum As Variant
    
    Application.EnableEvents = False
    ActiveWindow.DisplayHeadings = True
    
TryoutputName:
    outputName = Application.InputBox("Enter the new output name", "Add New output")
    
    If outputName = False Then GoTo endInsertNewOutputButton
        
TryRowNum:
    rowNum = Application.InputBox("Which row number would you like to insert the new output?", "Choose row number")
    
    If rowNum = False Then GoTo endInsertNewOutputButton
      
    If IsNumeric(Left(outputName, 1)) Then
        MsgBox "Output name cannot begin with a number."
        GoTo TryoutputName
    End If

    
    If IsNumeric(rowNum) Then
        If rowNum > OutputFileSht.Range("HeaderRow").row And rowNum < OutputFileSht.Range("FooterRow").row Then
            OutputFileSht.Cells(rowNum, OutputFileSht.Range("HeaderRow").Column).EntireRow.Insert
            OutputFileSht.Cells(rowNum, OutputFileSht.Range("HeaderRow").Column).Value = outputName
            OutputFileSht.Cells(rowNum, OutputFileSht.Range("OutputConstColumn").Column).Value = outputName
        Else
            MsgBox "Invalid row number. The row number must be within the bounds of the available output list area."
            GoTo TryRowNum
        End If
    End If
        
    
    Call FormatOutputSheet
    
endInsertNewOutputButton:
    ActiveWindow.DisplayHeadings = False
    Application.EnableEvents = True
      
End Function

'--------Commenting out Iterative Functionality for this version--------'

' Button used in OutputFileSht
' Hides output file sheet
' Unhides IterativeSht, enabling iterative mode
Function EnableIterativeModeButton() As Boolean

'' Copy output file path from output file sheet to iterative sheet
'IterativeSht.Range("OutputFilePath").Value = OutputFileSht.Range("OutputFilePath").Value
'
'IterativeSht.Visible = xlSheetVisible
'OutputFileSht.Visible = xlSheetHidden
'Sheets("Iterative Mode").Activate
MsgBox "Iterative mode has been disabled for this version of CASSYS, but will be available on the next version of CASSYS"

End Function
' Button used in IterativeSht
' Hides IterativeSht, disabling iterative mode
Function DisableIterativeModeButton() As Boolean

'--------Commenting out Iterative Functionality for this version--------'

'' Copy output file path from iterative mode sheet to output file sheet
'OutputFileSht.Range("OutputFilePath").Value = IterativeSht.Range("OutputFilePath").Value
'
'IterativeSht.Visible = xlSheetHidden
'OutputFileSht.Visible = xlSheetVisible
'
'Sheets("Output File").Activate

End Function

