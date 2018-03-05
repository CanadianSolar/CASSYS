VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputFileSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                                                       Climate FILE SHEET                                                             '
'------------------------------------------------------------------------------------------------------------------------------------'
' The Climate file sheet is the component of the interface where the user specifies and                                                '
' defines the Climate file that contains the data used for simulation. The worksheet is composed of four major sections:               '                                                                                                   '
' "Climate File Definition", where the user specifies the Climate file, how it is delimited and where the data starts;                   '
' "Climate Position", where the user specifies which parameter a column in the Climate file belongs to, by looking at the Climate Preview; '
' "Preview", which uses the built in Text to Columns feature to split the raw delmited file into readable columns                    '
' "Original Format", which shows the original, raw data. This data is used to generate the preview.                                  '
' The code for this sheet automatically detects TMY files (.tm2 or CSV-formatted TM3).                                               '

Option Explicit
Private Const PreviewInputLine = 26 ' The row where the input file preview begins
Private Const OriginalFormatLine = 39 ' The row where the raw input (delimited) data begins
Private Const NumInputHiders = 4 ' The number of shapes that appear when a TMY file is selected to hide user input fields

Private Sub Worksheet_Activate()

    Dim currentShtStatus As sheetStatus

    ' Resets the active sheet to the current sheet to prevent errors from .Select
    Me.Activate
    
    ' Upon activation selects the first editable cell and sets the zoom to 82 which allows the full sheet to be seen
    Range("RowsToSkip").Select
    ActiveWindow.Zoom = 82
        
    'Add comment on panel temperature cell if Use Measured Module Temperature is "True" on the losses sheet to warn the user that it is required
    Call PreModify(InputFileSht, currentShtStatus)
    
    ' This code section checks if 'Use Measured Module Temperature' is selected on the losses page
    ' If this option is selected then a comment is added in the cell for the Panel Temperature column number to show that it is a required field
    If LossesSht.Range("UseMeasuredValues").Value = "True" Then
        If InputFileSht.Range("TempPanel").Comment Is Nothing Then
            InputFileSht.Range("TempPanel").AddComment ("Required field: 'Use measured module temperature' was selected on the losses page.")
        End If
    Else
        If Not InputFileSht.Range("TempPanel").Comment Is Nothing Then
            InputFileSht.Range("TempPanel").Comment.Delete
        End If
    End If
    
    Call PostModify(InputFileSht, currentShtStatus)
    
End Sub

' WorkSheet_Change Function
'
' This function is called whenever a cell in the sheet
' changes
'
' Arguments Target - The range of the changed cell
'
' The purpose of this function is to call functions
' based on which cell is changed
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim cellName As String
    Dim validFilePath As Boolean
    Dim currentShtStatus As sheetStatus
    Dim colcell As Range
    
    Dim i As Integer
    
    ' Disable events to prevent an infinite loop when changing the worksheet
    Application.EnableEvents = False
    Call PreModify(InputFileSht, currentShtStatus)

    'If the delimeter cell, input file path is changed run the preview input function
    If Not Intersect(Target, Range("Delimeter")) Is Nothing Or Not Intersect(Target, Range("InputFilePath")) Is Nothing Then
        validFilePath = checkValidFilePath(InputFileSht, "Input", InputFileSht.Range("FullInputPath").Value)
        If validFilePath = True Then
            ' Check if the file being loaded is a tm2 file by checking the file extension (last 4 letters of the file path)
            For Each colcell In Range("InputColumnNums")
                If colcell.Value <> vbNullString Then
                    Call MsgBox("Warning. The column number inputs may not match with the new input file.", vbExclamation, "CASSYS: Check Column Numbers")
                    Exit For
                End If
            Next
            If Right(InputFileSht.Range("InputFilePath").Value, 4) = ".tm2" Then
                Call configureInputTMY(2)
            ElseIf Right(InputFileSht.Range("InputFilePath").Value, 4) = ".tm3" Then
                Call configureInputTMY(3)
            ElseIf Right(InputFileSht.Range("InputFilePath").Value, 4) = ".epw" Then
                Call configureInputEPW(1)
            Else
                For i = 1 To NumInputHiders
                    InputFileSht.Shapes("InputHide" & i).Visible = msoFalse
                Next
                InputFileSht.Range("TMYType").Value = 0
                Range("RowsToSkip").Value = 1 ' set back to a default
                PreviewInput
            End If
        Else
            For i = 1 To NumInputHiders
                InputFileSht.Shapes("InputHide" & i).Visible = msoFalse
            Next
            GoTo change_end
        End If
    End If
    
    If Not Intersect(Target, Range("InputColumnNums")) Is Nothing Or Not Intersect(Target, Range("TimeFormat")) Is Nothing Or Not Intersect(Target, Range("RowsToSkip")) Is Nothing Then
        ' Check that a file is specified before allowing column numers to be denoted
        ' Also check that a column number being inputted is valid (less than the maximum number of columns in the input file)
        If Target.count = 1 And Not Intersect(Target, Range("InputColumnNums")) Is Nothing Then
            ' If one column header number is being changed
            If Target.Value > (Range("lastInputColumn").Value) And Target.Value <> "N/A" Then
                ' If a file is selected but the user specifies a column number greater than the maximum number of columns in the input file
                If Range("lastInputColumn").Value <> 0 Then
                    Call MsgBox("Invalid input. The entered column number is greater than the number of columns in the provided input file.", vbExclamation, "CASSYS: Invalid Input")
                    Target.ClearContents
                    cellName = ExtractCellName(Target)
                    Call AddColumnHeader(InputFileSht.Range(cellName), InputFileSht.Range("Prev" & cellName))
                    GoTo change_end
                Else
                    ' If a file has not been specified in the input file path
                    Call MsgBox("Please select a valid input file before inputting a column number.", vbExclamation, "CASSYS: Invalid Input")
                    Target.ClearContents
                    GoTo change_end
                End If
            ElseIf (Range("GlobalRad").Value <> vbNullString And Range("HorIrradiance") <> vbNullString) Or (Range("GlobalRad").Value <> vbNullString And Range("Hor_Diffuse") <> vbNullString) Then
                Call MsgBox("Invalid input. CASSYS requires only one of tilted or horizontal irradiance. Please select only one.", vbExclamation, "CASSYS: Invalid Input")
                Target.ClearContents
                GoTo change_end
            End If
            
            ' If the column number was valid (the code did not stop executing before this stage) then add a column header label to the preview
            cellName = ExtractCellName(Target)
            Call AddColumnHeader(InputFileSht.Range(cellName), InputFileSht.Range("Prev" & cellName))
        Else
            ' If input file is being loaded then add all of the column headers at once
            Call AddColumnHeaders
        End If
     
        ' Preview the first and last dates from the input file
        If Not Intersect(Target, Range("TimeStamp")) Is Nothing Or Not Intersect(Target, Range("TimeStamp")) Is Nothing Then Call GetDates(InputFileSht.Range("FullInputPath").Value)
        
        ' Update the brown coloured input header line if the user changes the number of rows to skip
        If Not Intersect(Target, Range("RowsToSkip")) Is Nothing Then
            Call SplitToColumns(InputFileSht.Range("Delimeter").Value)
            Call GetDates(InputFileSht.Range("FullInputPath").Value)
        End If
        
        ' Update the time stamp format in the preview if a different time format is selected
        If Not Intersect(Target, Range("TimeFormat")) Is Nothing Then Call FormatPreviewTimeStamp
    End If
        
    ' Fill N/A if POA Irradiance if a column for POA is not selected,
    ' if selected fill the values from the O&S sheet and colour the cells.
    If Not Intersect(Target, Range("GlobalRad")) Is Nothing Then
        If Not InputFileSht.Range("GlobalRad").Value = vbNullString Then
            InputFileSht.Range("MeterTilt").Locked = False
            InputFileSht.Range("MeterAzimuth").Locked = False
            If (InputFileSht.Range("MeterTilt").Value = "" And InputFileSht.Range("MeterAzimuth").Value = "") Or (InputFileSht.Range("MeterTilt").Value = "N/A" And InputFileSht.Range("MeterAzimuth").Value = "N/A") Then
                InputFileSht.Range("MeterTilt").Value = Orientation_and_ShadingSht.Range("PlaneTilt").Value
                InputFileSht.Range("MeterAzimuth").Value = Orientation_and_ShadingSht.Range("Azimuth").Value
            End If
        ElseIf Not InputFileSht.Range("HorIrradiance").Value = vbNullString Then
            InputFileSht.Range("MeterTilt").Locked = True
            InputFileSht.Range("MeterAzimuth").Locked = True
            InputFileSht.Range("MeterTilt").Value = "N/A"
            InputFileSht.Range("MeterAzimuth").Value = "N/A"
            
        End If
    End If
        
    If Not Intersect(Target, Range("HorIrradiance")) Is Nothing Then
        If Not InputFileSht.Range("HorIrradiance").Value = vbNullString Then
            InputFileSht.Range("MeterTilt").Locked = True
            InputFileSht.Range("MeterAzimuth").Locked = True
            InputFileSht.Range("MeterTilt").Value = "N/A"
            InputFileSht.Range("MeterAzimuth").Value = "N/A"
        Else
            InputFileSht.Range("MeterTilt").Locked = False
            InputFileSht.Range("MeterAzimuth").Locked = False
            InputFileSht.Range("MeterTilt").ClearContents
            InputFileSht.Range("MeterAzimuth").ClearContents
        End If
    End If
    ' If the nominal time step is changed, re-check if the nominal time step matches what was determined from the file
    If Not Intersect(Target, Range("Interval")) Is Nothing Then
        Call CheckNomTimeStep
    End If
        
change_end:
    Call PostModify(InputFileSht, currentShtStatus)
    Application.EnableEvents = True
End Sub

' WorkSheet_FollowHyperlink
'
' This function is called whenever a hyperlink is
' clicked in the Input File page
'
' The purpose of this function is to call the InputFilePath '
' function if the browse link is clicked

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    Dim currentShtStatus As sheetStatus
    
    If Target.Range.Address = Range("InputBrowse").Address Then
        ' If inputbrowse link is selected run browse for input file
        Call GetInputFilePath
    End If
    
    'If Save link is selected run the save function
    If Target.Range.Address = Range("SaveInput").Address Then
        Call PreModify(InputFileSht, currentShtStatus)
        Call SaveXML
        Call PostModify(InputFileSht, currentShtStatus)
    End If
    
End Sub

' GetDates Function
'
' Arguments Path - the file path
'
' The purpose of this function is to get the first
' and last time stamps in the csv file and output
' it to the sheet

Sub GetDates(ByVal path As String)
    Dim List() As String ' the split string
    Dim fileNum As Integer ' the file pointer
    Dim inputLine As String ' file buffer
    Dim i As Long ' counter
    Dim first As String
    Dim last As String
    Dim second As String
    
    ' set initial values
    i = 0
    first = vbNullString
    last = vbNullString
    
    fileNum = FreeFile()
    
    ' if the time has not been assigned to a column clear the first and last dates
    If InputFileSht.Range("TimeStamp") = vbNullString Then
        InputFileSht.Range("FirstDate").Value = vbNullString
        InputFileSht.Range("SecondDate").Value = vbNullString
        InputFileSht.Range("LastDate").Value = vbNullString
        Exit Sub
    End If
    
    ' open the file
    If path = vbNullString Then Exit Sub
    Open path For Input As fileNum
    
    ' get first and last line
    While (Not EOF(fileNum))
        Line Input #fileNum, inputLine
        If i = InputFileSht.Range("RowsToSkip") Then
            first = inputLine
        End If
        If i = InputFileSht.Range("RowsToSkip") + 1 Then
            second = inputLine
        End If
        i = i + 1
    Wend
    last = inputLine
    
    ' get the values for the first and last date and assign them to their respective cells
    If first <> vbNullString Then
    List = Split(first, InputFileSht.Range("Delimeter").Value)
        InputFileSht.Range("FirstDate").Value = List(InputFileSht.Range("TimeStamp").Value - 1)
        InputFileSht.Range("FirstDate").NumberFormat = InputFileSht.Range("TimeFormat").Value
                
        List = Split(second, InputFileSht.Range("Delimeter").Value)
        InputFileSht.Range("SecondDate").Value = List(InputFileSht.Range("TimeStamp").Value - 1)
        InputFileSht.Range("SecondDate").NumberFormat = InputFileSht.Range("TimeFormat").Value
        
        List = Split(last, InputFileSht.Range("Delimeter").Value)
        InputFileSht.Range("LastDate").Value = List(InputFileSht.Range("TimeStamp").Value - 1)
        InputFileSht.Range("LastDate").NumberFormat = InputFileSht.Range("TimeFormat").Value
    End If
    
    Call CheckNomTimeStep
    
    Close #fileNum
    
End Sub

' Check Time Step
'
' Checks if the user defined nominal time step matches the time step read from the input file
Private Sub CheckNomTimeStep()

    If Not Round(((InputFileSht.Range("SecondDate").Value - InputFileSht.Range("FirstDate").Value) * 24 * 60), 2) = InputFileSht.Range("Interval").Value Then
        MsgBox ("CASSYS: The defined nominal time step value does not match the nominal time step determined from the file. Please check this value before proceeding in the Climate File Sheet.")
        InputFileSht.Range("Interval").Interior.Color = RGB(255, 0, 0)
    Else
        InputFileSht.Range("Interval").Interior.Color = RGB(176, 220, 231)
    End If

End Sub
' GetLine Function
'
' Arguments Path - the file path
'
' The purpose of this function is to open a file,
' read the first 10 lines, and output them to the
' user
Private Sub GetLine(ByVal path As String)

    Dim fileNum As Integer ' the file pointer
    Dim i As Integer ' counter
    Dim ReadLine As String ' file buffer
    
    
    fileNum = FreeFile()

    ' open the file for reading
    Open path For Input As fileNum
    
    ' Read the title row and the 10 rows after it
    For i = 1 To 11
        ' Clear the row of previous contents
        InputFileSht.Rows(i + OriginalFormatLine).Clear
        
        ' Add row Index
        InputFileSht.Cells(i + OriginalFormatLine, 2).Value = i
        
        ' Add right border to cell
        With InputFileSht.Cells(i + OriginalFormatLine, 2).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        With InputFileSht.Rows(i + OriginalFormatLine).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        ' center justify row index
        InputFileSht.Cells(i + OriginalFormatLine, 2).HorizontalAlignment = xlCenter
        
        ' If the file is not at the end read the line and output it
        If Not EOF(fileNum) Then
            Line Input #fileNum, ReadLine
            InputFileSht.Cells(i + OriginalFormatLine, 3).Value = ReadLine
            If InStr(1, ReadLine, "Date (MM/DD/YYYY),Time (HH:MM),ETR (W/m^2),ETRN (W/m^2),") Then Call configureInputTMY(3)
        End If
    Next i
    
    Close #fileNum
    
End Sub
    
' PreviewInput Function
'
' The purpose of this function is to open a csv file
' and output the first 10 lines to the user then call
' a function to split those lines based on the delimeter
' set by the user

 Sub PreviewInput()
 
    Dim currentShtStatus As sheetStatus

    Call PreModify(InputFileSht, currentShtStatus)
    
    ' Clear all data from previous input file previews
    InputFileSht.Range("lastInputColumn").Value = 0
    InputFileSht.Range("previewInputs").ClearContents
    InputFileSht.Range("previewInputs").NumberFormat = "General"
    InputFileSht.Range("InputDates").ClearContents
    
    ' Change browse path color back to white if the input file exists
    InputFileSht.Range("InputFilePath").Interior.Color = ColourWhite
    
    'Get the first 10 lines from the csv file
    GetLine InputFileSht.Range("FullInputPath").Value
    
    'Split the lines using the value in the delimeter cell
    Call SplitToColumns(InputFileSht.Range("Delimeter").Value)
    
     ' Fill in the 'first' and 'last' date preview cells
    GetDates InputFileSht.Range("FullInputPath").Value
    
    ' Format the preview dates according to the time format selection
    Call FormatPreviewTimeStamp
    
    Call AddColumnHeaders
 
    Call PostModify(InputFileSht, currentShtStatus)
    
End Sub
    
' Delimeter Function
'
' Arguments Delimeter - the string delimeter
'
'
' The purpose of this function is to split the string
' taken from rows in the file and output them into
' rows and columns with the column index above them

Sub SplitToColumns(ByVal Delimeter As String)
    Dim i As Integer
    Dim cell As Range

    For i = 1 To 11
        If i <= InputFileSht.Range("RowsToSkip").Value Then
            With InputFileSht.Rows(i + PreviewInputLine).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
        Else
            With InputFileSht.Rows(i + PreviewInputLine).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
              
        ' Add right border to cell
        With InputFileSht.Cells(i + PreviewInputLine, 2).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    Next i
        
    With InputFileSht.Rows(PreviewInputLine - 1 & ":" & PreviewInputLine).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ' Check if the data to parse is empty
    If Application.WorksheetFunction.CountA(InputFileSht.Range("C40:C50")) <> 0 Then
        Application.DisplayAlerts = False
        InputFileSht.Range("C40:C50").TextToColumns Destination:=Range("C27"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=Delimeter, FieldInfo:=Array(Array(Range("TimeStamp").Value, 5))
        Application.DisplayAlerts = True
        
        ' Finds the number of columns in the input file and inserts it into a cell named 'lastInputColumn' which is used for determining
        ' the upper bound of valid column number inputs from the user and the number of column number labels to appear in the Preview section
        If InputFileSht.Range("lastInputColumn").Value = 0 And Application.WorksheetFunction.CountA(InputFileSht.Range("previewInputs")) <> 0 Then
               For i = (PreviewInputLine + 1 + InputFileSht.Range("RowsToSkip")) To 37 Step 1
                   If Application.WorksheetFunction.CountA(InputFileSht.Range("C" & i, "T" & i)) <> 0 Then
                       InputFileSht.Range("lastInputColumn").Value = Application.WorksheetFunction.CountA(InputFileSht.Range("C" & i, "T" & i)) ' + 2
                       Exit Sub
                   End If
               Next i
        End If
    End If
    
End Sub
        
' AddColumn Header Function
'
' Arguments:
' NewHeader - the cell that contains the new column location of the header
' OldHeader - the cell that contains the old column location of the header
'
' The purpose of this function is to update the column header based on which
' cell changed and updated the column header given by the user in that cell

Sub AddColumnHeader(ByVal NewHeader As Range, ByVal OldHeader As Range)

    Dim compareCell As Range
    Dim cell As Range
    Dim i As Integer
    
    If NewHeader.Value = "" Then
        ' If the old column header is the same as the header that is being updated
        If InputFileSht.Cells(PreviewInputLine - 1, (OldHeader.Value + 2)).Value = NewHeader.Offset(0, -2).Value Then
            ' clear to old column header
            InputFileSht.Cells(PreviewInputLine - 1, (OldHeader.Value + 2)).Value = ""
        End If
        OldHeader.Value = 0
    ElseIf NewHeader.Value = "N/A" Then
        ' If the old column header is the same as the header that is being updated
        If InputFileSht.Cells(PreviewInputLine - 1, (OldHeader.Value + 2)).Value = NewHeader.Offset(0, -2).Value Then
            ' clear to old column header
            InputFileSht.Cells(PreviewInputLine - 1, (OldHeader.Value + 2)).Value = ""
        End If
        Application.EnableEvents = False
        NewHeader.Value = ""
        Application.EnableEvents = True
        OldHeader.Value = 0
    ElseIf IsNumeric(NewHeader.Value) = False Then
        ' If the Value is a string
        ' Display error message and reset cell value
        NewHeader.Value = OldHeader.Value
        MsgBox ("CASSYS: Column value cannot be a string value")
    ElseIf Not (NewHeader.Value = CInt(NewHeader.Value)) Then
        ' If the Value is a floating point number
        ' Display error message and reset cell value
        NewHeader.Value = OldHeader.Value
        MsgBox ("CASSYS: Column value cannot contain decimals")
    Else
        ' If the old column header is the same as the header that is being updated
        If InputFileSht.Cells(PreviewInputLine - 1, (NewHeader.Value + 2)).Value = NewHeader.Offset(0, -2).Value Or InputFileSht.Cells(PreviewInputLine - 1, (NewHeader.Value + 2)).Value = "" Then
            ' clear to old column header
            InputFileSht.Cells(PreviewInputLine - 1, (OldHeader.Value + 2)).Value = ""
            
            ' Add cell formatting (red color,right border and center justified)
            InputFileSht.Cells(PreviewInputLine - 1, (NewHeader.Value + 2)).Value = NewHeader.Offset(0, -2).Value
            InputFileSht.Cells(PreviewInputLine - 1, (NewHeader.Value + 2)).HorizontalAlignment = xlCenter
            InputFileSht.Cells(PreviewInputLine - 1, (NewHeader.Value + 2)).Font.Color = -16776961
            
            ' Replace old header value with new value
            OldHeader.Value = NewHeader.Value
        Else
            ' If there are two columns with the same column number
            MsgBox ("CASSYS: Error: Two input types have been assigned to the same column number in the Input File.")
            NewHeader.Value = OldHeader.Value
            
            'loop through all cells to check for duplicate matches and if found, change them to the value "N/A"
            For Each cell In InputFileSht.Range("InputColumnNums")
                For Each compareCell In InputFileSht.Range("InputColumnNums")
                    If cell.Address <> compareCell.Address And cell.Value = compareCell.Value Then
                        cell.Value = vbNullString
                    End If
                Next
            Next
        End If
    End If
      
End Sub
        
' InputFilePath Function
'
' The purpose of this function is to open a dialog
' box to allow the user to change the input file path
' NB: edited so that clicking the browse link still gives relative file path 01/02/16
Sub GetInputFilePath()
    
    Dim FOpen As Variant
    Dim FilePathLeft As String
    Dim FilePathLeft_csyx As String
    Dim currentShtStatus As sheetStatus
    
    FilePathLeft = Left(FOpen, Len(ThisWorkbook.path))
    FilePathLeft = Replace(FilePathLeft, "/", "\")
    FilePathLeft_csyx = Left(FOpen, Len(Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\"))))
    FilePathLeft_csyx = Replace(FilePathLeft_csyx, "/", "\")
    
    ' Get Open File Name using Dialog box for input file
    ChDir Application.ThisWorkbook.path
    FOpen = Application.GetOpenFilename(title:="Please choose an input file path", FileFilter:="CSV or TMY Input Files(*.csv;*.tm2; *.tm3; *.epw),*.csv;.tm2;.tm3;.epw," & "All Files (*.*),*.*")
    ' If FOpen is true it means that the user did not select Cancel
    If FOpen <> False Then
        Call PreModify(InputFileSht, currentShtStatus)
        Range("FullInputPath").Value = FOpen
        FilePathLeft = Left(FOpen, Len(ThisWorkbook.path))
        FilePathLeft = Replace(FilePathLeft, "/", "\")
        FilePathLeft_csyx = Left(FOpen, Len(Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\"))))
        FilePathLeft_csyx = Replace(FilePathLeft_csyx, "/", "\")

        If (Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\")) = FilePathLeft_csyx) And IntroSht.Range("LoadFilePath").Value <> "" Then
            InputFileSht.Range("InputFilePath").Value = Right(FOpen, Len(FOpen) - Len(FilePathLeft_csyx))
        ElseIf ThisWorkbook.path = FilePathLeft Then
            InputFileSht.Range("InputFilePath").Value = Right(FOpen, Len(FOpen) - Len(ThisWorkbook.path) - 1)
        Else:
            Range("InputFilePath").Value = FOpen
        End If
        Call PostModify(InputFileSht, currentShtStatus)
    End If
    
End Sub

' FormatPreviewTimeStamp Routine
'
' Updates the timestamp date format based on the user selected units

Sub FormatPreviewTimeStamp()
    
    Dim i As Integer
    Dim dateColumn As Integer
    Dim dateCell As Range
    Dim timeStampFormat As String
    
    ' If the column number for the timestamp is filled in
    If IsNumeric(InputFileSht.Range("TimeStamp").Value) And InputFileSht.Range("TimeStamp").Value <> vbNullString Then
        ' Get the time stamp format to decide how to preview the date
        timeStampFormat = InputFileSht.Range("TimeFormat").Value
        ' Get the column containing the dates
        dateColumn = InputFileSht.Range("TimeStamp").Value + 2
        
        ' Loop through the 10 preview rows
        For i = 28 To 37 Step 1
            Set dateCell = InputFileSht.Range(Cells(i, dateColumn).Address)
            
            ' Making sure that the column is actually filled with dates before formatting it
            If InStr(1, dateCell.Value, ":") <> 0 Or InStr(1, dateCell.Value, "/") <> 0 Then
                dateCell.NumberFormat = timeStampFormat
            Else
                ' The column is not correct or no data has been previewed yet
                Exit Sub
            End If
        Next i
    End If
    
End Sub

        
' Add column header function
'
' adds red text to label the column in the input file sheet preview
Sub AddColumnHeaders()
    Dim i As Integer
        
        ' add the column numbering to the preview
        For i = 3 To Range("lastInputColumn").Value + 2 Step 1
            InputFileSht.Cells(26, i).Value = i - 2
        Next i
    
        ' If the time cell is changed run the addcolumnheader to change the header of the inputted column number
        Call AddColumnHeader(Range("TimeStamp"), Range("PrevTimeStamp"))
        
        ' If the global radiation cell is changed run the addcolumnheader to change the header of the inputted column number
        Call AddColumnHeader(Range("GlobalRad"), Range("PrevGlobalRad"))
        
        ' If the ambient temperature cell is changed run the addcolumnheader to change the header of the inputted column number
        Call AddColumnHeader(Range("TempAmbient"), Range("PrevTempAmbient"))
        
        ' If the panel temperature cell is changed run the addcolumnheader to change the header of the inputted column number
        Call AddColumnHeader(Range("TempPanel"), Range("PrevTempPanel"))
        
        ' If the wind speed cell is changed run the addcolumnheader to change the header of the inputted column number
        Call AddColumnHeader(Range("WindSpeed"), Range("PrevWindSpeed"))
        
        ' If the horizontal irradiance cell is changed run the addcolumnheader to change the header of the inputted column number
        Call AddColumnHeader(Range("HorIrradiance"), Range("PrevHorIrradiance"))
        
        ' If the horizontal diffuse cell is changed run the addcolumnheader to change the header of the inputted column number
        Call AddColumnHeader(Range("Hor_Diffuse"), Range("PrevHor_Diffuse"))
    
End Sub

' Hides user input fields on the input file sheet
' and manually sets the necessary fields for TMY reading
' TMYType 2 is passed for .tm2 files and TMYType 3 is passed
' for TMY3 files

Private Sub configureInputTMY(TMYType As Integer)

    Dim i As Integer
    Dim currentShtStatus As sheetStatus
    
    ' Makes the input hiding shapes visible which blocks out the majority of the worksheet
    ' except for the browse file path
    For i = 1 To NumInputHiders
        InputFileSht.Shapes("InputHide" & i).Visible = msoTrue
    Next
    
    InputFileSht.Shapes("InputHide3").TextFrame.Characters.text = "TM" & TMYType & " file loaded. CASSYS will automatically read the file. No user input is required."

    
    ' Logs the TMYType for use by the engine
    InputFileSht.Range("TMYType").Value = TMYType
    
    ' Specify the number of rows to skip for either a TM2 or TM3 file
    If TMYType = 2 Then
        InputFileSht.Range("RowsToSkip").Value = 1
    Else
        ' TM3 files have data that begin on line 2
        InputFileSht.Range("RowsToSkip").Value = 2
    End If
    
    ' TMY files are always averaged at the end of the hour
    InputFileSht.Range("AveragedAt").Value = "End"
    InputFileSht.Range("Interval").Value = 60#
    
    ' Clear the input date preview which is unnecessary, and so that it does not remain
    ' if a new, non-TMY file is loaded
    InputFileSht.Range("InputDates").ClearContents
    
    ' Prevents the comment telling the user to put in a column number for Panel Temperature from showing up
    If Not InputFileSht.Range("TempPanel").Comment Is Nothing Then
        InputFileSht.Range("TempPanel").Comment.Delete
    End If
    
    ' Sets the default date format for TMY files
    InputFileSht.Range("TimeFormat").Value = "yyyy-MM-dd HH:mm:ss"
    
    Call PreModify(LossesSht, currentShtStatus)
    LossesSht.Range("UseMeasuredValues").Value = "False"
    Call PostModify(LossesSht, currentShtStatus)
    
End Sub

Private Sub configureInputEPW(TMYType As Integer)
    Dim i As Integer
    Dim currentShtStatus As sheetStatus
    ' will need to fix up the .csv file that will be used so that it only has the necessary columns??? maybe not
    
     If Range("Delimeter").Value <> "," Then
        Range("Delimeter").Value = ","
     End If
        
    ' Assuming that the EPW files are always in the same format
    ' Makes the input hiding shapes visible which blocks out the majority of the worksheet
    ' except for the browse file path
    For i = 2 To NumInputHiders
        InputFileSht.Shapes("InputHide" & i).Visible = msoTrue
    Next
    
    InputFileSht.Shapes("InputHide3").TextFrame.Characters.text = "EPW file loaded. CASSYS will automatically read the file. No user input is required."

    ' Logs the TMYType for use by the engine
    InputFileSht.Range("TMYType").Value = TMYType
    
    InputFileSht.Range("AveragedAt").Value = "End"
    InputFileSht.Range("Interval").Value = 60#
    
    ' Clear the input date preview which is unnecessary, and so that it does not remain if a new file is loaded
    InputFileSht.Range("InputDates").ClearContents
    
    ' Set default rows to skip
    InputFileSht.Range("RowsToSkip").Value = 8
    
    ' Clear preview as it is not needed
    InputFileSht.Range("previewInputs").ClearContents
    
End Sub


