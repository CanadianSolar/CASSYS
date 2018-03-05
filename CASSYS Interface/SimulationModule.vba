Attribute VB_Name = "SimulationModule"
'                       SIMULATION MODULE                   '
'-----------------------------------------------------------'
' This module contains the code that calls the CASSYS.exe   '
' engine that is responsible for performing the simulation. '
' It is also responsible for pre/post simulation actions    '
' such as error checking if the required fields used for    '
' simulation are present, and after simulation it copies    '
' the data written to the output file into the Detail page  '
' and calls the Summary sheet's CreateReportTable method.   '

Option Explicit
Private errMessage As String ' holds the simulation error messages in case simulation was unsuccessful
Private simErrorFatal As Range ' range containing the simulation error message (from the .exe file)
Private resultLastColumn As Long ' gets last column with data in the sheet
Private resultLastRow As Long 'gets last row with data in the result sheet


' Simulation Function
'
' The purpose of this function is to run a C# simulation
' program and update the result to the result sheet
Sub Simulation()

    Dim wsh As Object                                ' Declare Shell Object
    Set wsh = VBA.CreateObject("WScript.Shell")      ' Set as Shell
    Dim waitOnReturn As Boolean: waitOnReturn = True ' Declare shell parameters
    Dim windowStyle As Integer: windowStyle = 1      ' Window style set to hidden
    Dim runFile As Long                              ' holds error code if the shell fails
    Dim FOpen As String                              ' Declare File open path
    Dim path As Variant
    Dim simulateCurrentFile As Integer               ' If a program has already been loaded, ask if user wants to simulate loaded XML or another XML
    Dim xDoc As DOMDocument60
    Dim currentShtStatus As sheetStatus
   
    Set xDoc = New DOMDocument60
    path = IntroSht.Range("LoadFilePath").Value

    xDoc.validateOnParse = False
    Application.ScreenUpdating = False
    ' Get file to simulate

    'Application.DisplayAlerts = False
    Call SaveXML(True)
    FOpen = Application.ThisWorkbook.path & "\CASSYSTemp.csyd"
    Application.DisplayAlerts = True

    ' Simulate file

    ' Check if simulation program is in the same directory as CASSYS.xlsm
    If Not Len(Dir(Application.ThisWorkbook.path & "/CASSYS.exe")) = 0 Or Not Len(Dir(Application.ThisWorkbook.path & "/CASSYS Engine.exe")) = 0 Then
        'Check if all required fields for simulation are present before running simulation program
        If requiredFieldsFound(FOpen) = True Then
            If Not Len(Dir(Application.ThisWorkbook.path & "/CASSYS.exe")) = 0 Then
                runFile = wsh.Run(Chr(34) & Application.ThisWorkbook.path & "/CASSYS.exe " & Chr(34) & """" & FOpen & """", windowStyle, waitOnReturn)
                ResultSht.Activate
            Else
                runFile = wsh.Run(Chr(34) & Application.ThisWorkbook.path & "/CASSYS Engine.exe " & Chr(34) & """" & FOpen & """", windowStyle, waitOnReturn)
                ResultSht.Activate
            End If
        Else
            ' If required fields are missing then inform the user and give them a choice to load the file and fix the errors
            IntroSht.Activate
            Call MsgBox("The following required fields could not be found:" & vbNewLine & errMessage, vbOKOnly Or vbSystemModal Or vbCritical, "Unable to run simulation")
            errMessage = vbNullString
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Exit Sub
        End If
    Else
        ' If the simulation program could not be found then inform the user
        Call MsgBox("Simulation program could not be found. " _
        & "Please make sure CASSYS.exe is in the same directory as CASSYS.xlsm.", vbExclamation Or vbOKOnly)
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Exit Sub
    End If
   
' Update error log and results page with simulation data
    Call PreModify(ErrorSht, currentShtStatus)
    Call LoadErrorLog
    Call PostModify(ErrorSht, currentShtStatus)
    
    
    Set simErrorFatal = ErrorSht.Range("A:A").Find("FATAL", LookAt:=xlPart)
    ' Check if the simulation was able to run correctly without crashing (no fatal error)
    If Not simErrorFatal Is Nothing Then
        errMessage = Replace(simErrorFatal.Value, "Error Description: FATAL:", vbNullString)
        Call MsgBox("Simulation encountered the following errors and was unable to continue:" & vbNewLine & vbNewLine & errMessage, vbCritical, "CASSYS: Simulation unable to run")
        errMessage = vbNullString
        ErrorSht.Activate
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' If the simulation program ran properly then update the results sheet with simulation data
    Call Update(FOpen)
    Call UpdateLossDiagram
    Application.Calculate
    errMessage = vbNullString
    If ErrorSht.Range("A8").Value = vbNullString Then ErrorSht.Visible = xlSheetHidden
    Application.ScreenUpdating = True
 
End Sub

' LoadErrorLog Macro
'
' Recorded macro that retrieves the error log from the same parent directory as CASSYS and inserts it into
' the ErrorSht
Sub LoadErrorLog()

    Dim path As String
    Dim ErrLogQuery As QueryTable
    Dim pathName As String
    Dim currentShtStatus As sheetStatus
    
    Call PreModify(ErrorSht, currentShtStatus)
    
    ' First clear all cells on the simulation error sheet so that a new log can be loaded
    ErrorSht.Range(ErrorSht.Cells(7, 1), ErrorSht.Cells(ErrorSht.Rows.count, ErrorSht.Columns.count)).ClearContents
    path = Application.ThisWorkbook.path & "\ErrorLog.txt"
    
    ' If the file exists then create a data connection to import the simulation error log from the .txt file
    If Not Len(Dir(path)) = 0 And Not path = vbNullString Then
        pathName = "TEXT;" & path
        Set ErrLogQuery = ActiveSheet.QueryTables.Add( _
        Connection:=pathName, _
        Destination:=ErrorSht.Range("$A$7"))
        ErrLogQuery.Refresh (False)
        ErrLogQuery.Delete
        ErrorSht.Visible = xlSheetVisible
    End If
    
    Call PostModify(ErrorSht, currentShtStatus)

End Sub

' Update Function
'
' The purpose of this function is to copy the data
' generated by the simulation from the output file path
' to the results page
Sub Update(ByVal FOpen As Variant)

    ' Unhide the chart and results page
    Dim xDoc As DOMDocument60
    Dim xNode As IXMLDOMNode
    Dim FilePath As String
    Dim columnHeader As Range
     ' Holds the status of the sheet
    Set xDoc = New DOMDocument60  '
    Dim Wq As QueryTable
    Dim chartNum As Integer
    Dim currentShtStatus As sheetStatus
    
    xDoc.validateOnParse = False
  
    If xDoc.Load(FOpen) Then
        Application.EnableEvents = False
        ' Clear the sheet
        Call PreModify(ResultSht, currentShtStatus)
        
        Call PrintMessage("Loading Simulation Results...", MessageSht.Range("A1"))
        ResultSht.Range(ResultSht.Columns("C"), ResultSht.Columns(ResultSht.Columns.count)).ClearContents
    
        ' Getting the Ouptut file path from the Site XML sheet
        Set xNode = xDoc.SelectSingleNode("/Site/OutputFilePath")
        If Not xNode Is Nothing Then
            If xNode.HasChildNodes Then
                FilePath = Replace(xNode.ChildNodes.NextNode.NodeValue, "/", "\")
            End If
        End If
        
        Set xDoc = Nothing
        
        ' Check if the XML file returned a proper file path, if yes then load data
        If Not Len(Dir(FilePath)) = 0 And Not FilePath = vbNullString Then
            FilePath = "TEXT;" & FilePath
            
            ' Set the data connection to the output file path that was determined earlier
            Set Wq = ResultSht.QueryTables.Add( _
            Connection:=FilePath, _
            Destination:=ResultSht.Range("$C$1"))
            With Wq
                .RefreshStyle = xlInsertDeleteCells
                .TextFileParseType = xlDelimited
                .TextFileCommaDelimiter = True
                .Refresh (False)
                .Delete
            End With
            
            ' The Result Sheet columns 1,2,3 should be auto-filled with the formulae and filtering should be allowed
            resultLastRow = ResultSht.Range("D:D").Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
            If resultLastRow <> 2 Then ResultSht.Range("A2:B2").AutoFill Destination:=ResultSht.Range("A2:B" & resultLastRow)
            
            ResultSht.Range("A1:B1").EntireColumn.Hidden = True
            ' Clear date formatting on columns that are not supposed to have dates (due to residue formatting from previous simulations)
            
            ResultSht.Visible = xlSheetVisible
            
            ' Deletes the temporary file
            If Dir(FOpen) <> "" Then
                SetAttr FOpen, vbNormal
                Kill FOpen
            End If
        
            ' Make chart builder and data summary available to the user if the Results sheet is visible
            If ResultSht.Visible Then
                SummarySht.Visible = xlSheetVisible
            End If
        Else
             MsgBox "Unable to load results. Please specify a valid output file path on the Output page."
        End If
        Call PostModify(ResultSht, currentShtStatus)
        Application.EnableEvents = True
    Else
        ' File not found, inform the user
         MsgBox " Unable to load results. The site definition could not be found."
    End If
    
    MessageSht.Visible = xlSheetVeryHidden
    
    Call FormatAfterSimulation(ResultSht, 1, 2, 1)
    
    ' Clear chart builder to default values
    ChartConfigSht.Range("numYValues").Value = 1
    ChartConfigSht.Range("chartParams").Value = vbNullString
    ResultSht.Range("A1", ResultSht.Cells(1, resultLastColumn)).Name = "GraphParam"
    ChartConfigSht.Visible = True
    
      ' clear previous series
    For chartNum = 1 To 3 Step 1
        Do Until Charts("Chart" & chartNum).SeriesCollection.count = 0
            Charts("Chart" & chartNum).SeriesCollection(1).Delete
        Loop
    Next
    
    
    Application.EnableEvents = False
    On Error GoTo simCancel
    Call SummarySht.CreateReportTable(resultLastRow, resultLastColumn)
    Application.EnableEvents = True
    Exit Sub
    
simCancel:
    On Error GoTo 0
    SummarySht.Visible = xlSheetHidden
    Application.EnableEvents = True
    ResultSht.Activate

End Sub
'
' The purpose of this sub is to copy over the necessary
' loss values produced from the simulation
'
Sub UpdateLossDiagram()
    Dim currentSht As sheetStatus
    Dim xDoc As DOMDocument
    Dim xNode As IXMLDOMNode
    Dim nodePath As String
    Dim i As Integer
    Dim rootNode As Variant
    Dim kidNode As Variant
    
    Application.EnableEvents = False
    
    'LossDiagramValueSht.HorizontalGlobIrradiance
    i = 10
    
    Set xDoc = New DOMDocument
    
    
    If xDoc.Load(Application.ThisWorkbook.path & "\LossDiagramOutputs.xml") Then
        ' The document loaded, now do something with it
        Set rootNode = xDoc.DocumentElement
        ' make sure that desired sheet is accessable
        Call PreModify(LossDiagramValueSht, currentSht)
        
        ' list out values
        For Each kidNode In rootNode.ChildNodes
            LossDiagramValueSht.Cells(i, 15) = kidNode.tagName
            LossDiagramValueSht.Cells(i, 16) = kidNode.text
            i = i + 1
        Next kidNode
        
        Call PostModify(LossDiagramValueSht, currentSht)
        LossDiagramSht.Visible = xlSheetVisible
    Else
        ' The Document didnt load
        MsgBox "Unable to Load Temporary File: LossDiagramOutputs.xml"
        Exit Sub
    End If
    
    ' realease object reference to document
    Set xDoc = Nothing
    
    ' DELETE XML FILE
    If Application.ThisWorkbook.path & "\LossDiagramOutputs.xml" <> "" Then
        Kill Application.ThisWorkbook.path & "\LossDiagramOutputs.xml"
    End If
    
    Application.EnableEvents = True
    SummarySht.Activate
End Sub

' requiredFieldsFound function
'
' This function acts as the error checking
'
Function requiredFieldsFound(ByVal FOpen As String) As Boolean
    Dim xDoc As DOMDocument60
    requiredFieldsFound = True

    
   ' Checks that the important fields necessary to run the simulation are present in the XML file
   Set xDoc = New DOMDocument60
   xDoc.Load (FOpen)
    
    If Not xDoc.SelectSingleNode("//TMYType") Is Nothing Then
        If xDoc.SelectSingleNode("//TMYType").text <> 0 Then
            requiredFieldsFound = True
            Exit Function
        End If
    End If
       
    ' Check that input file path is defined and that the file exists
    ' NB: Check in directory of .csyx and CASSYS interface
    ' NB: Change current directory for simulation
    If Not xDoc.SelectSingleNode("//InputFilePath").text = vbNullString Then
        If Len((Dir$(xDoc.SelectSingleNode("//InputFilePath").text))) <> 0 Then
            requiredFieldsFound = True
        ElseIf Len((Dir$(Replace(Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\")), "\", "/") & xDoc.SelectSingleNode("//InputFilePath").text))) <> 0 Then
            requiredFieldsFound = True
            ChDir (Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\")))
        ElseIf Len(Dir$(Replace(ThisWorkbook.path, "\", "/") & "/" & xDoc.SelectSingleNode("//InputFilePath").text)) <> 0 Then
            requiredFieldsFound = True
            ChDir (ThisWorkbook.path)
        Else
            errMessage = errMessage & vbNewLine & "Climate file path"
            requiredFieldsFound = False
        End If
    Else
         errMessage = errMessage & vbNewLine & "Climate file path"
         requiredFieldsFound = False
    End If
   
    ' Check that either panel temperature or ambient temperature is defined
    If Not SystemSht.Range("ModuleModel").Value = "Please click 'Search' to select a module." Then
        If Not xDoc.SelectSingleNode("//TempAmbient") Is Nothing Or Not xDoc.SelectSingleNode("//TempPanel") Is Nothing Then
            If (xDoc.SelectSingleNode("//TempAmbient").text = "N/A" Or xDoc.SelectSingleNode("//TempAmbient").text = vbNullString) And _
            (xDoc.SelectSingleNode("//TempPanel").text = "N/A" Or xDoc.SelectSingleNode("//TempPanel").text = vbNullString) Then
                errMessage = errMessage & vbNewLine & "Climate file column number for Ambient Temperature or Panel Temperature"
                requiredFieldsFound = False
            Else
                ' If 'Use measured module values' on the losses sheet is "True" then a column for panel temperature must be defined.
                If LossesSht.Range("UseMeasuredValues").Value = "True" Then
                    If xDoc.SelectSingleNode("//TempPanel").text = vbNullString Or xDoc.SelectSingleNode("//TempPanel").text = "N/A" Then
                        errMessage = errMessage & vbNewLine & "Climate file column number for Panel Temperature"
                        requiredFieldsFound = False
                    End If
                End If
            End If
        End If
    End If
    
    ' Check that either Horizontal Irradiance or Global (POA) Irradiance is defined
    If Not xDoc.SelectSingleNode("//GlobalRad") Is Nothing And Not xDoc.SelectSingleNode("//HorIrradiance") Is Nothing Then
        If (xDoc.SelectSingleNode("//GlobalRad").text = "N/A" Or xDoc.SelectSingleNode("//GlobalRad").text = vbNullString) And _
        (xDoc.SelectSingleNode("//HorIrradiance").text = "N/A" Or xDoc.SelectSingleNode("//HorIrradiance").text = vbNullString) Then
            errMessage = errMessage & vbNewLine & "Climate file column number for POA Irradiance or Horizontal Irradiance"
            requiredFieldsFound = False
        End If
    End If
    
     ' Check that output file path is defined and that the file exists
    If xDoc.SelectSingleNode("//OutputFilePath").text = vbNullString Then
         errMessage = errMessage & vbNewLine & "Output file path"
         requiredFieldsFound = False
    End If
    
    Set xDoc = Nothing
    
End Function

Sub FormatAfterSimulation(ByRef simSht As Worksheet, ByVal headerRow As Integer, freezePaneColumn As Integer, freezePaneRow As Integer)
    
    Dim cell As Range
    Dim currentShtStatus As sheetStatus
    
    ' Automatically selects cell A1 of the sheet upon activation
    Range("D1").Select
    
    ActiveWindow.DisplayGridlines = True
    
    ' Format result sheet
    Call PreModify(simSht, currentShtStatus)
    
    
    For Each cell In simSht.Range(simSht.Cells(headerRow, 1), simSht.Cells(headerRow, simSht.Columns.count))
        If cell.Value <> vbNullString Then
            resultLastColumn = cell.Column
            cell.WrapText = True
            cell.HorizontalAlignment = xlCenter
            cell.VerticalAlignment = xlCenter
            cell.NumberFormat = "Comma"
       
           If InStr(1, cell.Value, "Timestamp") <> 0 Then
               cell.Offset(1, 0).ColumnWidth = 20
           ElseIf cell.EntireColumn.Hidden = False Then
               cell.EntireColumn.ColumnWidth = 13.71
               If InStr(1, cell.Value, "m2") <> 0 Then
                   With cell.Characters(InStr(1, cell.Value, "2"), 1).Font
                       .Subscript = False
                       .Superscript = True
                   End With
               End If
           End If

        Else
            Exit For
        End If
    Next
    
    Call PostModify(simSht, currentShtStatus)

End Sub


