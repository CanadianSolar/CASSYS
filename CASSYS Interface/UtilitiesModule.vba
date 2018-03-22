Attribute VB_Name = "UtilitiesModule"
Option Explicit
Option Compare Text

Public Const SubArrayHeight = 17 ' The section height of the sub-array information
Public Const PVDataHeight = 4 ' The height where the data starts in the pv module database
Public Const InvDataHeight = 2 ' The height where the data starts in the inverter database
Public Const IntroHeight = 4 ' The section height of the sub-array info on the intro sheet
Public Const ColourWhite = 16777215 ' The numerical constant representing the colour white
Public Const ColourBrightGreen = 9565049 ' The numerical constant representing bright green
Public Const ColourMediumGreen = 9762464 ' The numerical constant representing medium green
Public Const ColourThemeGreen = 5296276 ' The green used for headers


' ClearAll Function
'
' The purpose of this function is to clear
' the cell values to their default values
' and activate site sheet to start a new
' site definition
Sub ClearAll()

    Dim chartNum As Integer
    Dim introShtStatus As sheetStatus
    Dim currentShtStatus As sheetStatus
    
    ' Disable events to speed things up
    Application.EnableEvents = False
        
    ' Hide all irrelevant sheets that should only appear after simulation or loading
    SummarySht.Visible = xlSheetHidden
    ResultSht.Visible = xlSheetHidden
    ChartConfigSht.Visible = xlSheetHidden
    ReportSht.Visible = xlSheetHidden
    CompChart1.Visible = xlSheetHidden
    CompChart2.Visible = xlSheetHidden
    CompChart3.Visible = xlSheetHidden
    Inverter_DatabaseSht.Visible = xlSheetHidden
    PV_DatabaseSht.Visible = xlSheetHidden
    ErrorSht.Visible = xlSheetHidden
    MessageSht.Visible = xlSheetHidden
    IterativeSht.Visible = xlSheetHidden
    'OutputFileSht.Visible = xlSheetHidden

     ' Clear intro sheet
     Call PreModify(IntroSht, introShtStatus)
     IntroSht.Range("LoadFilePath").Value = vbNullString
     IntroSht.Range("SaveFilePath").Value = vbNullString
     
     
     ' Clear site sheet
     Call PreModify(SiteSht, currentShtStatus)
     
     SiteSht.Range("Name,Region,Country,City").Value = vbNullString
    
     ' Set site values to default
     Application.EnableEvents = True
     SiteSht.Range("Latitude").Value = 0
     SiteSht.Range("Longitude").Value = 0
     SiteSht.Range("Altitude").Value = 0
     SiteSht.Range("TimeZone").Value = 0
     SiteSht.Range("AlbJan,AlbFeb,AlbMar,AlbApr,AlbMay,AlbJun,AlbJul,AlbAug,AlbSep,AlbOct,AlbNov,AlbDec").Value = 0.2
     SiteSht.Range("AlbYearly").Value = 0.2
     Application.EnableEvents = False
     
     Call PostModify(SiteSht, currentShtStatus)
     
     ' Clear system page
     ' Set to 1 sub array
     
     Call PreModify(SystemSht, currentShtStatus)
     
     Application.EnableEvents = True
     SystemSht.Range("NumSubArray").Value = 1
     Application.EnableEvents = False
     SystemSht.Range("PVDataIndex").Value = -1
     SystemSht.Range("InvDataIndex").Value = -1
     
     SystemSht.Range("ModStr").Value = 1
     SystemSht.Range("NumStr").Value = 1
     SystemSht.Range("NumInv").Value = 1
          
     SystemSht.Range("PVLossFrac").Value = 0.015
     SystemSht.Range("ACWiringLossAtSTC").Value = "at STC"
     SystemSht.Range("InvLossFrac").Value = 0.015
     
     Call PostModify(SystemSht, currentShtStatus)
     
     ' Clear transformer sheet
     
     Call PreModify(TransformerSht, currentShtStatus)
     TransformerSht.Range("PNomTrf").Value = 0
     TransformerSht.Range("PIronLossTrf").Value = 0
     TransformerSht.Range("PFullLoadLss").Value = 0
     TransformerSht.Range("NightlyDisconnect").Value = "True"
     TransformerSht.Range("PIronLoss").Value = 0
     TransformerSht.Range("FIronLoss").Value = 0
     TransformerSht.Range("FResLoss").Value = 0
     
     Call PostModify(TransformerSht, currentShtStatus)
     
     'clear ASTM sheet
     Call PreModify(AstmSht, currentShtStatus)
     AstmSht.Range("SystemPmax").Value = 1
     AstmSht.Range("ASTMCoeffs").Value = 1
     AstmSht.Range("AstmMonthList").Value = 1
     Call PostModify(AstmSht, currentShtStatus)
     
     ' Clear orientation and shading sheet
     
     Call PreModify(Orientation_and_ShadingSht, currentShtStatus)
     
     ' Clear fixed tilted plane
     Orientation_and_ShadingSht.Range("PlaneTiltFix").Value = 30
     Orientation_and_ShadingSht.Range("AzimuthFix").Value = 0
     
    ' Clear Seasonal Adjustment info
    Orientation_and_ShadingSht.Range("AzimuthSeasonal").Value = 0
    Orientation_and_ShadingSht.Range("PlaneTiltSummer").Value = 30
    Orientation_and_ShadingSht.Range("PlaneTiltWinter").Value = 30
    Orientation_and_ShadingSht.Range("SummerMonth").Value = "Mar"
    Orientation_and_ShadingSht.Range("WinterMonth").Value = "Nov"
    Orientation_and_ShadingSht.Range("SummerDay").Value = 1
    Orientation_and_ShadingSht.Range("WinterDay").Value = 1
    
    ' Clear Unlimited Rows info
    Orientation_and_ShadingSht.Range("PlaneTilt").Value = 30
    Orientation_and_ShadingSht.Range("AzimuthFix").Value = 0
    Orientation_and_ShadingSht.Range("Pitch").Value = 1
    Orientation_and_ShadingSht.Range("CollBandWidth").Value = 1
    Orientation_and_ShadingSht.Range("TopInactive").Value = 0
    Orientation_and_ShadingSht.Range("BottomInactive").Value = 0
    Orientation_and_ShadingSht.Range("RowsBlock").Value = 1
    Orientation_and_ShadingSht.Range("BacktrackOptSAET").Value = "No"
    Orientation_and_ShadingSht.Range("BacktrackOptSAST").Value = "No"
    
    ' Clear Tracker info
    Orientation_and_ShadingSht.Range("RowsBlockSAET").Value = 1
    Orientation_and_ShadingSht.Range("WActiveSAET").Value = 1
    Orientation_and_ShadingSht.Range("MinTiltSAET").Value = -90
    Orientation_and_ShadingSht.Range("MaxTiltSAET").Value = 90
    Orientation_and_ShadingSht.Range("StrInWidSAET").Value = 4
    Orientation_and_ShadingSht.Range("CellSizeSAET").Value = 15.6
    
    Orientation_and_ShadingSht.Range("RowsBlockSAST").Value = 1
    Orientation_and_ShadingSht.Range("WActiveSAST").Value = 1
    Orientation_and_ShadingSht.Range("RotationMaxSAST").Value = 90
    Orientation_and_ShadingSht.Range("StrInWidSAST").Value = 4
    Orientation_and_ShadingSht.Range("CellSizeSAST").Value = 15.6
    
    Orientation_and_ShadingSht.Range("AxisTiltTART").Value = 30
    Orientation_and_ShadingSht.Range("AxisAzimuthTART").Value = 0
    Orientation_and_ShadingSht.Range("RotationMinTART").Value = -90
    Orientation_and_ShadingSht.Range("RotationMaxTART").Value = 90
    
    Orientation_and_ShadingSht.Range("PlaneTiltAVAT").Value = 30
    Orientation_and_ShadingSht.Range("MinAzimuthAVAT").Value = -90
    Orientation_and_ShadingSht.Range("MaxAzimuthAVAT").Value = 90
    
    Orientation_and_ShadingSht.Range("MinTiltTAXT").Value = 0
    Orientation_and_ShadingSht.Range("MaxTiltTAXT").Value = 90
    Orientation_and_ShadingSht.Range("MinAzimuthTAXT").Value = -90
    Orientation_and_ShadingSht.Range("MaxAzimuthTAXT").Value = 90
    
     
    ' Set orientation type
    Application.EnableEvents = True
    Orientation_and_ShadingSht.Range("OrientType").Value = "Fixed Tilted Plane"
    Application.EnableEvents = False
    
    Call PostModify(Orientation_and_ShadingSht, currentShtStatus)
    
    
     ' Clear Horizon Shading Sheet
     ' Sets the horizon to True with zero points
     ' Clear Horizon Shading Sheet
     ' Sets the horizon to True with zero points
    Call PreModify(Horizon_ShadingSht, currentShtStatus)
    Application.EnableEvents = True
    Range("DefineHorizonProfile").Value = "Yes"
    Call Horizon_ShadingSht.ClearHorizon
    Application.EnableEvents = False
    Call PostModify(Horizon_ShadingSht, currentShtStatus)
    
    ' Clear losses sheet
    Call PreModify(LossesSht, currentShtStatus)
    
    LossesSht.Range("ConsHLF").Value = 20
    LossesSht.Range("ConvHLF").Value = 0
    LossesSht.Range("UseMeasuredValues").Value = False

    LossesSht.Range("EfficiencyLoss").Value = 0.009
    LossesSht.Range("PowerLoss").Value = 0.02
    LossesSht.Range("LossFixedVoltage").Value = 0
    
    Application.EnableEvents = True
    LossesSht.Range("IAMSelection").Value = "ASHRAE"
    LossesSht.Range("bNaught").Value = 0.05
    Dim aoi As Integer
    Dim Iam As Double
    For aoi = 0 To 90 Step 5
        If aoi = 90 Then
            Iam = 0
        Else
            Iam = 1 - 0.05 * ((1 / Cos(WorksheetFunction.Pi / 180 * aoi)) - 1)
        End If
        LossesSht.Range("IAM_" & aoi).Value = Iam
    Next aoi
    
    Application.EnableEvents = False
    
    Call PostModify(LossesSht, currentShtStatus)

    ' set soiling loss calculation
     Call PreModify(SoilingSht, currentShtStatus)
     
     Application.EnableEvents = True
     SoilingSht.Range("SfreqVal") = "Yearly"
     Application.EnableEvents = False
     
     ' Clear soiling losses
     ' Clear yearly losses
     SoilingSht.Range("Yearly").Value = 0
     
     ' Clear monthly losses
     SoilingSht.Range("Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec").Value = 0
     
     Call PostModify(SoilingSht, currentShtStatus)
     
    ' Disable spectral model
     Call PreModify(SpectralSht, currentShtStatus)
     
     Application.EnableEvents = True
     SpectralSht.Range("UseSpectralModel") = "No"
     Application.EnableEvents = False
          
     ' Clear spectral modification values losses
     SpectralSht.Range("ktCorrectionValues").Value = 0
     
     Call PostModify(SpectralSht, currentShtStatus)
     
     ' Clear InputFilePath and Output file
     
     Call PreModify(InputFileSht, currentShtStatus)
     
     ' Hide shapes that block input when TMY file is loaded
     Dim i As Integer
     For i = 1 To 4
         InputFileSht.Shapes("InputHide" & i).Visible = msoFalse
     Next
     
     InputFileSht.Range("lastInputColumn").Value = 0
     InputFileSht.Range("InputFilePath").Value = vbNullString
     InputFileSht.Range("Interval").Value = 60
     InputFileSht.Range("FullInputPath").Value = vbNullString
     InputFileSht.Range("IncorrectClimateRowsAllowed").Value = 0
     
     ' Clear column headers, column selections and reset colours
     InputFileSht.Range("TimeStamp,GlobalRad,TempAmbient,TempPanel,WindSpeed,HorIrradiance,Hor_Diffuse,FirstDate,LastDate,previewInputs,MeterTilt,MeterAzimuth").ClearContents
     InputFileSht.Range("MeterTilt").Interior.ColorIndex = 0
     InputFileSht.Range("MeterAzimuth").Interior.ColorIndex = 0
     InputFileSht.Range("InputFilePath").Interior.Color = ColourWhite
     
     ' Clears colour formatting on input file sheet of brown 'rows to skip lines'
     InputFileSht.Range("RowsToSkip").Value = 1
     InputFileSht.Range("TMYType").Value = 0
     Call InputFileSht.SplitToColumns(",")
     Call PostModify(InputFileSht, currentShtStatus)
     
     Call PreModify(OutputFileSht, currentShtStatus)
    
     OutputFileSht.Range("OutputFilePath").Value = vbNullString
     
     Call PostModify(OutputFileSht, currentShtStatus)
    
     ' Clear Results Sheet
     Call PreModify(ResultSht, currentShtStatus)
     
     ' Clear all the auto-filled entries from a previous run
     ResultSht.Range(ResultSht.Columns("D"), ResultSht.Columns(ResultSht.Columns.count)).ClearContents
     ResultSht.Range(ResultSht.Rows(3), ResultSht.Rows(ResultSht.Rows.count)).ClearContents
        
     Call PostModify(ResultSht, currentShtStatus)
     
     ' Report sheet
     Call PreModify(ReportSht, currentShtStatus)
     
    ' Delete all data from previous reports
     ReportSht.Range("A1", "N" & ReportSht.Range("B" & Rows.count).End(xlUp).row).Delete
     Dim graph As Shape
     For Each graph In ReportSht.Shapes
        graph.Delete
     Next graph
     
     Call PostModify(ReportSht, currentShtStatus)
     
     ' Simulation error log
     Call PreModify(ErrorSht, currentShtStatus)
     
     ErrorSht.Rows("7:" & Rows.count).ClearContents
     ErrorSht.Columns("P:XFD").ClearContents
     
     Call PostModify(ErrorSht, currentShtStatus)
     
'--------Commenting out Iterative Functionality for this version--------'
     ' Iterative Mode Sheet
     
'     Call PreModify(IterativeSht, currentShtStatus)
'
'     IterativeSht.Range("ParamName").ClearContents
'     IterativeSht.Range("Start").ClearContents
'     IterativeSht.Range("End").ClearContents
'     IterativeSht.Range("Interval").ClearContents
'     IterativeSht.Range("OutputFilePath").Value = ""
'
'     Call PostModify(IterativeSht, currentShtStatus)
     
     Call PreModify(OutputFileSht, currentShtStatus)
     
     OutputFileSht.CheckBoxes("ControlAllChkBox").LockedText = False
     OutputFileSht.CheckBoxes("ControlAllChkBox").text = "-"
     Call ControlAllChkBox_Click
     
     'NB: Making sure all outputs go to "-", even if hidden
     If IntroSht.Range("ModeSelect").Value = "Radiation Mode" Then
         Call PVArrayChkBox_Click
         Call InverterChkBox_Click
         Call SystemLossesPerfChkBox_Click
         Call EfficienciesChkBox_Click
     End If
     
     OutputFileSht.Range("OutputFilePath").Interior.Color = ColourWhite
     OutputFileSht.Range("OutputFilePath").Value = vbNullString
     Call PostModify(OutputFileSht, currentShtStatus)
     
     SummarySht.Range(SummarySht.Cells(12, 1), SummarySht.Cells(SummarySht.Rows.count, SummarySht.Columns.count)).Delete
     SummarySht.Range("ViewDays").Value = "Monthly"
     MessageSht.Cells.ClearContents
     ' Restore intro sheet status
     
    Application.EnableEvents = True
    ' Clear chart builder to default values
    ChartConfigSht.Range("numYValues").Value = 1
    ChartConfigSht.Range("chartParams").Value = vbNullString
    
      ' clear previous series
    For chartNum = 1 To 3 Step 1
        Do Until Charts("Chart" & chartNum).SeriesCollection.count = 0
            Charts("Chart" & chartNum).SeriesCollection(1).Delete
        Loop
    Next
    
    
    Application.EnableEvents = False
     
     IntroSht.Activate
     Application.Calculate
     Call PostModify(IntroSht, introShtStatus)
     
    Application.EnableEvents = True
     
End Sub

' SaveAsPDF function
'
' The purpose of this function is to save the
' report sheet as a PDF file
Sub SaveAsPDF()
    Dim FSave As Variant ' Holds the file path to be saved to
    FSave = Application.GetSaveAsFilename(title:="Save As", FileFilter:="PDF file (*.pdf),*.pdf", InitialFileName:="CASSYS-Report")
    'Save as XML file
    If Not FSave = False Then
        ReportSht.ExportAsFixedFormat xlTypePDF, FSave
    End If
End Sub


' This sub prepares the workbook for release
Public Sub PrepareForRelease()

    Dim Sheet As Worksheet
    
    'Hide PVSyst equivalents on transformer sheet
    TransformerSht.Range("PVSystVals").EntireRow.Hidden = True
    TransformerSht.Range("ShowHidePV").Value = "Show PVSyst Equivalents"
    
    'Remove any user additions to PV module data base
    RemoveUserAddsToPVModuleDB
    
    ' Protect all worksheets
    EnableMacrosSht.Protect
    IntroSht.Protect
    IntroSht.Range("ModeSelect") = "Grid-Connected System" 'Change to a grid connected system
    SiteSht.Protect
    Orientation_and_ShadingSht.Protect
    Horizon_ShadingSht.Protect
    SystemSht.Protect
    LossesSht.Protect
    SoilingSht.Protect
    SpectralSht.Protect
    TransformerSht.Protect
    InputFileSht.Protect
    OutputFileSht.Unprotect
    ChartConfigSht.Protect
    ErrorSht.Protect
    AstmSht.Protect
    IterativeSht.Protect
    LossDiagramSht.Protect
    LossDiagramValueSht.Protect
    
    ' Hide gridlines and headings
    Application.ScreenUpdating = False
    For Each Sheet In ThisWorkbook.Worksheets
        Sheet.Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
    Next Sheet
    Application.ScreenUpdating = True
    
    ' Hide worksheets
    EnableMacrosSht.Visible = xlSheetVisible
    IntroSht.Visible = xlSheetHidden
    SiteSht.Visible = xlSheetHidden
    Orientation_and_ShadingSht.Visible = xlSheetHidden
    Horizon_ShadingSht.Visible = xlSheetHidden
    SystemSht.Visible = xlSheetHidden
    LossesSht.Visible = xlSheetHidden
    SoilingSht.Visible = xlSheetHidden
    SpectralSht.Visible = xlSheetHidden
    TransformerSht.Visible = xlSheetHidden
    InputFileSht.Visible = xlSheetHidden
    OutputFileSht.Visible = xlSheetHidden
    ResultSht.Visible = xlSheetHidden
    SummarySht.Visible = xlSheetHidden
    ReportSht.Visible = xlSheetHidden
    ChartConfigSht.Visible = xlSheetHidden
    CompChart1.Visible = xlSheetHidden
    CompChart2.Visible = xlSheetHidden
    CompChart3.Visible = xlSheetHidden
    ErrorSht.Visible = xlSheetHidden
    Inverter_DatabaseSht.Visible = xlSheetHidden
    PV_DatabaseSht.Visible = xlSheetHidden
    MessageSht.Visible = xlSheetHidden
    AstmSht.Visible = xlSheetHidden
    IterativeSht.Visible = xlSheetHidden
    LossDiagramSht.Visible = xlSheetHidden
    LossDiagramValueSht.Visible = xlSheetHidden
    
    ' Run the New function
    Call ClearAll
    
    ' Put focus on intro sheet
    EnableMacrosSht.Activate
    
End Sub

' This sub prepares the workbook for work
Public Sub PrepareForWork()

    Dim Sheet As Worksheet
    
    ' Unprotect all worksheets
    EnableMacrosSht.Unprotect
    IntroSht.Unprotect
    SiteSht.Unprotect
    Orientation_and_ShadingSht.Unprotect
    Horizon_ShadingSht.Unprotect
    SystemSht.Unprotect
    LossesSht.Unprotect
    SoilingSht.Unprotect
    SpectralSht.Unprotect
    TransformerSht.Unprotect
    InputFileSht.Unprotect
    OutputFileSht.Unprotect
    ResultSht.Unprotect
    ReportSht.Unprotect
    ChartConfigSht.Unprotect
    ErrorSht.Unprotect
    AstmSht.Unprotect
    IterativeSht.Unprotect
    LossDiagramSht.Unprotect
    LossDiagramValueSht.Unprotect
    
    ' Unhide gridlines and headings
    Application.ScreenUpdating = False
    For Each Sheet In ThisWorkbook.Worksheets
        Sheet.Activate
        ActiveWindow.DisplayGridlines = True
        ActiveWindow.DisplayHeadings = True
    Next Sheet
    Application.ScreenUpdating = True
    
    ' Show hidden sheets
    EnableMacrosSht.Visible = xlSheetVisible
    IntroSht.Visible = xlSheetVisible
    SiteSht.Visible = xlSheetVisible
    Orientation_and_ShadingSht.Visible = xlSheetVisible
    SystemSht.Visible = xlSheetVisible
    LossesSht.Visible = xlSheetVisible
    SoilingSht.Visible = xlSheetVisible
    SpectralSht.Visible = xlSheetVisible
    TransformerSht.Visible = xlSheetVisible
    InputFileSht.Visible = xlSheetVisible
    OutputFileSht.Visible = xlSheetVisible
    ResultSht.Visible = xlSheetVisible
    ReportSht.Visible = xlSheetVisible
    ChartConfigSht.Visible = xlSheetVisible
    CompChart1.Visible = xlSheetVisible
    CompChart2.Visible = xlSheetVisible
    CompChart3.Visible = xlSheetVisible
    ErrorSht.Visible = xlSheetVisible
    Inverter_DatabaseSht.Visible = xlSheetVisible
    PV_DatabaseSht.Visible = xlSheetVisible
    AstmSht.Visible = xlSheetVisible
    IterativeSht.Visible = xlSheetVisible
    LossDiagramSht.Visible = xlSheetVisible
    LossDiagramValueSht.Visible = xlSheetVisible
    
    ' Put focus on enable macros sheet
    EnableMacrosSht.Activate
  
End Sub

' Print a message to the Message sheet
Public Sub PrintMessage(Msg As String, Optional ByVal printLocation As Range)

    Dim IsScreenUpdating As Boolean
        
    IsScreenUpdating = Application.ScreenUpdating
    MessageSht.Visible = xlSheetVisible
    MessageSht.Activate
    Application.ScreenUpdating = True
    printLocation.Value = Msg
    printLocation.WrapText = False
    Application.ScreenUpdating = IsScreenUpdating
    
End Sub

' checkValidFilepath Function
'
' Checks the file path on the input file or output file sheets
' to ensure that they are valid when loaded
' NB: edited so the input file path is not written as part of the output file path 02/02/2016
Function checkValidFilePath(fileSheet As Worksheet, ByVal pathLabel As String, ByVal FilePath As String) As Boolean
    
    Dim currentShtStatus As sheetStatus
    Dim isValidFilePath As Boolean
    
    On Error GoTo fileNotFound:
    Call PreModify(fileSheet, currentShtStatus)
    
    isValidFilePath = Len((Dir$(FilePath))) <> 0 And (Right(FilePath, 4) = ".csv" Or (Right(FilePath, 4) = ".tm2" Or Right(FilePath, 4) = ".tm3" Or Right(FilePath, 4) = ".epw" And pathLabel = "Input") Or (Right(FilePath, 1) = "/" And pathLabel = "Output"))
    If isValidFilePath Then
        Range(pathLabel & "FilePath").Interior.Color = ColourWhite
        'NB: edited so the input file path is not written as part of the output file path
        If pathLabel = "Input" Then Range(pathLabel & "FilePath").Value = Replace(InputFileSht.Range("InputFilePath").Value, "/", "\")
        checkValidFilePath = True
    Else
        ' Clear residue inputs and preview if the file is incorrect
fileNotFound:
        If pathLabel = "Input" Then
            InputFileSht.Range("previewInputs").ClearContents
            InputFileSht.Range("InputColumnNums").ClearContents
            Range(pathLabel & "FilePath").Interior.Color = RGB(255, 0, 0)
        End If
        checkValidFilePath = False
    End If
    
    Call PostModify(fileSheet, currentShtStatus)

    
End Function
' Quicksort function
'
' Source: http://www.blueclaw-db.com/quick-sort.htm
'
' The purpose of this function is to efficiently sort a list
Public Sub QuickSort(strArray() As Variant, intBottom As Integer, intTop As Integer)
    Dim strPivot As Variant, strTemp As Variant
    Dim intBottomTemp As Integer, intTopTemp As Integer

    intBottomTemp = intBottom
    intTopTemp = intTop

    strPivot = strArray((intBottom + intTop) \ 2)

    While (intBottomTemp <= intTopTemp)

        '  comparison of the values is a descending sort
        While (strArray(intBottomTemp) < strPivot And intBottomTemp < intTop)
            intBottomTemp = intBottomTemp + 1
        Wend

        While (strPivot < strArray(intTopTemp) And intTopTemp > intBottom)
            intTopTemp = intTopTemp - 1
        Wend
        
        If intBottomTemp < intTopTemp Then
            strTemp = strArray(intBottomTemp)
            strArray(intBottomTemp) = strArray(intTopTemp)
            strArray(intTopTemp) = strTemp
        End If

        If intBottomTemp <= intTopTemp Then
            intBottomTemp = intBottomTemp + 1
            intTopTemp = intTopTemp - 1
        End If
  
    Wend

    ' the function calls itself until everything is in good order
    If (intBottom < intTopTemp) Then QuickSort strArray, intBottom, intTopTemp
    If (intBottomTemp < intTop) Then QuickSort strArray, intBottomTemp, intTop
    
End Sub

' The purpose of this function is to
' extract a named cell's name without its sheet label

Public Function ExtractCellName(ByRef cell As Range)
    
    Dim cellName As String
    
    'string manipulation to store the cell name
    'the .Name.Name property returns unnecessary infomation about the sheet, so then name must be extracted
    cellName = cell.Name.Name
    cellName = Right(cellName, Len(cellName) - InStr(cellName, "!"))
    ExtractCellName = cellName
    
End Function

' The purpose of this subroutine is to
' delete temporarily loaded sub-arrays
' from the inverter and PV module
' databases.

Public Sub DeleteTempArrays(ByRef databaseSht As Worksheet)
    
    Dim Rng As Range
    Dim IsScreenUpdating As Boolean
    
    IsScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    On Error GoTo skipdelete
    
    ' Filter all User_Added ones (these are the temporarily loaded ones)
    databaseSht.UsedRange.AutoFilter field:=1, Criteria1:="=User_Added"
    
    ' This Set action will produce an error if no cells were filtered, and the code will go to the skipdelete line
    Set Rng = databaseSht.AutoFilter.Range.Offset(1, 0).Resize(databaseSht.AutoFilter.Range.Rows.count - 1, 1).SpecialCells(xlCellTypeVisible)
    
    ' Delete the filtered User_Added cells
    Application.DisplayAlerts = False
    databaseSht.UsedRange.Offset(1, 0).Resize(databaseSht.UsedRange.Rows.count - 1).Rows.Delete
    Application.DisplayAlerts = True
    
    ' Restore visibility of the hidden cells from filtering
skipdelete:    databaseSht.Cells.AutoFilter
    Application.ScreenUpdating = IsScreenUpdating

        

End Sub

' NB: causes next button to select next unhidden sheet regardless of mode
Sub NextButton()

Dim i As Integer
Dim currSheet As String
Dim nextSheet As String
Dim visibleCounter As Integer
Dim passedCurrSheet As Boolean

For i = 1 To Worksheets.count
    If Worksheets(i).Visible = True Then
    If Worksheets(ActiveSheet.Index).Name = Worksheets(i).Name Then
    passedCurrSheet = True
    ElseIf passedCurrSheet = True Then
    nextSheet = Worksheets(i).Name
    Exit For
    End If
    End If
Next i
    
Sheets(nextSheet).Activate
    
End Sub

' NB: Back button selects previous unhidden sheet regardless of mode
Sub PrevButton()

Dim i As Integer
Dim currSheet As String
Dim nextSheet As String
Dim visibleCounter As Integer
Dim passedCurrSheet As Boolean

For i = Worksheets.count To 1 Step -1
    If Worksheets(i).Visible = True Then
    If Worksheets(ActiveSheet.Index).Name = Worksheets(i).Name Then
    passedCurrSheet = True
    ElseIf passedCurrSheet = True Then
    nextSheet = Worksheets(i).Name
    Exit For
    End If
    End If
Next i
    
Sheets(nextSheet).Activate

End Sub

' The SaveCodeModules sub is used to export all VBA modules to text files, for version control purposes
' The files are stored in the ./VBA directory
' Credits: code adapted from http://stackoverflow.com/questions/131605/best-way-to-do-version-control-for-ms-excel
Sub SaveCodeModules()

Dim i As Integer, sName As String

On Error GoTo ErrHandler

With ThisWorkbook.vbproject
    For i% = 1 To .VBComponents.count
        sName = .VBComponents(i%).CodeModule.Name
        .VBComponents(i%).Export Application.ActiveWorkbook.path & "/VBA/" & sName & ".vba"
    Next i
End With

On Error GoTo 0
Exit Sub

ErrHandler:
Resume

End Sub

'The purpose of this sub is to go through all PV module database inputs and remove any user additions

Private Sub RemoveUserAddsToPVModuleDB()
    Dim lastRow As Integer
    Dim i As Integer
    Dim cellVal As String
    
    ' Find the last unoccupied row in the PV Module Database
    lastRow = PV_DatabaseSht.Range("A" & Rows.count).End(xlUp).row + 1
    
    'In database the first entry is in row 5
    i = 5
    
    ' Loop through all inputs and if Origin value contains the word 'user' then delete the row
    Do Until i = lastRow
        cellVal = PV_DatabaseSht.Range("A" & i).Value
        If (InStr(cellVal, "User") <> 0) Then
            Range("A" & i).EntireRow.Delete
            lastRow = lastRow - 1  ' Decrease the amount of rows to check after deleting one
        Else
            i = i + 1
        End If
    Loop
End Sub

