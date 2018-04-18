Attribute VB_Name = "ToolsModule"
' The DevTools module contains functions to help with the development of the program:
' - subs to prepare the workbook for release or for work
' - subs to show range names


' This sub prepares the workbook for release
Public Sub PrepareForRelease()

    Dim Sheet As Worksheet
    
    'Hide PVSyst equivalents on transformer sheet
    TransformerSht.Range("PVSystVals").EntireRow.Hidden = True
    TransformerSht.Range("ShowHidePV").Value = "Show PVSyst Equivalents"
    
    'Remove any user additions to PV module data base
    'RemoveUserAddsToPVModuleDB
    
    ' Protect all worksheets
    EnableMacrosSht.Protect
    IntroSht.Protect
    IntroSht.Range("ModeSelect") = "Grid-Connected System" 'Change to a grid connected system
    SiteSht.Protect
    Orientation_and_ShadingSht.Protect
    BifacialSht.Protect
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
    BifacialSht.Visible = x1SheetHidden
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
    BifacialSht.Unprotect
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
    BifacialSht.Visible = x1SheetVisible
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

' ShowRangeNames provide a visual display of range names on the screen
Sub ShowRangeNames()
    
    ' Iterate over all the names in the workbook
    For Each nm In ThisWorkbook.Names
        
        ' Deal with potential errors
        On Error GoTo nextName
        If InStr(nm, "#REF!") <> 0 Then GoTo nextName                ' Ill-defined names
        If InStr(nm, "!") = 0 Then GoTo nextName                     ' No idea what those are, as I expect ranges to be of the form SheetName!RangeAddress, but they exist. Just skip them
        
        ' Find the range
        Set cl = Range(nm)
        
        ' Skip ranges not in current worksheet
        If cl.Worksheet.Name <> Application.ActiveSheet.Name Then GoTo nextName
        
        ' Find dimentions of range
        clLeft = cl.Left
        clTop = cl.Top
        clHeight = cl.Height
        clWidth = cl.Width
        
        ' Deal with ranges that are too big (such as entire rows)
        If clWidth > 1000 Then clWidth = 1000
        
        ' Add the shape
        Set s = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, clLeft, clTop, clWidth, clHeight)
        
        ' Name the shape and select it
        s.Name = "_shpNamedRng:" & nm.Name
        s.Select
        
        ' Set the text to the name of the range
        Selection.ShapeRange.TextFrame2.TextRange.Characters.text = nm.Name
        
        ' Format the text
        Selection.ShapeRange.TextFrame2.TextRange.Font.Italic = msoTrue
        Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 8
        Selection.ShapeRange.TextFrame2.MarginLeft = 2
        Selection.ShapeRange.TextFrame2.MarginTop = 2
        Selection.ShapeRange.TextFrame2.WordWrap = msoFalse
        Selection.ShapeRange.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
        Selection.ShapeRange.Fill.Visible = msoFalse
        With Selection.ShapeRange.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
            .Solid
        End With
            
        ' Format the shape
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 127, 127)
            .Weight = 2.5
        End With
        
nextName:
    Next nm
End Sub

' Delete the temporary shapes showing the range names
Sub HideRangeNames()
    For Each s In ActiveSheet.Shapes
        If InStr(s.Name, "_shpNamedRng") <> 0 Then
            s.Select
            Selection.Delete
        End If
    Next s
End Sub



