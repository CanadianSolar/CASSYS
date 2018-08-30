VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReportSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                   REPORT SHEET                  '
'-------------------------------------------------'
' The ExportReportToPDF function creates a report '
' using data from each worksheet and formats it   '
' so that it exports correctly to PDF.            '

Option Explicit
Public currentDocLength As Integer ' updating ReportSht length: relies on data in column B
Public lastShapeRow As Integer     ' used in GetLastShapeInRow and indicates the largest row number in which there is a shape/chart
Public shapeNum As Integer         ' Number of shapes in the current report sheet
Public rowsPerPage As Integer      ' indicates the current number of rows which are being used in the current pdf page
Public rowsAdded As Integer        ' indicates the number of rows in a sheet being added to the report sheet
Private Const pdfPageSize = 61     ' set number of rows that fit nicely in a pdf page

Sub ExportReportToPDF(ByVal fileName As String)
    Dim currentShtStatus As sheetStatus
    Dim i As Integer          ' integer used for looping
    Dim dataEndRow As Integer ' integer used to indicate the last row in ReportSht containing data
    Dim graph As Shape        ' used to represent shape objects in the report sheet and delete them
       
'    On Error GoTo fileAlreadyOpen
     
    ' Reset row count for pdf sheet
    rowsPerPage = 0
    
    ' Reset lastShapeRow
    lastShapeRow = 0
    
    ' Reset all the page breaks just in case they have been incorrectly defined
    ReportSht.ResetAllPageBreaks

    ' Speed up code running
    Call PreModify(ReportSht, currentShtStatus)
    
    ' Delete pictures in the sheet, charts copied to this page are considered pictures
    For Each graph In ReportSht.Shapes
        graph.Delete
    Next graph
    
    ' Delete all data from previous reports
    ReportSht.Range("A1", "N" & ReportSht.Range("B" & Rows.count).End(xlUp).row).Delete
    
    ' Copy, paste and format the various sheet sections on the report sheet
    SetPrintFormatSite
    If IntroSht.Range("ModeSelect").Value = "Grid-Connected System" Then
        SetPrintFormatOrient_and_Shading
        SetPrintFormatBifacial
        SetPrintFormatHorizon
        SetPrintFormatSystem
        SetPrintFormatLosses
        SetPrintFormatSoiling
        SetPrintFormatSpectral
        SetPrintFormatTransformer
        SetPrintFormatInput
        SetPrintFormatSummary
        SetPrintLossesDiagram
    ElseIf IntroSht.Range("ModeSelect").Value = "ASTM E2848 Regression" Then
        SetPrintASTM
        SetPrintFormatInput
        SetPrintFormatSummary
    ElseIf IntroSht.Range("ModeSelect").Value = "Radiation Mode" Then
        SetPrintFormatOrient_and_Shading
        SetPrintFormatHorizon
        SetPrintFormatInput
        SetPrintFormatSummary
    End If

    ' Find the last row of data which is used to determine the range to be exported to PDF
    dataEndRow = Application.Max(ReportSht.Range("B" & Rows.count).End(xlUp).row, lastShapeRow)
        
    ReportSht.Range("A" & dataEndRow + 1).PageBreak = xlPageBreakManual
    
    ' Setup how the page will look in the PDF
    With ReportSht.PageSetup
        .PrintArea = "$A$1:$N$" & dataEndRow
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .Zoom = False
    End With
    
    ' Activate the report sheet so that the ExportAsFixedFormat method can be called
    ReportSht.Visible = xlSheetVisible
    ReportSht.Activate

    On Error GoTo fileAlreadyOpen

    ReportSht.ExportAsFixedFormat _
        Type:=xlTypePDF, fileName:=fileName, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
    
    On Error GoTo 0
    GoTo normalEnd
        
fileAlreadyOpen:
    If Err.Number = -2147018887 Then
        MsgBox "A PDF of the same name is currently open. Please close the PDF file first and try again."
        Application.EnableEvents = True
'        ReportSht.Visible = xlHidden
        OutputFileSht.Activate
    Else
        MsgBox Err.Number
'        Resume Next
'        On Error GoTo 0
'        Resume
    End If

normalEnd:
    'Rehide the report sheet and restore application statuses
    ReportSht.Visible = xlSheetHidden
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Call PostModify(ReportSht, currentShtStatus)
    
    SummarySht.Activate
End Sub
'
' This function is called to copy and paste the desired
' contents from the site sheet into the report sheet
'
Private Sub SetPrintFormatSite()
    Dim currentShtStatus As sheetStatus
    Dim initialPasteRow As Integer ' the first row in which the data is pasted. Ranges are specified in offsets; changes only have to be made to the initialPasteRow to paste in a different place.
    
    initialPasteRow = 1
    
    Call PreModify(SiteSht, currentShtStatus)
    
    ' Copy over site sheet information and chart
    Call PasteHeader(SiteSht, initialPasteRow)
    Call CopyPaste(SiteSht.Range("PrintSite_Def"), ReportSht.Range("B" & initialPasteRow).Offset(2, 0))
    ' ensure that the chart will be visible
    If IntroSht.Range("ModeSelect") = "Grid-Connected System" Then
        Call CopyPasteChart(SiteSht, "AlbedoChart", Sheets("Site").ChartObjects("AlbedoChart"), ReportSht.Range("B" & initialPasteRow).Offset(22, 1))
    End If
    ' If monthly albedo is chosen, copy and paste monthly information on top of yearly
    If SiteSht.Range("AlbFreqVal").Value = "Monthly" Then
        Call CopyPaste(SiteSht.Range("PrintSite_AlbedoMonthly"), ReportSht.Range("B" & initialPasteRow).Offset(19, 0))
    End If
    
    Call PostModify(SiteSht, currentShtStatus)
    
    ' update current document length and set page break if needed
    GetLastShapeRow
    currentDocLength = Application.Max(ReportSht.Range("B" & Rows.count).End(xlUp).row, lastShapeRow)
    Call PageBreak(initialPasteRow, currentDocLength)
End Sub
'
' This function is called to copy and paste the desired
' contents from the Orientation and shading sheet into the report sheet
'
Private Sub SetPrintFormatOrient_and_Shading()
    Dim currentShtStatus As sheetStatus
    Dim initialPasteRow As Integer ' the first row in which the data is pasted. Ranges are specified in offsets; changes only have to be made to the initialPasteRow to paste in a different place.
    
    initialPasteRow = currentDocLength + 2 ' +2 to leave space between sections
    
    Call PreModify(Orientation_and_ShadingSht, currentShtStatus)
    
    Call PasteHeader(Orientation_and_ShadingSht, initialPasteRow)
    Call CopyPaste(Orientation_and_ShadingSht.Range("PrintOS_Names").SpecialCells(xlCellTypeVisible), ReportSht.Range("B" & initialPasteRow).Offset(2, 0))
    Call CopyPaste(Orientation_and_ShadingSht.Range("PrintOS_Vals").SpecialCells(xlCellTypeVisible), ReportSht.Range("B" & initialPasteRow).Offset(2, 3))
    
    Call PostModify(Orientation_and_ShadingSht, currentShtStatus)
    
    ' update current document length and set page break if needed
    GetLastShapeRow
    currentDocLength = Application.Max(ReportSht.Range("B" & Rows.count).End(xlUp).row, lastShapeRow)
    Call PageBreak(initialPasteRow, currentDocLength)
End Sub
'
' This function is called to copy and paste the desired
' contents from the Bifacial sheet into the report sheet
'
Private Sub SetPrintFormatBifacial()
    Dim currentShtStatus As sheetStatus
    Dim initialPasteRow As Integer ' The first row in which the data is pasted. Ranges are specified in offsets; changes only have to be made to the initialPasteRow to paste in a different place.
    
    initialPasteRow = currentDocLength + 2 ' +1 to leave a space between sections
    
    Call PreModify(BifacialSht, currentShtStatus)
    
    Call PasteHeader(BifacialSht, initialPasteRow)
    
    Call CopyPaste(BifacialSht.Range("PrintBifacial1"), ReportSht.Range("B" & initialPasteRow).Offset(2, 0))
    
    If BifacialSht.Range("UseBifacialModel").Value = "Yes" Then
        Call CopyPaste(BifacialSht.Range("PrintBifacial2"), ReportSht.Range("B" & initialPasteRow).Offset(4, 0))
    
        If BifacialSht.Range("BifAlbFreqVal").Value = "Yearly" Then
            Call CopyPaste(BifacialSht.Range("BifAlbYearly"), ReportSht.Range("F" & initialPasteRow).Offset(14, 0))
        ElseIf BifacialSht.Range("BifAlbFreqVal").Value = "Monthly" Then
            Call CopyPaste(BifacialSht.Range("PrintBifAlbMonthly"), ReportSht.Range("B" & initialPasteRow).Offset(14, 0))
        End If
    
        If BifacialSht.Range("BifAlbFreqVal").Value <> "Site" Then
            Call CopyPasteChart(BifacialSht, "BifAlbedoChart", Sheets("Bifacial").ChartObjects("BifAlbedoChart"), ReportSht.Range("B" & initialPasteRow).Offset(17, 1))
        End If
    End If
        
    ' update number of charts on page
    GetLastShapeRow
    If BifacialSht.Range("UseBifacialModel").Value = "Yes" And BifacialSht.Range("BifAlbFreqVal").Value <> "Site" Then
        ReportSht.Shapes(shapeNum).Width = 725
        ReportSht.Shapes(shapeNum).Cut
        ReportSht.Range("A" & initialPasteRow).Offset(17, 1).PasteSpecial (xlPasteAll)
    End If
    
    Call PostModify(BifacialSht, currentShtStatus)
    
    ' update current document length and set page break if needed
    GetLastShapeRow
    currentDocLength = Application.Max(ReportSht.Range("B" & Rows.count).End(xlUp).row, lastShapeRow)
    Call PageBreak(initialPasteRow, currentDocLength)
End Sub
'
' This function is called to copy and paste the desired
' contents from the Horizon sheet into the report sheet
'
Private Sub SetPrintFormatHorizon()
    Dim initialPasteRow As Integer
    Dim currentShtStatus As sheetStatus
    Dim chartWidth As Integer
    Dim cell As Range
    
    initialPasteRow = currentDocLength + 2
    
    Call PreModify(Horizon_ShadingSht, currentShtStatus)
    
    Call PasteHeader(Horizon_ShadingSht, initialPasteRow)
    Call CopyPaste(Horizon_ShadingSht.Range("PrintHorizon_Labels").SpecialCells(xlCellTypeVisible), ReportSht.Range("B" & initialPasteRow).Offset(2, 0))
    Call CopyPaste(Horizon_ShadingSht.Range("PrintHorizon_Values").SpecialCells(xlCellTypeVisible), ReportSht.Range("B" & initialPasteRow).Offset(2, 3), xlPasteValues)
    ReportSht.Range(ReportSht.Range("B" & initialPasteRow).Offset(2, 3), ReportSht.Range("B" & initialPasteRow).Offset(6, 3)).HorizontalAlignment = xlCenter
    
    For Each cell In ReportSht.Range(ReportSht.Range("B" & initialPasteRow).Offset(2, 3), ReportSht.Range("B" & initialPasteRow).Offset(6, 3)).Cells
        If IsEmpty(cell) = False Then
            cell.BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
        End If
    Next cell
        
    If Horizon_ShadingSht.Range("DefineHorizonProfile").Value = "Yes" Then
        Call CopyPaste(Horizon_ShadingSht.Range("PrintHorizon_Info").SpecialCells(xlCellTypeVisible), ReportSht.Range("B" & initialPasteRow).Offset(6, 2))
        Call CopyPasteChart(Horizon_ShadingSht, "HorizonChart", Sheets("Horizon").ChartObjects("HorizonChart"), ReportSht.Range("B" & initialPasteRow).Offset(2, 4))
        ' update number of shapes
        GetLastShapeRow
        ' resize the chart to fit on the page
        chartWidth = ReportSht.Shapes(shapeNum).Width
        ReportSht.Shapes(shapeNum).Width = chartWidth * 0.75

        'attempt to working around the moving transformer chart
        ReportSht.Shapes(shapeNum).Cut
        ReportSht.Range("G" & initialPasteRow).Offset(2, 0).PasteSpecial (xlPasteAll)
    End If
    
    Call PostModify(Horizon_ShadingSht, currentShtStatus)
    
    ' update current document length and set page break if needed
    currentDocLength = Application.Max(ReportSht.Range("B" & Rows.count).End(xlUp).row, lastShapeRow)
    Call PageBreak(initialPasteRow, currentDocLength)
End Sub
'
' This function is called to copy and paste the desired
' contents from the Systems sheet into the report sheet
'
' The systems sheet is also responsible for page breaking itself if
' the sheet itself exceeds the maximum allowed number of
' rows per pdf page
'
Private Sub SetPrintFormatSystem()
    Dim currentShtStatus As sheetStatus
    Dim lastRow As Integer         ' the last row with data
    Dim initialPasteRow As Integer ' the first row in which the data is pasted. Ranges are specified in offsets; changes only have to be made to the initialPasteRow to paste in a different place.
    Dim initialSubRow As Integer   ' the first row where sub arrays start to appear
    
    lastRow = SystemSht.Range("B" & Rows.count).End(xlUp).row
    initialPasteRow = currentDocLength + 2
    initialSubRow = initialPasteRow + 2 ' will increase now that adding more system summary values
    
    Call PreModify(SystemSht, currentShtStatus)
    
    ' Print out system summary
    Call PasteHeader(SystemSht, initialPasteRow)
    Call CopyPaste(SystemSht.Range("PrintSystem_NumSub"), ReportSht.Range("B" & initialPasteRow).Offset(2, 0))
    Call CopyPaste(SystemSht.Range("PrintSystem_NumSub_Vals"), ReportSht.Range("B" & initialPasteRow).Offset(2, 2))
    Call CopyPaste(SystemSht.Range("PrintSystem_PnomDC"), ReportSht.Range("B" & initialPasteRow).Offset(2, 4))
    Call CopyPaste(SystemSht.Range("SystemDC"), ReportSht.Range("B" & initialPasteRow).Offset(2, 6))
    Call CopyPaste(SystemSht.Range("PrintSystem_NumMods"), ReportSht.Range("B" & initialPasteRow).Offset(2, 8))
    Call CopyPaste(SystemSht.Range("PrintSystem_NumMods_Val"), ReportSht.Range("B" & initialPasteRow).Offset(2, 10))
    Call CopyPaste(SystemSht.Range("PrintSystem_NumInverters"), ReportSht.Range("B" & initialPasteRow).Offset(3, 0))
    Call CopyPaste(SystemSht.Range("PrintSystem_NumInverter_Vals"), ReportSht.Range("B" & initialPasteRow).Offset(3, 2))
    Call CopyPaste(SystemSht.Range("PrintSystem_PnomAC"), ReportSht.Range("B" & initialPasteRow).Offset(3, 4))
    Call CopyPaste(SystemSht.Range("SystemAC"), ReportSht.Range("B" & initialPasteRow).Offset(3, 6))
    Call CopyPaste(SystemSht.Range("PrintSystem_ACLossFraction"), ReportSht.Range("B" & initialPasteRow).Offset(3, 8))
    Call CopyPaste(SystemSht.Range("PrintSystem_ACLossFraction_Val"), ReportSht.Range("B" & initialPasteRow).Offset(3, 10), xlPasteValues) ' Drop down may cause issue
    ReportSht.Range("B" & initialPasteRow).Offset(3, 10).HorizontalAlignment = xlCenter
    ReportSht.Range("B" & initialPasteRow).Offset(3, 10).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
    ' no fill
    ReportSht.Range(ReportSht.Range("B" & initialPasteRow).Offset(1, 0), ReportSht.Range("B" & initialPasteRow).Offset(3, 10)).Interior.ColorIndex = xlNone
    
    ' Print out PV module model labels and values
    Call CopyPaste(SystemSht.Range(SystemSht.Range("PrintSystem_PVLabels"), SystemSht.Range("B" & lastRow)), ReportSht.Range("B" & initialPasteRow).Offset(5, 0))
    Call CopyPaste(SystemSht.Range(SystemSht.Range("PrintSystem_PV"), SystemSht.Range("C" & lastRow)), ReportSht.Range("D" & initialPasteRow).Offset(5, 0))
    
    ' Print out inverter model labels and values
    Call CopyPaste(SystemSht.Range(SystemSht.Range("PrintSystem_InvLabels"), SystemSht.Range("K" & lastRow)), ReportSht.Range("H" & initialPasteRow).Offset(5, 0))
    Call CopyPaste(SystemSht.Range(SystemSht.Range("PrintSystem_Inv"), SystemSht.Range("L" & lastRow)), ReportSht.Range("J" & initialPasteRow).Offset(5, 0))
    
    ' Print out row to fill in missing green filled cells
    Call CopyPaste(SystemSht.Range(SystemSht.Range("PrintSystem_SubHeader"), SystemSht.Range("I" & lastRow)), ReportSht.Range("C" & initialPasteRow).Offset(5, 0))
    Call CopyPaste(SystemSht.Range(SystemSht.Range("PrintSystem_SubHeader"), SystemSht.Range("I" & lastRow)), ReportSht.Range("I" & initialPasteRow).Offset(5, 0))
      
    Call PostModify(SystemSht, currentShtStatus)
    
    ' update current document length and set page break if needed
'    GetLastShapeRow
    currentDocLength = Application.Max(ReportSht.Range("B" & Rows.count).End(xlUp).row, lastShapeRow)
  
    ' if the size of the pasted sheet is larger than the pdfPageSize go through and set page breaks while updating the old doc length (rowsLeftOver)
    Dim rowsLeftOver As Integer
    If currentDocLength - initialPasteRow > pdfPageSize Then
        Dim count As Integer
        Dim i As Integer
        ' break at start of systems sheet then continue to break at necessary
        ReportSht.Range("A" & initialPasteRow).PageBreak = xlPageBreakManual
        ' reset rowsPerPage
        rowsPerPage = 0
        For i = initialSubRow To currentDocLength
            If InStr(ReportSht.Range("B" & i).Value, "SUB-ARRAY") <> 0 Then
                count = count + 1
                If count >= 4 Then
                    ReportSht.Range("A" & i).PageBreak = xlPageBreakManual
                    rowsLeftOver = i
                    'reset the count
                    count = 0
                End If
            End If
        Next i
        Call PageBreak(rowsLeftOver, currentDocLength)
    Else
        Call PageBreak(initialPasteRow, currentDocLength)
    End If
End Sub
'
' This function is called to copy and paste the desired
' contents from the Losses sheet into the report sheet
'
Private Sub SetPrintFormatLosses()
    Dim currentShtStatus As sheetStatus
    Dim cell As Range
    Dim initialPasteRow As Integer ' The first row in which the data is pasted. Ranges are specified in offsets; changes only have to be made to the initialPasteRow to paste in a different place.
    
    initialPasteRow = currentDocLength + 2 ' +2 to leave space between sections
    
    Call PreModify(LossesSht, currentShtStatus)
    
    Call PasteHeader(LossesSht, initialPasteRow)
    Call CopyPaste(LossesSht.Range("PrintLosses_Names").SpecialCells(xlCellTypeVisible), ReportSht.Range("B" & initialPasteRow).Offset(2, 0))
    Call CopyPaste(LossesSht.Range("PrintLosses_Vals").SpecialCells(xlCellTypeVisible), ReportSht.Range("F" & initialPasteRow).Offset(3, 0), xlPasteValues)
    ReportSht.Range(ReportSht.Range("F" & initialPasteRow).Offset(4, 0), ReportSht.Range("F" & initialPasteRow).Offset(17, 0)).HorizontalAlignment = xlCenter
    
    For Each cell In ReportSht.Range(ReportSht.Range("F" & initialPasteRow).Offset(4, 0), ReportSht.Range("F" & initialPasteRow).Offset(17, 0)).Cells
        If IsEmpty(cell) = False Then
            cell.BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
        End If
    Next cell
    
    If LossesSht.Range("IAMSelection").Value = "User Defined" Then
        Call CopyPaste(LossesSht.Range("PrintLosses_IAM_Table"), ReportSht.Range("I" & initialPasteRow).Offset(3, 0))
    Else
        Call CopyPaste(LossesSht.Range("PrintLosses_ASHRAE_Param"), ReportSht.Range("B" & initialPasteRow).Offset(18, 0))
        Call CopyPaste(LossesSht.Range("bNaught"), ReportSht.Range("F" & initialPasteRow).Offset(18, 0))
    End If
    
    Call PostModify(LossesSht, currentShtStatus)
    
    ' update current document length and set page break if needed
    GetLastShapeRow
    currentDocLength = Application.Max(ReportSht.Range("B" & Rows.count).End(xlUp).row, lastShapeRow)
    Call PageBreak(initialPasteRow, currentDocLength)
End Sub
'
' This function is called to copy and paste the desired
' contents from the Soiling sheet into the report sheet
'
Private Sub SetPrintFormatSoiling()
    Dim currentShtStatus As sheetStatus
    Dim initialPasteRow As Integer ' The first row in which the data is pasted. Ranges are specified in offsets; changes only have to be made to the initialPasteRow to paste in a different place.
    
    initialPasteRow = currentDocLength + 2 ' +1 to leave a space between sections
    
    Call PreModify(SoilingSht, currentShtStatus)
    
    Call PasteHeader(SoilingSht, initialPasteRow)
    Call CopyPaste(SoilingSht.Range("PrintSoiling1"), ReportSht.Range("B" & initialPasteRow).Offset(2, 0))
    
    If SoilingSht.Range("SfreqVal").Value = "Yearly" Then
        Call CopyPaste(SoilingSht.Range("SoilingYearly"), ReportSht.Range("D" & initialPasteRow).Offset(3, 0))
    Else
        Call CopyPaste(SoilingSht.Range("PrintSoilingMonthly"), ReportSht.Range("B" & initialPasteRow).Offset(4, 0))
    End If
    
    Call CopyPasteChart(SoilingSht, "SoilingChart", Sheets("Soiling").ChartObjects("SoilingChart"), ReportSht.Range("B" & initialPasteRow).Offset(7, 1))
    ' update number of charts on page
    GetLastShapeRow
    ReportSht.Shapes(shapeNum).Width = 725
    ReportSht.Shapes(shapeNum).Cut
    ReportSht.Range("A" & initialPasteRow).Offset(7, 1).PasteSpecial (xlPasteAll)
    
    Call PostModify(SoilingSht, currentShtStatus)
    
    ' update current document length and set page break if needed
    GetLastShapeRow
    currentDocLength = Application.Max(ReportSht.Range("B" & Rows.count).End(xlUp).row, lastShapeRow)
    Call PageBreak(initialPasteRow, currentDocLength)
End Sub
'
' This function is called to copy and paste the desired
' contents from the Spectral sheet into the report sheet
'
Private Sub SetPrintFormatSpectral()
    Dim currentShtStatus As sheetStatus
    Dim initialPasteRow As Integer ' The first row in which the data is pasted. Ranges are specified in offsets; changes only have to be made to the initialPasteRow to paste in a different place.
    Dim chartHeight As Integer
    
    initialPasteRow = currentDocLength + 2
    
    Call PreModify(SpectralSht, currentShtStatus)
    
    Call PasteHeader(SpectralSht, initialPasteRow)
    Call CopyPaste(SpectralSht.Range("PrintSpectral1"), ReportSht.Range("B" & initialPasteRow).Offset(2, 0))
    
    ' Print the rest of the sheet only if UseSpectralModel is set to Yes
    If SpectralSht.Range("UseSpectralModel").Value = "Yes" Then
        ' Copy and paste spectral info
        ReportSht.Range("B" & initialPasteRow).Offset(4, 0).Value = "kt"
        ReportSht.Range("B" & initialPasteRow).Offset(5, 0).Value = "Correction"
        Call CopyPaste(SpectralSht.Range("PrintSpectral2"), ReportSht.Range("C" & initialPasteRow).Offset(4, 0))
        
        ' Copy and paste spectral chart
        Call CopyPasteChart(SpectralSht, "ktSpectralChart", SpectralSht.ChartObjects("ktSpectralChart"), ReportSht.Range("D" & initialPasteRow).Offset(8, 0))
    End If
    
    Call PostModify(SpectralSht, currentShtStatus)
    
    ' update current document length and set page break if needed
    GetLastShapeRow
    currentDocLength = Application.Max(ReportSht.Range("B" & Rows.count).End(xlUp).row, lastShapeRow)
    Call PageBreak(initialPasteRow, currentDocLength)
End Sub

'
' This function is called to copy and paste the desired
' contents from the Transformer sheet into the report sheet
'
Private Sub SetPrintFormatTransformer()
    Dim currentShtStatus As sheetStatus
    Dim initialPasteRow As Integer ' The first row in which the data is pasted. Ranges are specified in offsets; changes only have to be made to the initialPasteRow to paste in a different place.
    Dim chartHeight As Integer
    
    initialPasteRow = currentDocLength + 2 ' 63 + 15
    
    Call PreModify(TransformerSht, currentShtStatus)
    
    Call PasteHeader(TransformerSht, initialPasteRow)
    Call CopyPaste(TransformerSht.Range("PrintTransformer1"), ReportSht.Range("B" & initialPasteRow).Offset(2, 0))
    Call CopyPaste(TransformerSht.Range("PrintTransformer2"), ReportSht.Range("G" & initialPasteRow).Offset(2, 0))
    
    'copy and paste transformer losses chart
    Call CopyPasteChart(TransformerSht, "TransformerChart", Sheets("Transformer").ChartObjects("TransformerChart"), ReportSht.Range("H" & initialPasteRow).Offset(2, 0))
    
    Call PostModify(TransformerSht, currentShtStatus)
    
    ' update number of charts/shapes and chart/shape last row. This is important for the resizing and cutting and pasting
    GetLastShapeRow

    ' resize the chart to fit on the page
    chartHeight = ReportSht.Shapes(shapeNum).Height
    ReportSht.Shapes(shapeNum).Height = chartHeight * 0.75

    'attempt to working around the moving transformer chart
    ReportSht.Shapes(shapeNum).Cut
    ReportSht.Range("H" & initialPasteRow).Offset(2, 1).PasteSpecial (xlPasteAll)
    
    ' update current document length and set page break if needed
    GetLastShapeRow
    currentDocLength = Application.Max(ReportSht.Range("B" & Rows.count).End(xlUp).row, lastShapeRow)
    Call PageBreak(initialPasteRow, currentDocLength)
End Sub
'
' This function is called to copy and paste the desired
' contents from the ASTM Regression sheet into the report sheet
' if selected
'
Private Sub SetPrintASTM()
    Dim currentShtStatus As sheetStatus
    Dim initialPasteRow As Integer
     
    initialPasteRow = currentDocLength + 2
    
    Call PreModify(AstmSht, currentShtStatus)
    
    Call PasteHeader(AstmSht, initialPasteRow)
    Call CopyPaste(AstmSht.Range("PrintASTM_Definition"), ReportSht.Range("B" & initialPasteRow).Offset(2, 0))
    Call CopyPaste(AstmSht.Range("PrintASTM_EAF"), ReportSht.Range("B" & initialPasteRow).Offset(13, 0))
    
    Call PostModify(AstmSht, currentShtStatus)
    
    ' update current document length and set page break if needed
    GetLastShapeRow
    currentDocLength = Application.Max(ReportSht.Range("B" & Rows.count).End(xlUp).row, lastShapeRow)
End Sub
'
' This function is called to copy and paste the desired
' contents from the Climate File sheet into the report sheet
'
Private Sub SetPrintFormatInput()
    Dim currentShtStatus As sheetStatus
    Dim initialPasteRow As Integer ' The first row in which the data is pasted. Ranges are specified in offsets; changes only have to be made to the initialPasteRow to paste in a different place.
    
    initialPasteRow = currentDocLength + 2
    
    Call PreModify(InputFileSht, currentShtStatus)
    
    Call PasteHeader(InputFileSht, initialPasteRow)
    If Right(InputFileSht.Range("InputFilePath").Value, 4) = ".csv" Then
        Call CopyPaste(InputFileSht.Range("PrintInput1"), ReportSht.Range("B" & initialPasteRow).Offset(2, 0))
        Call CopyPaste(InputFileSht.Range("PrintInput2"), ReportSht.Range("G" & initialPasteRow).Offset(6, 0))
        Call CopyPaste(InputFileSht.Range("PrintInput3"), ReportSht.Range("J" & initialPasteRow).Offset(6, 0))
        Call CopyPaste(InputFileSht.Range("PrintInput4"), ReportSht.Range("B" & initialPasteRow).Offset(16, 0))
        ReportSht.Range("D" & initialPasteRow).Offset(7, 0).Value = "Column"
        
        Call PostModify(InputFileSht, currentShtStatus)
    
        ' If there is the comment telling the user to label Panel Temperature, then delete it
        If Not ReportSht.Range("D" & initialPasteRow).Offset(12, 0).Comment Is Nothing Then ReportSht.Range("D" & initialPasteRow).Offset(12, 0).Comment.Delete
    
        ' Merge the First and Last date cells with their adjacent cells so that they can be seen entirely
        ReportSht.Range(Range("J" & initialPasteRow).Offset(10, 0), Range("L" & initialPasteRow).Offset(10, 0)).Merge
        ReportSht.Range(Range("J" & initialPasteRow).Offset(11, 0), Range("L" & initialPasteRow).Offset(11, 0)).Merge
    
        ' Add the correct colour to the headers and titles
'       ReportSht.Range(Range("B" & initialPasteRow), Range("M" & initialPasteRow)).Interior.Color = ColourThemeGreen
        ReportSht.Range(Range("G" & initialPasteRow).Offset(6, 0), Range("L" & initialPasteRow).Offset(6, 0)).Interior.Color = ColourThemeGreen
    
        ' Make the input file path has no fill just in case it was red previously due to a file path not being found
        ReportSht.Range("B" & initialPasteRow).Offset(15, 0).Interior.ColorIndex = xlNone
    
        ' Add borders to the time series characteristic box and the First and Last date boxes
        Call ReportSht.Range(Range("G" & initialPasteRow).Offset(6, 0), Range("L" & initialPasteRow).Offset(6, 0)).BorderAround(xlContinuous, xlThin, xlColorIndexAutomatic)
        Call ReportSht.Range(Range("G" & initialPasteRow).Offset(6, 0), Range("L" & initialPasteRow).Offset(11, 0)).BorderAround(xlContinuous, xlThin, xlColorIndexAutomatic)
        Call ReportSht.Range(Range("J" & initialPasteRow).Offset(10, 0), Range("L" & initialPasteRow).Offset(11, 0)).BorderAround(xlContinuous, xlThin, xlColorIndexAutomatic)
        Call ReportSht.Range(Range("J" & initialPasteRow).Offset(11, 0), Range("L" & initialPasteRow).Offset(11, 0)).BorderAround(xlContinuous, xlThin, xlColorIndexAutomatic)
    
        ' Left align the time format (yyy-MM-dd HH:mm:ss, etc.) so that it can be seen even when it is very long
        ReportSht.Range("E" & initialPasteRow).Offset(7, 0).HorizontalAlignment = xlLeft
        ReportSht.Range(Range("B" & initialPasteRow).Offset(4, 0), Range("B" & initialPasteRow).Offset(5, 0)).EntireRow.Delete
    Else
        Call CopyPaste(InputFileSht.Range("PrintInput4"), ReportSht.Range("B" & initialPasteRow).Offset(2, 0))
    End If

    ' update current document length and set page break if needed
    GetLastShapeRow
    currentDocLength = Application.Max(ReportSht.Range("B" & Rows.count).End(xlUp).row, lastShapeRow)
    Call PageBreak(initialPasteRow, currentDocLength)
End Sub
'
' This function is called to copy and paste the desired
' contents from the Summary sheet into the report sheet
'
Private Sub SetPrintFormatSummary()
    Dim currentShtStatus As sheetStatus
    Dim initialPasteRow As Integer
    Dim initialRow As Integer
    Dim tableRows As Integer
    Dim tableColumns As Integer
    Dim tempTableEnd As Integer
    Dim cell As Range
    Dim i As Integer
    
    initialPasteRow = currentDocLength + 2              ' change to +2 at end because there will always be data to read the proper length
    initialRow = initialPasteRow - rowsPerPage - 2      ' the initial row in the pdf page used for pagebreaking correctly DEAD CODE
    
    tableRows = SummarySht.Range("A" & Rows.count).End(xlUp).row - 10 ' -10 from (-10) due to SummarySht header
    tableColumns = SummarySht.Cells(12, Columns.count).End(xlToLeft).Column
    
    Call PreModify(SummarySht, currentShtStatus)
    
    Call PasteHeader(SummarySht, initialPasteRow)
    
    ' if the data summary will take up more than a page, start it on its own page
    If tableColumns >= 27 Then
        ReportSht.Range("A" & initialPasteRow).PageBreak = xlPageBreakManual
        rowsPerPage = 0
    End If
    
    For i = 2 To tableColumns Step 10
        tempTableEnd = i + 9
        If tempTableEnd > tableColumns Then
            tempTableEnd = tableColumns
        End If
        
        Call CopyPaste(SummarySht.Range("A12", "A" & (tableRows + 10)), ReportSht.Range("B" & initialPasteRow).Offset(2, 0))
        Call CopyPaste(SummarySht.Range(SummarySht.Cells(12, i), SummarySht.Cells(tableRows + 10, tempTableEnd)), ReportSht.Range("B" & initialPasteRow).Offset(2, 1))
        
        ReportSht.Range(ReportSht.Range("B" & initialPasteRow).Offset(2, 1), ReportSht.Range("B" & initialPasteRow).Offset(tableRows, tempTableEnd)).HorizontalAlignment = xlCenter
        
        For Each cell In ReportSht.Range(ReportSht.Range("B" & initialPasteRow).Offset(2, 1), ReportSht.Range("B" & initialPasteRow).Offset(tableRows, tempTableEnd)).Cells
            If cell.WrapText Then ReportSht.Range("B" & initialPasteRow).Offset(2, 1).Rows.AutoFit
        Next cell
        
        ' update the current doc length after every chart paste
        currentDocLength = ReportSht.Range("B" & Rows.count).End(xlUp).row
        
        ' check to make sure that the pasted table will fit onto a page and if not break
        Call PageBreak(initialPasteRow, currentDocLength)
        
        ' update initialPasteRow for next section of the table
        initialPasteRow = initialPasteRow + tableRows
        
'        ' update the initialRow when there is a page break
'        If currentDocLength - initialRow >= pdfPageSize Then
'            initialRow = initialPasteRow
'        End If
    Next i
    
    Call PostModify(SummarySht, currentShtStatus)
    
End Sub
'
' This function is called to copy and paste the desired
' contents from the Losses Diagram sheet into the report sheet
'
Private Sub SetPrintLossesDiagram()
    Dim currentShtStatus As sheetStatus
    Dim initialPasteRow As Integer

    initialPasteRow = currentDocLength + 2
    
    Call PreModify(LossDiagramSht, currentShtStatus)
    
    ' copy and paste losses diagram
    Call PasteHeader(LossDiagramSht, initialPasteRow)
    Call CopyPasteChart(LossDiagramSht, "LossDiagram", Sheets("Losses Diagram").ChartObjects("LossDiagram"), ReportSht.Range("C" & initialPasteRow).Offset(2, 0))
    
    Call PostModify(LossDiagramSht, currentShtStatus)
    
    ' update current document length and set page break if needed
    GetLastShapeRow
    currentDocLength = Application.Max(ReportSht.Range("B" & Rows.count).End(xlUp).row, lastShapeRow)
    Call PageBreak(initialPasteRow, currentDocLength)
End Sub
'
' CopyPaste subroutine
' Used to facilitate copying and pasting in one line instead of two
'
Private Sub CopyPaste(ByRef copyRange As Range, ByRef pasteRange As Range, Optional ByVal pasteType As Long = xlPasteAll)

    copyRange.Copy
    pasteRange.PasteSpecial (pasteType)
    Application.CutCopyMode = False
    
End Sub
'
' CopyPasteChart subroutine
' Used to facilitate copying and pasting of charts in one line instead of two
'
Private Sub CopyPasteChart(ByRef currentSht As Worksheet, ByVal chartName As String, ByRef copyChart As ChartObject, ByRef pasteRange As Range, Optional pasteType As Long = xlPasteAll)
    
    currentSht.ChartObjects(chartName).Activate
    
    copyChart.Copy
    pasteRange.PasteSpecial (pasteType)
    Application.CutCopyMode = False
    
    
    
End Sub
'
' This sub is called at the end of every function and
' places a page break when needed
'
Private Sub PageBreak(ByVal oldDocLength As Integer, ByVal newDocLength As Integer)
    oldDocLength = oldDocLength - 1         ' -1 is compensating for the empty row between all sections caused by initialPasteRow + 2
    rowsAdded = newDocLength - oldDocLength
    
    If rowsPerPage + rowsAdded <= pdfPageSize Then
        rowsPerPage = rowsPerPage + rowsAdded
    Else
        ReportSht.Range("A" & oldDocLength + 1).PageBreak = xlPageBreakManual
        rowsPerPage = rowsAdded + 1
        ' check if the rows in page are still larger than pdfPageSize
        Do While rowsPerPage >= pdfPageSize
            ReportSht.Range("A" & (oldDocLength + pdfPageSize - 4)).PageBreak = xlPageBreakManual
            ' update rowsPerPage and oldDocLength
            rowsPerPage = rowsPerPage - pdfPageSize
            oldDocLength = oldDocLength + pdfPageSize - 4
        Loop
    End If
End Sub
'
' This function is used to find the laste shape in the
' report sheet and update the largest row number taken up by the last shape
'
' This is important to determine what is actually the last row in the sheet
'
Function GetLastShapeRow() As Boolean
    shapeNum = ReportSht.Shapes.count
    If shapeNum = 0 Then
        lastShapeRow = 0
        Exit Function
    End If
    lastShapeRow = ReportSht.Shapes(shapeNum).BottomRightCell.row
End Function

Private Sub PasteHeader(ByRef currentSheet As Worksheet, ByVal initialPasteRow As Integer)
    ' copy and paste title of sheet
    Call CopyPaste(currentSheet.Range("B8"), ReportSht.Range("B" & initialPasteRow))
    ' fill header row
    ReportSht.Range(Range("B" & initialPasteRow), Range("N" & initialPasteRow)).Interior.Color = ColourThemeGreen
End Sub


