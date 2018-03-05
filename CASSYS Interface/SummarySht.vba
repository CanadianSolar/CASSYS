VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SummarySht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                   SUMMARY SHEET                   '
'---------------------------------------------------'
' The summary sheet is responsible for creating a   '
' pivot table generated from the simulation result  '
' data.                                             '

Public showSummary As Boolean ' If atleast one output parameter was chosen the summary sheet will be shown
Private Sub Worksheet_Activate()
    
    Me.Activate
    Range("ViewDays").Select

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim currentSheetStatus As sheetStatus
    
    ' Changes the Row Labels to either dates (showing individual days) or month-year based on the selection of daily or monthly
    If Target.Address = SummarySht.Range("ViewDays").Address Then
        If SummarySht.Range("A12").Value <> vbNullString Then
            Call PreModify(SummarySht, currentSheetStatus)
            If SummarySht.Range("ViewDays").Value = "Daily" Then
                SummarySht.PivotTables("ReportTable").PivotFields("Date").Orientation = xlRowField
                SummarySht.PivotTables("ReportTable").PivotFields("Month").Orientation = xlHidden
            ElseIf SummarySht.Range("ViewDays").Value = "Monthly" Then
                SummarySht.PivotTables("ReportTable").PivotFields("Month").Orientation = xlRowField
                SummarySht.PivotTables("ReportTable").PivotFields("Date").Orientation = xlHidden
            End If
            Call PostModify(SummarySht, currentSheetStatus)
        End If
    End If
End Sub

' CreateReportTable subroutine
'
' This function creates the pivot table for the Data Summary Sheet
' Each data column is summarized in an appropriate way based on
' the units of the data in that column. For example, unitless
' outputs are not summarized but wind velocity (m/s) is averaged
'
' It takes in two integers specifying the last column and last row of data
' to determine where to import the pivot table
Sub CreateReportTable(resultLastRow As Long, resultLastColumn As Long)

    Dim reportCache As PivotCache
    Dim ReportTable As PivotTable
    Dim resultRange As Range
    Dim outputParam As Range
    Dim summarizeParam As Boolean
    
    Dim showSummary As Boolean ' If atleast one output parameter was chosen the summary sheet will be shown
    
    Application.EnableEvents = False
    
    ' Assigning the result sheet range, and pivot table destination
    Set resultRange = ResultSht.Range("A1", ResultSht.Cells(resultLastRow, resultLastColumn))
    
    ' clearing destination
    SummarySht.Range(SummarySht.Cells(12, 1), SummarySht.Cells(SummarySht.Rows.count, SummarySht.Columns.count)).Delete
    
    ' Creating Pivot cache
    Set reportCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=resultRange.Address(, , , True))
    
    ' Creating pivot table
    Set ReportTable = reportCache.CreatePivotTable(SummarySht.Range("A12"), "ReportTable")
    
    ' showSummary is initially set to false; it will be set to true if atleast one output parameter was selected to be summarized
    showSummary = False
    
    ' ManualUpdate speeds up the addition of data fields to the pivot table
    ReportTable.ManualUpdate = True
    With ReportTable
        For Each reportfield In .PivotFields
        
            ' Add monthly row label
            If InStr(1, reportfield.Name, "Month") Then
                Call .AddFields(reportfield.Name)
                With ReportTable.PivotFields(reportfield.Name)
                    .Orientation = xlRowField
                    .Position = 1
                End With
            End If
                
            ' Loop through parameter names to see if the user selected 'Summarize'
            For Each outputParam In OutputFileSht.Range("OutputParam")
                If InStr(1, reportfield.Name, outputParam.Offset(0, OutputFileSht.Range("HeaderRow").Column - OutputFileSht.Range("OutputParam").Column).Value) Then
                    If outputParam.Value2 = "Summarize" Then
                        summarizeParam = True
                        showSummary = True
                        Exit For
                    End If
                End If
            Next
        
            ' If the parameter was selected to be summarized then add it to the pivot table
            If summarizeParam = True Then
                On Error GoTo nextParam          ' This prevents the pivot table creation from failing when the simulation returns invalid results such as NaN
                ' Decide how to summarize the parameter based on its units
                If InStr(1, reportfield.Name, "W/m2") <> 0 Then
                    ' If the units are W/m2, then convert it to kWh/m2 using the value of the nominal timestep
                    .CalculatedFields.Add Replace(reportfield.Name, "W/m2", "kWh/m2"), "=" & CStr(InputFileSht.Range("Interval").Value / 60000) & "*'" & reportfield.Name & "'", True
                    .PivotFields(Replace(reportfield.Name, "W/m2", "kWh/m2")).Orientation = xlDataField
                ElseIf InStr(1, reportfield.Name, "mmm-yyyy") Then
                    Call .AddFields(reportfield.Name)
                    With ReportTable.PivotFields(reportfield.Name)
                        .Orientation = xlRowField
                        .Position = 1
                    End With
                ElseIf InStr(1, reportfield.Name, "(kW)") <> 0 Then
                    .CalculatedFields.Add Replace(reportfield.Name, "kW", "kWh"), "=" & CStr(InputFileSht.Range("Interval").Value / 60) & "*'" & reportfield.Name & "'", True
                    .PivotFields(Replace(reportfield.Name, "kW", "kWh")).Orientation = xlDataField
                ElseIf InStr(1, reportfield.Name, "deg. C") <> 0 Then
                    ' Add temperature fields
                    Call .AddDataField(field:=reportfield, Function:=xlAverage)
                    .PivotFields(reportfield.Name).NumberFormat = "0.0"
                ElseIf InStr(1, reportfield.Name, "m/s") <> 0 Then
                    ' Add wind velocity
                    Call .AddDataField(field:=reportfield, Function:=xlAverage)
                    .PivotFields(reportfield.Name).NumberFormat = "0.0"
                ElseIf InStr(1, reportfield.Name, "(kWh)") <> 0 Then
                    ' Add Energy fields
                    .PivotFields(reportfield.Name).NumberFormat = "0.0"
                    Call .AddDataField(field:=reportfield, Function:=xlSum)
                End If
nextParam:
                On Error GoTo 0
            End If
            summarizeParam = False
        Next
    End With
    
    With ReportTable
        ' Refresh pivot table to show the newly added fields since manual update was turned off
        .PivotCache.Refresh
        ' Change 'Row Labels' to 'Time Period' (the row header that sorts the data)
        .CompactLayoutRowHeader = "Time Period"
        ' Show grand total at the bottom of the pivot table
        .ColumnGrand = True
        .ManualUpdate = False
        .TableStyle2 = "PivotStyleDark2"
        .ShowValuesRow = False
        
        ' Format each column in the pivot table
        For Each reportfield In .VisibleFields
            If reportfield.Orientation <> xlRowField And reportfield.Name <> "Values" Then
                If InStr(reportfield, "kWh") <> 0 Then
                    reportfield.NumberFormat = "#,##0"
                Else
                    reportfield.NumberFormat = "#,##0.0"
                End If
            End If
        Next reportfield
        
        
        ThisWorkbook.ShowPivotTableFieldList = False
    End With

    SummarySht.Visible = xlSheetHidden
    
    ' Format the report table
    If showSummary = True Then
        Call FormatAfterSimulation(SummarySht, 12, 1, 12)
        SummarySht.Visible = xlSheetVisible
        ActiveWindow.DisplayGridlines = False
        SummarySht.Activate
    Else
        ResultSht.Activate
        ActiveWindow.ScrollRow = 2
        ActiveWindow.ScrollColumn = 2
    End If
    
    Application.EnableEvents = True
     
End Sub
