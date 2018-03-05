VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChartConfigSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                                   CHART BUILDER                                         '
'-----------------------------------------------------------------------------------------'
'   The Chart Builder worksheet is activated after simulation,                            '
'   and allows the user to select outputs (generated from the simulation)                 '
'   to compare. The user is able to graph one X value against a maximum of 4              '
'   Y values. Once the user selects 'Display' after making their selection,               '
'   The chart is generated on a separate worksheet, named Chart 1, Chart 2, or Chart 3.   '

Option Explicit

' DispChart subroutine

' This subroutine is responsible for generating the chart
' after the user selects their desired comparison parameters.
' It is called by one of the three 'Display' buttons on the Chart
' builder sheet, and generates the chart on either Chart 1, 2, or 3.

Sub DispChart(ByVal chartNum As Integer)

    Dim Chart As Chart
    Dim chartName As String
    Dim i As Integer ' Iteration variable specifying the current Y-Value
    Dim XRange As Range ' contains the x axis labels
    Dim YRange(4) As Range  ' contains the y-axis labels

    
    ' Set reference to the calling chart
    Set Chart = Charts("Chart" & chartNum)
    chartName = "Chart" & chartNum
    
     ' Clear previous axis titles
    If Chart.Axes(xlValue).HasTitle = True Then Chart.Axes(xlValue).AxisTitle.Caption = vbNullString
    
    ' Clear previous series data
    Do Until Chart.SeriesCollection.count = 0
        Chart.SeriesCollection(1).Delete
    Loop

    ' Get the column corresponding to the selected X-Axis parameter
    Set XRange = ResultSht.Range("GraphParam").Find(Range(chartName & "X"), LookIn:=xlValues, SearchFormat:=False, MatchCase:=False)
    
    ' If the selected parameter cannot be found it is possible the parameter name relies on a formula
    If XRange Is Nothing Then Set XRange = ResultSht.Range("GraphParam").Find(Range(chartName & "X"), LookIn:=xlFormulas, SearchFormat:=False, MatchCase:=False)
    
    
    ' Get the columns corresponding to the selected Y-Axis parameters
    For i = 1 To 4 Step 1
        Set YRange(i) = ResultSht.Range("GraphParam").Find(Range(chartName & "Y" & i), LookIn:=xlValues, SearchFormat:=False, MatchCase:=False)
        
        ' If the selected parameter cannot be found it is possible that the parameter name relies on a formula
        If YRange(i) Is Nothing Then Set YRange(i) = ResultSht.Range("GraphParam").Find(Range(chartName & "Y" & i), LookIn:=xlFormulas, SearchFormat:=False, MatchCase:=False)
    Next i
    
    ' Loop through all Y Values
    For i = 1 To Range(chartName & "NumVal")
        
        With Chart.SeriesCollection.NewSeries
            
            ' Set orientation and graph type
            If InStr(1, XRange.Value, "Timestamp") <> 0 Or InStr(1, XRange.Value, "Date") <> 0 Or InStr(1, XRange.Value, "Month") <> 0 Then
                .ChartType = xlXYScatterSmoothNoMarkers
                Chart.Axes(xlCategory).TickLabels.Orientation = 90
                If InStr(1, XRange.Value, "Month") = 0 Then
                    Chart.Axes(xlCategory).TickLabels.NumberFormat = "yyyy-mm-dd"
                Else
                    Chart.Axes(xlCategory).TickLabels.NumberFormat = "mmm-yyyy"
                End If
            Else
                .MarkerStyle = -4168
                .MarkerSize = 3
                .ChartType = xlXYScatter
                Chart.Axes(xlCategory).TickLabels.Orientation = 0
                Chart.Axes(xlCategory).TickLabels.NumberFormat = "General"
            End If
        
            
            ' Add X-axis title
            Chart.Axes(xlCategory).HasTitle = True
            Chart.Axes(xlCategory).AxisTitle.Caption = XRange.Value
            
            ' Add Y-axis title
            Chart.Axes(xlValue).HasTitle = True
            
            If Chart.Axes(xlValue).AxisTitle.Caption = "Axis Title" Then
                Chart.Axes(xlValue).AxisTitle.Caption = YRange(i).Value
            Else
                Chart.Axes(xlValue).AxisTitle.Caption = YRange(i).Value & " & " & Charts("Chart " & chartNum).Axes(xlValue).AxisTitle.Caption
            End If
            
            ' Create title string
            If i = 1 Then
                Chart.ChartTitle.text = YRange(i).Value & " vs " & XRange.Value
            Else
                Chart.ChartTitle.text = YRange(i).Value & " & " & Chart.ChartTitle.text
            End If
            
            ' Add Title
            Chart.HasTitle = True
            .Name = YRange(i).Value & " vs " & XRange.Value
            
            If InStr(1, YRange(i).Value, "Date") <> 0 Or InStr(1, YRange(i).Value, "Month") <> 0 Or InStr(1, YRange(i).Value, "Timestamp") <> 0 Then
                If InStr(1, YRange(i).Value, "Date") <> 0 Or InStr(1, YRange(i).Value, "Timestamp") <> 0 Then Chart.Axes(xlValue).TickLabels.NumberFormat = "yyyy-mm-dd"
                If YRange(i).Value = "Month" Then Chart.Axes(xlValue).TickLabels.NumberFormat = "mmm-yyyy"
            Else
                Chart.Axes(xlValue).TickLabels.NumberFormat = "General"
            End If
            
            ' Sets the range where the graph fetches its data
            ' Value2 is important here as this ensures that formulas are read properly as well
            If YRange(i).Offset(i, 0).End(xlDown).row > 65535 Then
                .Values = "Details!" & YRange(i).Offset(1, 0).Address & ":" & YRange(i).End(xlDown).Address
                .XValues = "Details!" & XRange(i).Offset(1, 0).Address & ":" & XRange.End(xlDown).Address
            Else
                .Values = ResultSht.Range(YRange(i).Offset(1, 0).Address, YRange(i).End(xlDown).Address).Value2
                .XValues = ResultSht.Range(XRange.Offset(1, 0).Address, XRange.End(xlDown).Address).Value2
            End If

            
        End With
    Next i
    
    ' Delete the legend if only one series is present
    If Chart.SeriesCollection.count = 1 Then
        Chart.HasLegend = False
    Else
        Chart.HasLegend = True
    End If
    
    ' Make the chart visible and activate the correct chart sheet
    Chart.Visible = True
    Chart.Activate
    Exit Sub
      
' The user tried to graph with a selection of blank x and y values
ErrHandler:
    MsgBox "Please make sure the fields for all X and Y values are filled." & vbNewLine & "Make selections using the drop-down lists labelled 'X Values' and 'Y Values'."
      
End Sub

' DispChart1 Function
'
' The purpose of this function is to call DispChart
' and generate a chart on the sheet named Chart 1
' when the first 'Display' button is pressed

Sub DispChart1()
    Call DispChart(1)
End Sub
  
' DispChart2 Function
'
Sub DispChart2()
    Call DispChart(2)
End Sub

' DispChart3 Function
'
Sub DispChart3()
    Call DispChart(3)
End Sub
           
' WorkSheet_Change Function
' This function is called when a cell is changed
'
' The purpose of this function is to update the number
' of y values in a given chart whenever that number is
' changed

Private Sub Worksheet_Change(ByVal Target As Range)
    
    Dim i As Integer
    Dim j As Integer
    Dim currentShtStatus As sheetStatus
    
    Call PreModify(ChartConfigSht, currentShtStatus)
    
    ' Update number of y vales if the number of y values is changed
    For i = 1 To 3
        If Not Intersect(Target, Range("Chart" & i & "NumVal")) Is Nothing Then
            For j = 1 To 4
                If (j > Range("Chart" & i & "NumVal").Value) Then
                    ChartConfigSht.Rows(3 + j + (10 * i)).Hidden = True
                Else
                    ChartConfigSht.Rows(3 + j + (10 * i)).Hidden = False
                End If
            Next j
        End If
    Next i
   
    Call PostModify(ChartConfigSht, currentShtStatus)
    
End Sub
                    
                    
                    


