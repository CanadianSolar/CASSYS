VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LossesSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                       LOSSES SHEET                        '

' This worksheet is where the user specifies the loss       '
' characterization parameters for thermal losses, module    '
' losses and the IAM (ASHRAE Model parameter)               '


Option Explicit
Private Sub Worksheet_Activate()
 
    ' Resets the active sheet to this sheet
    Me.Activate
    ' Upon sheet activation the first editable cell is selected
    Range("UseMeasuredValues").Select
    
    Dim currentShtStatus As sheetStatus
    Dim i As Integer
    Dim numSubArrays As Integer
    Dim numRows As Integer
    Dim definitionAvailable As Boolean
    numSubArrays = SystemSht.Range("NumSubArray").Value
    numRows = 17
    
    For i = 0 To (numSubArrays - 1)
        If SystemSht.Range("DefnAvailable").Offset(i * numRows, 0).Value = "Yes" Then
            definitionAvailable = True
            Exit For
        End If
        If i = (numSubArrays - 1) Then
            definitionAvailable = False
        End If
    Next i
    
    Call PreModify(LossesSht, currentShtStatus)
    If Not definitionAvailable Then
        Dim gradient As LinearGradient
        Dim colorStop As colorStop
        Range("UsePan").Value = "No"
        Range("UsePan").Interior.Pattern = XlPattern.xlPatternLinearGradient
        Set gradient = Range("UsePan").Interior.gradient
        gradient.degree = 90
        Set colorStop = gradient.ColorStops.Add(0.01)
        colorStop.Color = RGB(204, 192, 218)
        colorStop.TintAndShade = 0.1
        Set colorStop = gradient.ColorStops.Add(0.99)
        colorStop.Color = RGB(204, 192, 218)
        colorStop.TintAndShade = -0.15

        
        Range("UsePan").Locked = True
        
    ElseIf definitionAvailable Then
        Range("UsePan").Locked = False
        Range("UsePan").Interior.Color = xlNone
    End If
    Call PostModify(LossesSht, currentShtStatus)
End Sub
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    Dim currentShtStatus As sheetStatus
    ' If Save is selected run Save Function
    If Target.Range.Address = Range("SaveLosses").Address Then
        Call PreModify(LossesSht, currentShtStatus)
        Call SaveXML
        Call PostModify(LossesSht, currentShtStatus)
    End If
    
End Sub

' WorkSheet_Change Function
' This function is called when a cell is changed but
' will only do anything when the Range("NumSubArray")
' is changed
'
' The purpose of this function is to show and hide the heat
' loss factor fields whenever the user wishes to use measured
' values or not by changing its field's value
Private Sub Worksheet_Change(ByVal Target As Range)
      
    Dim currentShtStatus As sheetStatus
   
    If Not Intersect(Target, Range("UseMeasuredValues")) Is Nothing Then
        Call PreModify(LossesSht, currentShtStatus)
        ' If use measured values is is true, hide user defined heat loss factor
        If (Range("UseMeasuredValues").Value = True) Then
            LossesSht.Range("HeatLossRows").EntireRow.Hidden = True
            LossesSht.Range("ReplaceHeatLossRows").EntireRow.Hidden = False
        ElseIf (Range("UseMeasuredValues").Value = False) Then
            ' If use measured values is is true, show user defined heat loss factor
            LossesSht.Range("HeatLossRows").EntireRow.Hidden = False
            LossesSht.Range("ReplaceHeatLossRows").EntireRow.Hidden = True
        End If
        Call PostModify(LossesSht, currentShtStatus)
    End If
   
    If Not Intersect(Target, Range("IAMSelection")) Is Nothing Then
        Call PreModify(LossesSht, currentShtStatus)
        If (Range("IAMSelection").Value = "ASHRAE") Then
            LossesSht.Range("ASHRAERow").EntireRow.Hidden = False
            LossesSht.Range("UserDefinedIAMRows").EntireRow.Hidden = True
            LossesSht.ChartObjects("IAMChart").Visible = False
        ElseIf (Range("IAMSelection").Value = "User Defined") Then
            LossesSht.Range("UserDefinedIAMRows").EntireRow.Hidden = False
            LossesSht.Range("ASHRAERow").EntireRow.Hidden = True
            LossesSht.ChartObjects("IAMChart").Visible = True
        End If
        Call PostModify(LossesSht, currentShtStatus)
    End If
End Sub
