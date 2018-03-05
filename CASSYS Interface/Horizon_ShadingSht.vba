VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Horizon_ShadingSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                          HORIZON SHADING SHEET                            '
'---------------------------------------------------------------------------'
' The Horizon Shading Sheet is where the user can input their horizon       '
' profile information.                                                      '
' The rows displayed are based on the number of points the user selects to  '
' define the horizon.                                                       '


Option Explicit
Private Sub Worksheet_Activate()
   
    ' Resets the active sheet to this sheet
    Me.Activate
    ' Upon sheet activation the first editable cell is selected
    Range("DefineHorizonProfile").Select
   
End Sub


Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    
    Dim currentShtStatus As sheetStatus
    ' If Save is selected run Save Function
    If Target.Range.Address = Range("SaveHS").Address Then
        Call PreModify(Horizon_ShadingSht, currentShtStatus)
        Call SaveXML
        Call PostModify(Horizon_ShadingSht, currentShtStatus)
    
    ElseIf Target.Range.Address = Range("ClearHProf").Address Then
        Call ClearHorizon
    End If
End Sub



Private Sub Worksheet_Change(ByVal Target As Range)

    Dim currentShtStatus As sheetStatus
    Dim n As Integer
    
    ' Used for determining how many rows to show, how many points to save to .csyx
    n = Range("NumHorPts").Value
    
    Call PreModify(Horizon_ShadingSht, currentShtStatus)
    
    ' Hiding and unhiding necessary rows
    If Not Intersect(Target, Range("NumHorPts")) Is Nothing Then
        
        ChartObjects("HorizonChart").Placement = xlFreeFloating
        
        Rows(Range("HAziFirst").row + n & ":" & Range("HAziFirst").row + 359).Hidden = True
        Rows(Range("HAziFirst").row & ":" & Range("HAziFirst").row + n - 1).Hidden = False
        
        If n <> 360 Then
            Range(Range("HAziFirst").Offset(n, 0).Address & ":" & Range("HAziFirst").Offset(359, 0).Address).ClearContents
            Range(Range("HElevFirst").Offset(n, 0).Address & ":" & Range("HElevFirst").Offset(359, 0).Address).ClearContents
        End If
        
        ChartObjects("HorizonChart").Placement = xlMoveAndSize
    
    ElseIf Not Intersect(Target, Range("DefineHorizonProfile")) Is Nothing Then
        If Range("DefineHorizonProfile").Value = "Yes" Then
            Range("NoHorProf").EntireRow.Hidden = True
            Range("YesHorProf").EntireRow.Hidden = False
            Horizon_ShadingSht.Shapes(2).Visible = True
            Rows(Range("HAziFirst").row + n & ":" & Range("HAziFirst").row + 359).Hidden = True
            ChartObjects("HorizonChart").Visible = True
        ElseIf Range("DefineHorizonProfile").Value = "No" Then
            If Range("NoHorProf").EntireRow.Hidden = True And Range("YesHorProf").EntireRow.Hidden = False Then
                'Call ClearHorizon
                Range("NoHorProf").EntireRow.Hidden = False
                Range("YesHorProf").EntireRow.Hidden = True
                ChartObjects("HorizonChart").Visible = False
            End If
        End If
    
    ElseIf Not Intersect(Target, Range("HAzi")) Is Nothing Or Not Intersect(Target, Range("HElev")) Is Nothing Then
        Horizon_ShadingSht.Range("HorizonAzi").Value = vbNullString
        Horizon_ShadingSht.Range("HorizonElev").Value = vbNullString
        If Horizon_ShadingSht.Range("NumHorPts").Value = 1 Then
            If IsEmpty(Horizon_ShadingSht.Range("HAziFirst")) = False And IsEmpty(Horizon_ShadingSht.Range("HElevFirst")) = False Then
                Horizon_ShadingSht.Range("HorizonAzi").Value = Range("HAziFirst").Value
                Horizon_ShadingSht.Range("HorizonElev").Value = Range("HElevFirst").Value
            End If
        Else
            Dim i As Integer
            Dim j As Integer
            Dim incompCount As Integer
            Range(Range("HAziFirst").Address & ":" & Range("HElevFirst").Offset(n - 1, 0).Address).Sort Key1:=Range(Range("HAziFirst").Address & ":" & Range("HAziFirst").Offset(n - 1, 0).Address), Order1:=xlAscending
            For i = n - 1 To 0 Step -1
                incompCount = 0
                ' Finding and deleting horizon values which are missing information
                If IsEmpty(Range("HAziFirst").Offset(i, 0)) = True Or IsEmpty(Range("HElevFirst").Offset(i, 0)) = True Then
                    incompCount = incompCount + 1
                ' Finding and deleting duplicate horizon azimuth values
                Else
                    For j = 0 To i - 1 Step 1
                        If Range("HAziFirst").Offset(i, 0).Value = Range("HAziFirst").Offset(j, 0).Value Then
                            If i <> j And IsEmpty(Horizon_ShadingSht.Range("HAziFirst").Offset(j, 0)) = False Then
                                incompCount = incompCount + 1
                            End If
                        End If
                    Next j
                End If
                If incompCount = 0 Then
                    Horizon_ShadingSht.Range("HorizonAzi").Value = Range("HAziFirst").Offset(i, 0).Value & Horizon_ShadingSht.Range("HorizonAzi").Value
                    Horizon_ShadingSht.Range("HorizonElev").Value = Range("HElevFirst").Offset(i, 0).Value & Horizon_ShadingSht.Range("HorizonElev").Value
                    If i <> 0 Then
                        Horizon_ShadingSht.Range("HorizonAzi").Value = "," & Horizon_ShadingSht.Range("HorizonAzi").Value
                        Horizon_ShadingSht.Range("HorizonElev").Value = "," & Horizon_ShadingSht.Range("HorizonElev").Value
                    End If
                End If
            Next i
        End If
    End If
    
    
    Call PostModify(Horizon_ShadingSht, currentShtStatus)

End Sub

Public Sub ClearHorizon()
    Dim currentShtStatus As sheetStatus
    Call PreModify(Horizon_ShadingSht, currentShtStatus)
    Range("HElev").ClearContents
    Range("HAzi").ClearContents
    Range("HAziFirst").Value = 0
    Range("HElevFirst").Value = 0
    Range("NumHorPts").Value = 1
    Call PostModify(Horizon_ShadingSht, currentShtStatus)
End Sub

