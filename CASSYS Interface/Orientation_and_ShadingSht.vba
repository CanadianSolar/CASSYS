VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Orientation_and_ShadingSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                  ORIENTATION AND SHADING SHEET                           '

' The orientation and shading sheet is the component of the interface      '
' where the user specifies the near shading and cell based shading limits. '
' User entry fields on this sheet are changed based on whether the user    '
' selects a Fixed Tilt model or an Unlimited Rows model.                   '

Option Explicit
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    
    Dim currentShtStatus As sheetStatus
    ' If Save is selected run Save Function
    If Target.Range.Address = Range("SaveOS").Address Then
        Call PreModify(Orientation_and_ShadingSht, currentShtStatus)
        Call SaveXML
        Call PostModify(Orientation_and_ShadingSht, currentShtStatus)
    End If
    
End Sub
Private Sub Worksheet_Activate()
    
    ' Upon worksheet activation, selects first editable cell
    Me.Activate
    Range("OrientType").Select
    

End Sub

' WorkSheet_Change Function
' This function is called when a cell is changed
'
' The purpose of the function is to show or hide
' the fields corresponding to the value chosen in
' the drop down list
' NB: changing code so when OrientType changes rows hide rather than changing number format 27/01/2016
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim currentShtStatus As sheetStatus
    Dim InputShtStatus As sheetStatus
    Dim BifacialShtStatus As sheetStatus
    Call PreModify(Orientation_and_ShadingSht, currentShtStatus)
    
    ' If the list value was changed
    If Not Intersect(Target, Range("OrientType")) Is Nothing Then

'--------Commenting out Iterative Functionality for this version--------'
'        ' Creating list of variable parameters for iterative mode
'        Call OrientationOutputValidation
        
        ' If the value was set to fixed tilted plane
        If (Range("OrientType").Value = "Fixed Tilted Plane") Then
            Orientation_and_ShadingSht.ChartObjects("Chart 5").Visible = False
            
            ' Hide Info Based on new named ranges
            Range("UnlimitedParam").EntireRow.Hidden = True
            Range("FixedTiltParam").EntireRow.Hidden = False
            Range("TrkEWParam").EntireRow.Hidden = True
            Range("TrkNSParam").EntireRow.Hidden = True
            Range("TrkVrtParam").EntireRow.Hidden = True
            Range("TrkTwoParam").EntireRow.Hidden = True
            Range("TrkTwoNSParam").EntireRow.Hidden = True
            Range("TrkTwoEWParam").EntireRow.Hidden = True
            Range("TiltnRollParam").EntireRow.Hidden = True
            Range("ArrayTypeDescribe") = "This array type only requires Plane Tilt and Azimuth."
            
            Call PreModify(InputFileSht, InputShtStatus)
            InputFileSht.Range("MeterAzimuthDescribe").Value = "    Angle the tilted irradiance meter faces with respect to true south, [+] if W, [-] if E"
            InputFileSht.Range("MeterTiltDescribe").Value = "    The tilt at which the tilted irradiance is measured "
            InputFileSht.Range("MeterTilt").Interior.Color = RGB(255, 255, 255)
            InputFileSht.Range("MeterAzimuth").Interior.Color = RGB(255, 255, 255)
            InputFileSht.Range("MeterTilt").Locked = False
            InputFileSht.Range("MeterAzimuth").Locked = False
            If InputFileSht.Range("MeterTilt").Value = "N/A" Then
                InputFileSht.Range("MeterTilt").Value = ""
                InputFileSht.Range("MeterAzimuth").Value = ""
            End If
            Call PostModify(InputFileSht, InputShtStatus)
            
            Call PreModify(BifacialSht, BifacialShtStatus)
            BifacialSht.Range("UseBifacialModel").Interior.Color = RGB(255, 255, 255)
            BifacialSht.Range("UseBifacialModel").Locked = False
            Call PostModify(BifacialSht, BifacialShtStatus)
            
        ElseIf (Range("OrientType").Value = "Fixed Tilted Plane Seasonal Adjustment") Then
            Orientation_and_ShadingSht.ChartObjects("Chart 5").Visible = False
            
            Range("UnlimitedParam").EntireRow.Hidden = True
            Range("FixedTiltParam").EntireRow.Hidden = True
            Range("TrkEWParam").EntireRow.Hidden = True
            Range("TrkNSParam").EntireRow.Hidden = True
            Range("TrkVrtParam").EntireRow.Hidden = True
            Range("TrkTwoParam").EntireRow.Hidden = True
            Range("TrkTwoNSParam").EntireRow.Hidden = True
            Range("TrkTwoEWParam").EntireRow.Hidden = True
            Range("TiltnRollParam").EntireRow.Hidden = True
            Range("FixedPlaneSeasonalParam").EntireRow.Hidden = False
            Range("ArrayTypeDescribe") = "This array type only requires Plane Tilt and Azimuth."
            
            Call PreModify(InputFileSht, InputShtStatus)
            InputFileSht.Range("MeterAzimuthDescribe").Value = "    Angle the tilted irradiance meter faces with respect to true south, [+] if W, [-] if E"
            InputFileSht.Range("MeterTiltDescribe").Value = "    The tilt at which the tilted irradiance is measured "
            InputFileSht.Range("MeterTilt").Interior.Color = RGB(255, 255, 255)
            InputFileSht.Range("MeterAzimuth").Interior.Color = RGB(255, 255, 255)
            InputFileSht.Range("MeterTilt").Locked = False
            InputFileSht.Range("MeterAzimuth").Locked = False
            If InputFileSht.Range("MeterTilt").Value = "N/A" Then
                InputFileSht.Range("MeterTilt").Value = ""
                InputFileSht.Range("MeterAzimuth").Value = ""
            End If
            Call PostModify(InputFileSht, InputShtStatus)
            
            Call PreModify(BifacialSht, BifacialShtStatus)
            BifacialSht.Range("UseBifacialModel").Interior.Color = RGB(204, 192, 218)
            BifacialSht.Range("UseBifacialModel").Value = "No"
            BifacialSht.Range("UseBifacialModel").Locked = True
            Call PostModify(BifacialSht, BifacialShtStatus)
            
        ElseIf (Range("OrientType").Value = "Unlimited Rows") Then
            Orientation_and_ShadingSht.ChartObjects("Chart 5").Visible = True

            ' Hide Info Based on new named ranges
            Range("UnlimitedParam").EntireRow.Hidden = False
            Range("FixedTiltParam").EntireRow.Hidden = True
            Range("TrkEWParam").EntireRow.Hidden = True
            Range("TrkNSParam").EntireRow.Hidden = True
            Range("TrkVrtParam").EntireRow.Hidden = True
            Range("TrkTwoParam").EntireRow.Hidden = True
            Range("TrkTwoNSParam").EntireRow.Hidden = True
            Range("TrkTwoEWParam").EntireRow.Hidden = True
            Range("TiltnRollParam").EntireRow.Hidden = True
            Range("ArrayTypeDescribe") = "This array type is used when rows are very long compared to their width"
            Call PreModify(InputFileSht, InputShtStatus)
            InputFileSht.Range("MeterAzimuthDescribe").Value = "    Angle the tilted irradiance meter faces with respect to true south, [+] if W, [-] if E"
            InputFileSht.Range("MeterTiltDescribe").Value = "    The tilt at which the tilted irradiance is measured "
            InputFileSht.Range("MeterTilt").Interior.Color = RGB(255, 255, 255)
            InputFileSht.Range("MeterAzimuth").Interior.Color = RGB(255, 255, 255)
            InputFileSht.Range("MeterTilt").Locked = False
            InputFileSht.Range("MeterAzimuth").Locked = False
            If InputFileSht.Range("MeterTilt").Value = "N/A" Then
                InputFileSht.Range("MeterTilt").Value = ""
                InputFileSht.Range("MeterAzimuth").Value = ""
            End If
            Call PostModify(InputFileSht, InputShtStatus)
            
            Call UpdateCellShading
            
            Call PreModify(BifacialSht, BifacialShtStatus)
            BifacialSht.Range("UseBifacialModel").Interior.Color = RGB(255, 255, 255)
            BifacialSht.Range("UseBifacialModel").Locked = False
            Call PostModify(BifacialSht, BifacialShtStatus)
        
        ElseIf (Range("OrientType").Value = "Single Axis Elevation Tracking (E-W)") Then
            Orientation_and_ShadingSht.ChartObjects("Chart 5").Visible = False
            ' Hide Info Based on new named ranges
            Range("UnlimitedParam").EntireRow.Hidden = True
            Range("FixedTiltParam").EntireRow.Hidden = True
            Range("TrkEWParam").EntireRow.Hidden = False
            Range("TrkNSParam").EntireRow.Hidden = True
            Range("TrkVrtParam").EntireRow.Hidden = True
            Range("TrkTwoParam").EntireRow.Hidden = True
            Range("TrkTwoNSParam").EntireRow.Hidden = True
            Range("TrkTwoEWParam").EntireRow.Hidden = True
            Range("TiltnRollParam").EntireRow.Hidden = True
            Range("ArrayTypeDescribe") = "This array type is used when modules track the sun, rotating on an East-West axis"
            
            ' Update backtracking and cell shading options depending upon the number of rows
            Call UpdateNumberOfRowsSAET
                        
            Call PreModify(InputFileSht, InputShtStatus)
            InputFileSht.Range("MeterAzimuthDescribe").Value = "    Tracker selected. Meter azimuth is equal to the tracker surface azimuth"
            InputFileSht.Range("MeterTiltDescribe").Value = "    Tracker selected. Meter tilt is equal to the tracker surface tilt "
            InputFileSht.Range("MeterTilt").Value = "N/A"
            InputFileSht.Range("MeterAzimuth").Value = "N/A"
            InputFileSht.Range("MeterTilt").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterAzimuth").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterTilt").Locked = True
            InputFileSht.Range("MeterAzimuth").Locked = True
            Call PostModify(InputFileSht, InputShtStatus)
                        
            Call UpdateCellShadingSAET
                        
            Call PreModify(BifacialSht, BifacialShtStatus)
            BifacialSht.Range("UseBifacialModel").Interior.Color = RGB(255, 255, 255)
            BifacialSht.Range("UseBifacialModel").Locked = False
            Call PostModify(BifacialSht, BifacialShtStatus)
        
        ElseIf (Range("OrientType").Value = "Single Axis Horizontal Tracking (N-S)") Then
            Orientation_and_ShadingSht.ChartObjects("Chart 5").Visible = False
            ' Hide Info Based on new named ranges
            Range("UnlimitedParam").EntireRow.Hidden = True
            Range("FixedTiltParam").EntireRow.Hidden = True
            Range("TrkEWParam").EntireRow.Hidden = True
            Range("TrkNSParam").EntireRow.Hidden = False
            Range("TrkVrtParam").EntireRow.Hidden = True
            Range("TrkTwoParam").EntireRow.Hidden = True
            Range("TrkTwoNSParam").EntireRow.Hidden = True
            Range("TrkTwoEWParam").EntireRow.Hidden = True
            Range("TiltnRollParam").EntireRow.Hidden = True
            Range("ArrayTypeDescribe") = "This array type is used when modules track the sun, rotating on a North-South axis"
            
            ' Update backtracking and cell shading options depending upon the number of rows
            Call UpdateNumberOfRowsSAST
            
            Call PreModify(InputFileSht, InputShtStatus)
            InputFileSht.Range("MeterAzimuthDescribe").Value = "    Tracker selected. Meter azimuth is equal to the tracker surface azimuth"
            InputFileSht.Range("MeterTiltDescribe").Value = "    Tracker selected. Meter tilt is equal to the tracker surface tilt "
            InputFileSht.Range("MeterTilt").Value = "N/A"
            InputFileSht.Range("MeterAzimuth").Value = "N/A"
            InputFileSht.Range("MeterTilt").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterAzimuth").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterTilt").Locked = True
            InputFileSht.Range("MeterAzimuth").Locked = True
            Call PostModify(InputFileSht, InputShtStatus)
            
            Call UpdateCellShadingSAST
                        
            Call PreModify(BifacialSht, BifacialShtStatus)
            BifacialSht.Range("UseBifacialModel").Interior.Color = RGB(255, 255, 255)
            BifacialSht.Range("UseBifacialModel").Locked = False
            Call PostModify(BifacialSht, BifacialShtStatus)
            
        ElseIf (Range("OrientType").Value = "Tilt and Roll Tracking") Then
            Orientation_and_ShadingSht.ChartObjects("Chart 5").Visible = False
            ' Hide Info Based on new named ranges
            Range("UnlimitedParam").EntireRow.Hidden = True
            Range("FixedTiltParam").EntireRow.Hidden = True
            Range("TrkEWParam").EntireRow.Hidden = True
            Range("TrkNSParam").EntireRow.Hidden = True
            Range("TrkVrtParam").EntireRow.Hidden = True
            Range("TrkTwoParam").EntireRow.Hidden = True
            Range("TrkTwoNSParam").EntireRow.Hidden = True
            Range("TrkTwoEWParam").EntireRow.Hidden = True
            Range("TiltnRollParam").EntireRow.Hidden = False
            Range("ArrayTypeDescribe") = "This array type is used when modules track the sun, rotating on a tilted axis"
            
            Call PreModify(InputFileSht, InputShtStatus)
            InputFileSht.Range("MeterAzimuthDescribe").Value = "    Tracker selected. Meter azimuth is equal to the tracker surface azimuth"
            InputFileSht.Range("MeterTiltDescribe").Value = "    Tracker selected. Meter tilt is equal to the tracker surface tilt "
            InputFileSht.Range("MeterTilt").Value = "N/A"
            InputFileSht.Range("MeterAzimuth").Value = "N/A"
            InputFileSht.Range("MeterTilt").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterAzimuth").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterTilt").Locked = True
            InputFileSht.Range("MeterAzimuth").Locked = True
            Call PostModify(InputFileSht, InputShtStatus)
        
            Call PreModify(BifacialSht, BifacialShtStatus)
            BifacialSht.Range("UseBifacialModel").Interior.Color = RGB(204, 192, 218)
            BifacialSht.Range("UseBifacialModel").Value = "No"
            BifacialSht.Range("UseBifacialModel").Locked = True
            Call PostModify(BifacialSht, BifacialShtStatus)
        
        ElseIf (Range("OrientType").Value = "Azimuth (Vertical Axis) Tracking") Then
            Orientation_and_ShadingSht.ChartObjects("Chart 5").Visible = False
            ' Hide Info Based on new named ranges
            Range("UnlimitedParam").EntireRow.Hidden = True
            Range("FixedTiltParam").EntireRow.Hidden = True
            Range("TrkEWParam").EntireRow.Hidden = True
            Range("TrkNSParam").EntireRow.Hidden = True
            Range("TrkVrtParam").EntireRow.Hidden = False
            Range("TrkTwoParam").EntireRow.Hidden = True
            Range("TrkTwoNSParam").EntireRow.Hidden = True
            Range("TiltnRollParam").EntireRow.Hidden = True
            Range("TrkTwoEWParam").EntireRow.Hidden = True
            Range("ArrayTypeDescribe") = "This array type is used when modules track the sun, rotating on a vertical axis"
            Call PreModify(InputFileSht, InputShtStatus)
            InputFileSht.Range("MeterAzimuthDescribe").Value = "    Tracker selected. Meter azimuth is equal to the tracker surface azimuth"
            InputFileSht.Range("MeterTiltDescribe").Value = "    Tracker selected. Meter tilt is equal to the tracker surface tilt "
            InputFileSht.Range("MeterTilt").Value = "N/A"
            InputFileSht.Range("MeterAzimuth").Value = "N/A"
            InputFileSht.Range("MeterTilt").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterAzimuth").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterTilt").Locked = True
            InputFileSht.Range("MeterAzimuth").Locked = True
            Call PostModify(InputFileSht, InputShtStatus)
        
        
            Call PreModify(BifacialSht, BifacialShtStatus)
            BifacialSht.Range("UseBifacialModel").Interior.Color = RGB(204, 192, 218)
            BifacialSht.Range("UseBifacialModel").Value = "No"
            BifacialSht.Range("UseBifacialModel").Locked = True
            Call PostModify(BifacialSht, BifacialShtStatus)
        
        ElseIf (Range("OrientType").Value = "Two Axis Tracking") Then
            Orientation_and_ShadingSht.ChartObjects("Chart 5").Visible = False
            ' Hide Info Based on new named ranges
            Range("UnlimitedParam").EntireRow.Hidden = True
            Range("FixedTiltParam").EntireRow.Hidden = True
            Range("TrkEWParam").EntireRow.Hidden = True
            Range("TrkNSParam").EntireRow.Hidden = True
            Range("TrkVrtParam").EntireRow.Hidden = True
            Range("TrkTwoParam").EntireRow.Hidden = False
            Range("TrkTwoNSParam").EntireRow.Hidden = True
            Range("TrkTwoEWParam").EntireRow.Hidden = True
            Range("TiltnRollParam").EntireRow.Hidden = True
            Range("ArrayTypeDescribe") = "This array type is used when modules track the sun, rotating both on a vertical and horizontal axis"
            Call PreModify(InputFileSht, InputShtStatus)
            InputFileSht.Range("MeterAzimuthDescribe").Value = "    Tracker selected. Meter azimuth is equal to the tracker surface azimuth"
            InputFileSht.Range("MeterTiltDescribe").Value = "    Tracker selected. Meter tilt is equal to the tracker surface tilt "
            InputFileSht.Range("MeterTilt").Value = "N/A"
            InputFileSht.Range("MeterAzimuth").Value = "N/A"
            InputFileSht.Range("MeterTilt").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterAzimuth").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterTilt").Locked = True
            InputFileSht.Range("MeterAzimuth").Locked = True
            Call PostModify(InputFileSht, InputShtStatus)
        
            Call PreModify(BifacialSht, BifacialShtStatus)
            BifacialSht.Range("UseBifacialModel").Interior.Color = RGB(204, 192, 218)
            BifacialSht.Range("UseBifacialModel").Value = "No"
            BifacialSht.Range("UseBifacialModel").Locked = True
            Call PostModify(BifacialSht, BifacialShtStatus)
        
        ElseIf (Range("OrientType").Value = "Tracking, Two Axis-Frame N-S") Then
            Orientation_and_ShadingSht.ChartObjects("Chart 5").Visible = False
            ' Hide Info Based on new named ranges
            Range("UnlimitedParam").EntireRow.Hidden = True
            Range("FixedTiltParam").EntireRow.Hidden = True
            Range("TrkEWParam").EntireRow.Hidden = True
            Range("TrkNSParam").EntireRow.Hidden = True
            Range("TrkVrtParam").EntireRow.Hidden = True
            Range("TrkTwoParam").EntireRow.Hidden = True
            Range("TrkTwoNSParam").EntireRow.Hidden = False
            Range("TrkTwoEWParam").EntireRow.Hidden = True
            Call PreModify(InputFileSht, InputShtStatus)
            InputFileSht.Range("MeterAzimuthDescribe").Value = "    Tracker selected. Meter azimuth is equal to the tracker surface azimuth"
            InputFileSht.Range("MeterTiltDescribe").Value = "    Tracker selected. Meter tilt is equal to the tracker surface tilt "
            InputFileSht.Range("MeterTilt").Value = "N/A"
            InputFileSht.Range("MeterAzimuth").Value = "N/A"
            InputFileSht.Range("MeterTilt").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterAzimuth").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterTilt").Locked = True
            InputFileSht.Range("MeterAzimuth").Locked = True
            Call PostModify(InputFileSht, InputShtStatus)
        
            Call PreModify(BifacialSht, BifacialShtStatus)
            BifacialSht.Range("UseBifacialModel").Interior.Color = RGB(204, 192, 218)
            BifacialSht.Range("UseBifacialModel").Value = "No"
            BifacialSht.Range("UseBifacialModel").Locked = True
            Call PostModify(BifacialSht, BifacialShtStatus)
        
        ElseIf (Range("OrientType").Value = "Tracking, Two Axis-Frame E-W") Then
            Orientation_and_ShadingSht.ChartObjects("Chart 5").Visible = False
            ' Hide Info Based on new named ranges
            Range("UnlimitedParam").EntireRow.Hidden = True
            Range("FixedTiltParam").EntireRow.Hidden = True
            Range("TrkEWParam").EntireRow.Hidden = True
            Range("TrkNSParam").EntireRow.Hidden = True
            Range("TrkVrtParam").EntireRow.Hidden = True
            Range("TrkTwoParam").EntireRow.Hidden = True
            Range("TrkTwoNSParam").EntireRow.Hidden = True
            Range("TrkTwoEWParam").EntireRow.Hidden = False
            Call PreModify(InputFileSht, InputShtStatus)
            InputFileSht.Range("MeterAzimuthDescribe").Value = "    Tracker selected. Meter azimuth is equal to the tracker surface azimuth"
            InputFileSht.Range("MeterTiltDescribe").Value = "    Tracker selected. Meter tilt is equal to the tracker surface tilt "
            InputFileSht.Range("MeterTilt").Value = "N/A"
            InputFileSht.Range("MeterAzimuth").Value = "N/A"
            InputFileSht.Range("MeterTilt").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterAzimuth").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterTilt").Locked = True
            InputFileSht.Range("MeterAzimuth").Locked = True
            Call PostModify(InputFileSht, InputShtStatus)
            
            Call PreModify(BifacialSht, BifacialShtStatus)
            BifacialSht.Range("UseBifacialModel").Interior.Color = RGB(204, 192, 218)
            BifacialSht.Range("UseBifacialModel").Value = "No"
            BifacialSht.Range("UseBifacialModel").Locked = True
            Call PostModify(BifacialSht, BifacialShtStatus)
            
        End If
        
    ' Seasonal tilt
    ElseIf Not Intersect(Target, Range("SeasonalTiltDates")) Is Nothing Then
        If (Range("SummerMonth").Value = Range("WinterMonth").Value And Range("SummerDay").Value = Range("WinterDay").Value) Then
            MsgBox "Summer Tilt and Winter Tilt cannot start on the same day", vbExclamation, "Invalid CASSYS Input"
            Range("WinterDay") = Range("WinterDay").Value + 1
        End If
    
    ' Cell based shading
    
    ElseIf Not Intersect(Target, Range("UseCellVal")) Is Nothing Then
        ' If use cell based shading is true, show cell based shading parmaters
        Call UpdateCellShading
        
    ElseIf Not Intersect(Target, Range("UseCellValSAET")) Is Nothing Then
        Call UpdateCellShadingSAET
   
    ElseIf Not Intersect(Target, Range("UseCellValSAST")) Is Nothing Then
        Call UpdateCellShadingSAST
        
    ElseIf Not Intersect(Target, Range("PlaneTilt")) Is Nothing Then
        Application.EnableEvents = False
        ' make sure that the plane tilt is a valid number
        If (Range("PlaneTilt").Value < 0 Or IsNumeric(Range("PlaneTilt").Value) = False) Then
            Range("PlaneTilt").Value = 0
        ElseIf (Range("PlaneTilt").Value > 90) Then
            Range("PlaneTilt").Value = 90
        End If
        Application.EnableEvents = True
    
    
    ElseIf Not Intersect(Target, Range("RowsBlockSAET")) Is Nothing Then
        Call UpdateNumberOfRowsSAET
        
    ElseIf Not Intersect(Target, Range("RowsBlockSAST")) Is Nothing Then
        Call UpdateNumberOfRowsSAST
    
    End If
    
    Call PostModify(Orientation_and_ShadingSht, currentShtStatus)
    
End Sub

'--------Commenting out Iterative Functionality for this version--------'

Sub OrientationOutputValidation()
'    Dim i As Integer
'
'    ' Creating list of available parameters for iteration sheet
'
'    ' unprotect sheet to allow cell validation
'    IterativeSht.Unprotect
'
'    IterativeSht.Range("Z:Z").ClearContents
'    IterativeSht.Range("AA:AA").ClearContents
'    i = 1
'    ' Writing orientation and shading related parameters to column Z and their respective XML path to column AA in Iteration mode sheet
'    If Range("OrientType").Value = "Fixed Tilted Plane" Or Range("OrientType").Value = "Unlimited Rows" Or Range("OrientType").Value = "Azimuth (Vertical Axis) Tracking" Then
'        IterativeSht.Range("Z" & IterativeSht.Range("Z" & Rows.Count).End(xlUp).row + 1).Value = "PlaneTilt (degrees)"
'        IterativeSht.Range("AA" & IterativeSht.Range("AA" & Rows.Count).End(xlUp).row + 1).Value = "/Site/Orientation_and_Shading/PlaneTilt"
'        i = i + 1
'    End If
'
'    ElseIf Range("OrientType").Value = "Fixed Tilted Plane" Or Range("OrientType").Value = "Fixed Tilted Plane Seasonal Adjustment" Or Range("OrientType").Value = "Unlimited Rows" Or Range("OrientType").Value = "Azimuth (Vertical Axis) Tracking" Then
'        IterativeSht.Range("Z" & IterativeSht.Range("Z" & Rows.Count).End(xlUp).row + 1).Value = "Azimuth (degrees)"
'        IterativeSht.Range("AA" & IterativeSht.Range("AA" & Rows.Count).End(xlUp).row + 1).Value = "/Site/Orientation_and_Shading/Azimuth"
'        i = i + 1
'    End If
'
'    ' Creating validation
'    With IterativeSht.Range("ParamName").Validation
'    .Delete
'    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'    xlBetween, Formula1:="=$Z$2:$Z$" & i
'    .IgnoreBlank = True
'    .InCellDropdown = True
'    .InputTitle = ""
'    .ErrorTitle = ""
'    .InputMessage = ""
'    .errorMessage = ""
'    .ShowInput = True
'    .ShowError = True
'    End With
'
'    ' Protecting sheet after setting cell validation
'    IterativeSht.Protect
'
'    ' Set output to first element in list
'    ' This ensures the output selected is available for the selected mode
'    IterativeSht.Range("ParamName").Value = IterativeSht.Range("Z2").Value

End Sub

' Respond to a change in the Use Cell-Based Shading Effect selection
' Unlimited rows case
Sub UpdateCellShading()
    If Range("UseCellVal").Value = "Yes" Then
        Range("CellBasedInput").EntireRow.Hidden = False
        Range("CellBasedBlank").EntireRow.Hidden = True
        Range("WidOfStr").Interior.Color = RGB(204, 192, 218)
    ElseIf Range("UseCellVal").Value = "No" Then
        ' if use cell based shading is false, hide cell based shading parmaters
        Range("CellBasedInput").EntireRow.Hidden = True
        Range("CellBasedBlank").EntireRow.Hidden = False
    End If
End Sub

' Respond to a change in the number of rows by updating the backtracking and cell-based shading effect options
' Single axis E-W case
Sub UpdateNumberOfRowsSAET()
    If Range("RowsBlockSAET") <= 1 Then
        Range("PitchSAET").Value = 0
        Range("BacktrackOptSAET").Value = "No"
        Range("UseCellValSAET").Value = "No"
        ' Lock Cells
        Range("PitchSAET").Locked = True
        Range("BacktrackOptSAET").Locked = True
        Range("UseCellValSAET").Locked = True
        ' Change colour of locked cells to purple
        Range("PitchSAET").Interior.Color = RGB(204, 192, 218)
        Range("BacktrackOptSAET").Interior.Color = RGB(204, 192, 218)
        Range("UseCellValSAET").Interior.Color = RGB(204, 192, 218)
    ElseIf Range("RowsBlockSAET").Value > 1 Then
        ' Unlock cells
        Range("PitchSAET").Locked = False
        Range("BacktrackOptSAET").Locked = False
        Range("UseCellValSAET").Locked = False
        ' Change colour of unlocked cells to white
        Range("PitchSAET").Interior.Color = RGB(255, 255, 255)
        Range("BacktrackOptSAET").Interior.Color = RGB(255, 255, 255)
        Range("UseCellValSAET").Interior.Color = RGB(255, 255, 255)
    End If
End Sub

' Respond to a change in the Use Cell-Based Shading Effect selection
' Single axis E-W case
Sub UpdateCellShadingSAET()
    If Range("UseCellValSAET").Value = "Yes" Then
        Range("CellBasedInputSAET").EntireRow.Hidden = False
        Range("CellBasedBlankSAET").EntireRow.Hidden = True
        Range("WidOfStrSAET").Interior.Color = RGB(204, 192, 218)
    ElseIf Range("UseCellValSAET").Value = "No" Then
        ' if use cell based shading is false, hide cell based shading parmaters
        Range("CellBasedInputSAET").EntireRow.Hidden = True
        Range("CellBasedBlankSAET").EntireRow.Hidden = False
    End If
End Sub

' Respond to a change in the number of rows by updating the backtracking and cell-based shading effect options
' Single axis N-S case
Sub UpdateNumberOfRowsSAST()
    If Range("RowsBlockSAST") <= 1 Then
        Range("PitchSAST").Value = 0
        Range("BacktrackOptSAST").Value = "No"
        Range("UseCellValSAST").Value = "No"
        ' Lock Cells
        Range("PitchSAST").Locked = True
        Range("BacktrackOptSAST").Locked = True
        Range("UseCellValSAST").Locked = True
        ' Change colour of locked cells to purple
        Range("PitchSAST").Interior.Color = RGB(204, 192, 218)
        Range("BacktrackOptSAST").Interior.Color = RGB(204, 192, 218)
        Range("UseCellValSAST").Interior.Color = RGB(204, 192, 218)
    ElseIf Range("RowsBlockSAST").Value > 1 Then
        ' Unlock cells
        Range("PitchSAST").Locked = False
        Range("BacktrackOptSAST").Locked = False
        Range("UseCellValSAST").Locked = False
        ' Change colour of unlocked cells to white
        Range("PitchSAST").Interior.Color = RGB(255, 255, 255)
        Range("BacktrackOptSAST").Interior.Color = RGB(255, 255, 255)
        Range("UseCellValSAST").Interior.Color = RGB(255, 255, 255)
    End If
End Sub

' Respond to a change in the Use Cell-Based Shading Effect selection
' Single axis N-S case
Sub UpdateCellShadingSAST()
    If Range("UseCellValSAST").Value = "Yes" Then
        Range("CellBasedInputSAST").EntireRow.Hidden = False
        Range("CellBasedBlankSAST").EntireRow.Hidden = True
        Range("WidOfStrSAST").Interior.Color = RGB(204, 192, 218)
    ElseIf Range("UseCellValSAST").Value = "No" Then
        ' if use cell based shading is false, hide cell based shading parmaters
        Range("CellBasedInputSAST").EntireRow.Hidden = True
        Range("CellBasedBlankSAST").EntireRow.Hidden = False
    End If
End Sub

