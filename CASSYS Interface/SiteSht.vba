VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SiteSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                   SITE SHEET               '
'--------------------------------------------'
' The site sheet is where the user enters    '
' general information about the site such    '
' as the geographical location and name,     '
' and also the yearly or monthly albedo.     '

Option Explicit
Private Sub Worksheet_Activate()
    
    ' Resets the active worksheet to the current sheet
    Me.Activate
    ' Upon activating this sheet, the field for the project name is automatically selected
    Range("Name").Select
    ' Upon activating the sheet hide the respective fields
    If (IntroSht.Range("ModeSelect") <> "ASTM E2848 Regression") Then
         Call SwitchFreq(Range("AlbFreqVal").Value)
    End If
    
End Sub

' WorkSheet_Change Function
' This function is called when a cell is changed
'
' The purpose of this function is to convert the
' latitude or longitude from decimal degrees to
' degree-minutes-seconds format and vice versa.
' After these values are calculatted they will
' be inserted into the corresponding fields

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim currentShtStatus As sheetStatus
    Dim OandSShtStatus As sheetStatus
    Dim degree As Double ' The degree value
    Dim minutes As Double ' The minute value
    Dim seconds As Double ' The seconds value
    Dim sign As String
    Dim aux As Double
    
    Application.EnableEvents = False
    Call PreModify(SiteSht, currentShtStatus)
    
    ' Convert from decimal to latitude/longitude
    If Not Intersect(Target, Range("Latitude")) Is Nothing And IsNumeric(Range("Latitude")) = True Then
        convDecimalToLatitude
        Call PreModify(Orientation_and_ShadingSht, OandSShtStatus)
        If Range("Latitude").Value >= 0 Then
            Orientation_and_ShadingSht.Range("AzimuthRefAVAT").Value = 0
          Orientation_and_ShadingSht.Range("AzimuthRefTAXT").Value = 0
            Orientation_and_ShadingSht.Range("AxisAzimuthSAET").Value = 90
        ElseIf Range("Latitude").Value < 0 Then
            Orientation_and_ShadingSht.Range("AzimuthRefAVAT").Value = 180
            Orientation_and_ShadingSht.Range("AzimuthRefTAXT").Value = 180
            Orientation_and_ShadingSht.Range("AxisAzimuthSAET").Value = -90
        End If
        Call PostModify(Orientation_and_ShadingSht, OandSShtStatus)
    End If
    
  
    If Not Intersect(Target, Range("Longitude")) Is Nothing And IsNumeric(Range("Longitude")) = True Then convDecimalToLongitude
    
    ' Convert from latitude/longitude to decimal
    If (Not Intersect(Target, Range("LatDMS")) Is Nothing Or Not Intersect(Target, Range("LatNS")) Is Nothing) And IsNumeric(Range("LatDeg")) And IsNumeric(Range("LatMin")) And IsNumeric(Range("LatSec")) Then
        ' Convert from degress, minutes seconds for latitude to decimal degrees
        convLatitudeToDecimal
        ' NB: Change azimuth reference when latitude changes
        Call PreModify(Orientation_and_ShadingSht, OandSShtStatus)
        If Range("LatNS").Value = "North" Then
            Orientation_and_ShadingSht.Range("AzimuthRefAVAT").Value = 0
            Orientation_and_ShadingSht.Range("AzimuthRefTAXT").Value = 0
            Orientation_and_ShadingSht.Range("AxisAzimuthSAET").Value = 90
        ElseIf Range("LatNS").Value = "South" Then
            Orientation_and_ShadingSht.Range("AzimuthRefAVAT").Value = 180
            Orientation_and_ShadingSht.Range("AzimuthRefTAXT").Value = 180
            Orientation_and_ShadingSht.Range("AxisAzimuthSAET").Value = -90
        End If
        Call PostModify(Orientation_and_ShadingSht, OandSShtStatus)
    End If
    
    
        
    
   
    If (Not Intersect(Target, Range("LongDMS")) Is Nothing Or Not Intersect(Target, Range("LongEW")) Is Nothing) And IsNumeric(Range("LongDeg")) And IsNumeric(Range("LongMin")) And IsNumeric(Range("LongSec")) Then
        ' Convert from degress, minutes seconds for latitude to decimal degrees
        convLongitudeToDecimal
    End If
    
    ' Hide and show rows based on selection of "Use Local Time" or Yearly/Monthly Albedo
    If Not Intersect(Target, Range("UseLocTime")) Is Nothing Then Call ShowHideRef(Range("UseLocTime").Value)
    If Not Intersect(Target, Range("AlbFreqVal")) Is Nothing Then Call SwitchFreq(Range("AlbFreqVal").Value)
   
    
    Call PostModify(SiteSht, currentShtStatus)
    Application.EnableEvents = True
    
End Sub

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    Dim currentShtStatus As sheetStatus
    'Call Functions from Hyperlinks
    ' If Save is selected run Save Function
    If Target.Range.Address = Range("SaveSite").Address Then
        Call PreModify(SiteSht, currentShtStatus)
        Call SaveXML
        Call PostModify(SiteSht, currentShtStatus)
    End If
    
End Sub

' ShowHideRef Function
'
' Arguments Selection - The value the user selected
'
' The purpose of this function is to show or hide the
' reference meridian field depending on what the user
' selected

Sub ShowHideRef(ByVal Selection As String)

    Dim currentShtStatus As sheetStatus
    
    Call PreModify(SiteSht, currentShtStatus)
    
    If (Selection = "No") Then
        ' Hide reference merdian cells and lock them
        Range("RefMerCells").NumberFormat = ";;;"
        Range("RefMer").Locked = True
        With Range("RefMer")
            .Borders.LineStyle = xlNone
            .Interior.Color = ColourWhite
        End With
        
    ElseIf (Selection = "Yes") Then
        ' show reference merdian cells and unlock them
        Range("RefMerCells").NumberFormat = "General"
        Range("RefMer").Locked = True
        With Range("RefMer")
            .Borders.LineStyle = xlContinuous
        End With
    End If
    
    Call PostModify(SiteSht, currentShtStatus)
    
End Sub

' SwitchFreq Function
'
' Arguments Selection - The value the user selected
'
' The purpose of this function is to show or hide the
' yearly or monthly albedo depending n what value the
' user selected
Sub SwitchFreq(ByVal Selection As String)

    Dim currentShtStatus As sheetStatus

    Call PreModify(SiteSht, currentShtStatus)
    ' If yearly albedo is selected show yearly value fields
    If (Selection = "Yearly") Then
        SiteSht.Range("AlbDefinitions").EntireRow.Hidden = False
        SiteSht.Range("YearlyAlbedo").EntireRow.Hidden = False
        SiteSht.Range("MonthlyAlbedo").EntireRow.Hidden = True
    ElseIf (Selection = "Monthly") Then
        SiteSht.Range("AlbDefinitions").EntireRow.Hidden = False
        ' If monthly albedo is selected show monthly value fields
        SiteSht.Range("YearlyAlbedo").EntireRow.Hidden = True
        SiteSht.Range("MonthlyAlbedo").EntireRow.Hidden = False
    ElseIf (Selection = "From Climate File") Then
        SiteSht.Range("AlbDefinitions").EntireRow.Hidden = True
    End If
    Call PostModify(SiteSht, currentShtStatus)
    
End Sub
Private Sub convDecimalToLatitude()

    Dim degree As Double ' The degree value
    Dim minutes As Double ' The minute value
    Dim seconds As Double ' The seconds value
    Dim sign As String
    Dim aux As Double

    ' Calculate the degree, minutes and seconds value
    ' Calculate for positive value
    If (Range("Latitude").Value >= 0) Then
        sign = "North"
        aux = Application.WorksheetFunction.Floor(Range("Latitude").Value, 1)
    Else
        sign = "South"
        aux = Application.WorksheetFunction.Ceiling(Range("Latitude").Value, 1)
    End If
    
    ' NB: Changed Math.Round to WorksheetFunction.RoundDown so minutes were rounded down, eliminating negative second values
    seconds = Math.Abs(Math.Round((Range("Latitude").Value - aux) * 3600))
    ' NB: Adding the RoundDown(degree) ensures that the LatDeg value is correctly rounded
    degree = Math.Abs(aux) + WorksheetFunction.RoundDown(seconds / 3600, 0)
    minutes = (seconds Mod 3600) / 60
    seconds = Math.Round((minutes - WorksheetFunction.RoundDown(minutes, 0)) * 60)
    minutes = WorksheetFunction.RoundDown(minutes, 0)
    
    
    ' Output the values to the cell
    Range("LatDeg").Value = degree
    Range("LatMin").Value = minutes
    Range("LatSec").Value = seconds
    Range("LatNS").Value = sign
     
End Sub

Private Sub convDecimalToLongitude()
    
    Dim degree As Double ' The degree value
    Dim minutes As Double ' The minute value
    Dim seconds As Double ' The seconds value
    Dim sign As String
    Dim aux As Double

    ' Calculate the degree, minutes and seconds value
    ' Calculate for positive value
    If (Range("Longitude").Value >= 0) Then
        sign = "East"
        aux = Application.WorksheetFunction.Floor(Range("Longitude").Value, 1)
    Else
        sign = "West"
        aux = Application.WorksheetFunction.Ceiling(Range("Longitude").Value, 1)
    End If
    
    ' NB: Changed Math.Round to WorksheetFunction.RoundDown so minutes were rounded down, eliminating negative second values
    seconds = Math.Abs(Math.Round((Range("Longitude").Value - aux) * 3600))
    ' NB: Adding the RoundDown(degree) ensures that the LongDeg value is correctly rounded
    degree = Math.Abs(aux) + WorksheetFunction.RoundDown(seconds / 3600, 0)
    minutes = (seconds Mod 3600) / 60
    seconds = Math.Round((minutes - WorksheetFunction.RoundDown(minutes, 0)) * 60)
    minutes = WorksheetFunction.RoundDown(minutes, 0)
    
    ' Output the values to the cell
    Range("LongDeg").Value = degree
    Range("LongMin").Value = minutes
    Range("LongSec").Value = seconds
    Range("LongEW").Value = sign

End Sub

Private Sub convLongitudeToDecimal()

    ' Convert from degress, minutes seconds for latitude to decimal degrees
    Range("Longitude").Value = Math.Abs(Range("LongDeg").Value) + Math.Abs(Range("LongMin") / 60) + Math.Abs(Range("LongSec") / 3600)
    If (Range("LongEW").Value = "West") Then
        Range("Longitude").Value = Range("Longitude").Value * (-1)
    End If

End Sub

Private Sub convLatitudeToDecimal()

    ' Convert from degress, minutes seconds for latitude to decimal degrees
    Range("Latitude").Value = Math.Abs(Range("LatDeg").Value) + Math.Abs(Range("LatMin") / 60) + Math.Abs(Range("LatSec") / 3600)
    If (Range("LatNS").Value = "South") Then
        Range("Latitude").Value = Range("Latitude").Value * (-1)
    End If

End Sub
