VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BifacialSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                 BIFACIAL SHEET             '
'--------------------------------------------'
' This is the sheet where users enter infor- '
' mation about the bifacial panel model      '

Option Explicit
Private Sub Worksheet_Activate()
    
    ' Resets the active worksheet to the current sheet
    Me.Activate
    ' Upon activating this sheet, the field for the bifacial option is automatically selected
    Range("UseBifacialModel").Select
    ' Upon activating the sheet hide the respective fields
    ' Call BifSwitchFreq(Range("BifAlbFreqVal").Value)
    
End Sub

' WorkSheet_Change Function
' This function is called when a cell is changed

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim currentShtStatus As sheetStatus
    
    Application.EnableEvents = False
    Call PreModify(BifacialSht, currentShtStatus)
    
    If Not Intersect(Target, Range("BifAlbFreqVal")) Is Nothing Then
        Call BifSwitchFreq(Range("BifAlbFreqVal").Value)
    ElseIf Not Intersect(Target, Range("UseBifacialModel")) Is Nothing Then
        If Range("UseBifacialModel").Value = "Yes" Then
            Range("UseBifacialModelRng").EntireRow.Hidden = False
            Range("NoUseBifacialModelRng").EntireRow.Hidden = True
            Call BifSwitchFreq(Range("BifAlbFreqVal").Value)
        Else
            Range("UseBifacialModelRng").EntireRow.Hidden = True
            Range("NoUseBifacialModelRng").EntireRow.Hidden = False
        End If
    End If
    
    Call PostModify(BifacialSht, currentShtStatus)
    Application.EnableEvents = True
    
End Sub

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    Dim currentShtStatus As sheetStatus
    'Call Functions from Hyperlinks
    ' If Save is selected run Save Function
    If Target.Range.Address = Range("SaveBifacial").Address Then
        Call PreModify(BifacialSht, currentShtStatus)
        Call SaveXML
        Call PostModify(BifacialSht, currentShtStatus)
    End If
    
End Sub


' BifSwitchFreq Function
'
' Arguments Selection - The value the user selected
'
' The purpose of this function is to show or hide the
' yearly or monthly albedo, or both, depending on what value the
' user selected
' If the user selected 'Site' then the site albedo is used and both sections are hidden
Sub BifSwitchFreq(ByVal Selection As String)

    Dim currentShtStatus As sheetStatus
    Call PreModify(BifacialSht, currentShtStatus)
    
    ' If yearly albedo is selected show yearly value fields
    If (Selection = "Yearly") Then
        BifacialSht.Range("BifYearlyAlbedo").EntireRow.Hidden = False
        BifacialSht.Range("BifMonthlyAlbedo").EntireRow.Hidden = True
        BifacialSht.Range("BifAlbedoGraph").EntireRow.Hidden = False
        ' If monthly albedo is selected show monthly value fields
    ElseIf (Selection = "Monthly") Then
        BifacialSht.Range("BifYearlyAlbedo").EntireRow.Hidden = True
        BifacialSht.Range("BifMonthlyAlbedo").EntireRow.Hidden = False
        BifacialSht.Range("BifAlbedoGraph").EntireRow.Hidden = False
    ElseIf (Selection = "Site") Then
        BifacialSht.Range("BifYearlyAlbedo").EntireRow.Hidden = True
        BifacialSht.Range("BifMonthlyAlbedo").EntireRow.Hidden = True
        BifacialSht.Range("BifAlbedoGraph").EntireRow.Hidden = True
    End If
    
    Call PostModify(BifacialSht, currentShtStatus)
End Sub
