VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpectralSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                              SPECTRAL SHEET                               '
'---------------------------------------------------------------------------'
' The Spectral Sheet is where the user can enter the parameters of a very   '
' simplified 'spectral' model, where plane of array irradiance is simply    '
' corrected according to the clearness index                                '

Option Explicit

Private Sub Worksheet_Activate()
   
    ' Resets the active sheet to this sheet
    Me.Activate
    ' Upon sheet activation the first editable cell is selected
    Range("UseSpectralModel").Select
   
End Sub

' Respond to clicks on hyperlinks
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    
    Dim currentShtStatus As sheetStatus
    ' If Save is selected run Save Function
    If Target.Range.Address = Range("SaveSpectral").Address Then
        Call PreModify(SpectralSht, currentShtStatus)
        Call SaveXML
        Call PostModify(SpectralSht, currentShtStatus)
    End If
    
End Sub

' Update the sheet in response to the selection of 'Use Spectral Model' dropdown
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim currentShtStatus As sheetStatus
    Dim n As Integer
    
    Call PreModify(SpectralSht, currentShtStatus)
    
    ' Hiding and unhiding necessary rows
    If Not Intersect(Target, Range("UseSpectralModel")) Is Nothing Then
        If Range("UseSpectralModel").Value = "Yes" Then
            Range("SpectralModelRng").EntireRow.Hidden = False
            Range("NoSpectralModelRng").EntireRow.Hidden = True
'            ChartObjects("HorizonChart").Visible = True
        ElseIf Range("UseSpectralModel").Value = "No" Then
            Range("SpectralModelRng").EntireRow.Hidden = True
            Range("NoSpectralModelRng").EntireRow.Hidden = False
'            ChartObjects("HorizonChart").Visible = True
        End If
    End If
    
    Call PostModify(SpectralSht, currentShtStatus)

End Sub


