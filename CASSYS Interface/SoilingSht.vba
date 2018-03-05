VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SoilingSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_Activate()
    
    Me.Activate
    Range("SfreqVal").Select

End Sub
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    
    Dim currentShtStatus As sheetStatus
    
    ' If Save is selected run Save Function
    If Target.Range.Address = Range("SaveSoiling").Address Then
        Call PreModify(SoilingSht, currentShtStatus)
        Call SaveXML
        Call PostModify(SoilingSht, currentShtStatus)
    End If
    
End Sub

' WorkSheet_Change Function
' This function is called when a cell is changed
'
' The purpose of the function is to show or hide
' the fields corresponding to the value chosen in
' the drop down list

Private Sub Worksheet_Change(ByVal Target As Range)
    
    Dim currentShtStatus As sheetStatus

    If Not Intersect(Target, Range("SfreqVal")) Is Nothing Then
    
        Call PreModify(SoilingSht, currentShtStatus)
        If Range("SfreqVal").Value = "Monthly" Then
             'Hide Yearly Losses
            SoilingSht.Rows("12:13").Hidden = True
            ' Show Monthly Losses
            SoilingSht.Rows("14:15").Hidden = False
        ElseIf Range("SfreqVal").Value = "Yearly" Then
             'Show Yearly Losses
            SoilingSht.Rows("12:13").Hidden = False
             'Hide Monthly Losses
            SoilingSht.Rows("14:15").Hidden = True
        End If
        Call PostModify(SoilingSht, currentShtStatus)
    End If
    
End Sub

