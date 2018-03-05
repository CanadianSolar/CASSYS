VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inverter_DatabaseSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                   INVERTER DATABASE               '

' The inverter database contains specifications for '
' inverters, which are imported from a .PAN file.   '
' The system sheet's inverter selection userforms   '
' draw their data from this database.               '

Private Sub Worksheet_Deactivate()
    Me.Visible = xlSheetHidden
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim currentShtStatus As sheetStatus
    Dim i As Integer
    Dim numRows As Integer
    numRows = Range("A" & Rows.count).End(xlUp).row
    For i = 3 To numRows
        ' If the Multi-curve cell reads "True" or "False" instead of having an "X" in the multi-curve or single-curve cell it is changed to be read properly
        If Range("AN" & i).Value <> vbNullString Then
            Application.EnableEvents = False
            Call PreModify(Inverter_DatabaseSht, currentShtStatus)
            If Range("AN" & i).Value = False Then
                Range("AN" & i).Value = ""
                Range("AO" & i).Value = "X"
            ElseIf Range("AN" & i).Value = True Then
                Range("AN" & i).Value = "X"
            End If
            Call PostModify(Inverter_DatabaseSht, currentShtStatus)
            Application.EnableEvents = True
        End If
    Next i
End Sub
        
