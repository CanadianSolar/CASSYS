VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LossDiagramSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------LossDiagramSht------------------'
'The purpose of this worksheet is to give the user '
'a visual representation of the losses that have   '
'occured in the system and the resultant energies  '
'available.                                        '

Private Sub Worksheet_Activate()
    Dim currentShtStatus As sheetStatus
    
    ' Resets the active worksheet to the current sheet
    Me.Activate
    Call PreModify(LossDiagramSht, currentShtStatus)
    Call LossDiagramValueSht.AxesAlignment
    Call PostModify(LossDiagramSht, currentShtStatus)
End Sub


