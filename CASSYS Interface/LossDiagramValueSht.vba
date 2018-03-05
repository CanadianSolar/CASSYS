VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LossDiagramValueSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------LossDiagramValueSht------------------'
'The purpose of this worksheet is to hold the values and'
'calculations necessary to display the loss diagram     '

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim currentShtStatus As sheetStatus
    Dim charObj As ChartObject

    If Target = LossDiagramValueSht.Range("HorizontalGlobIrradiance") Then
        Call PreModify(LossDiagramSht, currentShtStatus)
        Call AxesAlignment
        Call PostModify(LossDiagramSht, currentShtStatus)
    End If
End Sub

'The purpose of this sub is to align the primary
'and secondary axes in the losses diagram when a
'value is changed in the outlined table
'
Public Sub AxesAlignment()
    Dim chartObj As ChartObject

    Set chartObj = Sheets("Losses Diagram").ChartObjects(1)

    With chartObj
        chartObj.Chart.Axes(xlValue, xlPrimary).MaximumScale = Range("ArrayNomEnergy").Value + (Range("ArrayNomEnergy").Value * 0.25)
        chartObj.Chart.Axes(xlValue, xlSecondary).MaximumScale = Range("Effective_POA_Radiation").Value + (Range("Effective_POA_Radiation").Value * 0.25)
    End With

End Sub



