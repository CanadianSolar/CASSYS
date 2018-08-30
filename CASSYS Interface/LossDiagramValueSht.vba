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
' The algorithm is as follows:
' - find the max of the radiation and energy parts
' - adjust the axes so that the max is at 90% of full scale
' - and the effective POA radiation (axis 2) and PV nominal energy (on axis 1) align
' This translates into Geff/scale2 = Enom/scale1
' and max(Gmax/scale2, Emax/scale1) = 0.9
' which is solved for scale1 and scale2
Function AxesAlignment() As Boolean
    Dim chartObj As ChartObject
    Dim Gmax, Emax As Double
    Dim Scale1, Scale2 As Double

    ' If there was an issue in the simulation, the data may not be there in which case this function will crash
    ' If that's the case, recover gracefully
    On Error GoTo EndAxesAlignment
    
    ' Adjust axes
    Gmax = WorksheetFunction.Max(Range("LossDiagramRadiations"))
    Emax = WorksheetFunction.Max(Range("LossDiagramEnergies"))
    Scale1 = WorksheetFunction.Max(Gmax * Range("ArrayNomEnergy").Value / Range("Effective_POA_Radiation").Value, Emax) / 0.9
    Scale2 = Scale1 * Range("Effective_POA_Radiation").Value / Range("ArrayNomEnergy").Value
    Set chartObj = Sheets("Losses Diagram").ChartObjects(1)

    With chartObj
        chartObj.Chart.Axes(xlValue, xlPrimary).MaximumScale = Scale1
        chartObj.Chart.Axes(xlValue, xlSecondary).MaximumScale = Scale2
    End With
    
EndAxesAlignment:
    On Error GoTo 0
End Function


