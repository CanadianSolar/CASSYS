VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransformerSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'           TRANSFORMER SHEET            '
'----------------------------------------'
' The function of the transformer sheet  '
' is to collect information about losses '
' and power/voltage relating to the      '
' transformer.                           '

Option Explicit
Private Sub Worksheet_Activate()
    
    ' Resets the active sheet to this worksheet
    Me.Activate
    
    'Load and caryy out PVSyst Equivalent calculations
    calcPVSystEquivalents
    
    ' Automatically selects the first editable cell on this worksheet
    Range("PIronLossTrf").Select
    
End Sub

'Update PVSyst equivalent calculations when transformer sheet is selected
Private Sub calcPVSystEquivalents()

   CASSYStoPVSystIronLoss
   CASSYStoPVSystResistiveLosses
   
End Sub

' Worksheet_FollowHyperLink
'
' This sub is activated when a link is clicked on the transformer sheet
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    
    Dim currentShtStatus As sheetStatus
    Call PreModify(TransformerSht, currentShtStatus)
    ' If Save is selected run the Save function
    If Target.Range.Address = Range("SaveTransformer").Address Then
        Call SaveXML
    ElseIf Target.Range.Address = Range("ShowHidePV").Address Then
        Call togglecols
    End If
    Call PostModify(TransformerSht, currentShtStatus)
    
End Sub

' Hides or shows the PVSyst equivalent values on the transformer sheet
Function togglecols() As Boolean
    
    Range("PVSystVals").EntireRow.Hidden = Not Range("PVSystVals").EntireRow.Hidden
    If Range("PVSystVals").EntireRow.Hidden = True Then
        Range("ShowHidePV").Value = "Show PVSyst Equivalents"
    Else
        Range("ShowHidePV").Value = "Hide PVSyst Equivalents"
    End If

End Function
'
'Called when a cell is changed.
'More specifically this change function will allow
'the user to put in either CASSYS or PVSyst values
'and will change the corresponding CASSYS or PVSyst
'values.
'
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim currentShtStatus As sheetStatus
    
    Application.EnableEvents = False
    Call PreModify(TransformerSht, currentShtStatus)
    
    'CASSYS values --> therfore converting them to PVSyst Values
    If Not Intersect(Target, Range("PIronLossTrf")) Is Nothing And IsNumeric(Range("PIronLossTrf")) = True Then
        CASSYStoPVSystIronLoss
        CASSYStoPVSystResistiveLosses
        
    ElseIf Not Intersect(Target, Range("PFullLoadLss")) Is Nothing And IsNumeric(Range("PFullLoadLss")) = True Then
        CASSYStoPVSystResistiveLosses
        
    End If
    
    'PVSyst inputs --> therfore converting them to CASSYS values
    If Not Intersect(Target, Range("PIronLoss")) Is Nothing And IsNumeric(Range("PIronLoss")) Then
        PVSystIronLossPercentCalculation
        PVSystToCASSYSIronLoss
        
    ElseIf Not Intersect(Target, Range("FIronLoss")) Is Nothing And IsNumeric(Range("FIronLoss")) Then
        PVSystIronLossWattCalculation
        PVSystToCASSYSIronLoss
        
    ElseIf Not Intersect(Target, Range("FResLoss")) Is Nothing And IsNumeric(Range("FResLoss")) Then
        PVSystToCASSYSFullLoadLss
        
    End If
    
    'Update PVSyst Resistive/Inductive Losses when any variable is changed
    CASSYStoPVSystResistiveLosses
    
    'Warning if nominal power input for transformer is less than the nominal inverter ac power from systems sheet
    If Not Intersect(Target, Range("PNomTrf")) Is Nothing And Range("PNomTrf").Value < Worksheets("System").Range("SystemAC").Value Then
        Range("TrfWarning").Value = "Warning: Nominal power of transformer is less than that of the inverters"
        Range("TrfWarning").Font.Color = RGB(255, 0, 0)
    Else
         Range("TrfWarning").ClearContents
    End If
    
    Call PostModify(TransformerSht, currentShtStatus)
    Application.EnableEvents = True
End Sub

'Convert from PVSyst iron loss input to CASSYS iron loss input
Private Sub PVSystToCASSYSIronLoss()
    Range("PIronLossTrf").Value = Range("PIronLoss").Value
End Sub

'Convert from PVSyst input losses to CASSYS full load loss
Private Sub PVSystToCASSYSFullLoadLss()
    If (Range("ACCapSTC").Value = 0) Then
        Range("PFullLoadLss").Value = 0
    Else
        Range("PFullLoadLss").Value = ((((Range("FResLoss").Value) * ((Range("PNomTrf").Value * 1000) ^ 2)) / (Range("ACCapSTC").Value * 1000)) + (Range("PIronLoss").Value * 1000)) / 1000
    End If
End Sub

'Convert CASSYS Iron Loss to PVSyst Iron Loss equivalent
Private Sub CASSYStoPVSystIronLoss()
    Range("PIronLoss").Value = Range("PIronLossTrf").Value
    PVSystIronLossPercentCalculation
End Sub

'Convert CASSYS inputs to PVSyst Resistive/inductive losses percentage
Private Sub CASSYStoPVSystResistiveLosses()

    If (Range("PNomTrf").Value = 0) Then
        Range("FResLoss").Value = 0
    Else
        Range("FResLoss").Value = ((Range("PFullLoadLss").Value - Range("PIronLossTrf")) * 1000) * (Range("ACCapSTC").Value * 1000 / ((Range("PNomTrf").Value * 1000) ^ 2))
    End If
    
End Sub

'Convert PVSyst Iron loss in kW to percentage
Private Sub PVSystIronLossPercentCalculation()
    If (Range("ACCapSTC").Value = 0) Then
        Range("FIronLoss").Value = 0
    Else
        Range("FIronLoss").Value = Range("PIronLossTrf").Value / Range("ACCapSTC").Value
    End If
End Sub

'Convert PVSyst Iron loss in percentage to kW
Private Sub PVSystIronLossWattCalculation()
    Range("PIronLoss").Value = Range("FIronLoss").Value * Range("ACCapSTC").Value
End Sub
