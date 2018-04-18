Attribute VB_Name = "CheckBoxModule"
'                   CheckBoxModule              '
'-----------------------------------------------'
' Collection of all subroutines to control      '
' checkboxes on the output file sheet           '

' checkBoxActions Function
'
' Generic function that specifies actions to take after checking a checkbox on the output file page
' Uses relative ranges and offsets so that adding or removing columns from the output file sheet will not affect functionality of checkboxes
' OutputFormatModule is responsible for naming the ranges referred to in this function

Sub checkBoxActions(ByVal checkBoxName As String, ByVal sectionName As String)

    Dim shtStatus As sheetStatus 'Sheet status used for pre/post modify
    Dim Chkbox As CheckBox ' Represents a single checkbox object on the sheet
    Dim unitsColumnOffset As Integer ' The difference of columns between the column containing the units and the column with the selection boxes
    Dim paramColumnOffset ' The difference of columns between the column with the display names of the available outputs and the column of the selection boxes
    Dim sectionBlock As Range ' A range containing all of the selection boxes in a single section
    Dim aCell As Range ' represents a single cell in the range sectionBlock, used with the For Each loop
    Dim currentShtStatus As sheetStatus
    
    ' Set offsets
    paramColumnOffset = Range("OutputParam").Column - Range("HeaderRow").Column
    unitsColumnOffset = OutputFileSht.Range("UnitsColumn").Column - OutputFileSht.Range("OutputParam").Column
    
    ' Specify the range that will be changed
    Set sectionBlock = OutputFileSht.Range(OutputFileSht.Range(sectionName & "_SectionStart").Offset(0, paramColumnOffset), OutputFileSht.Range(sectionName & "_SectionEnd").Offset(-1, 0))
    
    Call PreModify(OutputFileSht, currentShtStatus)
    
    Set Chkbox = ActiveSheet.CheckBoxes(checkBoxName)  'The checkbox being checked/unchecked
    
    With Chkbox
        'Allows for change of text
        .LockedText = False
        
        If InStr(1, .text, "Summarize") <> 0 Then
            ' If the user wants to change all the values to Summarize
            For Each aCell In sectionBlock
                ' Only changes specific parameters with the checkbox
                ' Parameters with units of "°", "%", "Unitless" are not summarized; Power Injected to grid is not summarized even though it has units of kW
                If Not aCell.Offset(0, unitsColumnOffset).Value = "°" _
                And Not aCell.Offset(0, unitsColumnOffset).Value = "Unitless" _
                And Not aCell.Offset(0, unitsColumnOffset).Value = "%" _
                And Not aCell.Name.Name = "Power_Injected_into_Grid" Then
                    OutputFileSht.Cells(aCell.row, OutputFileSht.Range("OutputParam").Column).Value = "Summarize"
                End If
            Next
            .text = "Click to set the entire section to 'Detail'"
            .Value = 1
        ElseIf InStr(1, .text, "Detail") <> 0 Then
            ' If the user wants to change all the values to Detail
            sectionBlock.Value = "Detail"
            .text = "   Click to set the entire section to '-'"
            .Value = 2
        Else
            ' If the user wants to change all the values to '-'
            sectionBlock.Value = "-"
            .text = "Click to set the entire section to 'Summarize'"
            .Value = 0
        End If
        
    End With
    
    Call OutputFileSht.ChangeCellColour(OutputFileSht.Range("OutputParam"))
    Call PostModify(OutputFileSht, currentShtStatus)
    
End Sub

Sub MeteorologicalChkBox_Click()
    'Controls the Metereological and Sun Position Data section
    Call checkBoxActions("Meteorological CheckBox", "Meteorological")
End Sub
Sub TrackerChkBox_Click()
    'Controls the Tracker Or Collector Parameters section
    Call checkBoxActions("Tracker CheckBox", "Tracker")
End Sub
Sub IrradianceCollectorPlane_Click()
    'Controls the Incident Irradiance in Collector Plane section
    Call checkBoxActions("Irradiance Checkbox", "Irradiance")
End Sub

Sub ShadingCheckBox_Click()
    'Controls the Shading Section
    Call checkBoxActions("ShadingChkBox", "Shading")
End Sub

Sub IncidentEnergyChkBox_Click()
    'Controls Incident Energy Factors Section
    Call checkBoxActions("IncidentEnergy CheckBox", "Incident")
End Sub

Sub BifacialChkBox_Click()
    ' Controls the Bifacial section
    Call checkBoxActions("BifacialChkBox", "Bifacial")
End Sub

Sub PVArrayChkBox_Click()
    ' Controls the PV Array section
    Call checkBoxActions("PVArrayChkBox", "Pv")
End Sub

Sub InverterChkBox_Click()
    ' Controls the Inverter Section
    Call checkBoxActions("InverterChkBox", "Inverter")
End Sub

Sub SystemLossesPerfChkBox_Click()
    ' Controls the System Wide Losses and Performance section
    Call checkBoxActions("SystemLossesChkBox", "System")
End Sub

Sub EfficienciesChkBox_Click()
    ' Controls the Efficiencies section
    Call checkBoxActions("EfficienciesChkBox", "Efficiencies")
End Sub

'NB: Adjusted so control all only controls visible checkboxes 29/01/2016
Sub ControlAllChkBox_Click()

    ' Togggles all checkboxes when clicked to the next state
    Dim shtStatus As sheetStatus
    Dim Chkbox As CheckBox
    Dim currentShtStatus As sheetStatus
    
    Call PreModify(OutputFileSht, currentShtStatus)
    
    For Each Chkbox In OutputFileSht.CheckBoxes
        If Chkbox.Name <> "ControlAllChkBox" Then
            With Chkbox
               Chkbox.text = OutputFileSht.CheckBoxes("ControlAllChkBox").text
            End With
        End If
    Next
    
    ' In grid connected mode
    If IntroSht.Range("ModeSelect") = "Grid-Connected System" Then
        ' Change all checkboxes
        Call MeteorologicalChkBox_Click
        Call TrackerChkBox_Click
        Call IrradianceCollectorPlane_Click
        Call ShadingCheckBox_Click
        Call IncidentEnergyChkBox_Click
        Call BifacialChkBox_Click
        Call PVArrayChkBox_Click
        Call InverterChkBox_Click
        Call SystemLossesPerfChkBox_Click
        Call EfficienciesChkBox_Click
    
    ' If in RadOnlyMode
    ElseIf IntroSht.Range("ModeSelect") = "Radiation Mode" Then
        ' Change only checkboxes related to radiation only systems
        Call MeteorologicalChkBox_Click
        Call IrradianceCollectorPlane_Click
        Call TrackerChkBox_Click
    ' If ASTM system mode
    ElseIf IntroSht.Range("ModeSelect") = "ASTM E2848 Regression" Then
        Call MeteorologicalChkBox_Click
        Call IrradianceCollectorPlane_Click
        Call SystemLossesPerfChkBox_Click
    End If
        
    
    With OutputFileSht.CheckBoxes("ControlAllChkBox")
        ' Allows for change of text
        .LockedText = False
        ' Change the control all checkbox to match the text of the other checkboxes
        If InStr(1, .text, "Detail") <> 0 Then
            .text = "Click to set all sections and outputs to '-'"
            .Value = 2
        ElseIf InStr(1, .text, "-") <> 0 Then
            .text = "   Click to set all sections and outputs to 'Summarize'"
            .Value = 0
        Else
            .text = "Click to set all sections and outputs to 'Detail'"
            .Value = 1
        End If
    End With

    Call PostModify(OutputFileSht, currentShtStatus)

End Sub


