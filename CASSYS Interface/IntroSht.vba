VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IntroSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                       INTRO SHEET                           '

' The intro sheet houses links to each worksheet              '
' and the action buttons (Load, Save, Save As, Simulate, New) '



Option Explicit
Private Sub Worksheet_Activate()

    ' Resets the active sheet to the current sheet to prevent errors from .Select
    Me.Activate
    
    ' Upon sheet activation, puts focus on cell containing the file name
    Range("IntroFileName").Select
    
    
End Sub

' WorkSheet_FollowHyperlink
'
' This function is called whenever a hyperlink is
' clicked in the Intro page
'
' The purpose of this function is to Call functions
' that correspond to the hyperlink value

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    
    ' Navigates to the Module Database page when the respective link is clicked
    If Target.Name = "Module Database" Then
        PV_DatabaseSht.Visible = xlSheetVisible
        PV_DatabaseSht.Activate
        PV_DatabaseSht.Range("F" & PVDataHeight + 1).Select
    ' Navigates to the Module Database page when the respective link is clicked
    ElseIf Target.Name = "Inverter Database" Then
        Inverter_DatabaseSht.Visible = xlSheetVisible
        Inverter_DatabaseSht.Activate
        Inverter_DatabaseSht.Range("F" & InvDataHeight + 1).Select
    End If
    
End Sub

' NB: Causes mode select to change the visible sheets
' Activated when a cell is changed
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim currentShtStatus As sheetStatus
    Application.EnableEvents = False
    Call PreModify(IntroSht, currentShtStatus)
    
    ' Goes to the mode-changing sub
    If Not Intersect(Target, Range("ModeSelect")) Is Nothing Then Call SwitchMode(Range("ModeSelect").Value)

    
    Call PostModify(IntroSht, currentShtStatus)
    Application.EnableEvents = True
End Sub

Sub SwitchMode(ByVal Selection As String)
    Dim SiteShtStatus As sheetStatus
    Dim InputFileShtStatus As sheetStatus
    Dim OutputShtStatus As sheetStatus
    Dim currentShtStatus As sheetStatus
    Dim Orientation_and_ShadingShtStatus As sheetStatus
    
'--------Commenting out Iterative Functionality for this version--------'
'    ' Creating list of outputs for Iteration Mode output
'    Call IterativeOutputValidation

    ' If CASSYS is in "Radiation Mode" all sheets which define the system disappear as well as AstmSht
    ' as do their hyperlinks on the intro page
    If (Selection = "Radiation Mode") Then
        
        Call PreModify(IntroSht, currentShtStatus)
        SystemSht.Visible = xlSheetHidden
        LossesSht.Visible = xlSheetHidden
        SoilingSht.Visible = xlSheetHidden
        SpectralSht.Visible = xlSheetHidden
        TransformerSht.Visible = xlSheetHidden
        Orientation_and_ShadingSht.Visible = xlSheetVisible
        Horizon_ShadingSht.Visible = xlSheetVisible
        AstmSht.Visible = xlSheetHidden
        
        IntroSht.Range("ASTM_only").EntireRow.Hidden = True
        IntroSht.Range("ASTM_spaces").EntireRow.Hidden = True
        IntroSht.Range("ASTM_hyperlink_hide").EntireRow.Hidden = False
        IntroSht.Range("Grid_Only").EntireRow.Hidden = True
        IntroSht.Range("Rad_Only_Empty").EntireRow.Hidden = False
        Call PostModify(IntroSht, currentShtStatus)
        
        'unlocked and unshading columns that were not used in ASTM system
        Call PreModify(InputFileSht, InputFileShtStatus)
        InputFileSht.Range("ASTM_locked").Locked = False
        InputFileSht.Range("ASTM_locked").Interior.Color = RGB(255, 255, 255)
        Call PostModify(InputFileSht, InputFileShtStatus)
        
        Call PreModify(OutputFileSht, OutputShtStatus)
        OutputFileSht.Range("OutputSht_ASTM_hide").EntireRow.Hidden = False
        OutputFileSht.Range("GridConnectedOutputs").EntireRow.Hidden = True
        OutputFileSht.CheckBoxes("PVArrayChkBox").Visible = False
        OutputFileSht.CheckBoxes("InverterChkBox").Visible = False
        OutputFileSht.CheckBoxes("SystemLossesChkBox").Visible = False
        OutputFileSht.CheckBoxes("EfficienciesChkBox").Visible = False
        OutputFileSht.CheckBoxes("IncidentEnergy CheckBox").Visible = False
        OutputFileSht.CheckBoxes("ShadingChkBox").Visible = False
        Call PostModify(OutputFileSht, OutputShtStatus)
        
        'unhiding site file parameters not needed for ASTM mode
        Call PreModify(SiteSht, SiteShtStatus)
        SiteSht.Range("Site_ASTM_hide").EntireRow.Hidden = False
        Call PostModify(SiteSht, SiteShtStatus)
        
        Call PreModify(Orientation_and_ShadingSht, Orientation_and_ShadingShtStatus)
        With Orientation_and_ShadingSht.Range("OrientType").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=OrientRadOnly"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .errorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
        
        Orientation_and_ShadingSht.Range("RowsBlockSAET").Value = 1
        Orientation_and_ShadingSht.Range("RowsBlockSAST").Value = 1
        Orientation_and_ShadingSht.Range("RowsBlockSAET").Interior.Color = RGB(204, 192, 218)
        Orientation_and_ShadingSht.Range("RowsBlockSAST").Interior.Color = RGB(204, 192, 218)
        Orientation_and_ShadingSht.Range("RowsBlockSAET").Locked = True
        Orientation_and_ShadingSht.Range("RowsBlockSAST").Locked = True
        
        ' Set orientation type to Fixed Titled Plane by default
        Application.EnableEvents = True
        Orientation_and_ShadingSht.Range("OrientType").Value = Orientation_and_ShadingSht.Range("OrientRadOnly").Cells(1, 1).Value
        Application.EnableEvents = False
        
        Call PostModify(Orientation_and_ShadingSht, Orientation_and_ShadingShtStatus)
    
    ' If CASSYS is in "ASTM E2848 Regression" all sheets except AstmSht disappear
    ' as do their hyperlinks on the intro page
    ElseIf (Selection = "ASTM E2848 Regression") Then
        'hiding unnecessary sheets
        Call PreModify(IntroSht, currentShtStatus)
        AstmSht.Visible = xlSheetVisible
        Orientation_and_ShadingSht.Visible = xlSheetHidden
        Horizon_ShadingSht.Visible = xlSheetHidden
        SystemSht.Visible = xlSheetHidden
        LossesSht.Visible = xlSheetHidden
        SoilingSht.Visible = xlSheetHidden
        SpectralSht.Visible = xlSheetHidden
        TransformerSht.Visible = xlSheetHidden
        ResultSht.Visible = xlSheetHidden
        ReportSht.Visible = xlSheetHidden
        'hiding unnecessary hyperlinks & spacing
        IntroSht.Range("ASTM_only").EntireRow.Hidden = False
        IntroSht.Range("ASTM_spaces").EntireRow.Hidden = False
        IntroSht.Range("ASTM_hyperlink_hide").EntireRow.Hidden = True
        IntroSht.Range("Rad_Only_Empty").EntireRow.Hidden = True
        Call PostModify(IntroSht, currentShtStatus)
        
        'locking and shading input to parameters not required
        Call PreModify(InputFileSht, InputFileShtStatus)
        Application.EnableEvents = True
        InputFileSht.Range("ASTM_locked").Value = vbNullString
        Application.EnableEvents = False
        InputFileSht.Range("ASTM_locked").Locked = True
        InputFileSht.Range("ASTM_locked").Interior.Color = RGB(204, 192, 218)
        Call PostModify(InputFileSht, InputFileShtStatus)
        
        'hiding unncessary output parameters
        Call PreModify(OutputFileSht, OutputShtStatus)
        OutputFileSht.Range("GridConnectedOutputs").EntireRow.Hidden = False
        OutputFileSht.CheckBoxes("SystemLossesChkBox").Visible = True
        OutputFileSht.CheckBoxes("Tracker Checkbox").Visible = False
        OutputFileSht.CheckBoxes("PVArrayChkBox").Visible = False
        OutputFileSht.CheckBoxes("InverterChkBox").Visible = False
        OutputFileSht.CheckBoxes("EfficienciesChkBox").Visible = False
        OutputFileSht.CheckBoxes("IncidentEnergy CheckBox").Visible = False
        OutputFileSht.CheckBoxes("ShadingChkBox").Visible = False
        OutputFileSht.Range("OutputSht_ASTM_hide").EntireRow.Hidden = True
        Call PostModify(OutputFileSht, OutputShtStatus)
        
        'hiding site file parameters not needed for ASTM mode
        Call PreModify(SiteSht, SiteShtStatus)
        SiteSht.Range("Site_ASTM_hide").EntireRow.Hidden = True
        Call PostModify(SiteSht, SiteShtStatus)
        
    ' CASSYS is the same as before when in "Grid-Connected System"
    ' Astm sheet is hidden
    Else
        Call PreModify(IntroSht, currentShtStatus)
        SystemSht.Visible = xlSheetVisible
        LossesSht.Visible = xlSheetVisible
        SoilingSht.Visible = xlSheetVisible
        SpectralSht.Visible = xlSheetVisible
        TransformerSht.Visible = xlSheetVisible
        Horizon_ShadingSht.Visible = xlSheetVisible
        Orientation_and_ShadingSht.Visible = xlSheetVisible
        AstmSht.Visible = xlSheetHidden
        
        IntroSht.Range("ASTM_only").EntireRow.Hidden = True
        IntroSht.Range("ASTM_spaces").EntireRow.Hidden = True
        IntroSht.Range("ASTM_hyperlink_hide").EntireRow.Hidden = False
        IntroSht.Range("Grid_Only").EntireRow.Hidden = False
        IntroSht.Range("Rad_Only_Empty").EntireRow.Hidden = True
        Call PostModify(IntroSht, currentShtStatus)
        
        'unlocked and unshading columns that were not used in ASTM system
        Call PreModify(InputFileSht, InputFileShtStatus)
        InputFileSht.Range("ASTM_locked").Locked = False
        InputFileSht.Range("ASTM_locked").Interior.Color = RGB(255, 255, 255)
        Call PostModify(InputFileSht, InputFileShtStatus)
        
        Call PreModify(OutputFileSht, OutputShtStatus)
        OutputFileSht.Range("OutputSht_ASTM_hide").EntireRow.Hidden = False
        OutputFileSht.Range("GridConnectedOutputs").EntireRow.Hidden = False
        OutputFileSht.CheckBoxes("Tracker Checkbox").Visible = True
        OutputFileSht.CheckBoxes("PVArrayChkBox").Visible = True
        OutputFileSht.CheckBoxes("InverterChkBox").Visible = True
        OutputFileSht.CheckBoxes("SystemLossesChkBox").Visible = True
        OutputFileSht.CheckBoxes("EfficienciesChkBox").Visible = True
        OutputFileSht.CheckBoxes("IncidentEnergy CheckBox").Visible = True
        OutputFileSht.CheckBoxes("ShadingChkBox").Visible = True
        Call PostModify(OutputFileSht, OutputShtStatus)
        
        'unhiding site file parameters not needed for ASTM mode
        Call PreModify(SiteSht, SiteShtStatus)
        SiteSht.Range("Site_ASTM_hide").EntireRow.Hidden = False
        Call PostModify(SiteSht, SiteShtStatus)
        
        Call PreModify(Orientation_and_ShadingSht, Orientation_and_ShadingShtStatus)
        With Orientation_and_ShadingSht.Range("OrientType").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=OrientList"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .errorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
        
        Orientation_and_ShadingSht.Range("RowsBlockSAET").Interior.Color = RGB(255, 255, 255)
        Orientation_and_ShadingSht.Range("RowsBlockSAST").Interior.Color = RGB(255, 255, 255)
        Orientation_and_ShadingSht.Range("RowsBlockSAET").Locked = False
        Orientation_and_ShadingSht.Range("RowsBlockSAST").Locked = False
        
        Call PostModify(Orientation_and_ShadingSht, Orientation_and_ShadingShtStatus)
    End If
    
End Sub

'--------Commenting out Iterative Functionality for this version--------'

Sub IterativeOutputValidation()
'    Dim c As Range
'    Dim i As Integer
'
'    ' unprotect sheet to allow cell validation
'    IterativeSht.Unprotect
'
'    ' Creating list of available parameters for iteration sheet
'
'    IterativeSht.Range("W:W").ClearContents
'    i = 1
'    ' Writing all mode available outputs to column W in Iteration mode sheet
'    If IntroSht.Range("ModeSelect").Value = "Grid-Connected System" Then
'        For Each c In OutputFileSht.Range("Grid_Con_OutputNames")
'            ' Only output parameters with a summary option are allowed
'            If (OutputFileSht.Range("E" & c.row).Validation.Formula1 = "=SummaryOption") Then
'                i = i + 1
'                IterativeSht.Range("W" & CStr(i)).Value = c.Value
'            End If
'        Next
'    ElseIf IntroSht.Range("ModeSelect").Value = "Radiation Mode" Then
'        For Each c In OutputFileSht.Range("RadiationOutputNames")
'            ' Only output parameters with a summary option are allowed
'            If (OutputFileSht.Range("E" & c.row).Validation.Formula1 = "=SummaryOption") Then
'                i = i + 1
'                IterativeSht.Range("W" & CStr(i)).Value = c.Value
'            End If
'        Next
'    ElseIf IntroSht.Range("ModeSelect").Value = "ASTM E2848 Regression" Then
'        For Each c In OutputFileSht.Range("ASTMOutputNames")
'            ' Only output parameters with a summary option are allowed
'            If (OutputFileSht.Range("E" & c.row).Validation.Formula1 = "=SummaryOption") Then
'                i = i + 1
'                IterativeSht.Range("W" & CStr(i)).Value = c.Value
'            End If
'        Next
'    End If
'
'
'
'    ' Creating validation
'    With IterativeSht.Range("IterativeOutputParam").Validation
'    .Delete
'    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'    xlBetween, Formula1:="=$W$2:$W$" & i
'    .IgnoreBlank = True
'    .InCellDropdown = True
'    .InputTitle = ""
'    .ErrorTitle = ""
'    .InputMessage = ""
'    .errorMessage = ""
'    .ShowInput = True
'    .ShowError = True
'    End With
'
'    ' Protecting sheet after setting cell validation
'    IterativeSht.Protect
'
'    ' Set output to first element in list
'    ' This ensures the output selected is available for the selected mode
'    IterativeSht.Range("IterativeOutputParam").Value = IterativeSht.Range("W2").Value

End Sub


