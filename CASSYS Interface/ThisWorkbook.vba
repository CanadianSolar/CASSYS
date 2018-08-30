VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'               This Workbook               '
'-------------------------------------------'
' The routines contained on this page are   '
' triggered when the workbook is opened or  '
' closed.                                   '

Option Explicit
' Workbook_BeforeClose
'
' Event is triggered when the user clicks 'X' to close
' the Excel window
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    Dim doNotSave As Integer
    Dim currSht As Worksheet
    
    ' NB: Stops pop up asking user to save
    BypassBeforeSave = True
    Application.EnableEvents = True
    
    ' To prevent errors when trying to save
    On Error Resume Next
    Application.ScreenUpdating = False
    
    ' Check if the user wants to save the existing CSYX file before exiting the Excel fie (Do not save by default to allow loading first time when workbook opens)
    doNotSave = vbNo
    If Not IntroSht.Range("SaveFilePath").Value = vbNullString Then
    doNotSave = MsgBox("You will now exit the CASSYS Interface. Would you like to save the Site file you were working on?", vbYesNo + vbQuestion, "CASSYS: Save the Site file?")
    End If
    If doNotSave = vbYes Then
    Call SaveXML
    End If
    
    ' Shows the enable macros sheet
    Application.ScreenUpdating = False
    
    Application.DisplayAlerts = False
    ActiveWorkbook.Saved = True
            
End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    Application.DisplayAlerts = False

    ' If user is exiting CASSYS, BeforeSave bypassed
    ' so not asked if they want to save
    If Not BypassBeforeSave Then
        Dim answer As Integer
        answer = MsgBox("The CASSYS Interface (cassys.xlsm) cannot be saved. To save a project, please use the Save or Save As buttons in the Intro tab.", vbOKOnly, "CASSYS: Save")
        Cancel = True
        Application.DisplayAlerts = True
    End If

End Sub

' Workbook_Open
'
' This event is triggered right when the workbook is opened
Private Sub Workbook_Open()
    Call WorkbookOpen
End Sub

' This function contains what goes into Workbook_Open but is accessible from other parts of the program
Function WorkbookOpen() As Boolean
    
    On Error Resume Next
    ' If macros are not enabled, the "Enable macros" sheet will show and stay.
    ' Otherwise, this code will execute and show the Intro sheet
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Make all the sheets visible
    IntroSht.Visible = xlSheetVisible
    EnableMacrosSht.Visible = xlSheetVeryHidden
    SiteSht.Visible = xlSheetVisible
    Orientation_and_ShadingSht.Visible = xlSheetVisible
    BifacialSht.Visible = xlSheetVisible
    Horizon_ShadingSht.Visible = xlSheetVisible
    SystemSht.Visible = xlSheetVisible
    LossesSht.Visible = xlSheetVisible
    SoilingSht.Visible = xlSheetVisible
    SpectralSht.Visible = xlSheetVisible
    TransformerSht.Visible = xlSheetVisible
    InputFileSht.Visible = xlSheetVisible
    OutputFileSht.Visible = xlSheetVisible
    ResultSht.Visible = xlSheetVisible
    SummarySht.Visible = xlSheetVisible
    
    ' Ensuring all unneccessary sheets are hidden
    ErrorSht.Visible = xlSheetHidden
    ResultSht.Visible = xlSheetHidden
    SummarySht.Visible = xlSheetHidden
    ReportSht.Visible = xlSheetHidden
    ChartConfigSht.Visible = xlSheetHidden
   Inverter_DatabaseSht.Visible = xlSheetHidden
    PV_DatabaseSht.Visible = xlSheetHidden
    CompChart1.Visible = xlSheetHidden
    CompChart2.Visible = xlSheetHidden
    CompChart3.Visible = xlSheetHidden
    IterativeSht.Visible = xlSheetHidden
    LossDiagramSht.Visible = xlSheetHidden
    LossDiagramValueSht.Visible = xlSheetHidden
    
    
    ' Version number for current development of CASSYS (this is for extra protection, in case it gets overwritten)
    IntroSht.Unprotect
    'NB: increased version number
    IntroSht.Range("Version").Value = "1.5.2"
    IntroSht.Protect
    IntroSht.Activate
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ThisWorkbook.Saved = True
    Application.DisplayAlerts = False
    'ThisWorkbook.ChangeFileAccess Mode:=xlReadOnly
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
 
End Function
