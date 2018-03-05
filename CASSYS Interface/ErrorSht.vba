VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                                     ERROR LOG                                           '
'-----------------------------------------------------------------------------------------'
' The error log collects information about critical errors encountered simulation         '
' or load attempts, as well as displaying warnings that the user should be aware of.      '
' The error log is only activated if errors were encountered after loading or simulation. '
' During loading, the errorLogger subroutine is responsible for writing messages to this  '
' sheet; during simulation, the error log created by the engine is imported using Excel's '
' text import function.                                                                   '

Option Explicit

Private Sub Worksheet_Activate()

    ' The worksheet is activated again to reset the active sheet to the current sheet
    ' to prevent .Select from producing an error message
    Me.Activate
     ' Automatically selects cell A1 of the simulation error log upon sheet activation
    Range("A1").Select
    
End Sub


