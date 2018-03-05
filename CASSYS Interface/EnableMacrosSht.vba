VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnableMacrosSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                   ENABLE MACROS SHEET                   '
'---------------------------------------------------------'
' This worksheet is used to safeguard against situations  '
' where the user's instance of Excel has disabled macros. '
' This sheet displays a message that tells the user to    '
' enable macros. This sheet is hidden by macros everytime '
' the workbook opens, and is unhidden everytime CASSYS    '
' closes. Consequently, when macros are disabled this     '
' sheet will not be hidden upon the workbook open and     '
' the user will see this worksheet and cannot access      '
' other worksheets.                                       '
