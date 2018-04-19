Attribute VB_Name = "PrePostModifyModule"
' Objects of this type are used to store information            '
' about the status of a sheet: whether it is protected,         '
' whether the calculation is automatic and whether the screen   '
' updating is automatic or manual. Before modifying a sheet     '
' the relevant information is stored in a SheetStatus variable. '
' After the modification is done the sheet is restored to its   '
' original state with the help of the SheetStatus information.  '
'                                                               '
'      *** DO NOT CHANGE ANYTHING IN THIS MODULE AS IT ***      '
'      ***  IS CENTRAL TO THE PROPER OPERATION OF THE  ***      '
'      ***                 PROGRAM                     ***      '
'                                                               '
'***************************************************************'

Public Type sheetStatus
    IsScreenUpdating As Boolean
    IsCalculationAuto As Variant
    IsProtected As Boolean
    ActSht As Worksheet
End Type

' Prepare a sheet for modification:
' suspend screen updating, disable automatic calculation,
' and remove protection.
Public Sub PreModify(Sheet As Worksheet, ByRef status As sheetStatus)
    
    ' Disable screen updating so that screen does not flicker
    status.IsScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' Disable automatic calculation to speed up the process
    status.IsCalculationAuto = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    ' Keep track of active sheet and set current sheet as active
    Set status.ActSht = Application.ActiveSheet
    Sheet.Activate
   
    ' Remove protection from sheet
    status.IsProtected = Sheet.ProtectContents
    If (status.IsProtected) Then Sheet.Unprotect
End Sub

' Restore a sheet's status after modifications:
' restore screen updating, enable automatic calculation, and
' restore protection if they were there in the first place
' Note: if Protected is FALSE, protection is not restored
Public Sub PostModify(Sheet As Worksheet, ByRef status As sheetStatus)

    If (status.IsProtected) Then Sheet.Protect
    status.ActSht.Activate
    Application.Calculation = status.IsCalculationAuto
    Application.ScreenUpdating = status.IsScreenUpdating
End Sub



