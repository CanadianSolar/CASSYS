VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'                   SYSTEM SHEET                '
'-----------------------------------------------'
' The system sheet is where the PV-Arrays and   '
' inverters used for simulation are defined.    '
' This sheet draws its data directly from the   '
' Inverter and PV module databases.


' NB: The values in cells SystemAC and SystemDC '
' are dictated by the value of cell "SelectMode"'
' in the intro sheet. If the mode is set to     '
' "Radiation Mode" the values are set to zero,  '
' otherwise the values follow the same formula  '
' as before. This is controlled by an if-       '
' statement in the cell. If there is a better   '
' way of changing this for the save function    '
' it should be implemented.                     '


Option Explicit
Private Sub Worksheet_Activate()

    ' Resets the active sheet to this sheet
    Me.Activate
    ' Upon activating the worksheet, the first editable cell is selected
    Range("NumSubArray").Select

End Sub

' WorkSheet_Change Function
' This function is called when a cell is changed but
' will only do anything when the Range("NumSubArray")
' is changed
'

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim currentShtStatus As sheetStatus
    Dim IsEnableEvents As Boolean
    Dim i As Integer
    Dim subArrays As Integer
    Dim numRows As Integer
    
    'Only trigger event when the NumSubArray cell has been changed
    If Intersect(Target, Range("NumSubArray")) Is Nothing Then
        Exit Sub
    End If
    
    IsEnableEvents = Application.EnableEvents
    'Disable events to prevent an infinte recursive loop
    Application.EnableEvents = False
    
    Call PreModify(SystemSht, currentShtStatus)
    
    Call UpdateNumArrays(Target)
    
    Call PostModify(SystemSht, currentShtStatus)
    
    ' Restore events to what they were before
    Application.EnableEvents = IsEnableEvents
    
    
End Sub

' WorkSheet_FollowHyperlink Function
' This function is called when a hyperlink
' is clicked on
'
' The purpose of this function is to open up
' the userform that allows specific modules
' and inverters of a specific manufacturer to
' be selected

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    Dim i As Integer ' Counter variable
    Dim rowOffset As Integer
    Dim currentShtStatus As sheetStatus
  
    Call PreModify(SystemSht, currentShtStatus)
    
    For i = 1 To Range("NumSubArray")
        rowOffset = (i - 1) * SubArrayHeight
        
        ' If that search link from the PV side is clicked
        If Target.SubAddress = SystemSht.Range("PVSearch").Offset(rowOffset, 0).Address Then
            ' Store the sub-array value and open the form
            Range("PVModuleIndex").Value = i
            Call OpenPVForm
        ElseIf Target.SubAddress = SystemSht.Range("InvSearch").Offset(rowOffset, 0).Address Then
         ' If that search link from the Inverter side is clicked
            ' Store the sub-array value and open the form
            Range("InverterIndex").Value = i
            Call OpenInvForm
        End If
    Next i
    
    If Target.Range.Address = Range("SaveSystem").Address Then
        Call SaveXML
    End If
    
    Call PostModify(SystemSht, currentShtStatus)
    
End Sub
' UpdateNumArrays subroutine
'
' This routine expands or truncates the worksheet space
' based on how many sub-arrays the user has specified in
' the 'Number of Sub Arrays' field.
' It does this by simply copying the formulas and cells
' of the first sub array data section.
Private Sub UpdateNumArrays(ByRef Target As Range)

    Dim i As Integer ' Counter
    Dim j As Integer ' Counter
    Dim rowOffset  ' auxiliary offset for location of sub-array section
    Dim Hlink As Hyperlink
    Dim newNumSubArray As Integer
    Dim oldNumSubArray As Integer
 
    'Error Catching
    'Prevent Values from being less than 1
    If Range("NumSubArray").Value < 1 Then
        Range("NumSubArray").Value = 1
        Range("NumSubArray").Value = CInt(Range("NumSubArray").Value)
    ElseIf IsNumeric(Range("NumSubArray")) = False Then
        'If the value inputted is not a number return it to its original value
        Range("NumSubArray").Value = Range("TempSubArray").Value
    Else
        'Make sure the number is an integer
        Range("NumSubArray").Value = CInt(Range("NumSubArray").Value)
    End If
    
    'Range("TempSubArray") contains the number of sub-array before the NumSubArray cell is changed
    'First case - Add in sub-arrays if the new value is greater than the old one
    oldNumSubArray = Range("TempSubArray").Value
    newNumSubArray = Range("NumSubArray").Value
    If (newNumSubArray > oldNumSubArray) Then
        'Unhide rows to fit new sub-array
        Range(Cells(13 + oldNumSubArray * SubArrayHeight, 1), Cells(12 + newNumSubArray * SubArrayHeight, 1)).EntireRow.Hidden = False
        ' Copy the first array
        SystemSht.Range("SystemArray").Copy
        ' Paste the new arrays
        For i = oldNumSubArray + 1 To newNumSubArray
            ' Paste it as the new array
            rowOffset = (i - 1) * SubArrayHeight
            Cells(13 + (i - 1) * SubArrayHeight, 1).Select
            SystemSht.Paste
            SystemSht.Range("SubTitle").Offset(rowOffset, 0).Value = "SUB-ARRAY " & i
        Next i
        Application.CutCopyMode = False
        Range("NumSubArray").Select
    'Second case - Remove sub-arrays if the new vales is less than the old one
    ElseIf (newNumSubArray < oldNumSubArray) Then
        'Delete sub-array cells
        Range(Cells(13 + newNumSubArray * SubArrayHeight, 1), Cells(12 + oldNumSubArray * SubArrayHeight, 1)).EntireRow.Clear
        'Hide corresponding rows
        Range(Cells(13 + newNumSubArray * SubArrayHeight, 1), Cells(12 + oldNumSubArray * SubArrayHeight, 1)).EntireRow.Hidden = True
    End If
    
    'Update the number of arrays
    Range("TempSubArray").Value = Range("NumSubArray").Value
    
    ' Make sure that the number of strings, inverters and modules in a string are integer numbers
    For i = 1 To Range("NumSubArray")
        rowOffset = (i - 1) * SubArrayHeight
        If Not Intersect(Target, Range("ModStr").Offset(rowOffset, 0)) Is Nothing Then
            If IsNumeric(Range("ModStr").Offset(rowOffset, 0)) = False Then
                Range("ModStr").Offset(rowOffset, 0).Value = 1
            Else
                Range("ModStr").Offset(rowOffset, 0).Value = CInt(Range("ModStr").Offset(rowOffset, 0).Value)
                If Range("ModStr").Offset(rowOffset, 0).Value < 1 Then
                    Range("ModStr").Offset(rowOffset, 0).Value = 1
                End If
            End If
        ElseIf Not Intersect(Target, Range("NumStr").Offset(rowOffset, 0)) Is Nothing Then
            If IsNumeric(Range("NumStr").Offset(rowOffset, 0)) = False Then
                Range("NumStr").Offset(rowOffset, 0).Value = 1
            Else
                Range("NumStr").Offset(rowOffset, 0).Value = CInt(Range("NumStr").Offset(rowOffset, 0).Value)
                If Range("NumStr").Offset(rowOffset, 0).Value < 1 Then
                    Range("NumStr").Offset(rowOffset, 0).Value = 1
                End If
            End If
        ElseIf Not Intersect(Target, Range("NumInv").Offset(rowOffset, 0)) Is Nothing Then
            If IsNumeric(Range("NumInv").Offset(rowOffset, 0)) = False Then
                Range("NumInv").Offset(rowOffset, 0).Value = 1
            Else
                Range("NumInv").Offset(rowOffset, 0).Value = CInt(Range("NumInv").Offset(rowOffset, 0).Value)
                If Range("NumInv").Offset(rowOffset, 0).Value < 1 Then
                    Range("NumInv").Offset(rowOffset, 0).Value = 1
                End If
            End If
        End If
    Next i
    
    ' Update the hyperlinks
    For i = 1 To SystemSht.Range("NumSubArray").Value
        rowOffset = (i - 1) * SubArrayHeight
        ' update the pv module hyperlinks
        SystemSht.Hyperlinks.Add SystemSht.Range("PVSearch").Offset(rowOffset, 0), vbNullString, SystemSht.Range("PVSearch").Offset(rowOffset, 0).Address, "Search For a Specific PV Module"
        
        SystemSht.Range("PVSearch").Offset(rowOffset, 0).Font.Underline = xlUnderlineStyleNone
        SystemSht.Range("PVSearch").Offset(rowOffset, 0).Font.Bold = True
        With SystemSht.Range("PVSearch").Offset(rowOffset, 0).Font
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
        End With
        
        ' update the inverter hyperlinks
        SystemSht.Hyperlinks.Add SystemSht.Range("InvSearch").Offset(rowOffset, 0), vbNullString, SystemSht.Range("InvSearch").Offset(rowOffset, 0).Address, "Search For a Specific Inverter"
        
        SystemSht.Range("InvSearch").Offset(rowOffset, 0).Font.Underline = xlUnderlineStyleNone
        SystemSht.Range("InvSearch").Offset(rowOffset, 0).Font.Bold = True
        With SystemSht.Range("InvSearch").Offset(rowOffset, 0).Font
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
        End With
    Next i
    
End Sub



