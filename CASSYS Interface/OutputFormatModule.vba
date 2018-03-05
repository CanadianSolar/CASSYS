Attribute VB_Name = "OutputFormatModule"
' FormatOutputSheet Function
'
' This purpose of this subroutine is to eliminate the time consuming and unwieldy process of
' reformatting the output sheet and altering the code of various modules everytime the
' output sheet is changed (adding/reordering/removing sections or available outputs)
'
Sub FormatOutputSheet()

    Dim FooterRow As Range
    Dim currentShtStatus As sheetStatus
    Dim FoundFooter As Boolean
    
    Application.EnableEvents = False
    Call PreModify(OutputFileSht, currentShtStatus)
    
    ' Finds footer row assumming 'Version' is in the footer row
    
    For Each outParam In OutputFileSht.Range("A1:Z500")
        If InStr(1, outParam.Value, "Version") Then
            Set FooterRow = outParam
            FoundFooter = True
        End If
    Next
    
    ' If the footer or header was not found (due to some reformmating that altered the header and footer identifiers,
    ' alerts the formatter to change the vba code.
    
    If FoundFooter = False Then
        MsgBox "Footer not found. Please Alter VBA OutputFormatModule with footer identifier"
    End If
    
    Call DeleteExistingNames
    Call RenameKeyFeatures(FooterRow)
    Call NameCellsUsingOutputs
    Call NameAndFormatSections(FooterRow)
    Call PostModify(OutputFileSht, currentShtStatus)
    
    Application.EnableEvents = True
    
End Sub

' DeleteExistingNames function
'
' The purpose of this subroutine is to delete all existing names on the output sheet
' To ensure that name conflicts do not happen and to remove unnecessary past names
' in the case that an output is deleted

Sub DeleteExistingNames()

    Dim aName As Name

    For Each aName In Application.ThisWorkbook.Names
        If InStr(1, aName.RefersTo, "OutputFile!") <> 0 Then
                aName.Delete
        End If
    Next aName

End Sub


' RenameKeyFeatures function
'
' The purpose of this subroutine is to rename some important cells after the
' DeleteExistingNames function is called.

Sub RenameKeyFeatures(ByRef FooterRow As Range)
    
    Dim outParam As Range
    Dim firstOutParam As Range
    Dim lastOutParam As Range
    Dim foundfirstOutParam As Boolean
    Dim cellAddress As Variant
    Dim cellAddresses() As String
    Dim finalOutputRange As Range

    'Output Param is the column range of the first 'Yes'/'No' box to the last 'Yes'/'No' box
    
    For Each outParam In OutputFileSht.Range("A1:Z500")
        'Find first and last 'yes'/'no' boxes
        If Not Intersect(outParam, OutputFileSht.Cells.SpecialCells(xlCellTypeAllValidation)) Is Nothing Then
            If foundfirstOutParam = False Then
                Set firstOutParam = outParam
                foundfirstOutParam = True
            End If
            Set lastOutParam = outParam
        End If
    
        'Rename important features like "Save" and "Browse"
        If outParam.Value = "Save" Then outParam.Name = "SaveOutput"
        If outParam.Value = "Browse" Then outParam.Name = "OutputBrowse"
        If InStr(1, outParam.Value, "AVAILABLE OUTPUTS") Then outParam.Name = "HeaderRow"
        If InStr(1, outParam.Value, "Units") Then outParam.Name = "UnitsColumn"
    Next
    
    'Deletes empty rows
    For Each outParam In OutputFileSht.Range(firstOutParam, lastOutParam)
        If WorksheetFunction.CountA(outParam.Rows.EntireRow) = 0 Then
    
            outParam.EntireRow.Delete
    
        End If
    Next
    
    'Gets string concatenation of cell addresses of all 'Yes'/'No' boxes
    For Each outParam In OutputFileSht.Range(firstOutParam, lastOutParam)
        If Not Intersect(outParam, OutputFileSht.Cells.SpecialCells(xlCellTypeAllValidation)) Is Nothing Then
                If outParam.Offset(0, OutputFileSht.Range("HeaderRow").Column - firstOutParam.Column).Font.Bold = True Then
                    If outParam.Value = vbNullString Then
                        outParam.Validation.Delete
                        outParam.Borders(xlEdgeLeft).LineStyle = xlNone
                        outParam.Borders(xlEdgeRight).LineStyle = xlNone
    
                    End If
                Else
                    If outParam.Value = vbNullString Then
                        outParam.Value = "-"
                    End If
    
                End If
                If Not Intersect(outParam, OutputFileSht.Cells.SpecialCells(xlCellTypeAllValidation)) Is Nothing Then
                    cellAddress = cellAddress & Replace(outParam.Address, "$", vbNullString) & ","
                End If
        End If
    Next
    
    ' Removes last comma and splits the string into an array
    cellAddress = Left(cellAddress, Len(cellAddress) - 1)
    cellAddresses() = Split(cellAddress, ",")
    
     ' Uses Union to loop through the string array and assign all addresses to a single range
     ' This is necessary since passing cellAddress to the Range() function directly has a limit of 255 characters
    
     For Each cellAddress In cellAddresses
         If finalOutputRange Is Nothing Then
            Set finalOutputRange = OutputFileSht.Range(cellAddresses(0))
         End If
         Set finalOutputRange = Union(finalOutputRange, OutputFileSht.Range(cellAddress))
    
     Next
     finalOutputRange.Name = "OutputParam"
    
    ' Assumes that the OutputFilePath bar will always be one row above the browse button
    FooterRow.Name = "FooterRow"
    OutputFileSht.Range("OutputBrowse").Offset(-1, 0).Name = "OutputFilePath"
    ' Places the values for the data validation to make the 'Yes'/'No' boxes in a spot that will not be used or altered
    OutputFileSht.Range("FF1010:FF1012").Name = "SummaryOption"
    OutputFileSht.Range("FF1011:FF1012").Name = "NoSummaryOption"
    OutputFileSht.Range("FF1010").Value = "Summarize"
    OutputFileSht.Range("FF1011").Value = "Detail"
    OutputFileSht.Range("FF1012").Value = "-"
    
    OutputFileSht.Range("V2").Name = "OutputConstColumn"

End Sub

' NameCellsUsingOutputs function
'
' The purpose of this subroutine is to name the 'Yes'/'No' with their
' corresponding output name. This is important for the InsertValue_BoolToYesNo
' function of the load module.

Sub NameCellsUsingOutputs()

    Dim outParam As Range
    Dim paramName As String
    
    ' All Yes/No boxes are named according to the output name which is found 2 cells to the left of the boxes
    For Each outParam In OutputFileSht.Range("OutputParam")
        If outParam.Value <> vbNullString Then
            If Not Intersect(outParam, OutputFileSht.Cells.SpecialCells(xlCellTypeAllValidation)) Is Nothing Then
    
                ' Gets the output name, and replaces invalid name characters with underscores (spaces are also replaced with underscores)
                ' For example, the Yes/No box for the output 'AC-side Efficiency' will be named 'AC_side_Efficiency'
    
                paramName = OutputFileSht.Cells(outParam.row, OutputFileSht.Range("OutputConstColumn").Column)
                paramName = Trim(paramName)
                paramName = Replace(paramName, "(", "_")
                paramName = Replace(paramName, ")", "_")
                paramName = Replace(paramName, "-", "_")
                paramName = Replace(paramName, "/", "_")
                paramName = Replace(paramName, "*", "_")
                paramName = Replace(paramName, "&", "_")
                
                If paramName <> vbNullString Then
                    outParam.Name = Replace(paramName, " ", "_")
                End If
                
            End If
        End If
    Next

End Sub


' Resizes Rows to be correctly formatted; also names sections for use by CheckBox module and WorkSheet_Change subroutines
' Name of cell one beneath the bolded header cells will become (First word of Header)_SectionStart,
' Name of cell two to the right the next section's header will be named (First word of Header)_SectionEnd
' This allows us to easily refer to the range of cells in the section block


Sub NameAndFormatSections(ByRef FooterRow As Range)

    Dim currentHeader As String
    Dim lastHeader As String
    Dim outParam As Range
    
    'This is required for the section naming to work properly (when it reaches the last section it requires another bold row to name the section end)
    If WorksheetFunction.CountA(FooterRow.Offset(-1, 0).Rows.EntireRow) = 0 Then
        FooterRow.Offset(-1, 0).EntireRow.Font.Bold = True
    Else
        FooterRow.EntireRow.Insert
        FooterRow.Offset(-1, 0).EntireRow.Validation.Delete
        FooterRow.Offset(-1, 0).EntireRow.Font.Bold = True
        FooterRow.Offset(-1, 0).EntireRow.Interior.Color = ColourWhite
    End If
    
    ' Section naming may fail if two sections start with the same word.
    
    For Each outParam In OutputFileSht.Range(Left(OutputFileSht.Range("HeaderRow").Address(0, 0), 1) & OutputFileSht.Range("HeaderRow").row, Left(OutputFileSht.Range("HeaderRow").Address(0, 0), 1) & OutputFileSht.Range("FooterRow").row)
        If outParam.Address <> FooterRow.Address Then
            If outParam.Font.Bold = True Then
                ' If the current row is a section header row (section header rows are bold)
                lastHeader = currentHeader
                currentHeader = outParam.Value
                        
                If InStr(currentHeader, " ") Then
                    ' Name the start of the section by taking the first word of the section name and appending '_SectionStart'
                    outParam.Offset(1, 0).Name = StrConv(Left(currentHeader, InStr(currentHeader, " ") - 1), vbProperCase) + "_SectionStart"
                Else
                    outParam.Offset(1, 0).Name = StrConv(currentHeader, vbProperCase) + "_SectionStart"
                End If
    
                If InStr(lastHeader, " ") Then
                    outParam.Offset(0, OutputFileSht.Range("OutputParam").Column - OutputFileSht.Range("HeaderRow").Column).Name = StrConv(Left(lastHeader, InStr(lastHeader, " ") - 1), vbProperCase) + "_SectionEnd"
                Else
                    outParam.Offset(0, OutputFileSht.Range("OutputParam").Column - OutputFileSht.Range("HeaderRow").Column).Name = StrConv(lastHeader, vbProperCase) + "_SectionEnd"
                End If
                
                outParam.Rows.RowHeight = 30
    
            Else
                outParam.Rows.RowHeight = 15.75
            End If
        End If
    Next
    
    ' Reformats the footer row height
    FooterRow.Rows.RowHeight = 15.75
    ' Hides all of the rows beneath the footer row
    OutputFileSht.Range("C" & FooterRow.Offset(1, 0).row, "C" & 1048576).Rows.EntireRow.Hidden = True

End Sub






