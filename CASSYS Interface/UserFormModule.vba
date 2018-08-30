Attribute VB_Name = "UserFormModule"
Option Explicit

' OpenPVForm Function
'
' The purpose of this function is to open the
' UF_SelectPVModule userform to allow for module
' selection
Function OpenPVForm() As Boolean
    UF_SelectPVModule.Show
End Function

' OpenInvForm Function
'
' The purpose of this function is to open the
' UF_SelectInverter userform to allow for inverter
' selection
Function OpenInvForm() As Boolean
    UF_SelectInverter.Show
End Function

' OpenAddPVForm Function
'
' The purpose of this function is to open the
' UF_AddInverter userform to allow for creating
' new inverters
Function OpenAddPVForm() As Boolean
    UF_AddPVModule.Show
End Function

' OpenAddInvForm Function
'
' The purpose of this function is to open the
' UF_AddInverter userform to allow for creating
' new inverters
Function OpenAddInvForm() As Boolean
    UF_AddInverterOptions.Show
End Function
' OpenAddModuleOptions Function
'
' The purpose of this function is to open the
' UF_AddModuleOptions userform to allow the user to
' choose between importing or defining a module
Function OpenAddModuleOptions() As Boolean
    UF_AddModuleOptions.Show
End Function

' Save module database
Function SaveModuleDatabase() As Boolean
    Call SaveDatabase("module", PV_DatabaseSht)
End Function

' Save inverter database
Function SaveInverterDatabase() As Boolean
    Call SaveDatabase("inverter", Inverter_DatabaseSht)
End Function

' Generic function to save database
' This saves CASSYS, first removing the current model and then overwriting the current xlsm
Function SaveDatabase(dbName As String, sht As Worksheet) As Boolean

    Dim answer As VbMsgBoxResult
    Dim shtStatus As sheetStatus
    
    ' First ask for confirmation
    answer = MsgBox("Saving the " & dbName & " database will erase the current model. Continue?", vbYesNo)
    If answer = vbNo Then Exit Function
    
    ' Call PreModify
    Call PreModify(sht, shtStatus)
    
    ' Wipe out model
    Call PrepareForRelease
    
    ' Save file
    Call SaveCASSYSWorkbook
    
    ' Do as if the workbook is re-open, so that the normal tabs show
    Call ThisWorkbook.WorkbookOpen
    
    ' Open the inverter database tab
    sht.Visible = xlSheetVisible
    sht.Activate
    
    ' Restore workbook status
    Call PostModify(sht, shtStatus)

End Function
' OpenAddInvOptions Function
'
' The purpose of this function is to open the
' UF_AddModuleOptions userform to allow the user to
' choose between importing or defining an inverter
Function OpenAddInvOptions() As Boolean
    UF_AddInverterOptions.Show
End Function

' PVIndex Function
'
' Arguments
' Manu - The name of the Manufacturer
' Model - The name of the Model
' Source - The name of the Data/Version Source
'
' Returns Integer - The Index of the pv module that matches
' the name of the model, manufacturer and data source of
' the arguments or 0 if the module is not found
'
' The purpose of this function is to find the Index of the
' requested model based on the arguments given
Function PVIndex(ByVal Manu As String, ByVal Model As String, ByVal Source As String) As Integer

    Dim modelRange As Range ' The range of the model without any blank cells
    Dim themodule As Range ' The range counter
    Dim first As Range
    Dim getColumn As Integer ' Gets the column the model list is in
    Dim getRow As Long
   
    ' Get the row and column of the model list
    getRow = PV_DatabaseSht.Range("Model").row
    getColumn = PV_DatabaseSht.Range("Model").Column
    
    ' Use the getRow and getColumn to get the range of the manufacturers without any blank cells
    Set modelRange = Range(PV_DatabaseSht.Cells(getRow, getColumn), PV_DatabaseSht.Cells(getRow, getColumn).End(xlDown))

    'Try to find the module
    Set themodule = modelRange.Find(Model, LookIn:=xlValues, LookAt:=xlWhole)
    Set first = themodule

    ' If the module exists
    If Not themodule Is Nothing Then
        ' If the data source and manufacturer match return the row Index
        ' NB: changed if statement to match InvIndex Function
        If themodule.Offset(0, -1).Value = Manu And themodule.Offset(0, -2).Value = Source Or themodule.Offset(0, -2).Value = "User_Added" Then
            PVIndex = themodule.row - PVDataHeight
        Else
            ' If not continue until it is found or for some reason does not exist
            Do
                ' Finds the next instance of the module
                Set themodule = modelRange.FindNext(themodule)
            Loop While (Not themodule Is Nothing) And Not (themodule.Offset(0, -1).Value = Manu And (themodule.Offset(0, -2).Value = Source Or themodule.Offset(0, -2).Value = "User_Added")) And themodule.Address <> first.Address
            
            ' If the module does exist, return the row Index
            ' NB: Adjusted if statement to match InvIndex Function again
            If Not themodule Is Nothing And themodule.Address <> first.Address Then
                PVIndex = themodule.row - PVDataHeight
            Else
                ' If not return error value
                ' NB: Changed error value to 0 from -1, only one error value probably better
                PVIndex = 0
            End If
        End If
    Else
        ' If not return error value
        PVIndex = 0
    End If
    
End Function

' InvIndex Function
'
' Arguments
' Manu - The name of the Manufacturer
' Model - The name of the Model
' Source - The name of the Data/Version Source
'
' Returns Integer - The Index of the inverter model that matches
' the name of the model, manufacturer and data source of
' the arguments or 0 if the inverter is not found
'
' The purpose of this function is to find the Index of the
' requested model based on the arguments given
Function InvIndex(ByVal Manu As String, ByVal Model As String, ByVal Source As String) As Integer

    Dim modelRange As Range ' The range of the model without any blank cells
    Dim themodule As Range ' The range counter
    Dim first As Range
    Dim getColumn As Integer ' Gets the column the model list is in
    Dim getRow As Integer ' Gets the first row of the model list
    
    ' Get the row and column of the model list
    getRow = Inverter_DatabaseSht.Range("Inverter").row
    getColumn = Inverter_DatabaseSht.Range("Inverter").Column
    ' Use the getRow and getColumn to get the range of the manufacturers without any blank cells
    Set modelRange = Range(Inverter_DatabaseSht.Cells(getRow, getColumn), Inverter_DatabaseSht.Cells(getRow, getColumn).End(xlDown))

    'Try to find the module
    Set themodule = modelRange.Find(Model, LookIn:=xlValues, LookAt:=xlWhole)
    Set first = themodule
    
    ' If the module exists
    If Not themodule Is Nothing Then
        ' If the data source and manufacturer match return the row Index
        If themodule.Offset(0, -1).Value = Manu And themodule.Offset(0, -2).Value = Source Or themodule.Offset(0, -2).Value = "User_Added" Then
            InvIndex = themodule.row - InvDataHeight
        Else
            ' If not continue until it is found or for some reason does not exist
            Do
                ' Finds the next instance of the module
                Set themodule = modelRange.FindNext(themodule)
            Loop While (Not themodule Is Nothing) And Not (themodule.Offset(0, -1).Value = Manu And (themodule.Offset(0, -2).Value = Source Or themodule.Offset(0, -2).Value = "User_Added")) And themodule.Address <> first.Address
            
            ' If the module does exist return the row Index
            If Not themodule Is Nothing And themodule.Address <> first.Address Then
                InvIndex = themodule.row - InvDataHeight
            Else
                ' If not return error value
                InvIndex = 0
            End If
        End If
    Else
        ' If not return error value
        InvIndex = 0
    End If
    
End Function


