Attribute VB_Name = "UserFormModule"
Option Explicit

' OpenPVForm Function
'
' The purpose of this function is to open the
' UF_SelectPVModule userform to allow for module
' selection
Sub OpenPVForm()
    UF_SelectPVModule.Show
End Sub

' OpenInvForm Function
'
' The purpose of this function is to open the
' UF_SelectInverter userform to allow for inverter
' selection
Sub OpenInvForm()
    UF_SelectInverter.Show
End Sub

' OpenAddPVForm Function
'
' The purpose of this function is to open the
' UF_AddInverter userform to allow for creating
' new inverters
Sub OpenAddPVForm()
    UF_AddPVModule.Show
End Sub

' OpenAddInvForm Function
'
' The purpose of this function is to open the
' UF_AddInverter userform to allow for creating
' new inverters
Sub OpenAddInvForm()
    UF_AddInverterOptions.Show
End Sub
' OpenAddModuleOptions Function
'
' The purpose of this function is to open the
' UF_AddModuleOptions userform to allow the user to
' choose between importing or defining a module
Sub OpenAddModuleOptions()
    UF_AddModuleOptions.Show
End Sub
' OpenAddInvOptions Function
'
' The purpose of this function is to open the
' UF_AddModuleOptions userform to allow the user to
' choose between importing or defining an inverter
Sub OpenAddInvOptions()
    UF_AddInverterOptions.Show
End Sub

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


