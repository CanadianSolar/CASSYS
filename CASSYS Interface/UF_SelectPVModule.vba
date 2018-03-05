VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_SelectPVModule 
   Caption         =   "CASSYS - Select Module:"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5880
   OleObjectBlob   =   "UF_SelectPVModule.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_SelectPVModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SubArrayHeight = 17 ' The section height of the sub-array information
Private Const PVDataHeight = 4 ' The height where the data starts in the pv module database


' Cancel_Click Function
' This function is called when the cancel button is clicked
'
' The purpose of this function is to close the userform if
' the user clicks cancel
Private Sub Cancel_Click()
    Unload Me
End Sub



' Manu_Change Function
' This function is called when the value in the Manu ComboBox is changed
'
' The purpose of this function is to update the model ComboBox's contents
' so that they correspond to the models from the selected manufacturer
Private Sub Manu_Change()
    Dim cmodel As Range ' The range counter
    Dim modelRange As Range ' The range of the model without any blank cells
    Dim getColumn As Integer ' Gets the column the model list is in
    Dim getRow As Integer ' Gets the first row of the model list
    Dim isUnique As Integer ' This value is used to represent if the manufacturer is unique or not, with 1 being unique
    Dim i As Integer
    Dim theList() As Variant
    
    ' Get the row and column of the model list
    getRow = PV_DatabaseSht.Range("Model").row
    getColumn = PV_DatabaseSht.Range("Model").Column
    
    ' Use the getRow and getColumn to get the range of the manufacturers without any blank cells
    Set modelRange = Range(PV_DatabaseSht.Cells(getRow, getColumn), PV_DatabaseSht.Cells(getRow, getColumn).End(xlDown))
    
    'If the text in the ComboBox is in the list then
    If Manu.ListIndex > -1 Then
        'Clear the original contents
        Me.Model.Clear
        
        ' Add the manufacturer if it is unique
        With CreateObject("Scripting.Dictionary")
            For Each cmodel In modelRange
            ' Default to true
                If Not .exists(cmodel.Value) And cmodel.Offset(0, -1).Value = Me.Manu.Value Then
                    .Add cmodel.Value, Nothing
                End If
    
            Next cmodel
        
            ReDim theList(.count)
        
            theList = .keys
            
            ' Sorts the array
            Call QuickSort(theList, 0, .count - 1)
            
            ' Adds the list to the ComboBox
            Me.Model.List = theList
        End With
    Else
        'Clear the list if the text in the ComboBox is incorrect
        Me.Model.Clear
    End If
End Sub

'Model_Change Function
' This function is called when the value in the Model ComboBox is changed
'
' The purpose of this function is to update the labels based on the values
' given in the database of the chosen model
Private Sub Model_Change()
    Dim cmodel As Range ' The range counter
    Dim modelRange As Range ' The range of the model without any blank cells
    Dim themodule As Range ' The range counter
    Dim first As Range
    Dim getColumn As Integer ' Gets the column the model list is in
    Dim getRow As Integer ' Gets the first row of the model list
    Dim isUnique As Integer ' This value is used to represent if the manufacturer is unique or not, with 1 being unique
    Dim i As Integer
    Dim getSource As String
    Dim theList() As Variant
     
    ' Get the row and column of the model list
    getRow = PV_DatabaseSht.Range("Model").row
    getColumn = PV_DatabaseSht.Range("Model").Column
    
    getSource = Me.Source.Value
    
    ' Use the getRow and getColumn to get the range of the manufacturers without any blank cells
    Set modelRange = Range(PV_DatabaseSht.Cells(getRow, getColumn), PV_DatabaseSht.Cells(getRow, getColumn).End(xlDown))
    
    'Try to find the module
    Set themodule = modelRange.Find(Me.Model.Value, LookIn:=xlValues, LookAt:=xlWhole)
    Set first = themodule
    
    'If the module is in the database
    If Not themodule Is Nothing Then
        'And If the manufacturer is correct
        If themodule.Offset(0, -1).Value = Me.Manu.Value Then
            'Update the corresponding labels
            Me.MaxPow.Caption = themodule.Offset(0, 5).Value
            Me.CurrentPmpp.Caption = themodule.Offset(0, 13).Value
            Me.VoltagePmpp.Caption = themodule.Offset(0, 12).Value
            Me.MaxCurr.Caption = themodule.Offset(0, 15).Value
            Me.MaxVolt.Caption = themodule.Offset(0, 14).Value
            Me.Rshunt.Caption = themodule.Offset(0, 21).Value
            Me.Rseries.Caption = themodule.Offset(0, 22).Value
        Else
            'If it is not correct keep trying to find it until it hits the end
            Do
                Set themodule = modelRange.FindNext(themodule)
            Loop While Not themodule Is Nothing And Not themodule.Offset(0, -1).Value = Me.Manu.Value And themodule.Address <> first.Address
            
            'If the model exists update the labels
            If Not themodule Is Nothing And themodule.Address <> first.Address Then
                Me.MaxPow.Caption = themodule.Offset(0, 5).Value
                Me.CurrentPmpp.Caption = themodule.Offset(0, 13).Value
                Me.VoltagePmpp.Caption = themodule.Offset(0, 12).Value
                Me.MaxCurr.Caption = themodule.Offset(0, 15).Value
                Me.MaxVolt.Caption = themodule.Offset(0, 14).Value
                Me.Rshunt.Caption = themodule.Offset(0, 21).Value
                Me.Rseries.Caption = themodule.Offset(0, 22).Value
            Else
                'If not clear the labels
                Me.MaxPow.Caption = 0
                Me.CurrentPmpp.Caption = 0
                Me.VoltagePmpp.Caption = 0
                Me.MaxCurr.Caption = 0
                Me.MaxVolt.Caption = 0
                Me.Rshunt.Caption = 0
                Me.Rseries.Caption = 0
            End If
        End If
    Else
       ' If not clear the labels
        Me.MaxPow.Caption = 0
        Me.CurrentPmpp.Caption = 0
        Me.VoltagePmpp.Caption = 0
        Me.MaxCurr.Caption = 0
        Me.MaxVolt.Caption = 0
        Me.Rshunt.Caption = 0
        Me.Rseries.Caption = 0
    End If
    
        'If the text in the ComboBox is in the list then
    If Model.ListIndex > -1 Then
        'Clear the original contents
        Me.Source.Clear
    
    ' Go through each model and only add it if its manufacturer is correct
    
    ' Add the manufacturer if it is unique
        With CreateObject("Scripting.Dictionary")
            For Each cmodel In modelRange
            ' Default to true
                If Not .exists(cmodel.Offset(0, -2).Value) And cmodel.Offset.Value = Me.Model.Value Then
                    .Add cmodel.Offset(0, -2).Value, Nothing
                End If
    
            Next cmodel
        
            ReDim theList(.count)
        
            theList = .keys

            Call QuickSort(theList, 0, .count - 1)

            Me.Source.List = theList
        End With
        Me.Source.Value = getSource
    Else
        'Clear the list if the text in the ComboBox is incorrect
        Me.Source.Clear
    End If
    
     
End Sub

Private Sub NewModuleButton_Click()
    UF_AddModuleOptions.Show
End Sub

Private Sub Source_Change()
    Dim getIndex As Integer ' the Index of the model
    
    ' If the value is not in the range clear the box
    If Not Source.ListIndex > -1 Then
        Me.Source.Value = vbNullString
    Else
        ' If it is get the Index
        getIndex = PVIndex(Me.Manu.Value, Me.Model.Value, Me.Source.Value)
        
        ' If found output the values to the form
        If Not (getIndex = 0) Then
            Me.MaxPow.Caption = PV_DatabaseSht.Cells(getIndex + PVDataHeight, 8).Value
            Me.CurrentPmpp.Caption = PV_DatabaseSht.Cells(getIndex + PVDataHeight, 16).Value
            Me.VoltagePmpp.Caption = PV_DatabaseSht.Cells(getIndex + PVDataHeight, 15).Value
            Me.MaxCurr.Caption = PV_DatabaseSht.Cells(getIndex + PVDataHeight, 18).Value
            Me.MaxVolt.Caption = PV_DatabaseSht.Cells(getIndex + PVDataHeight, 27).Value
            Me.Rshunt.Caption = PV_DatabaseSht.Cells(getIndex + PVDataHeight, 24).Value
            Me.Rseries.Caption = PV_DatabaseSht.Cells(getIndex + PVDataHeight, 25).Value
        End If
    End If
End Sub

' Paste_Click Function
' This function is called when the select button is pressed
'
' The purpose of this function is to update the value of the
' Sub-Array Model cell to that of the chosen model in the form
Private Sub Paste_Click()
    Dim getIndex As Integer
    Dim Err As String
    Err = vbNullString
    'If the model has a selected entry update the value in the SystemSht page and exit
    If Not Me.Model.Value = vbNullString And Not Me.Source.Value = vbNullString Then
        
        ' Find the Index of the entry
        getIndex = PVIndex(Me.Manu.Value, Me.Model.Value, Me.Source.Value)
        
        'If it exists update the Index cell
        If Not getIndex = 0 Then
            SystemSht.Range("PVDataIndex").Offset(((Range("PVModuleIndex").Value - 1) * SubArrayHeight), 0).Value = getIndex
        Else
            'If not display error message
            '
        End If
        Unload Me
    Else
        'If not display error message
        If Me.Manu.Value = vbNullString Then
            Err = "Manufacturer" & Constants.vbCrLf
        End If
        
        If Me.Model.Value = vbNullString Then
            Err = Err & "Model" & Constants.vbCrLf
        End If
        
        If Me.Source.Value = vbNullString Then
            Err = Err & "Version Origin" & Constants.vbCrLf
        End If
        
        MsgBox "Please select a:" & Constants.vbCrLf & Err
    End If
End Sub


' UserForm_Initialize
' This function is called when the UserForm is opened.
'
' The purpose of this function is to intialize the labels
' captions and the contents of the manufacturer ComboBox
Private Sub UserForm_Initialize()
    Dim cmanu As Range ' The range counter
    Dim i As Integer ' Counter variable
    Dim getRow As Integer ' Gets the first row of the model list
    Dim getColumn As Integer ' Gets the column the model list is in
    Dim isUnique As Integer ' This value is used to represent if the manufacturer is unique or not, with 1 being unique
    Dim modelRange As Range ' The range of the manufacturers without the whitespace
    Dim theList() As Variant
    
    ' Get the row and column of the model list
    getRow = PV_DatabaseSht.Range("Model").row
    getColumn = PV_DatabaseSht.Range("Model").Column
    
    ' Use the getRow and getColumn to get the range of the manufacturers without any blank cells
    Set modelRange = Range(PV_DatabaseSht.Cells(getRow, getColumn), PV_DatabaseSht.Cells(getRow, getColumn).End(xlDown))
    
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
    ' Add the manufacturer if it is unique
    With CreateObject("Scripting.Dictionary")
        For Each cmanu In modelRange.Offset(0, -1)
        ' Default to true
            If Not .exists(cmanu.Value) Then
                .Add cmanu.Value, Nothing
            End If
    
        Next cmanu
        
        ReDim theList(.count)
        
        theList = .keys

        Call QuickSort(theList, 0, .count - 1)

        Me.Manu.List = theList

    End With
    
    If Range("PVDataIndex").Value <> -1 Then
        Me.Manu.Value = SystemSht.Range("ModuleManu").Offset(((Range("PVModuleIndex").Value - 1) * SubArrayHeight), 0).Value
        Me.Model.Value = SystemSht.Range("ModuleModel").Offset(((Range("PVModuleIndex").Value - 1) * SubArrayHeight), 0).Value
        Me.Source.Value = SystemSht.Range("ModuleSource").Offset(((Range("PVModuleIndex").Value - 1) * SubArrayHeight), 0).Value
    Else
        Me.Manu.Value = "Please first select a manufacturer. -->"
    End If
    
End Sub


