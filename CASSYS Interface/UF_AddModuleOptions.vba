VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_AddModuleOptions 
   Caption         =   "CASSYS - Add New PV Modules"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8310
   OleObjectBlob   =   "UF_AddModuleOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_AddModuleOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numberOfPan As Integer                  ' Number of PAN files selected by the user

Private Sub DefineNewModuleButton_Click()
    Me.Hide
    UF_AddPVModule.Show
End Sub

'Since 1.5.2, added batch import of PAN files; this function returns the number of files selected by the user
Public Function getNumberOfPan() As Integer
    getNumberOfPan = numberOfPan
End Function

Public Sub ImportPANFileButton_Click()
    
    Dim isImportSuccessful As Integer
    Dim FOpen As Variant

    Dim dupModuleRepeat As Integer          ' Used to remember user choice for duplicate PAN files: 0 = not initialized, 1 = overwrite, 2 = skip
    dupModuleRepeat = 0
    Dim numberOfOverwritten As Integer
    Dim numberOfSkipped As Integer
    Dim numberOfAdded As Integer
         
    FOpen = Application.GetOpenFilename(Title:="Please choose a .PAN file to import", FileFilter:="PAN Files(*.PAN),*.PAN;.pan," & "All Files (*.*),*.*", MultiSelect:=True)
     
    numberOfPan = 1
    numberOfOverwritten = 0
    numberOfSkipped = 0
    numberOfAdded = 0
    
    ' If user has pressed the Cancel button, exit gracefully
    If IsArray(FOpen) = False Then
        Unload Me
        Exit Sub
    End If
    
    While numberOfPan <= UBound(FOpen)
    
        If FOpen(numberOfPan) <> False Then
            isImportSuccessful = ParsePANFile(FOpen(numberOfPan), dupModuleRepeat)
            numberOfPan = numberOfPan + 1
            If isImportSuccessful = -1 Then
                numberOfOverwritten = numberOfOverwritten + 1
            ElseIf isImportSuccessful = 0 Then
                numberOfSkipped = numberOfSkipped + 1
            ElseIf isImportSuccessful = 1 Then
                numberOfAdded = numberOfAdded + 1
            End If
        Else
            Exit Sub
        End If
    Wend
    
    ' Give feedback to users, how many pv modules are succesfully proceesed
  
    Call MsgBox(numberOfPan - 1 & " Pan files requested for import." & Chr(10) & numberOfOverwritten & " Overwritten" & Chr(10) & numberOfAdded & " Added" & Chr(10) & numberOfSkipped & " Skipped")
    Unload Me
      
End Sub

