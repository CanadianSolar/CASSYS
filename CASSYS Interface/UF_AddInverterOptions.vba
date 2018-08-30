VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_AddInverterOptions 
   Caption         =   "CASSYS - Add Inverter:"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8310
   OleObjectBlob   =   "UF_AddInverterOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_AddInverterOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DefineNewInverterButton_Click()
    Me.Hide
    UF_AddInverter.Show
End Sub

Public Sub ImportONDFileButton_Click()
    
    Dim isOndImportSuccessful As Integer
    Dim FOpen As Variant
    Dim dupInverterRepeat As Integer          ' Used to remember user choice for duplicate OND files: 0 = not initialized, 1 = overwrite, 2 = skip
    Dim numberOfOnd As Integer
    Dim numberOfOndOverwritten As Integer
    Dim numberOfOndSkipped As Integer
    Dim numberOfOndAdded As Integer
    
    
    FOpen = Application.GetOpenFilename(Title:="Please choose a .OND file to import", FileFilter:="OND Files(*.OND),*.OND;.ond,", MultiSelect:=True)
    
    dupInverterRepeat = 0
    numberOfOnd = 1
    numberOfOverwritten = 0
    numberOfSkipped = 0
    numberOfAdded = 0
        
    ' If user has pressed the Cancel button, exit gracefully
    If IsArray(FOpen) = False Then
        Unload Me
        Exit Sub
    End If
        
    While numberOfOnd <= UBound(FOpen)
 
        If FOpen(numberOfOnd) <> False Then
            isOndImportSuccessful = ParseONDFile(FOpen(numberOfOnd), dupInverterRepeat)
            numberOfOnd = numberOfOnd + 1
            If isOndImportSuccessful = -1 Then
                numberOfOndOverwritten = numberOfOndOverwritten + 1
            ElseIf isOndImportSuccessful = 0 Then
                numberOfOndSkipped = numberOfOndSkipped + 1
            ElseIf isOndImportSuccessful = 1 Then
                numberOfOndAdded = numberOfOndAdded + 1
            End If
        Else
            Exit Sub
        End If
        
    Wend
    
    Call MsgBox(numberOfOnd - 1 & " OND files requested for import." & Chr(10) & numberOfOndOverwritten & " Overwritten" & Chr(10) & numberOfOndAdded & " Added" & Chr(10) & numberOfOndSkipped & " Skipped")
    Unload Me
    
End Sub



