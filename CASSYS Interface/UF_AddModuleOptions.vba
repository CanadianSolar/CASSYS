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
Private Sub DefineNewModuleButton_Click()
    Me.Hide
    UF_AddPVModule.Show
End Sub

Private Sub ImportPANFileButton_Click()
    
    Dim isImportSuccessful As Boolean
    Dim FOpen As Variant
    FOpen = Application.GetOpenFilename(title:="Please choose a .PAN file to import", FileFilter:="PAN Files(*.PAN),*.PAN;.pan," & "All Files (*.*),*.*")
    
    If FOpen <> False Then
        isImportSuccessful = ParsePANFile(FOpen)
    Else
        Exit Sub
    End If
    
    If isImportSuccessful = True Then
        Call MsgBox("PV Module successfully imported.", , "CASSYS: Import PAN File")
        Unload Me
        'Unload UF_SelectPVModule
        'UF_SelectPVModule.Show
    Else
        Call MsgBox("The PAN file was not correctly defined. Import was unsucessful.", vbExclamation, "CASSYS: Import PAN File")
    End If
    
End Sub

