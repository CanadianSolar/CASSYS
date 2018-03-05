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

Private Sub ImportONDFileButton_Click()
    
    Dim isImportSuccessful As Boolean
    Dim FOpen As Variant
    FOpen = Application.GetOpenFilename(title:="Please choose a .OND file to import", FileFilter:="OND Files(*.OND),*.OND;.ond," & "All Files (*.*),*.*")
    
    If FOpen <> False Then
        isImportSuccessful = ParseONDFile(FOpen)
    Else
        Exit Sub
    End If
    
    If isImportSuccessful = True Then
        Call MsgBox("Inverter successfully imported.", , "CASSYS: Import OND File")
        Unload Me
        'Unload UF_SelectInverter
        'UF_SelectInverter.Show
    Else
        Call MsgBox("The OND file was not correctly defined. Import was unsucessful.", vbExclamation, "CASSYS: Import OND File")
    End If
    
End Sub


