VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_AddSameModuleOptions 
   Caption         =   "CASSYS Import existing PV module "
   ClientHeight    =   2490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7350
   OleObjectBlob   =   "UF_AddSameModuleOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_AddSameModuleOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Define the parameters, when user select which option should be used for import existing pan files: 1 = overwrite or 2 = skip
Dim choiceType As Integer

' This function returns the choice selected by the user for overwriting or skipping an existing PAN file
Public Function getChoice() As Integer
    getChoice = choiceType
End Function

' This function returns whether the choice made by the user (overwrite/skip) applies to all future occurrences
Public Function getChoiceRepeat() As Boolean
    getChoiceRepeat = CheckBox1.Value
End Function

' When 'Overwrite' is chosen
Public Sub CommandButton1_Click()
    choiceType = 1
    UF_AddSameModuleOptions.Hide
End Sub

' When 'Skip' is chosen
Public Sub CommandButton2_Click()
    choiceType = 2
    UF_AddSameModuleOptions.Hide
End Sub

