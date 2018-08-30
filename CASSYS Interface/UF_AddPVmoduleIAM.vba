VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_AddPVmoduleIAM 
   Caption         =   "CASSYS - Add Module IAM:"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3105
   OleObjectBlob   =   "UF_AddPVmoduleIAM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_AddPVmoduleIAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' The purpose of this function is to check the fields all contains valid values
Public Sub AddIAM_Click()
               
    ' Check if the texts are in the correct range
    If Me.AOI1.Value < 0 Or Me.AOI2.Value < 0 Or Me.AOI3.Value < 0 Or Me.AOI4.Value < 0 Or Me.AOI5.Value < 0 Or Me.AOI6.Value < 0 Or Me.AOI7.Value < 0 Or Me.AOI8.Value < 0 Or Me.AOI9.Value < 0 Or _
       Me.AOI1.Value > 90 Or Me.AOI2.Value > 90 Or Me.AOI3.Value > 90 Or Me.AOI4.Value > 90 Or Me.AOI5.Value > 90 Or Me.AOI6.Value > 90 Or Me.AOI7.Value > 90 Or Me.AOI8.Value > 90 Or Me.AOI9.Value > 90 Then
        MsgBox " The incident angle is defined incorrectly, it should be in the range of 0 to 90"
                   
    ElseIf Me.Mod1.Value < 0 Or Me.Mod2.Value < 0 Or Me.Mod3.Value < 0 Or Me.Mod4.Value < 0 Or Me.Mod5.Value < 0 Or Me.Mod6.Value < 0 Or Me.Mod7.Value < 0 Or Me.Mod8.Value < 0 Or Me.Mod9.Value < 0 Or _
       Me.Mod1.Value > 1.5 Or Me.Mod2.Value > 1.5 Or Me.Mod3.Value > 1.5 Or Me.Mod4.Value > 1.5 Or Me.Mod5.Value > 1.5 Or Me.Mod6.Value > 1.5 Or Me.Mod7.Value > 1.5 Or Me.Mod8.Value > 1.5 Or Me.Mod9.Value > 1.5 Then
        MsgBox " The incident angle modifier is out of range, it should be in the range of 0 to 1.5"
    
    Else
        UF_AddPVmoduleIAM.Hide
    
    End If
    
End Sub


' Cancel_Click Function
' This function is called when the cancel button is clicked
'
' The Cancel button has the same effect as the OK button, for simplicity
Private Sub CancelIAM_Click()
    UF_AddPVmoduleIAM.Hide
End Sub




