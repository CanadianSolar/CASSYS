VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_AddPVModule 
   Caption         =   "CASSYS - Add Module:"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7275
   OleObjectBlob   =   "UF_AddPVModule.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_AddPVModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Add_PVModule function
'
' The purpose of this function is to add a new
' PV module using the values defined in the fields
' of the userform
Sub Add_PVModule()
    Dim last As Long ' last line in the sheet
    
    last = PV_DatabaseSht.Range("C" & PV_DatabaseSht.Rows.count).End(xlUp).row + 1
    
    ' add pv module
    PV_DatabaseSht.Range("C" & last).Value = Me.Model.Value
    
    ' add manufacturer
    PV_DatabaseSht.Range("B" & last).Value = Me.Manu.Value
    
    ' add data origin
    PV_DatabaseSht.Range("A" & last).Value = "User_Defined"
    
    ' add pv module
    PV_DatabaseSht.Range("H" & last).Value = Me.PNom.Value
    
    ' add module technology type
    PV_DatabaseSht.Range("J" & last).Value = Me.Tech.Value
    
    ' add number of cells in series
    PV_DatabaseSht.Range("K" & last).Value = Me.CellsinS.Value
    
    ' add number of cells in parallel
    PV_DatabaseSht.Range("L" & last).Value = Me.CellsinP.Value
    
    ' add gref
    PV_DatabaseSht.Range("M" & last).Value = Me.Gref.Value
    
    ' add tref
    PV_DatabaseSht.Range("N" & last).Value = Me.Tref.Value
    
    ' add vmpp
    PV_DatabaseSht.Range("O" & last).Value = Me.Vmpp.Value
    
    ' add impp
    PV_DatabaseSht.Range("P" & last).Value = Me.Impp.Value
    
    ' add voc
    PV_DatabaseSht.Range("Q" & last).Value = Me.Voc.Value
    
    ' add isc
    PV_DatabaseSht.Range("R" & last).Value = Me.Isc.Value
    
    ' add mIsc
    PV_DatabaseSht.Range("S" & last).Value = Me.mIsc.Value
    
    ' add mPmpp
    PV_DatabaseSht.Range("U" & last).Value = Me.mPmpp.Value
    
    ' add mVco
    PV_DatabaseSht.Range("T" & last).Value = Me.mVco.Value
    
    ' add Rsh0
    PV_DatabaseSht.Range("V" & last).Value = Me.Rsh0.Value
    
    ' add Rshexp
    PV_DatabaseSht.Range("W" & last).Value = Me.Rshexp.Value
    
    ' add Rshunt
    PV_DatabaseSht.Range("X" & last).Value = Me.Rshunt.Value
    
    ' add Rseries
    PV_DatabaseSht.Range("Y" & last).Value = Me.Rseries.Value
    
    ' add the numnber of bypass diodes
    PV_DatabaseSht.Range("AB" & last).Value = Me.NumDiodes.Value
    
    ' add bypass diode voltage
    PV_DatabaseSht.Range("AC" & last).Value = Me.DiodeVolt.Value
    
    PV_DatabaseSht.Rows(last).Interior.ColorIndex = 6
    
    If UF_SelectPVModule.Visible = True Then
        Unload Me
        Unload UF_SelectPVModule
        UF_SelectPVModule.Show
        Exit Sub
    End If
    
    Unload Me
End Sub

' Add_Click Function
' This function is called when the add button is clicked
'
' The purpose of this function is to call the add pv module
' function if the fields all contains valid values and are
' not empty
Private Sub Add_Click()
    Dim isValid As Boolean
    isValid = True
    ' check if any of the fields are empty
    If Me.Model.Value = vbNullString Or Me.Manu.Value = vbNullString Then
       MsgBox " The module is defined incorrectly."
       isValid = False
    ElseIf Me.PNom.Value = vbNullString Or Me.Tech.Value = vbNullString Or Me.CellsinS.Value = vbNullString Or Me.CellsinP.Value = vbNullString Or Me.Gref.Value = vbNullString Or Me.Tref.Value = vbNullString Or Me.Vmpp.Value = vbNullString Or Me.Impp.Value = vbNullString Or Me.Voc.Value = vbNullString Or Me.Isc.Value = vbNullString Then
        MsgBox " The module is defined incorrectly."
        isValid = False
    ElseIf Me.mIsc.Value = vbNullString Or Me.mPmpp.Value = vbNullString Or Me.mVco.Value = vbNullString Or Me.Rsh0.Value = vbNullString Or Me.Rshexp.Value = vbNullString Or Me.Rshunt.Value = vbNullString Or Me.Rseries.Value = vbNullString Or Me.NumDiodes.Value = vbNullString Or Me.DiodeVolt.Value = vbNullString Then
        MsgBox " The module is defined incorrectly."
        isValid = False
    End If
    
    ' check if the values are valid
    If isValid Then
        If IsNumeric(Me.PNom.Value) = False Or IsNumeric(Me.CellsinS.Value) = False Or IsNumeric(Me.CellsinP.Value) = False Or IsNumeric(Me.Gref.Value) = False Or IsNumeric(Me.Tref.Value) = False Or IsNumeric(Me.Vmpp.Value) = False Or IsNumeric(Me.Impp.Value) = False Or IsNumeric(Me.Voc.Value) = False Or IsNumeric(Me.Isc.Value) = False Then
            MsgBox " The module has invalid inputs"
            isValid = False
        ElseIf IsNumeric(Me.mIsc.Value) = False Or IsNumeric(Me.mPmpp.Value) = False Or IsNumeric(Me.mVco.Value) = False Or IsNumeric(Me.Rsh0.Value) = False Or IsNumeric(Me.Rshexp.Value) = False Or IsNumeric(Me.Rshunt.Value) = False Or IsNumeric(Me.Rseries.Value) = False Or IsNumeric(Me.NumDiodes.Value) = False Or IsNumeric(Me.DiodeVolt.Value) = False Then
            MsgBox " The module has invalid inputs"
            isValid = False
        End If
    End If
    
    If isValid Then
        Call Add_PVModule
    End If
End Sub

' Cancel_Click Function
' This function is called when the cancel button is clicked
'
' The purpose of this function is to close the userform if
' the user clicks cancel
Private Sub Cancel_Click()
    Unload Me
End Sub


' UserForm_Initialize
' This function is called when the UserForm is opened.
'
' The purpose of this function is to set the
' focus to the first cell
Private Sub UserForm_Initialize()
   Me.Model.SetFocus
End Sub

