VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_AddInverter 
   Caption         =   "CASSYS - Add Inverter:"
   ClientHeight    =   11115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7110
   OleObjectBlob   =   "UF_AddInverter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_AddInverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Add_Inverter function
'
' The purpose of this function is to add a new
' inverter using the values defined in the fields
' of the userform
Sub Add_Inverter()
    Dim last As Long ' last line in the sheet
    
    last = Inverter_DatabaseSht.Range("C" & Inverter_DatabaseSht.Rows.count).End(xlUp).row + 1
    ' add inverter model
    Inverter_DatabaseSht.Range("C" & last).Value = Me.Model.Value
    
    ' add manufacturer
    Inverter_DatabaseSht.Range("B" & last).Value = Me.Manu.Value
    
    ' add data origin
    Inverter_DatabaseSht.Range("A" & last).Value = "User_Defined"
    
    ' add nominal power
    Inverter_DatabaseSht.Range("H" & last).Value = Me.PNomAC.Value
    
    ' add max power
    Inverter_DatabaseSht.Range("I" & last).Value = Me.PMaxAC.Value
    
    ' add nominal current
    Inverter_DatabaseSht.Range("J" & last).Value = Me.INom.Value
    
    ' add max current
    Inverter_DatabaseSht.Range("K" & last).Value = Me.IMax.Value
    
    ' add output
    Inverter_DatabaseSht.Range("L" & last).Value = Me.Output.Value
    
    ' add inverter type
    Inverter_DatabaseSht.Range("M" & last).Value = Me.TypeField.Value
    
    ' add frequency
    Inverter_DatabaseSht.Range("N" & last).Value = Me.Freq.Value
    
    ' add threshold
    Inverter_DatabaseSht.Range("R" & last).Value = Me.Thresh.Value
    
    ' add Min MPP
    Inverter_DatabaseSht.Range("T" & last).Value = Me.MinMPP.Value
    
    ' add Max MPP
    Inverter_DatabaseSht.Range("U" & last).Value = Me.MaxMPP.Value
    
    ' add minimum voltage
    Inverter_DatabaseSht.Range("W" & last).Value = Me.MinV.Value
    
    ' add oper
    Inverter_DatabaseSht.Range("AA" & last).Value = Me.Oper.Value
    
    'add bipolar in
    Inverter_DatabaseSht.Range("BT" & last).Value = Me.BipolarIn.Value
    
    
    If Me.EffMultiPage.Value = 0 Then
        ' add single curve values
        
        Inverter_DatabaseSht.Range("AP" & last).Value = Me.SingleMax.Value
        
        Inverter_DatabaseSht.Range("AQ" & last).Value = Me.SingleEuro.Value
    
        ' add data points
        Inverter_DatabaseSht.Range("AS" & last).Value = Me.SingleIN2.Value
    
        Inverter_DatabaseSht.Range("AU" & last).Value = Me.SingleIN3.Value
    
        Inverter_DatabaseSht.Range("AW" & last).Value = Me.SingleIN4.Value
    
        Inverter_DatabaseSht.Range("AY" & last).Value = Me.SingleIN5.Value
    
        Inverter_DatabaseSht.Range("BA" & last).Value = Me.SingleIN6.Value
    
        Inverter_DatabaseSht.Range("BC" & last).Value = Me.SingleIN7.Value
    
        Inverter_DatabaseSht.Range("BE" & last).Value = Me.SingleIN8.Value
    
        Inverter_DatabaseSht.Range("AT" & last).Value = Me.SingleEffic2.Value
    
        Inverter_DatabaseSht.Range("AV" & last).Value = Me.SingleEffic3.Value

        Inverter_DatabaseSht.Range("AX" & last).Value = Me.SingleEffic4.Value
    
        Inverter_DatabaseSht.Range("AZ" & last).Value = Me.SingleEffic5.Value
    
        Inverter_DatabaseSht.Range("BB" & last).Value = Me.SingleEffic6.Value
    
        Inverter_DatabaseSht.Range("BD" & last).Value = Me.SingleEffic7.Value
    
        Inverter_DatabaseSht.Range("BF" & last).Value = Me.SingleEffic8.Value
    ElseIf Me.EffMultiPage.Value = 1 Then
        'add low voltage curve values
        ' add single curve values
        
        Inverter_DatabaseSht.Range("AO" & last).Value = "X"
        
        
        ' add low voltage
        
        
        Inverter_DatabaseSht.Range("BX" & last).Value = Me.LowVolt.Value
    
        ' add data points
        Inverter_DatabaseSht.Range("CA" & last).Value = Me.LowIN2.Value
    
        Inverter_DatabaseSht.Range("CC" & last).Value = Me.LowIN3.Value
    
        Inverter_DatabaseSht.Range("CE" & last).Value = Me.LowIN4.Value
    
        Inverter_DatabaseSht.Range("CG" & last).Value = Me.LowIN5.Value
    
        Inverter_DatabaseSht.Range("CI" & last).Value = Me.LowIN6.Value
    
        Inverter_DatabaseSht.Range("CK" & last).Value = Me.LowIN7.Value
    
        Inverter_DatabaseSht.Range("CM" & last).Value = Me.LowIN8.Value
    
        Inverter_DatabaseSht.Range("CB" & last).Value = Me.LowEffic2.Value
    
        Inverter_DatabaseSht.Range("CD" & last).Value = Me.LowEffic3.Value

        Inverter_DatabaseSht.Range("CF" & last).Value = Me.LowEffic4.Value
    
        Inverter_DatabaseSht.Range("CH" & last).Value = Me.LowEffic5.Value
    
        Inverter_DatabaseSht.Range("CJ" & last).Value = Me.LowEffic6.Value
    
        Inverter_DatabaseSht.Range("CL" & last).Value = Me.LowEffic7.Value
    
        Inverter_DatabaseSht.Range("CN" & last).Value = Me.LowEffic8.Value
        
        
        ' add medium voltage curve values
        
        
        Inverter_DatabaseSht.Range("CO" & last).Value = Me.MedVolt.Value
    
        ' add data points
        Inverter_DatabaseSht.Range("CR" & last).Value = Me.MedIN2.Value
    
        Inverter_DatabaseSht.Range("CT" & last).Value = Me.MedIN3.Value
    
        Inverter_DatabaseSht.Range("CV" & last).Value = Me.MedIN4.Value
    
        Inverter_DatabaseSht.Range("CX" & last).Value = Me.MedIN5.Value
    
        Inverter_DatabaseSht.Range("CZ" & last).Value = Me.MedIN6.Value
    
        Inverter_DatabaseSht.Range("DB" & last).Value = Me.MedIN7.Value
    
        Inverter_DatabaseSht.Range("DD" & last).Value = Me.MedIN8.Value
    
        Inverter_DatabaseSht.Range("CS" & last).Value = Me.MedEffic2.Value
    
        Inverter_DatabaseSht.Range("CU" & last).Value = Me.MedEffic3.Value

        Inverter_DatabaseSht.Range("CW" & last).Value = Me.MedEffic4.Value
    
        Inverter_DatabaseSht.Range("CY" & last).Value = Me.MedEffic5.Value
    
        Inverter_DatabaseSht.Range("DA" & last).Value = Me.MedEffic6.Value
    
        Inverter_DatabaseSht.Range("DC" & last).Value = Me.MedEffic7.Value
    
        Inverter_DatabaseSht.Range("DE" & last).Value = Me.MedEffic8.Value
        
        
        ' add high voltage curve values
        
        
        Inverter_DatabaseSht.Range("DF" & last).Value = Me.HighVolt.Value
    
        ' add data points
        Inverter_DatabaseSht.Range("DI" & last).Value = Me.HighIN2.Value
    
        Inverter_DatabaseSht.Range("DK" & last).Value = Me.HighIN3.Value
    
        Inverter_DatabaseSht.Range("DM" & last).Value = Me.HighIN4.Value
    
        Inverter_DatabaseSht.Range("DO" & last).Value = Me.HighIN5.Value
    
        Inverter_DatabaseSht.Range("DQ" & last).Value = Me.HighIN6.Value
    
        Inverter_DatabaseSht.Range("DS" & last).Value = Me.HighIN7.Value
    
        Inverter_DatabaseSht.Range("DU" & last).Value = Me.HighIN8.Value
    
        Inverter_DatabaseSht.Range("DJ" & last).Value = Me.HighEffic2.Value
    
        Inverter_DatabaseSht.Range("DL" & last).Value = Me.HighEffic3.Value

        Inverter_DatabaseSht.Range("DN" & last).Value = Me.HighEffic4.Value
    
        Inverter_DatabaseSht.Range("DP" & last).Value = Me.HighEffic5.Value
    
        Inverter_DatabaseSht.Range("DR" & last).Value = Me.HighEffic6.Value
    
        Inverter_DatabaseSht.Range("DT" & last).Value = Me.HighEffic7.Value
    
        Inverter_DatabaseSht.Range("DV" & last).Value = Me.HighEffic8.Value
    End If
    
    Inverter_DatabaseSht.Rows(last).Interior.ColorIndex = 6
    
    If UF_SelectInverter.Visible = True Then
        Unload Me
        Unload UF_SelectInverter
        UF_SelectInverter.Show
        Exit Sub
    End If
    
    Unload Me
    
End Sub

' Add_Click Function
' This function is called when the add button is clicked
'
' The purpose of this function is to call the add inverter
' function if the fields all contains valid values and are
' not empty
Private Sub Add_Click()
    Dim isValid As Boolean
    isValid = True
    
    
    ' check if the inverter info is defined
    If (Me.Model.Value = vbNullString Or Me.Manu.Value = vbNullString) Then
        MsgBox "The inverter is not sufficiently defined."
    ElseIf (Me.PNomAC.Value = vbNullString Or Me.PMaxAC.Value = vbNullString Or Me.INom.Value = vbNullString Or Me.IMax.Value = vbNullString Or Me.Output.Value = vbNullString Or Me.TypeField.Value = vbNullString Or Me.BipolarIn.Value = vbNullString Or Me.Freq.Value = vbNullString Or Me.Thresh.Value = vbNullString Or Me.MinMPP.Value = vbNullString Or Me.MaxMPP.Value = vbNullString Or Me.MinV.Value = vbNullString Or Me.Oper.Value = vbNullString) Then
        MsgBox "The inverter is not sufficiently defined."
    End If
    
    ' check if the efficiency curves are defined
    
    ' Iff single efficiencty curve is chosen
    If Me.EffMultiPage.Value = 0 Then
        ' check if the single voltage curve is defined
        If (Me.SingleMax.Value = vbNullString Or Me.SingleIN2.Value = vbNullString Or Me.SingleIN3.Value = vbNullString Or Me.SingleIN4.Value = vbNullString Or Me.SingleIN5.Value = vbNullString Or Me.SingleIN6.Value = vbNullString Or Me.SingleIN7.Value = vbNullString Or Me.SingleIN8.Value = vbNullString) Then
            MsgBox "The efficiency curves are not fully defined"
            isValid = False
        ElseIf (Me.SingleEuro.Value = vbNullString Or Me.SingleEffic2.Value = vbNullString Or Me.SingleEffic3.Value = vbNullString Or Me.SingleEffic4.Value = vbNullString Or Me.SingleEffic5.Value = vbNullString Or Me.SingleEffic6.Value = vbNullString Or Me.SingleEffic7.Value = vbNullString Or Me.SingleEffic8.Value = vbNullString) Then
            MsgBox "The efficiency curves are not fully defined"
            isValid = False
        End If
    ElseIf Me.EffMultiPage.Value = 1 Then
        ' if multiple curves are chosen
    
        ' check if the low volatge curve is defined
        If (Me.LowVolt.Value = vbNullString Or Me.LowIN2.Value = vbNullString Or Me.LowIN3.Value = vbNullString Or Me.LowIN4.Value = vbNullString Or Me.LowIN5.Value = vbNullString Or Me.LowIN6.Value = vbNullString Or Me.LowIN7.Value = vbNullString Or Me.LowIN8.Value = vbNullString) Then
            MsgBox " The efficiency curves for low voltage are not fully defined"
            isValid = False
        ElseIf (Me.LowEffic2.Value = vbNullString Or Me.LowEffic3.Value = vbNullString Or Me.LowEffic4.Value = vbNullString Or Me.LowEffic5.Value = vbNullString Or Me.LowEffic6.Value = vbNullString Or Me.LowEffic7.Value = vbNullString Or Me.LowEffic8.Value = vbNullString) Then
            MsgBox " The efficiency curves for low voltage are not fully defined"
            isValid = False
        ElseIf (Me.MedVolt.Value = vbNullString Or Me.MedIN2.Value = vbNullString Or Me.MedIN3.Value = vbNullString Or Me.MedIN4.Value = vbNullString Or Me.MedIN5.Value = vbNullString Or Me.MedIN6.Value = vbNullString Or Me.MedIN7.Value = vbNullString Or Me.MedIN8.Value = vbNullString) Then
            ' check if the medium voltage curve is defined
            MsgBox " The efficiency curves for medium voltage are not fully defined"
            isValid = False
        ElseIf (Me.MedEffic2.Value = vbNullString Or Me.MedEffic3.Value = vbNullString Or Me.MedEffic4.Value = vbNullString Or Me.MedEffic5.Value = vbNullString Or Me.MedEffic6.Value = vbNullString Or Me.MedEffic7.Value = vbNullString Or Me.MedEffic8.Value = vbNullString) Then
            MsgBox " The efficiency curves for medium voltage are not fully defined"
            isValid = False
        ElseIf (Me.HighVolt.Value = vbNullString Or Me.HighIN2.Value = vbNullString Or Me.HighIN3.Value = vbNullString Or Me.HighIN4.Value = vbNullString Or Me.HighIN5.Value = vbNullString Or Me.HighIN6.Value = vbNullString Or Me.HighIN7.Value = vbNullString Or Me.HighIN8.Value = vbNullString) Then
            ' check if the high voltage curve is defined
            MsgBox " The efficiency curves for high voltage are not fully defined"
            isValid = False
        ElseIf (Me.HighEffic2.Value = vbNullString Or Me.HighEffic3.Value = vbNullString Or Me.HighEffic4.Value = vbNullString Or Me.HighEffic5.Value = vbNullString Or Me.HighEffic6.Value = vbNullString Or Me.HighEffic7.Value = vbNullString Or Me.HighEffic8.Value = vbNullString) Then
            MsgBox " The efficiency curves for high voltage are not fully defined"
            isValid = False
        End If
    End If
    
    If isValid Then
        ' check if the values are numeric
        If IsNumeric(Me.PNomAC.Value) = False Or IsNumeric(Me.PMaxAC.Value) = False Or IsNumeric(Me.INom.Value) = False Or IsNumeric(Me.IMax.Value) = False Or IsNumeric(Me.Thresh.Value) = False Or IsNumeric(Me.MinMPP.Value) = False Or IsNumeric(Me.MaxMPP.Value) = False Or IsNumeric(Me.MinV.Value) = False Then
            MsgBox "The inverter has invalid values"
            isValid = False
        End If
    
        If Me.EffMultiPage.Value = 0 Then
            ' check if the single voltage curve is defined
            If (IsNumeric(Me.SingleMax.Value) = False Or IsNumeric(Me.SingleIN2.Value) = False Or IsNumeric(Me.SingleIN3.Value) = False Or IsNumeric(Me.SingleIN4.Value) = False Or IsNumeric(Me.SingleIN5.Value) = False Or IsNumeric(Me.SingleIN6.Value) = False Or IsNumeric(Me.SingleIN7.Value) = False Or IsNumeric(Me.SingleIN8.Value) = False) Then
                MsgBox "The efficiency curves have invalid values"
                isValid = False
            ElseIf (IsNumeric(Me.SingleEuro.Value) = False Or IsNumeric(Me.SingleEffic2.Value) = False Or IsNumeric(Me.SingleEffic3.Value) = False Or IsNumeric(Me.SingleEffic4.Value) = False Or IsNumeric(Me.SingleEffic5.Value) = False Or IsNumeric(Me.SingleEffic6.Value) = False Or IsNumeric(Me.SingleEffic7.Value) = False Or IsNumeric(Me.SingleEffic8.Value) = False) Then
                MsgBox "The efficiency curves have invalid values"
                isValid = False
            End If
        ElseIf Me.EffMultiPage.Value = 1 Then
            ' if multiple curves are chosen
    
            ' check if the low volatge curve is defined
            If (IsNumeric(Me.LowVolt.Value) = False Or IsNumeric(Me.LowIN2.Value) = False Or IsNumeric(Me.LowIN3.Value) = False Or IsNumeric(Me.LowIN4.Value) = False Or IsNumeric(Me.LowIN5.Value) = False Or IsNumeric(Me.LowIN6.Value) = False Or IsNumeric(Me.LowIN7.Value) = False Or IsNumeric(Me.LowIN8.Value) = False) Then
                MsgBox "The efficiency curves for low voltage have invalid values"
                isValid = False
            ElseIf (IsNumeric(Me.LowEffic2.Value) = False Or IsNumeric(Me.LowEffic3.Value) = False Or IsNumeric(Me.LowEffic4.Value) = False Or IsNumeric(Me.LowEffic5.Value) = False Or IsNumeric(Me.LowEffic6.Value) = False Or IsNumeric(Me.LowEffic7.Value) = False Or IsNumeric(Me.LowEffic8.Value) = False) Then
                MsgBox "The efficiency curves for low voltage have invalid values"
                isValid = False
            ElseIf (IsNumeric(Me.MedVolt.Value) = False Or IsNumeric(Me.MedIN2.Value) = False Or IsNumeric(Me.MedIN3.Value) = False Or IsNumeric(Me.MedIN4.Value) = False Or IsNumeric(Me.MedIN5.Value) = False Or IsNumeric(Me.MedIN6.Value) = False Or IsNumeric(Me.MedIN7.Value) = False Or IsNumeric(Me.MedIN8.Value) = False) Then
                ' check if the Medium voltage curve is defined
                MsgBox "The efficiency curves for Medium voltage have invalid values"
                isValid = False
            ElseIf (IsNumeric(Me.MedEffic2.Value) = False Or IsNumeric(Me.MedEffic3.Value) = False Or IsNumeric(Me.MedEffic4.Value) = False Or IsNumeric(Me.MedEffic5.Value) = False Or IsNumeric(Me.MedEffic6.Value) = False Or IsNumeric(Me.MedEffic7.Value) = False Or IsNumeric(Me.MedEffic8.Value) = False) Then
                MsgBox "The efficiency curves for IsNumeric(Medium voltage have invalid values"
                isValid = False
            ElseIf (IsNumeric(Me.HighVolt.Value) = False Or IsNumeric(Me.HighIN2.Value) = False Or IsNumeric(Me.HighIN3.Value) = False Or IsNumeric(Me.HighIN4.Value) = False Or IsNumeric(Me.HighIN5.Value) = False Or IsNumeric(Me.HighIN6.Value) = False Or IsNumeric(Me.HighIN7.Value) = False Or IsNumeric(Me.HighIN8.Value) = False) Then
                ' check if the high voltage curve is defined
                MsgBox " The efficiency curves for high voltage have invalid values"
                isValid = False
            ElseIf (IsNumeric(Me.HighEffic2.Value) = False Or IsNumeric(Me.HighEffic3.Value) = False Or IsNumeric(Me.HighEffic4.Value) = False Or IsNumeric(Me.HighEffic5.Value) = False Or IsNumeric(Me.HighEffic6.Value) = False Or IsNumeric(Me.HighEffic7.Value) = False Or IsNumeric(Me.HighEffic8.Value) = False) Then
                MsgBox " The efficiency curves for high voltage have invalid values"
                isValid = False
            End If
        End If
    End If
    
    If isValid Then
        If Not Me.Oper.ListIndex > -1 Or Not Me.BipolarIn.ListIndex > -1 Then
            MsgBox "The inverter has invalid values selected in the combo box"
        End If
    End If
    
    If isValid Then
       Call Add_Inverter
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

' EffMultiPage_Change Function
' This function is called when the efficiency curve page is changed
'
' The purpose of this function is to set the focus to the first field
' and if multiple curve is selected, set the multipage to the first
' page
Private Sub EffMultiPage_Change()
    If EffMultiPage.Value = 0 Then
        Me.SingleMax.SetFocus
    Else
        Me.VoltMulti.Value = 0
        Me.LowVolt.SetFocus
    End If
End Sub

' UserForm_Initialize
' This function is called when the UserForm is opened.
'
' The purpose of this function is to intialize
' the contents of the ComboBoxes , set the initial
' page in the efficiency curve multipage and set the
' focus to the first cell
Private Sub UserForm_Initialize()
    Me.EffMultiPage.Value = 0
    Me.Model.SetFocus
    
    Me.BipolarIn.AddItem "Yes"
    Me.BipolarIn.AddItem "No"
    
    Me.Oper.AddItem "MPPT"
    Me.Oper.AddItem "DCDC"
End Sub

' VoltMulti_Change Function
' This function is called when the voltage page is changed
'
' The purpose of this function is to set the focus to the
' first field of the page
Private Sub VoltMulti_Change()
    If VoltMulti.Value = 0 Then
        Me.LowVolt.SetFocus
    ElseIf VoltMulti.Value = 1 Then
        Me.MedVolt.SetFocus
    ElseIf VoltMulti.Value = 2 Then
        Me.HighVolt.SetFocus
    End If
End Sub
