Attribute VB_Name = "SaveModule"
Option Explicit

' SaveXML Function
'
' The purpose of this function is to save the site
' definitions on each of the worksheets to an XML file

Sub SaveXML(Optional ByVal saveTemp As Boolean)

    Dim newdoc As DOMDocument60 ' The XML Document
 
    Dim rootElement As IXMLDOMElement ' The root element of the XML file
    Dim AstmRegress As IXMLDOMElement ' Element used for Astm information in XML file
    Dim AstmCoeffs As IXMLDOMElement ' Element used for Astm factors in XML file
    Dim Eaf As IXMLDOMElement ' Element used for Eaf element in XML file
    Dim infoElement As IXMLDOMElement ' Element Used for site information for the XML file
    Dim orientElement As IXMLDOMElement ' Element used for orientation and shading info
    Dim subElement As IXMLDOMElement ' Element used for the sub arrays in the system
    Dim systemDCElement As IXMLDOMElement ' Element used for the DC Power of the System
    Dim systemACElement As IXMLDOMElement ' Element used for the AC Power of the System
    Dim pvElement As IXMLDOMElement ' Element used for the PV module data
    Dim invElement As IXMLDOMElement ' Element used for the inverter data
    Dim iamElement As IXMLDOMElement
    Dim effElement As IXMLDOMElement ' Element used for efficiency data
    Dim voltElement As IXMLDOMElement ' Element used for voltage data
    Dim transElement As IXMLDOMElement ' Element used for transformer data
    Dim lossElement As IXMLDOMElement ' Element used for loss data
    Dim losstypeElement As IXMLDOMElement ' Element used to hold soiling loss data
    Dim inputElement As IXMLDOMElement ' Element used for defining input file style
    Dim outputElement As IXMLDOMElement ' Element used for defining output file style
    'Dim Iterations As IXMLDOMElement ' Element used for holding iterative mode parameters   '--------Commenting out Iterative Functionality for this version--------'
    'Dim Iteration1 As IXMLDOMElement ' Element used for holding first iterative element     '--------Commenting out Iterative Functionality for this version--------'
    Dim element As IXMLDOMElement ' Element used to add data to other elements
   
    Dim attr As IXMLDOMAttribute ' Attributes used for the elements
    Dim text As IXMLDOMText ' Text values for each element
    Dim response As Integer
    'Counters
    Dim i As Integer
    Dim j As Integer

    Dim FSave As Variant ' Holds the file path to be saved to
    Dim FoundPath As Boolean ' Specifies if save path was found successfully
    
    'Create document
    Set newdoc = New DOMDocument60
    
    If IntroSht.Range("ModeSelect").Value = "ASTM E2848 Regression" Then
    
        Set rootElement = newdoc.createElement("Site")
        newdoc.appendChild rootElement
        '-->Add Site Info
        Call Add_Element(newdoc, rootElement, infoElement, IntroSht.Range("Version"))
        Call Add_Element(newdoc, rootElement, infoElement, IntroSht.Range("ModeSelect"))
        Call Add_Element(newdoc, rootElement, infoElement, _
        Range("Name,Country,Region,City"))
        
        '-->Add ASTM E2848 Regression Info
        Set AstmRegress = newdoc.createElement("ASTMRegress")
        rootElement.appendChild AstmRegress
        Call Add_Element(newdoc, AstmRegress, element, AstmSht.Range("SystemPmax"))
        
        '-->ASTM E2848 Regression Factors
        Set AstmCoeffs = newdoc.createElement("ASTMCoeffs")
        AstmRegress.appendChild AstmCoeffs
        ' Set all null regression factors to zero
        Dim coeff As Range
        For Each coeff In Range("AstmCoeffs")
            If IsEmpty(coeff.Value) = True Then
                coeff.Value = 0
            End If
        Next
        Call Add_Element(newdoc, AstmCoeffs, element, AstmSht.Range("ASTMCoeffs"))
        
        '-->Add Empirical Adjustment Factors (EAF)
        Set Eaf = newdoc.createElement("EAF")
        AstmRegress.appendChild Eaf
        ' Set all null EAF values to one
        For Each coeff In Range("AstmMonthList")
            If IsEmpty(coeff.Value) = True Then
                coeff.Value = 1
            End If
        Next
        Call Add_Element(newdoc, Eaf, element, AstmSht.Range("AstmMonthList"))
        
       
    Else
        If SiteSht.Range("UseLocTime").Value = "Yes" And SiteSht.Range("TimeZone").Value = 0 And saveTemp = False Then
            response = MsgBox("Your Timezone (hours GMT) value on the Site page is zero." & " If this value is correct, press 'Yes' to continue.", vbYesNo Or vbInformation, "Save Verification")
        If response = vbNo Then Exit Sub
        End If
        'Create site root
        Set rootElement = newdoc.createElement("Site")
        newdoc.appendChild rootElement
       
        Call Add_Site_Info(newdoc, rootElement, infoElement)
        
        
        ' Add Orientation and Shading Info
        Set orientElement = newdoc.createElement("Orientation_and_Shading")
        rootElement.appendChild orientElement
        
        Set attr = newdoc.createAttribute("ArrayType")
        attr.NodeValue = Orientation_and_ShadingSht.Range("OrientType").Value
        orientElement.setAttributeNode attr
        
        
        Call Add_Orientation_Shading_Info(newdoc, orientElement, element)
       
        'Add System Info
        Set infoElement = newdoc.createElement("System")
        rootElement.appendChild infoElement
        
           
        Set attr = newdoc.createAttribute("TotalArrays")
        attr.NodeValue = SystemSht.Range("NumSubArray").Value
        infoElement.setAttributeNode attr
        
        Set systemACElement = newdoc.createElement("ACWiringLossAtSTC")
        Call Add_Element(newdoc, infoElement, systemACElement, SystemSht.Range("ACWiringLossAtSTC"))
        
        If IntroSht.Range("ModeSelect").Value = "Radiation Mode" Then
            Set systemDCElement = newdoc.createElement("SystemDC")
            Call Add_Element(newdoc, infoElement, systemDCElement, SystemSht.Range("SystemDC"))
        
            Set systemACElement = newdoc.createElement("SystemAC")
            Call Add_Element(newdoc, infoElement, systemACElement, SystemSht.Range("SystemAC"))
            
        Else
            Set systemDCElement = newdoc.createElement("SystemDC")
            Call Add_Element(newdoc, infoElement, systemDCElement, SystemSht.Range("SystemDC"))
        
            Set systemACElement = newdoc.createElement("SystemAC")
            Call Add_Element(newdoc, infoElement, systemACElement, SystemSht.Range("SystemAC"))
           
            For i = 1 To SystemSht.Range("NumSubArray").Value
                'Add SubArray Info
                Set subElement = newdoc.createElement("SubArray" & i)
                infoElement.appendChild subElement
            
                '--> Add PV Module Elements
                'Add PV Module Info
                Set pvElement = newdoc.createElement("PVModule")
                subElement.appendChild pvElement
                Call Add_PV_Info(newdoc, pvElement, element, iamElement, i)
            
                '-->Add Inverter Elements
                'Add Inverter Info
                Set invElement = newdoc.createElement("Inverter")
                subElement.appendChild invElement
                Call Add_Inv_Info(newdoc, invElement, element, effElement, voltElement, i)
            Next i
    
            '--> Add Transformer Elements
            Set transElement = newdoc.createElement("Transformer")
            infoElement.appendChild transElement
        
            'Add Transformer Losses Info And Value
            Call Add_Trans_Info(newdoc, transElement, element)
           
            '--> Add Losses Element
            Set lossElement = newdoc.createElement("Losses")
            infoElement.appendChild lossElement
        
    
            'Add Losses Info And Value
            Call Add_Loss_Info(newdoc, lossElement, element, losstypeElement)
        
        End If
        
    End If 'end of mode select
    
        '-->Add InputFilePath
        Call Add_Element(newdoc, rootElement, infoElement, InputFileSht.Range("InputFilePath"))
                 
        '-->Add InputFileStyle
        Set inputElement = newdoc.createElement("InputFileStyle")
        rootElement.appendChild inputElement
        Call Add_Input_Info(newdoc, inputElement, element)
        
        '-->Add OutputFilePath
        Call Add_Output_File_Path(newdoc, rootElement, infoElement)
        
        ' Creating output file style element
        Set outputElement = newdoc.createElement("OutputFileStyle")
        rootElement.appendChild outputElement
        
'--------Commenting out Iterative Functionality for this version--------'

'        ' If iterative mode is enabled, the output file path and output file style is read from the iterative sheet
'        If Worksheets("Iterative Mode").Visible = True Then
'            '-->Add Iterative Mode if enabled
'            Set Iterations = newdoc.createElement("Iterations")
'            rootElement.appendChild Iterations
'            Set Iteration1 = newdoc.createElement("Iteration1")
'            Iterations.appendChild Iteration1
'            Call Add_Iterative_Mode(newdoc, Iteration1, outputElement, element, attr)
'        Else
            '-->Add Output File Style
            Call Add_Output_Info(newdoc, outputElement, element)
'        End If
        
        Call Indent(newdoc, newdoc.ChildNodes.NextNode, 1)

        If saveTemp = True Then
            FSave = Application.ThisWorkbook.path & "\CASSYSTemp.csyd"
            newdoc.Save (FSave)
            Exit Sub
        End If
        
        If IntroSht.Range("ChooseSaveAs").Value = "True" Then ' If the user has selected Save As
            FSave = Application.GetSaveAsFilename(title:="Save As", FileFilter:="CASSYS Site file (*.csyx),*.csyx", InitialFileName:=Worksheets("Site").Range("Name").Value)
            If (FSave <> False) Then
                FoundPath = True
            Else
                FoundPath = False
            End If
        Else ' If the user has selected Save
            If IntroSht.Range("SaveFilePath").Value = vbNullString Then
                FSave = Application.GetSaveAsFilename(title:="Save As", FileFilter:="CASSYS Site file (*.csyx),*.csyx", InitialFileName:=Worksheets("Site").Range("Name").Value)
                If (FSave <> False) Then
                    FoundPath = True
                Else
                    FoundPath = False
                End If
            Else
                FSave = IntroSht.Range("SaveFilePath").Value
                FoundPath = True
            End If
        End If

    
        'Save as XML file
    If (FoundPath = True) Then
        newdoc.Save (FSave)
        IntroSht.Unprotect
        IntroSht.Range("SaveFilePath") = FSave
        IntroSht.Protect
    Else
    'Do Nothing
    End If

    Set newdoc = Nothing

    
End Sub

' Add_Site_Info Function
'
' Arguments:
' newDoc - The XML Document
' rootElement - The root element of the XML file
' infoElement - Element Used for site information for XML
'
' The purpose of this function is to add the site information
' to the XML document
Sub Add_Site_Info(ByRef newdoc As DOMDocument60, ByRef rootElement As IXMLDOMElement, ByRef infoElement As IXMLDOMElement)

    Dim albedoElement As IXMLDOMElement
    Dim element As IXMLDOMElement
    Dim attr As IXMLDOMAttribute
    Dim text As IXMLDOMText ' Text values for each element
    
    ' Add Version Element And Valye
    Call Add_Element(newdoc, rootElement, infoElement, IntroSht.Range("Version"))
    Call Add_Element(newdoc, rootElement, infoElement, IntroSht.Range("ModeSelect"))
    Call Add_Element(newdoc, rootElement, infoElement, _
    Range("Name,Country,Region,City,Latitude,Longitude,Altitude,TimeZone,UseLocTime,RefMer,TransEnum"))

    
    Set albedoElement = newdoc.createElement("Albedo")
    Set attr = newdoc.createAttribute("Frequency")
    attr.NodeValue = (SiteSht.Range("AlbFreqVal").Value)
    albedoElement.setAttributeNode attr
    
    rootElement.appendChild albedoElement
    
    Call Add_Albedo_Info(newdoc, albedoElement, element)
    
End Sub


' Add_Orientation_Shading_Info Function
'
' Arguments:
' newDoc - The XML Document
' orientElement - Element used for orientation and shading info
' element - Element used to add data to other elements
'
' The purpose of this function is to add the site information
' to the XML document
' NB: changing function to save all attributes for all system types. 28/01/2016
Sub Add_Orientation_Shading_Info(ByRef newdoc As DOMDocument60, ByRef orientElement As IXMLDOMElement, ByRef element As IXMLDOMElement)

    If Orientation_and_ShadingSht.Range("OrientType").Value = "Fixed Tilted Plane" Then
        ' Add Plane tilt and Azimuth
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("PlaneTiltFix,AzimuthFix"))
    
    ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Fixed Tilted Plane Seasonal Adjustment" Then
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("SeasonalAdjustmentParams,AzimuthSeasonal"))
    
    ' If the array type is Unlimited Rows add more variables
    ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Unlimited Rows" Then
        ' Add Plane Tilt
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("PlaneTilt,Azimuth"))
        
        ' Add Pitch
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("Pitch,CollBandWidth,TopInactive,BottomInactive,GAOR,ShadingLimit"))
        
        ' Cell Based Shading
        If Orientation_and_ShadingSht.Range("UseCellVal") = "Yes" Then
            Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("UseCellVal,StrInWid,CellSize,WidOfStr"))
        Else
            Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("UseCellVal"))
        End If
        ' Add number of rows in a block
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("RowsBlock"))
    
    ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Single Axis Elevation Tracking (E-W)" Then
        ' Add Azimuth
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("AxisTiltSAET,AxisAzimuthSAET"))
        
        'Add Tilt Limits
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("MinTiltSAET,MaxTiltSAET"))
        
        'Saving Shading parameters
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("RowsBlockSAET,PitchSAET,WActiveSAET"))
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("BacktrackOptSAET"))
        If Range("UseCellValSAET").Value = "Yes" Then
            Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("UseCellValSAET,StrInWidSAET,CellSizeSAET,WidOfStrSAET"))
        ElseIf Range("UseCellValSAET").Value = "No" Then
            Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("UseCellValSAET"))
        End If
        
    
    ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Single Axis Horizontal Tracking (N-S)" Then
        'Add Tilt and Azimuth
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("AxisTiltSAST,AxisAzimuthSAST"))
        
        'Add Rotation Limits
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("RotationMaxSAST"))
        
        'Saving shading parameters
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("RowsBlockSAST,PitchSAST,WActiveSAST"))
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("BacktrackOptSAST"))
        If Range("UseCellValSAST").Value = "Yes" Then
            Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("UseCellValSAST,StrInWidSAST,CellSizeSAST,WidOfStrSAST"))
        ElseIf Range("UseCellValSAST").Value = "No" Then
            Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("UseCellValSAST"))
        End If
        
    ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Tilt and Roll Tracking" Then
        ' Add Azimuth
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("AxisTiltTART,AxisAzimuthTART"))
        
        'Add Tilt Limits
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("RotationMinTART,RotationMaxTART"))
    
    ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Azimuth (Vertical Axis) Tracking" Then
        'Add Tilt
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("PlaneTiltAVAT"))
        
        'Add Azimuth Limits
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("AzimuthRefAVAT,MinAzimuthAVAT,MaxAzimuthAVAT"))
        
    ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Two Axis Tracking" Then
        'Add Tilt Limits
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("MinTiltTAXT,MaxTiltTAXT"))
        
        'Add Azimuth Limits
        Call Add_Element(newdoc, orientElement, element, Orientation_and_ShadingSht.Range("AzimuthRefTAXT,MinAzimuthTAXT,MaxAzimuthTAXT"))
    
    End If
    
    'Add Horizon information
    
    Call Add_Element(newdoc, orientElement, element, Horizon_ShadingSht.Range("DefineHorizonProfile"))
    If Range("DefineHorizonProfile").Value = "Yes" Then
        Dim i As Integer
        Dim j As Integer
        Dim incompCount As Integer
        Dim delCount As Integer
        delCount = 0
        For i = Range("NumHorPts").Value - 1 To 0 Step -1
            incompCount = 0
            ' Finding and deleting horizon values which are missing information
            If IsEmpty(Range("HAziFirst").Offset(i, 0)) = True Or IsEmpty(Range("HElevFirst").Offset(i, 0)) = True Then
                incompCount = incompCount + 1
            ' Finding and deleting duplicate horizon azimuth values
            Else
                For j = 0 To i - 1 Step 1
                    If Range("HAziFirst").Offset(i, 0).Value = Range("HAziFirst").Offset(j, 0).Value Then
                        If i <> j And IsEmpty(Horizon_ShadingSht.Range("C" & j)) = False Then
                            incompCount = incompCount + 1
                        End If
                    End If
                Next j
            End If
            If incompCount <> 0 Then
                delCount = delCount + 1
            End If
        Next i
        
        Call Add_Element(newdoc, orientElement, element, Horizon_ShadingSht.Range("HorizonAzi"))
        Call Add_Element(newdoc, orientElement, element, Horizon_ShadingSht.Range("HorizonElev"))
        If delCount > 0 Then
            MsgBox Prompt:="There were duplicate horizon azimuth values or azimuth and horizon elevation values not associated with an elevation or azimuth, respectively. The unassociated values were deleted, and the first of the duplicates was taken by CASSYS", Buttons:=vbExclamation, title:="Duplicate and Unassociated Horizon Values"
        End If
    End If
    
    
    
    
End Sub


' Add_PV_Info Function
'
' Arguments:
' newDoc - The XML Document
' pvElement - Element used for the PV module data
' element - Element used to add data to other elements
' i - counter used to determine which sub-array the PV Module is in
'
' The purpose of this function is to add the PV module information
' to the XML document
' NB: adjusted so module file name is added to xml document
Sub Add_PV_Info(ByRef newdoc As DOMDocument60, ByRef pvElement As IXMLDOMElement, ByRef element As IXMLDOMElement, ByRef iamElement As IXMLDOMElement, ByVal i As Integer)
    
    Dim j As Integer
    Dim auxOffset As Integer
    auxOffset = (i - 1) * SubArrayHeight
    
    'Set Model Info and Value
    Call AddArray_Element(newdoc, pvElement, element, "Module", SystemSht.Range("ModuleModel").Offset(auxOffset, 0).Value)
           
    'Set Manufacturer Info and Value
    Call AddArray_Element(newdoc, pvElement, element, "Manufacturer", SystemSht.Range("ModuleManu").Offset(auxOffset, 0).Value)
                      
    'Set Data Source Info and Value
    Call AddArray_Element(newdoc, pvElement, element, "Origin", PV_DatabaseSht.Cells(SystemSht.Range("PVDataIndex").Offset(auxOffset, 0).Value + PVDataHeight, 1).Value)

    'NB: Causes program to save PV Module's file name to the xml
    'Set File Name Info and Value
    Call AddArray_Element(newdoc, pvElement, element, "FileName", PV_DatabaseSht.Cells(SystemSht.Range("PVDataIndex").Offset(auxOffset, 0).Value + PVDataHeight, 4).Value)
    
    'Add PV Element Info and Value
    For j = 8 To 39
        Call AddArray_Element(newdoc, pvElement, element, PV_DatabaseSht.Cells(PVDataHeight - 2, j).Value, PV_DatabaseSht.Cells(SystemSht.Range("PVDataIndex").Offset(auxOffset, 0).Value + PVDataHeight, j).Value)
    Next j
    
    If SystemSht.Range("DefnAvailable").Offset(auxOffset, 0).Value = "Yes" And LossesSht.Range("UsePAN").Value = "Yes" Then
        Set iamElement = newdoc.createElement("IAMDefinition")
        pvElement.appendChild iamElement
        
        For j = 64 To 81
            If PV_DatabaseSht.Cells(SystemSht.Range("PVDataIndex").Offset(auxOffset, 0).Value + PVDataHeight, j).Value <> vbNullString Then
                Call AddArray_Element(newdoc, iamElement, element, PV_DatabaseSht.Cells(PVDataHeight - 2, j).Value, PV_DatabaseSht.Cells(SystemSht.Range("PVDataIndex").Offset(auxOffset, 0).Value + PVDataHeight, j).Value)
            End If
        Next j
    End If
    
    
    'Set Number of Modules Info and Value
    Call AddArray_Element(newdoc, pvElement, element, "NumModules", SystemSht.Range("NumMod").Offset(auxOffset, 0).Value)
           
    'Set Modules in a String Info and Value
    Call AddArray_Element(newdoc, pvElement, element, "ModulesInString", SystemSht.Range("ModStr").Offset(auxOffset, 0).Value)
           
    'Set Number of String Info and Value
    Call AddArray_Element(newdoc, pvElement, element, "NumStrings", SystemSht.Range("NumStr").Offset(auxOffset, 0).Value)
    
    'Set Loss Fraction Info and Value
    Call AddArray_Element(newdoc, pvElement, element, "LossFraction", SystemSht.Range("PVLossFrac").Offset(auxOffset, 0).Value)
    
    'Set Global Wiring Resistance Info and Value
    Call AddArray_Element(newdoc, pvElement, element, "GlobWireResist", SystemSht.Range("GWR").Offset(auxOffset, 0).Value)
    
End Sub

' Add_Trans_Info Function
'
' Arguments:
' newDoc - The XML Document
' transElement - Element used for the transformer data
' element - Element used to add data to other elements
'
' The purpose of this function is to add the transformer information
' to the XML document
Sub Add_Trans_Info(ByRef newdoc As DOMDocument60, ByRef transElement As IXMLDOMElement, ByRef element As IXMLDOMElement)
    ' Add the Nominal Power of the transformer
    Call Add_Element(newdoc, transElement, element, TransformerSht.Range("PNomTrf,PIronLossTrf,PFullLoadLss,PResLssTrf,ACCapSTC,NightlyDisconnect"))
End Sub


' Add_Inv_Info Function
'
' Arguments:
' newDoc - The XML Document
' invElement - Element used for the inverter data
' element - Element used to add data to other elements
' effElement - Element used for efficiency data
' voltElement - Element used for volatage data
' i - counter used to determine which sub-array the PV Module is in
'
' The purpose of this function is to add the inverter information
' to the XML document
Sub Add_Inv_Info(ByRef newdoc As DOMDocument60, ByRef invElement As IXMLDOMElement, ByRef element As IXMLDOMElement, ByRef effElement As IXMLDOMElement, ByRef voltElement As IXMLDOMElement, ByVal i As Integer)

    Dim j As Integer ' counter
    Dim auxOffset As Integer
    auxOffset = (i - 1) * SubArrayHeight
    Dim attr As IXMLDOMAttribute 'Attributes used for the elements
    Dim text As IXMLDOMText ' Text values for each element
    
    'Set Model Info and Value
    Call AddArray_Element(newdoc, invElement, element, "Model", SystemSht.Range("InverterList").Offset(auxOffset, 0).Value)
           
    'Set Manufacturer Info and Value
    Call AddArray_Element(newdoc, invElement, element, "Manufacturer", SystemSht.Range("InverterManu").Offset(auxOffset, 0).Value)
           
    'Set Data Source Info and Value
    Call AddArray_Element(newdoc, invElement, element, "Origin", Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, 1).Value)
    
    'NB: adding file name to inverter data
    'Set File Name Info and Value
    Call AddArray_Element(newdoc, invElement, element, "FileName", SystemSht.Range("InverterFileName").Offset(auxOffset, 0).Value)
       
    ' Add Info display on system page
    For j = 8 To 14
       
       If (Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, j).Value = vbNullString) Then
           Call AddArray_Element(newdoc, invElement, element, Inverter_DatabaseSht.Cells(InvDataHeight, j).Value, 0)
       Else
           Call AddArray_Element(newdoc, invElement, element, Inverter_DatabaseSht.Cells(InvDataHeight, j).Value, Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, j).Value)
       End If
    Next j
    'Add Threshold Info
    Call AddArray_Element(newdoc, invElement, element, "Threshold", Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, 18).Value)
    
    'Add MinMPP Info
    Call AddArray_Element(newdoc, invElement, element, "MinMPP", Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, 20).Value)
    
    'Add MaxMPP Info
    Call AddArray_Element(newdoc, invElement, element, "MaxMPP", Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, 21).Value)
    
    'Add Max.V Info
    Call AddArray_Element(newdoc, invElement, element, "Max.V", Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, 22).Value)
    
    'Add Min.V Info
    Call AddArray_Element(newdoc, invElement, element, "Min.V", Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, 23).Value)
    
    'Add Oper. Info
    Call AddArray_Element(newdoc, invElement, element, "Oper.", Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, 27).Value)
           
    'Set Number of Inverters Info and Value
    Call AddArray_Element(newdoc, invElement, element, "NumInverters", SystemSht.Range("NumInv").Offset(auxOffset, 0).Value)
           
    'Add Multi-Curve Info
    
    If (Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, 40).Value) = "X" Then
        Call AddArray_Element(newdoc, invElement, element, "MultiCurve", "True")
    Else
        Call AddArray_Element(newdoc, invElement, element, "MultiCurve", "False")
    End If
    
    'Add efficiency Info
    Set effElement = newdoc.createElement("Efficiency")
    invElement.appendChild effElement
      
    If (Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, 40).Value) = "X" Then
        'Add low voltage info
        Set voltElement = newdoc.createElement("Low")
        effElement.appendChild voltElement
        
        
        Set attr = newdoc.createAttribute("Voltage")
        attr.NodeValue = Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, 76).Value
        voltElement.setAttributeNode attr
           
        'Add low voltage efficiencies
        For j = 77 To 92
            If Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, j).Value <> vbNullString Then
                Call AddArray_Element(newdoc, voltElement, element, Inverter_DatabaseSht.Cells(InvDataHeight, j).Value, Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, j).Value)
            End If
        Next j
          
        Set text = newdoc.createTextNode(Constants.vbCrLf)
        effElement.appendChild text
        
        'Add Medium voltage info
        Set voltElement = newdoc.createElement("Med")
        effElement.appendChild voltElement
        
           
        Set attr = newdoc.createAttribute("Voltage")
        attr.NodeValue = Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, 93).Value
        voltElement.setAttributeNode attr
           
        'Add medium voltage info
        For j = 94 To 109
            If Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, j).Value <> vbNullString Then
                Call AddArray_Element(newdoc, voltElement, element, Inverter_DatabaseSht.Cells(InvDataHeight, j).Value, Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, j).Value)
            End If
        Next j
        
        'Add High voltage info
        Set voltElement = newdoc.createElement("High")
        effElement.appendChild voltElement
        
        Set attr = newdoc.createAttribute("Voltage")
        attr.NodeValue = Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, 110).Value
        voltElement.setAttributeNode attr
           
        'Add high voltage info
        For j = 111 To 126
            If Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, j).Value <> vbNullString Then
                Call AddArray_Element(newdoc, voltElement, element, Inverter_DatabaseSht.Cells(InvDataHeight, j).Value, Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, j).Value)
            End If
        Next j
    Else
        ' Add Single curve Voltage Info
        Set voltElement = newdoc.createElement("EffCurve")
        effElement.appendChild voltElement
         
        ' add single curve efficiences
        For j = 42 To 58
            If Not j = 44 Then
                Call AddArray_Element(newdoc, voltElement, element, Inverter_DatabaseSht.Cells(InvDataHeight, j).Value, Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, j).Value)
            End If
        Next j
    End If
    
    'Add Bipolar Input Info
    Call AddArray_Element(newdoc, invElement, element, "BipolarInput", Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, 72).Value)
    
    'Add Loss Fraction Info
    Call AddArray_Element(newdoc, invElement, element, "LossFraction", SystemSht.Range("InvLossFrac").Offset(auxOffset, 0).Value)
    
End Sub

' Add_Loss_Info Function
'
' Arguments:
' newDoc - The XML Document
' lossElement - Element used for loss data
' element - Element used to add data to other elements
' losstypeElement - Element used to hold soiling loss data
'
' The purpose of this function is to add the losses information
' to the XML file

Sub Add_Loss_Info(ByRef newdoc As DOMDocument60, ByRef lossElement As IXMLDOMElement, ByRef element As IXMLDOMElement, ByRef losstypeElement As IXMLDOMElement)
    Dim attr As IXMLDOMAttribute 'Attributes used for the elements
    Dim IAMCell As Range
    Dim AOICell As Range
    Dim text As IXMLDOMText ' Text values for each element
    
    ' Add Thermal losses info
    Set losstypeElement = newdoc.createElement("ThermalLosses")
     
    lossElement.appendChild losstypeElement
    
    ' Add the info on whether or not to use measured values
    Call Add_Element(newdoc, losstypeElement, element, LossesSht.Range("UseMeasuredValues"))
    ' Add the Constant Heat Loss Factor
    
    If (LossesSht.Range("UseMeasuredValues").Value = False) Then
        Call Add_Element(newdoc, losstypeElement, element, LossesSht.Range("ConsHLF,ConvHLF"))
    End If
    
    ' Add module quality info
    Set losstypeElement = newdoc.createElement("ModuleQualityLosses")
    lossElement.appendChild losstypeElement

    ' Add Efficiency loss info
    Call Add_Element(newdoc, losstypeElement, element, LossesSht.Range("EfficiencyLoss"))
    
    'Add the Module mismatch losses info
    Set losstypeElement = newdoc.createElement("ModuleMismatchLosses")

    lossElement.appendChild losstypeElement
    
    ' Add the power loss
    Call Add_Element(newdoc, losstypeElement, element, LossesSht.Range("PowerLoss"))
    
    ' Add the losses at a fixed voltage
    Call Add_Element(newdoc, losstypeElement, element, LossesSht.Range("LossFixedVoltage"))
    
    ' Add the Incidence Angle Modifier
    Set losstypeElement = newdoc.createElement("IncidenceAngleModifier")
    
    ' Specify how the IAM is defined, using an attribute: ASHRAE or User Defined
    Set attr = newdoc.createAttribute("IAMSelection")
    If LossesSht.Range("IAMSelection").Value = "ASHRAE" Then
        attr.NodeValue = "ASHRAE"
    ElseIf LossesSht.Range("IAMSelection").Value = "User Defined" Then
        attr.NodeValue = "User Defined"
    End If
    losstypeElement.setAttributeNode attr
    
    ' Append the IncidenceAngleModifier node to the losses node
    lossElement.appendChild losstypeElement
    
    If LossesSht.Range("IAMSelection").Value = "ASHRAE" Then
        'Add the bNaught value
        Call Add_Element(newdoc, losstypeElement, element, LossesSht.Range("bNaught"))
    Else
        ' Check if IAM 0 and IAM 90 are defined, if not then set them to default values
        If LossesSht.Range("IAM_0").Value = vbNullString Then LossesSht.Range("IAM_0").Value = 1
        If LossesSht.Range("IAM_90").Value = vbNullString Then LossesSht.Range("IAM_90").Value = 0
        ' Add the user defined values (blank values are ignored and not saved as nodes)
        For Each IAMCell In LossesSht.Range("IAMRange")
            If IAMCell.Value <> vbNullString Then Call Add_Element(newdoc, losstypeElement, element, IAMCell)
        Next
    End If
    
    ' Add the soiling losses info
    Set losstypeElement = newdoc.createElement("SoilingLosses")
    lossElement.appendChild losstypeElement
    
    'If the selected frequency is yearly
    If SoilingSht.Range("SIndex") = "1" Then
        ' Set the soiling losses frequency attribute to yearly
        Set attr = newdoc.createAttribute("Frequency")
        attr.NodeValue = "Yearly"
        losstypeElement.setAttributeNode attr
    
        ' Add the yearly percent loss
        Call Add_Element(newdoc, losstypeElement, element, SoilingSht.Range("Yearly"))
    ElseIf SoilingSht.Range("SIndex") = "2" Then
        'If the selected frequency is monthly
        
        ' Set the soiling losses frequency attribute to monthly
        Set attr = newdoc.createAttribute("Frequency")
        attr.NodeValue = "Monthly"
        losstypeElement.setAttributeNode attr
        
        'Add the percent loss for each month (in order of course)
        Call Add_Element(newdoc, losstypeElement, element, SoilingSht.Range("Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec"))
    End If
    
End Sub


' Add_Input_Info Function
'
' Arguments:
' newDoc - The XML Document
' inputElement - Element used for defining input file style
' element - Element used to add data to other elements
'
' The purpose of this function is to add the input file format
' information to the XML document
Sub Add_Input_Info(ByRef newdoc As DOMDocument60, ByRef inputElement As IXMLDOMElement, ByRef element As IXMLDOMElement)
    If IntroSht.Range("ModeSelect") = "ASTM E2848 Regression" Then
        Call Add_Element(newdoc, inputElement, element, _
        InputFileSht.Range("RowsToSkip,Delimeter,TimeFormat,AveragedAt,Interval,TMYType, InputColumnNums, IncorrectClimateRowsAllowed"))
    Else
        Call Add_Element(newdoc, inputElement, element, _
        InputFileSht.Range("RowsToSkip,Delimeter,TimeFormat,AveragedAt,Interval,TMYType,MeterTilt,MeterAzimuth, InputColumnNums, IncorrectClimateRowsAllowed"))
    End If
End Sub

' Add_Output_Info Function
'
' Arguments:
' newDoc - The XML Document
' rootElement - Parent element
' infoElement - Element used to store output file path
'
' The purpose of this function is to add the output file path to the XML document
Sub Add_Output_File_Path(ByRef newdoc As DOMDocument60, ByRef rootElement As IXMLDOMElement, ByRef infoElement As IXMLDOMElement)

'--------Commenting out Iterative Functionality for this version--------'

'    ' Determine if in iterative mode
'    If Worksheets("Iterative Mode").Visible = True Then
'        Call Add_Element(newdoc, rootElement, infoElement, IterativeSht.Range("OutputFilePath"))
'    Else
        Call Add_Element(newdoc, rootElement, infoElement, OutputFileSht.Range("OutputFilePath"))
'    End If
End Sub

' Add_Output_Info Function
'
' Arguments:
' newDoc - The XML Document
' outputElement - Element used for defining output file style
' element - Element used to add data to other elements
'
' The purpose of this function is to add the output file format
' information to the XML document
Sub Add_Output_Info(ByRef newdoc As DOMDocument60, ByRef outputElement As IXMLDOMElement, ByRef element As IXMLDOMElement)
    Dim outParam As Range
    Call Add_Element(newdoc, outputElement, element, OutputFileSht.Range("OutputParam"), True)
End Sub

Sub Add_Albedo_Info(ByRef newdoc As DOMDocument60, ByRef albedoElement As IXMLDOMElement, ByRef element As IXMLDOMElement)

    ' If albedo is set to yearly add yearly albedo values
    Dim monthlyMax As Integer
    Dim monthlyMin As Integer
    monthlyMax = 16
    monthlyMin = 5
    Dim i As Integer
    If (SiteSht.Range("AlbFreqVal").Value = "Yearly") Then
        Call AddArray_Element(newdoc, albedoElement, element, SiteSht.Range("AlbFreqVal"), SiteSht.Range("AlbYearly"))
    ElseIf (SiteSht.Range("AlbFreqVal").Value = "Monthly") Then
        ' If albedo is set to monthly add monthly albedo values
        For i = monthlyMin To monthlyMax
            Call AddArray_Element(newdoc, albedoElement, element, SiteSht.Cells(29, i).Value, SiteSht.Range("AlbJan,AlbFeb,AlbMar,AlbApr,AlbMay,AlbJun,AlbJul,AlbAug,AlbSep,AlbOct,AlbNov,AlbDec"))
        Next i
    End If
    
End Sub

'--------Commenting out Iterative Functionality for this version--------'

' Add_Iterative_Mode Function
'
' Arguments:
' newDoc - The XML Document
' Iteration1 - Element used to store iteration attributes (start, end, interval)
' Iterations - parent element to Iteration1
' outputElement - Element used for defining output file style
' element - Element used to add data to other elements
'
' The purpose of this function is to add the Iterative mode parameters
' information to the XML document
Sub Add_Iterative_Mode(ByRef newdoc As DOMDocument60, ByRef Iteration1 As IXMLDOMElement, ByRef outputElement As IXMLDOMElement, ByRef element As IXMLDOMElement, ByRef attr As IXMLDOMAttribute)
'    Dim outParam As Range

'    ' Add output file style
'    ' Set output to Summarize and write to csyx file
'    Application.ScreenUpdating = False
'    OutputFileSht.Range("E" & Range("IterOutputRowNum").Value).Value = "Summarize"
'    Call Add_Element(newdoc, outputElement, element, OutputFileSht.Range("E" & Range("IterOutputRowNum").Value), True)
'
'
'    ' Add Iterative node attributes
'    Set attr = newdoc.createAttribute("ParamPath")
'    attr.NodeValue = IterativeSht.Range("ParamPath").Value
'    Iteration1.setAttributeNode attr
'
'    Set attr = newdoc.createAttribute("Start")
'    attr.NodeValue = IterativeSht.Range("Start").Value
'    Iteration1.setAttributeNode attr
'
'    Set attr = newdoc.createAttribute("End")
'    attr.NodeValue = IterativeSht.Range("End").Value
'    Iteration1.setAttributeNode attr
'
'    Set attr = newdoc.createAttribute("Interval")
'    attr.NodeValue = IterativeSht.Range("Interval").Value
'    Iteration1.setAttributeNode attr
    End Sub

' recursively indents and adds newlines to the xml file
Sub Indent(ByRef newdoc As DOMDocument60, ByRef parent As IXMLDOMNode, ByVal length As Integer)
    Dim text As IXMLDOMText 'Text values for each element
    Dim child As IXMLDOMNode
    
    If parent.HasChildNodes Then
        ' go through the parent nodes children
        For Each child In parent.ChildNodes
            If child.NodeType = NODE_ELEMENT Then
                ' add newline
                Set text = newdoc.createTextNode(Constants.vbNewLine)
                parent.InsertBefore text, child
            
                ' add indents
                Set text = newdoc.createTextNode(String(length, Constants.vbTab))
                parent.InsertBefore text, child
                
                ' add newline and indents after last node
                If child.NextSibling Is Nothing Then
                   Set text = newdoc.createTextNode(Constants.vbNewLine)
                   parent.appendChild text
                   
                   Set text = newdoc.createTextNode(String(length - 1, Constants.vbTab))
                   parent.appendChild text
                End If
            End If
            
            ' if the child node has children call this function
            If child.HasChildNodes Then
                Call Indent(newdoc, child, length + 1)
            End If
        Next
    End If
    
End Sub

' Add_Element Function
'
' Arguments:
' newDoc - The XML Document
' Parent - the Parent node
' Child - the Child node
' Name - the name of the element to be added
' Value - the text value of the element to be added
'
' The purpose of this function is to append the child node
' to its parent and add text to that same child node
Sub Add_Element(ByRef newdoc As DOMDocument60, ByRef parent As IXMLDOMElement, ByRef child As IXMLDOMElement, ByRef rngToSave, Optional ByVal saveOutput As Boolean)

    Dim attr As IXMLDOMAttribute
    Dim text As IXMLDOMText 'Text values for each element
    Dim cell As Range
    Dim cellName As String

    For Each cell In rngToSave
        cellName = ExtractCellName(cell)
        ' Create the node and append it to its parent
        Set child = newdoc.createElement(cellName)
        
        
        If saveOutput = True Then
            ' NB: moved If cell.Value = "Summarize"... to front, so only ouputs that were marked to summarize or detail are saved to the .csyx file
            ' NB: Added "cell.EntireRow.Hidden" condition so outputs only save if relevant to the selected mode
            If cell.EntireRow.Hidden = False And cell.Value = "Summarize" Or cell.EntireRow.Hidden = False And cell.Value = "Detail" Then
                Set text = newdoc.createTextNode("True")
                child.appendChild text
                
                
                If cell.Offset(0, OutputFileSht.Range("UnitsColumn").Column - OutputFileSht.Range("OutputParam").Column).Value <> vbNullString Then
                    Set attr = newdoc.createAttribute("Units")
                    attr.NodeValue = Replace(cell.Offset(0, OutputFileSht.Range("UnitsColumn").Column - OutputFileSht.Range("OutputParam").Column).Value, "°", "deg.")
                    child.setAttributeNode attr
                Else
                    Set attr = newdoc.createAttribute("Units")
                    child.setAttributeNode attr
                End If
            
                Set attr = newdoc.createAttribute("DisplayName")
                attr.NodeValue = cell.Offset(0, OutputFileSht.Range("HeaderRow").Column - OutputFileSht.Range("OutputParam").Column).Value
                child.setAttributeNode attr
            
                Set attr = newdoc.createAttribute("Summarize")
            
                If cell.Value = "Summarize" Then
                    attr.NodeValue = "True"
                Else
                    attr.NodeValue = "False"
                End If
            
                child.setAttributeNode attr
            
                parent.appendChild child
            
            Else
                Set text = newdoc.createTextNode("False")
                child.appendChild text
            End If
        Else
            parent.appendChild child
            
            ' Create the text for the node and append it
            If (parent.nodeName = "InputFileStyle" And cell.Value2 = vbNullString) Then
                If child.nodeName <> "TMYType" Then
                    Set text = newdoc.createTextNode("N/A")
                    child.appendChild text
                Else
                    Set text = newdoc.createTextNode(0)
                    child.appendChild text
                End If
            ElseIf cell.Address = Range("InputFilePath").Address Then
                'Replace \ with / so C# will accept the path
                Set text = newdoc.createTextNode(Replace(Range("InputFilePath").Value, "\", "/"))
                child.appendChild text
            ElseIf cell.Address = Range("OutputFilePath").Address Then
                Set text = newdoc.createTextNode(Replace(Range("OutputFilePath").Value, "\", "/"))
                child.appendChild text
            ElseIf cell.Value = "Yes" Then
                Set text = newdoc.createTextNode("True")
                child.appendChild text
            ElseIf cell.Value = "No" Then
                Set text = newdoc.createTextNode("False")
                child.appendChild text
            ElseIf cell.Address = Range("ACWiringLossAtSTC").Address Then
                If (Range("ACWiringLossAtSTC").Value = "at STC") Then
                    Set text = newdoc.createTextNode("True")
                Else
                    Set text = newdoc.createTextNode("False")
                End If
                child.appendChild text
                
'--------Commenting out Iterative Functionality for this version--------'
'
'            ElseIf cell.Name.Name = "IterativeOutputParam" Then
'                Set text = newdoc.createTextNode("True")
'                child.appendChild text
'
'                ' Setting timestamp attribute
'                Set attr = newdoc.createAttribute("Units")
'                attr.NodeValue = IterativeSht.Range("X" & Application.WorksheetFunction.Match(IterativeSht.Range("IterativeOutputParam").Value, IterativeSht.Range("W:W"), 0)).Value
'                child.setAttributeNode attr
'
'                ' Setting summary attribute
'                Set attr = newdoc.createAttribute("Summarize")
'                attr.NodeValue = "True"
'                child.setAttributeNode attr
'
'                parent.appendChild child
            Else
                Set text = newdoc.createTextNode(cell.Value)
                child.appendChild text
                
            End If
        End If
    Next

End Sub

' AddArray_Element Function
'
' Arguments:
' newDoc - The XML Document
' Parent - the Parent node
' Child - the Child node
' Name - the name of the element to be added
' Value - the text value of the element to be added
'
' The purpose of this function is to append the child node
' to its parent and add text to that same child node

Sub AddArray_Element(ByRef newdoc As DOMDocument60, ByRef parent As IXMLDOMElement, ByRef child As IXMLDOMElement, ByVal Name As String, ByVal Value As String)
    Dim text As IXMLDOMText 'Text values for each element

    ' Create the node and append it to its parent
    Set child = newdoc.createElement(Name)
    parent.appendChild child

    ' Create the text for the node and append it
    If (parent.nodeName = "InputFileStyle" And Value = vbNullString) Then
        Set text = newdoc.createTextNode("N/A")
        child.appendChild text
    Else
        Set text = newdoc.createTextNode(Value)
        child.appendChild text
    End If

End Sub







