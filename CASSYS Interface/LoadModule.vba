Attribute VB_Name = "LoadModule"
'           LOAD MODULE         '
'-------------------------------'
' The load module's purpose is to let the user browse for a  '
' file to load. It contains various versions of
Option Explicit

Enum LOAD_STATUS
    ItemNotFound = 0
    ItemFound = 1
    filePathFound = True
    filePathNotFound = False
End Enum

' GetFileToLoad Function
'
' The purpose of this function is to get the file path of the XML file to be loaded
' and insert it into a cell on the intro sheet

Function GetFileToLoad()

    Dim browsePath As Variant
    Dim currentShtStatus As sheetStatus
    Dim doNotSave As Integer
    
    ' Check if the user wants to save the existing CSYX file before moving to another one (Do not save by default to allow loading first time when workbook opens)
    doNotSave = vbYes
    If Not IntroSht.Range("SaveFilePath").Value = vbNullString Then
        doNotSave = MsgBox("Are you sure you want to load a new file? Click Yes to proceed without saving the existing file, and No to save the file", vbYesNo + vbQuestion, "Save Reminder")
    End If
    
    ' If yes, then open browswer window to find the file
    If doNotSave = vbYes Then
        ' The browse box begins in the same directory as CASSYS.xlsm
        ChDir Application.ThisWorkbook.path
        browsePath = Application.GetOpenFilename(Title:=" CASSYS: Please choose a Site definition to load", FileFilter:="CASSYS Site file (*.csyx),*.csyx")
            
        If (browsePath <> False) Then
            Call PreModify(IntroSht, currentShtStatus)
            IntroSht.Range("LoadFilePath").Value = browsePath
            IntroSht.Range("SaveFilePath").Value = browsePath
            Call PostModify(IntroSht, currentShtStatus)
            GetFileToLoad = True
        Else
            GetFileToLoad = False
        End If
    Else
        ' If no, Save the existing document
        Call SaveXML
    End If
    
End Function


' Load Function
'
' The purpose of this function is to load data from
' the CSYX site file into their respective
' sheets

Function Load() As Boolean
   
    
    Dim newdoc As DOMDocument60
    Set newdoc = New DOMDocument60
    newdoc.validateOnParse = False
    Dim currentShtStatus As sheetStatus
    Dim errorShtStatus As sheetStatus
    
    ' Get the load file path and load the DOMDocument60
    newdoc.Load IntroSht.Range("LoadFilePath").Value
    
    Call PreModify(ErrorSht, errorShtStatus)
    
    ' Set ModeSelect
    ' ModeSelect indicates whether the file is a full system simulation, just for radiation calculation, or ASTM 2848 Regression
    Call PreModify(IntroSht, currentShtStatus)
    Application.EnableEvents = True
            
    ' Starting with v 1.0.0, ModeSelect is stored in the csyx file
    If Not newdoc.SelectSingleNode("//Site/ModeSelect") Is Nothing Then
        IntroSht.Range("ModeSelect").Value = newdoc.SelectSingleNode("//Site/ModeSelect").text

    ' In previous versions, the mode is determined by System DC and System AC sizes
    ' DC and AC sizes both zero means the file is in radiation calculation mode
    ElseIf Not newdoc.SelectSingleNode("//Site/System/SystemAC") Is Nothing And Not newdoc.SelectSingleNode("//Site/System/SystemDC") Is Nothing Then
        If newdoc.SelectSingleNode("//Site/System/SystemAC").text = 0 And newdoc.SelectSingleNode("//Site/System/SystemDC").text = 0 Then
            IntroSht.Range("ModeSelect").Value = "Radiation Mode"
        Else
            IntroSht.Range("ModeSelect").Value = "Grid-Connected System"
        End If
    Else
        IntroSht.Range("ModeSelect").Value = "Grid-Connected System"
    End If
    
    Application.EnableEvents = False
    Call PostModify(IntroSht, currentShtStatus)
    
    If IntroSht.Range("ModeSelect") = "ASTM E2848 Regression" Then 'Check if in ASTM E2848
        
        ' Add ASTM System Information
        Call PreModify(AstmSht, currentShtStatus)
        Call loadASTMSht(newdoc)
        Call PostModify(AstmSht, currentShtStatus)

    Else 'Grid-connected System or Radiation Mode
     
        ' Add Orientation And Shading Values
        Call PreModify(Orientation_and_ShadingSht, currentShtStatus)
        Call loadOSSht(newdoc)
        Call PostModify(Orientation_and_ShadingSht, currentShtStatus)
        
        Call PreModify(BifacialSht, currentShtStatus)
        Call loadBifacialSht(newdoc)
        Call PostModify(BifacialSht, currentShtStatus)
        
        Call PreModify(Horizon_ShadingSht, currentShtStatus)
        Call loadHorizon_ShadingSht(newdoc)
        Call PostModify(Horizon_ShadingSht, currentShtStatus)
            
        ' Add Losses Information (only if in grid-connected mode)
        If IntroSht.Range("ModeSelect").Value = "Grid-Connected System" Then
            
            ' Add System And Array Values
            Call PreModify(SystemSht, currentShtStatus)
            Call loadSystemSht(newdoc)
            Call PostModify(SystemSht, currentShtStatus)
            
            Call PreModify(LossesSht, currentShtStatus)
            Call loadLossesSht(newdoc)
            Call PostModify(LossesSht, currentShtStatus)
        
            ' Add Soiling Losses Information
            Call PreModify(SoilingSht, currentShtStatus)
            Call loadSoilingSht(newdoc)
            Call PostModify(SoilingSht, currentShtStatus)
        
            ' Add Spectral Information
            Call PreModify(SpectralSht, currentShtStatus)
            Call loadSpectralSht(newdoc)
            Call PostModify(SpectralSht, currentShtStatus)
        
            ' Add Transformer Information
            Call PreModify(TransformerSht, currentShtStatus)
            Call loadTransformerSht(newdoc)
            Call PostModify(TransformerSht, currentShtStatus)
            
        End If
    End If
    ' End of mode dependent loading
     
    ' These sheets are loaded from csyx independent of mode
     
    ' Add Site Values
    Call PreModify(SiteSht, currentShtStatus)
    Call loadSiteSht(newdoc)
    Call PostModify(SiteSht, currentShtStatus)
    
    ' Add Input File Information
    Call PreModify(InputFileSht, currentShtStatus)
    Call loadInputSht(newdoc)
    Call PostModify(InputFileSht, currentShtStatus)
    
    ' Add output file Path
    Call LoadOutputFilePath(newdoc)
    
'--------Commenting out Iterative Functionality for this version--------'
    
    ' Output parameters and output file path are presented on iterations sheet in iteration mode
'    If Not newdoc.SelectSingleNode("//Site/Iterations") Is Nothing Then
'        ' Add Iterative Mode Information
'        Call PreModify(IterativeSht, currentShtStatus)
'        Call loadIterativeSht(newdoc)
'        Call PostModify(IterativeSht, currentShtStatus)
'    Else
        ' Add Output File Information
        Call PreModify(OutputFileSht, currentShtStatus)
        Call loadOutputSht(newdoc)
        Call PostModify(OutputFileSht, currentShtStatus)
'    End If
    
    Call PostModify(ErrorSht, errorShtStatus)
    
End Function
' Loads the Site sheet
Private Sub loadSiteSht(ByRef newdoc As DOMDocument60)
    Dim Prefix As String
    ' Loaded required parameters for Grid-Connected System and Radiation Mode
    ' Until version 1.5.2 this info is under Site, after it's under Site/SiteDef
    If (StrComp(newdoc.SelectSingleNode("//Site/Version").text, "1.5.2") < 0) Then
      Prefix = "//Site"
    Else
      Prefix = "//Site/SiteDef"
    End If
    If IntroSht.Range("ModeSelect") <> "ASTM E2848 Regression" Then
        Call InsertValue(newdoc, Prefix + "/*", SiteSht.Range("Latitude,Longitude,Altitude,TimeZone"))
        Call InsertValue_BoolToYesNo(newdoc, Prefix + "/UseLocTime", SiteSht.Range("UseLocTime"))
        
        If Not newdoc.SelectSingleNode(Prefix + "/TransEnum") Is Nothing Then
            If newdoc.SelectSingleNode(Prefix + "/TransEnum").HasChildNodes Then
                If newdoc.SelectSingleNode(Prefix + "/TransEnum").ChildNodes.NextNode.NodeValue = 0 Then
                    SiteSht.Range("TransVal").Value = "Hay"
                ElseIf newdoc.SelectSingleNode(Prefix + "/TransEnum").ChildNodes.NextNode.NodeValue = 1 Then
                    SiteSht.Range("TransVal").Value = "Perez"
                End If
            End If
        End If
        
        ' add albedo type
        Call InsertAttribute(newdoc, Prefix + "/Albedo", "Frequency", SiteSht.Range("AlbFreqVal"))
        
        ' add monthly or yearly albedo
        If SiteSht.Range("AlbFreqVal").Value = "Yearly" Then
            SiteSht.Range("AlbYearly").Value = newdoc.SelectSingleNode(Prefix + "/Albedo/Yearly").text
            Call SiteSht.SwitchFreq(SiteSht.Range("AlbFreqVal"))
        ElseIf SiteSht.Range("AlbFreqVal").Value = "Monthly" Then
            SiteSht.Range("AlbJan").Value = newdoc.SelectSingleNode(Prefix + "/Albedo/Jan").text
            SiteSht.Range("AlbFeb").Value = newdoc.SelectSingleNode(Prefix + "/Albedo/Feb").text
            SiteSht.Range("AlbMar").Value = newdoc.SelectSingleNode(Prefix + "/Albedo/Mar").text
            SiteSht.Range("AlbApr").Value = newdoc.SelectSingleNode(Prefix + "/Albedo/Apr").text
            SiteSht.Range("AlbMay").Value = newdoc.SelectSingleNode(Prefix + "/Albedo/May").text
            SiteSht.Range("AlbJun").Value = newdoc.SelectSingleNode(Prefix + "/Albedo/Jun").text
            SiteSht.Range("AlbJul").Value = newdoc.SelectSingleNode(Prefix + "/Albedo/Jul").text
            SiteSht.Range("AlbAug").Value = newdoc.SelectSingleNode(Prefix + "/Albedo/Aug").text
            SiteSht.Range("AlbSep").Value = newdoc.SelectSingleNode(Prefix + "/Albedo/Sep").text
            SiteSht.Range("AlbOct").Value = newdoc.SelectSingleNode(Prefix + "/Albedo/Oct").text
            SiteSht.Range("AlbNov").Value = newdoc.SelectSingleNode(Prefix + "/Albedo/Nov").text
            SiteSht.Range("AlbDec").Value = newdoc.SelectSingleNode(Prefix + "/Albedo/Dec").text
            Call SiteSht.SwitchFreq(SiteSht.Range("AlbFreqVal"))
        End If
    End If
    
    ' Occurs independent of system mode
    Call InsertValue(newdoc, Prefix + "/*", SiteSht.Range("Name,Country,Region,City"))
    

End Sub

' Loads the Orientaton and Shading Sheet
'NB: adding other modes to load function 28/01/2016
Private Sub loadOSSht(ByRef newdoc As DOMDocument60)
    
    Application.EnableEvents = True
    Call InsertAttribute(newdoc, "//Site/Orientation_and_Shading", "ArrayType", Orientation_and_ShadingSht.Range("OrientType"))
    Application.EnableEvents = False
    
    If (StrComp(newdoc.SelectSingleNode("//Site/Version").text, "0.9.3") >= 0) Then
        If Orientation_and_ShadingSht.Range("OrientType").Value = "Fixed Tilted Plane" Then
            If (StrComp(newdoc.SelectSingleNode("//Site/Version").text, "1.5.0") >= 0) Then    ' This is only present for version starting 1.5.0
                Call InsertValue(newdoc, "//Site/Orientation_and_Shading/*", Orientation_and_ShadingSht.Range("PlaneTiltFix,AzimuthFix,CollBandWidthFix"))
            Else
                Call InsertValue(newdoc, "//Site/Orientation_and_Shading/*", Orientation_and_ShadingSht.Range("PlaneTiltFix,AzimuthFix"))
                Orientation_and_ShadingSht.Range("CollBandWidthFix") = 1                       ' Assume default width
            End If
        
        ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Fixed Tilted Plane Seasonal Adjustment" Then
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/*", Orientation_and_ShadingSht.Range("SeasonalAdjustmentParams,AzimuthSeasonal"))

        ' Add properties for unlimited rows
        ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Unlimited Rows" Then
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/PlaneTilt", Orientation_and_ShadingSht.Range("PlaneTilt"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/Azimuth", Orientation_and_ShadingSht.Range("Azimuth"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/Pitch", Orientation_and_ShadingSht.Range("Pitch"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/CollBandWidth", Orientation_and_ShadingSht.Range("CollBandWidth"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/TopInactive", Orientation_and_ShadingSht.Range("TopInactive"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/BottomInactive", Orientation_and_ShadingSht.Range("BottomInactive"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/RowsBlock", Orientation_and_ShadingSht.Range("RowsBlock"))
            Call InsertValue_BoolToYesNo(newdoc, "//Site/Orientation_and_Shading/UseCellVal", Orientation_and_ShadingSht.Range("UseCellVal"))
        
        ' Adds values required when 'Use Cell Shading' is set to 'Yes'
            Call Orientation_and_ShadingSht.UpdateCellShading
            If Orientation_and_ShadingSht.Range("UseCellVal").Value = "Yes" Then
                Call InsertValue(newdoc, "//Site/Orientation_and_Shading/StrInWid", Orientation_and_ShadingSht.Range("StrInWid"))
                Call InsertValue(newdoc, "//Site/Orientation_and_Shading/CellSize", Orientation_and_ShadingSht.Range("CellSize"))
            End If
        
        'Adds properties for tracking
        ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Single Axis Elevation Tracking (E-W)" Then
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/AxisTiltSAET", Orientation_and_ShadingSht.Range("AxisTiltSAET"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/AxisAzimuthSAET", Orientation_and_ShadingSht.Range("AxisAzimuthSAET"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/MinTiltSAET", Orientation_and_ShadingSht.Range("MinTiltSAET"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/MaxTiltSAET", Orientation_and_ShadingSht.Range("MaxTiltSAET"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/RowsBlockSAET", Orientation_and_ShadingSht.Range("RowsBlockSAET"))
            Call Orientation_and_ShadingSht.UpdateNumberOfRowsSAET
            If Range("RowsBlockSAET").Value > 1 Then
                Call InsertValue(newdoc, "//Site/Orientation_and_Shading/PitchSAET", Orientation_and_ShadingSht.Range("PitchSAET"))
                Call InsertValue(newdoc, "//Site/Orientation_and_Shading/WActiveSAET", Orientation_and_ShadingSht.Range("WActiveSAET"))
                Call InsertValue_BoolToYesNo(newdoc, "//Site/Orientation_and_Shading/BacktrackOptSAET", Orientation_and_ShadingSht.Range("BacktrackOptSAET"))
                Call InsertValue_BoolToYesNo(newdoc, "//Site/Orientation_and_Shading/UseCellValSAET", Orientation_and_ShadingSht.Range("UseCellValSAET"))
                Call Orientation_and_ShadingSht.UpdateCellShadingSAET
                If Range("UseCellValSAET").Value = "Yes" Then
                    Call InsertValue(newdoc, "//Site/Orientation_and_Shading/StrInWidSAET", Orientation_and_ShadingSht.Range("StrInWidSAET"))
                    Call InsertValue(newdoc, "//Site/Orientation_and_Shading/CellSizeSAET", Orientation_and_ShadingSht.Range("CellSizeSAET"))
                End If
            End If
        
        ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Single Axis Horizontal Tracking (N-S)" Then
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/AxisTiltSAST", Orientation_and_ShadingSht.Range("AxisTiltSAST"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/AxisAzimuthSAST", Orientation_and_ShadingSht.Range("AxisAzimuthSAST"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/RotationMaxSAST", Orientation_and_ShadingSht.Range("RotationMaxSAST"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/RowsBlockSAST", Orientation_and_ShadingSht.Range("RowsBlockSAST"))
            Call Orientation_and_ShadingSht.UpdateNumberOfRowsSAST
            If Range("RowsBlockSAST").Value > 1 Then
                Call InsertValue(newdoc, "//Site/Orientation_and_Shading/PitchSAST", Orientation_and_ShadingSht.Range("PitchSAST"))
                Call InsertValue(newdoc, "//Site/Orientation_and_Shading/WActiveSAST", Orientation_and_ShadingSht.Range("WActiveSAST"))
                Call InsertValue_BoolToYesNo(newdoc, "//Site/Orientation_and_Shading/BacktrackOptSAST", Orientation_and_ShadingSht.Range("BacktrackOptSAST"))
                Call InsertValue_BoolToYesNo(newdoc, "//Site/Orientation_and_Shading/UseCellValSAST", Orientation_and_ShadingSht.Range("UseCellValSAST"))
                Call Orientation_and_ShadingSht.UpdateCellShadingSAST
                If Range("UseCellValSAST").Value = "Yes" Then
                    Call InsertValue(newdoc, "//Site/Orientation_and_Shading/StrInWidSAST", Orientation_and_ShadingSht.Range("StrInWidSAST"))
                    Call InsertValue(newdoc, "//Site/Orientation_and_Shading/CellSizeSAST", Orientation_and_ShadingSht.Range("CellSizeSAST"))
                End If
            End If
            
         ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Tilt and Roll Tracking" Then
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/AxisTiltTART", Orientation_and_ShadingSht.Range("AxisTiltTART"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/AxisAzimuthTART", Orientation_and_ShadingSht.Range("AxisAzimuthTART"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/RotationMinTART", Orientation_and_ShadingSht.Range("RotationMinTART"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/RotationMaxTART", Orientation_and_ShadingSht.Range("RotationMaxTART"))
                    
        ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Azimuth (Vertical Axis) Tracking" Then
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/PlaneTiltAVAT", Orientation_and_ShadingSht.Range("PlaneTiltAVAT"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/AzimuthRefAVAT", Orientation_and_ShadingSht.Range("AzimuthRefAVAT"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/MinAzimuthAVAT", Orientation_and_ShadingSht.Range("MinAzimuthAVAT"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/MaxAzimuthAVAT", Orientation_and_ShadingSht.Range("MaxAzimuthAVAT"))
        
        ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Two Axis Tracking" Then
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/MinTiltTAXT", Orientation_and_ShadingSht.Range("MinTiltTAXT"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/MaxTiltTAXT", Orientation_and_ShadingSht.Range("MaxTiltTAXT"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/AzimuthRefTAXT", Orientation_and_ShadingSht.Range("AzimuthRefTAXT"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/MinAzimuthTAXT", Orientation_and_ShadingSht.Range("MinAzimuthTAXT"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/MaxAzimuthTAXT", Orientation_and_ShadingSht.Range("MaxAzimuthTAXT"))
        End If
        
    Else
        ' If loading an older file, adapts the file to be loaded to the new format
        If Orientation_and_ShadingSht.Range("OrientType").Value = "Fixed Tilted Plane" Then
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/*", Orientation_and_ShadingSht.Range("PlaneTiltFix,AzimuthFix"))
    
        ' Add properties for unlimited rows
        ElseIf Orientation_and_ShadingSht.Range("OrientType").Value = "Unlimited Rows" Then
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/PlaneTilt", Orientation_and_ShadingSht.Range("PlaneTilt"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/Azimuth", Orientation_and_ShadingSht.Range("Azimuth"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/Pitch", Orientation_and_ShadingSht.Range("Pitch"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/CollBandWidth", Orientation_and_ShadingSht.Range("CollBandWidth"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/TopInactive", Orientation_and_ShadingSht.Range("TopInactive"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/BottomInactive", Orientation_and_ShadingSht.Range("BottomInactive"))
            Call InsertValue(newdoc, "//Site/Orientation_and_Shading/RowsBlock", Orientation_and_ShadingSht.Range("RowsBlock"))
            Call InsertValue_BoolToYesNo(newdoc, "//Site/Orientation_and_Shading/UseCellVal", Orientation_and_ShadingSht.Range("UseCellVal"))
        
            ' Adds values required when 'Use Cell Shading' is set to 'Yes'
            If Orientation_and_ShadingSht.Range("UseCellVal").Value = "Yes" Then
                Call InsertValue(newdoc, "//Site/Orientation_and_Shading/StrInWid", Orientation_and_ShadingSht.Range("StrInWid"))
                Call InsertValue(newdoc, "//Site/Orientation_and_Shading/CellSize", Orientation_and_ShadingSht.Range("CellSize"))
            End If
        End If
    End If
    
   
   
End Sub

' Loads the Bifacial sheet
Private Sub loadBifacialSht(ByRef newdoc As DOMDocument60)
    
    Dim EventsEnabled As Boolean
    EventsEnabled = Application.EnableEvents
    
    Application.EnableEvents = True              ' Required to update the Bifacial sheet depending upon the value of UseBifacialModel
    
    ' No bifacial information in files prior to version 1.5.0
    If (StrComp(newdoc.SelectSingleNode("//Site/Version").text, "1.5.0") < 0) Then
        BifacialSht.Range("UseBifacialModel") = "No"
    
    ' Bifacial information present
    Else
        ' Check if spectral model is used
        Call InsertValue_BoolToYesNo(newdoc, "//Site/Bifacial/UseBifacialModel", BifacialSht.Range("UseBifacialModel"))
    
        If BifacialSht.Range("UseBifacialModel").Value = "Yes" Then
            
            Call InsertValue(newdoc, "//Site/Bifacial/GroundClearance", BifacialSht.Range("GroundClearance"))
            Call InsertValue(newdoc, "//Site/Bifacial/StructBlockingFactor", BifacialSht.Range("StructBlockingFactor"))
            Call InsertValue(newdoc, "//Site/Bifacial/PanelTransFactor", BifacialSht.Range("PanelTransFactor"))
            Call InsertValue(newdoc, "//Site/Bifacial/BifacialityFactor", BifacialSht.Range("BifacialityFactor"))
            
            ' add albedo type
            Call InsertAttribute(newdoc, "//Site/Bifacial/BifAlbedo", "Frequency", BifacialSht.Range("BifAlbFreqVal"))
        
            ' add monthly or yearly albedo
            If BifacialSht.Range("BifAlbFreqVal").Value = "Site" Then
                Call BifacialSht.BifSwitchFreq(BifacialSht.Range("BifAlbFreqVal"))
            ElseIf BifacialSht.Range("BifAlbFreqVal").Value = "Yearly" Then
                BifacialSht.Range("BifAlbYearly").Value = newdoc.SelectSingleNode("//Site/Bifacial/BifAlbedo/Yearly").text
                Call BifacialSht.BifSwitchFreq(BifacialSht.Range("BifAlbFreqVal"))
            ElseIf BifacialSht.Range("BifAlbFreqVal").Value = "Monthly" Then
                BifacialSht.Range("BifAlbJan").Value = newdoc.SelectSingleNode("//Site/Bifacial/BifAlbedo/Jan").text
                BifacialSht.Range("BifAlbFeb").Value = newdoc.SelectSingleNode("//Site/Bifacial/BifAlbedo/Feb").text
                BifacialSht.Range("BifAlbMar").Value = newdoc.SelectSingleNode("//Site/Bifacial/BifAlbedo/Mar").text
                BifacialSht.Range("BifAlbApr").Value = newdoc.SelectSingleNode("//Site/Bifacial/BifAlbedo/Apr").text
                BifacialSht.Range("BifAlbMay").Value = newdoc.SelectSingleNode("//Site/Bifacial/BifAlbedo/May").text
                BifacialSht.Range("BifAlbJun").Value = newdoc.SelectSingleNode("//Site/Bifacial/BifAlbedo/Jun").text
                BifacialSht.Range("BifAlbJul").Value = newdoc.SelectSingleNode("//Site/Bifacial/BifAlbedo/Jul").text
                BifacialSht.Range("BifAlbAug").Value = newdoc.SelectSingleNode("//Site/Bifacial/BifAlbedo/Aug").text
                BifacialSht.Range("BifAlbSep").Value = newdoc.SelectSingleNode("//Site/Bifacial/BifAlbedo/Sep").text
                BifacialSht.Range("BifAlbOct").Value = newdoc.SelectSingleNode("//Site/Bifacial/BifAlbedo/Oct").text
                BifacialSht.Range("BifAlbNov").Value = newdoc.SelectSingleNode("//Site/Bifacial/BifAlbedo/Nov").text
                BifacialSht.Range("BifAlbDec").Value = newdoc.SelectSingleNode("//Site/Bifacial/BifAlbedo/Dec").text
                Call BifacialSht.BifSwitchFreq(BifacialSht.Range("BifAlbFreqVal"))
            End If
    
        End If
    End If
    
    Application.EnableEvents = EventsEnabled
    

End Sub

'Loads the Horizon Shading Sheet
Private Sub loadHorizon_ShadingSht(ByRef newdoc As DOMDocument60)

    If Not newdoc.SelectSingleNode("//Site/Orientation_and_Shading/DefineHorizonProfile") Is Nothing Then
        Application.EnableEvents = True
        Call InsertValue_BoolToYesNo(newdoc, "//Site/Orientation_and_Shading/DefineHorizonProfile", Horizon_ShadingSht.Range("DefineHorizonProfile"))
        Application.EnableEvents = False
        If Range("DefineHorizonProfile").Value = "Yes" Then
            If Not newdoc.SelectSingleNode("//Site/Orientation_and_Shading/HorizonAzi") Is Nothing Then
                Range("NumHorPts").Value = WorksheetFunction.Min(Len(newdoc.SelectSingleNode("//Site/Orientation_and_Shading/HorizonAzi").text) - Len(Replace(newdoc.SelectSingleNode("//Site/Orientation_and_Shading/HorizonAzi").text, ",", "")) + 1, _
                                                                 Len(newdoc.SelectSingleNode("//Site/Orientation_and_Shading/HorizonElev").text) - Len(Replace(newdoc.SelectSingleNode("//Site/Orientation_and_Shading/HorizonElev").text, ",", "")) + 1)
                Dim n As Integer
                n = Range("NumHorPts").Value
                Range(Range("HElevFirst").Address & ":" & Range("HElevFirst").Offset(n - 1, 0).Address).Value = WorksheetFunction.Transpose(Split(newdoc.SelectSingleNode("//Site/Orientation_and_Shading/HorizonElev").text, ","))
                Range(Range("HAziFirst").Address & ":" & Range("HAziFirst").Offset(n - 1, 0).Address).Value = WorksheetFunction.Transpose(Split(newdoc.SelectSingleNode("//Site/Orientation_and_Shading/HorizonAzi").text, ","))
                If (Len(newdoc.SelectSingleNode("//Site/Orientation_and_Shading/HorizonAzi").text) - Len(Replace(newdoc.SelectSingleNode("//Site/Orientation_and_Shading/HorizonAzi").text, ",", ""))) <> (Len(newdoc.SelectSingleNode("//Site/Orientation_and_Shading/HorizonElev").text) - Len(Replace(newdoc.SelectSingleNode("//Site/Orientation_and_Shading/HorizonElev").text, ",", ""))) Then
                    Call ErrorLogger("Horizon: Unequal number of horizon azimuth and elevation values. Only values with respective elevation and azimuth values loaded.")
                End If
    
            Else
                Call Horizon_ShadingSht.ClearHorizon
            End If
        End If
    Else
        ' Loading from a previous version where horizon shading was not defined
        Application.EnableEvents = True
        Call Horizon_ShadingSht.ClearHorizon
        Range("DefineHorizonProfile").Value = "No"
        Application.EnableEvents = False
    End If
End Sub
' Loads the System Sheet
Private Sub loadSystemSht(ByRef newdoc As DOMDocument60)

    ' Add System And Array Values
    Call LoadSystemArrays(newdoc)

End Sub
' Loads the Losses sheet
Private Sub loadLossesSht(ByRef newdoc As DOMDocument60)
    Dim IAMNode As IXMLDOMNode
    Dim i As Integer
    Dim usePAN As Boolean
    Dim SiteLosses As String
    
    ' Deal with change of format from 1.5.1 to 1.5.2:
    ' In versions prior to 1.5.2, information is stored (erroneously) under //Site/System/Losses
    ' Starting in version 1.5.2, information is stored under //Site/Losses
    If (StrComp(newdoc.SelectSingleNode("//Site/Version").text, "1.5.2") < 0) Then
      SiteLosses = "//Site/System/Losses"
    Else
      SiteLosses = "//Site/Losses"
    End If
    
    Application.EnableEvents = True
    LossesSht.Range("IAMRange").ClearContents
    Call InsertValue(newdoc, SiteLosses + "/ThermalLosses/UseMeasuredValues", LossesSht.Range("UseMeasuredValues"))
    
    If LossesSht.Range("UseMeasuredValues").Value = False Then
        Call InsertValue(newdoc, SiteLosses + "/ThermalLosses/ConsHLF", LossesSht.Range("ConsHLF"))
        Call InsertValue(newdoc, SiteLosses + "/ThermalLosses/ConvHLF", LossesSht.Range("ConvHLF"))
    End If
 
    Call InsertValue(newdoc, SiteLosses + "/ModuleQualityLosses/EfficiencyLoss", LossesSht.Range("EfficiencyLoss"))
    
    ' No Module LID and Ageing information in files prior to version 1.5.2
    If (StrComp(newdoc.SelectSingleNode("//Site/Version").text, "1.5.2") < 0) Then
        LossesSht.Range("ModuleLID") = 0
        LossesSht.Range("ModuleAgeing") = 0
    Else
        Call InsertValue(newdoc, SiteLosses + "/ModuleQualityLosses/ModuleLID", LossesSht.Range("ModuleLID"))
        Call InsertValue(newdoc, SiteLosses + "/ModuleQualityLosses/ModuleAgeing", LossesSht.Range("ModuleAgeing"))
    End If
    
    Call InsertValue(newdoc, SiteLosses + "/ModuleMismatchLosses/*", LossesSht.Range("PowerLoss,LossFixedVoltage"))
    
    ' For version control. files defined before version 0.9.1 do not have user defined IAM.
    If Not newdoc.SelectSingleNode("//Version").text = "0.9" Then
        ' Load IAMSelection: ASHRAE or User Defined
        Call InsertAttribute(newdoc, "//IncidenceAngleModifier", "IAMSelection", LossesSht.Range("IAMSelection"))
        
        If LossesSht.Range("IAMSelection").Value = "User Defined" Then
            Call InsertValue(newdoc, "//IncidenceAngleModifier/*", LossesSht.Range("IAMRange"), False, True)
        ElseIf LossesSht.Range("IAMSelection").Value = "ASHRAE" Then
            Call InsertValue(newdoc, SiteLosses + "/IncidenceAngleModifier/bNaught", LossesSht.Range("bNaught"))
        End If
    Else
        ' Load an older site file
        LossesSht.Range("IAMSelection").Value = "ASHRAE"
        Call InsertValue(newdoc, "//bNaught", LossesSht.Range("bNaught"))
    End If
    For i = 0 To SystemSht.Range("NumSubArray").Value
        If Not newdoc.SelectSingleNode("//Site/SubArray" & i & "/PVModule/IAMDefinition") Is Nothing Then
            usePAN = True
            Exit For
        End If
        If i = SystemSht.Range("NumSubArray").Value Then
            usePAN = False
        End If
    Next i
    If usePAN Then
        LossesSht.Range("UsePAN").Value = "Yes"
    End If

End Sub
' Loads the Soiling Sheet
Private Sub loadSoilingSht(ByRef newdoc As DOMDocument60)
    
    Dim SiteSoilingLosses As String
    
    ' Deal with change of format from 1.5.1 to 1.5.2:
    ' In versions prior to 1.5.2, information is stored (erroneously) under //Site/System/Losses/SoilingLosses
    ' Starting in version 1.5.2, information is stored under //Site/SoilingLosses
    If (StrComp(newdoc.SelectSingleNode("//Site/Version").text, "1.5.2") < 0) Then
      SiteSoilingLosses = "//Site/System/Losses/SoilingLosses"
    Else
      SiteSoilingLosses = "//Site/SoilingLosses"
    End If
    
    Call InsertAttribute(newdoc, SiteSoilingLosses, "Frequency", SoilingSht.Range("SfreqVal"))
    
    ' Add monthly or yearly soiling
    If SoilingSht.Range("SfreqVal").Value = "Yearly" Then
        ' Fix bug from versions < 1.5.0: SoilingYearly may actually be stored in PrintSoilingYearly
        Dim xmlNodelist As IXMLDOMNodeList
        Dim tmpPath As String
        tmpPath = SiteSoilingLosses + "/PrintSoilingYearly"
        Set xmlNodelist = newdoc.SelectNodes(tmpPath)
        If xmlNodelist.Item(0) Is Nothing Then
            tmpPath = SiteSoilingLosses + "/Yearly"
        End If
        ' End bug fix
        SoilingSht.Range("SoilingYearly").Value = newdoc.SelectSingleNode(tmpPath).text
    ElseIf SoilingSht.Range("SfreqVal").Value = "Monthly" Then
        SoilingSht.Range("SoilingJan").Value = newdoc.SelectSingleNode(SiteSoilingLosses + "/Jan").text
        SoilingSht.Range("SoilingFeb").Value = newdoc.SelectSingleNode(SiteSoilingLosses + "/Feb").text
        SoilingSht.Range("SoilingMar").Value = newdoc.SelectSingleNode(SiteSoilingLosses + "/Mar").text
        SoilingSht.Range("SoilingApr").Value = newdoc.SelectSingleNode(SiteSoilingLosses + "/Apr").text
        SoilingSht.Range("SoilingMay").Value = newdoc.SelectSingleNode(SiteSoilingLosses + "/May").text
        SoilingSht.Range("SoilingJun").Value = newdoc.SelectSingleNode(SiteSoilingLosses + "/Jun").text
        SoilingSht.Range("SoilingJul").Value = newdoc.SelectSingleNode(SiteSoilingLosses + "/Jul").text
        SoilingSht.Range("SoilingAug").Value = newdoc.SelectSingleNode(SiteSoilingLosses + "/Aug").text
        SoilingSht.Range("SoilingSep").Value = newdoc.SelectSingleNode(SiteSoilingLosses + "/Sep").text
        SoilingSht.Range("SoilingOct").Value = newdoc.SelectSingleNode(SiteSoilingLosses + "/Oct").text
        SoilingSht.Range("SoilingNov").Value = newdoc.SelectSingleNode(SiteSoilingLosses + "/Nov").text
        SoilingSht.Range("SoilingDec").Value = newdoc.SelectSingleNode(SiteSoilingLosses + "/Dec").text
    End If

End Sub
' Loads the Spectral Sheet
Private Sub loadSpectralSht(ByRef newdoc As DOMDocument60)
    
    Dim EventsEnabled As Boolean
    EventsEnabled = Application.EnableEvents
    
    Application.EnableEvents = True              ' Required to update the Spectral sheet depending upon the value of UseSpectralModel
    
    ' No spectral information in files prior to version 1.4.0
    If (StrComp(newdoc.SelectSingleNode("//Site/Version").text, "1.4.0") < 0) Then
        SpectralSht.Range("UseSpectralModel") = "No"
    
    ' Spectral information present
    Else
    
        ' Check if spectral model is used
        Call InsertValue_BoolToYesNo(newdoc, "//Site/Spectral/UseSpectralModel", SpectralSht.Range("UseSpectralModel"))
    
        ' Add values of soiling as a function of clearness index
        ' The double call to WorksheetFunction.Transpose magically transforms the strings into values (!)
        If SpectralSht.Range("UseSpectralModel").Value = "Yes" Then
            Range("ktCorrectionValues").Value = WorksheetFunction.Transpose(WorksheetFunction.Transpose(Split(newdoc.SelectSingleNode("//Site/Spectral/ClearnessIndex/ktCorrection").text, ",")))
        End If
    End If
    
    Application.EnableEvents = EventsEnabled
End Sub
' Loads the Transformer Sheet
Private Sub loadTransformerSht(ByRef newdoc As DOMDocument60)

    Dim SiteTransformer As String
    
    ' Deal with change of format from 1.5.1 to 1.5.2:
    ' In versions prior to 1.5.2, information is stored (erroneously) under //Site/System/Transformer
    ' Starting in version 1.5.2, information is stored under //Site/Transformer
    If (StrComp(newdoc.SelectSingleNode("//Site/Version").text, "1.5.2") < 0) Then
      SiteTransformer = "//Site/System/Transformer"
    Else
      SiteTransformer = "//Site/Transformer"
    End If
      
    Application.EnableEvents = False
    Call InsertValue(newdoc, SiteTransformer + "/PNomTrf", TransformerSht.Range("PNomTrf"))
    Call InsertValue(newdoc, SiteTransformer + "/PIronLossTrf", TransformerSht.Range("PIronLossTrf"))
    Call InsertValue(newdoc, SiteTransformer + "/NightlyDisconnect", TransformerSht.Range("NightlyDisconnect"))
    
    'Versions older than 1.2.0 have a PGlobLossTrf value saved for PFullLoadLss
    If (StrComp(newdoc.SelectSingleNode("//Version").text, "1.2.0") < 0) Then
        Range("PFullLoadLss").Value = newdoc.SelectSingleNode(SiteTransformer + "/PGlobLossTrf").text
    Else
         Call InsertValue(newdoc, SiteTransformer + "/PFullLoadLss", TransformerSht.Range("PFullLoadLss"))
    End If
    Application.EnableEvents = True

End Sub
' Loads the Input File sheet
Private Sub loadInputSht(ByRef newdoc As DOMDocument60)
    Dim validFilePath As Boolean
    Dim inputFilePath As String
    Dim relativeFilePath As String
    ' For displaying only a relative file path
    Dim InputFileName As String
    ' Variables to figure out whether the input file is in the same directory as the csyx or CASSYS
    Dim FilePathLeft As String
    Dim FilePathLeft_csyx As String
    
    On Error Resume Next
    Application.EnableEvents = False
    
    InputFileSht.Range("lastInputColumn").Value = 0
    InputFileSht.Range("InputColumnNums").ClearContents
    
        
    ' Check if input file path is defined. If not, then make it red and notify user with a message on the error sheet
    inputFilePath = newdoc.SelectSingleNode("//Site/InputFilePath").text ' InputFileSht.Range("InputFilePath").Value
    inputFilePath = Replace(inputFilePath, "/", "\")
    ' Defines input file name
    FilePathLeft = Left(inputFilePath, Len(ThisWorkbook.path))
    FilePathLeft = Replace(FilePathLeft, "/", "\")
    FilePathLeft_csyx = Left(inputFilePath, Len(Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\"))))
    FilePathLeft_csyx = Replace(FilePathLeft_csyx, "/", "\")
    
    ' If the file path is in the same directory as csyx or CASSYS or is in a folder further down, the file path is given as a relative file path
    If ThisWorkbook.path = FilePathLeft Then
        InputFileName = Right(inputFilePath, Len(inputFilePath) - Len(ThisWorkbook.path) - 1)
    ElseIf Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\")) = FilePathLeft_csyx Then
        InputFileName = Right(inputFilePath, Len(inputFilePath) - Len(FilePathLeft_csyx))
    ' If the file is stored somewhere else entirely, the whole file path is stored
    Else
        InputFileName = inputFilePath
    End If

    ' Displays input file name
    InputFileSht.Range("InputFilePath").Value = InputFileName
    validFilePath = checkValidFilePath(InputFileSht, "Input", inputFilePath)
    ' If the file path is not correct then check the relative path (check the directory as CASSYS.xlsm is contained in)
    If validFilePath = False Then
        relativeFilePath = Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\")) & Right(inputFilePath, Len(FilePathLeft_csyx))
        validFilePath = checkValidFilePath(InputFileSht, "Input", relativeFilePath)
        If validFilePath = True Then
            inputFilePath = relativeFilePath
            InputFileName = Right(inputFilePath, Len(inputFilePath) - Len(FilePathLeft_csyx))
            InputFileSht.Range("InputFilePath").Value = InputFileName
        Else:
            relativeFilePath = ThisWorkbook.path & "\" & Right(inputFilePath, Len(FilePathLeft))
            validFilePath = checkValidFilePath(InputFileSht, "Input", relativeFilePath)
            If validFilePath = True Then
                inputFilePath = relativeFilePath
                InputFileName = Right(inputFilePath, Len(inputFilePath) - Len(ThisWorkbook.path) - 1)
                InputFileSht.Range("InputFilePath").Value = InputFileName
            ' Looks for file in location of .csyx file as well as in location of CASSYS itself
            Else:
                relativeFilePath = Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\")) & Right(inputFilePath, Len(inputFilePath) - InStrRev(inputFilePath, "\"))
                validFilePath = checkValidFilePath(InputFileSht, "Input", relativeFilePath)
                If validFilePath = True Then
                    inputFilePath = relativeFilePath
                    InputFileName = Right(inputFilePath, Len(inputFilePath) - Len(Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\"))))
                    InputFileSht.Range("InputFilePath").Value = InputFileName
                Else:
                    relativeFilePath = ThisWorkbook.path & "\" & Right(inputFilePath, Len(inputFilePath) - InStrRev(inputFilePath, "\"))
                    validFilePath = checkValidFilePath(InputFileSht, "Input", relativeFilePath)
                    If validFilePath = True Then
                        inputFilePath = relativeFilePath
                        InputFileName = Right(inputFilePath, Len(inputFilePath) - Len(ThisWorkbook.path) - 1)
                        InputFileSht.Range("InputFilePath").Value = InputFileName
                    End If
                End If
            End If
        End If
    Else
        ' Change browse path color back to white if the input file exists
        InputFileSht.Range("InputFilePath").Interior.Color = ColourWhite
        If checkValidFilePath(InputFileSht, "Input", CurDir() & "\" & inputFilePath) = False Then
            Range("FullInputPath").Value = inputFilePath
        Else
            Range("FullInputPath").Value = CurDir() & "\" & inputFilePath
        End If
    End If
    
    ' Continues to load input file sheet even if input file is invalid
    ' If this is a new version,
    If IntroSht.Range("ModeSelect") = "ASTM E2848 Regression" Then
        Call InsertValue(newdoc, "//Site/InputFileStyle/*", InputFileSht.Range("RowsToSkip,Delimeter,TimeFormat,AveragedAt,Interval,ASTMInputRange,TMYType"))
    Else
        Call InsertValue(newdoc, "//Site/InputFileStyle/*", InputFileSht.Range("RowsToSkip,Delimeter,TimeFormat,AveragedAt,Interval,InputColumnNums,TMYType"))
    End If
    
    If (Not newdoc.SelectSingleNode("//Site/InputFileStyle/IncorrectClimateRowsAllowed") Is Nothing) Then
        Call InsertValue(newdoc, "//Site/InputFileStyle/IncorrectClimateRowsAllowed", InputFileSht.Range("IncorrectClimateRowsAllowed"))
    Else
        InputFileSht.Range("IncorrectClimateRowsAllowed").Value = 0
    End If
    
    ' Not version specific.
    Call InputFileSht.ChangeClimateFile
    'Call InputFileSht.PreviewInput
    'InputFileSht.GetDates (inputFilePath)
    
    ' Version Control for 0.9.2 and 0.9.3, it is possible to load Tilt and Meter Azimuth
    If (newdoc.SelectSingleNode("//Version").text = "0.9.2" Or newdoc.SelectSingleNode("//Version").text = "0.9.3") Then
        Call InsertValue(newdoc, "//Site/InputFileStyle/*", InputFileSht.Range("MeterTilt,MeterAzimuth"))
        If (InputFileSht.Range("MeterTilt").Value = "N/A" Or InputFileSht.Range("MeterTilt").Value = "N/A") And InputFileSht.Range("GlobalRad").Value = vbNullString Then
            InputFileSht.Range("MeterTilt").Interior.Color = RGB(176, 220, 231)
            InputFileSht.Range("MeterAzimuth").Interior.Color = RGB(176, 220, 231)
        End If
    End If
    If validFilePath = False Then
        Call ErrorLogger("Climate file path: file not specified or was invalid.")
    End If
    
    Application.EnableEvents = True

End Sub

' Loads output file path to iterative and output file sheet
Private Sub LoadOutputFilePath(ByRef newdoc As DOMDocument60)
    
    Dim OutputFilePath As String
    Dim validDirectory As Boolean
    Dim relativeFilePath As String
    ' Variables to figure out whether the input file is in the same directory as the csyx or CASSYS
    Dim FilePathLeft As String
    Dim FilePathLeft_csyx As String
    
    'on error resume next
    Application.EnableEvents = False
    Call InsertValue(newdoc, "//Site/OutputFilePath", OutputFileSht.Range("OutputFilePath"))
    
    'Check if output file path is defined. If not, then make it red and notify user.
    OutputFilePath = ThisWorkbook.path & "\" & OutputFileSht.Range("OutputFilePath").Value
    validDirectory = checkValidFilePath(OutputFileSht, "Output", Left(OutputFilePath, InStrRev(OutputFilePath, "/")))
    
    ' Defines input file name
    FilePathLeft = Left(OutputFilePath, Len(ThisWorkbook.path))
    FilePathLeft = Replace(FilePathLeft, "/", "\")
    FilePathLeft_csyx = Left(OutputFilePath, Len(Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\"))))
    FilePathLeft_csyx = Replace(FilePathLeft_csyx, "/", "\")
    
    If validDirectory = False And OutputFileSht.Range("OutputFilePath").Value <> vbNullString Then
        ' If the file path is in the same directory as csyx or CASSYS or is in a folder further down, the file path is given as a relative file path
        If ThisWorkbook.path = FilePathLeft Then
            OutputFileSht.Range("OutputFilePath").Value = Right(OutputFilePath, Len(OutputFilePath) - Len(ThisWorkbook.path) - 1)
        ElseIf (Left(IntroSht.Range("LoadFilePath").Value, InStrRev(IntroSht.Range("LoadFilePath").Value, "\")) = FilePathLeft_csyx) And IntroSht.Range("LoadFilePath").Value <> "" Then
            OutputFileSht.Range("OutputFilePath").Value = Right(OutputFilePath, Len(OutputFilePath) - Len(FilePathLeft_csyx))
        ' If the file is stored somewhere else entirely, the whole file path is stored
        Else
            OutputFileSht.Range("OutputFilePath").Value = OutputFilePath
        End If
        
    End If
    
'    IterativeSht.Range("OutputFilePath").Value = OutputFileSht.Range("OutputFilePath").Value

    Application.EnableEvents = True

End Sub


' Loads the Output File Sheet
'NB: edited so the OutputFilePath cell displays the proper output file path
Private Sub loadOutputSht(ByRef newdoc As DOMDocument60)
    ' Set sheet to visible
    OutputFileSht.Visible = xlSheetVisible
    
    ' Loading output parameters
    Call InsertValue_OutputSelections(newdoc, "//Site/OutputFileStyle/*", OutputFileSht.Range("OutputParam"))
    Call OutputFileSht.ChangeCellColour(OutputFileSht.Range("OutputParam"))
    
End Sub
' Loads the Sub-Arrays (PV modules/Inverters) on the System sheet
' and checks for discrepancies/errors when loading from the database
' NB: adjusted so the file names of modules are added and taken from csyx file or database and displayed
Private Sub LoadSystemArrays(ByRef newdoc As DOMDocument60)
    
    Dim sysNode As IXMLDOMNode ' The system nodes
    Dim subarrNode As IXMLDOMNode ' The sub arry node
    Dim pvNode As IXMLDOMNode ' The pv array node
    Dim pvParam As IXMLDOMNode ' The specifications of the pv array
    Dim invNode As IXMLDOMNode ' The inverter node
    Dim efficiencyCurve As IXMLDOMNode 'Used for looping through XML
    Dim efficiencyCurveParam As IXMLDOMNode
    Dim lastRow As Long ' last row in the PV database
    Dim ColumnIndex As Integer ' the column Index of a parameter on the PV database sheet
    Dim j As Integer 'Counter
    
     ' Sheet status used for pre/post modify
    Dim pvExists As Boolean ' Whether or not the pv array exists
    Dim invExists As Boolean 'Whether or not the inverter exists
    Dim Model As String ' the pv/inverter model
    Dim Manu As String ' the pv/inverter manufacturer
    Dim Source As String ' the pv/inverter origin
    Dim Check As String ' the value from the database to be checked against the value in the pv/inverter nodes
    Dim getIndex As Integer ' the Index of the pv array or inverter
    Dim i As Integer ' counter
    Dim auxOffset As Integer
    Dim theParam As Range
    
    Dim fileName As String ' The name of the file from which the pv array or inverter comes
    
    ' Add System and Sub-Array Info
    
    ' Enter number of arrays
    ' For this cell only, events have to be enabled so that the system sheet gets properly resized
    Application.EnableEvents = True
    Call InsertAttribute(newdoc, "//Site/System", "TotalArrays", SystemSht.Range("NumSubArray"))
    
    If (Not newdoc.SelectSingleNode("//Site/System/ACWiringLossAtSTC") Is Nothing) Then
        If (newdoc.SelectSingleNode("//Site/System/ACWiringLossAtSTC").text = "False") Then
            SystemSht.Range("ACWiringLossAtSTC").Value = "at Pnom"
        Else
            SystemSht.Range("ACWiringLossAtSTC").Value = "at STC"
        End If
    Else
        SystemSht.Range("ACWiringLossAtSTC").Value = "at STC"
    End If
    
    'System sheet does not load if file is in radiation mode
    If IntroSht.Range("ModeSelect").Value = "Radiation Mode" Then Exit Sub
    
    For i = 1 To SystemSht.Range("NumSubArray").Value
        Application.EnableEvents = True       ' DJT whether this line is necessary should be investigated
        auxOffset = (i - 1) * SubArrayHeight
        pvExists = True
        
        ' Prefix for PV module info
        Dim fixPre As String
        fixPre = "//Site/System/SubArray" & i & "/PVModule"
    
        Set subarrNode = newdoc.SelectSingleNode(fixPre)

        If subarrNode.HasChildNodes Then
        
            ' Get model, manufacturer and source
            ' NB: and file name
        
            Set sysNode = newdoc.SelectSingleNode(fixPre & "/Module")
            If sysNode.HasChildNodes Then
                Model = sysNode.ChildNodes.NextNode.NodeValue
            Else
                pvExists = False
            End If
        
            Set sysNode = newdoc.SelectSingleNode(fixPre & "/Manufacturer")
            If sysNode.HasChildNodes Then
                Manu = sysNode.ChildNodes.NextNode.NodeValue
            Else
                pvExists = False
            End If
        
            Set sysNode = newdoc.SelectSingleNode(fixPre & "/Origin")
            If sysNode.HasChildNodes Then
                Source = sysNode.ChildNodes.NextNode.NodeValue
            Else
                pvExists = False
            End If
        
            ' NB: Gives value to file name, unless file name does not exist in csyx document
            Set sysNode = newdoc.SelectSingleNode(fixPre & "/FileName")
            If Not sysNode Is Nothing Then
                If sysNode.HasChildNodes Then
                    fileName = sysNode.ChildNodes.NextNode.NodeValue
                Else:
                    fileName = ""
                End If
            Else:
                fileName = ""
            End If

        
        
            If pvExists Then
                ' get the Index of the pv module with the module, manufacturer and version values
                getIndex = PVIndex(Manu, Model, Source)
            
                ' If the module is found
                If Not getIndex = 0 Then
                    ' Update Index
                    SystemSht.Range("PVDataIndex").Offset(auxOffset, 0).Value = getIndex
                Else
                    ' If not set pvExists to false and tell user that it does not exist
                    pvExists = False
                    ' NB: removed line: Call ErrorLogger("Warning: The Module " & Model & " was not found. This module has been temporarily added to the database.")
                    lastRow = PV_DatabaseSht.Range("A" & PV_DatabaseSht.Rows.count).End(xlUp).row + 1
                    ColumnIndex = 8
                    
                    PV_DatabaseSht.Cells(lastRow, 1).Value = "User_Added"
                    PV_DatabaseSht.Cells(lastRow, 2).Value = Manu
                    PV_DatabaseSht.Cells(lastRow, 3).Value = Model
                    ' NB: displays Module file name
                    If fileName = "" Then
                        PV_DatabaseSht.Cells(lastRow, 4).Value = IntroSht.Range("IntroFileName").Value2
                    Else:
                        PV_DatabaseSht.Cells(lastRow, 4).Value = fileName
                    End If
                    
                    On Error Resume Next
                                    
                    PV_DatabaseSht.Cells(lastRow, 8).Value = newdoc.SelectSingleNode(fixPre + "/Pnom").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 9).Value = newdoc.SelectSingleNode(fixPre + "/PnomToLow").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 10).Value = newdoc.SelectSingleNode(fixPre + "/PnomToUp").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 10).Value = newdoc.SelectSingleNode(fixPre + "/Toleranz").ChildNodes.NextNode.NodeValue
                                    
                    PV_DatabaseSht.Cells(lastRow, 11).Value = newdoc.SelectSingleNode(fixPre + "/LIDloss").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 12).Value = newdoc.SelectSingleNode(fixPre + "/Technology").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 13).Value = newdoc.SelectSingleNode(fixPre + "/CellsinS").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 14).Value = newdoc.SelectSingleNode(fixPre + "/CellsinP").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 15).Value = newdoc.SelectSingleNode(fixPre + "/Gref").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 16).Value = newdoc.SelectSingleNode(fixPre + "/Tref").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 17).Value = newdoc.SelectSingleNode(fixPre + "/Vmpp").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 18).Value = newdoc.SelectSingleNode(fixPre + "/Impp").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 19).Value = newdoc.SelectSingleNode(fixPre + "/Voc").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 20).Value = newdoc.SelectSingleNode(fixPre + "/Isc").ChildNodes.NextNode.NodeValue

                    PV_DatabaseSht.Cells(lastRow, 21).Value = newdoc.SelectSingleNode(fixPre + "/mIsc").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 22).Value = newdoc.SelectSingleNode(fixPre + "/mVco").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 23).Value = newdoc.SelectSingleNode(fixPre + "/mPmpp").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 24).Value = newdoc.SelectSingleNode(fixPre + "/Rsh0").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 25).Value = newdoc.SelectSingleNode(fixPre + "/Rshexp").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 26).Value = newdoc.SelectSingleNode(fixPre + "/Rshunt").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 27).Value = newdoc.SelectSingleNode(fixPre + "/Rserie").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 28).Value = newdoc.SelectSingleNode(fixPre + "/Gamma").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 29).Value = newdoc.SelectSingleNode(fixPre + "/muGamma").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 30).Value = newdoc.SelectSingleNode(fixPre + "/RelEffic800").ChildNodes.NextNode.NodeValue
                    
                    PV_DatabaseSht.Cells(lastRow, 31).Value = newdoc.SelectSingleNode(fixPre + "/RelEffic700").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 32).Value = newdoc.SelectSingleNode(fixPre + "/RelEffic600").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 33).Value = newdoc.SelectSingleNode(fixPre + "/RelEffic400").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 34).Value = newdoc.SelectSingleNode(fixPre + "/RelEffic200").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 35).Value = newdoc.SelectSingleNode(fixPre + "/Vmax").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 36).Value = newdoc.SelectSingleNode(fixPre + "/ByPassDiodes").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 37).Value = newdoc.SelectSingleNode(fixPre + "/ByPassDiodeVoltage").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 38).Value = newdoc.SelectSingleNode(fixPre + "/Brev").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 39).Value = newdoc.SelectSingleNode(fixPre + "/Length").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 40).Value = newdoc.SelectSingleNode(fixPre + "/Width").ChildNodes.NextNode.NodeValue
                    
                    PV_DatabaseSht.Cells(lastRow, 41).Value = newdoc.SelectSingleNode(fixPre + "/Depth").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 42).Value = newdoc.SelectSingleNode(fixPre + "/Weight").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 43).Value = newdoc.SelectSingleNode(fixPre + "/AreaM").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 44).Value = newdoc.SelectSingleNode(fixPre + "/Cellarea").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 45).Value = newdoc.SelectSingleNode(fixPre + "/Area").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 47).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/AOI1").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 48).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/Mod1").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 49).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/AOI2").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 50).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/Mod2").ChildNodes.NextNode.NodeValue
                    
                    PV_DatabaseSht.Cells(lastRow, 51).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/AOI3").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 52).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/Mod3").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 53).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/AOI4").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 54).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/Mod4").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 55).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/AOI5").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 56).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/Mod5").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 57).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/AOI6").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 58).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/Mod6").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 59).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/AOI7").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 60).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/Mod7").ChildNodes.NextNode.NodeValue
                    
                    PV_DatabaseSht.Cells(lastRow, 61).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/AOI8").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 62).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/Mod8").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 63).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/AOI9").ChildNodes.NextNode.NodeValue
                    PV_DatabaseSht.Cells(lastRow, 64).Value = newdoc.SelectSingleNode(fixPre + "/IAMDefinition/Mod9").ChildNodes.NextNode.NodeValue
                    
                    On Error GoTo 0
                             
                    SystemSht.Range("PVDataIndex").Offset(auxOffset, 0).Value = lastRow - 4
                End If
            End If
           
            Call InsertArrayValue(newdoc, fixPre & "/ModulesInString", SystemSht.Range("ModStr").Offset(auxOffset, 0))
            Call InsertArrayValue(newdoc, fixPre & "/NumStrings", SystemSht.Range("NumStr").Offset(auxOffset, 0))
            Call InsertArrayValue(newdoc, fixPre & "/LossFraction", SystemSht.Range("PVLossFrac").Offset(auxOffset, 0))
         
            If pvExists Then
                ' go through each of the pv module's parameter nodes
                For Each pvNode In subarrNode.ChildNodes
                
                    ' Get the column id of the node
                    Set theParam = PV_DatabaseSht.Range("PVParam").Find(pvNode.nodeName, LookIn:=xlValues, LookAt:=xlWhole)
                
                    ' If the node corresponds to a column
                    If Not theParam Is Nothing Then
                        'Get the value of the cell at the column and row Index
                        Check = PV_DatabaseSht.Cells(SystemSht.Range("PVDataIndex").Offset(auxOffset, 0).Value + PVDataHeight, theParam.Column).Value
                        ' If the values do not match inform the user
                        If pvNode.HasChildNodes Then
                            If (Not (Check = pvNode.ChildNodes.NextNode.NodeValue)) And Not pvNode.nodeName = "Origin" And Not pvNode.nodeName = "Model" And Not pvNode.nodeName = "Manufacturer" Then
                                Call ErrorLogger("Module Model " & Model & ": Discrepancy at " & pvNode.nodeName & ". Value is " & pvNode.ChildNodes.NextNode.NodeValue & " instead of " & Worksheets("PV Module Database").Cells(SystemSht.Range("PVDataIndex").Offset(auxOffset, 0).Value + PVDataHeight, theParam.Column) & ".")
                            End If
                        End If
                    End If
                Next
            End If
        End If
    
        Set subarrNode = newdoc.SelectSingleNode("//Site/System/SubArray" & i & "/Inverter")
    
        If subarrNode.HasChildNodes Then
        
            ' Get inverter model, manufacturer and origin
        
            invExists = True
            Set sysNode = newdoc.SelectSingleNode("//Site/System/SubArray" & i & "/Inverter/Model")
            If sysNode.HasChildNodes Then
                Model = sysNode.ChildNodes.NextNode.NodeValue
            Else
                invExists = False
            End If
        
            Set sysNode = newdoc.SelectSingleNode("//Site/System/SubArray" & i & "/Inverter/Manufacturer")
            If sysNode.HasChildNodes Then
                Manu = sysNode.ChildNodes.NextNode.NodeValue
            Else
                invExists = False
            End If
        
            Set sysNode = newdoc.SelectSingleNode("//Site/System/SubArray" & i & "/Inverter/Origin")
            If sysNode.HasChildNodes Then
                Source = sysNode.ChildNodes.NextNode.NodeValue
            Else
                invExists = False
            End If
        
            ' NB: Gives value to file name, unless file name does not exist in csyx document
            Set sysNode = newdoc.SelectSingleNode("//Site/System/SubArray" & i & "/Inverter/FileName")
            If Not sysNode Is Nothing Then
                If sysNode.HasChildNodes Then
                    fileName = sysNode.ChildNodes.NextNode.NodeValue
                Else:
                    fileName = ""
                End If
            Else
                fileName = ""
            End If
            
            
                        
            If invExists Then
                ' Get the Index of the pv module with the module, manufacturer and version values
                getIndex = InvIndex(Manu, Model, Source)

                ' If the Inverter is Found.
                If Not getIndex = 0 Then
                    SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value = getIndex
                Else
                    ' If not set invExists's cell to false and tell user that it does not exist
                    invExists = False
                    ' NB: Removed line: Call ErrorLogger("Warning: The Inverter " & Model & " was not found. This Inverter has been temporarily added to the database.")
                    lastRow = Inverter_DatabaseSht.Range("A" & Inverter_DatabaseSht.Rows.count).End(xlUp).row + 1
                    ColumnIndex = 8
                
                    Inverter_DatabaseSht.Cells(lastRow, 1).Value = "User_Added"
                    Inverter_DatabaseSht.Cells(lastRow, 2).Value = Manu
                    Inverter_DatabaseSht.Cells(lastRow, 3).Value = Model
                    ' NB: Displays inverter file name. If inverter file isn't in csyx csyx file name is displayed
                    If fileName = "" Then
                        Inverter_DatabaseSht.Cells(lastRow, 4).Value = IntroSht.Range("IntroFileName").Value2
                    Else:
                        Inverter_DatabaseSht.Cells(lastRow, 4).Value = fileName
                    End If
                
                    ' NB: determines which Node the parameters start on depending on whether or not fileName exists
                    If fileName = "" Then
                        ' NB: Changing so values populate correct fields every time
                        ' NB: If file name node exists, the inverter parameters are found the same way as before
                        For j = 3 To subarrNode.ChildNodes.length - 1 Step 1
                
                            If subarrNode.ChildNodes(j).nodeName = "Efficiency" Then
                                For Each efficiencyCurve In subarrNode.ChildNodes(j).ChildNodes
                                    ' If the inverter has more than one efficiency curve
                                    If Not efficiencyCurve.nodeName = "EffCurve" Then
                                        If efficiencyCurve.nodeName = "Low" Then
                                            Inverter_DatabaseSht.Cells(lastRow, 76).Value = efficiencyCurve.Attributes.Item(0).text
                                            ColumnIndex = 77
                                        End If
                                        If efficiencyCurve.nodeName = "Med" Then
                                            Inverter_DatabaseSht.Cells(lastRow, 93).Value = efficiencyCurve.Attributes.Item(0).text
                                            ColumnIndex = 94
                                        End If
                                        If efficiencyCurve.nodeName = "High" Then
                                            Inverter_DatabaseSht.Cells(lastRow, 110).Value = efficiencyCurve.Attributes.Item(0).text
                                            ColumnIndex = 111
                                        End If
                                    End If
                                    
                                    For Each efficiencyCurveParam In efficiencyCurve.ChildNodes
                                        If ColumnIndex = 41 Then ColumnIndex = ColumnIndex + 1
                                   
                                        Inverter_DatabaseSht.Cells(lastRow, ColumnIndex).Value = efficiencyCurveParam.text
            
                                        ColumnIndex = ColumnIndex + 1
                                        ' Skip Res
                                        If ColumnIndex = 44 Then ColumnIndex = ColumnIndex + 1
                                    
                                        ' Adding Bipolar inputs into 'Technological General Features' column
                                        If ColumnIndex = 59 Then Inverter_DatabaseSht.Cells(lastRow, 72).Value = subarrNode.ChildNodes(j).NextSibling.text
                                    
                                    Next
                                Next
                            Else
                         
                                ' If loop has reached the last column then stop inputting more information
                                If subarrNode.ChildNodes(j).nodeName = "BipolarInput" Then
                                    Inverter_DatabaseSht.Cells(lastRow, 72).Value = subarrNode.ChildNodes(j).text
                                    Exit For
                                End If
                        
                                    Inverter_DatabaseSht.Cells(lastRow, ColumnIndex).Value = subarrNode.ChildNodes(j).text
                                
                                    If ColumnIndex = 40 Then
                                        If subarrNode.ChildNodes(j).text = "True" Then Inverter_DatabaseSht.Cells(lastRow, ColumnIndex).Value = "X"
                                    End If
                                
                                    ' Overwrite 'NumInverters' since that is not supposed to be in the database
                                    If subarrNode.ChildNodes(j).nodeName = "NumInverters" Then
                                       ColumnIndex = ColumnIndex - 1
                                    End If
                        
                                    ' counter incremented to go to next column that needs a value inserted
                                    ColumnIndex = ColumnIndex + 1
                                
                                    ' Skip PNomDC, PMaxDC, Req
                                    If ColumnIndex = 15 Then ColumnIndex = ColumnIndex + 3
                                
                                    ' Skip NomVoltage
                                    If ColumnIndex = 19 Then ColumnIndex = ColumnIndex + 1
                                
            
                                    ' Skip NomCurrent,  MaxCurrent,  Req
                                    If ColumnIndex = 24 Then ColumnIndex = ColumnIndex + 3
                                
                                    ' Skip PMaxBehave, VMppBehave, NbDC, Nb, String,  Master,  NUnits,  DCIsol,  Dcswitch, Acswitch, Disconn, ENS
                                    If ColumnIndex = 28 Then ColumnIndex = ColumnIndex + 12
                                                         
                            End If
                        Next
                    Else:
                        ' NB: adjusts to taking from next column down if File Name node exists
                        For j = 4 To subarrNode.ChildNodes.length - 1 Step 1
                
                            If subarrNode.ChildNodes(j).nodeName = "Efficiency" Then
                                For Each efficiencyCurve In subarrNode.ChildNodes(j).ChildNodes
                                    ' If the inverter has more than one efficiency curve
                                    If Not efficiencyCurve.nodeName = "EffCurve" Then
                                        If efficiencyCurve.nodeName = "Low" Then
                                            Inverter_DatabaseSht.Cells(lastRow, 76).Value = efficiencyCurve.Attributes.Item(0).text
                                            ColumnIndex = 77
                                        End If
                                        If efficiencyCurve.nodeName = "Med" Then
                                            Inverter_DatabaseSht.Cells(lastRow, 93).Value = efficiencyCurve.Attributes.Item(0).text
                                            ColumnIndex = 94
                                        End If
                                        If efficiencyCurve.nodeName = "High" Then
                                            Inverter_DatabaseSht.Cells(lastRow, 110).Value = efficiencyCurve.Attributes.Item(0).text
                                            ColumnIndex = 111
                                        End If
                                    End If
                                    
                                    For Each efficiencyCurveParam In efficiencyCurve.ChildNodes
                                        If ColumnIndex = 41 Then ColumnIndex = ColumnIndex + 1
                                   
                                        Inverter_DatabaseSht.Cells(lastRow, ColumnIndex).Value = efficiencyCurveParam.text
                
                                        ColumnIndex = ColumnIndex + 1
                                        ' Skip Res
                                        If ColumnIndex = 44 Then ColumnIndex = ColumnIndex + 1
                                    
                                        ' Adding Bipolar inputs into 'Technological General Features' column
                                        If ColumnIndex = 59 Then Inverter_DatabaseSht.Cells(lastRow, 72).Value = subarrNode.ChildNodes(j).NextSibling.text
                                
                                    Next
                                Next
                            Else
                         
                                ' If loop has reached the last column then stop inputting more information
                                If subarrNode.ChildNodes(j).nodeName = "BipolarInput" Then
                                    Inverter_DatabaseSht.Cells(lastRow, 72).Value = subarrNode.ChildNodes(j).text
                                    Exit For
                                End If
                    
                                    Inverter_DatabaseSht.Cells(lastRow, ColumnIndex).Value = subarrNode.ChildNodes(j).text
                                
                                    If ColumnIndex = 40 Then
                                        If subarrNode.ChildNodes(j).text = "True" Then Inverter_DatabaseSht.Cells(lastRow, ColumnIndex).Value = "X"
                                    End If
                                
                                    ' Overwrite 'NumInverters' since that is not supposed to be in the database
                                    If subarrNode.ChildNodes(j).nodeName = "NumInverters" Then
                                       ColumnIndex = ColumnIndex - 1
                                    End If
                        
                                    ' counter incremented to go to next column that needs a value inserted
                                    ColumnIndex = ColumnIndex + 1
                                
                                    ' Skip PNomDC, PMaxDC, Req
                                    If ColumnIndex = 15 Then ColumnIndex = ColumnIndex + 3
                                
                                    ' Skip NomVoltage
                                    If ColumnIndex = 19 Then ColumnIndex = ColumnIndex + 1
                                
            
                                    ' Skip NomCurrent,  MaxCurrent,  Req
                                    If ColumnIndex = 24 Then ColumnIndex = ColumnIndex + 3
                                
                                    ' Skip PMaxBehave, VMppBehave, NbDC, Nb, String,  Master,  NUnits,  DCIsol,  Dcswitch, Acswitch, Disconn, ENS
                                    If ColumnIndex = 28 Then ColumnIndex = ColumnIndex + 12
                                                     
                            End If
                        Next
                    End If
                         
                    SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value = lastRow - 5
                End If
            End If
            getIndex = InvIndex(Manu, Model, Source)
            'NB: does not change InvDataIndex, which was causing a lack of inverter info to display incorrect inverter
            If Not getIndex = 0 Then
                SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value = getIndex
            End If
        
            Call InsertArrayValue(newdoc, "//Site/System/SubArray" & i & "/Inverter/NumInverters", SystemSht.Range("NumInv").Offset(auxOffset, 0))
            Call InsertArrayValue(newdoc, "//Site/System/SubArray" & i & "/Inverter/LossFraction", SystemSht.Range("InvLossFrac").Offset(auxOffset, 0))
        
            If invExists Then
                ' go through each of the inverter's nodes
                For Each invNode In subarrNode.ChildNodes

                    ' Get the column id of the node
                    Set theParam = Inverter_DatabaseSht.Range("InvParam").Find(invNode.nodeName, LookIn:=xlValues, LookAt:=xlWhole)
                
                    ' If the node corresponds to a column
                    If Not theParam Is Nothing Then
                        'Get the value of the cell at the column and row Index
                        Check = Inverter_DatabaseSht.Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, theParam.Column)
                    
                        ' if it doesn't match tell user
                        If invNode.HasChildNodes Then
                            If Not (Check = invNode.ChildNodes.NextNode.NodeValue) And Not invNode.nodeName = "Origin" And Not invNode.nodeName = "Model" And Not invNode.nodeName = "Manufacturer" Then
                                Call ErrorLogger("Inverter Model " & Model & ": Discrepancy at " & invNode.nodeName & ". Value is " & invNode.ChildNodes.NextNode.NodeValue & " instead of " & Worksheets("Inverter Database").Cells(SystemSht.Range("InvDataIndex").Offset(auxOffset, 0).Value + InvDataHeight, theParam.Column) & ".")
                            End If
                        End If
                    End If
                Next
            End If
        End If
    Next i
    
End Sub

Private Sub loadASTMSht(ByRef newdoc As DOMDocument60)
        Call InsertValue(newdoc, "//Site/ASTMRegress/SystemPmax", AstmSht.Range("SystemPmax"))
        Call InsertValue(newdoc, "//Site/ASTMRegress/ASTMCoeffs/*", AstmSht.Range("ASTMCoeffs"))
        'Call InsertValue(newdoc, "//Site/ASTMRegress/EAF/*", AstmSht.Range("AstmMonthList"))
        AstmSht.Range("EAFJan").Value = newdoc.SelectSingleNode("//Site/ASTMRegress/EAF/Jan").text
        AstmSht.Range("EAFFeb").Value = newdoc.SelectSingleNode("//Site/ASTMRegress/EAF/Feb").text
        AstmSht.Range("EAFMar").Value = newdoc.SelectSingleNode("//Site/ASTMRegress/EAF/Mar").text
        AstmSht.Range("EAFApr").Value = newdoc.SelectSingleNode("//Site/ASTMRegress/EAF/Apr").text
        AstmSht.Range("EAFMay").Value = newdoc.SelectSingleNode("//Site/ASTMRegress/EAF/May").text
        AstmSht.Range("EAFJun").Value = newdoc.SelectSingleNode("//Site/ASTMRegress/EAF/Jun").text
        AstmSht.Range("EAFJul").Value = newdoc.SelectSingleNode("//Site/ASTMRegress/EAF/Jul").text
        AstmSht.Range("EAFAug").Value = newdoc.SelectSingleNode("//Site/ASTMRegress/EAF/Aug").text
        AstmSht.Range("EAFSep").Value = newdoc.SelectSingleNode("//Site/ASTMRegress/EAF/Sep").text
        AstmSht.Range("EAFOct").Value = newdoc.SelectSingleNode("//Site/ASTMRegress/EAF/Oct").text
        AstmSht.Range("EAFNov").Value = newdoc.SelectSingleNode("//Site/ASTMRegress/EAF/Nov").text
        AstmSht.Range("EAFDec").Value = newdoc.SelectSingleNode("//Site/ASTMRegress/EAF/Dec").text
       
End Sub

' Loads parameters of iterative mode sheet if iterative mode node is present in csyx file
Private Sub loadIterativeSht(ByRef newdoc As DOMDocument60)

''--------Commenting out Iterative Functionality for this version--------'

'    ' Unhide iterative sheet and hide output file sheet
'    IterativeSht.Visible = xlSheetVisible
'    OutputFileSht.Visible = xlSheetHidden
'
'    Dim c As Range
'    Call InsertAttribute(newdoc, "//Site/Iterations/Iteration1", "ParamPath", IterativeSht.Range("ParamPath"))
'    Call InsertAttribute(newdoc, "//Site/Iterations/Iteration1", "Start", IterativeSht.Range("Start"))
'    Call InsertAttribute(newdoc, "//Site/Iterations/Iteration1", "End", IterativeSht.Range("End"))
'    Call InsertAttribute(newdoc, "//Site/Iterations/Iteration1", "Interval", IterativeSht.Range("Interval"))
'
'    'loading paramName from paramPath
'    Range("ParamName").Value = IterativeSht.Range("Z" & Application.WorksheetFunction.Match(Range("ParamPath").Value, IterativeSht.Range("AA:AA"), 0)).Value
'
'    ' Load output parameter
'    For Each c In Range("OutputParam")
'        If Not newdoc.SelectSingleNode("//Site/Iterations" & c.Name.Name) Is Nothing Then
'            IterativeSht.Range("IterativeOutputParam").Value = OutputFileSht.Range("B" & c.row).Value
'            Exit Sub
'        End If
'    Next
End Sub


' InsertValue Function
'
' Arguments
'
' newDoc - the XML Document
' node - holds the selected node
' path - the path to the node
' insertLocations - the fields that the node values are being inserted to
' xmLNodeList - contains the list of all nodes in the selected path
' aCell - an individual cell in the insertLocations range
    
' The purpose of this function is to get the values of all nodes in the selected path
' and insert the values into their corresponding fields specified by insertLocations
' End the path with '*' when selecting child nodes so that it properly selects childnodes with SelectNodes
    
' This function works regardless of version since it is based on name checking.
Private Sub InsertValue(ByRef newdoc As DOMDocument60, ByVal path As String, ByVal insertLocations As Range, Optional ByVal insertOutputs As Boolean, Optional ByVal surpressAlerts As Boolean)
    
    Dim xmlNodelist As IXMLDOMNodeList
    Dim aCell As Range
    Dim cellName As String
    Dim loadStatus As LOAD_STATUS
    Dim shtStatus As sheetStatus
    Dim child As IXMLDOMNode
    Dim i As Integer
    i = 0
    
    Set xmlNodelist = newdoc.SelectNodes(path)
    
    'Loops through each cell in the range of cells where information is to be inserted
    For Each aCell In insertLocations
    
        cellName = ExtractCellName(aCell)
        ' If the cell name and node name in the XML file match, the value is inserted into the correct place
        If Not xmlNodelist.Item(i) Is Nothing Then
            If cellName = xmlNodelist.Item(i).nodeName Then
                aCell.Value = xmlNodelist.Item(i).text
                loadStatus = ItemFound
            Else
                ' The node values were in the incorrect order or did not exist.
                loadStatus = ItemNotFound
                ' Another For Each loop is used to see if the node actually exists, just in a different location
                For Each child In xmlNodelist
                    If cellName = child.nodeName Then
                        aCell.Value = child.text
                        loadStatus = ItemFound
                    End If
                Next
            End If
            i = i + 1
            If loadStatus = ItemNotFound Then
                If surpressAlerts <> True Then
                    Call ErrorLogger("Field not found: " & cellName)
                End If
                'If the value was not found, Index i is reset back to zero to prevent error 91 (Object not set) when i is greater than the max Index
                i = 0
            End If
        Else
            Call ErrorLogger("Field not found: " & cellName)
        End If
    Next
    
End Sub
    
' InsertArrayValue function
'
' The same as the insertvalue function except it does not check node names, since the number of arrays is dynamic and subsequent arrays do not
' have names
Private Sub InsertArrayValue(ByRef newdoc As DOMDocument60, ByVal path As String, ByVal insertLocations As Range)
    
    Dim xmlNodelist As IXMLDOMNodeList
    Dim aCell As Range
    Dim shtStatus As sheetStatus
    
    Set xmlNodelist = newdoc.SelectNodes(path)
    
    For Each aCell In insertLocations
        aCell.Value = xmlNodelist.NextNode.text
    Next
     
End Sub
    
' InsertValue_BoolToYesNo
'
' This subroutine is responsible for loading values from the CSYX (XML) file whose values are booleans, and changing them to 'Yes'/'No'
' on the page where they are loaded

Private Sub InsertValue_BoolToYesNo(ByRef newdoc As DOMDocument60, ByVal path As String, ByVal insertLocations As Range)

    Dim xmlNodelist As IXMLDOMNodeList
    Dim child As IXMLDOMNode
    Dim aCell As Range
    Dim cellName As String
    Dim loadStatus As LOAD_STATUS
    Dim i As Integer
    
    Set xmlNodelist = newdoc.SelectNodes(path)

    ' If the output list cannot be found in the .csyx file
    If xmlNodelist.length = 0 Then
        MsgBox "The available output list is not consistent with the current version. CASSYS has stopped loading the output definition."
        Exit Sub
    End If
    
    For Each aCell In insertLocations
        cellName = ExtractCellName(aCell)
        
        ' Just in case some nodes are missing, the Index does not exceed the number of nodes in the node list
        If i = xmlNodelist.length Then
            i = 0
        End If

        If cellName = xmlNodelist.Item(i).nodeName Then
            loadStatus = ItemFound
            ' load normally (changing bools to yes and no)
            If xmlNodelist.Item(i).text = "True" Then
                aCell.Value = "Yes"
            Else
                aCell.Value = "No"
            End If
        Else
            ' In case the nodes are in a different order
            ' the For Each loop checks the rest of the sibling nodes
            loadStatus = ItemNotFound
            For Each child In xmlNodelist
                If cellName = child.nodeName Then
                    loadStatus = ItemFound
                    If xmlNodelist.Item(i).text = "True" Then
                        aCell.Value = "Yes"
                    Else
                        aCell.Value = "No"
                    End If
                End If
            Next
        End If
        i = i + 1
        If loadStatus = ItemNotFound Then Call ErrorLogger("Field not found: " & cellName)
    Next

End Sub


' InsertValue_OutputSelections Function
'
' Arguments
'
' newDoc - the XML Document
' node - holds the selected node
' path - the path to the node
' insertLocations - the fields that the node values are being inserted to
' xmLNodeList - contains the list of all nodes in the selected path
' aCell - an individual cell in the insertLocations range
    
' The purpose of this function is to get the "True/False" values of all nodes
' in the selected path, change them to "Yes/No" respectively
' and insert the values into their corresponding fields specified by insertLocations
' End the path with '*' so that it properly selects childnodes with SelectNodes

'NB: Got rid of call to error logger if output parameters were missing 29/01/2016
    
    
Private Sub InsertValue_OutputSelections(ByRef newdoc As DOMDocument60, ByVal path As String, ByVal insertLocations As Range)

    Dim outputList As IXMLDOMNodeList
    Dim child As IXMLDOMNode
    Dim aCell As Range
    Dim cellName As String
    Dim loadStatus As LOAD_STATUS
    Dim i As Integer
    Dim j As Integer
    
    Set outputList = newdoc.SelectNodes(path)
    
    ' If the output list cannot be found in the .csyx file
    If outputList.length = 0 Then
        MsgBox "The available output list is not consistent with the current version. CASSYS has stopped loading the output definition."
        Exit Sub
    End If

    For Each aCell In insertLocations
        cellName = ExtractCellName(aCell)
        
        ' If csyx version < 1.3.0, then account for name changes made for v1.3.0
        If (StrComp(newdoc.SelectSingleNode("//Site/Version").text, "1.3.0") < 0) Then
            If cellName = "Inverter_Loss_Due_to_High_Voltage_Threshold" Then
                cellName = "Inverter_Loss_Due_to_Nominal_Inv._Voltage"

            ElseIf cellName = "Inverter_Loss_Due_to_High_Power_Threshold" Then
                cellName = "Inverter_Loss_Due_to_Nominal_Inv._Power"

            ElseIf cellName = "Inverter_Loss_Due_to_Low_Voltage_Threshold" Then
                cellName = "Inverter_Loss_Due_to_Voltage_Threshold"

            ElseIf cellName = "Inverter_Loss_Due_to_Low_Power_Threshold" Then
                cellName = "Inverter_Loss_Due_to_Power_Threshold"

            End If
        End If
        
        ' Just in case some nodes are missing, the Index does not exceed the number of nodes in the node list
        If i = outputList.length Then
            i = 0
        End If
        
        If cellName = outputList.Item(i).nodeName Then
            loadStatus = ItemFound
            ' loading outputs, have to check attributes as well before deciding what to load
            If outputList.Item(i).text = "True" Then
                If Not outputList.Item(i).Attributes.getNamedItem("Summarize") Is Nothing Then
                    If outputList.Item(i).Attributes.getNamedItem("Summarize").text = "True" Then
                        aCell.Value = "Summarize"
                    Else
                        aCell.Value = "Detail"
                    End If
                Else
                    aCell.Value = "Detail"
                End If
            Else
                aCell.Value = "-"
            End If
        Else
            loadStatus = ItemNotFound
            ' NB: Using j instead of i since i does not necessarily correspond to the correct output parameter
            ' NB: j ensures that the "Summarize" or "Detail" property is loaded correctly as it always corresponds tothe correct parameter within the "for" loop
            j = 0
            For Each child In outputList
                If cellName = child.nodeName Then
                    loadStatus = ItemFound
                    ' loading outputs, have to check attributes as well before deciding what to load
                    If outputList.Item(j).text = "True" Then
                        If Not outputList.Item(j).Attributes.getNamedItem("Summarize") Is Nothing Then
                            If outputList.Item(j).Attributes.getNamedItem("Summarize").text = "True" Then
                                aCell.Value = "Summarize"
                            Else
                                aCell.Value = "Detail"
                            End If
                        Else
                            aCell.Value = "Detail"
                        End If
                    Else
                        aCell.Value = "-"
                    End If
                End If
                j = j + 1
            Next
        End If
        i = i + 1
    Next
 
    ' For compatibility with previous versions: force Input_Timestamp and Timestamp_Used_for_Simulation to 'Detail'
    Range("Input_Timestamp") = "Detail"
    Range("Timestamp_Used_for_Simulation") = "Detail"

End Sub
    
' InsertAttribute Function
'
' Arguments
'
' newDoc - the XML Document
' node - holds the selected node
' Path - the path to the node
' attr - the attribute name
' insertLocations - the field that the node value is being inserted to
'
' The purpose of this function is to get the value of an attribute
' of a specific node and insert it into the range of the
' corresponding field
' for example 'Yearly' and 'Monthly' are attributes that specify
' the current selection of the user and consequently what to save and load
    
Private Sub InsertAttribute(ByRef newdoc As DOMDocument60, ByVal path As String, ByVal attr As String, ByVal insertLocations As Range)

    Dim shtStatus As sheetStatus
    Dim node As IXMLDOMNode

    Set node = newdoc.SelectSingleNode(path)
    
    ' Get the attribute of the selected node
    If Not node Is Nothing Then
        If node.HasChildNodes Or Worksheets("Iterative Mode").Visible = True Then
            insertLocations.Value = node.Attributes.getNamedItem(attr).text
        End If
    End If
    
End Sub

' ErrorLogger Function
'
' Logs errors onto the Error Sheet while loading

Sub ErrorLogger(ByVal errorMessage As String)

    Dim lastRow As Integer
    
    If Application.WorksheetFunction.CountA(ErrorSht.Range("A7", "A" & Rows.count)) <> 0 Then
        ErrorSht.Range("LastRow").Offset(1, 0).Name = "LastRow"
        lastRow = ErrorSht.Range("LastRow").row
        ErrorSht.Range("A" & lastRow).Value = errorMessage
    Else
        ' The first line logs the file name of the site being loaded (value2 is needed since the file name comes from a formula)
        ErrorSht.Range("A7").Formula = "=""Notable events encountered while loading "" & IntroFileName & "" :"""
        ErrorSht.Range("A9").Value = errorMessage
        ErrorSht.Range("A9").Name = "LastRow"
    End If
    
    ErrorSht.Range("ErrorsEncountered").Value = "True"
    
End Sub




