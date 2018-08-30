Attribute VB_Name = "DatabaseImportModule"
'                   DATABASE IMPORT MODULE                  '
'==========================================================='
' This module contains code to import .PAN or .OND files to '
' the inverter or module databases by decrypting the Real48 '
' format that these files are encoded in. The first part of '
' a PAN file is not encrypted; this section contains the    '
' general information such as the panel model and source.   '
' After this section is a section encrypted in the Real48   '
' Delphi type, which must be parsed every 6 bits to return  '
' a meaningful number.                                      '

Option Explicit
Private Const numParams = 27 ' number of parameters for importing a PV module
Private Const CRMarker = "D" 'The hexadecimal byte representation of a carriage return <CR>
Private Const semiColonMarker = "3B" ' The hexadecimal byte representation of a semi-colon(;)
Private Const dotMarker = "9" ' The hexadecimal byte representation of a '.'
Private Const doubleDotMarker = "A" ' The hexadecimal byte representation of '..'
Private Const forwardSlashMarker = "2F" ' The hexadecimal byte representation of '/'
Private Const verticalBarMarker = "A6" ' The hexadecimal byte representation of a vertical bar with a break ()
Private Const MPPTTrue = "1" ' If the byte at the MPPT check position is 02, then the inverter operates on MPPT
Private Const valueFormat = "0.000" ' The format specifier used to format the parsed values
Private Const databaseStandardEffStart = 41 ' The column index on the inverter database sheet where the standard (single) efficiency curve starts
Private Const databaseLowEffStart = 76 ' The column index on the inverter database sheet where the low voltage efficiency curve starts
Private Const databaseMedEffStart = 93 ' The column index on the inverter database sheet where the medium voltage efficiency curve starts
Private Const databaseHighEffStart = 110 ' The column index on the inverter database sheet where the high voltage efficiency curve starts

' ParsePANFile function
'
'
' This module is responsible for deciding which type of PAN file was loaded and calling a respective function to parse
' This function retuns 1 if module added, -1 is module replaced and 0 if module skipped
Function ParsePANFile(ByVal PANFilePath As String, ByRef dupModuleRepeat As Integer) As Integer
    
    Dim firstLine As String
    
    On Error GoTo invalidPANFile
    
    ' Open and read first line of PAN File
    Open PANFilePath For Input As #1
    Line Input #1, firstLine
    Close #1
    
    If (StrComp(firstLine, "PVObject_=pvModule") = 0) Then
        ParsePANFile = ParseTextPANFile(PANFilePath, dupModuleRepeat)
    Else
        ParsePANFile = ParseBinaryPANFile(PANFilePath, dupModuleRepeat)
    End If
    
    Exit Function:
    
invalidPANFile:
    On Error GoTo 0
    ParsePANFile = 0
    
End Function

' ParseTextPANFile function
'
'
' This module is responsible for parsing text based PAN files and populating the module database sheet
' This particular format of PAN file is used by PvSyst in versions greater than or equal to 6.4 and follow a format similar to YAML
' This function retuns 1 if module added, -1 if module replaced and 0 if module skipped
Function ParseTextPANFile(ByVal PANFilePath As String, ByRef dupModuleRepeat As Integer) As Integer
    
    Dim textLine As String
    Dim lineArray() As String
    Dim currentDict As String
    Dim modRow As Integer                           ' row where the module data is going to be copied
    Dim IAMPoint() As String
    Dim currentShtStatus As sheetStatus             ' Sheet status used for Pre/Post modifying
    Dim IAMDefined As Boolean                       ' IAM is defined for module
    Dim OperPoints As Boolean                       ' Used to skip over Oper Points section as they are not needed in module database
    
    'Define Dictionaries (require Reference: Microsoft Scripting Runtime)
    Dim PvModule As New Scripting.Dictionary
    Dim PvCommercial As New Scripting.Dictionary
    Dim Technol As New Scripting.Dictionary
    Dim PVIAM As New Scripting.Dictionary
    
    Dim DupModuleAction As Integer                  ' Action selected by user if a duplicate module is found (1 = overwrite, 2 = skip)
    Dim Pos As Integer                              ' Position of a substring within a string
    Dim PVsyst_ver As String                        ' Version of PVsyst read from the file name rather than from the file itself
    
    ParseTextPANFile = False
    On Error GoTo invalidPANFile
    
    Open PANFilePath For Input As #1
    
    IAMDefined = False
    currentDict = "PvModule"
    Do Until EOF(1)
        Line Input #1, textLine
        'Debug.Print (textLine)
        If (InStr(textLine, "pvModule") <> 0) Then
            currentDict = "PvModule"
        ElseIf (InStr(textLine, "pvCommercial") <> 0) Then
            currentDict = "PvCommercial"
        ElseIf (InStr(textLine, "Technol") <> 0) Then
            currentDict = "technol"
        ElseIf (InStr(textLine, "PVObject_IAM") <> 0) Then
            currentDict = "PVIAM"
            IAMDefined = True
        ElseIf (InStr(textLine, "OperPoints, list") <> 0) Then
            OperPoints = True
        ElseIf (InStr(textLine, "End of List OperPoints") <> 0) Then
            OperPoints = False
        End If
        
        ' Populating dictionaries textLine is not currently in operPoints section
        If (InStr(textLine, "=") <> 0 And Not OperPoints) Then
            lineArray = Split(textLine, "=")
            Select Case currentDict
                Case "PvModule"
                    PvModule.Add Trim(lineArray(0)), Trim(lineArray(1))
                Case "PvCommercial"
                    PvCommercial.Add Trim(lineArray(0)), Trim(lineArray(1))
                Case "technol"
                    Technol.Add Trim(lineArray(0)), Trim(lineArray(1))
                Case "PVIAM"
                    PVIAM.Add Trim(lineArray(0)), Trim(lineArray(1))
            End Select
        End If
    Loop
    
    Close #1
    
    ' Canadian Solar modules exported from our own database: PVsyst version can be in the name of the file
    ' rather than in the file itself
    If (PANFilePath Like "*_V#[._]##*") Then
        Pos = 1
        Do
            Pos = InStr(Pos, PANFilePath, "_V")
            PVsyst_ver = Mid(PANFilePath, Pos + 2, 4)
        Loop While Mid(PVsyst_ver, 2, 1) <> "_" And Mid(PVsyst_ver, 2, 1) <> "."
        If Mid(PVsyst_ver, 2, 1) = "_" Then
            PVsyst_ver = Replace(PVsyst_ver, "_", ".")
        End If
        PvModule("Version") = PVsyst_ver
    End If
          
    ' Check if the module already exists, based on Manufacturer, Model, and Origin
    Dim getIndex As Integer ' the Index of the module
    getIndex = PVIndex(PvCommercial("Manufacturer"), PvCommercial("Model"), "PVsyst " & PvModule("Version"))
    
    ' Initialize row where to write module to zero
    modRow = 0
    
    ' If module already exists, ask whether to overwrite or skip
    Dim overWrite As Integer            ' 1 if existing module should be replaced, 2 if skipped
    
    If getIndex <> 0 Then
        If dupModuleRepeat = 0 Then
            ' Show the form
            UF_AddSameModuleOptions.CheckBox1 = False
            UF_AddSameModuleOptions.Label2.Caption = "PVsyst " & PvModule("Version") & " | " & PvCommercial("Manufacturer") & " | " & PvCommercial("Model")
            UF_AddSameModuleOptions.Show
            ' Retrieve user choice
            overWrite = UF_AddSameModuleOptions.getChoice()
            If UF_AddSameModuleOptions.getChoiceRepeat() Then
                dupModuleRepeat = overWrite                  ' Remember which option was selected
            End If
        Else
            overWrite = dupModuleRepeat
        End If
        
        ' If overwrite option is selected, module will be copied at same index
        If (overWrite = 1) Then
            modRow = getIndex + 4
            ParseTextPANFile = -1
        Else
            ParseTextPANFile = 0
        End If
        
    ' If module doesn't exist, it will be added to end of table
    Else
        modRow = PV_DatabaseSht.Range("A" & Rows.count).End(xlUp).row + 1
        ParseTextPANFile = 1
    End If
    
    ' Add module
    If (modRow <> 0) Then
            
        Call PreModify(PV_DatabaseSht, currentShtStatus)
            
        On Error Resume Next
        PV_DatabaseSht.Cells(modRow, 1).Value = "PVsyst " & PvModule("Version")
        PV_DatabaseSht.Cells(modRow, 2).Value = PvCommercial("Manufacturer")
        PV_DatabaseSht.Cells(modRow, 3).Value = PvCommercial("Model")
        PV_DatabaseSht.Cells(modRow, 4).Value = Right(PANFilePath, Len(PANFilePath) - InStrRev(PANFilePath, "\"))
        PV_DatabaseSht.Cells(modRow, 5).Value = PvCommercial("DataSource")
        PV_DatabaseSht.Cells(modRow, 6).Value = PvCommercial("YearBeg")
        PV_DatabaseSht.Cells(modRow, 7).Value = PvCommercial("YearEnd")
        PV_DatabaseSht.Cells(modRow, 8).Value = Technol("PNom")
        PV_DatabaseSht.Cells(modRow, 9).Value = Technol("PNomTolLow")
        PV_DatabaseSht.Cells(modRow, 10).Value = Technol("PNomTolUp")
        PV_DatabaseSht.Cells(modRow, 11).Value = Technol("LIDLoss")
        PV_DatabaseSht.Cells(modRow, 12).Value = Technol("Technol")
        PV_DatabaseSht.Cells(modRow, 13).Value = Technol("NCelS")
        PV_DatabaseSht.Cells(modRow, 14).Value = Technol("NCelP")
        PV_DatabaseSht.Cells(modRow, 15).Value = Technol("GRef")
        PV_DatabaseSht.Cells(modRow, 16).Value = Technol("TRef")
        PV_DatabaseSht.Cells(modRow, 17).Value = Technol("Vmp")
        PV_DatabaseSht.Cells(modRow, 18).Value = Technol("Imp")
        PV_DatabaseSht.Cells(modRow, 19).Value = Technol("Voc")
        PV_DatabaseSht.Cells(modRow, 20).Value = Technol("Isc")
        PV_DatabaseSht.Cells(modRow, 21).Value = Technol("muISC")
        PV_DatabaseSht.Cells(modRow, 22).Value = Technol("muVocSpec")
        PV_DatabaseSht.Cells(modRow, 23).Value = Technol("muPmpReq")
        PV_DatabaseSht.Cells(modRow, 24).Value = Technol("Rp_0")
        PV_DatabaseSht.Cells(modRow, 25).Value = Technol("Rp_Exp")
        PV_DatabaseSht.Cells(modRow, 26).Value = Technol("RShunt")
        PV_DatabaseSht.Cells(modRow, 27).Value = Technol("RSerie")
        PV_DatabaseSht.Cells(modRow, 28).Value = Technol("Gamma")
        PV_DatabaseSht.Cells(modRow, 29).Value = Technol("muGamma")
        PV_DatabaseSht.Cells(modRow, 30).Value = Technol("RelEffic800")
        PV_DatabaseSht.Cells(modRow, 31).Value = Technol("RelEffic700")
        PV_DatabaseSht.Cells(modRow, 32).Value = Technol("RelEffic600")
        PV_DatabaseSht.Cells(modRow, 33).Value = Technol("RelEffic400")
        PV_DatabaseSht.Cells(modRow, 34).Value = Technol("RelEffic200")

        PV_DatabaseSht.Cells(modRow, 35).Value = Technol("VMaxIEC")
        PV_DatabaseSht.Cells(modRow, 36).Value = Technol("NDiode")
        PV_DatabaseSht.Cells(modRow, 37).Value = Technol("VRevDiode")
        PV_DatabaseSht.Cells(modRow, 38).Value = Technol("BRev")

        PV_DatabaseSht.Cells(modRow, 39).Value = CDbl(PvCommercial("Height")) * 1000
        PV_DatabaseSht.Cells(modRow, 40).Value = CDbl(PvCommercial("Width")) * 1000
        PV_DatabaseSht.Cells(modRow, 41).Value = CDbl(PvCommercial("Depth")) * 1000
        PV_DatabaseSht.Cells(modRow, 42).Value = PvCommercial("Weight")
        PV_DatabaseSht.Cells(modRow, 43).Value = CDbl(PvCommercial("Width")) * CDbl(PvCommercial("Height"))

        PV_DatabaseSht.Cells(modRow, 44).Value = Technol("CellArea")
        PV_DatabaseSht.Cells(modRow, 45).Value = CInt(Technol("NCelS")) * CInt(Technol("NCelP")) * CDbl(Technol("CellArea")) / 10000
        
        If IAMDefined Then
            IAMPoint = Split(PVIAM("Point_1"), ",")
            PV_DatabaseSht.Cells(modRow, 47).Value = IAMPoint(0)
            PV_DatabaseSht.Cells(modRow, 48).Value = IAMPoint(1)
            
            IAMPoint = Split(PVIAM("Point_2"), ",")
            PV_DatabaseSht.Cells(modRow, 49).Value = IAMPoint(0)
            PV_DatabaseSht.Cells(modRow, 50).Value = IAMPoint(1)
            
            IAMPoint = Split(PVIAM("Point_3"), ",")
            PV_DatabaseSht.Cells(modRow, 51).Value = IAMPoint(0)
            PV_DatabaseSht.Cells(modRow, 52).Value = IAMPoint(1)
            
            IAMPoint = Split(PVIAM("Point_4"), ",")
            PV_DatabaseSht.Cells(modRow, 53).Value = IAMPoint(0)
            PV_DatabaseSht.Cells(modRow, 54).Value = IAMPoint(1)
            
            IAMPoint = Split(PVIAM("Point_5"), ",")
            PV_DatabaseSht.Cells(modRow, 55).Value = IAMPoint(0)
            PV_DatabaseSht.Cells(modRow, 56).Value = IAMPoint(1)
            
            IAMPoint = Split(PVIAM("Point_6"), ",")
            PV_DatabaseSht.Cells(modRow, 57).Value = IAMPoint(0)
            PV_DatabaseSht.Cells(modRow, 58).Value = IAMPoint(1)
            
            IAMPoint = Split(PVIAM("Point_7"), ",")
            PV_DatabaseSht.Cells(modRow, 59).Value = IAMPoint(0)
            PV_DatabaseSht.Cells(modRow, 60).Value = IAMPoint(1)
            
            IAMPoint = Split(PVIAM("Point_8"), ",")
            PV_DatabaseSht.Cells(modRow, 61).Value = IAMPoint(0)
            PV_DatabaseSht.Cells(modRow, 62).Value = IAMPoint(1)
            
            IAMPoint = Split(PVIAM("Point_9"), ",")
            PV_DatabaseSht.Cells(modRow, 63).Value = IAMPoint(0)
            PV_DatabaseSht.Cells(modRow, 64).Value = IAMPoint(1)
            
            ' Error if 10 IAMPoints are defined, as CASSYS does not know how to handle this
            If (PVIAM("Point_10") <> "") Then
                Call Err.Raise(6001, "CASSYS.DatabaseImportModule.ParseTextPANFile", "PVIAM has more than 9 points. CASSYS will terminate loading PAN file.")
            End If
            
        End If
        
        On Error GoTo 0

        Call PostModify(PV_DatabaseSht, currentShtStatus)
    End If
    
    Set PvModule = Nothing
    Set PvCommercial = Nothing
    Set Technol = Nothing
    Set PVIAM = Nothing

    Exit Function

    ' Return false for an unsuccessful parse
invalidPANFile:
    On Error GoTo 0
    ParseTextPANFile = 0
    
End Function

' ParseBinaryPANFile function
'
'
' This module is responsible for parsing binary PAN files and populating the module database sheet
' This particular format of PAN file is used by PvSyst in versions less than 6.4
' This function retuns 1 if module added, -1 if module replaced and 0 if module skipped
Function ParseBinaryPANFile(ByVal PANFilePath As String, ByRef dupModuleRepeat As Integer) As Integer
    
    Dim byteArray() As Byte ' All the bytes in the file are written into this array; each index represents one hexadecimal byte
    Dim modRow As Integer   ' row where the module data is going to be copied
    'Dim lastRow As Integer ' The last row in the PV module database to decide where to write the new PV module
    Dim currentShtStatus As sheetStatus ' Sheet status used for Pre/Post modifying
    
    ' Counter and index variables for the first part of the file
    Dim i As Integer ' loop index control variable
    Dim CRCounter As Integer ' Counts number of carriage returns
    Dim dotMarkerCounter As Integer
    Dim real48StartIndex As Integer ' denotes the start of the section encoded in Delphi Real48
    Dim manuStartIndex As Integer ' denotes the first byte where the manufacturer is written
    Dim panelStartIndex As Integer ' denotes the first byte where the panel model is written
    Dim sourceStartIndex As Integer ' denotes the first byte where the panel source is written
    Dim versionStartIndex As Integer ' denotes the first byte where the PAN file version is written
    Dim versionEndIndex As Integer ' denotes the last byte that the PAN file version is written
    Dim yearStartIndex As Integer ' denotes the first byte where the year is written
    Dim techStartIndex As Integer ' denotes the first byte where the technology type is written
    Dim cellSinPStartIndex As Integer ' denotes the first byte where the parallel cell specification is written
    Dim cellSinSStartIndex As Integer ' denotes the first byte where the series cell specification size is written
    Dim byPassDiodesStartIndex As Integer ' denotes the first byte where the bypass diodes is written
    Dim iamStartIndex As Integer
    Dim Pos As Integer   ' Position of a substring within a string
    Dim PVsyst_ver As String      ' Version of PVsyst read from the file name rather than from the file itself
    
    
    ' Parsed parameters from the file
    Dim Manufacturer As String
    Dim Model As String
    Dim Source As String
    Dim version As String
    Dim inYear As String
    Dim technology As String
    Dim CellsinS As String
    Dim CellsinP As String
    Dim byPassDiodes As String
    Dim PNom As String
    Dim Gref As String
    Dim Toleranz As String
    Dim AreaM As String
    Dim CellArea As String
    Dim vMax As String
    Dim Tref As String
    Dim Isc As String
    Dim mIsc As String
    Dim Voc As String
    Dim mVco As String
    Dim Impp As String
    Dim Vmpp As String
    Dim BypassDiodeVoltage As String
    Dim Rshunt As String
    Dim Rshexp As String
    Dim mPmpp As String
    Dim Rseries As String
    Dim Rsh0 As String
    Dim aoi(0 To 8) As String
    Dim Modifier(0 To 8) As String
         
    On Error GoTo invalidPANFile ' catches an Index Out of Range exception due to a file being empty or invalid
         
    ' This section of code is used to find the indices of each parameter
    ' in the first section (not in real48 format) since each parameter
    ' can be of variable length. To find out the number to bytes to be
    ' parsed for a desired parameter, the start index of the
    ' desired parameter is subtracted from the start index of the next parameter.
    
    ' Read all bytes from the PAN file into an array
    byteArray = BinaryInputReader(PANFilePath)
    
    manuStartIndex = FindMarkerIndex(semiColonMarker, 0, byteArray)
    panelStartIndex = FindMarkerIndex(dotMarker, 0, byteArray)
    sourceStartIndex = FindMarkerIndex(dotMarker, panelStartIndex, byteArray)
    versionStartIndex = FindMarkerIndex(doubleDotMarker, sourceStartIndex, byteArray)
    versionEndIndex = FindMarkerIndex(semiColonMarker, versionStartIndex, byteArray)
    yearStartIndex = FindMarkerIndex(semiColonMarker, versionEndIndex + 1, byteArray)
    techStartIndex = FindMarkerIndex(doubleDotMarker, yearStartIndex, byteArray)
    cellSinSStartIndex = FindMarkerIndex(semiColonMarker, techStartIndex, byteArray)
    cellSinPStartIndex = FindMarkerIndex(semiColonMarker, cellSinSStartIndex, byteArray)
    byPassDiodesStartIndex = FindMarkerIndex(semiColonMarker, cellSinPStartIndex, byteArray)
    
    
    ' Carriage return counter; the real48 section always starts after 3 occurrences of CR bytes
    CRCounter = 0
    For i = 0 To UBound(byteArray)
        If Hex$(byteArray(i)) = CRMarker Then
            CRCounter = CRCounter + 1
        End If
        
        If CRCounter = 3 Then
            ' Add two to Index of CR (format is <CR><LF><first real48 bit>)
            real48StartIndex = i + 2
            Exit For
        End If
    Next i
    
    ' Parse the first part of the PAN file (not encrypted, bytes can be directly converted from Byte --> String)
    ' The number of bytes to extract for a parameter can be found by
    ' subtracting the start index of the current parameter from the start Index of the following one
    Manufacturer = ByteArrayToString(ExtractByteParameter(byteArray, manuStartIndex, panelStartIndex - manuStartIndex))
    Model = ByteArrayToString(ExtractByteParameter(byteArray, panelStartIndex, sourceStartIndex - panelStartIndex))
    Source = ByteArrayToString(ExtractByteParameter(byteArray, sourceStartIndex, versionStartIndex - sourceStartIndex - 3))
    version = Replace(ByteArrayToString(ExtractByteParameter(byteArray, versionStartIndex, versionEndIndex - versionStartIndex - 1)), "Version", "PVsyst")
    inYear = ByteArrayToString(ExtractByteParameter(byteArray, yearStartIndex, 4))
    technology = ByteArrayToString(ExtractByteParameter(byteArray, techStartIndex, cellSinSStartIndex - techStartIndex - 1))
    CellsinS = ByteArrayToString(ExtractByteParameter(byteArray, cellSinSStartIndex, cellSinPStartIndex - cellSinSStartIndex - 1))
    CellsinP = ByteArrayToString(ExtractByteParameter(byteArray, cellSinPStartIndex, byPassDiodesStartIndex - cellSinPStartIndex - 1))
    byPassDiodes = ByteArrayToString(ExtractByteParameter(byteArray, byPassDiodesStartIndex, 1))
    
    ' Parse Delphi Real48 (Each parameter is always 6 bits starting at this section), and format them to one decimal place
    PNom = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 0), 6)), valueFormat)
    vMax = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 1), 6)), valueFormat)
    Toleranz = Format(100 * Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 2), 6)), "0.0") ' This parameter is displayed as a percentage
    ' In some PAN files the toleranz is 0 but stored as a larger number
    If Toleranz > 100 Then Toleranz = 0#
    AreaM = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 3), 6)), valueFormat)
    CellArea = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 4), 6)), valueFormat)
    Gref = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 5), 6)), valueFormat)
    Tref = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 6), 6)), valueFormat)
    Isc = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 8), 6)), valueFormat)
    mIsc = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 9), 6)), valueFormat)
    Voc = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 10), 6)), valueFormat)
    mVco = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 11), 6)), valueFormat)
    Impp = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 12), 6)), valueFormat)
    Vmpp = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 13), 6)), valueFormat)
    BypassDiodeVoltage = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 14), 6)), valueFormat)
    Rshunt = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 17), 6)), valueFormat)
    Rseries = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 18), 6)), valueFormat)
    Rsh0 = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 23), 6)), valueFormat)
    Rshexp = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 24), 6)), valueFormat)
    mPmpp = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 25), 6)), valueFormat)
    
    If checkValidVersion(version) = False Then
        dotMarkerCounter = 0
        For i = real48StartIndex + 170 To UBound(byteArray)
            If Hex$(byteArray(i)) = dotMarker Then
                dotMarkerCounter = dotMarkerCounter + 1
            End If
        
            If dotMarkerCounter = 2 Then
                ' Add two to index of CR (format is <CR><LF><first real48 bit>)
                iamStartIndex = i + 4
                Exit For
            End If
        Next i
        If dotMarkerCounter = 2 Then
            Call ExtractIAMProfile(iamStartIndex, byteArray, aoi, Modifier)
        End If
    End If
       
    ' Canadian Solar modules exported from our own database: PVsyst version can be in the name of the file
    ' rather than in the file itself
    If (PANFilePath Like "*_V#[_.]##*") Then
        Pos = 1
        ' Loop in case there is more than one _V in the file
        Do
            Pos = InStr(Pos, PANFilePath, "_V")
            PVsyst_ver = Mid(PANFilePath, Pos + 2, 4)
        Loop While Mid(PVsyst_ver, 2, 1) <> "_" And Mid(PVsyst_ver, 2, 1) <> "."
        If Mid(PVsyst_ver, 2, 1) = "_" Then
            PVsyst_ver = Replace(PVsyst_ver, "_", ".")
        End If
        version = "PVsyst " & PVsyst_ver
    End If
   
    ' Write to PV Database
    On Error GoTo 0
     
    ' Check if the module already exists, based on Manufacturer, Model, and Origin
    Dim getIndex As Integer ' the Index of the module
    getIndex = PVIndex(Left(Manufacturer, Len(Manufacturer) - 1), Left(Model, Len(Model) - 1), version)
    
    ' Initialize row where to write module to zero
    modRow = 0
    ' If module already exists, ask whether to overwrite or skip
    Dim overWrite As Integer            ' 1 if existing module should be replaced, 2 if skipped
    
    If getIndex <> 0 Then
        If dupModuleRepeat = 0 Then
            ' Show the form
            UF_AddSameModuleOptions.CheckBox1 = False
            UF_AddSameModuleOptions.Label2.Caption = version & " | " & Left(Manufacturer, Len(Manufacturer) - 1) & " | " & Left(Model, Len(Model) - 1)
            UF_AddSameModuleOptions.Show
            
            ' Retrieve user choice
            overWrite = UF_AddSameModuleOptions.getChoice()
            If UF_AddSameModuleOptions.getChoiceRepeat() Then
                dupModuleRepeat = overWrite                  ' Remember which option was selected
            End If
        Else
            overWrite = dupModuleRepeat
        End If
        
        ' If overwrite option is selected, module will be copied at same index
        If (overWrite = 1) Then
            modRow = getIndex + 4
            ParseBinaryPANFile = -1
        Else
            ParseBinaryPANFile = 0
        End If
        
    ' If module doesn't exist, it will be added to end of table
    Else
        modRow = PV_DatabaseSht.Range("A" & Rows.count).End(xlUp).row + 1
        ParseBinaryPANFile = 1
    End If
    
    
    ' Add module
    If (modRow <> 0) Then
    
    Call PreModify(PV_DatabaseSht, currentShtStatus)
    
    ' Extract the PAN file name from the full path
    ' Updated module database, added colomns(TolToLow,TolToUp,Gamma   muGamma RelEffic800 RelEffic700 RelEffic600 RelEffic400 RelEffic200) and deleted some colomns ( price,etc).

    PV_DatabaseSht.Cells(modRow, 4).Value = Right(PANFilePath, Len(PANFilePath) - InStrRev(PANFilePath, "\"))
    PV_DatabaseSht.Cells(modRow, 2).Value = Left(Manufacturer, Len(Manufacturer) - 1) 'Left() is required to trim the special whitespace character at the end
    PV_DatabaseSht.Cells(modRow, 3).Value = Left(Model, Len(Model) - 1)
    PV_DatabaseSht.Cells(modRow, 5).Value = Source
    PV_DatabaseSht.Cells(modRow, 1).Value = version
    PV_DatabaseSht.Cells(modRow, 6).Value = inYear
    PV_DatabaseSht.Cells(modRow, 12).Value = technology
    PV_DatabaseSht.Cells(modRow, 13).Value = CellsinS
    PV_DatabaseSht.Cells(modRow, 14).Value = CellsinP
    PV_DatabaseSht.Cells(modRow, 36).Value = byPassDiodes
    PV_DatabaseSht.Cells(modRow, 8).Value = PNom
    PV_DatabaseSht.Cells(modRow, 15).Value = Gref
    PV_DatabaseSht.Cells(modRow, 10).Value = Toleranz
    PV_DatabaseSht.Cells(modRow, 43).Value = AreaM
    PV_DatabaseSht.Cells(modRow, 44).Value = CellArea
    PV_DatabaseSht.Cells(modRow, 35).Value = vMax
    PV_DatabaseSht.Cells(modRow, 16).Value = Tref
    PV_DatabaseSht.Cells(modRow, 20).Value = Isc
    PV_DatabaseSht.Cells(modRow, 21).Value = mIsc
    PV_DatabaseSht.Cells(modRow, 19).Value = Voc
    PV_DatabaseSht.Cells(modRow, 22).Value = mVco
    PV_DatabaseSht.Cells(modRow, 18).Value = Impp
    PV_DatabaseSht.Cells(modRow, 17).Value = Vmpp
    PV_DatabaseSht.Cells(modRow, 37).Value = BypassDiodeVoltage
    PV_DatabaseSht.Cells(modRow, 26).Value = Rshunt
    PV_DatabaseSht.Cells(modRow, 27).Value = Rseries
    PV_DatabaseSht.Cells(modRow, 24).Value = Rsh0
    PV_DatabaseSht.Cells(modRow, 25).Value = Rshexp
    PV_DatabaseSht.Cells(modRow, 23).Value = mPmpp
    PV_DatabaseSht.Cells(modRow, 1).EntireRow.RowHeight = 15
    
    If checkValidVersion(version) = False And dotMarkerCounter = 2 Then
        For i = 0 To 8
            PV_DatabaseSht.Cells(modRow, 47 + 2 * i).Value = aoi(i)
            PV_DatabaseSht.Cells(modRow, 47 + 2 * i + 1).Value = Modifier(i)
        Next i
    End If
    
 End If
    
    Call PostModify(PV_DatabaseSht, currentShtStatus)
    
    Exit Function

    ' Return false for an unsuccessful parse
invalidPANFile:
    On Error GoTo 0
    ParseBinaryPANFile = 0
    
End Function

' ParseONDFile function
'
'
' This module is responsible for determining the type of OND file and parsing it respectively
Function ParseONDFile(ByVal ONDFilePath As String, ByRef dupInverterRepeat As Integer) As Integer
    
    Dim firstLine As String
    
    On Error GoTo invalidONDFile
    
    ' Open and read first line of OND File
    Open ONDFilePath For Input As #1
    Line Input #1, firstLine
    Close #1
    
    If (StrComp(firstLine, "PVObject_=pvGInverter") = 0) Then
        ParseONDFile = ParseTextONDFile(ONDFilePath, dupInverterRepeat)
    Else
        ParseONDFile = ParseBinaryONDFile(ONDFilePath, dupInverterRepeat)
    End If
    
    Exit Function
    
invalidONDFile:
    On Error GoTo 0
    ParseONDFile = 0
    
End Function

' ParseTextONDFile function
'
'
' This module is responsible for parsing text based OND files and populating the inverter database sheet
' This particular format of OND file is used by PvSyst in versions greater than or equal to 6.4 and follow a format similar to YAML
Function ParseTextONDFile(ByVal ONDFilePath As String, ByRef dupInverterRepeat As Integer) As Integer
    Dim textLine As String
    Dim lineArray() As String
    Dim currentDict As String
    Dim modRow As Integer                           ' row where the module data is going to be copied
    Dim EffPoint() As String
    Dim SingleEffCurve As Boolean
    Dim currentShtStatus As sheetStatus ' Sheet status used for Pre/Post modifying
    
    'Define Dictionaries (require Reference: Microsoft Scripting Runtime)
    Dim PvGInverter As New Scripting.Dictionary
    Dim PvCommercial As New Scripting.Dictionary
    Dim Converter As New Scripting.Dictionary
    Dim ProfilPIO As New Scripting.Dictionary                    ' Efficiency curve for single efficiency curve systems
    Dim ProfilPIOV1 As New Scripting.Dictionary                  ' Efficiency curve 1 for multi-efficiency curve systems
    Dim ProfilPIOV2 As New Scripting.Dictionary                  ' Efficiency curve 2 for multi-efficiency curve systems
    Dim ProfilPIOV3 As New Scripting.Dictionary                  ' Efficiency curve 3 for multi-efficiency curve systems
    
      
    ' Initialize as false and turn to true if configuration finishes
    ParseTextONDFile = False
    On Error GoTo invalidONDFile
    
    Open ONDFilePath For Input As #1
    
    currentDict = "PvModule"
    SingleEffCurve = False
    Do Until EOF(1)
        Line Input #1, textLine
        If (InStr(textLine, "pvGInverter") <> 0) Then
            currentDict = "pvGInverter"
        ElseIf (InStr(textLine, "pvCommercial") <> 0) Then
            currentDict = "pvCommercial"
        ElseIf (InStr(textLine, "Converter") <> 0) Then
            currentDict = "Converter"
        ElseIf (InStr(textLine, "ProfilPIOV1") <> 0) Then
            currentDict = "ProfilPIOV1"
        ElseIf (InStr(textLine, "ProfilPIOV2") <> 0) Then
            currentDict = "ProfilPIOV2"
        ElseIf (InStr(textLine, "ProfilPIOV3") <> 0) Then
            currentDict = "ProfilPIOV3"
        ElseIf (InStr(textLine, "ProfilPIO") <> 0) Then
            currentDict = "ProfilPIO"
            SingleEffCurve = True
        End If
    
        ' Populating dictionaries
        If (InStr(textLine, "=") <> 0) Then
            lineArray = Split(textLine, "=")
            Select Case currentDict
                Case "pvGInverter"
                    PvGInverter.Add Trim(lineArray(0)), Trim(lineArray(1))
                Case "pvCommercial"
                    PvCommercial.Add Trim(lineArray(0)), Trim(lineArray(1))
                Case "Converter"
                    Converter.Add Trim(lineArray(0)), Trim(lineArray(1))
                Case "ProfilPIO"
                    ProfilPIO.Add Trim(lineArray(0)), Trim(lineArray(1))
                Case "ProfilPIOV1"
                    ProfilPIOV1.Add Trim(lineArray(0)), Trim(lineArray(1))
                Case "ProfilPIOV2"
                    ProfilPIOV2.Add Trim(lineArray(0)), Trim(lineArray(1))
                Case "ProfilPIOV3"
                    ProfilPIOV3.Add Trim(lineArray(0)), Trim(lineArray(1))
            End Select
        End If
    Loop
    
    Close #1
    
    
    
    
    
    
    ' Check if the inverter already exists, based on Manufacturer, Model, and Origin
    Dim getIndex As Integer ' the Index of the inverter
    getIndex = InvIndex(PvCommercial("Manufacturer"), PvCommercial("Model"), "User_Added " & PvGInverter("Version"))
    
    ' Initialize row where to write module to zero
    modRow = 0
    
    ' If module already exists, ask whether to overwrite or skip
    Dim overWrite As Integer            ' 1 if existing module should be replaced, 2 if skipped
    
    If getIndex <> 0 Then
        If dupInverterRepeat = 0 Then
            ' Show the form
            UF_AddSameInverterOptions.CheckBox1 = False
            UF_AddSameInverterOptions.Label2.Caption = "User_Added " & PvGInverter("Version") & " | " & PvCommercial("Manufacturer") & " | " & PvCommercial("Model")
            UF_AddSameInverterOptions.Show
            ' Retrieve user choice
            overWrite = UF_AddSameInverterOptions.getChoice()
            If UF_AddSameInverterOptions.getChoiceRepeat() Then
                dupInverterRepeat = overWrite                  ' Remember which option was selected
            End If
        Else
            overWrite = dupInverterRepeat
        End If
        
        ' If overwrite option is selected, module will be copied at same index
        If (overWrite = 1) Then
            modRow = getIndex + 2
            ParseTextONDFile = -1
        Else
            ParseTextONDFile = 0
        End If
        
    ' If module doesn't exist, it will be added to end of table
    Else
        modRow = Inverter_DatabaseSht.Range("A" & Rows.count).End(xlUp).row + 1
        ParseTextONDFile = 1
    End If
    
    ' Add module
    If (modRow <> 0) Then
            
      
        ' Using dictionary to populate inverter database sheet
        Call PreModify(Inverter_DatabaseSht, currentShtStatus)
                        
        Inverter_DatabaseSht.Cells(modRow, 1).Value = "User_Added " & PvGInverter("Version")
        Inverter_DatabaseSht.Cells(modRow, 2).Value = PvCommercial("Manufacturer")
        Inverter_DatabaseSht.Cells(modRow, 3).Value = PvCommercial("Model")
        Inverter_DatabaseSht.Cells(modRow, 4).Value = Right(ONDFilePath, Len(ONDFilePath) - InStrRev(ONDFilePath, "\"))
        Inverter_DatabaseSht.Cells(modRow, 5).Value = PvCommercial("DataSource")
        Inverter_DatabaseSht.Cells(modRow, 6).Value = PvCommercial("YearBeg")
        Inverter_DatabaseSht.Cells(modRow, 7).Value = PvCommercial("YearEnd")
        Inverter_DatabaseSht.Cells(modRow, 8).Value = Converter("PNomConv")
        Inverter_DatabaseSht.Cells(modRow, 9).Value = Converter("PMaxOUT")
        Inverter_DatabaseSht.Cells(modRow, 10).Value = Converter("INomAC")
        Inverter_DatabaseSht.Cells(modRow, 11).Value = Converter("IMaxAC")
        Inverter_DatabaseSht.Cells(modRow, 12).Value = Converter("VOutConv")
        Inverter_DatabaseSht.Cells(modRow, 13).Value = Converter("MonoTri")
        Inverter_DatabaseSht.Cells(modRow, 15).Value = Converter("PNomDC")
        Inverter_DatabaseSht.Cells(modRow, 16).Value = Converter("PMaxDC")
        Inverter_DatabaseSht.Cells(modRow, 18).Value = Converter("PSeuil")
        Inverter_DatabaseSht.Cells(modRow, 19).Value = Converter("VmppNom")
        Inverter_DatabaseSht.Cells(modRow, 20).Value = Converter("VMppMin")
        Inverter_DatabaseSht.Cells(modRow, 21).Value = Converter("VMPPMax")
        Inverter_DatabaseSht.Cells(modRow, 22).Value = Converter("VAbsMax")
        Inverter_DatabaseSht.Cells(modRow, 24).Value = Converter("INomDC")
        Inverter_DatabaseSht.Cells(modRow, 25).Value = Converter("IMaxDC")
        Inverter_DatabaseSht.Cells(modRow, 27).Value = Converter("ModeOper")
        Inverter_DatabaseSht.Cells(modRow, 28).Value = Converter("CompPMax")
        Inverter_DatabaseSht.Cells(modRow, 29).Value = Converter("CompVMax")
        Inverter_DatabaseSht.Cells(modRow, 33).Value = Converter("MasterSlave")
        Inverter_DatabaseSht.Cells(modRow, 34).Value = Converter("NbInputs")
        Inverter_DatabaseSht.Cells(modRow, 42).Value = Converter("EfficMax")
        Inverter_DatabaseSht.Cells(modRow, 43).Value = Converter("EfficEuro")
        Inverter_DatabaseSht.Cells(modRow, 44).Value = Converter("FResNorm")
        
        If (SingleEffCurve) Then
            ' Single Efficiency curve data
            Call AssignEffPoints(ProfilPIO, 2, 8, 45, modRow)
            
        Else
            ' 3 efficiency curves definition
            Inverter_DatabaseSht.Cells(modRow, 40).Value = "X"
            ' Assigning Input voltages
            EffPoint = Split(Converter("VNomEff"), ",")
            Inverter_DatabaseSht.Cells(modRow, 76).Value = EffPoint(0)  ' Low
            Inverter_DatabaseSht.Cells(modRow, 93).Value = EffPoint(1)  ' Med
            Inverter_DatabaseSht.Cells(modRow, 110).Value = EffPoint(2) ' High
            ' Assigning Max Efficiency Voltage
            EffPoint = Split(Converter("EfficMaxV"), ",")
            Inverter_DatabaseSht.Cells(modRow, 77).Value = EffPoint(0)  ' Low
            Inverter_DatabaseSht.Cells(modRow, 94).Value = EffPoint(1)  ' Med
            Inverter_DatabaseSht.Cells(modRow, 111).Value = EffPoint(2) ' High
            ' Assigning Max Euro Efficiency Voltage
            EffPoint = Split(Converter("EfficEuroV"), ",")
            Inverter_DatabaseSht.Cells(modRow, 78).Value = EffPoint(0)  ' Low
            Inverter_DatabaseSht.Cells(modRow, 95).Value = EffPoint(1)  ' Med
            Inverter_DatabaseSht.Cells(modRow, 112).Value = EffPoint(2) ' High
            ' Assigning points for 3 efficiency curves
            Call AssignEffPoints(ProfilPIOV1, 2, 8, 79, modRow)
            Call AssignEffPoints(ProfilPIOV2, 2, 8, 96, modRow)
            Call AssignEffPoints(ProfilPIOV3, 2, 8, 113, modRow)
        End If
        
        ' The below set of values are not important to the simulation so skip them if they cannot be read
        On Error Resume Next
        
        Inverter_DatabaseSht.Cells(modRow, 59).Value = PvCommercial("Height")
        Inverter_DatabaseSht.Cells(modRow, 60).Value = PvCommercial("Width")
        Inverter_DatabaseSht.Cells(modRow, 61).Value = PvCommercial("Depth")
        Inverter_DatabaseSht.Cells(modRow, 62).Value = PvCommercial("Weight")
        Inverter_DatabaseSht.Cells(modRow, 63).Value = PvCommercial("Currency")
        Inverter_DatabaseSht.Cells(modRow, 68).Value = PvCommercial("NPieces")
        Inverter_DatabaseSht.Cells(modRow, 70).Value = Right(PvCommercial("Str_1"), Len(PvCommercial("Str_1")) - 11)
        
        ' Resume proper error processing
        On Error GoTo invalidONDFile
        
        ' Determining if inverter has bipolar inputs
        If (StrComp(PvCommercial("Str_4"), "Bipolar inputs") = 0) Then
            Inverter_DatabaseSht.Cells(modRow, 72).Value = "TRUE"
        Else
            Inverter_DatabaseSht.Cells(modRow, 72).Value = "FALSE"
        End If
        
        Inverter_DatabaseSht.Cells(modRow, 73).Value = PvCommercial("Comment")
        
    End If
        
        Call PostModify(Inverter_DatabaseSht, currentShtStatus)
        
        Exit Function
    
' The code only reaches here if an error due to formatting was encountered
invalidONDFile:
    On Error GoTo 0
    ParseTextONDFile = 0

End Function

' AssignEffPoints
'
'
' This module is responsible for splitting efficiency points into their values and assigning them to specified cells in the OND database
Private Sub AssignEffPoints(ByRef dict As Dictionary, ByVal startIndex As Integer, ByVal endIndex As Integer, ByVal firstColumnIndex As Integer, ByVal row As Integer)
    ' Dict                              ' Dictionary with points to be split
    ' startIndex                        ' First point to be written to database
    ' endIndex                          ' Last point to be written to database
    ' firstColumnIndex                  ' Column to begin writting values
    ' row                               ' Row the values will be written to
    
    Dim counter As Integer              ' Counter to be used in for loop
    Dim EffPoint() As String            ' Array used to hold values upon splitting point
    Dim keyName As String               ' String used to told key in dictionary i.e. "Point_1"
    Dim Index As Integer                ' Column the value will be placed in
    Index = firstColumnIndex
    
    For counter = startIndex To endIndex
        keyName = "Point_" & CStr(counter)
        EffPoint = Split(dict(keyName), ",")
        Inverter_DatabaseSht.Cells(row, Index).Value = EffPoint(0) / 1000
        Index = Index + 1
        Inverter_DatabaseSht.Cells(row, Index).Value = CDbl(EffPoint(1)) / CDbl(EffPoint(0)) * 100
        Index = Index + 1
    Next counter
End Sub

' ParseBinaryONDFile function
'
'
' This module is responsible for parsing binary OND files and populating the inverter database sheet
' This particular format of OND file is used by PvSyst in versions less than 6.4
Function ParseBinaryONDFile(ByVal ONDFilePath As String, ByRef dupInverterRepeat As Integer) As Integer
   
   
    Dim byteArray() As Byte ' The byte array used to store all the bytes in the file; each Index represents one byte
    Dim modRow As Integer                           ' row where the module data is going to be copied
    Dim currentShtStatus As sheetStatus
    
    ' Counter and Index variables for the first part of the file
    Dim i As Integer ' Loop Index control variable
    Dim aCell As Range ' for each loop control variable
    Dim aByte As Variant ' for each loop control variable
    Dim phaseByte ' The single byte which changes based on the phase type (Mono, Bi, Tri)
    Dim real48StartIndex As Integer ' The Index of the byte marking the beginning of the real48 section
    Dim manuStartIndex As Integer
    Dim invStartIndex As Integer
    Dim sourceStartIndex As Integer
    Dim versionStartIndex As Integer
    Dim versionEndIndex As Integer
    Dim yearStartIndex As Integer
    Dim effCurveStartIndex As Integer
    Dim effAttributeStartIndex As Integer
    Dim isMultiCurved As Boolean ' Boolean check to check if an inverter is multi-curved or single-curved
    
    ' Efficiency curve arrays; multi-curved arrays have all 4 curves stored in the OND file
    ' single-curved arrays only have the standard efficiency curve
    Dim standardEffCurve(0 To 17) As String ' This curve has an extra spot for an 'X' to mark a single-curve inverter
    Dim lowEffCurve(0 To 16) As String
    Dim medEffCurve(0 To 16) As String
    Dim highEffCurve(0 To 16) As String
     
    ' Parsed inverter parameters from the file
    Dim Manufacturer As String
    Dim Model As String
    Dim Source As String
    Dim version As String
    Dim inYear As String
    Dim Operation As String
    Dim phaseType As String
    Dim PNomAC As String
    Dim Output As String
    Dim MinMPP As String
    Dim MaxMPP As String
    Dim MaxVoltage As String
    Dim Threshold As String
    Dim PMaxAC As String
    Dim StandardEffMax As String
    Dim StandardEuro As String
    Dim NomCurrentDC As String
    Dim NomVoltageDC As String
    Dim MinV As String
    Dim PNomDC As String
    Dim PMaxDC As String
    Dim MaxCurrent As String
    Dim InomAC As String
    Dim IMaxAC As String
    Dim Res As String
    Dim remarkNotes As String
    Dim bipolarInputs As String
    
   
    On Error GoTo invalidONDFile ' catches an Index Out of Range exception due to a file being empty or invalid
    
    byteArray = BinaryInputReader(ONDFilePath) ' Reads all bytes from the file into a byte array

    ' Finds the start indices of each parameter in the non-encoded part of the file
    manuStartIndex = FindMarkerIndex(semiColonMarker, 0, byteArray)
    invStartIndex = FindMarkerIndex(dotMarker, 0, byteArray)
    sourceStartIndex = FindMarkerIndex(dotMarker, invStartIndex, byteArray)
    versionStartIndex = FindMarkerIndex(doubleDotMarker, sourceStartIndex, byteArray)
    versionEndIndex = FindMarkerIndex(semiColonMarker, versionStartIndex, byteArray)
    yearStartIndex = FindMarkerIndex(semiColonMarker, versionEndIndex + 1, byteArray)
    
    ' The real48 section always begins 6 bytes after the forward slash
    real48StartIndex = FindMarkerIndex(forwardSlashMarker, yearStartIndex, byteArray) + 6
    
    ' Get the operation type (MPPT is denoted by a "01" hexadecimal byte 4 positions before the start of the real48 section)
    If byteArray(real48StartIndex - 4) = MPPTTrue Then
        Operation = "MPPT"
    Else
        Operation = "Fixed Voltage" ' If the byte is "02" then it operates on fixed voltage instead
    End If
    
    ' Get the phase type (mono, bi, tri) based on the value of the phase byte
    phaseByte = byteArray(real48StartIndex - 1) ' this byte is always found one byte before the real48 section
    Select Case phaseByte
        Case "1"
            phaseType = "Mono"
        Case "2"
            phaseType = "Tri"
        Case "3"
            phaseType = "Bi"
    End Select
    
    If phaseType = vbNullString Then GoTo invalidONDFile ' Additional error checking
    
     ' Parse the first part of the PAN file (not encrypted, bytes can be directly converted from Byte --> String)
    Manufacturer = ByteArrayToString(ExtractByteParameter(byteArray, manuStartIndex, invStartIndex - manuStartIndex))
    Model = ByteArrayToString(ExtractByteParameter(byteArray, invStartIndex, sourceStartIndex - invStartIndex))
    Source = ByteArrayToString(ExtractByteParameter(byteArray, sourceStartIndex, versionStartIndex - sourceStartIndex - 3))
    version = Replace(ByteArrayToString(ExtractByteParameter(byteArray, versionStartIndex, versionEndIndex - versionStartIndex - 1)), "Version", "User_Added")
    inYear = ByteArrayToString(ExtractByteParameter(byteArray, yearStartIndex, 4))
    
    
    
    ' Parse the first Real48 Section (general inverter specifications)
    PNomAC = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 0), 6))
    Output = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 1), 6))
    MinMPP = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 2), 6))
    MaxMPP = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 3), 6))
    MaxVoltage = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 4), 6))
    Threshold = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 5), 6))
    PMaxAC = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 6), 6))
    StandardEffMax = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 7), 6))
    StandardEuro = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 8), 6))
    NomCurrentDC = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 10), 6))
    NomVoltageDC = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 11), 6))
    MinV = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 12), 6))
    PNomDC = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 13), 6))
    PMaxDC = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 14), 6))
    MaxCurrent = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 15), 6))
    InomAC = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 16), 6))
    IMaxAC = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 17), 6))
    Res = Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(real48StartIndex, 18), 6))
    
    
    ' Extract 7 points from the efficiency curve; each point consists of a pair of  P In(DC) and an Efficiency value at that power level
    standardEffCurve(0) = "X"
    standardEffCurve(1) = Format(StandardEffMax, valueFormat)
    standardEffCurve(2) = Format(StandardEuro, valueFormat)
    standardEffCurve(3) = Res
  
    
    If checkValidVersion(version) = False Then
        effCurveStartIndex = FindMarkerIndex(forwardSlashMarker, real48StartIndex + 108, byteArray) + 44
        Call ExtractEfficiencyCurves(effCurveStartIndex, byteArray, standardEffCurve)
        
        isMultiCurved = (UBound(byteArray) > 1024)

        If isMultiCurved Then ' If there is more than one efficiency curve there will be more than one verticalBarMarker; effcurvestartIndex will not be 0.

            ' Parse the inverter efficiency curve general attribtes (low, med, high v, max and euro efficiencies)
            effAttributeStartIndex = FindMarkerIndex(forwardSlashMarker, effCurveStartIndex + 240, byteArray) + 2
            standardEffCurve(0) = vbNullString
            lowEffCurve(0) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 0), 6)), "[=0]0;.#") ' Get Low Voltage
            medEffCurve(0) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 1), 6)), "[=0]0;.#") ' Get Med Voltage
            highEffCurve(0) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 2), 6)), "[=0]0;.#") ' Get High Voltage
            lowEffCurve(1) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 3), 6)), valueFormat) ' Get Low max eff
            medEffCurve(1) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 4), 6)), valueFormat) ' Get Med max eff
            highEffCurve(1) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 5), 6)), valueFormat) ' Get High max eff
            lowEffCurve(2) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 6), 6)), valueFormat) ' Get Low Euro Eff
            medEffCurve(2) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 7), 6)), valueFormat) ' Get Med Euro Eff
            highEffCurve(2) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 8), 6)), valueFormat) ' Get High Euro Eff


            effCurveStartIndex = FindMarkerIndex(forwardSlashMarker, effAttributeStartIndex, byteArray) + 44 ' Find the start of the efficiency curves and offset to skip the first curve point, which is not stored.
            Call ExtractEfficiencyCurves(effCurveStartIndex, byteArray, lowEffCurve)

            effCurveStartIndex = FindMarkerIndex(forwardSlashMarker, effCurveStartIndex + 240, byteArray) + 44
            Call ExtractEfficiencyCurves(effCurveStartIndex, byteArray, medEffCurve)

            effCurveStartIndex = FindMarkerIndex(forwardSlashMarker, effCurveStartIndex + 240, byteArray) + 44
            Call ExtractEfficiencyCurves(effCurveStartIndex, byteArray, highEffCurve)

        End If
    Else
        effCurveStartIndex = FindMarkerIndex(verticalBarMarker, 0, byteArray) + 34
        Call ExtractEfficiencyCurves(effCurveStartIndex, byteArray, standardEffCurve)
    
        isMultiCurved = (FindMarkerIndex(verticalBarMarker, effCurveStartIndex + 1, byteArray) > 0)
        
        If isMultiCurved Then ' If there is more than one efficiency curve there will be more than one verticalBarMarker; effcurvestartIndex will not be 0.

            ' Parse the inverter efficiency curve general attribtes (low, med, high v, max and euro efficiencies)
            effAttributeStartIndex = FindMarkerIndex(forwardSlashMarker, effCurveStartIndex + 240, byteArray) + 2
            standardEffCurve(0) = vbNullString
            lowEffCurve(0) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 0), 6)), "[=0]0;.#") ' Get Low Voltage
            medEffCurve(0) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 1), 6)), "[=0]0;.#") ' Get Med Voltage
            highEffCurve(0) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 2), 6)), "[=0]0;.#") ' Get High Voltage
            lowEffCurve(1) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 3), 6)), valueFormat) ' Get Low max eff
            medEffCurve(1) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 4), 6)), valueFormat) ' Get Med max eff
            highEffCurve(1) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 5), 6)), valueFormat) ' Get High max eff
            lowEffCurve(2) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 6), 6)), valueFormat) ' Get Low Euro Eff
            medEffCurve(2) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 7), 6)), valueFormat) ' Get Med Euro Eff
            highEffCurve(2) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effAttributeStartIndex, 8), 6)), valueFormat) ' Get High Euro Eff
 

            effCurveStartIndex = FindMarkerIndex(verticalBarMarker, effCurveStartIndex, byteArray) + 34 ' Find the start of the efficiency curves and offset to skip the first curve point, which is not stored.
            Call ExtractEfficiencyCurves(effCurveStartIndex, byteArray, lowEffCurve)

            effCurveStartIndex = FindMarkerIndex(verticalBarMarker, effCurveStartIndex, byteArray) + 34
            Call ExtractEfficiencyCurves(effCurveStartIndex, byteArray, medEffCurve)

            effCurveStartIndex = FindMarkerIndex(verticalBarMarker, effCurveStartIndex, byteArray) + 34
            Call ExtractEfficiencyCurves(effCurveStartIndex, byteArray, highEffCurve)

       End If
    End If
    
    For i = effCurveStartIndex To UBound(byteArray)
        remarkNotes = remarkNotes & " " & Hex$(byteArray(i))
    Next i
    
    If InStr(1, remarkNotes, "42 69 70 6F 6C 61 72") Then
        bipolarInputs = "TRUE"
    Else
        bipolarInputs = "FALSE"
    End If
    
    
    
    
    ' Check if the inverter already exists, based on Manufacturer, Model, and Origin
    Dim getIndex As Integer ' the Index of the inverter
    getIndex = InvIndex(Left(Manufacturer, Len(Manufacturer) - 1), Left(Model, Len(Model) - 1), version)
    
    ' Initialize row where to write module to zero
    modRow = 0
    
    ' If module already exists, ask whether to overwrite or skip
    Dim overWrite As Integer            ' 1 if existing module should be replaced, 2 if skipped
    
    If getIndex <> 0 Then
        If dupInverterRepeat = 0 Then
            ' Show the form
            UF_AddSameInverterOptions.CheckBox1 = False
            UF_AddSameInverterOptions.Label2.Caption = version & " | " & Left(Manufacturer, Len(Manufacturer) - 1) & " | " & Left(Model, Len(Model) - 1)
            UF_AddSameInverterOptions.Show
            ' Retrieve user choice
            overWrite = UF_AddSameInverterOptions.getChoice()
            If UF_AddSameInverterOptions.getChoiceRepeat() Then
                dupInverterRepeat = overWrite                  ' Remember which option was selected
            End If
        Else
            overWrite = dupInverterRepeat
        End If
        
        ' If overwrite option is selected, module will be copied at same index
        If (overWrite = 1) Then
            modRow = getIndex + 2
            ParseBinaryONDFile = -1
        Else
            ParseBinaryONDFile = 0
        End If
        
    ' If module doesn't exist, it will be added to end of table
    Else
        modRow = Inverter_DatabaseSht.Range("A" & Rows.count).End(xlUp).row + 1
        ParseBinaryONDFile = 1
    End If
    
    ' Add module
    If (modRow <> 0) Then
    
     
    
        ' Write to Inverter Database
        On Error GoTo 0
        Call PreModify(Inverter_DatabaseSht, currentShtStatus)
         
        ' Extract the PAN file name from the full path
        Inverter_DatabaseSht.Cells(modRow, 4).Value = Right(ONDFilePath, Len(ONDFilePath) - InStrRev(ONDFilePath, "\"))
        Inverter_DatabaseSht.Cells(modRow, 2).Value = Left(Manufacturer, Len(Manufacturer) - 1) 'Left() is required to trim the special whitespace character at the end
        Inverter_DatabaseSht.Cells(modRow, 3).Value = Left(Model, Len(Model) - 1)
        Inverter_DatabaseSht.Cells(modRow, 5).Value = Source
        Inverter_DatabaseSht.Cells(modRow, 1).Value = version
        Inverter_DatabaseSht.Cells(modRow, 6).Value = inYear
        Inverter_DatabaseSht.Cells(modRow, 27).Value = Operation
        Inverter_DatabaseSht.Cells(modRow, 13).Value = phaseType
        Inverter_DatabaseSht.Cells(modRow, 8).Value = PNomAC
        Inverter_DatabaseSht.Cells(modRow, 22).Value = MaxVoltage
        Inverter_DatabaseSht.Cells(modRow, 23).Value = MinV
        Inverter_DatabaseSht.Cells(modRow, 12).Value = Output
        Inverter_DatabaseSht.Cells(modRow, 20).Value = MinMPP
        Inverter_DatabaseSht.Cells(modRow, 21).Value = MaxMPP
        Inverter_DatabaseSht.Cells(modRow, 18).Value = Threshold
        Inverter_DatabaseSht.Cells(modRow, 9).Value = PMaxAC
        Inverter_DatabaseSht.Cells(modRow, 24).Value = NomCurrentDC
        Inverter_DatabaseSht.Cells(modRow, 19).Value = NomVoltageDC
        Inverter_DatabaseSht.Cells(modRow, 15).Value = PNomDC
        Inverter_DatabaseSht.Cells(modRow, 16).Value = PMaxDC
        Inverter_DatabaseSht.Cells(modRow, 25).Value = MaxCurrent
        Inverter_DatabaseSht.Cells(modRow, 10).Value = InomAC
        Inverter_DatabaseSht.Cells(modRow, 11).Value = IMaxAC
        Inverter_DatabaseSht.Cells(modRow, 44).Value = Res
        Inverter_DatabaseSht.Cells(modRow, 72).Value = bipolarInputs
        Inverter_DatabaseSht.Cells(modRow, 1).EntireRow.RowHeight = 15
        
        
        Call WriteEffCurveToDatabase(modRow, databaseStandardEffStart, standardEffCurve)
        
        ' If the inverter is multi-curved then write the low, med and high curves to the database as well
        If isMultiCurved Then
            Inverter_DatabaseSht.Cells(modRow, 40).Value = "X"
            Call WriteEffCurveToDatabase(modRow, databaseLowEffStart, lowEffCurve)
            Call WriteEffCurveToDatabase(modRow, databaseMedEffStart, medEffCurve)
            Call WriteEffCurveToDatabase(modRow, databaseHighEffStart, highEffCurve)
        End If
        
        ' Clear all invalid values (-9999)
        For Each aCell In Inverter_DatabaseSht.Range("A" & modRow, "Y" & modRow)
            If aCell.Value = -9999 Then aCell.Value = vbNullString
        Next
    End If
 
    Call PostModify(Inverter_DatabaseSht, currentShtStatus)
    

    Exit Function
    
' The code only reaches here if an error due to formatting was encountered
invalidONDFile:
    On Error GoTo 0
    ParseBinaryONDFile = 0
      
End Function

' Reads the input file using binary access (PAN or OND files)
' and returns an array of bytes.
' Each Index in the byte array stores a single hexadecimal byte
Private Function BinaryInputReader(ByVal PANFilePath As String) As Byte()
    
    Dim aByte As Byte
    Dim i As Integer
    Dim fileNumber As Integer
    Dim byteArray() As Byte
    i = 0
    
    fileNumber = FreeFile ' Specifies the next free file number
    
    Open PANFilePath For Binary Access Read As fileNumber
    ' Set array size (0 to the number of bytes signified by the length of file (LOF))
    ReDim byteArray(0 To LOF(fileNumber))
    
    Do While Not EOF(fileNumber)
        Get fileNumber, , aByte
        byteArray(i) = aByte
        i = i + 1
    Loop
    Close fileNumber
    
    ' Return an array of bytes
    BinaryInputReader = byteArray
    
End Function
' Code adapted from http://www.freevbcode.com/ShowCode.asp?ID=209
' Converts bytes to readable string format
Public Function ByteArrayToString(bytArray() As Byte) As String
    
    Dim sAns As String
    Dim iPos As String
    
    sAns = StrConv(bytArray, vbUnicode)
    iPos = InStr(sAns, Chr(0))
    If iPos > 0 Then sAns = Left(sAns, iPos - 1)
    
    ByteArrayToString = sAns
 
End Function
 
' This function is used to convert the Delphi Real48 encoding
' into a double after taking in an array of bytes
'
' This function was adapted from the C# code
' found at the following website:
' http://stackoverflow.com/questions/2506942/convert-delphi-real48-to-c-sharp-double
' from an answer posted by users heinrich5991 and mat-mcloughlin

 Private Function Real48ToDouble(real48() As Byte) As Double
    
    If real48(0) = 0 Then
        Real48ToDouble = 0# ' If the first bit in the 6 bit sequence is 0 then the whole number is 0
        Exit Function
    End If
    
    Dim i As Integer
    Dim exponent As Double
    Dim mantissa As Double ' The decimal part of the number
    
    
    exponent = real48(0) - 129#
    mantissa = 0#
    
    For i = 1 To 4 Step 1
        mantissa = mantissa + real48(i)
        mantissa = mantissa * 0.00390625
    Next i
    
    mantissa = mantissa + (real48(5) And &H7F)
    mantissa = mantissa * 0.0078125
    mantissa = mantissa + 1#
    
    If (real48(5) And &H80) = &H80 Then ' logical bitwise checks
        mantissa = -mantissa ' Sign checking
    End If
    
    mantissa = mantissa * Application.WorksheetFunction.Power(2#, exponent)
    
    Real48ToDouble = mantissa
          
 End Function
' FindMarkerIndex function
'
' loops through the byte array starting at the specified position and
' returns the Index of the first occurrence of a hexadecimal marker
' passed to the function
Private Function FindMarkerIndex(ByVal markerFind As String, ByVal startIndex As Integer, byteArray() As Byte)
    
    Dim i As Integer
    Dim markerIndex As Integer
     
    For i = startIndex To UBound(byteArray)
        If Hex$(byteArray(i)) = markerFind Then
            markerIndex = i + 1
            Exit For
        End If
    Next i
    
    FindMarkerIndex = markerIndex
    
End Function

' This function extracts bytes that form a single parameter from the original byte array (contains the bytes from the whole file)
' into a smaller byte array that it returns.
Private Function ExtractByteParameter(byteArray() As Byte, ByVal startIndex As Integer, ByVal numBytes As Integer) As Byte()
    Dim i As Integer
    Dim j As Integer
    Dim paramByteSequence() As Byte ' The smaller array that is returned, only contains the number of bytes specified and extracted
    
    j = 0
    ReDim paramByteSequence(0 To numBytes - 1) ' Set the array size based on how many bytes are to be extracted
    For i = startIndex To startIndex + numBytes - 1 ' Loop through the byte array starting at the specified Index for the provided number of bytes
        paramByteSequence(j) = byteArray(i)
        j = j + 1
    Next i

    ExtractByteParameter = paramByteSequence
    
End Function
' A function that returns the specific Index of the parameter to be extracted in the array
' based on the start Index of the Real48 encoded section and an offset multiplier which
' denotes the parameter's position as the number of 6 bit sections it is away from
' the first parameter in the Real48 section
Private Function getParamIndex(ByVal startIndex As Integer, ByVal offsetNum As Integer)
    getParamIndex = startIndex + 6 * offsetNum
End Function

Private Sub ExtractEfficiencyCurves(ByVal effCurveStartIndex As Integer, byteArray() As Byte, ByRef effCurveArray() As String)

    Dim j As Integer
    Dim i As Integer
      
    If UBound(effCurveArray) = 16 Then
        j = 3 ' Fill low, med, high eff curves (0 to 2 reserved for voltage, max and euro eff)
    Else
        j = 4 ' Fill standard single eff curve (0 to 3 reserved for voltage, max, euro eff and res)
    End If
    
    For i = 0 To 34 Step 5
        effCurveArray(j) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effCurveStartIndex, i), 6)) / 1000, valueFormat) ' Parse the PIn(DC) in watts, convert to kW
        
        If effCurveArray(j) <> 0 Then
            effCurveArray(j + 1) = Format(100 * Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(effCurveStartIndex, i + 1), 6)) / (effCurveArray(j) * 1000), valueFormat) ' Parse PIn(DC)*Efficiency and divide by PIn(DC) to get Efficiency
        Else
            effCurveArray(j + 1) = vbNullString
            effCurveArray(j) = vbNullString
        End If
        
        j = j + 2
    Next i
    
End Sub

Private Sub ExtractIAMProfile(ByVal iamStartIndex As Integer, byteArray() As Byte, ByRef aoi() As String, ByRef Modifier() As String)
    Dim j As Integer
    Dim i As Integer
    
    For i = 0 To 44 Step 5
        aoi(j) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(iamStartIndex, i), 6)), valueFormat)
        
        If aoi(j) <> vbNullString Then
            Modifier(j) = Format(Real48ToDouble(ExtractByteParameter(byteArray, getParamIndex(iamStartIndex, i + 1), 6)), valueFormat)
        Else
            aoi(j) = vbNullString
            Modifier(j) = vbNullString
        End If
        
        j = j + 1
    Next i

End Sub
' WriteEffCurveToDatabase subroutine
'
' The values in the efficiency curve arrays already correspond to the database format
' so a loop is the most practical way to write to the database
Private Sub WriteEffCurveToDatabase(ByVal lastRow As Integer, ByVal effCurveStartIndex, ByRef effCurveArray() As String)
    Dim databaseColumn As Integer
    Dim curveEntry As Integer
    Dim arraySize As Integer
    
    arraySize = UBound(effCurveArray)
    curveEntry = 0
    
    ' The effCurveStartIndex here corresponds to the column in the inverter database where the efficiency curve
    ' begins, and not the Index of bytes in the array
    For databaseColumn = effCurveStartIndex To effCurveStartIndex + arraySize Step 1
        Inverter_DatabaseSht.Cells(lastRow, databaseColumn).Value = effCurveArray(curveEntry)
        curveEntry = curveEntry + 1
    Next databaseColumn
    
End Sub
' This function is used to check the version number from the OND file (or PAN file if applicable)
' since versions 6 and up are currently not supported.
Private Function checkValidVersion(ByVal versionString As String)
Dim validversioncheck As Boolean
    
    checkValidVersion = True
    validversioncheck = CInt(Mid(versionString, InStr(1, versionString, ".") - 1, 1)) < 6
    If validversioncheck = False Then
        checkValidVersion = False
    End If
End Function

