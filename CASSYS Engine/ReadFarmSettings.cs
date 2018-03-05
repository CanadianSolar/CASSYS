// CASSYS - Grid connected PV system modelling software  
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: ReadFarmSettings
// 
// Revision History:
// AP - 2014-11-27: Version 0.9
// AP - 2015-05-21: Version 0.9.2 Added reading capability for 0.9.2
// NA - 2017-06-09: Version 1.1  Changes made to the structure of the file to include system mode parameter and addition of ASTM Mode
//
// Description:
// This class uses the information provided in the .CSYX file and dictates the number
// of Sub-Arrays, the Input File Scheme, the Output File Scheme and general simulation
// settings used in the program file.
// The class also collects the Input and Output file paths and assigns them to 
// variables used in the main program.    
//                        
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
// 
// 
///////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Text;
using System.Threading;
using System.Data;

namespace CASSYS
{
    public class ReadFarmSettings
    {
        // Inputs or Parameters for the ReadFarmSettings Class
        public static String EngineVersion = "1.3.1";               // The supported versions of CASSYS CSYX Files.
        public static XmlDocument doc;                              // The .CSYX document that contains the Site, System, etc. definitions
        public static String CASSYSCSYXVersion;                     // The CASSYS .CSYX Version Number obtained from the .CSYX file
        public static bool UseDiffMeasured;                         // Using the Measured Diffuse on Horizontal Value
        public static bool UseMeasuredTemp;                         // Using measured temperature for the panels
        public static bool UsePOA;                                  // Boolean that indicates if the program should use tilted irradiance to simulate
        public static bool UseGHI;                                  // Boolean that indicates if the program should use horizontal irradiance to simulate
        public static bool UseWindSpeed;                            // Boolean indicating if the program has the Wind Speed available to use
        public static bool tempAmbDefined;                          // Boolean indicating if the program has Temp Ambient available to use
        public static bool batchMode = false;                       // Determines if the program is being run in batch mode (CMD prompt arguments for IO files) or not
        public static string outputMode = "csv";                    // Determines if output should be a comma seperated string (str), a csv file (csv), or a csv file with multiple runs with varying parameter (var)
        public static string currentYear = DateTime.Now.Year.ToString();   // Gets the current year that the program is run for copywrite message

        // Gathering Input and Output Configurations for the ReadFarmSettings Class
        public static String SimInputFile;                          // Input file path
        public static String SimOutputFile;                         // Output file path
        public static int SubArrayCount;                            // Number of system sub-arrays, default: 1
        public static int ClimateFileRowsToSkip;                    // Number of rows to skip
        public static int TMYType;                                  // If a TMY file is loaded, it specifies either 2 or 3 for .tm2 or .tm3 format
        public static string Delim;                                 // The character used in the input file between adjacent values
        public static int[] ClimateRefPos;                          // The positions of input data from the user input file
        //public static bool NoSystemDefined = false;               // Checks if no system was defined and allow user to simulate irradiance values only
        public static string SystemMode = "GridConnected";          // The system type will determine which calculations are required
        public static int IncClimateRowsAllowed;                    // Number of incorrectly formatted climate file rows CASSYS can skip over until simulation stops

        // Output configuration for the Program
        public static Dictionary<String, dynamic> Outputlist;       // Creating a dictionary to hold all output values;
        public static List<String> OutputScheme;                    // String Array that holds values based on user request
        public static String OutputHeader = null;                   // The header of the output file, Always begins with TimeStamp.
        public static DataTable outputTable;                        // For variable parameter mode the output data held in a datatable
        public static int runNumber;                                // Current simulation being ran for variable parameter mode

        // Finding and assigning the Simulation Input and Output file name 
        public static void AssignIOFileNames()
        {
            // Getting the Input file Path
            SimInputFile = GetInnerText("Site", "InputFilePath", _Error: ErrLevel.FATAL);
            // Getting the Output file Path
            SimOutputFile = GetInnerText("Site", "OutputFilePath", _Error: ErrLevel.FATAL);
        }

        // Finding the mode of the simulation
        public static void GetSimulationMode()
        {
            // If the ModeSelect node is available, the mode of the simulation is determined from the node value. 
            // If the node is unavailable, it is determined to be a grid connected system by default.
            if (GetInnerText("Site", "ModeSelect", _Error: ErrLevel.WARNING, _default: "GridConnected", _VersionNum: "0.9.3") == "Radiation Mode")
            {
                SystemMode = "Radiation";
            }
            else if (GetInnerText("Site", "ModeSelect", _Error: ErrLevel.WARNING, _default: "GridConnected", _VersionNum: "1.0.1") == "ASTM E2848 Regression")
            {
                SystemMode = "ASTME2848Regression";
            }

            // In older versions the ModeSelect Node is not available, and the SystemDC node is checked.
            if (GetInnerText("Site", "ModeSelect", _Error: ErrLevel.WARNING, _default: "N/A", _VersionNum: "0.9.3") == "N/A")
            {
                if (double.Parse(GetInnerText("System", "SystemDC", ErrLevel.WARNING, _VersionNum: "0.9.2", _default: "1")) == 0)
                {
                    SystemMode = "Radiation";
                }
            }


        }

        // Finding and assigning the SubArray Count from the attribute of the SubArray Tag
        public static void AssignSubArrayCount()
        {
            if (string.Compare(SystemMode, "GridConnected") == 0)
            {
                // Getting the Number of Sub-Arrays for this file
                SubArrayCount = int.Parse(GetAttribute("System", "TotalArrays", _Error: ErrLevel.FATAL));
            }
        }

        // Finding and Assigning the Input file style as per the SimSettingsFile
        public static void AssignInputFileSchema()
        {
            try
            {
                // Collecting file specific information.
                Delim = GetInnerText("InputFile", "Delimeter", _Error: ErrLevel.FATAL);
                Util.AveragedAt = GetInnerText("InputFile", "AveragedAt", _Error: ErrLevel.FATAL);
                Util.timeFormat = GetInnerText("InputFile", "TimeFormat", _Error: ErrLevel.FATAL);
                Util.timeStep = double.Parse(GetInnerText("InputFile", "Interval", _Error: ErrLevel.FATAL));
                ClimateFileRowsToSkip = int.Parse(GetInnerText("InputFile", "RowsToSkip", _Error: ErrLevel.WARNING, _default: "0"));
                IncClimateRowsAllowed = int.Parse(GetInnerText("InputFile", "IncorrectClimateRowsAllowed", _Error: ErrLevel.INTERNAL, _default: "0"));
                TMYType = int.Parse(GetInnerText("InputFile", "TMYType", _default: "-1"));

                // Initializing the array to use as a holder for column numbers.
                ClimateRefPos = new int[30];

                // Notifying the user of the year change in the dates from the TMY3 file.
                if (TMYType == 3)
                {
                    ErrorLogger.Log("This is a TMY3 file. The year will be changed to 1990 to ensure the climate data is in chronological order.", ErrLevel.WARNING);
                }

                if (TMYType == 1)
                {
                    ErrorLogger.Log("This is a EPW file. The year will be changed to 2017 to ensure the climate data is in chronological order.", ErrLevel.WARNING);
                }

                // Collecting weather variable locations in the file
                if ((TMYType != 1) && (TMYType != 2) && (TMYType != 3))
                {
                    ClimateRefPos[0] = int.Parse(GetInnerText("InputFile", "TimeStamp", _Error: ErrLevel.FATAL));
                    UsePOA = Int32.TryParse(GetInnerText("InputFile", "GlobalRad", _Error: ErrLevel.WARNING, _default: "N/A"), out ClimateRefPos[2]);
                    UseGHI = Int32.TryParse(GetInnerText("InputFile", "HorIrradiance", _Error: ErrLevel.WARNING, _default: "N/A"), out ClimateRefPos[1]);
                    tempAmbDefined = Int32.TryParse(GetInnerText("InputFile", "TempAmbient", _Error: ErrLevel.FATAL), out ClimateRefPos[3]);
                    UseMeasuredTemp = Int32.TryParse(GetInnerText("InputFile", "TempPanel", _default: "N/A", _Error: ErrLevel.WARNING), out ClimateRefPos[4]);
                    UseWindSpeed = Int32.TryParse(GetInnerText("InputFile", "WindSpeed", _default: "N/A", _Error: ErrLevel.WARNING), out ClimateRefPos[5]);

                    // Check if Horizontal Irradiance is provided for use in simulation.
                    if (UseGHI)
                    {
                        // Check if Diffuse Measured is defined, if Global Horizontal is provided.
                        UseDiffMeasured = Int32.TryParse(GetInnerText("InputFile", "Hor_Diffuse", _Error: ErrLevel.WARNING), out ClimateRefPos[6]);
                    }

                    // Check if at least, and only one type of Irradiance is available to continue the simulation.
                    if (UsePOA == UseGHI)
                    {
                        // If both tilted, and horizontal are provided.
                        if (UsePOA)
                        {
                            ErrorLogger.Log("Column Numbers for both Global Tilted and Horizontal Irradiance have been provided. Please select one of these inputs to run the simulation.", ErrLevel.FATAL);
                        }
                        else
                        {
                            // If both are not provided.
                            ErrorLogger.Log("You have provided insufficient definitions for irradiance. Please check the Climate File Tab.", ErrLevel.FATAL);
                        }
                    }

                    // Check if at least one type of temperature is available to continue with the simulation.
                    if (tempAmbDefined == false && UseMeasuredTemp == false && string.Compare(SystemMode, "GridConnected") == 0)
                    {
                        ErrorLogger.Log("CASSYS did not find definitions for a temperature column in the Climate File. Please define a measured panel temperature or measured ambient temperature column.", ErrLevel.FATAL);
                    }
                }
            }
            catch (FormatException)
            {
                ErrorLogger.Log("The column number for Time Stamp is incorrectly defined. Please check your Input file definition.", ErrLevel.FATAL);
            }
        }

        // Finding and Assigning the Output file style as per the SimSettingsFile
        public static void AssignOutputFileSchema()
        {
            String OutputPath = "/Site/OutputFileStyle/*";
            XmlNodeList xnlist = doc.SelectNodes(OutputPath);
            Outputlist = new Dictionary<String, dynamic>();
            OutputScheme = new List<String>();

            // Loading all possible Output Values
            foreach (XmlNode outNode in xnlist)
            {
                if (Convert.ToBoolean(outNode.InnerText))
                {
                    // In variable parameter mode, the timestamps are only written for the first simulation
                    if (!((outNode.Name == "Input_Timestamp" || outNode.Name == "Timestamp_Used_for_Simulation") && ReadFarmSettings.outputMode == "var" && ReadFarmSettings.runNumber != 1))
                    {
                        // Gathering Relevant nodes
                        Outputlist.Add(outNode.Name, null);

                        // Creating a list of all items
                        OutputScheme.Add(outNode.Name);
                    }

                    // Creating the output header, using the display name attribute of each output, or if individual sub-array performance is requested
                    // providing the header for each PV-side or Inv-side of Sub-Array for Current, Power, and Voltage
                    if (outNode.Name == "ShowSubInv")
                    {
                        String SubArrayTitle = null;

                        for (int arrayCount = 0; arrayCount < SubArrayCount; arrayCount++)
                        {
                            SubArrayTitle += "SubArray " + (arrayCount + 1) + " Inv. 3-Phase Power (W),";
                        }

                        OutputHeader += SubArrayTitle;
                    }
                    else if (outNode.Name == "ShowSubInvV")
                    {
                        String SubArrayTitle = null;

                        for (int arrayCount = 0; arrayCount < SubArrayCount; arrayCount++)
                        {
                            SubArrayTitle += "SubArray " + (arrayCount + 1) + " 1-Phase Inv. Voltage (V)," + "SubArray ";
                        }

                        OutputHeader += SubArrayTitle;
                    }
                    else if (outNode.Name == "ShowSubInvC")
                    {
                        String SubArrayTitle = null;

                        for (int arrayCount = 0; arrayCount < SubArrayCount; arrayCount++)
                        {
                            SubArrayTitle += "SubArray " + (arrayCount + 1) + " 1-Phase Inv. Current (V)," + "SubArray ";
                        }

                        OutputHeader += SubArrayTitle;
                    }
                    else if (outNode.Name == "Sub_Array_Performance")
                    {
                        String SubArrayTitle = null;

                        for (int arrayCount = 0; arrayCount < SubArrayCount; arrayCount++)
                        {
                            SubArrayTitle += "SubArray " + (arrayCount + 1) + " PV. Voltage (V)," + "SubArray " + (arrayCount + 1) + " PV. Current (A)," + "SubArray " + (arrayCount + 1) + " PV. Power (kW),";
                        }

                        OutputHeader += SubArrayTitle;
                    }
                    else if (outNode.Name == "SubDCV")
                    {
                        String SubArrayTitle = null;

                        for (int arrayCount = 0; arrayCount < SubArrayCount; arrayCount++)
                        {
                            SubArrayTitle += "SubArray " + (arrayCount + 1) + " PV. Voltage (V),";
                        }

                        OutputHeader += SubArrayTitle;

                    }
                    else if (outNode.Name == "SubDCC")
                    {
                        String SubArrayTitle = null;

                        for (int arrayCount = 0; arrayCount < SubArrayCount; arrayCount++)
                        {
                            SubArrayTitle += "SubArray " + (arrayCount + 1) + " PV. Current (A),";
                        }

                        OutputHeader += SubArrayTitle;
                    }
                    // In variable parameter mode, the timestamps are only written for the first simulation
                    else if (!((outNode.Name == "Input_Timestamp" || outNode.Name == "Timestamp_Used_for_Simulation") && ReadFarmSettings.outputMode == "var" && ReadFarmSettings.runNumber != 1))
                    {
                        OutputHeader += outNode.Attributes["DisplayName"].Value + " (" + outNode.Attributes["Units"].Value + ")" + ",";
                    }
                }
            }
        }
        // Returns the value of the node, if the node exists
        public static String GetInnerText(String Path, String NodeName, ErrLevel _Error = ErrLevel.WARNING, String _VersionNum = "0.9", int _ArrayNum = 0, String _default = "0")
        {
            try
            {
                if (String.Compare(EngineVersion, _VersionNum)>=0)
                {
                    // Determine the Path of the .CSYX requested
                    switch (Path)
                    {
                        case "Site":
                            Path = "/Site/" + NodeName;
                            break;
                        case "ASTM":
                            Path = "/Site/ASTMRegress/" + NodeName;
                            break;
                        case "ASTM/Coeffs":
                            Path = "/Site/ASTMRegress/ASTMCoeffs/" + NodeName;
                            break;
                        case "ASTM/EAF":
                            Path = "/Site/ASTMRegress/EAF/" + NodeName;
                            break;
                        case "Albedo":
                            Path = "/Site/Albedo/" + NodeName;
                            break;
                        case "O&S":
                            Path = "/Site/Orientation_and_Shading/" + NodeName;
                            break;
                        case "System":
                            Path = "/Site/System/" + NodeName;
                            break;
                        case "PV":
                            Path = "/Site/System/" + "SubArray" + _ArrayNum + "/PVModule/" + NodeName;
                            break;
                        case "Inverter":
                            Path = "/Site/System/" + "SubArray" + _ArrayNum + "/Inverter/" + NodeName;
                            break;
                        case "Transformer":
                            Path = "/Site/System/Transformer/" + NodeName;
                            break;
                        case "Losses":
                            Path = "/Site/System/Losses/" + NodeName;
                            break;
                        case "InputFile":
                            Path = "/Site/InputFileStyle/" + NodeName;
                            break;
                        case "OutputFile":
                            Path = "/Site/OutputFileStyle/" + NodeName;
                            break;
                        case "Iterations":
                            Path = "/Site/Iterations/" + NodeName;
                            break;
                    }
                    // Check if the .CSYX Blank, if it is, return the default value
                    if (doc.SelectSingleNode(Path).InnerText == "")
                    {
                        if (_Error == ErrLevel.FATAL)
                        {
                            ErrorLogger.Log(NodeName + " is not defined. CASSYS requires this value to run.", ErrLevel.FATAL);
                            return "N/A";
                        }
                        else if(_Error == ErrLevel.WARNING)
                        {
                            ErrorLogger.Log("Warning: " + NodeName + " is not defined for this file. CASSYS assigned " + _default + " for this value.", ErrLevel.WARNING);
                            return _default;
                        }
                        else
                        {
                            return _default;
                        }
                    }
                    else
                    {
                        return doc.SelectSingleNode(Path).InnerText;
                    }
                }
                else
                {
                    ErrorLogger.Log(NodeName + " is not supported in this version of CASSYS. Please update your CASSYS Site file using the latest version available at https://github.com/CanadianSolar/CASSYS", ErrLevel.WARNING);
                    return null;
                }
            }
            catch (NullReferenceException)
            {
                if (_Error == ErrLevel.WARNING || _Error == ErrLevel.INTERNAL)
                {
                    return _default;
                }
                else
                {
                    ErrorLogger.Log(NodeName + " is not defined. CASSYS requires this value to run.", ErrLevel.FATAL);
                    return "N/A";
                }
            }
        }

        // Returns the attribute of the node, if the node and an attribute exist
        public static String GetAttribute(String Path, String AttributeName, ErrLevel _Error = ErrLevel.WARNING, String _VersionNum = "0.9", String _Adder = null, int _ArrayNum = 0)
        {
            try
            {
                if (String.Compare(EngineVersion, _VersionNum) >= 0)
                {
                    switch (Path)
                    {
                        case "Site":
                            Path = "/Site" + _Adder;
                            break;
                        case "Albedo":
                            Path = "/Site/Albedo" + _Adder;
                            break;
                        case "O&S":
                            Path = "/Site/Orientation_and_Shading" + _Adder;
                            break;
                        case "System":
                            Path = "/Site/System" + _Adder;
                            break;
                        case "PV":
                            Path = "/Site/System/" + "SubArray" + _ArrayNum + "/PVModule" + _Adder;
                            break;
                        case "Inverter":
                            Path = "/Site/System/" + "SubArray" + _ArrayNum + "/Inverter" + _Adder;
                            break;
                        case "Losses":
                            Path = "/Site/System/Losses" + _Adder;
                            break;
                        case "InputFile":
                            Path = "/Site/InputFileStyle" + _Adder;
                            break;
                        case "OutputFile":
                            Path = "/Site/OutputFileStyle" + _Adder;
                            break;
                        case "Iteration1":
                            Path = "/Site/Iterations/Iteration1" + _Adder;
                            break;
                    }

                    return doc.SelectSingleNode(Path).Attributes[AttributeName].Value;
                }
                else
                {
                    ErrorLogger.Log(AttributeName + " is not available in this version of CASSYS. Please update to the latest version available at https://github.com/CanadianSolar/CASSYS", ErrLevel.WARNING);
                    return null;
                }
            }
            catch (NullReferenceException)
            {
                if (_Error == ErrLevel.FATAL)
                {
                    ErrorLogger.Log(AttributeName + " in " + Path + " is not defined. CASSYS requires this value to run.", ErrLevel.FATAL);
                    return "N/A";
                }
                else
                {
                    ErrorLogger.Log(AttributeName + " in " + Path + " is not defined. CASSYS assigned 0 for this value.", ErrLevel.WARNING);
                    return "0";
                }
            }
        }

        // Returns the Version of the CSYX file to be simulated
        public static void CheckCSYXVersion()
        {
            CASSYSCSYXVersion = doc.SelectSingleNode("/Site/Version").InnerXml;

            // Check if version number is specified.
            if (CASSYSCSYXVersion == "")
            {
                ErrorLogger.Log("The file does not have a valid version number. Please check the site file.", ErrLevel.FATAL);
            }


            // CASSYS Version check, if the version does not match, the program should warn the user.
            if (String.Compare(EngineVersion, CASSYSCSYXVersion)<0)
            {
                ErrorLogger.Log("You are using an older version of the CASSYS Engine. Please update to the latest version available at https://github.com/CanadianSolar/CASSYS", ErrLevel.FATAL);
            }

            // Display the CSYX Version Number to the User
            Console.WriteLine("CASSYS Site File Version: " + CASSYSCSYXVersion);
        }

        // Header is the title and other elements shown when the simulation begins.
        public static void ShowHeader()
        {
            // Assigning a title to the console window.
            Console.Title = "CASSYS - Canadian Solar System Simulation Program for Grid-Connected PV Systems";

            // Show the following messages to the user
            Console.WriteLine("-------------------------------------------------------------------------------");
            Console.WriteLine("CASSYS - Canadian Solar System Simulation Program for Grid-Connected PV Systems");
            Console.WriteLine("Copyright 2015 - " + currentYear + " CanadianSolar, All rights reserved.");
            Console.WriteLine("CASSYS Engine Version: " + EngineVersion);
            Console.WriteLine("Full License: https://github.com/CanadianSolar/CASSYS/blob/master/LICENSE");
            Console.WriteLine("-------------------------------------------------------------------------------");
        }

        // Footer is responsible to close the console window after 3 seconds, unless a key is pressed.
        public static void ShowFooter()
        {
            // Let the user know that the window can be kept open for longer...
            Console.WriteLine("Status: Press any key in the next second to keep this window open.");

            // Counter to ensure console only waits for 3 seconds
            int counter = 0;

            // Keep the console window open as long as a key is pressed.
            while ((!Console.KeyAvailable) && (counter < 20))
            {
                Thread.Sleep(50);
                counter++;
            }

            // Once the loop ends check if the window should be kept open.
            if (Console.KeyAvailable)
            {
                Console.WriteLine("Status: Press any key to close this window.");
                Console.ReadKey();
                Console.ReadKey();
            }
        }
    }
}
