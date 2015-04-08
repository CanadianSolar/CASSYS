// CASSYS - Grid connected PV system modelling software 
// Version 0.9 
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: ReadFarmSettings
// 
// Revision History:
// AP - 2014-11-27: Version 0.9
//
// Description:
// This class uses the information provided in the XML file and dictates the number
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

namespace CASSYS
{
    static class ReadFarmSettings
    {
        // Inputs or Parameters for the ReadFarmSettings Class
        public static string CurrentVersion = "0.9";            // The current version of CASSYS XML.
        public static XmlDocument doc;                          // The XML document that contains the Site, System, etc. definitions
        public static String CASSYSCSYXVersion;                 // The CASSYS XML Version Number obtained from the XML file
        public static bool UseDiffMeasured;                     // Using the Measured Diffuse on Horizontal Value
        public static bool UseMeasuredTemp;                     // Using measured temperature for the panels
        public static bool UsePOA;                              // Boolean that indicates if the program should use tilted irradiance to simulate
        public static bool UseGHI;                              // Boolean that indicates if the program should use horizontal irradiance to simulate
        public static bool UseWindSpeed;                        // Boolean indicating if the program has the Wind Speed available to use
        public static bool tempAmbDefined;                      // Boolean indicating if the program has Temp Ambient available to use
        public static bool batchMode = false;                   // Determines if the program is being run in batch mode (CMD prompt arguments for IO files) or not

        // Gathering Input and Output Configurations for the ReadFarmSettings Class
        public static String SimInputFile;                      // Input file path
        public static String SimOutputFile;                     // Output file path
        public static int SubArrayCount;                        // Number of system sub-arrays, default: 1
        public static int ClimateFileRowsToSkip;                // Number of rows to skip
        public static int TMYType;                              // If a TMY file is loaded, it specifies either 2 or 3 for .tm2 or .tm3 format
        public static string delim;                             // The character used in the input file between adjacent values
        public static int[] ClimateRefPos;                      // The positions of input data from the user input file
        
        // Output configuration for the Program
        public static Dictionary<String, dynamic> Outputlist;   // Creating a dictionary to hold all output values;
        public static List<String> OutputScheme;                // String Array that holds values based on user request
        public static String OutputHeader = null;               // The header of the output file, Always begins with TimeStamp.

        // Finding and assigning the Simulation Input and Output file name 
        public static void AssignIOFileNames()
        {
            // Getting the Input file Path
            SimInputFile = GetInnerText("Site", "InputFilePath", _Error: ErrLevel.FATAL);
            // Getting the Output file Path
            SimOutputFile = GetInnerText("Site", "OutputFilePath", _Error: ErrLevel.FATAL);
        }

        // Finding and assigning the SubArray Count from the attribute of the SubArray Tag
        public static void AssignSubArrayCount()
        {
            // Getting the Number of Sub-Arrays for this file
            SubArrayCount = int.Parse(GetXMLAttribute("System", "TotalArrays", _Error: ErrLevel.FATAL));
        }

        // Finding and Assigning the Input file style as per the SimSettingsFile
        public static void AssignInputFileSchema()
        {
            try
            {
                // Collecting file specific information.
                delim = GetInnerText("InputFile", "Delimeter", _Error: ErrLevel.FATAL);
                Util.AveragedAt = GetInnerText("InputFile", "AveragedAt", _Error: ErrLevel.FATAL);
                Util.timeFormat = GetInnerText("InputFile", "TimeFormat", _Error: ErrLevel.FATAL);
                Util.timeStep = double.Parse(GetInnerText("InputFile", "Interval", _Error: ErrLevel.FATAL));
                ClimateFileRowsToSkip = int.Parse(GetInnerText("InputFile", "RowsToSkip", _Error: ErrLevel.WARNING, _default: "0"));
                TMYType = int.Parse(GetInnerText("InputFile", "TMYType",_default:"-1"));

                // Initializing the array to use as a holder for column numbers.
                ClimateRefPos = new int[30];

                // Collecting weather variable locations in the file
                if((TMYType != 2) && (TMYType != 3))
                {
                    ClimateRefPos[0] = int.Parse(GetInnerText("InputFile", "TimeStamp", _Error: ErrLevel.FATAL));
                    UsePOA = Int32.TryParse(GetInnerText("InputFile", "GlobalRad", _Error: ErrLevel.WARNING, _default: "N/A"), out ClimateRefPos[2]);
                    UseGHI = Int32.TryParse(GetInnerText("InputFile", "HorIrradiance", _Error: ErrLevel.WARNING, _default: "N/A"), out ClimateRefPos[1]);
                    tempAmbDefined = Int32.TryParse(GetInnerText("InputFile", "TempAmbient", _Error: ErrLevel.FATAL),out ClimateRefPos[3]);
                    UseMeasuredTemp = Int32.TryParse(GetInnerText("InputFile", "TempPanel", _Error: ErrLevel.WARNING), out ClimateRefPos[4]);
                    UseWindSpeed = Int32.TryParse(GetInnerText("InputFile", "WindSpeed", _Error: ErrLevel.WARNING), out ClimateRefPos[5]);

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
                    if (tempAmbDefined == false && UseMeasuredTemp == false)
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
                    // Gathering Relevant nodes
                    Outputlist.Add(outNode.Name, null);
                    
                    // Creating a list of all items
                    OutputScheme.Add(outNode.Name);
                    
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
                            SubArrayTitle +=  "SubArray " + (arrayCount + 1) + " 1-Phase Inv. Voltage (V)," + "SubArray ";
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
                    else
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
                // Determine the Path of the XML requested
                    switch (Path)
                    {
                        case "Site":
                            Path = "/Site/" + NodeName;
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
                    }
                    
                    // Check if the XML Blank, if it is, return the default value
                    if (doc.SelectSingleNode(Path).InnerText == "")
                    {
                        if (_Error == ErrLevel.FATAL)
                        {
                            ErrorLogger.Log(NodeName + " is not defined. CASSYS requires this value to run.", ErrLevel.FATAL);
                            return "N/A";
                        }
                        else
                        {
                            ErrorLogger.Log(NodeName + " is not defined for this file. CASSYS assigned " + _default + " for this value.", ErrLevel.WARNING);
                            return _default;
                        }
                    }
                    else
                    {
                        return doc.SelectSingleNode(Path).InnerText;
                    }
            }
            catch (NullReferenceException)
            {
                if (_Error == ErrLevel.FATAL)
                {
                    ErrorLogger.Log(NodeName + " in " + Path + " is not defined in this XML Version. CASSYS requires this value to run.", ErrLevel.FATAL);
                    return "N/A";
                }
                else
                {
                    ErrorLogger.Log(NodeName + " in " + Path + " is not defined for this file. This may be because you are using site files created with an older version of CASSYS. CASSYS assigned " + _default + " for this value.", ErrLevel.WARNING);
                    return _default;
                }
            }
        }

        // Returns the attribute of the node, if the node and an attribute exist
        public static String GetXMLAttribute(String Path, String AttributeName, ErrLevel _Error = ErrLevel.WARNING, String _VersionNum = "0.9", String _Adder = null, int _ArrayNum = 0)
        {
            try
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
                }

                return doc.SelectSingleNode(Path).Attributes[AttributeName].Value;
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

            // Check if version number if specified.
            if (CASSYSCSYXVersion == "")
            {
                ErrorLogger.Log("The file does not have a valid version number. Please check the site file.", ErrLevel.FATAL);
            }


            // CASSYS Version check, if the version does not match, the program should exit.
            if (CASSYSCSYXVersion != CurrentVersion)
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
            Console.WriteLine("Copyright 2015 CanadianSolar, All rights reserved.");
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