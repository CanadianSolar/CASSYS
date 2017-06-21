// CASSYS - Grid connected PV system modelling software 
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Simulation.cs
// 
// Revision History:
// AP - 2014-10-14: Version 0.9
// AP - 2015-04-20: Veresion 0.9.1 Added user defined IAM profile for PV Module
// AP - 2015-05-22: Version 0.9.2 Added ability to perform irradiance only calculation
// AP - 2015-06-12: Version 0.9.2 Added separate detranspose method if meter and panel tilt do not match
// NA - 2017-06-09: Version 1.0.1 Modularity increased. Simulation determines mode and creates appropriate simulation instances.
// 
// Description 
// The main program reads the input file and output file and interacts with 
// all other classes that pertain to a grid connected solar farm 
// (PV Array, Inverter, Transformers, etc.)
//                                                  
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
//
//
///////////////////////////////////////////////////////////////////////////////


using System;
using System.Xml;
using System.IO;
using System.Diagnostics;

namespace CASSYS
{
    class Simulation
    {
        // List and definition of variables used in the program file
        // Instance of all one-time objects, i.e. weather and radiation related objects
        SimMeteo SimMet = new SimMeteo();
        RadiationProc SimRadProc;
        GridConnectedSystem SimGridSys;
        ASTME2848 SimASTM;

        // Blank constructor for the CASSYS Program
        public Simulation()
        {
        }

        // Method to start the simulation software run.
        public void Simulate(String XMLFileName, String _Input = null, String _Output = null)
        {
            // Adding timer to calculate elapsed time for the simulation.
            Stopwatch timeTaken = new Stopwatch();
            timeTaken.Start();

            // Get .CSYX file name from CMDPrompt, and Load document
            XmlDocument SiteSettings = new XmlDocument();
            try
            {
                SiteSettings.Load(XMLFileName);
                ErrorLogger.RunFileName = XMLFileName;
                ErrorLogger.Clean();
            }
            catch (Exception)
            {
                ErrorLogger.Log("CASSYS was unable to access or load the Site XML file. Simulation has ended.", ErrLevel.FATAL);
            }

            // ReadFarmSettings reads the overall configuration of the farm
            ReadFarmSettings.doc = SiteSettings;
            ReadFarmSettings.CheckCSYXVersion();

            // Obtain IO File names, either from the command prompt, or from .CSYX
            if ((_Input == null) && (_Output == null))
            {
                ReadFarmSettings.AssignIOFileNames();
            }
            else
            {
                ReadFarmSettings.SimInputFile = _Input;
                ReadFarmSettings.SimOutputFile = _Output;
            }

            // collecting input and output file locations and Input/Output file location
            ReadFarmSettings.GetSimulationMode();

            // Inform user about site configuration:
            Console.WriteLine("Status: Configuring parameters");

            // Creating the SimMeteo Object to process the input file, and streamWriter to Write to the Output File
            try
            {

                // Instantiating the relevant simulation class based on the simulation mode
                switch (ReadFarmSettings.SystemMode)
                {
                    case "ASTME2848Regression":
                        SimASTM = new ASTME2848();
                        SimASTM.Config();
                        break;

                    case "GridConnected":
                        // Assign the Sub-Array Count for grid connected systems.
                        ReadFarmSettings.AssignSubArrayCount();
                        SimGridSys = new GridConnectedSystem();
                        SimGridSys.Config();
                        goto case "Radiation";

                    case "Radiation":
                        SimRadProc = new RadiationProc();
                        SimRadProc.Config();
                        break;

                    default:
                        ErrorLogger.Log("Invalid Simulation mode found. CASSYS will exit.", ErrLevel.FATAL);
                        break;
                }

                // Reading and assigning the input file schema and the output file schema
                ReadFarmSettings.AssignOutputFileSchema();
                ReadFarmSettings.AssignInputFileSchema();

                // Notifying user of the Configuration status
                if (ErrorLogger.numWarnings == 0)
                {
                    Console.WriteLine("Status: Configuration OK.");
                }
                else
                {
                    Console.WriteLine("Status: There were problems encountered during configuration.");
                    Console.WriteLine("        Please see the error log file for details.");
                    Console.WriteLine("        CASSYS has configured parameters with default values.");
                }

                // Reading the first line and Writing the first line of the output file (Headers of columns)
                try
                {
                    StreamWriter OutputFileWriter;

                    // if in batch mode, use the appending overload of StreamWriter
                    if (File.Exists(ReadFarmSettings.SimOutputFile) && (ReadFarmSettings.batchMode))
                    {
                        OutputFileWriter = new StreamWriter(ReadFarmSettings.SimOutputFile, true);
                    }
                    else
                    {
                        // Writing the headers to the new output file.
                        OutputFileWriter = new StreamWriter(ReadFarmSettings.SimOutputFile);
                        OutputFileWriter.WriteLine(ReadFarmSettings.OutputHeader);
                    }

                    // Read through the input file and perform calculations
                    Console.Write("Status: Simulation Running");

                    // Creating the input file parser
                    SimMeteo SimMeteoParser = new SimMeteo();
                    SimMeteoParser.Config();

                    // Read through the Input File and perform the relevant simulation
                    while (!SimMeteoParser.InputFileReader.EndOfData)
                    {

                        SimMeteoParser.Calculate();

                        // Instantiating the relevant simulation class based on the simulation mode
                        switch (ReadFarmSettings.SystemMode)
                        {

                            case "ASTME2848Regression":
                                SimASTM.Calculate(SimMeteoParser);
                                break;

                            case "GridConnected":
                                SimRadProc.Calculate(SimMeteoParser);
                                SimGridSys.Calculate(SimRadProc, SimMeteoParser);
                                break;
                            case "Radiation":
                                SimRadProc.Calculate(SimMeteoParser);
                                break;

                            default:
                                ErrorLogger.Log("Invalid Simulation mode found. CASSYS will exit.", ErrLevel.FATAL);
                                break;
                        }

                        try
                        {
                            // Assembling and writing the line containing all output values
                            OutputFileWriter.WriteLine(String.Join(",", GetOutputLine()));
                        }
                        catch (IOException ex)
                        {
                            ErrorLogger.Log(ex, ErrLevel.WARNING);
                        }
                    }

                    // Clean out the buffer of the writer to ensure all entries are written to the output file.
                    OutputFileWriter.Flush();
                    SimMeteoParser.InputFileReader.Dispose();

                    timeTaken.Stop();

                    Console.WriteLine("");
                    Console.WriteLine("Status: Complete. Simulation took " + timeTaken.ElapsedMilliseconds / 1000D + " seconds.");

                }
                catch (IOException)
                {
                    ErrorLogger.Log("The output file name " + ReadFarmSettings.SimOutputFile + " is not accesible or is not available at the location provided.", ErrLevel.FATAL);
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("CASSYS encountered an unexpected error: " + ex.Message, ErrLevel.FATAL);
            }
        }

        // Gather the information for the output as per user selection.
        String GetOutputLine()
        {
            // Constructing the OutputLine to be written to string;
            string OutputLine = null;

            foreach (String required in ReadFarmSettings.OutputScheme)
            {
                if (ReadFarmSettings.Outputlist[required] != null)
                {
                    string temp = ReadFarmSettings.Outputlist[required].ToString() + ",";
                    OutputLine += temp;
                }
                else
                {
                    OutputLine += "Not Available,";
                }
            }

            return OutputLine;
        }
    }
}