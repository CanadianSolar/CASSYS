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
// NA - 2017-08-24: Version 1.2.0 Updated to include Iterative Mode. 
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
using System.Data;
using System.Windows.Forms;

namespace CASSYS
{
    class Simulation
    {
        // List and definition of variables used in the program file
        // Instance of all one-time objects, i.e. weather and radiation related objects
        SimMeteo SimMet = new SimMeteo();
        LossDiagram LossDiagram = new LossDiagram();            // initialize LossDiagram class only when needed
        RadiationProc SimRadProc;
        GridConnectedSystem SimGridSys;             
        ASTME2848 SimASTM;
        public String OutputString = "";
        StreamWriter OutputFileWriter;                          // stream writer used to write to output csv file
                        
        // Blank constructor for the CASSYS Program
        public Simulation()
        {
        }

        // Method to start the simulation software run.
        public void Simulate(XmlDocument SiteSettings, String _Input = null, String _Output = null)
        {
            // Adding timer to calculate elapsed time for the simulation.
            Stopwatch timeTaken = new Stopwatch();
            timeTaken.Start();

            // Delete error log just before simulation
            if(File.Exists(Application.StartupPath + "/ErrorLog.txt"))
            {
                File.Delete(Application.StartupPath + "/ErrorLog.txt");
            }

           // ReadFarmSettings reads the overall configuration of the farm
            ReadFarmSettings.doc = SiteSettings;

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

            // Collecting input and output file locations and Input/Output file location
            ReadFarmSettings.GetSimulationMode();

            // Inform user about site configuration:
            Console.WriteLine("Status: Configuring parameters");

            // Creating the SimMeteo Object to process the input file, and streamWriter to Write to the Output File
            try
            {
                // Reading and assigning the input file schema before configuring the inputs
                ReadFarmSettings.AssignInputFileSchema();

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

                try
                {
                    // Read through the input file and perform calculations
                    Console.Write("Status: Simulation Running");

                    // Creating the input file parser
                    SimMeteo SimMeteoParser = new SimMeteo();
                    SimMeteoParser.Config();

                    // Create StreamWriter
                    if (ReadFarmSettings.outputMode == "str")
                    {
                        // Assigning Headers to output String
                        OutputString = ReadFarmSettings.OutputHeader.TrimEnd(',') + '\n';
                    }
                    // Create StreamWriter and write output to .csv file
                    else if (ReadFarmSettings.outputMode == "csv")
                    {
                        // If in batch mode, use the appending overload of StreamWriter
                        if ((File.Exists(ReadFarmSettings.SimOutputFile)) && (ReadFarmSettings.batchMode))
                        {
                            OutputFileWriter = new StreamWriter(ReadFarmSettings.SimOutputFile, true);
                        }
                        else
                        {
                            // Writing the headers to the new output file.
                            OutputFileWriter = new StreamWriter(ReadFarmSettings.SimOutputFile);

                            // Writing the headers to the new output file.
                            OutputFileWriter.WriteLine(ReadFarmSettings.OutputHeader);
                        }
                    }
                    // The following is executed if using iterative mode
                    else
                    {
                        // Row for output header created only during first simulation run
                        if (ReadFarmSettings.runNumber == 1)
                        {
                            ReadFarmSettings.outputTable.Rows.InsertAt(ReadFarmSettings.outputTable.NewRow(), 0);
                        }

                        // Placing output header in data table
                        ReadFarmSettings.outputTable.Rows[0][ReadFarmSettings.runNumber -1] = ReadFarmSettings.OutputHeader.TrimEnd(',');
                    }

                    // Read through the Input File and perform the relevant simulation
                    while (!SimMeteoParser.InputFileReader.EndOfData)
                    {
                        SimMeteoParser.Calculate();

                        // If input file line could not be read, go to next line
                        if (!SimMeteoParser.inputRead)
                        {
                            continue;
                        }

                        // running required calculations based on simulation mode
                        switch (ReadFarmSettings.SystemMode)
                        {
                            case "ASTME2848Regression":
                                SimASTM.Calculate(SimMeteoParser);
                                break;

                            case "GridConnected":
                                SimRadProc.Calculate(SimMeteoParser);
                                SimGridSys.Calculate(SimRadProc, SimMeteoParser);
                                LossDiagram.Calculate();
                                break;

                            case "Radiation":
                                SimRadProc.Calculate(SimMeteoParser);
                                break;

                            default:
                                ErrorLogger.Log("Invalid Simulation mode found. CASSYS will exit.", ErrLevel.FATAL);
                                break;
                        }
                        // Only create output string in DLL mode, as creating string causes significant lag
                        if (ReadFarmSettings.outputMode == "str")
                        {
                            OutputString += String.Join(",", GetOutputLine()) + "\n";
                        }
                        // Write to output file
                        else if (ReadFarmSettings.outputMode == "csv")
                        {
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
                        // Using Iterative Mode: Writes output to a datatable
                        else
                        {
                            // Create row if this is the first simulation run
                            if (ReadFarmSettings.runNumber == 1)
                            {
                                ReadFarmSettings.outputTable.Rows.InsertAt(ReadFarmSettings.outputTable.NewRow(), ReadFarmSettings.outputTable.Rows.Count);
                            }

                            // Write output values to data table
                            ReadFarmSettings.outputTable.Rows[ErrorLogger.iterationCount][ReadFarmSettings.runNumber-1] = String.Join(",", GetOutputLine());
                        }
                    }

                    if (ReadFarmSettings.outputMode == "str")
                    {
                        // Trim newline character from the end of output
                        OutputString = OutputString.TrimEnd('\r', '\n');
                    }
                    // Write output to .csv file
                    else if (ReadFarmSettings.outputMode == "csv")
                    {
                        try
                        {
                            // Clean out the buffer of the writer to ensure all entries are written to the output file.
                            OutputFileWriter.Flush();
                            OutputFileWriter.Dispose();
                            SimMeteoParser.InputFileReader.Dispose();

                        }
                        catch (IOException ex)
                        {
                            ErrorLogger.Log(ex, ErrLevel.WARNING);
                        }
                    }

                    // If simulation done for Grid connected mode then write the calculated loss values to temp output file 
                    if(ReadFarmSettings.SystemMode == "GridConnected")
                    {
                        LossDiagram.AssignLossOutputs();
                    }

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
                    // In variable parameter mode, the timestamps are only written for the first simulation
                    string temp = ReadFarmSettings.Outputlist[required].ToString() + ",";
                    OutputLine += temp;
                }
                else
                {
                    OutputLine += "Not Available,";
                }
            }
            return OutputLine.TrimEnd(',');
        }
    }
}