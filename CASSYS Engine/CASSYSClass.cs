// CASSYS - Grid connected PV system modelling software 
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: CASSYS.cs
// 
// Revision History:
// AP - 2015-01-15: Version 0.9
//
// Description 
// The main program calls the Simulation class and provides its run method with 
// the .CSYX file it should simulate. 
//                                                  
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
// N/A
//
///////////////////////////////////////////////////////////////////////////////

using System;
using System.Xml;
using System.IO;
using System.Diagnostics;
using System.Data;
using System.Collections.Generic;
using System.Linq;

namespace CASSYS
{
   public class CASSYSClass
    {
        static Simulation PVPlant;
        static XmlDocument SiteSettings;                // CSYX file site specifications
        static String SimOutputFilePath;
        static String SimInputFilePath;

        static string paramPath;
        static double start;
        static double end;
        static double interval;
        static StreamWriter OutputFileWriter;

        public static void Main(string[] args)
        {

            // Showing the Header of the Simulation Program in the Console
            if (ReadFarmSettings.outputMode != "str")
            {
                ReadFarmSettings.ShowHeader();
            }

            // Declaring a new simulation object
            PVPlant = new Simulation();
            SiteSettings = new XmlDocument();

            // Load CSYX file
            try
            {
                if (args.Length == 0)
                {
                    // Set batch mode to true, and Ask user for all the arguments
                    Console.WriteLine("Status: You are running CASSYS in batch mode.");
                    Console.WriteLine("Status: You will need to provide a CASSYS Site File (.CSYX), Input File, and Output File Path.");
                    Console.WriteLine("Enter CASSYS Site (.CSYX) File Path: ");
                    String SimFilePath = Console.ReadLine();
                    Console.WriteLine("Enter Input file path (.csv file): ");
                    SimInputFilePath = Console.ReadLine();
                    Console.WriteLine("Enter an Output file path (.csv file, Note: if the file exists this will append your results to the file): ");
                    SimOutputFilePath = Console.ReadLine();

                    SiteSettings.Load(SimFilePath.Replace("\\", "/"));
                    ErrorLogger.RunFileName = SimFilePath;
                    ErrorLogger.Clean();
                }
                else
                {
                    // Get .CSYX file name from CMDPrompt, and Load document
                    SiteSettings.Load(args[0]);
                    ErrorLogger.RunFileName = args[0];
                }
            
            // Assigning Xml document to ReadFarmSettings to allow reading of XML document. Check CSYX version
            ReadFarmSettings.doc = SiteSettings;
            ReadFarmSettings.CheckCSYXVersion();  

            // Variable Parameter mode
            if (ReadFarmSettings.doc.SelectSingleNode("/Site/Iterations/Iteration1") != null)
            {
                variableParameters(args);
            }
            else if (args.Length == 1)
            {
                // Running the Simulation based on the CASSYS Configuration File provided
                ReadFarmSettings.batchMode = false;
                PVPlant.Simulate(SiteSettings);
                // Show the end of simulation message, window should be kept open for longer
                ReadFarmSettings.ShowFooter();
            }
            else if (args.Length == 3)
            {
                // Set batch mode to true, and Run from command prompt arguments directly
                ReadFarmSettings.batchMode = true;
                PVPlant.Simulate(SiteSettings, _Input: args[1], _Output: args[2]);
            }
            else if (args.Length == 0)
            {
                PVPlant.Simulate(SiteSettings, _Input: SimInputFilePath.Replace("\\", "/"), _Output: SimOutputFilePath.Replace("\\", "/"));
            }
            else
            {
                // CASSYS needs a site file name to run, so warn the user and exit the program.
                ErrorLogger.Log("No site file provided for CASSYS to simulate. Please select a valid .CSYX file.", ErrLevel.FATAL);
            }
            }
            catch (Exception)
            {
                ErrorLogger.Log("CASSYS was unable to access or load the Site XML file. Simulation has ended.", ErrLevel.FATAL);
            }
        }

        public string GetOutputString()
        {
            return PVPlant.OutputString;
        }

        static void variableParameters(string[] args)
        {
            ReadFarmSettings.outputMode = "var";
            ReadFarmSettings.batchMode = true;

            // Assign input and output file path for simulation run.
            // File paths have already been assigned in case of no input arguments 
            if (args.Length == 1)
            {
                SimInputFilePath = null;
                SimOutputFilePath = null;
            }
            else if (args.Length == 3)
            {
                SimInputFilePath = args[1];
                SimOutputFilePath = args[2];
            }

            // Gathering information on parameter to be changed
            paramPath = ReadFarmSettings.GetAttribute("Iteration1", "ParamPath", _Error: ErrLevel.FATAL);
            start = double.Parse(ReadFarmSettings.GetAttribute("Iteration1", "Start", _Error: ErrLevel.FATAL));
            end = double.Parse(ReadFarmSettings.GetAttribute("Iteration1", "End", _Error: ErrLevel.FATAL));
            interval = double.Parse(ReadFarmSettings.GetAttribute("Iteration1", "Interval", _Error: ErrLevel.FATAL));

            // Number of simulations ran
            ReadFarmSettings.runNumber = 1;

            // Data table to hold output for various runs
            ReadFarmSettings.outputTable = new DataTable();

            // Row used to hold parameter information
            ReadFarmSettings.outputTable.Rows.InsertAt(ReadFarmSettings.outputTable.NewRow(), 0);

            // Loop through values of variable parameter
            for (double value = start; value <= end; value = value + interval)
            {
                // iterationCount used to determine current row in data base
                ErrorLogger.iterationCount = 0;
                // Reset outputheader between runs
                ReadFarmSettings.OutputHeader = null;

                // Vary parameter in XML object
                SiteSettings.SelectSingleNode(paramPath).InnerText = Convert.ToString(value);

                // Creating column to store data from run
                // A single column stores the all data from the simulation run
                ReadFarmSettings.outputTable.Columns.Add(paramPath + "=" + value, typeof(String));

                PVPlant.Simulate(SiteSettings, SimInputFilePath, SimOutputFilePath);
                ReadFarmSettings.runNumber++;
            }

            try
            {
                // Write data table output to csv file
                OutputFileWriter = new StreamWriter(ReadFarmSettings.SimOutputFile);
                WriteDataTable(ReadFarmSettings.outputTable, OutputFileWriter);
                OutputFileWriter.Dispose();
            }
            catch (IOException ex)
            {
                ErrorLogger.Log("Error occured while creating to output file. Error: " + ex, ErrLevel.FATAL);
            }
        }

        static void WriteDataTable(DataTable SourceTable, TextWriter writer)
        {
            string columnStr = "";
            string firstSimCommas = new string(',', ReadFarmSettings.OutputScheme.Count + 2);
            string StandardCommas = new string(',', ReadFarmSettings.OutputScheme.Count);
            foreach(DataColumn column in SourceTable.Columns)
            {
                if (column.Ordinal == 0)
                    columnStr += column + firstSimCommas;
                else
                    columnStr += column + StandardCommas;
            }

            writer.WriteLine(columnStr);
         
            IEnumerable<String> items = null;

            foreach (DataRow row in SourceTable.Rows)
            {
                items = row.ItemArray.Select(o => o.ToString());
                writer.WriteLine(String.Join(",", items));
            }

            writer.Flush();
        }

        private static string QuoteValue(string value)
        {
            return String.Concat("\"", value.Replace("\"", "\"\""), "\"");
        }
    }
}
