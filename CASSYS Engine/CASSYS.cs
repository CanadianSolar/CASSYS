// CASSYS - Grid connected PV system modelling software 
// Version 0.9  
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
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using CASSYS;
using System.Xml.XPath;

namespace CASSYS
{
    class CASSYS
    {
        static void Main(string[] args)
        {
            // Showing the Header of the Simulation Program in the Console
            ReadFarmSettings.ShowHeader();

            // Declaring a new simulation object
            Simulation PVPlant = new Simulation();
            
            if (args.Length == 1)
            {
                // Running the Simulation based on the CASSYS Configuration File provided
                ReadFarmSettings.batchMode = false;
                PVPlant.Simulate(args[0]);

                // Show the end of simulation message, window should be kept open for longer
                ReadFarmSettings.ShowFooter();
            }
            else if (args.Length == 0)
            {
                // Set batch mode to true, and Ask user for all the arguments
                Console.WriteLine("Status: You are running CASSYS in batch mode.");
                Console.WriteLine("Status: You will need to provide a CASSYS Site File (.CSYX), Input File, and Output File Path.");
                Console.WriteLine("Enter CASSYS Site (.CSYX) File Path: ");
                String SimFilePath = Console.ReadLine();
                Console.WriteLine("Enter Input file path (.csv file): ");
                String SimInputFilePath = Console.ReadLine();
                Console.WriteLine("Enter an Output file path (.csv file, Note: if the file exists this will append your results to the file): ");
                String SimOutputFilePath = Console.ReadLine();
                
                // Send arguments to simulation object and run.
                PVPlant.Simulate(SimFilePath.Replace("\\", "/"), _Input: SimInputFilePath.Replace("\\", "/"), _Output: SimOutputFilePath.Replace("\\", "/"));
            }
            else if (args.Length == 3)
            {
                // Set batch mode to true, and Run from command prompt arguments directly
                ReadFarmSettings.batchMode = true;
                PVPlant.Simulate(args[0], _Input: args[1], _Output: args[2]);
            }
            else
            {
                // CASSYS needs a site file name to run, so warn the user and exit the program.
                ErrorLogger.Log("No site file provided for CASSYS to simulate. Please select a valid .CSYX file.", ErrLevel.FATAL);
            }
        }
    }
}
