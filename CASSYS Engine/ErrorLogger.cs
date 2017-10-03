// CASSYS - Grid connected PV system modelling software 
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Error Logger Class
// 
// Revision History:
// AP - 2014-10-14: Version 0.9
//
// Description:
// The error logging class contains a list of exceptions used by the program.
// It also contains the Error Logger method and specifics related to its Log file
// and the entries written to it.
//
// 
///////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Runtime.Serialization;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace CASSYS
{
    // List of custom exceptions for CASSYS
    // Generic CASSYSException - Used to throw exception is any problems occur during simulation.
    class CASSYSException : Exception, ISerializable
    {
        public CASSYSException()
            : base()
        {
        }

        public CASSYSException(string message)
            : base(message)
        {
        }

        public CASSYSException(string message, Exception inner)
            : base(message, inner)
        {
        }

        protected CASSYSException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

    }

    // Newton-Raphson Limit Exception - Used to throw exceptions if the N-R Method does not converge
    class NRException : Exception, ISerializable
    {
        public NRException()
            : base()
        {
        }

        public NRException(string message)
            : base("The Newton-Raphson Method did not converge within allowed iterations while calculating " + message)
        {
        }

        public NRException(string message, Exception inner)
            : base(message, inner)
        {
        }

        protected NRException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }

    // Blank Inner-node value for a given .CSYX node
    class CASSYSBlankXMLException : Exception, ISerializable
    {
        public CASSYSBlankXMLException()
            : base()
        {
        }

        public CASSYSBlankXMLException(string message)
            : base("The value for " + message + " is missing in the definition file. Simulation assumed 0 value.")
        {
        }

        public CASSYSBlankXMLException(string message, Exception inner)
            : base(message, inner)
        {
        }

        protected CASSYSBlankXMLException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }

    // Error Logging function
    public enum ErrLevel {INTERNAL, WARNING, FATAL};                    // Error Logging Level used to determine if the program should exit or not

    // Error Logging Method
    public static class ErrorLogger
    {
        // Locally used variables for the static class
        public static String RunFileName;                              // This is the name of the .CSYX file for this simulation (set during beginning of run)
        public static int numWarnings;                                 // Number of warnings issued to the User
        public static int iterationCount;                              // Number of iterations [#] 

        // Write to Error Log and output message to user (Type 1: Uses exceptions.)
        public static void Log(Exception err, ErrLevel Level)
        {
            if (Level == ErrLevel.INTERNAL)
                return;

            try
            {
                // Append to the file
                StreamWriter OutputStream = new StreamWriter(Application.StartupPath + "/ErrorLog.txt", true);

                // Attempt to write to file and close it
                OutputStream.WriteLine("Site Configuration File: " + RunFileName);
                OutputStream.WriteLine("Input line: " + iterationCount);
                OutputStream.WriteLine("Error Description: " + Level + ": " + err.Message);
                OutputStream.WriteLine("--------------------------------------------------------");
                OutputStream.Close();
            }
            catch (Exception e)
            {
                // If there was a problem tell the user
                Console.WriteLine("The Error Log is not accessible!" + e.ToString());
                Level = ErrLevel.FATAL;
            }

            // Using switch statements to determine if the program should exit
            switch (Level)
            {
                case ErrLevel.FATAL:
                    Console.WriteLine("An error occurred while running CASSYS. Please check the Error Log at " + Application.StartupPath + " for details.");
                    Environment.Exit(1);
                    break;
                case ErrLevel.WARNING:
                    numWarnings++;
                    break;
            }
        }

        // Write to Error Log and output message to user (Type 2: Does not use exceptions.)
        public static void Log(String Message, ErrLevel Level)
        {
            if (Level == ErrLevel.INTERNAL)
                return;

            // Append to the file
            StreamWriter OutputStream = new StreamWriter(Application.StartupPath + "/ErrorLog.txt", true);

            // Attempt to write to file and close it
            OutputStream.WriteLine("Site Configuration File: " + RunFileName);
            OutputStream.WriteLine("Input File line: " + iterationCount);
            OutputStream.WriteLine("Error Description: " + Level + ": " + Message);
            OutputStream.WriteLine("--------------------------------------------------------");
            OutputStream.Close();

            // Using switch statements to determine if the program should exit
            switch (Level)
            {
                case ErrLevel.FATAL:
                    Console.WriteLine("CASSYS has stopped the simulation. Please check the end of the Error Log at " + Application.StartupPath + " for details.");
                    Environment.Exit(1);
                    break;
                case ErrLevel.WARNING:
                    numWarnings++;
                    break;
            }
        }

        // Check assertion, and if not true, then write to error log.
        public static void Assert(String Message, bool check, ErrLevel Level)
        {
            if (check)
            {
                // Do nothing.
            }
            else
            {
                ErrorLogger.Log(Message, Level);
            }

        }

        // Method to clean out the log file and close it before writing the errors for the the simulation that is currently running.
        public static void Clean()
        {
            // Clean out the log file so new errors may be written into it
            StreamWriter OutputStream = new StreamWriter(Application.StartupPath + "/ErrorLog.txt", false);

            // Adding header to the file once it has been cleaned
            OutputStream.WriteLine("CASSYS - Simulation Error Log for : " + RunFileName);
            OutputStream.Close();
        }
    }
}
