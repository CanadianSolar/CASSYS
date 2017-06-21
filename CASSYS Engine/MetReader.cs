﻿// CASSYS - Grid connected PV system modelling software 
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: SimMeteo.cs, MetReader.cs
// 
// Revision History:
// ML - 2015-03-10: Version 0.9
// NA - 2017-06-09: Version 1.1 - Changes made to the structure of the file to include class specific outputs and calc methods
//
// Description 
// This class is used to read different input file types (TMY, and User defined)
// and assign the meteorological values to the SimMeteo class, used in the simulation.
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
using System.IO;
using System.Text;
using Microsoft.VisualBasic.FileIO;


namespace CASSYS
{
    // Class that collects all the Meteorological Data
    class SimMeteo
    {
        // Initializing output variables
        public double HGlo = double.NaN;                            // Horiztal Global Irradiance [W/m^2]
        public double HDiff = double.NaN;                           // Horizontal Diffuse Irradiance [W/m^2]
        public double TGlo = double.NaN;                            // Titled Irradiance [W/m^2]
        public double TAmbient = double.NaN;                        // Ambient temperature [C]
        public double WindSpeed = double.NaN;                       // Windspeed [m/s]
        public double TModMeasured = double.NaN;                    // Measured module temperature [C]
        public String TimeStamp = null;                             // Timestamp initialized to null
        public int lastDOY = 0;                                     // Holds the day of year that was simulated in the last interval.
        public int DayOfYear;                                       // Holds current day of year
        public double HourOfDay;                                    // Holds current hour of day
        public int MonthOfYear;                                     // Holds current month of year
        public double TimeStepEnd;                                  // Time at which timestamp begins
        public double TimeStepBeg;                                  // Time at which timestamp ends
        public TextFieldParser InputFileReader;                     // Used to read climate values from input file
        

        // Blank Constructor for SimMeteo
        public SimMeteo()
        {
        }
        
        // Reads the input file and assigns the available outputs and time stamp
        public void Calculate()
        {
            // Keeping a track of how many times this method was accessed, and printing progress on console window.
            if (ErrorLogger.iterationCount % 1000 == 0)
            {
                Console.Write('.');
            }
            ErrorLogger.iterationCount++;

            // Assigning outputs of the SimMeteo class based on file type.
            if (ReadFarmSettings.TMYType == 2)
            {
                ParseTM2Line();
            }
            else if (ReadFarmSettings.TMYType == 3)
            {
                ParseTM3Line();
            }
            else
            {
                ParseCSVLine();
            }

            // Determining DayOfYear, HOD, MOY, TimeStamp end and beginning and assigning these to the SimMeteoClass
            Utilities.TSBreak(TimeStamp, out DayOfYear, out HourOfDay, out MonthOfYear, out TimeStepEnd, out TimeStepBeg);

            // Assigning outputs for this interval.
            AssignOutputs();

        }

        // These outputs will be written to the file from this class.
        public void AssignOutputs()
        {
            //  Assigning all outputs their corresponding values;
            ReadFarmSettings.Outputlist["Input_Timestamp"] = TimeStamp;
            ReadFarmSettings.Outputlist["Timestamp_Used_for_Simulation"] = String.Format("{0:u}", TimeStamp).Replace('Z', ' ');
            ReadFarmSettings.Outputlist["Global_Irradiance_in_Array_Plane"] = TGlo;
            ReadFarmSettings.Outputlist["Ambient_Temperature"] = TAmbient;
            ReadFarmSettings.Outputlist["Wind_Velocity"] = WindSpeed;
            ReadFarmSettings.Outputlist["Measured_Module_Temperature__deg_C_"] = TModMeasured;
            ReadFarmSettings.Outputlist["Average_Ambient_Temperature_deg_C_"] = TAmbient;
        }

        // Reads .TM2 files and provides the values to Simulation Object
        public void ParseTM2Line()
        {

            // Initializing the array used to store the fields from the TMY file; the field number corresponds to the column number in the Excel file (minus one because of zero index)
            string[] fields = InputFileReader.ReadFields();
            DateTime tStamp = new DateTime();

            // Format the fields
            for (int i = 0; i < fields.Length; i++)
            {
                //Remove leading zeroes in negative values
                if (fields[i].Contains('-'))
                {
                    fields[i] = fields[i].TrimStart('-', '0');
                    fields[i] = '-' + fields[i];
                }

                //Remove leading zeroes from non-zero values
                if ((fields[i].TrimStart('0')).Length != 0)
                {
                    fields[i] = fields[i].TrimStart('0');
                }
                //Remove leading zeroes from zero values
                else
                {
                    fields[i] = "0";
                }
            }
            //Write raw data into array
            if (TimeStamp == null)
            {
                tStamp = new DateTime(Int32.Parse("19" + fields[0]), Int32.Parse(fields[1]), Int32.Parse(fields[2]), Int32.Parse(fields[3]), 0, 0);
            }
            else
            {
                tStamp = DateTime.Parse(TimeStamp).AddHours(1);
            }

            // Parse fields 
            TimeStamp = (tStamp.ToString("yyyy-MM-dd HH:mm:ss"));
            Util.timeFormat = "yyyy-MM-dd HH:mm:ss";
            HGlo = double.Parse(fields[6]);
            HDiff = double.Parse(fields[12]);
            TAmbient = double.Parse(fields[33]) / 10.0;
            WindSpeed = double.Parse(fields[48]) / 10.0;
            ReadFarmSettings.UsePOA = false;
            ReadFarmSettings.UseDiffMeasured = true;
            ReadFarmSettings.UseWindSpeed = true;
            ReadFarmSettings.UseMeasuredTemp = false;

        }

        // Reads CSV Files and provides the values to Simulation Object
        public void ParseCSVLine()
        {
            // Read Input file line and split line based on the delimiter, and assign variables as defined above
            string[] inputLineDelimited = InputFileReader.ReadFields();

            try
            {
                // Get the Inputs from the Input file as assigned by the .CSYX file
                // The Input order is setup in weatherRefPos and then the input line is broken into its constituents based on the user assignment
                TimeStamp = inputLineDelimited[ReadFarmSettings.ClimateRefPos[0] - 1];

                if (ReadFarmSettings.UsePOA)
                {
                    TGlo = double.Parse(inputLineDelimited[ReadFarmSettings.ClimateRefPos[2] - 1]);
                }
                else
                {
                    if (ReadFarmSettings.UseDiffMeasured)
                    {
                        HDiff = double.Parse(inputLineDelimited[ReadFarmSettings.ClimateRefPos[6] - 1]);
                    }
                        HGlo = double.Parse(inputLineDelimited[ReadFarmSettings.ClimateRefPos[1] - 1]);
                }

                // If no system is defined proceed as normal to read and simulate file.
                if (ReadFarmSettings.SystemMode != "Radiation")
                {
                    // Measured temperature is available, try and access the value from the Input file else assign not a number status
                    if (ReadFarmSettings.UseMeasuredTemp)
                    {
                        TModMeasured = double.Parse(inputLineDelimited[ReadFarmSettings.ClimateRefPos[4] - 1]);
                    }
                    else
                    {
                        TModMeasured = double.NaN;
                    }

                    TAmbient = double.Parse(inputLineDelimited[ReadFarmSettings.ClimateRefPos[3] - 1]);

                    if (ReadFarmSettings.UseWindSpeed)
                    {
                        WindSpeed = double.Parse(inputLineDelimited[ReadFarmSettings.ClimateRefPos[5] - 1]);
                    }
                    else
                    {
                        WindSpeed = double.NaN;
                    }
                }
            }
            catch (IndexOutOfRangeException)
            {
                ErrorLogger.Log("One of the Input columns was not defined correctly. Please check your Input file definition. Simulation has ended.", ErrLevel.FATAL);
            }
            catch (FormatException)
            {
                ErrorLogger.Log("Incorrect format for values in the Input String. Please check your Input file at the Input line specified above. Simulation has ended.", ErrLevel.FATAL);
            }
            catch (CASSYSException ex)
            {
                ErrorLogger.Log(ex, ErrLevel.WARNING);
            }
        }

        // Reads .TM3 Files and provides the value to Simulation Object
        public void ParseTM3Line()
        {

            // Read Input file line and split line based on the delimiter, and assign variables as defined above
            string[] inputLineDelimited = InputFileReader.ReadFields();
            DateTime simDateTime = new DateTime();

            try
            {
                // Get the Inputs from the Input file as assigned by the .CSYX file
                // The Input order is set up in weatherRefPos and then the input line is broken into its constituents based on the user assignment


                simDateTime = DateTime.Parse(inputLineDelimited[0]).AddHours(Double.Parse(inputLineDelimited[1].Substring(0, inputLineDelimited[1].IndexOf(':'))));
                simDateTime = new DateTime(1990, simDateTime.Month, simDateTime.Day, simDateTime.Hour, simDateTime.Minute, simDateTime.Second);

                if (simDateTime.DayOfYear >= lastDOY)
                {
                    lastDOY = simDateTime.DayOfYear;
                }
                else
                {
                    simDateTime = new DateTime(simDateTime.Year + 1, simDateTime.Month, simDateTime.Day, simDateTime.Hour, simDateTime.Minute, simDateTime.Second);
                    lastDOY = simDateTime.DayOfYear;
                }
                
                TimeStamp = simDateTime.ToString("yyyy-MM-dd HH:mm:ss");

                Util.timeFormat = "yyyy-MM-dd HH:mm:ss";
                TAmbient = double.Parse(inputLineDelimited[31]);
                HDiff = double.Parse(inputLineDelimited[10]);
                HGlo = double.Parse(inputLineDelimited[4]);
                WindSpeed = double.Parse(inputLineDelimited[46]);
                ReadFarmSettings.UsePOA = false;
                ReadFarmSettings.UseDiffMeasured = true;
                ReadFarmSettings.UseWindSpeed = true;
                ReadFarmSettings.UseMeasuredTemp = false;
            }
            catch (IndexOutOfRangeException)
            {
                ErrorLogger.Log("One of the Input columns was not defined correctly. Please check your Input file definition. Simulation has ended.", ErrLevel.FATAL);
            }
            catch (FormatException)
            {
                ErrorLogger.Log("Incorrect format for values in the Input String. Please check your Input file at the Input line specified above. Simulation has ended.", ErrLevel.FATAL);
            }
            catch (CASSYSException ex)
            {
                ErrorLogger.Log(ex, ErrLevel.WARNING);
            }
        }

        // Configuration method for the SimMeteo Class.
        public void Config()
        {
            // Try to access the input file and exit if there are any issues.
            try
            {
                InputFileReader = new TextFieldParser(ReadFarmSettings.SimInputFile);
            
                // Skip # of rows that are defined by the User in the .CSYX
                for (int i = 0; i < ReadFarmSettings.ClimateFileRowsToSkip; i++)
                {
                    InputFileReader.ReadLine();
                }

                // Configuring the delimiter for the file.
                InputFileReader.SetDelimiters(ReadFarmSettings.delim);

                if (ReadFarmSettings.TMYType == 2)
                {
                    InputFileReader.TextFieldType = FieldType.FixedWidth;
                    InputFileReader.SetFieldWidths(3, 2, 2, 2, 4, 4, 4, 1, 1, 4, 1, 1, 4, 1, 1, 4, 1, 1, 4, 1, 1, 4, 1, 1, 4, 1, 1, 2, 1, 1, 2, 1, 1, 4, 1, 1, 4, 1, 1, 3, 1, 1, 4, 1, 1, 3, 1, 1, 3, 1, 1, 4, 1, 1, 5, 1, 1, 1, 3, 1, 1, 3, 1, 1, 3, 1, 1, 2, 1, 1);
                }
            }
            catch (Exception)
            {
                ErrorLogger.Log("There was a problem accessing the Climate File. Please check if the file is open or if the path to the file is valid.", ErrLevel.FATAL);
            }
        }
    } 
}
