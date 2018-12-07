// CASSYS - Grid connected PV system modelling software 
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
using System.Linq;
using Microsoft.VisualBasic.FileIO;


namespace CASSYS
{
    // Class that collects all the Meteorological Data
    public class SimMeteo
    {
        // Initializing output variables
        public double HGlo = double.NaN;                            // Horiztal Global Irradiance [W/m^2]
        public double HDiff = double.NaN;                           // Horizontal Diffuse Irradiance [W/m^2]
        public double TGlo = double.NaN;                            // Titled Irradiance [W/m^2]
        public double TAmbient = double.NaN;                        // Ambient temperature [C]
        public double WindSpeed = double.NaN;                       // Windspeed [m/s]
        public double TModMeasured = double.NaN;                    // Measured module temperature [C]
        public double Albedo = double.NaN;                          // Albedo []
        public String TimeStamp = null;                             // Timestamp initialized to null
        public int Year;                                            // Holds current year
        public int DayOfYear;                                       // Holds current day of year
        public int DayOfMonth;                                      // Holds current day of month
        public double HourOfDay;                                    // Holds current hour of day
        public double minuteInterval;                               // Holds the interval at which the weather was collected
        public int MonthOfYear;                                     // Holds current month of year
        public double TimeStepEnd;                                  // Time at which timestamp begins
        public double TimeStepBeg;                                  // Time at which timestamp ends
        public TextFieldParser InputFileReader;                     // Used to read climate values from input file
        public bool inputRead;                                      // Used to skip calculations and output file for a timestamp if unable to read meteological data
        int numOfSkippedInput = 0;                                  // Number of input lines CASSYS was unable to read

        // Blank Constructor for SimMeteo
        public SimMeteo()
        {
        }

        // Reads the input file and assigns the available outputs and time stamp
        public void Calculate()
        {
            inputRead = true;
            // Keeping a track of how many times this method was accessed, and printing progress on console window.
            if (ErrorLogger.iterationCount % 1000 == 0)
            {
                Console.Write('.');
            }
            ErrorLogger.iterationCount++;

            // Assigning outputs of the SimMeteo class based on file type.
            if(ReadFarmSettings.TMYType == 1)
            {
                ParseEPWLine();
            }
            else if (ReadFarmSettings.TMYType == 2)
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
            Utilities.TSBreak(TimeStamp, out DayOfYear, out HourOfDay, out Year, out MonthOfYear, out TimeStepEnd, out TimeStepBeg, this);

            // Do not assign outputs if unable to read line of input file
            if (!inputRead)
            {
                if (numOfSkippedInput > ReadFarmSettings.IncClimateRowsAllowed)
                {
                    ErrorLogger.Log("CASSYS was unable to read a total of " + ReadFarmSettings.IncClimateRowsAllowed + " climate file lines. Please correct climate file format. CASSYS has exited.", ErrLevel.FATAL);
                }
                else
                {
                    numOfSkippedInput++;
                }
                return;
            }

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

        // Reads EPW Files and provides the values to simulation object
        public void ParseEPWLine()
         {
            // Read Input file line and split line based on the delimiter, and assign variables as defined above
            string[] inputLineDelimited = InputFileReader.ReadFields();
            DateTime dateAndTime;
            try
            {
                Year = int.Parse(inputLineDelimited[0]); // not used
                MonthOfYear = int.Parse(inputLineDelimited[1]);
                DayOfMonth = int.Parse(inputLineDelimited[2]);
                HourOfDay = double.Parse(inputLineDelimited[3]);
                minuteInterval = double.Parse(inputLineDelimited[4]);

                // EPW files are assumed to be in hourly intervals, so this is a check to ensure the assumption is true
                if(minuteInterval != 60 & minuteInterval != 0)
                {
                    ErrorLogger.Log("EPW file is not in 60 minute intervals", ErrLevel.FATAL);
                }

                // Year is permanantly set to 2017 to prevent chronological error 
                dateAndTime = new DateTime(2017, MonthOfYear, DayOfMonth).AddHours(HourOfDay);

                TAmbient = double.Parse(inputLineDelimited[6]);
                HGlo = double.Parse(inputLineDelimited[13]);
                HDiff = double.Parse(inputLineDelimited[15]);
                WindSpeed = double.Parse(inputLineDelimited[21]);
                
                TimeStamp = dateAndTime.ToString("yyyy-MM-dd HH:mm:ss");
                Util.timeFormat = "yyyy-MM-dd HH:mm:ss";
                ReadFarmSettings.UsePOA = false;
                ReadFarmSettings.UseDiffMeasured = true;
                ReadFarmSettings.UseWindSpeed = true;
                ReadFarmSettings.UseMeasuredTemp = false;
            }
            catch
            {
                ErrorLogger.Log("Error produced in loading EPW", ErrLevel.FATAL);
                inputRead = false;
            }
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
                // As of v. 1.3.1 date is handled first to correct potential issues with Feb 28 of leap year
                simDateTime = DateTime.Parse(inputLineDelimited[0]);
                simDateTime = new DateTime(1990, simDateTime.Month, simDateTime.Day).AddHours(Double.Parse(inputLineDelimited[1].Substring(0, inputLineDelimited[1].IndexOf(':'))));

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
                ErrorLogger.Log("One of the Input columns was not defined correctly. Please check your Input file definition. Row was skipped.", ErrLevel.WARNING);
                inputRead = false;
            }
            catch (FormatException)
            {
                ErrorLogger.Log("Incorrect format for values in the Input String. Please check your Input file at the Input line specified above. Row was skipped.", ErrLevel.WARNING);
                inputRead = false;
            }
            catch (CASSYSException ex)
            {
                ErrorLogger.Log(ex, ErrLevel.WARNING);
            }
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

                    if (ReadFarmSettings.UseMeasuredAlbedo)
                    {
                        Albedo = double.Parse(inputLineDelimited[ReadFarmSettings.ClimateRefPos[7] - 1]);
                    }
                    else
                    {
                        Albedo = double.NaN;
                    }
                }
            }
            catch (IndexOutOfRangeException)
            {
                ErrorLogger.Log("Incorrect number of fields in the above line. Row was skipped.", ErrLevel.WARNING);
                inputRead = false;
            }
            catch (FormatException)
            {
                ErrorLogger.Log("Incorrect format for values in the Input String. Please check your Input file at the Input line specified above. Row was skipped.", ErrLevel.WARNING);
                inputRead = false;
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
                InputFileReader.SetDelimiters(ReadFarmSettings.Delim);

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
