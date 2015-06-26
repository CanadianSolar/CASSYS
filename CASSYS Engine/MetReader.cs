// CASSYS - Grid connected PV system modelling software 
// Version 0.9  
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: SimMeteo.cs, MetReader.cs
// 
// Revision History:
// ML - 2015-03-10: Version 0.9
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
using System.Text;
using Microsoft.VisualBasic.FileIO;


namespace CASSYS
{
    // Class that collects all the Meteorological Data
    class SimMeteo
    {
        // Set expected meteorological data fields 
        private double hGlo = double.NaN;                           // Horizontal Global Irradiance [W/m^2]
        private double hDiff = double.NaN;                          // Diffuse Horizontal Irradiance [W/m^2]
        private double tGlo = double.NaN;                           // Tilted Global Irradiance  [W/m^2]
        private double tAmbient = double.NaN;                       // Ambient temperature measured [deg C]  
        private double windSpeed = double.NaN;                      // Measured wind speed [m/s]
        private double tModMeasured = double.NaN;                   // Module Temperature Measured [deg C]
        private String timeStamp = null;                            // Time stamp Initializer [yyyy-mm-dd hh:mm:ss]

        // Set properties
        public double HGlo { get { return hGlo; } set { hGlo = value; } }
        public double HDiff { get { return hDiff; } set { hDiff = value; } }
        public double TGlo { get { return tGlo; } set { tGlo = value; } }
        public double TAmbient { get { return tAmbient; } set { tAmbient = value; } }
        public double WindSpeed { get { return windSpeed; } set { windSpeed = value; } }
        public double TModMeasured { get { return tModMeasured; } set { tModMeasured = value; } }
        public String TimeStamp { get { return timeStamp; } set { timeStamp = value; } }

    }

    // This is a reader that will read TMY file types, and also read Measured/User defined Input Files
    static class MetReader
    {  
        // Reads .TM2 files and provides the values to Simulation Object
        public static void ParseTM2Line(TextFieldParser InputFileReader,SimMeteo SimEnvironment)
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
            if (SimEnvironment.TimeStamp == null)
            {
                tStamp = new DateTime(Int32.Parse("19" + fields[0]), Int32.Parse(fields[1]), Int32.Parse(fields[2]), Int32.Parse(fields[3]), 0, 0); 
            }
            else
            {
                tStamp = DateTime.Parse(SimEnvironment.TimeStamp).AddHours(1);
            }

            // Parse fields 
            SimEnvironment.TimeStamp = (tStamp.ToString("yyyy-MM-dd HH:mm:ss"));
            Util.timeFormat = "yyyy-MM-dd HH:mm:ss";
            SimEnvironment.HGlo = double.Parse(fields[6]);
            SimEnvironment.HDiff = double.Parse(fields[12]);
            SimEnvironment.TAmbient = double.Parse(fields[33]) / 10.0;
            SimEnvironment.WindSpeed = double.Parse(fields[48])/10.0;
            ReadFarmSettings.UsePOA = false;
            ReadFarmSettings.UseDiffMeasured = true;
            ReadFarmSettings.UseWindSpeed = true;
            ReadFarmSettings.UseMeasuredTemp = false;
        }
       
        // Reads CSV Files and provides the values to Simulation Object
        public static void ParseCSVLine(TextFieldParser InputFileReader, SimMeteo SimEnvironment)
        {   
            // Read Input file line and split line based on the delimiter, and assign variables as defined above
            string[] inputLineDelimited = InputFileReader.ReadFields();
           
            try
            {
                // Get the Inputs from the Input file as assigned by the .CSYX file
                // The Input order is setup in weatherRefPos and then the input line is broken into its constituents based on the user assignment
                SimEnvironment.TimeStamp = inputLineDelimited[ReadFarmSettings.ClimateRefPos[0] - 1];
                
                if (ReadFarmSettings.UsePOA)
                {
                    SimEnvironment.TGlo = double.Parse(inputLineDelimited[ReadFarmSettings.ClimateRefPos[2] - 1]);
                }
                else
                {
                    if (ReadFarmSettings.UseDiffMeasured)
                    {
                        SimEnvironment.HDiff = double.Parse(inputLineDelimited[ReadFarmSettings.ClimateRefPos[6] - 1]);
                    }
                    SimEnvironment.HGlo = double.Parse(inputLineDelimited[ReadFarmSettings.ClimateRefPos[1] - 1]);
                }

                // If no system is defined proceed as normal to read and simulate file.
                if (!ReadFarmSettings.NoSystemDefined)
                {
                    // Measured temperature is available, try and access the value from the Input file else assign not a number status
                    if (ReadFarmSettings.UseMeasuredTemp)
                    {
                        SimEnvironment.TModMeasured = double.Parse(inputLineDelimited[ReadFarmSettings.ClimateRefPos[4] - 1]);
                    }
                    else
                    {
                        SimEnvironment.TModMeasured = double.NaN;
                    }

                    SimEnvironment.TAmbient = double.Parse(inputLineDelimited[ReadFarmSettings.ClimateRefPos[3] - 1]);

                    if (ReadFarmSettings.UseWindSpeed)
                    {
                        SimEnvironment.WindSpeed = double.Parse(inputLineDelimited[ReadFarmSettings.ClimateRefPos[5] - 1]);
                    }
                    else
                    {
                        SimEnvironment.WindSpeed = double.NaN;
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
        public static void ParseTM3Line(TextFieldParser InputFileReader, SimMeteo SimEnvironment)
        {

            // Read Input file line and split line based on the delimiter, and assign variables as defined above
            string[] inputLineDelimited = InputFileReader.ReadFields();
            DateTime simDateTime = new DateTime();

            try
            {
                // Get the Inputs from the Input file as assigned by the .CSYX file
                // The Input order is set up in weatherRefPos and then the input line is broken into its constituents based on the user assignment
                if (SimEnvironment.TimeStamp == null)
                {
                    simDateTime = DateTime.Parse(inputLineDelimited[0]);
                }
                else
                {
                    simDateTime = DateTime.Parse(SimEnvironment.TimeStamp).AddHours(Double.Parse(inputLineDelimited[1].Substring(0, inputLineDelimited[1].IndexOf(':'))));
                }

                SimEnvironment.TimeStamp = simDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                Util.timeFormat = "yyyy-MM-dd HH:mm:ss";
                SimEnvironment.TAmbient = double.Parse(inputLineDelimited[31]);
                SimEnvironment.HDiff = double.Parse(inputLineDelimited[10]);
                SimEnvironment.HGlo = double.Parse(inputLineDelimited[4]);
                SimEnvironment.WindSpeed = double.Parse(inputLineDelimited[46]);
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
    }
   
}
