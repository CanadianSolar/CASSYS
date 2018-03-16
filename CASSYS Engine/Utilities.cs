// CASSYS - Grid connected PV system modelling software 
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Utilities Class
// 
// Revision History:
// AP - 2014-10-14: Version 0.9
//
// Description
// The Utility class contains functions that are either of general use or that
// are used by other objects in the library. These may include conversions between
// physical units, determining certain values such as day or month, etc and other 
// general purpose calculations/set-ups. The Utility class also has a Util class
// which includes all constants (physical or otherwise) used in the program.
//
//  
///////////////////////////////////////////////////////////////////////////////

using System;

namespace CASSYS
{
    // The class Util holds all the constants used in the program 
    public class Util
    {
        // Physical constants
        static public double SOLAR_CONST = 1367;              // Solar constant in W/m2 
        static public double BOLTZMANN_CONST = 1.38066e-23;   // Boltzmann constant in J/K 
        static public double COULOMB = 1.60218e-19;           // Elementary Charge 
        static public double RGAS = 8314.41;                  // Universal gas constant in J/kmol/K 
        static public double MEAN_EARTH_RADIUS = 6371000;     // Mean earth radius [m] 

        // Constants Used in the PV Array Class
        static public double SiBANDGAP = 1.12;                // The silicon band-gap [eV]
        static public double ELEMCHARGE = 1.602e-19;          // The elementary charge of the electron [Coulomb]
        static public double BOLTZMANNCONST = 1.38e-23;       // The Boltzmann constant [m^2 kg s^-2 K^-1] 
        static public double GOLDEN = 0.61803398875;          // Golden ratio, used in GetVmpp Method
        static public int NRLIMIT = 30;                       // Maximum number of iterations allowed for NR Methods
        static public double DiffInciAng = 1.04719755;        // 60 degrees expressed in radians [radians], assumed incidence angle for diffuse component of irradiance

        static public int NUM_GROUND_SEGS = 100;              // Number of segments into which to divide up the ground for ground irradiance calculation [#]
        static public int NUM_CELLS_PANEL = 6;                // Number of cells in a PV panel

        // Other numerical constants
        static public double COS_LITTLE = 0.999993908;        // Cosine of 0.2 degrees. Used in Astro Class.
        static public int BADDATA = -9999;                    // Bad data marker based on Arithmetic errors during calculation
        static public double timeStep;                        // Time step between time stamps [minutes], set in ReadFarmSettings  
        static public String AveragedAt;                      // The type of averaging used for the interval of the time stamp

        // Unit conversions
        static public double DTOR = Math.PI / 180;
        static public double RTOD = 180 / Math.PI;
        static public double HAtoR = 12 / Math.PI;
        static public double NoonHour = 12.00;

        // Other constants (String, boolean, etc.)
        static public String timeFormat;                      // The time stamp format used in the input file
        static public bool keepWindowOpen;                    // Keeping the window open long enough
    }

    public static class Utilities
    {
        static public DateTime CurrentTimeStamp;
        static public DateTime cachedTimeStamp;

        // Gets the day of the year based on a given date
        public static void TSBreak(String TimeStamp, out int dayOfYear, out double hour, out int year, out int month, out double nextTimeStampHour, out double baseTimeStampHour, SimMeteo simMeteoParser)
        {
            try
            {
                CurrentTimeStamp = DateTime.ParseExact(TimeStamp, Util.timeFormat, null);

                // Checks ensure the time series is always progressing forward.
                if (ErrorLogger.iterationCount != 1)
                {
                    if (CurrentTimeStamp != cachedTimeStamp)
                    {
                        // Check if the time stamps are going back in time
                        if (DateTime.Compare(CurrentTimeStamp, cachedTimeStamp) < 0)
                        {
                            ErrorLogger.Log("Time stamps in the Input File go backwards in time. Please check your input file. CASSYS has ended.", ErrLevel.FATAL);

                        }
                    }
                }
                else
                {
                    // Get the next expected time stamp
                    cachedTimeStamp = CurrentTimeStamp;
                }

                // Next and Base time stamps are used to check if the sun-rise and sun-set event occurs in between the time stamps under consideration
                DateTime nextTimeStamp = DateTime.ParseExact(TimeStamp, Util.timeFormat, null);
                DateTime baseTimeStamp = DateTime.ParseExact(TimeStamp, Util.timeFormat, null);

                switch (Util.AveragedAt)
                {
                    case "Beginning":
                        baseTimeStamp = CurrentTimeStamp;
                        nextTimeStamp = baseTimeStamp.AddMinutes(Util.timeStep);
                        CurrentTimeStamp = CurrentTimeStamp.AddMinutes(Util.timeStep / 2D);

                        break;

                    case "End":
                        nextTimeStamp = CurrentTimeStamp;
                        baseTimeStamp = CurrentTimeStamp.AddMinutes(-Util.timeStep);
                        CurrentTimeStamp = CurrentTimeStamp.AddMinutes(-Util.timeStep / 2D);
                        break;

                    default:
                        baseTimeStamp = CurrentTimeStamp.AddMinutes(-Util.timeStep/2D);
                        nextTimeStamp = CurrentTimeStamp.AddMinutes(Util.timeStep/2D);
                        CurrentTimeStamp = CurrentTimeStamp.AddMinutes(0);
                        break;
                }

                dayOfYear = CurrentTimeStamp.DayOfYear;

                // Allowing for Leap Years - Assumes February 29 as Feb 28 and all other days as their day number during a normal year
                if ((CurrentTimeStamp.Month > 2) && (DateTime.IsLeapYear(CurrentTimeStamp.Year)))
                {
                    if (dayOfYear > 59)
                    {
                        dayOfYear = CurrentTimeStamp.DayOfYear - 1;
                    }
                }
                hour = CurrentTimeStamp.Hour + CurrentTimeStamp.Minute / 60D + CurrentTimeStamp.Second / 3600D;
                year = CurrentTimeStamp.Year;
                month = CurrentTimeStamp.Month;

                baseTimeStampHour = baseTimeStamp.Hour + baseTimeStamp.Minute / 60D + baseTimeStamp.Second / 3600D;
                nextTimeStampHour = nextTimeStamp.Hour + nextTimeStamp.Minute / 60D + nextTimeStamp.Second / 3600D;
            }
            catch (FormatException)
            {
                dayOfYear = 0;
                hour = 0;
                year = 0;
                month = 0;
                baseTimeStampHour = 0;
                nextTimeStampHour = 0;
                ErrorLogger.Log(TimeStamp + " was not recognized a valid DateTime. The date-time was expected in " + Util.timeFormat + " format. Please check Site definition file. Row was skipped", ErrLevel.WARNING);
                simMeteoParser.inputRead = false;
            }
        }
        
        // Converts a temperature in degrees Centigrade to degrees Kelvin
        public static double ConvertCtoK(double TempinC)
        {
            double TempinK = TempinC + 273.15;
            return TempinK;
        }

        // Converts an angle given in degrees to radians
        public static double ConvertDtoR(double AngleinD)
        {
            double AngleinR = (Math.PI / 180) * AngleinD;
            return AngleinR;
        }

        // Converts a given angle in Radians to Degrees
        public static double ConvertRtoD(double AngleinR)
        {
            double AngleinD = AngleinR / (Math.PI / 180);
            return AngleinD;
        }

        // Converts kiloWatts to Watts
        public static double ConvertkWtoW(double PowinkW)
        {
            double PowinW = PowinkW / 1000;
            return PowinW;
        }

        // Converts Watts to kiloWatts
        public static double ConvertWtokW(double PowinW)
        {
            double PowinkW = PowinW / 1000;
            return PowinkW;
        }
        
        // Truncates a scientific notation value to a specified number of integers
        public static double Truncate(double d, int places)
        {
            if (d == 0)
            {
                return 0;
            }
            else
            {
                double scale = Math.Pow(10, Math.Floor(Math.Log10(Math.Abs(d))) + 1 - places);
                return scale * Math.Truncate(d / scale);
            }
        }

        

    }
}
