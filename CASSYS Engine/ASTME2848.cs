// CASSYS - Grid connected PV system modelling software 
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: ASTM E2848 System Class
// 
// Revision History:
// NA - 2017-06-09: First release
//
// Description:
// The ASTM E2848 Class calculates the PV AC Power generated using the ASTM E2848 Standard Test.
//                             
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
// Ref 1: ASTM International, "Standard Test Method for Reporting Photovoltaic Non-Concentrator System Performance", 2011.
// http://www.astm.org/cgi-bin/resolver.cgi?E2848-13
///////////////////////////////////////////////////////////////////////////////

using System;

namespace CASSYS
{
    class ASTME2848
    {
        // Parameters of the ASTM E2848 System
        // The regression factors assume that tilted Irradiance is in W/m^2, ambient tempeature is in deg C, wind speed is in m/s
        double itsPmax;                             // Maximum power that can be produced by grid [kW]
        double itsA1;                               // Regression Factor a1
        double itsA2;                               // Regression Factor a2
        double itsA3;                               // Regression Factor a3
        double itsA4;                               // Regression Factor a4
        double[] itsEAF = new double[12];           // Holds Monthly Emperical Adjustement Factors [unitless]

        // Output variables calculated
        public double ACPower = 0;                  // AC power produced by PV system [kW]

        // Blank constructor
        public ASTME2848()
        {
        }

        // Constructor for Testing
        public ASTME2848(double Pmax, double A1, double A2, double A3, double A4, double[] EAF)
        {
            itsPmax = Pmax;
            itsA1 = A1;
            itsA2 = A2;
            itsA3 = A3;
            itsA4 = A4;
            itsEAF = EAF;
        }

        // Calculates AC power of system
        public void Calculate
            (
            SimMeteo SimMet                          // Meteological data required for ACPower calculation
            )
        {
            // Initializing to default
            ACPower = double.NaN;

            // Log errors if incorrect parameters are found
            // No check is made for radiation as it can be negative at night in some cases
            if (SimMet.MonthOfYear < 1 || SimMet.MonthOfYear > 12)
            {
                ErrorLogger.Log("ASTM E2848 Calculate: Month must be betwen 1 and 12", ErrLevel.WARNING);
            }
            else if (SimMet.TAmbient < -273.15)
            {
                ErrorLogger.Log("ASTM E2848 Calculate: Ambient temperature must be greater than 0K.", ErrLevel.WARNING);
            }
            else if (SimMet.WindSpeed < 0)
            {
                ErrorLogger.Log("ASTM E2848 Calculate: Windspeed must be greater than 0 m/s.", ErrLevel.WARNING);
            }
            else
            {
                // If inputs are valid continue with power calculation using the ASTM E2848 Equation (Ref 1)
                ACPower = Math.Min(SimMet.TGlo * (itsA1 + itsA2 * SimMet.TGlo + itsA3 * SimMet.TAmbient + itsA4 * SimMet.WindSpeed) * itsEAF[SimMet.MonthOfYear - 1], itsPmax);
                ACPower = Math.Max(ACPower, 0);   
            }

            // Assigning Outputs for this class.
            AssignOutputs();
        }


        // Config will assign parameter variables their values as obtained from the .CSYX file
        public void Config()
        {
            try
            {
                // Gathering all the parameters from ASTM E2848 Element 
                itsPmax = double.Parse(ReadFarmSettings.GetInnerText("ASTM", "SystemPmax", _Error: ErrLevel.FATAL));
                itsA1 = double.Parse(ReadFarmSettings.GetInnerText("ASTM/Coeffs", "ASTM1", _Error: ErrLevel.FATAL));
                itsA2 = double.Parse(ReadFarmSettings.GetInnerText("ASTM/Coeffs", "ASTM2", _Error: ErrLevel.FATAL));
                itsA3 = double.Parse(ReadFarmSettings.GetInnerText("ASTM/Coeffs", "ASTM3", _Error: ErrLevel.FATAL));
                itsA4 = double.Parse(ReadFarmSettings.GetInnerText("ASTM/Coeffs", "ASTM4", _Error: ErrLevel.FATAL));

                // Looping through and assigning all EAF parameters from EAF Element
                string[] months = new string[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
                for (int i = 0; i < 12; i++)
                {
                    itsEAF[i] = double.Parse(ReadFarmSettings.GetInnerText("ASTM/EAF", months[i], _Error: ErrLevel.FATAL));
                }
            }
            catch (Exception e)
            {
                ErrorLogger.Log("ASTM E2848 Config: " + e.Message, ErrLevel.FATAL);
            }
        }

        // Assigns output parameters relating to the ASTM E2848 class
        public void AssignOutputs()
        {
            ReadFarmSettings.Outputlist["Power_Injected_into_Grid"] = ACPower;
            ReadFarmSettings.Outputlist["Energy_Injected_into_Grid"] = ACPower * Util.timeStep / 60;
        }
    }
}
