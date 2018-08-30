// CASSYS - Grid connected PV system modelling software
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Sun Class
// 
// Revision History:
// DT - 2014-10-19: Version 0.9
//
// Description: 
// The Sun class is an object used to compute solar zenith and azimuth angles.
//                              
///////////////////////////////////////////////////////////////////////////////
// 
//                              
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
//
// Ref 1: Duffie JA and Beckman WA (1991) Solar Engineering of Thermal
//     Processes, Second Edition. John Wiley & Sons.
//
///////////////////////////////////////////////////////////////////////////////
//
// 
///////////////////////////////////////////////////////////////////////////////

using System;

namespace CASSYS
{
    class Sun
    {
        // Sun definition variables
        public double itsSLat;                 // Latitude of site [°N]                      
        public double itsSLong;                // Longitude of site [°E]
        public double itsMLong;                // Longitude of reference meridian [°E] - required when standard time is used
        public bool itsLSTFlag;                // Flag indicating whether calculations are in Local Standard Time (false = local solar time)
        public double itsSurfaceSlope;         // Slope of the collector

        // Current day calculation, cached to speed calculation up
        int itsCurrentDayOfYear = 0;           // Current day of year held in memory
        double itsCurrentDecl;                 // Declination for current day
        double itsCurrentNExtra;               // Current extraterrestrial radiation

        // Output variables calculated
        public double Zenith;                  // Zenith angle of sun [rad]
        public double Azimuth;                 // Azimuth of sun [rad, >0 facing E]
        public double AirMass;                 // Air mass [#]
        public double NExtra;                  // Extraterrestrial normal Irradiance [W/m2]
        public double AppSunsetHour;           // Sunset hour (uses Tilt to calculate the Sunset Hour)
        public double AppSunriseHour;          // Sunrise hour (uses Tilt to calculate the Sunset Hour)
        public double TrueSunSetHour;          // The true Sunset hour
        public double TrueSunRiseHour;         // The true sunrise hour

        // Blank Sun constructor
        public Sun
            (
            )
        {
        }

        // Sun constructor with parameters
        public Sun
            (
              double SLat                         // Latitude of site [°N]                      
            , double SLong                        // Longitude of site [°E]
            , double MLong                        // Longitude of reference meridian [°E] - required when standard time is used
            , bool LSTFlag                        // Flag indicating whether calculations are in Local Standard Time (false = local solar time)
            )
        {
            this.itsSLat = SLat;
            this.itsSLong = SLong;
            this.itsMLong = MLong;
            this.itsLSTFlag = LSTFlag;
        }

        // Calculation for inverter output power, using efficiency curve
        public void Calculate
            (
              int DayOfYear                           // Day of year (1-365)
            , double Hour                             // Hour of day, in decimal format (11.75 = 11:45 a.m.)
            )
        {
            double itsSLatR = Utilities.ConvertDtoR(itsSLat);
            double itsSLongR = Utilities.ConvertDtoR(itsSLong);
            double itsMLongR = Utilities.ConvertDtoR(itsMLong);

            try
            {
                if (DayOfYear < 1 || DayOfYear > 365 || Hour < 0 || Hour > 24)
                {

                    throw new CASSYSException("Sun.Calculate: Invalid time stamp for sun position calculation");
                }
            }
            catch (CASSYSException cs)
            {
                ErrorLogger.Log(cs, ErrLevel.FATAL);
            }

            // Compute declination and normal extraterrestrial Irradiance if day has changed
            // Compute Sunrise and Sunset hour angles
            if (DayOfYear != itsCurrentDayOfYear)
            {
                itsCurrentDayOfYear = DayOfYear;
                itsCurrentDecl = Astro.GetDeclination(itsCurrentDayOfYear);
                itsCurrentNExtra = Astro.GetNormExtra(itsCurrentDayOfYear);

                // Variables used to hold the apparent/true sunset and sunrise hour angles    
                double appSunRiseHA;                        // Hour angle for Sunrise [radians]
                double appSunsetHA;                         // Hour angle for Sunset  [radians]
                double trueSunsetHA;                        // True Sunset Hour angle [radians]

                // Invoking the Tilt method to get the values
                Tilt.CalcApparentSunsetHourAngle(itsSLatR, itsCurrentDecl, itsSurfaceSlope, Azimuth, out appSunRiseHA, out appSunsetHA, out trueSunsetHA);

                // Assigning to the output values
                AppSunriseHour = Math.Abs(appSunRiseHA) * Util.HAtoR;
                AppSunsetHour = Util.NoonHour + appSunsetHA * Util.HAtoR;
                TrueSunSetHour = Util.NoonHour + trueSunsetHA * Util.HAtoR ; 
                TrueSunRiseHour = TrueSunSetHour - Astro.GetDayLength(itsSLatR, itsCurrentDecl);

                // If using local standard time then modify the sunrise and sunset to match the local time stamp.
                if (itsLSTFlag)
                {
                    TrueSunSetHour -= Astro.GetATmsST(DayOfYear, itsSLongR, itsMLongR) / 60; // Going from solar to local time
                    TrueSunRiseHour = TrueSunSetHour - Astro.GetDayLength(itsSLatR, itsCurrentDecl);
                }
            }

            // Compute hour angle
            double SolarTime = Hour;
            if (itsLSTFlag)
            {
                SolarTime += Astro.GetATmsST(DayOfYear, itsSLongR, itsMLongR) / 60; // Going from local to solar time
            }

            double HourAngle = Astro.GetHourAngle(SolarTime);

            // Compute azimuth and zenith angles
            Astro.CalcSunPositionHourAngle(itsCurrentDecl, HourAngle, itsSLatR, out Zenith, out Azimuth);

            // Compute normal extraterrestrial Irradiance
            NExtra = itsCurrentNExtra;

            // Compute air mass
            AirMass = Astro.GetAirMass(Zenith);
        }

        // Config will assign parameter variables their values as obtained from the .CSYX file
        public void Config()
        {
            // Gathering the parameters for the Sun Class
            //itsSurfaceSlope = Util.DTOR * double.Parse(ReadFarmSettings.GetInnerText("O&S", "PlaneTilt")); TODO: Re-evaluate.
            itsSLat = double.Parse(ReadFarmSettings.GetInnerText("SiteDef", "Latitude"));
            itsSLong = double.Parse(ReadFarmSettings.GetInnerText("SiteDef", "Longitude"));
            itsLSTFlag = Convert.ToBoolean(ReadFarmSettings.GetInnerText("SiteDef", "UseLocTime"));

            // If Local Standard Time is to be used, get the reference meridian for the "Standard Time" of the region
            if (itsLSTFlag)
            {
                itsMLong = double.Parse(ReadFarmSettings.GetInnerText("SiteDef", "RefMer"));
            }
        }

        // Compute the clearness index. Duffie and Beckman (1991)
        public static double GetClearnessIndex      // (o) clearness index [0-1]
            (
              double HGlo                           // (i) global irradiance on horizontal [W/m2]
            , double NExtra                         // (i) normal extraterrestrial irradiance [W/m2]
            , double Zenith                         // (i) zenith angle of sun [radians]
            )
        {
            double kt;                              // Clearness index defined locally to be returned later
            if (Zenith >= Math.PI / 2)
                kt = 1.0;
            else
                kt = HGlo / (NExtra * Math.Cos(Zenith));

            kt = Math.Min(kt, 1.0);
            kt = Math.Max(kt, 0.0);
            return kt;
        }

        // Compute the diffuse fraction given the clearness index, using the Orgill and Hollands formula
        // Duffie, J.A., and Beckman, W.A., Solar Engineering of Thermal
        // Processes, 2nd edition, John Wiley & Sons (1991), p. 81
        public static double GetDiffuseFraction     // (o) diffuse fraction (Orgill and Hollands formula) [0-1]
            (
                double kt                           // (i) clearness index [0-1]
            )
        {
            double kd;                              // Diffuse fraction defined locally to be returned later
            if (kt < 0.35)
                kd = 1.0 - 0.249 * kt;
            else if (kt < 0.75)
                kd = 1.557 - 1.84 * kt;
            else
                kd = 0.177;
            return kd;
        }
    }
}
