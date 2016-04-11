// CASSYS - Grid connected PV system modelling software 
// Version 0.9  
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Astro class
// 
// Revision History:
// DT - 2014-10-19: Version 0.9
//
// Description: 
// The Astro class contains a set of methods to calculate sun position, 
// air mass, day length, etc.
//                              
///////////////////////////////////////////////////////////////////////////////
// 
//                              
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
//
// Braun JE and Mitchell JC (1983) Solar Geometry for fixed and
//     Tracking Surfaces. Solar Energy 31,5, 439-444.
// Duffie JA and Beckman WA (1991) Solar Engineering of Thermal
//     Processes, Second Edition. John Wiley & Sons.
// Iqbal M (1993) An introduction to solar radiation. Elsevier.
// Kasten F (1966) A New Table and Approximation Formula
//				for the Relative Optical Air Mass. Arch. Meteorol.
//				Geophys. Bioklimataol., B14, 206-233.
// McQuiston FC and Parker JD (1982) Heating, Ventilating, and Air
//     Conditioning. John Wiley & Sons.
// 
///////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace CASSYS
{
    public static class Astro
    {

        // Private static members: constants used in the code. 
        private static double DTOR = Math.PI / 180;                      // Degree to Radians conversion
        private static double RTOD = 180 / Math.PI;                      // Radians to Degree conversion
        private static double SOLAR_CONST = 1367;                        // Solar constant [W/m2]
        private static double MEAN_EARTH_RADIUS = 6371000;               // Mean earth radius [m]
        private static double AUTOm = 1.49598e+11;                       // Conversion from Astronomical Unit (AU) to m

        // Private function.  Day angle [radians] eqn 1.2.2 Iqbal (1993)
        private static double _GetDayAngle
            ( 
            int DayOfYear                           // Day of year [1, 365]
            )
        {
            return 2 * Math.PI * (DayOfYear - 1) / 365D;
        }

        // Daily astronomical quantities
        // Declination Eqn 1.3.1 from Iqbal (1993)
        public static double GetDeclination         // (o) declination [radians]
            ( 
            int DayOfYear                           // (i) day of year [1,365] 
            )
        {
            double Declination, DA;

            if (DayOfYear < 0 || DayOfYear > 365)
            {
                throw new System.ArgumentException("GetDeclination: Invalid day of year for declination calculation");
            }

            DA = _GetDayAngle(DayOfYear);             // Day angle 
            Declination = (0.006918 - 0.399912 * Math.Cos(DA) + 0.070257 * Math.Sin(DA)
                - 0.006758 * Math.Cos(2 * DA) + 0.000907 * Math.Sin(2 * DA)
                - 0.002697 * Math.Cos(3 * DA) + 0.00148 * Math.Sin(3 * DA));
            return Declination;
        }

        // Hourly astronomical quantities
        // Hour angle [radians] as a function of time of day
        // 15 degrees per hour, morning negative, afternoon positive
        public static double GetHourAngle           // (o) hour angle [radians] 
            ( 
            double Time                             // (i) local apparent time [decimal hour]    
            )
        {
          return DTOR*15*(Time-12);
        }

        // Time as a function of hour angle in radians
        // 15 degrees per hour, morning negative, afternoon positive
        static double GetTimeHA                      // (o) Local Apparent time [decimal hour] 
            ( 
            double HourAngle                         // (i) hour angle [radians] 
            )
        {
            return HourAngle/(DTOR*15) + 12;
        }

        // Position of sun: zenith and azimuth as a function of day and hour
        // Computes declination and hour angle and calls second form of same function
        public static void CalcSunPosition
            ( int DayOfYear                           // (i) day of year (Julian day) [1,365] 
            , double Hour                             // (i) local apparent time [decimal hour] 
            , double Lat                              // (i) latitude [radians]   
            , out double Zenith                       // (o) zenith angle of sun [radians]  
            , out double Azimuth                      // (o) azimuth angle of sun [radians] 
            )
        {                                   
          // Get declination and hour angle 
          double Decl, HourAngle;
  
            if (DayOfYear < 0 || DayOfYear > 365)
            {
                throw new System.ArgumentException("CalcSunPosition: Invalid day of year in calculation of sun position");
            }
            if (Lat < -Math.PI/2 || Lat > Math.PI/2)
            {
                throw new System.ArgumentException("CalcSunPosition: Invalid latitude in calculation of sun position");
            }
  
          Decl = GetDeclination(DayOfYear);
          HourAngle = GetHourAngle(Hour);
  
          // Call second form 
          CalcSunPositionHourAngle(Decl, HourAngle, Lat, out Zenith, out Azimuth);
        }

        // Position of sun: zenith and azimuth as a function of declination and hour
        // angle
        // Zenith angle: Duffie and Beckman (1991), Eq. 1.6.5
        // For sinAzimuth, see Braun and Mitchell (1983) formula (4)
        // For cosAzimuth, see McQuiston and Parker (1982), formulae (5.4) and (5.1)
        // Note: in Braun and Mitchell, the convention for solar azimuth angle is 0 
        // points towards the equator and positive is to the west. For a point 
        // located at the equator 0 points south. In the northern hemisphere that
        // convention is equivalent to that in CASSYS. In the southern hemisphere 
        // Azimuth_CASSYS = PI - Azimuth_BM (modulo 2*PI), therefore sinAzimuth is the
        // same but cosAzimuth changes sign

        public static void CalcSunPositionHourAngle
            ( double Decl                             // (i) declination [radians] 
            , double HourAngle                        // (i) hour angle [radians] 
            , double Lat                              // (i) latitude [radians] 
            , out double Zenith                       // (o) zenith angle of sun [radians] 
            , out double Azimuth                      // (o) azimuth angle of sun [radians] 
            )
        {
          // Declarations 
          double cosZenith;                           // Cosine of the Zenith Angle [radians]
          double sinAzimuth;                          // Sine of the Azimuth Angle [radians]
          double cosAzimuth;                          // Cosine of the Azimuth Angle [radians]
  
            if (Lat < -Math.PI/2 || Lat > Math.PI/2)
            {
                throw new System.ArgumentException("CalcSunPositionHourAngle: Invalid latitude in calculation of sun position");
            }
  
          // Compute zenith angle 
          cosZenith = Math.Cos(Lat)*Math.Cos(Decl)*Math.Cos(HourAngle)+Math.Sin(Lat)*Math.Sin(Decl);
          cosZenith = Math.Min(cosZenith,  1.0);
          cosZenith = Math.Max(cosZenith, -1.0);
          Zenith = Math.Acos(cosZenith);
  
          // Compute azimuth angle 
          sinAzimuth = Math.Sin(HourAngle)*Math.Cos(Decl);
          cosAzimuth = Math.Cos(HourAngle)*Math.Cos(Decl)*Math.Sin(Lat)-Math.Sin(Decl)*Math.Cos(Lat);
          Azimuth = Math.Atan2(sinAzimuth, cosAzimuth);
        }

        // Distance Calculations
        // Calculate the Sun - Earth Distance.  Iqbal (1983) eqn. 1.2.1.
        public static double GetEccentricityCorrFactor  // (o) eccentricity correction factor [] 
            ( 
            int DayOfYear                             // (i) day of year [1,365] 
            )
        {
            double DA;

            if (DayOfYear < 0 || DayOfYear > 365)
            {
                throw new System.ArgumentException("GetEccentricityCorrFactor: Invalid day of year in calculation of eccentricity correction factor");
            }


            DA = _GetDayAngle(DayOfYear);

            return (1.000110 + 0.034221 * Math.Cos(DA)
                   + 0.00128 * Math.Sin(DA) + 0.000719 * Math.Cos(2 * DA)
                   + 0.000077 * Math.Sin(2 * DA));
        }

        // Calculate the Sun - Earth Distance.  Duffie and Beckman (1991)
        public static double GetSunEarthDistance      // (o) distance from sun to earth [m] 
            (
            int Day                                  // (i) day of year (Julian day) [1,365] 
            )
        {
            return AUTOm / Math.Sqrt(GetEccentricityCorrFactor(Day));
        }

        // Functions Dealing With Time
        //---------------------------------------------------------------------------
        // Standard time correction
        // GetATmsST = Apparent Time MinuS Standard Time, in minutes
        // Duffie & Beckman (1991), eq. 1.5.3a and 1.5.3b

        public static double GetATmsST                // (o) Apparent (solar) Time Minus Standard Time [min]  
            ( 
            int DayOfYear                           // (i) day of year [1, 365] 
            , double SLong                          // (i) site longitude [radians, E > 0]  
            , double MLong                          // (i) meridian longitude [radians, E > 0] 
            )
        {
            // 
            double b;                               // Day angle [rad]
            double e;                               // Please see model for complete description.
            double DLong;                           // Difference in site longitude and meridian longitude [rad]
  
            if (DayOfYear < 0 || DayOfYear > 365)
            {
                throw new System.ArgumentException("GetATmsST: Invalid day of year in calculation of apparent minus standard time");
            }
  
            b = _GetDayAngle(DayOfYear);
            e = 229.18 * (0.000075 + 0.001868*Math.Cos(b) - 0.032077*Math.Sin(b)
                   - 0.014615*Math.Cos(2*b) - 0.04089*Math.Sin(2*b));
            DLong = (SLong - MLong)*RTOD;
            if (DLong >  180) DLong = DLong - 360;
            if (DLong < -180) DLong = DLong + 360;

            return 4*DLong+e;
        }

        ///////////////////////////////////////////////////////////////////////////////
        // Time conversions from apparent to standard time and vice-versa
        // Times are expressed in hours, ATmsST is expressed in minutes

        public static double GetLATime              // (o) Local Apparent (solar) Time [decimal hour] 
            ( 
            double LSTime                           // (i) Local Standard Time [decimal hour] 
            , double ATmsST                         // (i) Apparent Time MinuS Standard Time [minutes] 
            )
        {
            return LSTime + ATmsST / 60;
        }

        public static double GetLSTime                // (o) Local Standard Time [decimal hour] 
            ( double LATime                           // (i) Local Apparent Time [decimal hour] 
            , double ATmsST                           // (i) Apparent Time MinuS Standard Time [minutes] 
            )
        {
            return LATime - ATmsST / 60;
        }

        // Extraterrestrial Irradiance and air mass 
        // Normal extraterrestrial irradiance [W/m2]
        // Duffie and Beckman (1991). eq. 1.4.1          
        public static double GetNormExtra            // (o) normal extraterrestrial irradiance [W/m2] 
            (
            int DayOfYear                            // (i) day of year [1, 365] 
            )
        {
            if (DayOfYear < 0 || DayOfYear > 365)
            {
                throw new System.ArgumentException("GetNormExtra: Invalid day of year in calculation of extraterrestrial irradiance");
            }

            return SOLAR_CONST * GetEccentricityCorrFactor(DayOfYear);
        }

        // Compute the air mass for a given zenith angle. See Kasten (1966).
        public static double GetAirMass              // (o) air mass [#] 
            (
            double Zenith                            // (i) zenith angle of sun [radians] 
            )
        {
            if (Zenith < 0 || Zenith > Math.PI)
            {
                throw new System.ArgumentException("GetAirMass: Invalid zenith value.");
            }

            if (Zenith < Math.PI / 2)
                return Math.Max(1 / (Math.Cos(Zenith) + 0.15 * Math.Pow(93.885 - RTOD * Zenith, -1.253)), 1);
            else
                return 1 / (0.15 * Math.Pow(3.885, -1.253));
        }

        // Sunset and day length, Sunset hour angle [radians]. See Duffie and Beckman (1991) eq. 1.6.10

        public static double GetSunsetHourAngle       // (o) sunset hour angle [radians] 
            ( double Lat                              // (i) latitude [radians, N > 0]  
            , double Decl                             // (i) declination [radians]   
            )
        { 
          double Sunset, cosSunset;
          try
          {
              if (Lat < -Math.PI / 2 || Lat > Math.PI / 2)
              {
                  throw new System.ArgumentException("GetSunsetHourAngle: Invalid latitude");
              }
          }
          catch (ArgumentException ae)
          {
              ErrorLogger.Log(ae, ErrLevel.FATAL);
          }
  
          cosSunset = -Math.Tan(Lat)*Math.Tan(Decl);
          if (cosSunset > 1)
            Sunset = 0;
          else if (cosSunset < -1)
            Sunset = Math.PI;
          else
            Sunset = Math.Acos(cosSunset);
  
          return Sunset;
        }

        // Compute day length in hours
        public static double GetDayLength             // (o) length of day [decimal hours] 
            ( double Lat                              // (i) latitude [radians, N > 0]  
            , double Decl                             // (i) declination [radians] 
            )
        {
            double Sunset;
            Sunset = GetSunsetHourAngle(Lat, Decl);
            return Sunset * RTOD * 2D / 15D;
        }

        // Compute distance between two points, in m, knowing their latitude and longitude, in radians
        // Meeus (1991), formulae 10.1 and 16.2
        public static double GetDistance              // (o) distance between two points [m] 
            ( double Lat1                             // (i) latitude of point 1 [radians, N > 0]  
            , double Long1                            // (i) longitude of point 1 [radians, E > 0]  
            , double Lat2                             // (i) latitude of point 2 [radians, N > 0]  
            , double Long2                            // (i) longitude of point 2 [radians, E > 0]  
            )
        { 
            double cosangle;
            double angle;
            double aux1;
            double aux2;
          
            // Calculate cosine of angular distance between the two points 
            cosangle = Math.Sin(Lat1)*Math.Sin(Lat2)+Math.Cos(Lat1)*Math.Cos(Lat2)*Math.Cos(Long1-Long2);
          
            // If angular distance less than 0.2 degrees, use approximation 16.2
            if (cosangle > Util.COS_LITTLE)
            {                          
                aux1 = (Long1-Long2)*Math.Cos((Lat1+Lat2)/2);
                aux2 = Lat1-Lat2;
                angle = Math.Sqrt(aux1*aux1+aux2*aux2);
            }
            else
            {
                angle = Math.Acos(cosangle);
            }
            return angle*MEAN_EARTH_RADIUS; 
        }
    }
}

