// CASSYS - Grid connected PV system modelling software  
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Tilt class
// 
// Revision History:
// DT - 2014-10-20: Version 0.9
//
// Description: 
// The Tilt class contains a set of methods to calculate solar radiation on 
// tilted surfaces, as well as the object to encapsulate them. 
//                              
///////////////////////////////////////////////////////////////////////////////
// 
//                              
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
//
// Ref 1: Duffie JA and Beckman WA (1991) Solar Engineering of Thermal
//     Processes, Second Edition. John Wiley & Sons.
// 
// Ref 2: (Website, Accessed 2014-10) PV Modelling Collaborative: 
//     http://pvpmc.org/modeling-steps/shading-soiling-and-reflection-losses/incident-angle-reflection-losses/ashre-model/
///////////////////////////////////////////////////////////////////////////////

using System;

namespace CASSYS
{
    public static class Tilt
    {
        // Incidence angle on a tilted surface
        // The first form returns the cosine of the angle, the second form calls the
        // first form and returns the angle itself [radians]
        // Duffie and Beckman (1991) eqn. 1.6.3

       // Output variables
       public static double IncidenceAngle;           // Angle of incidence for the beam 

       public static double GetCosIncidenceAngle      // (o) cos of angle of incidence of beam radiation on surface [#] 
            ( double SunZenith                        // (i) zenith angle of sun [radians] 
            , double SunAzimuth                       // (i) azimuth angle of sun [radians] 
            , double SurfaceSlope                     // (i) slope of surface [radians] 
            , double SurfaceAzimuth                   // (i) azimuth of surface [radians] 
            )
        {
            double cosInc;                            // cos of angle of incidence of beam radiation on surface [#]
           
            if (SunZenith < 0 || SunZenith > Math.PI)
            {
                throw new System.ArgumentException("GetCosIncidenceAngle: Invalid sun zenith.");
            }
            if (SunAzimuth < -Math.PI || SunAzimuth > Math.PI)
            {
                throw new System.ArgumentException("GetCosIncidenceAngle: Invalid sun azimuth.");
            }
            //if (SurfaceSlope * Util.DTOR < 0 || SurfaceSlope * Util.DTOR > Math.PI)
            if (SurfaceSlope < 0 || SurfaceSlope > Math.PI)
            {
                throw new System.ArgumentException("GetCosIncidenceAngle: Invalid surface slope.");
            }
            // Allow SurfaceAzimuth < -Math.PI and SurfaceAzimuth > Math.PI for bifacial modelling

            // To prevent Incidence angle from being returned if sun is below horizon
            if (SunZenith > Math.PI / 2)
            {
                return 0;
            }

            // Nominal case
            cosInc = Math.Cos(SunZenith) * Math.Cos(SurfaceSlope)
                   + Math.Sin(SunZenith) * Math.Sin(SurfaceSlope) * Math.Cos(SunAzimuth - SurfaceAzimuth);
            cosInc = Math.Min(cosInc, 1.0);
            cosInc = Math.Max(cosInc, -1.0);

            return cosInc;
        }

        public static double GetIncidenceAngle        // (o) incidence angle of beam radiation on surface [radians] 
            ( double SunZenith                        // (i) zenith angle of sun [radians] 
            , double SunAzimuth                       // (i) azimuth angle of sun [radians] 
            , double SurfaceSlope                     // (i) slope of surface [radians] 
            , double SurfaceAzimuth                   // (i) azimuth of surface [radians] 
            )
        {
            // Calculating the IncidenceAngle using GetCosIncidenceAngle Method
            return Math.Acos(GetCosIncidenceAngle(SunZenith, SunAzimuth,
                SurfaceSlope, SurfaceAzimuth));
        }

        // Calculates the Incidence Angle Modifier using ASHRAE Parameter, see Ref. 2
        public static double GetASHRAEIAM
            (
              double Bo                                     // ASHRAE Parameter [#]
            , double InciAng                                // Incidence Angle [radians]
            )
        {
            return Math.Cos(InciAng) > Bo / (1 + Bo) ? Math.Max((1 - Bo * (1 / Math.Cos(InciAng) - 1)), 0) : 0;
        }

        // Compute profile angle.
        // Duffie and Beckman (1991) eqn 1.6.12
        public static double GetProfileAngle          // (o) profile angle for a surface [radians] 
            ( double SunZenith                        // (i) zenith angle of sun [radians] 
            , double SunAzimuth                       // (i) azimuth angle of sun [radians] 
            , double SurfaceAzimuth                   // (i) azimuth of surface [radians] 
            )
        {
            // Catching any errors
            // Allow SurfaceAzimuth < -Math.PI and SurfaceAzimuth > Math.PI for bifacial modelling
            if (SunAzimuth < -Math.PI || SunAzimuth > Math.PI)
            {
                throw new CASSYSException("GetProfileAngle: Invalid sun azimuth.");
            }

            if (SunZenith == 0)
            {
                return Math.PI / 2;
            }
            else if (Math.Abs(SunAzimuth - SurfaceAzimuth) == Math.PI / 2)
            {
                return Math.PI / 2;
            }
            else
            {
                return Math.Atan(Math.Tan(Math.PI / 2 - SunZenith) / Math.Cos(SunAzimuth - SurfaceAzimuth));
            }
        }

        // Compute apparent sunset on a titled surface facing the equator
        // Duffie and Beckman (1991) eqn 2.19.3b
        public static double GetApparentSunsetHourAngleEquator      // (o) apparent sunset hour angle [radians] 
            ( double Slope                            // (i) slope of receiving surface [radians] 
            , double Lat                              // (i) latitude [radians, N > 0] 
            , double Decl                             // (i) declination [radians] 
            )
        {
            double IsNorth, Aux1, Aux2;

            IsNorth = (Lat > 0) ? 1 : -1;
            Aux1 = Astro.GetSunsetHourAngle(Lat, Decl);
            Aux2 = Astro.GetSunsetHourAngle(Lat-IsNorth*Slope, Decl);
            return Math.Min(Aux1, Aux2);
        }

        // Compute apparent sunrise and sunset times on a tilted surface of any orientation
        // Duffie and Beckman (1991) eqn 2.20.5e to 2.20.5i
        public static void CalcApparentSunsetHourAngle
            ( double Lat                              // (i) latitude [radians] 
            , double Decl                             // (i) declination [radians] 
            , double Slope                            // (i) slope of receiving surface [radians] 
            , double Azimuth                          // (i) azimuth angle of receiving surface [radians] 
            , out double AppSunrise                   // (o) apparent sunrise hour angle [radians] 
            , out double AppSunset                    // (o) apparent sunset hour angle [radians] 
            , out double Sunset                       // (o) true sunset hour angle [radians] 
            )
        {
            // Declarations 
            double aux1, aux2, aux3, disc;
            double cosInc, SunZenith, SunAzimuth;
            double A, B, C;                        

            // Calculate true sunset and parameters A, B and C 
            Sunset = Astro.GetSunsetHourAngle(Lat, Decl);
            A = Math.Cos(Slope)+Math.Tan(Lat)*Math.Cos(Azimuth)*Math.Sin(Slope);
            B = Math.Cos(Sunset)*Math.Cos(Slope)+Math.Tan(Decl)*Math.Sin(Slope)*Math.Cos(Azimuth);
            C = Math.Sin(Slope)*Math.Sin(Azimuth)/Math.Cos(Lat);

            // Pathological case: the sun does not rise 
            if (Sunset == 0)
            {
                AppSunrise = AppSunset = 0;
                return;
            }

            // Normal case 
            aux1 = A*A + C*C;
            disc = aux1 - B*B;

            // Check discriminant. If it is positive or zero, the sun rises (tangentially)
            // on the surface. In that case, calculate apparent sunrise and apparent
            // sunset 
            if (disc >= 0)
            {
                aux2 = C*Math.Sqrt(disc)/aux1;
                aux3 = A*B/aux1;
                AppSunrise = Math.Min(Sunset, Math.Acos(aux3+aux2));
                AppSunset  = Math.Min(Sunset, Math.Acos(aux3-aux2));
                if ((A > 0 && B > 0) || (A >= B))
                    AppSunrise = -AppSunrise;
                else
                    AppSunset = -AppSunset;
            }

            // If discriminant is negative, the surface is either always or never
            // illuminated during the day. To find which case applies, compute the
            // cosine of the angle of incidence on the surface at noon. If it is less
            // than zero, then the surface is always in the dark. If it is greater
            // than zero, then the surface is always illuminated 
            else
            {
                Astro.CalcSunPositionHourAngle(Decl, 0, Lat, out SunZenith, out SunAzimuth);
                cosInc = GetCosIncidenceAngle(SunZenith, SunAzimuth, Slope, Azimuth);
                if (cosInc < 0)
                    AppSunrise = AppSunset = 0;
                else
                {
                    AppSunrise = -Sunset;
                    AppSunset  =  Sunset;
                }
            }
        }
    }
}
