// CASSYS - Grid connected PV system modelling software 
// Version 0.9  
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Splitter Class
// 
// Revision History:
// DJT - 2014-11-10: Version 0.9
//
// Description:
// This class calculates beam and diffuse irradiance, given global irradiance
//                             
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Text;
using System.IO;
using CASSYS;

namespace CASSYS
{
    class Splitter
    {

        // Private static members: constants used in the code. 
        private static double DTOR = Math.PI / 180; // Degree to Radians conversion

        // Output variables
        public double HGlo;                         // global horizontal irradiance [W/m2]
        public double HDif;                         // diffuse horizontal irradiance [W/m2]
        public double NDir;                         // direct normal irradiance [W/m2]
        public double HDir;                         // direct horizontal irradiance [W/m2]

        // Splitter constructor
        public Splitter
            (
            )
        {
        }

        // Calculation method
        // There are 6 ways to call the Calculate method; for each, only some of the parameters are required to calculate global, beam and diffuse. The inputs can be
        // 1. global horizontal
        // 2. global horizontal and diffuse horizontal
        // 3. global horizontal and direct horizontal
        // 4. global horizontal and direct normal
        // 5. diffuse horizontal and direct horizontal
        // 6. diffuse horizontal and direct normal
        public void Calculate
            (
              double Zenith                 // zenith angle of sun [radians]
            , double _HGlo = double.NaN     // global horizontal irradiance [W/m2]
            , double _HDif = double.NaN     // diffuse horizontal irradiance [W/m2]
            , double _NDir = double.NaN     // direct normal irradiance [W/m2]
            , double _HDir = double.NaN     // direct horizontal irradiance [W/m2]
            , double NExtra = double.NaN    // normal extraterrestrial irradiance [W/m2]
            )
        {
            // Set all values to an invalid 
            // Check that at least some of the radiation inputs are properly defined
            try
            {
                // Either HGlo has to be defined, or HDif and one of the two direct components
                if (double.IsNaN(_HGlo) && (double.IsNaN(_HDif) || (double.IsNaN(HDir) && double.IsNaN(NDir))))
                    throw new CASSYSException("Splitter: insufficient number of inputs defined.");

                // If global is defined, cannot have both diffuse and direct defined
                if (!double.IsNaN(_HGlo) && !double.IsNaN(_HDif) && ((!double.IsNaN(_HDir)) || !double.IsNaN(_NDir)))
                    throw new CASSYSException("Splitter: cannot specify both diffuse and direct when global is used.");

                // Cannot have direct normal and direct horizontal defined
                if (!double.IsNaN(_HDir) && !double.IsNaN(_NDir))
                    throw new CASSYSException("Splitter: cannot specify both direct normal and direct horizontal.");

                // If global is defined but neither diffuse and direct are defined,
                // then additional inputs are required to calculate them
                if (!double.IsNaN(_HGlo) && !double.IsNaN(_HDif) && !double.IsNaN(_HDir) && !double.IsNaN(_NDir))
                    if (NExtra == double.NaN)
                        throw new CASSYSException("Splitter: Extraterrestrial irradiance is required when only global horizontal is specified.");
            }
            catch (CASSYSException cs)
            {
                ErrorLogger.Log(cs, ErrLevel.FATAL);
            }

            // Calculate cos of zenith angle
            double cosZ = Math.Cos(Zenith);

            // First case: only global horizontal is defined
            // Calculate diffuse using the Hollands and Orgill correlation
            try
            {
                // If only HGlo is specified this case is used, first check the values provided in the program:
                if (!double.IsNaN(_HGlo) && double.IsNaN(_HDif) && double.IsNaN(_NDir) && double.IsNaN(_HDir))
                {
                    // Initialize value of HGlo
                    HGlo = _HGlo;

                    // If sun below horizon, direct is zero and diffuse is global
                    // Changed this to sun below 87.5° as high zenith angles sometimes caused problems of
                    // high direct on tilted surfaces
                    if (Zenith > 87.5 * DTOR || HGlo <= 0)
                    {
                        HDif = HGlo;
                        HDir = NDir = 0;
                    }

                    // Compute diffuse fraction
                    else
                    {
                        double kt = HGlo / NExtra / cosZ;
                        double kd;
                        kd = GetDiffuseFraction(kt);
                        kd = Math.Min(kd, 1.0);
                        kd = Math.Max(kd, 0.0);

                        // Compute diffuse and direct on horizontal
                        HDif = HGlo * kd;
                        HDir = HGlo - HDif;
                        NDir = HDir / cosZ;
                    }

                    // Limit beam normal to clear sky value
                    // ASHRAE clear sky model (ASHRAE Handbook - Fundamentals, 2013, ch. 14) with tau_b = 0.245, taud = 2.611, a_b = 0.668 and a_d = 0.227
                    // (these values are from Flagstaff, AZ, for the month of June, and lead to one of the highest beam/extraterrestrial ratios worldwide)
                    double AirMass = Astro.GetAirMass(Zenith);
                    double NDir_cs = NExtra * Math.Exp(-0.245 * Math.Pow(AirMass, 0.668));
                    NDir = Math.Min(NDir, NDir_cs);
                    HDir = NDir * cosZ;
                    HDif = HGlo - HDir;

                }

                // Second case: global horizontal and diffuse horizontal are defined then this
                else if (_HGlo != double.NaN && _HDif != double.NaN)
                {
                    // If sun below horizon, direct is zero and diffuse is global
                    // Changed this to sun below 87.5° as high zenith angles sometimes caused problems of
                    // high direct on tilted surfaces
                    if (Zenith > 87.5 * DTOR || _HGlo <= 0)
                    {
                        HGlo = _HGlo;
                        HDif = _HGlo;
                        HDir = NDir = 0;
                    }
                    else
                    {
                        HGlo = _HGlo;
                        HDif = Math.Min(_HGlo, _HDif);
                        HDir = HGlo - HDif;
                        NDir = HDir / cosZ;
                    }
                }

                // Third case: global horizontal and direct horizontal are defined
                else if (_HGlo != double.NaN && _HDir != double.NaN)
                {
                    HGlo = _HGlo;
                    HDir = Math.Min(_HGlo, _HDir);
                    HDif = HGlo - HDir;
                    NDir = HDir / cosZ;
                }

                // Fourth case: global horizontal and direct normal are defined
                else if (_HGlo != double.NaN && _NDir != double.NaN)
                {
                    HGlo = _HGlo;
                    HDir = Math.Min(HGlo, NDir * cosZ);
                    HDif = HGlo - HDir;
                    NDir = HDir / cosZ;
                }

                // Fifth case: diffuse horizontal and direct horizontal are defined
                else if (_HDif != double.NaN && _HDir != double.NaN)
                {
                    HDif = _HDif;
                    HDir = _HDir;
                    HGlo = HDif + HDir;
                    NDir = HDir / cosZ;
                }

                // Sixth case: diffuse horizontal and direct normal are defined
                else if (_HDif != double.NaN && _NDir != double.NaN)
                {
                    HDif = _HDif;
                    NDir = _NDir;
                    HDir = NDir * cosZ;
                    HGlo = HDif + HDir;
                }

                // Other cases: should never get there
                else
                {
                    throw new CASSYSException("Splitter: unexpected case encountered.");
                }
            }
            catch (CASSYSException ex)
            {
                ErrorLogger.Log(ex, ErrLevel.FATAL);
            }

            
        }

        // Compute the clearness index. Duffie and Beckman (1991)
        double GetClearnessIndex            // (o) clearness index [0-1] 
            (double HGlo                    // (i) global irradiance on horizontal [W/m2] 
            , double NExtra                 // (i) normal extraterrestrial irradiance [W/m2] 
            , double Zenith                 // (i) zenith angle of sun [radians] 
            )
        {
            if (Zenith > Math.PI / 2)
                return 1.0;
            else
                return HGlo / (NExtra * Math.Cos(Zenith));
        }


        // Compute the diffuse fraction given the clearness index, using the
        // Orgill and Hollands formula
        // Duffie, J.A., and Beckman, W.A., Solar Engineering of Thermal
        // Processes, 2nd edition, John Wiley & Sons (1991), p. 81
        double GetDiffuseFraction               // (o) diffuse fraction (Orgill and Hollands formula) [0-1] 
            (
            double kt                          // (i) clearness index [0-1] 
            )
        {

            double kd;                          // Diffuse fraction defined locally to be returned later
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