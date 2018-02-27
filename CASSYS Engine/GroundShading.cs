// CASSYS - Grid connected PV system modelling software
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: GroundShading.cs
//
// Revision History:
//
// Description:
// This class is responsible for the simulation of the shading effects on the
// beam and diffuse components of ground irradiance for bifacial modules
//
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
// Code adapted from https://github.com/NREL/bifacialvf which is based on
// https://www.nrel.gov/docs/fy17osti/67847.pdf
///////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;

namespace CASSYS
{
    // This enumeration is used to define the type of row for which to calculate
    public enum RowType { INTERIOR = 0, FRONT = 1, BACK = 2, SINGLE = 3 };

    class GroundShading
    {
        // Input variables
        double itsPanelWidth;                       // Panel bandwidth [m]
        double itsClearance;                        // Panel ground clearance [panel slope lengths]
        double itsPitch;                            // The distance between the rows [panel slope lengths]
        double transFactor;                         // Transmission factor [#]
        RowType itsRowType;                         // Row type based on the position of row relative to others [unitless]
        TrackMode itsTrackMode;                     // The tracking mode, used to determine when to calculate sky config factors [unitless]
        Shading itsShading;                         // Used to calculate partial shading on front/back of module
        int n;                                      // Number of segments into which to divide up the ground [#]

        // Ground Shading local variables/arrays and intermediate calculation variables and arrays
        int[] backGroundSH;                         // Ground shade factors for ground segments to the back, 0 = not shaded, 1 = shaded [#]
        int[] frontGroundSH;                        // Ground shade factors for ground segments to the front, 0 = not shaded, 1 = shaded [#]
        double[] backSkyConfigFactors;              // Fraction of isotropic diffuse sky radiation present on ground segments to the back [#]
        double[] frontSkyConfigFactors;             // Fraction of isotropic diffuse sky radiation present on ground segments to the front [#]
        string modShad;                             // TODO: remove later
        string skyConfig;                           // TODO: remove later

        // Output variables
        public double pvFrontSH;                    // Fraction of the front surface of the PV panel that is shaded [# from 0-1]
        public double pvBackSH;                     // Fraction of the back surface of the PV panel that is shaded [# from 0-1]
        public double maxShadow;                    // Maximum shadow length projected to front (-) or back (+) from front of module row [panel slope lengths]
        public double[] backGroundGHI;              // Sum of irradiance components for each of the n ground segments to back of the PV row [W/m2]
        public double[] frontGroundGHI;             // Sum of irradiance components for each of the n ground segments to front of the PV row [W/m2]

        // Ground shading constructor
        public GroundShading()
        {

        }

        // Config manages calculations and initializations that need only to be run once
        public void Config
            (
              int n                         // Number of segments into which to divide up the ground [#]
            )
        {
            this.n = n;

            // TODO: allow user to define?
            transFactor = 0;
            // TODO: support multiple types of rows?
            itsRowType = RowType.INTERIOR;

            // Create and configure tracker so that track mode can be read and interpreted
            Tracker SimTracker = new Tracker();
            SimTracker.Config();
            itsTrackMode = SimTracker.itsTrackMode;

            // Read in values from .csyx file
            itsPanelWidth = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "CollBandWidth", ErrLevel.FATAL));
            itsPitch = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "Pitch", ErrLevel.FATAL));
            itsClearance = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "CollGroundClearance", ErrLevel.FATAL));

            // Convert ground clearance and pitch to panel slope lengths
            itsPitch /= itsPanelWidth;
            itsClearance /= itsPanelWidth;

            // Initialize arrays
            backGroundSH = new int[n];
            frontGroundSH = new int[n];
            backGroundGHI = new double[n];
            frontGroundGHI = new double[n];
            backSkyConfigFactors = new double[n];
            frontSkyConfigFactors = new double[n];

            // Calculate sky configuration factors if not a tracking system; otherwise, will be done in Calculate()
            if (itsTrackMode == TrackMode.NOAT)
            {
                double PanelTilt = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PlaneTilt", ErrLevel.FATAL));
                CalcSkyConfigFactors(PanelTilt);
            }
        }

        // Calculate manages calculations that need to be run for each time step
        public void Calculate
            (
              double SunZenith                                  // The zenith position of the sun with 0 being normal to the earth [radians]
            , double SunAzimuth                                 // The azimuth position of the sun relative to 0 being true south. Positive if west, negative if east [radians]
            , double PanelTilt                                  // The angle between the surface tilt of the module and the ground [radians]
            , double PanelAzimuth                               // The azimuth direction in which the surface is facing. Positive if west, negative if east [radians]
            , double HDif                                       // Diffuse horizontal irradiance [W/m2]
            , double HDir                                       // Direct horizontal irradiance [W/m2]
            , Shading SimShading                                // Used to calculate front and back partial shading
            , DateTime ts                                       // EC: delete later, only for printing convenience
            )
        {
            itsShading = SimShading;

            // Calculate sky configuration factors if a tracking system; otherwise, has already been done in Config()
            if (itsTrackMode != TrackMode.NOAT)
            {
                CalcSkyConfigFactors(PanelTilt);
            }

            // EC: delete printing code later
            modShad = Environment.NewLine + ts.ToString();
            string shBack = Environment.NewLine + ts.ToString();
            string irrBack = Environment.NewLine + ts.ToString();

            // Calculate shading for ground underneath PV modules
            CalcGroundShading(SunZenith, SunAzimuth, PanelTilt, PanelAzimuth);
            modShad += "," + pvFrontSH + "," + pvBackSH;

            // Calculate irradiance for ground underneath PV modules
            for (int i = 0; i < n; i++)
            {
                // Add diffuse sky component viewed by ground
                backGroundGHI[i] = HDif * backSkyConfigFactors[i];
                // Add direct beam component, depending on shading
                if (backGroundSH[i] == 0)
                {
                    backGroundGHI[i] += HDir;
                }
                else
                {
                    backGroundGHI[i] += HDir * transFactor;
                }

                // Add diffuse sky component viewed by ground
                frontGroundGHI[i] = HDif * frontSkyConfigFactors[i];
                // Add direct beam component, depending on shading
                if (frontGroundSH[i] == 0)
                {
                    frontGroundGHI[i] += HDir;
                }
                else
                {
                    frontGroundGHI[i] += HDir * transFactor;
                }
                irrBack += "," + backGroundGHI[i].ToString();
                shBack += "," + backGroundSH[i].ToString();
            }
            //File.AppendAllText("irrBack.csv", irrBack);
            //File.AppendAllText("shBack.csv", shBack);
            //File.AppendAllText("modShad.csv", modShad);
        }

        // Divides the ground between two PV rows into n segments and determines the fraction of isotropic diffuse sky radiation present on each segment
        public void CalcSkyConfigFactors
            (
              double PanelTilt                                  // The angle between the surface tilt of the module and the ground [radians]
            )
        {
            if (itsClearance < 0.0)
            {
                ErrorLogger.Log("Ground clearance of panel cannot be negative.", ErrLevel.FATAL);
            }

            if (itsRowType == RowType.INTERIOR)
            {
                // Divide the row-to-row spacing into n intervals for calculating ground shade factors
                double delta = itsPitch / n;
                // Initialize horizontal dimension x to provide midpoint intervals
                double x = -delta / 2.0;

                double skyAhead = 0;
                double skyAbove = 0;
                double skyBehind = 0;

                for (int i = 0; i < n; i++)
                {
                    x += delta;

                    // Calculate sky configuration factors ahead, above, and behind the ground segment.
                    // Directions are split into three so that view can extend freely backward and forward, until view is blocked.
                    skyAhead = CalcSkyConfigDirection(PanelTilt, x, -1);
                    skyAbove = CalcSkyConfigDirection(PanelTilt, x, 0);
                    skyBehind = CalcSkyConfigDirection(PanelTilt, x, 1);

                    // Sum sky configuration, using same values for both back and front since we assume homogeneity for interior rows
                    backSkyConfigFactors[i] = skyAhead + skyAbove + skyBehind;
                    frontSkyConfigFactors[i] = skyAhead + skyAbove + skyBehind;
                    skyConfig += Environment.NewLine + i + "," + skyAhead + "," + skyAbove + "," + skyBehind + "," + backSkyConfigFactors[i];
                }
                File.WriteAllText("skyConfig.csv", skyConfig);
            }
            else
            {
                ErrorLogger.Log("Incorrect row type.", ErrLevel.FATAL);
            }
        }

        double CalcSkyConfigDirection
            (
              double PanelTilt                                  // The angle between the surface tilt of the module and the ground [radians]
            , double x                                          // Horizontal dimension in the row-to-row ground area
            , double direction                                  // The direction in which to move along the x-axis [-1, 0, 1]
            )
        {
            double h = Math.Sin(PanelTilt);                     // Vertical height of sloped PV panel [panel slope lengths]
            double x1 = Math.Cos(PanelTilt);                    // Horizontal distance from front of panel to back of panel [panel slope lengths]
            double d = itsPitch - x1;                           // Horizontal distance from back of one row to front of the next [panel slope lengths]

            double offset = direction;                          // Initialize offset to begin at first unit of given direction
            double skyPatch = 0;                                // Configuration factor for view of sky in single row-to-row area
            double skySum = 0;                                  // Configuration factor for all sky views in given direction

            double angA = 0;
            double angB = 0;
            double angC = 0;
            double angD = 0;
            double beta1 = 0;
            double beta2 = 0;

            // Sum sky configuration factors until sky can no longer be seen. For direction = 0, only do the calculation once.
            do
            {
                // Angle from ground point to top of panel P
                angA = Math.Atan2(h + itsClearance, (offset + 1) * itsPitch + x1 - x);
                if (angA < 0.0)
                {
                    angA += Math.PI;
                }
                // Angle from ground point to bottom of panel P
                angB = Math.Atan2(itsClearance, (offset + 1) * itsPitch - x);
                if (angB < 0.0)
                {
                    angB += Math.PI;
                }
                beta1 = Math.Max(angA, angB);

                // Angle from ground point to top of panel Q
                angC = Math.Atan2(h + itsClearance, offset * itsPitch + x1 - x);
                if (angC < 0.0)
                {
                    angC += Math.PI;
                }
                // Angle from ground point to bottom of panel Q
                angD = Math.Atan2(itsClearance, offset * itsPitch - x);
                if (angD < 0.0)
                {
                    angD += Math.PI;
                }
                beta2 = Math.Min(angC, angD);

                skyPatch = 0;
                // If there is an opening in the sky through which the sun is seen, calculate view factor of sky patch
                if (beta2 > beta1)
                {
                    skyPatch = 0.5 * (Math.Cos(beta1) - Math.Cos(beta2));
                }
                skySum += skyPatch;
                offset += direction;
            } while (skyPatch > 0.0000000001 && offset != 0);

            return skySum;
        }

        // Divides the ground between two PV rows into n segments and determines direct beam shading (0 = not shaded, 1 = shaded) for each segment
        public void CalcGroundShading
            (
              double SunZenith                                  // The zenith position of the sun with 0 being normal to the earth [radians]
            , double SunAzimuth                                 // The azimuth position of the sun relative to 0 being true south. Positive if west, negative if east [radians]
            , double PanelTilt                                  // The angle between the surface tilt of the module and the ground [radians]
            , double PanelAzimuth                               // The azimuth direction in which the surface is facing. Positive if west, negative if east [radians]
            )
        {
            double h = Math.Sin(PanelTilt);                 // Vertical height of sloped PV panel [panel slope lengths]
            double x1 = Math.Cos(PanelTilt);                // Horizontal distance from front of panel to back of panel [panel slope lengths]
            double d = itsPitch - x1;                       // Horizontal distance from back of one row to front of the next [panel slope lengths]

            double SunElevation = (Math.PI / 2) - SunZenith;
            // Horizontal length of shadow normal to row from module top to bottom (base of triangle formed by beam of sun and height of module top from bottom)
            double Lh = (h / Math.Tan(SunElevation)) * Math.Cos(PanelAzimuth - SunAzimuth);
            // Horizontal length of shadow normal to row from module top to ground (base of triangle formed by beam of sun and height of module top from ground)
            double Lhc = ((h + itsClearance) / Math.Tan(SunElevation)) * Math.Cos(PanelAzimuth - SunAzimuth);
            // Horizontal length of shadow normal to row from module bottom to ground (base of triangle formed by beam of sun and height of module bottom from ground)
            double Lc = (itsClearance / Math.Tan(SunElevation)) * Math.Cos(PanelAzimuth - SunAzimuth);

            double s1Start = 0;                             // Shading start position for first potential shading segment
            double s1End = 0;                               // Shading end position for first potential shading segment
            double s2Start = 0;                             // Shading start position for second potential shading segment
            double s2End = 0;                               // Shading end position for second potential shading segment
            double SStart = 0;                              // Shading start position for placeholder? segment
            double SEnd = 0;                                // Shading start position for placeholder? segment

            if (itsRowType == RowType.INTERIOR)
            {
                double FrontPA = Tilt.GetProfileAngle(SunZenith, SunAzimuth, PanelAzimuth) * Util.RTOD;
                double BackPA = Tilt.GetProfileAngle(SunZenith, SunAzimuth, PanelAzimuth + Math.PI) * Util.RTOD;
                modShad += "," + (SunZenith * Util.RTOD) + "," + (SunAzimuth * Util.RTOD) + "," + (itsShading.FrontSLA * Util.RTOD) + "," + FrontPA + "," + (itsShading.BackSLA * Util.RTOD) + "," + BackPA;
                // Sun below horizon, everything completely shaded
                if (SunElevation < 0)
                {
                    pvFrontSH = 1.0;
                    pvBackSH = 1.0;
                    s1Start = 0.0;
                    s1End = itsPitch;
                }
                // Front side of PV module partially shaded, back completely shaded, ground completely shaded
                else if (Lh > d)
                {
                    pvFrontSH = itsShading.GetFrontShadedFraction(SunZenith, SunAzimuth, PanelTilt);
                    pvBackSH = 1.0;
                    s1Start = 0.0;
                    s1End = itsPitch;
                }
                // Front side of PV module completely shaded, back partially shaded, ground completely shaded
                else if (Lh < -(itsPitch + x1))
                {
                    pvFrontSH = 1.0;
                    pvBackSH = itsShading.GetBackShadedFraction(SunZenith, SunAzimuth, PanelTilt);
                    s1Start = 0.0;
                    s1End = itsPitch;
                }
                // Assume ground is partially shaded
                else
                {
                    // Shadow to back of row - module front unshaded, back shaded
                    if (Lhc >= 0.0)
                    {
                        pvFrontSH = 0.0;
                        pvBackSH = 1.0;
                        SStart = Lc;
                        SEnd = Lhc + x1;
                        // Put shadow in correct row-to-row space if needed
                        while (SStart > itsPitch)
                        {
                            SStart -= itsPitch;
                            SEnd -= itsPitch;
                        }
                        s1Start = SStart;
                        s1End = SEnd;
                        // Need to use two shade areas. Transpose the area that extends beyond itsPitch to the front of the row-to-row space
                        if (s1End > itsPitch)
                        {
                            s1End = itsPitch;
                            s2Start = 0.0;
                            s2End = SEnd - itsPitch;
                            if (s2End > s1Start)
                            {
                                ErrorLogger.Log("Unexpected shading coordinates encountered.", ErrLevel.FATAL);
                            }
                        }
                    }
                    // Shadow to front of row - either front or back might be shaded, depending on tilt and other factors
                    else
                    {
                        // Sun hits front of module. Shadow cast by bottom of module extends further forward than shadow cast by top
                        if (Lc < Lhc + x1)
                        {
                            pvFrontSH = 0.0;
                            pvBackSH = 1.0;
                            SStart = Lc;
                            SEnd = Lhc + x1;
                        }
                        // Sun hits back of module. Shadow cast by top of module extends further forward than shadow cast by bottom
                        else
                        {
                            pvFrontSH = 1.0;
                            pvBackSH = 0.0;
                            SStart = Lhc + x1;
                            SEnd = Lc;
                        }
                        // Put shadow in correct row-to-row space if needed
                        while (SStart < 0.0)
                        {
                            SStart += itsPitch;
                            SEnd += itsPitch;
                        }
                        s1Start = SStart;
                        s1End = SEnd;
                        // Need to use two shade areas. Transpose the area that extends beyond itsPitch to the front of the row-to-row space
                        if (s1End > itsPitch)
                        {
                            s1End = itsPitch;
                            s2Start = 0.0;
                            s2End = SEnd - itsPitch;
                            if (s2End > s1Start)
                            {
                                ErrorLogger.Log("Unexpected shading coordinates encountered.", ErrLevel.FATAL);
                            }
                        }
                    }
                }

                // Divide the row-to-row spacing into n intervals for calculating ground shade factors
                double delta = itsPitch / n;
                // Initialize horizontal dimension x to provide midpoint intervals
                double x = -delta / 2.0;

                // Determine whether shaded or sunny for each n ground segments
                // TODO: improve accuracy (especially for n < 100) by setting 0 or 1 if < or > 50% of segment is shaded
                for (int i = 0; i < n; i++)
                {
                    x += delta;
                    if ((x >= s1Start && x < s1End) || (x >= s2Start && x < s2End))
                    {
                        // x within a shaded interval, so set both groundSH to 1 to indicate shaded
                        backGroundSH[i] = 1;
                        frontGroundSH[i] = 1;
                    }
                    else
                    {
                        // x not within a shaded interval, so set both groundSH to 0 to indicate sunny
                        backGroundSH[i] = 0;
                        frontGroundSH[i] = 0;
                    }
                }
            }
            else
            {
                ErrorLogger.Log("Incorrect row type.", ErrLevel.FATAL);
            }
            // Determine maximum shadow length projected from the front of the PV module row
            maxShadow = (Math.Abs(s1Start) > Math.Abs(s1End)) ? s1Start : s1End;
        }
    }
}
