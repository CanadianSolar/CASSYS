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
// Ref 1: Marion, B.; Ayala S.; Deline, C. "Bifacial PV View Factor model"
//      National Renewable Energy Laboratory
//      https://github.com/NREL/bifacialvf
//
// Ref 2: Marion, B. et al. "A Practical Irradiance Model for Bifacial PV Modules"
//      National Renewable Energy Laboratory, 2017
//      https://www.nrel.gov/docs/fy17osti/67847.pdf
///////////////////////////////////////////////////////////////////////////////

using System;
using System.IO;

namespace CASSYS
{
    class GroundShading
    {
        // Parameters for the ground shading class
        double itsArrayBW;                          // Array bandwidth [m]
        double itsClearance;                        // Array ground clearance [panel slope lengths]
        double itsPitch;                            // The distance between the rows [panel slope lengths]
        double itsPanelTilt;                        // The angle between the surface tilt of the module and the ground [radians]
        double itsPanelAzimuth;                     // The angle between horizontal projection of normal to module surface and true South [radians]
        double transFactor;                         // Fraction of light that is transmitted through the array [#]
        public int numGroundSegs;                   // Number of segments into which to divide up the ground [#]
        int numRows;                                // Number of rows in the system [#]
        TrackMode itsTrackMode;                     // Used to determine tracking mode
        RowType itsRowType;                         // Used to differentiate row type (interior or single)

        // Ground shading local variables/arrays and intermediate calculation variables and arrays
        int[] frontGroundSH;                        // Ground shade factors for ground segments to the front of the row, 0 = not shaded, 1 = shaded [#]
        int[] rearGroundSH;                         // Ground shade factors for ground segments to the rear of the row, 0 = not shaded, 1 = shaded [#]
        double[] frontSkyViewFactors;               // Fraction of isotropic diffuse sky radiation present on ground segments to the front of the row [#]
        double[] rearSkyViewFactors;                // Fraction of isotropic diffuse sky radiation present on ground segments to the rear of the row [#]

        // Output variables
        public double[] frontGroundGHI;             // Global irradiance for each of the ground segments to front of the row [W/m2]
        public double[] rearGroundGHI;              // Global irradiance for each of the ground segments to rear of the row [W/m2]
        public double aveGroundGHI;                 // Average global irradiance on ground segment to the rear of row [W/m2]

        // Ground shading constructor
        public GroundShading()
        {

        }

        // Config manages calculations and initializations that need only to be run once
        public void Config()
        {
            bool useBifacial = Convert.ToBoolean(ReadFarmSettings.GetInnerText("Bifacial", "UseBifacialModel", ErrLevel.FATAL));

            if (useBifacial)
            {
                // Number of segments into which to divide up the ground [#]
                numGroundSegs = Util.NUM_GROUND_SEGS;

                switch (ReadFarmSettings.GetAttribute("O&S", "ArrayType", ErrLevel.FATAL))
                {
                    // In all cases, pitch and clearance must be normalized to panel slope lengths
                    case "Fixed Tilted Plane":
                        itsTrackMode = TrackMode.NOAT;
                        if (String.Compare(ReadFarmSettings.CASSYSCSYXVersion, "0.9.3") >= 0)
                        {
                            itsPanelTilt = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PlaneTiltFix", ErrLevel.FATAL));
                        }
                        else
                        {
                            itsPanelTilt = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PlaneTilt", ErrLevel.FATAL));
                        }
                        // itsPitch will be assigned in the below (numRows == 1) conditional
                        itsArrayBW = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "CollBandWidthFix", ErrLevel.FATAL));
                        itsClearance = Convert.ToDouble(ReadFarmSettings.GetInnerText("Bifacial", "GroundClearance", ErrLevel.FATAL)) / itsArrayBW;
                        numRows = 1;
                        break;
                    case "Unlimited Rows":
                        itsTrackMode = TrackMode.NOAT;
                        itsPanelTilt = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PlaneTilt", ErrLevel.FATAL));
                        itsArrayBW = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "CollBandWidth", ErrLevel.FATAL));
                        itsPitch = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "Pitch", ErrLevel.FATAL)) / itsArrayBW;
                        itsClearance = Convert.ToDouble(ReadFarmSettings.GetInnerText("Bifacial", "GroundClearance", ErrLevel.FATAL)) / itsArrayBW;
                        numRows = int.Parse(ReadFarmSettings.GetInnerText("O&S", "RowsBlock", ErrLevel.FATAL));
                        break;
                    case "Single Axis Elevation Tracking (E-W)":
                        itsTrackMode = TrackMode.SAXT;
                        itsArrayBW = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "WActiveSAET", ErrLevel.FATAL));
                        itsPitch = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PitchSAET", ErrLevel.FATAL)) / itsArrayBW;
                        numRows = int.Parse(ReadFarmSettings.GetInnerText("O&S", "RowsBlockSAET", ErrLevel.FATAL));
                        break;
                    case "Single Axis Horizontal Tracking (N-S)":
                        itsTrackMode = TrackMode.SAXT;
                        itsArrayBW = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "WActiveSAST", ErrLevel.FATAL));
                        itsPitch = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PitchSAST", ErrLevel.FATAL)) / itsArrayBW;
                        numRows = int.Parse(ReadFarmSettings.GetInnerText("O&S", "RowsBlockSAST", ErrLevel.FATAL));
                        break;
                    default:
                        ErrorLogger.Log("Bifacial is not supported for the selected orientation and shading.", ErrLevel.FATAL);
                        break;
                }

                transFactor = Convert.ToDouble(ReadFarmSettings.GetInnerText("Bifacial", "PanelTransFactor", ErrLevel.FATAL));

                if (numRows == 1)
                {
                    // Pitch is needed for a single row because of ground patch calculations and geometry. Take value 100x greater than array bandwidth.
                    itsPitch = 100;
                    itsRowType = RowType.SINGLE;
                }
                else
                {
                    itsRowType = RowType.INTERIOR;
                }

                // Initialize arrays
                frontGroundSH = new int[numGroundSegs];
                rearGroundSH = new int[numGroundSegs];

                frontSkyViewFactors = new double[numGroundSegs];
                rearSkyViewFactors = new double[numGroundSegs];

                frontGroundGHI = new double[numGroundSegs];
                rearGroundGHI = new double[numGroundSegs];

                // Calculate sky view factors for diffuse shading. Stays constant for non-tracking systems, so done here in Config()
                if (itsTrackMode == TrackMode.NOAT)
                {
                    CalcSkyViewFactors();
                }
            }
        }

        // Calculate manages calculations that need to be run for each time step
        public void Calculate
            (
              double SunZenith                                  // The zenith position of the sun with 0 being normal to the earth [radians]
            , double SunAzimuth                                 // The azimuth position of the sun relative to 0 being true south. Positive if west, negative if east [radians]
            , double PanelTilt                                  // The angle between the surface tilt of the module and the ground [radians]
            , double PanelAzimuth                               // The azimuth direction in which the surface is facing. Positive if west, negative if east [radians]
            , double Clearance                                  // Array ground clearance [m]
            , double HDir                                       // Direct horizontal irradiance [W/m2]
            , double HDif                                       // Diffuse horizontal irradiance [W/m2]
            , DateTime ts                                       // Time stamp analyzed, used for printing .csv files
            )
        {
            // For tracking systems, panel tilt, azimuth, and ground clearance will change at each time step
            itsPanelTilt = PanelTilt;
            itsPanelAzimuth = PanelAzimuth;
            itsClearance = Clearance / itsArrayBW;              // Convert to panel slope lengths

            if (itsTrackMode != TrackMode.NOAT)
            {
                // Calculate sky view factors for diffuse shading. Changes every time step for tracking systems, so done here in Calculate()
                CalcSkyViewFactors();
            }

            // Calculate beam shading for ground underneath PV modules. Three different area of ground are accounted for: interior rows, front of first row, and back of last row
            CalcGroundShading(SunZenith, SunAzimuth);

            // Calculate global irradiance for ground underneath PV modules, to the front and to the rear
            aveGroundGHI = 0;
            for (int i = 0; i < numGroundSegs; i++)
            {
                // Add diffuse sky component viewed by ground
                frontGroundGHI[i] = HDif * frontSkyViewFactors[i];
                // Add direct beam component, depending on shading
                if (frontGroundSH[i] == 0)
                {
                    frontGroundGHI[i] += HDir;
                }
                else
                {
                    frontGroundGHI[i] += HDir * transFactor;
                }

                // Add diffuse sky component viewed by ground
                rearGroundGHI[i] = HDif * rearSkyViewFactors[i];
                // Add direct beam component, depending on shading
                if (rearGroundSH[i] == 0)
                {
                    rearGroundGHI[i] += HDir;
                }
                else
                {
                    rearGroundGHI[i] += HDir * transFactor;
                }

                // Compute average ground irradiance
                aveGroundGHI += rearGroundGHI[i] / numGroundSegs;
            }

            // Option to print details of the model in .csv files. Only recommended for single day simulations.
            // PrintModel(ts, PanelAzimuth);
        }

        // Divides the ground between two PV rows into n segments and determines the fraction of isotropic diffuse sky radiation present on each segment
        void CalcSkyViewFactors()
        {
            // Divide the row-to-row spacing into n intervals for calculating ground shade factors
            double delta = itsPitch / numGroundSegs;
            // Initialize horizontal dimension x to provide midpoint intervals
            double x = 0;

            double skyAhead = 0;
            double skyAbove = 0;
            double skyBehind = 0;

            for (int i = 0; i < numGroundSegs; i++)
            {
                x = (i + 0.5) * delta;

                if (itsRowType == RowType.SINGLE)
                {
                    // Calculate sky view factor without front limiting panel to model front row.
                    skyBehind = CalcSkyViewDirection(x, RowType.LAST, 1);
                    skyAbove = CalcSkyViewDirection(x, RowType.FIRST, 0);
                    frontSkyViewFactors[i] = skyAbove + skyBehind;

                    // Calculate sky view factor without back limiting panel to model back row.
                    skyAhead = CalcSkyViewDirection(x, RowType.FIRST, -1);
                    skyAbove = CalcSkyViewDirection(x, RowType.LAST, 0);
                    rearSkyViewFactors[i] = skyAbove + skyAhead;
                }
                else
                {
                    // Calculate and summarize sky view factors ahead, above, and behind the ground segment for interior rows
                    // Directions are split into three so that view can extend freely backward and forward, until view is blocked.
                    skyAhead = CalcSkyViewDirection(x, RowType.INTERIOR, -1);
                    skyBehind = CalcSkyViewDirection(x, RowType.INTERIOR, 1);
                    skyAbove = CalcSkyViewDirection(x, RowType.INTERIOR, 0);

                    // Assume homogeneity to the front and rear of interior rows
                    frontSkyViewFactors[i] = skyAhead + skyAbove + skyBehind;
                    rearSkyViewFactors[i] = skyAhead + skyAbove + skyBehind;
                }
            }
        }

        double CalcSkyViewDirection
            (
              double x                                          // Horizontal dimension in the row-to-row ground area
            , RowType rowType                                   // The position of the row being calculated relative to others [unitless]
            , double direction                                  // The direction in which to move along the x-axis [-1, 0, 1]
            )
        {
            double h = Math.Sin(itsPanelTilt);                  // Vertical height of sloped PV panel [panel slope lengths]
            double b = Math.Cos(itsPanelTilt);                  // Horizontal distance from front of panel to back of panel [panel slope lengths]

            double offset = direction;                          // Initialize offset to begin at first unit of given direction
            double skyPatch = 0;                                // View factor for view of sky in single row-to-row area
            double skySum = 0;                                  // View factor for all sky views in given direction

            double angA = 0;
            double angB = 0;
            double angC = 0;
            double angD = 0;
            double beta1 = 0;                                   // Start of ground's field of view that sees the sky segment
            double beta2 = 0;                                   // End of ground's field of view that sees the sky segment

            // Sum sky view factors until sky view factor contributes <= 1% of sum
            // Only loop the calculation for rows extending forward or backward, so break loop when direction = 0.
            do
            {
                // Set back limiting angle to 0 since there is no row behind.
                if (rowType == RowType.LAST)
                {
                    beta1 = 0;
                }
                else
                {
                    // Angle from ground point to top of panel P
                    angA = Math.Atan2(h + itsClearance, (offset + 1) * itsPitch + b - x);
                    // Angle from ground point to bottom of panel P
                    angB = Math.Atan2(itsClearance, (offset + 1) * itsPitch - x);

                    beta1 = Math.Max(angA, angB);
                }

                // Set front limiting angle to PI since there is no row ahead.
                if (rowType == RowType.FIRST)
                {
                    beta2 = Math.PI;
                }
                else
                {
                    // Angle from ground point to top of panel Q
                    angC = Math.Atan2(h + itsClearance, offset * itsPitch + b - x);
                    // Angle from ground point to bottom of panel Q
                    angD = Math.Atan2(itsClearance, offset * itsPitch - x);

                    beta2 = Math.Min(angC, angD);
                }

                // If there is an opening in the sky through which the sun is seen, calculate view factor of sky patch
                skyPatch = (beta2 > beta1) ? RadiationProc.GetViewFactor(beta1, beta2) : 0;

                skySum += skyPatch;
                offset += direction;
            } while (offset != 0 && skyPatch > (0.01 * skySum));

            return skySum;
        }

        // Divides the ground between two PV rows into n segments and determines direct beam shading (0 = not shaded, 1 = shaded) for each segment
        void CalcGroundShading
            (
              double SunZenith                                  // The zenith position of the sun with 0 being normal to the earth [radians]
            , double SunAzimuth                                 // The azimuth position of the sun relative to 0 being true south. Positive if west, negative if east [radians]
            )
        {
            // When sun is below horizon, set everything to shaded
            if (SunZenith > (Math.PI / 2))
            {
                for (int i = 0; i < numGroundSegs; i++)
                {
                    frontGroundSH[i] = 1;
                    rearGroundSH[i] = 1;
                }
            }
            else
            {
                double h = Math.Sin(itsPanelTilt);                  // Vertical height of sloped PV panel [panel slope lengths]
                double b = Math.Cos(itsPanelTilt);                  // Horizontal distance from front of panel to back of panel [panel slope lengths]

                double FrontPA = Tilt.GetProfileAngle(SunZenith, SunAzimuth, itsPanelAzimuth);

                double Lh = h / Math.Tan(FrontPA);                          // Base of triangle formed by beam of sun and height of module top from bottom
                double Lc = itsClearance / Math.Tan(FrontPA);               // Base of triangle formed by beam of sun and height of module bottom from ground
                double Lhc = (h + itsClearance) / Math.Tan(FrontPA);        // Base of triangle formed by beam of sun and height of module top from ground

                double s1Start = 0;                             // Shading start position for first potential shading segment
                double s1End = 0;                               // Shading end position for first potential shading segment
                double s2Start = 0;                             // Shading start position for second potential shading segment
                double s2End = 0;                               // Shading end position for second potential shading segment
                double SStart = 0;                              // Shading start position for placeholder segment
                double SEnd = 0;                                // Shading start position for placeholder segment

                // Divide the row-to-row spacing into n intervals for calculating ground shade factors
                double delta = itsPitch / numGroundSegs;
                // Initialize horizontal dimension x to provide midpoint intervals
                double x = 0;

                if (itsRowType == RowType.SINGLE)
                {
                    // Calculate front ground shading
                    // Front side of PV module completely sunny, ground partially shaded
                    if (Lh > 0.0)
                    {
                        s1Start = Lc;
                        s1End = Lhc + b;
                    }
                    // Front side of PV module completely shaded, ground completely shaded
                    else if (Lh < -(itsPitch + b))
                    {
                        s1Start = -itsPitch;
                        s1End = itsPitch;
                    }
                    // Shadow to front of row - either front or back might be shaded, depending on tilt and other factors
                    else
                    {
                        // Sun hits front of module. Shadow cast by bottom of module extends further forward than shadow cast by top
                        if (Lc < Lhc + b)
                        {
                            s1Start = Lc;
                            s1End = Lhc + b;
                        }
                        // Sun hits back of module. Shadow cast by top of module extends further forward than shadow cast by bottom
                        else
                        {
                            s1Start = Lhc + b;
                            s1End = Lc;
                        }
                    }

                    // Determine whether shaded or sunny for each n ground segments
                    // TODO: improve accuracy (especially for n < 100) by setting 1 only if > 50% of segment is shaded
                    for (int i = 0; i < numGroundSegs; i++)
                    {
                        // Offset x coordinate by -itsPitch because row ahead is being measured
                        x = (i + 0.5) * delta - itsPitch;
                        if (x >= s1Start && x < s1End)
                        {
                            // x within a shaded interval, so set to 1 to indicate shaded
                            frontGroundSH[i] = 1;
                        }
                        else
                        {
                            // x not within a shaded interval, so set to 0 to indicate sunny
                            frontGroundSH[i] = 0;
                        }
                    }

                    // Calculate rear ground shading
                    // Back side of PV module completely shaded, ground completely shaded
                    if (Lh > itsPitch - b)
                    {
                        s1Start = 0.0;
                        s1End = itsPitch;
                    }
                    // Shadow to front of row - either front or back might be shaded, depending on tilt and other factors
                    else
                    {
                        // Sun hits front of module. Shadow cast by bottom of module extends further forward than shadow cast by top
                        if (Lc < Lhc + b)
                        {
                            s1Start = Lc;
                            s1End = Lhc + b;
                        }
                        // Sun hits back of module. Shadow cast by top of module extends further forward than shadow cast by bottom
                        else
                        {
                            s1Start = Lhc + b;
                            s1End = Lc;
                        }
                    }

                    // Determine whether shaded or sunny for each n ground segments
                    // TODO: improve accuracy (especially for n < 100) by setting 1 only if > 50% of segment is shaded
                    for (int i = 0; i < numGroundSegs; i++)
                    {
                        x = (i + 0.5) * delta;
                        if (x >= s1Start && x < s1End)
                        {
                            // x within a shaded interval, so set to 1 to indicate shaded
                            rearGroundSH[i] = 1;
                        }
                        else
                        {
                            // x not within a shaded interval, so set to 0 to indicate sunny
                            rearGroundSH[i] = 0;
                        }
                    }
                }
                else
                {
                    // Calculate interior ground shading
                    // Front side of PV module partially shaded, back completely shaded, ground completely shaded
                    if (Lh > itsPitch - b)
                    {
                        s1Start = 0.0;
                        s1End = itsPitch;
                    }
                    // Front side of PV module completely shaded, back partially shaded, ground completely shaded
                    else if (Lh < -(itsPitch + b))
                    {
                        s1Start = 0.0;
                        s1End = itsPitch;
                    }
                    // Assume ground is partially shaded
                    else
                    {
                        // Shadow to back of row - module front unshaded, back shaded
                        if (Lhc >= 0.0)
                        {
                            SStart = Lc;
                            SEnd = Lhc + b;
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
                                if (s2End - s1Start > 0.000001)
                                {
                                    ErrorLogger.Log("Unexpected shading coordinates encountered.", ErrLevel.FATAL);
                                }
                            }
                        }
                        // Shadow to front of row - either front or back might be shaded, depending on tilt and other factors
                        else
                        {
                            // Sun hits front of module. Shadow cast by bottom of module extends further forward than shadow cast by top
                            if (Lc < Lhc + b)
                            {
                                SStart = Lc;
                                SEnd = Lhc + b;
                            }
                            // Sun hits back of module. Shadow cast by top of module extends further forward than shadow cast by bottom
                            else
                            {
                                SStart = Lhc + b;
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
                                if (s2End - s1Start > 0.000001)
                                {
                                    ErrorLogger.Log("Unexpected shading coordinates encountered.", ErrLevel.FATAL);
                                }
                            }
                        }
                    }

                    // Determine whether shaded or sunny for each n ground segments
                    // TODO: improve accuracy (especially for n < 100) by setting 1 only if > 50% of segment is shaded
                    for (int i = 0; i < numGroundSegs; i++)
                    {
                        x = (i + 0.5) * delta;

                        // Assume homogeneity to the front and rear of interior rows
                        if ((x >= s1Start && x < s1End) || (x >= s2Start && x < s2End))
                        {
                            // x within a shaded interval, so set to 1 to indicate shaded
                            frontGroundSH[i] = 1;
                            rearGroundSH[i] = 1;
                        }
                        else
                        {
                            // x not within a shaded interval, so set to 0 to indicate sunny
                            frontGroundSH[i] = 0;
                            rearGroundSH[i] = 0;
                        }
                    }
                }
            }
        }

        void PrintModel
            (
              DateTime ts                                       // Time stamp analyzed, used for printing the model
            , double PanelAzimuth                               // The azimuth direction in which the surface is facing. Positive if west, negative if east [radians]
            )
        {
            string skyViewAll = Environment.NewLine + ts;
            string skyViewOne = "";

            string irrFront = Environment.NewLine + ts;
            string irrRear = Environment.NewLine + ts;

            string setup = Environment.NewLine + ts + "," + (PanelAzimuth * Util.RTOD) + "," + (itsPanelTilt * Util.RTOD) + "," + (itsPitch * itsArrayBW) + "," + (itsClearance * itsArrayBW) + "," + itsArrayBW;

            for (int i = 0; i < numGroundSegs; i++)
            {
                irrFront += "," + frontGroundGHI[i];
                irrRear += "," + rearGroundGHI[i];

                skyViewAll += "," + rearSkyViewFactors[i];
                skyViewOne += Environment.NewLine + i + "," + frontSkyViewFactors[i] + "," + rearSkyViewFactors[i];
            }

            // Profile of sky view factors received by ground
            File.AppendAllText("skyViewAll.csv", skyViewAll);
            File.WriteAllText("skyViewOne.csv", skyViewOne);

            // Profile of irradiance received by ground
            File.AppendAllText("irrFront.csv", irrFront);
            File.AppendAllText("irrRear.csv", irrRear);

            // Print details about the simulation geometry
            File.AppendAllText("setup.csv", setup);
        }
    }
}
