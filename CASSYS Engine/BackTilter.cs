// CASSYS - Grid connected PV system modelling software
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Back Tilter Class
//
// Revision History:
//
// Description:
// This class is responsible for the simulation of back side irradiance,
// whose components are direct, diffuse, front reflected, and ground reflected
// irradiance.
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
    class BackTilter
    {
        // Parameters for the back tilter class
        double itsArrayBW;                          // Array bandwidth [m]
        double itsClearance;                        // Array ground clearance [panel slope lengths]
        double itsPitch;                            // The distance between the rows [panel slope lengths]
        double itsPanelTilt;                        // The angle between the surface tilt of the module and the ground [radians]
        double itsBo;                               // ASHRAE Parameter used for IAM calculation [#]
        int numCellRows;                            // Number of cell rows on back of array [#]

        // Back tilter local variables/arrays and intermediate calculation variables and arrays
        double structLossFactor;                    // Percentage loss attributed to shading from back obstructions [#]
        double[] itsMonthlyAlbedo;                  // Array to store monthly albedo values [#]
        double[] backGlo;                           // Tilted global irradiance for each cell row on back of array [W/m2]
        double[] backDir;                           // Tilted beam irradiance for each cell row on back of array [W/m2]
        double[] backDif;                           // Tilted diffuse irradiance for each cell row on back of array [W/m2]
        double[] backFroRef;                        // Tilted front-panel-reflected irradiance for each cell row on back of array [W/m2]
        double[] backGroRef;                        // Tilted ground-reflected irradiance for each cell row on back of array [W/m2]
        public double IAMDir;                       // Incidence Angle Modifier for beam irradiance [#]
        public double IAMDif;                       // Incidence Angle Modifier for diffuse irradiance [#]
        public double IAMRef;                       // Incidence Angle Modifier for reflected irradiances [#]
        public double IrrInhomogeneity;             // The inhomogeneity of values that are present across the back of the array [W/m2]

        // Output variables
        public double Albedo;                       // Monthly albedo value [#]
        public double BGlo;                         // Effective back tilted global irradiance [W/m2]
        public double BDir;                         // Back tilted beam irradiance [W/m2]
        public double BDif;                         // Back tilted diffuse irradiance [W/m2]
        public double BFroRef;                      // Back tilted front-panel-reflected irradiance [W/m2]
        public double BGroRef;                      // Back tilted ground-reflected irradiance [W/m2]

        // Back Tilter constructor
        public BackTilter()
        {

        }

        // Config manages calculations and initializations that need only to be run once
        public void Config()
        {
            bool useBifacial = Convert.ToBoolean(ReadFarmSettings.GetInnerText("Bifacial", "UseBifacialModel", ErrLevel.FATAL));

            if (useBifacial)
            {
                switch (ReadFarmSettings.GetAttribute("O&S", "ArrayType", ErrLevel.FATAL))
                {
                    // In all cases, pitch must be normalized to panel slope lengths
                    case "Unlimited Rows":
                        itsPanelTilt = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PlaneTilt", ErrLevel.FATAL));
                        itsArrayBW = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "CollBandWidth", ErrLevel.FATAL));
                        itsPitch = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "Pitch", ErrLevel.FATAL)) / itsArrayBW;
                        itsClearance = Convert.ToDouble(ReadFarmSettings.GetInnerText("Bifacial", "GroundClearance", ErrLevel.FATAL)) / itsArrayBW;

                        // Find number of cell rows on back of array [#]
                        numCellRows = Util.NUM_CELLS_PANEL * int.Parse(ReadFarmSettings.GetInnerText("O&S", "StrInWid", ErrLevel.WARNING, _default: "1"));
                        break;
                    case "Single Axis Elevation Tracking (E-W)":
                        itsArrayBW = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "WActiveSAET", ErrLevel.FATAL));
                        itsPitch = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PitchSAET", ErrLevel.FATAL)) / itsArrayBW;

                        // Find number of cell rows on back of array [#]
                        numCellRows = Util.NUM_CELLS_PANEL * int.Parse(ReadFarmSettings.GetInnerText("O&S", "StrInWidSAET", ErrLevel.WARNING, _default: "1"));
                        break;
                    case "Single Axis Horizontal Tracking (N-S)":
                        itsArrayBW = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "WActiveSAST", ErrLevel.FATAL));
                        itsPitch = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PitchSAST", ErrLevel.FATAL)) / itsArrayBW;

                        // Find number of cell rows on back of array [#]
                        numCellRows = Util.NUM_CELLS_PANEL * int.Parse(ReadFarmSettings.GetInnerText("O&S", "StrInWidSAST", ErrLevel.WARNING, _default: "1"));
                        break;
                    default:
                        ErrorLogger.Log("Bifacial is not supported for the selected orientation and shading.", ErrLevel.FATAL);
                        break;
                }

                itsBo = Convert.ToDouble(ReadFarmSettings.GetInnerText("Losses", "IncidenceAngleModifier/bNaught", _Error: ErrLevel.WARNING, _default: "0.05"));
                structLossFactor = Convert.ToDouble(ReadFarmSettings.GetInnerText("Bifacial", "StructBlockingFactor", ErrLevel.FATAL));

                // Initialize arrays
                backGlo = new double[numCellRows];
                backDir = new double[numCellRows];
                backDif = new double[numCellRows];
                backFroRef = new double[numCellRows];
                backGroRef = new double[numCellRows];

                ConfigAlbedo();
            }
        }

        // Calculate manages calculations that need to be run for each time step
        public void Calculate
            (
              double PanelTilt                      // The angle between the surface tilt of the module and the ground [radians]
            , double Clearance                      // Array ground clearance [panel slope lengths]
            , double HDif                           // Diffuse horizontal irradiance [W/m2]
            , double TDifRef                        // Front reflected diffuse irradiance [W/m2]
            , double[] groundGHI                    // Sum of irradiance components for each of the interior ground segments [W/m2]
            , double backSH                         // Fraction of the back surface of the PV array that is unshaded [#]
            , int month                             // Current month, used for getting albedo value [#]
            , double TDir                           // Back tilted beam irradiance [W/m2]
            , DateTime ts                           // Time stamp analyzed, used for printing .csv files
            )
        {
            // For tracking systems, panel tilt and ground clearance will change at each time step
            itsPanelTilt = PanelTilt;
            itsClearance = Clearance / itsArrayBW;           // Convert to panel slope lengths

            int numGroundSegs = Util.NUM_GROUND_SEGS;

            double h = Math.Sin(itsPanelTilt);               // Vertical height of sloped PV panel [panel slope lengths]
            double b = Math.Cos(itsPanelTilt);               // Horizontal distance from front of panel to back of panel [panel slope lengths]

            // Calculate x, y coordinates of bottom and top edges of PV row behind the current PV row so that portions of sky and ground viewed by
            // the PV cell may be determined. Coordinates are relative to (0,0) being the ground point below the lower front edge of current PV row.
            // The row behind the current row is in the positive x direction.
            double bottomX = itsPitch;                       // x value for point on bottom edge of PV panel behind current row
            double bottomY = itsClearance;                   // y value for point on bottom edge of PV panel behind current row
            double topX = bottomX + b;                       // x value for point on top edge of PV panel behind current row
            double topY = bottomY + h;                       // y value for point on top edge of PV panel behind current row

            double aveGroundGHI = 0;                         // Ground irradiance received averaged over the row-to-row area
            for (int k = 0; k < numGroundSegs; k++)
            {
                aveGroundGHI += groundGHI[k] / numGroundSegs;
            }

            // Get albedo value for the month
            Albedo = itsMonthlyAlbedo[month];

            // Accumulate diffuse, reflected, and beam irradiance components for each cell row over its field of view of PI radians
            for (int i = 0; i < numCellRows; i++)
            {
                double cellX = b * (i + 0.5) / numCellRows;                                     // x value for location of cell
                double cellY = itsClearance + h * (i + 0.5) / numCellRows;                      // y value for location of cell
                double elevUp = Math.Atan((topY - cellY) / (topX - cellX));                     // Elevation angle from PV cell to top of PV panel
                double elevDown = Math.Atan((cellY - bottomY) / (bottomX - cellX));             // Elevation angle from PV cell to bottom of PV panel

                int stopSky = Convert.ToInt32((itsPanelTilt - elevUp) * Util.RTOD);             // Last whole degree in arc range that sees sky; first is 0 [degrees]
                int startGround = Convert.ToInt32((itsPanelTilt + elevDown) * Util.RTOD);       // First whole degree in arc range that sees ground; last is 180 [degrees]

                // Add sky diffuse component, corrected for AOI
                backDif[i] = RadiationProc.GetViewFactor(0, stopSky * Util.DTOR) * HDif * IAMDif;

                // Add front surface reflected component, corrected for AOI
                backFroRef[i] = RadiationProc.GetViewFactor(stopSky * Util.DTOR, startGround * Util.DTOR) * TDifRef * IAMRef;

                backGroRef[i] = 0;
                // Add ground reflected component, corrected for albedo and AOI
                for (int j = startGround; j < 180; j++)
                {
                    // Get start and ending elevations for this (j, j + 1) pair
                    double startElevDown = elevDown + (j - startGround) * Util.DTOR;
                    double stopElevDown = elevDown + (j + 1 - startGround) * Util.DTOR;

                    // Projection onto ground in positive x direction
                    double projX2 = cellX + cellY / Math.Tan(startElevDown);
                    double projX1 = cellX + cellY / Math.Tan(stopElevDown);

                    // Initialize and get actualGroundGHI value
                    double actualGroundGHI = 0;
                    if (Math.Abs(projX1 - projX2) > 0.99 * itsPitch)
                    {
                        // Use average GHI if projection approximates or exceeds the pitch
                        actualGroundGHI = aveGroundGHI;
                    }
                    else
                    {
                        // Normalize projections and multiply by n
                        projX1 = numGroundSegs * projX1 / itsPitch;
                        projX2 = numGroundSegs * projX2 / itsPitch;

                        // Shift array indices to be within interval [0, n)
                        while (projX1 < 0 || projX2 < 0)
                        {
                            projX1 += numGroundSegs;
                            projX2 += numGroundSegs;
                        }
                        projX1 %= numGroundSegs;
                        projX2 %= numGroundSegs;

                        // Determine indices (truncate values) for use with groundGHI arrays
                        int index1 = Convert.ToInt32(Math.Floor(projX1));
                        int index2 = Convert.ToInt32(Math.Floor(projX2));

                        if (index1 == index2)
                        {
                            // Use single value if projection falls within a single segment of ground
                            actualGroundGHI = groundGHI[index1];
                        }
                        else
                        {
                            // Sum the irradiances on the ground if the projection falls across multiple segments
                            for (int k = index1; k <= index2; k++)
                            {
                                if (k == index1)
                                {
                                    actualGroundGHI += groundGHI[k] * (k + 1.0 - projX1);
                                }
                                else if (k == index2)
                                {
                                    actualGroundGHI += groundGHI[k] * (projX2 - k);
                                }
                                else
                                {
                                    actualGroundGHI += groundGHI[k];
                                }
                            }
                            // Get irradiance on ground in the 1 degree field of view
                            actualGroundGHI /= projX2 - projX1;
                        }
                    }
                    backGroRef[i] += RadiationProc.GetViewFactor(j * Util.DTOR, (j + 1) * Util.DTOR) * actualGroundGHI * Albedo * IAMRef;
                }

                // Cell is fully shaded if > 1, fully unshaded if < 0, otherwise fractionally shaded
                double backShade = (1.0 - backSH) * numCellRows - i;
                backShade = Math.Min(backShade, 1.0);
                backShade = Math.Max(backShade, 0.0);

                // Add beam irradiance, corrected for back shading and AOI
                backDir[i] = TDir * (1.0 - backShade) * IAMDir;

                // Correct each component for structure shading losses
                backDif[i] *= (1.0 - structLossFactor);
                backFroRef[i] *= (1.0 - structLossFactor);
                backGroRef[i] *= (1.0 - structLossFactor);
                backDir[i] *= (1.0 - structLossFactor);

                // Sum all components to get global back irradiance
                backGlo[i] = backDif[i] + backFroRef[i] + backGroRef[i] + backDir[i];
            }

            BDif = 0;
            BFroRef = 0;
            BGroRef = 0;
            BDir = 0;

            double maxGlo = backGlo[0];
            double minGlo = backGlo[0];

            for (int i = 0; i < numCellRows; i++)
            {
                // Calculate mean irradiance components
                BDif += backDif[i] / numCellRows;
                BFroRef += backFroRef[i] / numCellRows;
                BGroRef += backGroRef[i] / numCellRows;
                BDir += backDir[i] / numCellRows;

                // Find the max and min global irradiance values
                maxGlo = Math.Max(maxGlo, backGlo[i]);
                minGlo = Math.Min(minGlo, backGlo[i]);
            }
            BGlo = BDif + BFroRef + BGroRef + BDir;

            // Calculate the homogeneity of values as the range normalized by the sum
            IrrInhomogeneity = (BGlo > 0) ? 1 - ((maxGlo - minGlo) / BGlo) : 0;

            // Option to print details of the model in .csv files (takes about 3 seconds)
            PrintModel(ts);
        }

        // Config the albedo value based on monthly/yearly values defined on file
        void ConfigAlbedo()
        {
            // Determine whether to use site albedo values, or inter-row specific albedo values
            string whichAlbedo;
            if (Convert.ToBoolean(ReadFarmSettings.GetInnerText("BifAlbedo", "UseBifAlb", ErrLevel.FATAL)))
            {
                whichAlbedo = "BifAlbedo";
            }
            else
            {
                whichAlbedo = "Albedo";
            }

            // Initializing the expected list
            itsMonthlyAlbedo = new double[13];
            if (ReadFarmSettings.GetAttribute(whichAlbedo, "Frequency", ErrLevel.WARNING) == "Monthly")
            {
                // Using the month number as the index, populate the albedo vales from each corresponding node
                itsMonthlyAlbedo[1] = double.Parse(ReadFarmSettings.GetInnerText(whichAlbedo, "Jan", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[2] = double.Parse(ReadFarmSettings.GetInnerText(whichAlbedo, "Feb", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[3] = double.Parse(ReadFarmSettings.GetInnerText(whichAlbedo, "Mar", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[4] = double.Parse(ReadFarmSettings.GetInnerText(whichAlbedo, "Apr", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[5] = double.Parse(ReadFarmSettings.GetInnerText(whichAlbedo, "May", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[6] = double.Parse(ReadFarmSettings.GetInnerText(whichAlbedo, "Jun", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[7] = double.Parse(ReadFarmSettings.GetInnerText(whichAlbedo, "Jul", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[8] = double.Parse(ReadFarmSettings.GetInnerText(whichAlbedo, "Aug", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[9] = double.Parse(ReadFarmSettings.GetInnerText(whichAlbedo, "Sep", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[10] = double.Parse(ReadFarmSettings.GetInnerText(whichAlbedo, "Oct", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[11] = double.Parse(ReadFarmSettings.GetInnerText(whichAlbedo, "Nov", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[12] = double.Parse(ReadFarmSettings.GetInnerText(whichAlbedo, "Dec", ErrLevel.WARNING, _default: "0.2"));
            }
            else if (ReadFarmSettings.GetAttribute(whichAlbedo, "Frequency", ErrLevel.WARNING) == "Yearly")
            {
                itsMonthlyAlbedo[1] = Convert.ToDouble(ReadFarmSettings.GetInnerText(whichAlbedo, "Yearly", ErrLevel.WARNING, _default: "0.2"));

                // Apply the same albedo to all months
                for (int i = 3; i < itsMonthlyAlbedo.Length + 1; i++)
                {
                    itsMonthlyAlbedo[i - 1] = itsMonthlyAlbedo[1];
                }
            }
        }

        void PrintModel
            (
              DateTime ts                                       // Time stamp analyzed
            )
        {
            string irrBackSide = Environment.NewLine + ts;
            for (int i = 0; i < numCellRows; i++)
            {
                irrBackSide += "," + backDif[i] + "," + backFroRef[i] + "," + backGroRef[i] + "," + backDir[i] + "," + backGlo[i];
            }
            File.AppendAllText("irrBackSide.csv", irrBackSide);
        }
    }
}
