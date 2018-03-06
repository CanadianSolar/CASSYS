// CASSYS - Grid connected PV system modelling software
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Tilter Class
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
// Code adapted from https://github.com/NREL/bifacialvf which is based on
// https://www.nrel.gov/docs/fy17osti/67847.pdf
///////////////////////////////////////////////////////////////////////////////

using System;
using System.IO;

namespace CASSYS
{
    class BackTilter
    {
        // Parameters for the back tilter class
        int cellRows;                               // Number of cell rows on back of array [#]

        double[] backGlo;                           // Back tilted global irradiance for each cell row on back of array [W/m2]
        double[] backDir;                           // Back tilted beam irradiance for each cell row on back of array [W/m2]
        double[] backDif;                           // Back tilted diffuse irradiance for each cell row on back of array [W/m2]
        double[] backFroRef;                        // Back tilted front-panel-reflected irradiance for each cell row on back of array [W/m2]
        double[] backGroRef;                        // Back tilted ground-reflected irradiance for each cell row on back of array [W/m2]

        // Output variables
        public double BGlo;                         // Back tilted global irradiance [W/m2]
        public double BDir;                         // Back tilted beam irradiance [W/m2]
        public double BDif;                         // Back tilted diffuse irradiance [W/m2]
        public double BFroRef;                      // Back tilted front-panel-reflected irradiance [W/m2]
        public double BGroRef;                      // Back tilted ground-reflected irradiance [W/m2]

        // Back Tilter constructor
        public BackTilter()
        {

        }

        // Config manages calculations and initializations that need only to be run once
        public void Config
            (
              int n                                 // Number of cell rows on back of array [#]
            )
        {
            cellRows = n;

            // Initialize arrays
            backGlo = new double[cellRows];
            backDir = new double[cellRows];
            backDif = new double[cellRows];
            backFroRef = new double[cellRows];
            backGroRef = new double[cellRows];
        }

        // Calculate manages calculations that need to be run for each time step
        public void Calculate
            (
              double PanelTilt                      // The angle between the surface tilt of the module and the ground [radians]
            , double Pitch                          // The distance between the rows [panel slope lengths]
            , double Clearance                      // Array ground clearance [panel slope lengths]
            , double HDif                           // Diffuse horizontal irradiance [W/m2]
            , double TGloEff                        // Front irradiance adjusted for incidence angle and soiling [W/m2]
            , double[] midGroundGHI                 // Sum of irradiance components for each of the n ground segments in the middle PV rows [W/m2]
            , double midBackSH                      // Fraction of the back surface of the PV panel that is shaded [#]
            , int numGroundSegs                     // Number of segments into which to divide up the ground [#]
            , double albedo                         // Albedo value for the current month [#]
            , double incAngle                       // Angle of incidence for the beam to back of panel [radians]
            , double TBackDir                       // Back tilted beam irradiance [W/m2]
            , DateTime ts                           // Time stamp analyzed, used for printing .csv files
            )
        {
            // Assume back surface material is glass
            // EC: used to index large array of correction values - might not need
            //double refractionIndex = 1.526;

            double h = Math.Sin(PanelTilt);                  // Vertical height of sloped PV panel [panel slope lengths]
            double b = Math.Cos(PanelTilt);                  // Horizontal distance from front of panel to back of panel [panel slope lengths]
            double d = Pitch - b;                            // Horizontal distance from back of one row to front of the next [panel slope lengths]

            // Calculate x, y coordinates of bottom and top edges of PV row behind the desired PV row so that portions of sky and ground viewed by
            // the PV cell may be determined. Coordinates are relative to (0,0) being the ground point below the lower front edge of desired PV row.
            // The row behind the desired row is in the positive x direction.
            double bottomX = Pitch;                          // x value for point on bottom edge of PV panel behind current row
            double bottomY = Clearance;                      // y value for point on bottom edge of PV panel behind current row
            double topX = bottomX + b;                       // x value for point on top edge of PV panel behind current row
            double topY = bottomY + h;                       // y value for point on top edge of PV panel behind current row

            double viewFactor = 0;
            double actualGroundGHI = 0;
            // Calculate diffuse, reflected, and beam irradiance components for each cell row over its field of view of PI radians
            for (int i = 0; i < cellRows; i++)
            {
                double cellX = b * (i + 0.5) / cellRows;                                        // x value for location of cell
                double cellY = Clearance + h * (i + 0.5) / cellRows;                            // y value for location of cell
                double elevUp = Math.Atan((topY - cellY) / (topX - cellX));                     // Elevation angle from PV cell to top of PV panel
                double elevDown = Math.Atan((cellY - bottomY) / (bottomX - cellX));             // Elevation angle from PV cell to bottom of PV panel
                //double L = (bottomX - cellX) / Math.Cos(elevDown);                              // Diagonal distance from PV cell to bottom of module in row behind

                // EC: right now it rounds up or down... should I floor stopSky and ceil startGround?
                int stopSky = Convert.ToInt32((PanelTilt - elevUp) * Util.RTOD);                // Last whole degree in arc range that sees sky; first is 0 [degrees]
                int startGround = Convert.ToInt32((PanelTilt + elevDown) * Util.RTOD);          // First whole degree in arc range that sees ground; last is 180 [degrees]
                //Console.Write("\n0-" + stopSky + ", " + stopSky + "-" + startGround + ", " + startGround + "-180");

                backDif[i] = 0;
                backFroRef[i] = 0;
                backGroRef[i] = 0;
                backDir[i] = 0;

                // Add sky diffuse component
                // EC: without AOI could be BDif += 0.5 * (Math.Cos(0) - Math.Cos(stopSky * Util.DTOR)) * HDif;
                for (int j = 0; j < stopSky; j++)
                {
                    // EC: need to properly calculate AOI
                    double AOIcorr = 1;
                    viewFactor = 0.5 * (Math.Cos(j * Util.DTOR) - Math.Cos((j + 1) * Util.DTOR));
                    backDif[i] += viewFactor * HDif * AOIcorr;
                }

                // Add front surface reflected component
                for (int j = stopSky; j < startGround; j++)
                {
                    //double startAlpha = elevUp + elevDown - (j - stopSky) * Util.DTOR;
                    //double stopAlpha = elevUp + elevDown - (j + 1 - stopSky) * Util.DTOR;

                    //double m = L * Math.Sin(startAlpha);
                    //double theta = Math.PI - elevDown - (Math.PI / 2.0 - startAlpha) - PanelTilt;
                    //projX2 = m / Math.Cos(theta);
                    //m = L * Math.Sin(stopAlpha);
                    //theta = Math.PI - elevDown - (Math.PI / 2.0 - stopAlpha) - PanelTilt;
                    //projX1 = m / Math.Cos(theta);
                    //projX1 = Math.Max(projX1, 0.0);

                    // Get reflected irradiance from PV module in the 1 degree field of view.
                    // TGloEff is already corrected for shading and AOI. So do we need to know midFrontSH???
                    //double pvReflected = ((projX2 - projX1) * TGloEff) / (projX2 - projX1); // * (1.0 - midFrontSH);
                    double pvReflected = TGloEff; // * (1.0 - midFrontSH);

                    // EC: need to properly calculate AOI
                    double AOIcorr = 1;
                    viewFactor = 0.5 * (Math.Cos(j * Util.DTOR) - Math.Cos((j + 1) * Util.DTOR));
                    backFroRef[i] += viewFactor * pvReflected * AOIcorr;
                }

                // Add ground reflected component
                for (int j = startGround; j < 180; j++)
                {
                    // Get start and ending elevations for this (j, j + 1) pair
                    double startElevDown = elevDown + (j - startGround) * Util.DTOR;
                    double stopElevDown = elevDown + (j + 1 - startGround) * Util.DTOR;
                    // Projection onto ground in positive x direction
                    double projX2 = cellX + cellY / Math.Tan(startElevDown);
                    double projX1 = cellX + cellY / Math.Tan(stopElevDown);

                    // Initialize and get actualGroundGHI value
                    actualGroundGHI = 0;
                    if (Math.Abs(projX1 - projX2) > 0.99 * Pitch)
                    {
                        // Use average GHI if projection approximates the pitch
                        for (int k = 0; k < numGroundSegs; k++)
                        {
                            actualGroundGHI += midGroundGHI[k] / numGroundSegs;
                        }
                    }
                    else
                    {
                        // Normalize projections and multiply by n
                        projX1 = numGroundSegs * projX1 / Pitch;
                        projX2 = numGroundSegs * projX2 / Pitch;

                        // Shift array indices to be within interval [0, n)
                        while (projX1 < 0 || projX2 < 0)
                        {
                            projX1 += numGroundSegs;
                            projX2 += numGroundSegs;
                        }
                        projX1 %= numGroundSegs;
                        projX2 %= numGroundSegs;
                        //Console.WriteLine("projX1 = " + projX1 + ", projX2 = " + projX2);

                        // Determine indices (truncate values) for use with groundGHI arrays
                        int index1 = Convert.ToInt32(Math.Floor(projX1));
                        int index2 = Convert.ToInt32(Math.Floor(projX2));
                        //Console.WriteLine("index1 = " + index1 + ", index2 = " + index2);

                        if (index1 == index2)
                        {
                            // Use single value if projection falls within a single segment of ground
                            actualGroundGHI = midGroundGHI[index1];
                        }
                        else
                        {
                            // Sum the irradiances on the ground if the projection falls across multiple segments
                            for (int k = index1; k <= index2; k++)
                            {
                                if (k == index1)
                                {
                                    actualGroundGHI += midGroundGHI[k] * (k + 1.0 - projX1);
                                }
                                else if (k == index2)
                                {
                                    actualGroundGHI += midGroundGHI[k] * (projX2 - k);
                                }
                                else
                                {
                                    actualGroundGHI += midGroundGHI[k];
                                }
                            }
                            // Get irradiance on ground in the 1 degree field of view
                            actualGroundGHI /= projX2 - projX1;
                        }
                    }
                    // EC: need to properly calculate AOI
                    double AOIcorr = 1;
                    viewFactor = 0.5 * (Math.Cos(j * Util.DTOR) - Math.Cos((j + 1) * Util.DTOR));
                    backGroRef[i] += viewFactor * actualGroundGHI * albedo * AOIcorr;
                }

                // Cell is fully shaded if > 1, fully unshaded if < 0, otherwise fractionally shaded
                double cellShade = midBackSH * cellRows - i;
                cellShade = Math.Min(cellShade, 1.0);
                cellShade = Math.Max(cellShade, 0.0);

                // Add beam irradiance, corrected for shading and AOI
                if (incAngle < (Math.PI / 2))
                {
                    // EC: need to properly calculate AOI
                    double AOIcorr = 1;
                    backDir[i] += (1.0 - cellShade) * TBackDir * AOIcorr;
                }

                // Sum all components to get global back irradiance
                backGlo[i] = backDif[i] + backFroRef[i] + backGroRef[i] + backDir[i];
            }

            BDif = 0;
            BFroRef = 0;
            BGroRef = 0;
            BDir = 0;
            // Calculate sums for irradiance components
            for (int i = 0; i < cellRows; i++)
            {
                BDif += backDif[i];
                BFroRef += backFroRef[i];
                BGroRef += backGroRef[i];
                BDir += backDir[i];
            }
            BGlo = BDif + BFroRef + BGroRef + BDir;
            // Option to print details of the model in .csv files (takes about 1 second)
            PrintModel(ts);
        }

        void PrintModel
            (
              DateTime ts                                       // Time stamp analyzed
            )
        {
            string irrBackSide = Environment.NewLine + ts;
            for (int i = 0; i < cellRows; i++)
            {
                irrBackSide += "," + backDif[i] + "," + backFroRef[i] + "," + backGroRef[i] + "," + backDir[i] + "," + backGlo[i];
            }
            File.AppendAllText("irrBackSide.csv", irrBackSide);
        }
    }
}
