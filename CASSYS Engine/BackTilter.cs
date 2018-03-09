﻿// CASSYS - Grid connected PV system modelling software
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
        int numCellRows;                            // Number of cell rows on back of array [#]

        double[] backGlo;                           // Back tilted global irradiance for each cell row on back of array [W/m2]
        double[] backDir;                           // Back tilted beam irradiance for each cell row on back of array [W/m2]
        double[] backDif;                           // Back tilted diffuse irradiance for each cell row on back of array [W/m2]
        double[] backFroRef;                        // Back tilted front-panel-reflected irradiance for each cell row on back of array [W/m2]
        double[] backGroRef;                        // Back tilted ground-reflected irradiance for each cell row on back of array [W/m2]

        double aveGroundGHI;

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
              int cellRows                          // Number of cell rows on back of array [#]
            )
        {
            numCellRows = cellRows;

            // Initialize arrays
            backGlo = new double[numCellRows];
            backDir = new double[numCellRows];
            backDif = new double[numCellRows];
            backFroRef = new double[numCellRows];
            backGroRef = new double[numCellRows];
        }

        // Calculate manages calculations that need to be run for each time step
        public void Calculate
            (
              double PanelTilt                      // The angle between the surface tilt of the module and the ground [radians]
            , double Pitch                          // The distance between the rows [panel slope lengths]
            , double Clearance                      // Array ground clearance [panel slope lengths]
            , double HDif                           // Diffuse horizontal irradiance [W/m2]
            , double TGloEff                        // Front tilted effective irradiance [W/m2]
            , double Bo                             // ASHRAE Parameter used for IAM calculation [#]
            , double[] midGroundGHI                 // Sum of irradiance components for each of the ground segments in the middle PV rows [W/m2]
            , double midBackSH                      // Fraction of the back surface of the PV panel that is shaded [#]
            , double midFrontSH                     // Fraction of the front surface of the PV panel that is shaded [#]
            , int numGroundSegs                     // Number of segments into which to divide up the ground [#]
            , double albedo                         // Albedo value for the current month [#]
            , double InciAng                        // Angle of incidence for the beam to back of panel [radians]
            , double TBackDir                       // Back tilted beam irradiance [W/m2]
            , DateTime ts                           // Time stamp analyzed, used for printing .csv files
            )
        {
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

            double configFactor = 0;

            // Assume front surface material is glass
            double refractionIndex = 1.526;
            // Reflectance at normal incidence (Duffie and Beckman, p.217)
            double reflectFactor = Math.Pow((refractionIndex - 1.0) / (refractionIndex + 1.0), 2.0);

            aveGroundGHI = 0;
            for (int k = 0; k < numGroundSegs; k++)
            {
                aveGroundGHI += midGroundGHI[k] / numGroundSegs;
            }
            aveGroundGHI *= albedo;

            //double actualGroundGHI = 0;
            // Calculate diffuse, reflected, and beam irradiance components for each cell row over its field of view of PI radians
            for (int i = 0; i < numCellRows; i++)
            {
                double cellX = b * (i + 0.5) / numCellRows;                                     // x value for location of cell
                double cellY = Clearance + h * (i + 0.5) / numCellRows;                         // y value for location of cell
                double elevUp = Math.Atan((topY - cellY) / (topX - cellX));                     // Elevation angle from PV cell to top of PV panel
                double elevDown = Math.Atan((cellY - bottomY) / (bottomX - cellX));             // Elevation angle from PV cell to bottom of PV panel

                int stopSky = Convert.ToInt32((PanelTilt - elevUp) * Util.RTOD);                // Last whole degree in arc range that sees sky; first is 0 [degrees]
                int startGround = Convert.ToInt32((PanelTilt + elevDown) * Util.RTOD);          // First whole degree in arc range that sees ground; last is 180 [degrees]

                backDif[i] = 0;
                backFroRef[i] = 0;
                backGroRef[i] = 0;
                backDir[i] = 0;

                // Add sky diffuse component
                for (int j = 0; j < stopSky; j++)
                {
                    // Get Incidence Angle modifier for current field of view
                    double IAMDif = Tilt.GetASHRAEIAM(Bo, j * Util.DTOR);

                    configFactor = 0.5 * (Math.Cos(j * Util.DTOR) - Math.Cos((j + 1) * Util.DTOR));
                    backDif[i] += configFactor * HDif * IAMDif;
                }

                // Add front surface reflected component
                for (int j = stopSky; j < startGround; j++)
                {
                    // Calculate reflected irradiance from PV module
                    // EC: not fully confident in how midFrontSH is factored in... does not account for WHERE the shading occurs on the panel
                    double pvReflected = TGloEff * (1.0 - reflectFactor) * (1.0 - midFrontSH);

                    // Get Incidence Angle modifier for current field of view
                    double IAMFroRef = Tilt.GetASHRAEIAM(Bo, j * Util.DTOR);

                    configFactor = 0.5 * (Math.Cos(j * Util.DTOR) - Math.Cos((j + 1) * Util.DTOR));
                    backFroRef[i] += configFactor * pvReflected * IAMFroRef;
                }

                // Add ground reflected component
                // Calculate and summarize ground configuration factors ahead, below, and behind the cell.
                // Directions are split into three so that view can independently extend backward and forward.
                backGroRef[i] += CalcGroundConfigDirection(-1, Pitch, PanelTilt, cellX, cellY, midGroundGHI, Bo, albedo, numGroundSegs);
                backGroRef[i] += CalcGroundConfigDirection(0, Pitch, PanelTilt, cellX, cellY, midGroundGHI, Bo, albedo, numGroundSegs);
                backGroRef[i] += CalcGroundConfigDirection(1, Pitch, PanelTilt, cellX, cellY, midGroundGHI, Bo, albedo, numGroundSegs);

                //for (int j = startGround; j < 180; j++)
                //{
                //    // Get start and ending elevations for this (j, j + 1) pair
                //    double startElevDown = elevDown + (j - startGround) * Util.DTOR;
                //    double stopElevDown = elevDown + (j + 1 - startGround) * Util.DTOR;
                //    // Projection onto ground in positive x direction
                //    double projX2 = cellX + cellY / Math.Tan(startElevDown);
                //    double projX1 = cellX + cellY / Math.Tan(stopElevDown);

                //    // Initialize and get actualGroundGHI value
                //    actualGroundGHI = 0;
                //    if (Math.Abs(projX1 - projX2) > 0.99 * Pitch)
                //    {
                //        // Use average GHI if projection approximates the pitch
                //        for (int k = 0; k < numGroundSegs; k++)
                //        {
                //            actualGroundGHI += midGroundGHI[k] / numGroundSegs;
                //        }
                //    }
                //    else
                //    {
                //        // Normalize projections and multiply by n
                //        projX1 = numGroundSegs * projX1 / Pitch;
                //        projX2 = numGroundSegs * projX2 / Pitch;

                //        // Shift array indices to be within interval [0, n)
                //        while (projX1 < 0 || projX2 < 0)
                //        {
                //            projX1 += numGroundSegs;
                //            projX2 += numGroundSegs;
                //        }
                //        projX1 %= numGroundSegs;
                //        projX2 %= numGroundSegs;

                //        // Determine indices (truncate values) for use with groundGHI arrays
                //        int index1 = Convert.ToInt32(Math.Floor(projX1));
                //        int index2 = Convert.ToInt32(Math.Floor(projX2));

                //        if (index1 == index2)
                //        {
                //            // Use single value if projection falls within a single segment of ground
                //            actualGroundGHI = midGroundGHI[index1];
                //        }
                //        else
                //        {
                //            // Sum the irradiances on the ground if the projection falls across multiple segments
                //            for (int k = index1; k <= index2; k++)
                //            {
                //                if (k == index1)
                //                {
                //                    actualGroundGHI += midGroundGHI[k] * (k + 1.0 - projX1);
                //                }
                //                else if (k == index2)
                //                {
                //                    actualGroundGHI += midGroundGHI[k] * (projX2 - k);
                //                }
                //                else
                //                {
                //                    actualGroundGHI += midGroundGHI[k];
                //                }
                //            }
                //            // Get irradiance on ground in the 1 degree field of view
                //            actualGroundGHI /= projX2 - projX1;
                //        }
                //    }
                //    double IAMGroRef = Math.Cos(j * Util.DTOR) > Bo / (1 + Bo) ? Math.Max((1 - Bo * (1 / Math.Cos(j * Util.DTOR) - 1)), 0) : 0;
                //    configFactor = 0.5 * (Math.Cos(j * Util.DTOR) - Math.Cos((j + 1) * Util.DTOR));
                //    backGroRef[i] += configFactor * actualGroundGHI * albedo * IAMGroRef;
                //}

                // Cell is fully shaded if > 1, fully unshaded if < 0, otherwise fractionally shaded
                double backShade = midBackSH * numCellRows - i;
                backShade = Math.Min(backShade, 1.0);
                backShade = Math.Max(backShade, 0.0);

                // Add beam irradiance, corrected for shading and AOI
                if (InciAng < (Math.PI / 2))
                {
                    // Get Incidence Angle modifier for current field of view
                    double IAMDir = Tilt.GetASHRAEIAM(Bo, InciAng);

                    backDir[i] += (1.0 - backShade) * TBackDir * IAMDir;
                }

                // Sum all components to get global back irradiance
                backGlo[i] = backDif[i] + backFroRef[i] + backGroRef[i] + backDir[i];
            }

            BDif = 0;
            BFroRef = 0;
            BGroRef = 0;
            BDir = 0;
            // Calculate mean irradiance components
            // EC: divide by numCellRows?
            for (int i = 0; i < numCellRows; i++)
            {
                BDif += backDif[i] / numCellRows;
                BFroRef += backFroRef[i] / numCellRows;
                BGroRef += backGroRef[i] / numCellRows;
                BDir += backDir[i] / numCellRows;
            }
            BGlo = BDif + BFroRef + BGroRef + BDir;
            // Option to print details of the model in .csv files (takes about 3 seconds)
            PrintModel(ts);
        }

        double CalcGroundConfigDirection
            (
              int direction                         // The direction in which to move along the x-axis [-1, 0, 1]
            , double Pitch                          // The distance between the rows [panel slope lengths]
            , double PanelTilt                      // The angle between the surface tilt of the module and the ground [radians]
            , double cellX                          // x value for location of cell
            , double cellY                          // y value for location of cell
            , double[] midGroundGHI                 // Global irradiance for each of the ground segments in the middle PV rows [W/m2]
            , double Bo                             // ASHRAE Parameter used for IAM calculation [#]
            , double albedo                         // Albedo value for the current month [#]
            , int numGroundSegs                     // Number of segments into which to divide up the ground [#]
            )
        {
            // Divide the row-to-row spacing into n intervals
            double delta = Pitch / numGroundSegs;

            double theta1 = 0;
            double theta2 = 0;
            double beta1 = 0;
            double beta2 = 0;

            int offset = direction;                             // Initialize offset to begin at first unit of given direction
            double groundPatch = 0;                             // Configuration factor for view of ground in single segment
            double groundSum = 0;                               // Configuration factor for all ground views in given direction

            // Sum ground configuration factors until ground is out of sight (from ahead) or ??? (from behind)
            // Only loop the calculation for rows extending forward or backward, so break loop when direction = 0.
            do
            {
                double x1 = offset * delta;                     // x value for start of ground segment
                double x2 = (offset + 1) * delta;               // x value for end of ground segment

                theta1 = Math.Atan(cellY / (x1 - cellX));       // Elevation angle from the start of the ground segment to the cell
                theta2 = Math.Atan(cellY / (x2 - cellX));       // Elevation angle from the end of the ground segment to the cell
                if (theta1 < 0)
                {
                    theta1 = Math.PI + theta1;
                }
                if (theta2 < 0)
                {
                    theta2 = Math.PI + theta2;
                }
                // Calculate the roughly equivalent ("j", "j + 1") field of view pair = (beta1, beta2)
                beta1 = theta1 + PanelTilt;                     // End of cell's field of view that sees the ground segment
                beta2 = theta2 + PanelTilt;                     // Start of cell's field of view that sees the ground segment

                // Determine corresponding value within range [0, numGroundSegments) with which to index array
                int index = offset;
                while (index < 0)
                {
                    index += numGroundSegs;
                }
                index %= numGroundSegs;

                // Get Incidence Angle modifier for current field of view
                double IAMGroRef = Tilt.GetASHRAEIAM(Bo, beta2);

                double configFactor = 0.5 * (Math.Cos(beta2) - Math.Cos(beta1));
                groundPatch = configFactor * midGroundGHI[index] * albedo * IAMGroRef;// * delta;

                groundSum += groundPatch;
                offset += direction;
            } while (offset != 0 && beta1 < Math.PI && (beta1 - beta2 > 0.001));//(Math.Abs(offset) < numGroundSegs * 2));// && groundPatch > (0.001 * groundSum));
            //Console.Write(offset + "\t");

            return groundSum;
        }

        void PrintModel
            (
              DateTime ts                                       // Time stamp analyzed
            )
        {
            string irrBackSide = Environment.NewLine + ts;
            string viewFactor = Environment.NewLine + ts + "," + aveGroundGHI + "," + BGroRef;
            for (int i = 0; i < numCellRows; i++)
            {
                irrBackSide += "," + backDif[i] + "," + backFroRef[i] + "," + backGroRef[i] + "," + backDir[i] + "," + backGlo[i];
            }
            File.AppendAllText("irrBackSide.csv", irrBackSide);
            File.AppendAllText("viewFactor.csv", viewFactor);
        }
    }
}
