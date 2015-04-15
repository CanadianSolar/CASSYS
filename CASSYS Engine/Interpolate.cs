// CASSYS - Grid connected PV system modelling software 
// Version 0.9 
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: ReadFarmSettings
// 
// Revision History:
// AP - 2014-9-18: Version 0.9
//
// Inputs: Array of X, Y values
//       
// Outputs: Interpolated Y values based on requested X 
// 
// 
//                              
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
//
// 
///////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Text;

namespace CASSYS
{

    public class Interpolate
    {
        // LINEAR INTERPOLATION
        // Given two arrays (x and y) and a x value, compute a y value by linear interpolation. 
        // The interpolation is guaranteed to work only if the arrays are sorted
        public static double Linear(double[] x, double[] y, double xvalue)
        {
            // The number of points available for interpolation.
            int npts = x.Length;

            // Sort arrays based in ascending order 
            Array.Sort(x, y);

            // Declarations
            double dx, slope;
            int i, j;

            // Pathological case
            if (npts < 2)
                return y[0];

            // Linear extrapolation to the left 
            if (xvalue < x[0])
            {
                slope = (y[1] - y[0]) / (x[1] - x[0]);
                dx = xvalue - x[0];
                return (y[0] + dx * slope)/xvalue;
            }

            // Linear extrapolation to the right
            if (xvalue > x[npts - 1])
            {
                slope = (y[npts - 1] - y[npts - 2]) / (x[npts - 1] - x[npts - 2]);
                dx = xvalue - x[npts - 1];
                return (y[npts - 1] + dx * slope)/xvalue;
            }

            // Other cases: bracket spot
            i = 0;
            j = npts - 1;
            while (j > i + 1)
            {
                int k = (i + j) / 2;
                if (xvalue < x[k])
                    j = k;
                else
                    i = k;
            }

            // Interpolate 
            return (y[i] * (x[j] - xvalue) + y[j] * (xvalue - x[i])) / (xvalue * (x[j] - x[i]));
        }

        // QUADRATIC INTERPOLATION
        // This is a quick Lagrangian polynomial method of interpolating for the value of y for given x-value
        public static double Quadratic(double[] x, double[] y, double xvalue)
        {
            // The number of points available for interpolation.
            int npts = x.Length;

            // Sort array in ascending order
            Array.Sort(x, y);

            // Initialize the new value for Y
            double newY = 0;

            // Use Lagrange Polynomial to determine the new y-value
            newY = y[0] * (xvalue - x[1]) * (xvalue - x[2]) / ((x[0] - x[1]) * (x[0] - x[2]));
            newY += y[1] * (xvalue - x[0]) * (xvalue - x[2]) / ((x[1] - x[0]) * (x[1] - x[2]));
            newY += y[2] * (xvalue - x[0]) * (xvalue - x[1]) / ((x[2] - x[0]) * (x[2] - x[1]));

            return newY;
        }

    }
}