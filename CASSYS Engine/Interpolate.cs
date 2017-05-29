// CASSYS - Grid connected PV system modelling software 
// Version 0.9 
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: ReadFarmSettings
// 
// Revision History:
// AP - 2014-9-18: Version 0.9
// AP - 2015-04-22: Version 0.9.1 - Added the Bezier Interpolation Method
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
                return (y[0] + dx * slope);
            }

            // Linear extrapolation to the right
            if (xvalue > x[npts - 1])
            {
                slope = (y[npts - 1] - y[npts - 2]) / (x[npts - 1] - x[npts - 2]);
                dx = xvalue - x[npts - 1];
                return (y[npts - 1] + dx * slope);
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
            return (y[i] * (x[j] - xvalue) + y[j] * (xvalue - x[i])) / ((x[j] - x[i]));
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

        // BEZIER INTERPOLATION
        // This is a 4-point method of Bezier Interpolation 
        public static double Bezier(double[] xa, double[] ya, double x, int n)
        {

            //find interval x is in
            int pos = -1;
            for (int i = 0; i < n - 1; i++)
            {
                if (x >= xa[i] && x <= xa[i + 1])
                {
                    pos = i;
                }
            }
            if (pos == -1)
            {
                //can not extrapolate
                // ErrorLogger.Log("The Bezier Interpolation cannot algorithm tried to perform an extrapolation. CASSYS has ended.");
            }

            //evaluate on the given interval 
            Point pt;
            if (pos == 0)
            {
                pt = EvaluateBezier(xa, ya, 0, 0, 1, 2, -1, x);
            }
            else if (pos == n - 2)
            {
                pt = EvaluateBezier(xa, ya, pos - 1, pos, pos + 1, pos + 1, -1, x);
            }

            else
            {
                pt = EvaluateBezier(xa, ya, pos - 1, pos, pos + 1, pos + 2, -1, x);
            }

            //return Y value found
            return pt.y;
        }

        // Struct to hold X and Y values for a Point
        struct Point
        {
            public double x;
            public double y;
        }

        // XYDist calculates the Distance between two points
        private static double XYDist(Point a, Point b)
        {
            return Math.Sqrt((a.x - b.x) * (a.x - b.x) + (a.y - b.y) * (a.y - b.y));
        }

        // Bezier 4 Point Method
        private static Point Bezier4Point(Point p1, Point p2, Point p3, Point p4, double mu, double xv = 0)
        {
            double mum1, mum13, mu3;
            Point p;

            if (mu != -1)
            {
                mum1 = 1 - mu;
                mum13 = mum1 * mum1 * mum1;
                mu3 = mu * mu * mu;
                p.x = mum13 * p1.x + 3 * mu * mum1 * mum1 * p2.x + 3 * mu * mu * mum1 * p3.x + mu3 * p4.x;
            }
            // interpolate to x value
            else
            {
                //use Newtons method to approximate mu
                double high, low, fp;
                high = 1;
                low = 0;
                mu = (xv - p1.x) / (p4.x - p1.x);
                mum1 = 1 - mu;
                mum13 = mum1 * mum1 * mum1;
                mu3 = mu * mu * mu;
                p.x = mum13 * p1.x + 3 * mu * mum1 * mum1 * p2.x + 3 * mu * mu * mum1 * p3.x + mu3 * p4.x;
                while (Math.Abs(xv - p.x) > 0.01)
                {
                    fp = -3 * (p1.x - 3 * p2.x + 3 * p3.x - p4.x) * mu * mu + 6 * (p1.x - 2 * p2.x + p3.x) * mu - 3 * (p1.x - p2.x);

                    mu = (xv - p.x + fp * mu) / fp;
                    mum1 = 1 - mu;
                    mum13 = mum1 * mum1 * mum1;
                    mu3 = mu * mu * mu;
                    if (mu < 0 || mu > 1)
                    {
                        mu = (high + low) / 2;
                        mum1 = 1 - mu;
                        mum13 = mum1 * mum1 * mum1;
                        mu3 = mu * mu * mu;
                    }
                    p.x = mum13 * p1.x + 3 * mu * mum1 * mum1 * p2.x + 3 * mu * mu * mum1 * p3.x + mu3 * p4.x;

                    if (xv > p.x)
                        low = mu;
                    else
                        high = mu;

                }
            }
            //calculate interpolated Y
            p.y = mum13 * p1.y + 3 * mu * mum1 * mum1 * p2.y + 3 * mu * mu * mum1 * p3.y + mu3 * p4.y;

            return p;
        }

        // Evaluates the Bezier curve based on arrays of x and y values
        private static Point EvaluateBezier(double[] xarr, double[] yarr, int i0, int i1, int i2, int i3, double t, double xv = 0)
        {
            //create new points
            Point[] pts = new Point[4];
            Point[] bz = new Point[4];
            pts[0].x = xarr[i0];
            pts[0].y = yarr[i0];
            pts[1].x = xarr[i1];
            pts[1].y = yarr[i1];
            pts[2].x = xarr[i2];
            pts[2].y = yarr[i2];
            pts[3].x = xarr[i3];
            pts[3].y = yarr[i3];

            //determine distances between points
            double d01, d12, d23, d02, d13, f;
            d01 = XYDist(pts[0], pts[1]);
            d12 = XYDist(pts[1], pts[2]);
            d23 = XYDist(pts[2], pts[3]);
            d02 = XYDist(pts[0], pts[2]);
            d13 = XYDist(pts[1], pts[3]);

            bz[0] = pts[1];
            //determine distance case
            if ((d02 / 6 < d12 / 2) && (d13 / 6 < d12 / 2))
            {
                //'this is the normal case where both 1/6th vectors are less than half of d12
                if (i0 != i1)
                {
                    f = 1 / 6D;
                }
                else
                {
                    f = 1 / 3D;   //for endpoint intervals'
                }
                bz[1].x = pts[1].x + (pts[2].x - pts[0].x) * f;
                bz[1].y = pts[1].y + (pts[2].y - pts[0].y) * f;

                if (i2 != i3)
                {
                    f = 1 / 6D;
                }
                else
                {
                    f = 1 / 3D;   //for endpoint intervals
                }
                bz[2].x = pts[2].x + (pts[1].x - pts[3].x) * f;
                bz[2].y = pts[2].y + (pts[1].y - pts[3].y) * f;
            }
            else if ((d02 / 6 >= d12 / 2) && (d13 / 6 >= d12 / 2))
            {
                //this is the case where both 1/6th vectors are > than half of d12
                bz[1].x = pts[1].x + (pts[2].x - pts[0].x) * (d12 / 2 / d02);
                bz[1].y = pts[1].y + (pts[2].y - pts[0].y) * (d12 / 2 / d02);
                bz[2].x = pts[2].x + (pts[1].x - pts[3].x) * (d12 / 2 / d13);
                bz[2].y = pts[2].y + (pts[1].y - pts[3].y) * (d12 / 2 / d13);
            }
            else if (d02 / 6 >= d12 / 2)
            {
                //'for this case d02/6 is more than half of d12, so the d13/6 vector needs to be reduced
                bz[1].x = pts[1].x + (pts[2].x - pts[0].x) * (d12 / 2 / d02);
                bz[1].y = pts[1].y + (pts[2].y - pts[0].y) * (d12 / 2 / d02);
                bz[2].x = pts[2].x + (pts[1].x - pts[3].x) * (d12 / 2 / d13 * (d13 / d02));
                bz[2].y = pts[2].y + (pts[1].y - pts[3].y) * (d12 / 2 / d13 * (d13 / d02));
            }
            else
            {
                bz[1].x = pts[1].x + (pts[2].x - pts[0].x) * (d12 / 2 / d02 * (d02 / d13));
                bz[1].y = pts[1].y + (pts[2].y - pts[0].y) * (d12 / 2 / d02 * (d02 / d13));
                bz[2].x = pts[2].x + (pts[1].x - pts[3].x) * (d12 / 2 / d13);
                bz[2].y = pts[2].y + (pts[1].y - pts[3].y) * (d12 / 2 / d13);
            }
            bz[3] = pts[2];
            Point pt;

            //get interpolated point
            pt = Bezier4Point(bz[0], bz[1], bz[2], bz[3], t, xv);
            return pt;
        }
    }
}