// CASSYS - Grid connected PV system modelling software 
// Version 0.9.3  
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Tracker.cs
// 
// Revision History:
// AP - 2016-01-26: Version 0.9.3
// NB - 2016-02-17: Updated equations for Tilt and Roll tracker
//
// Description:
// This class is reponsible for the simulation of trackers, specifically the 
// following types:
//  1)	Single Axis Elevation Tracker (E-W, N-S axis), and Tilt and Roll Tracker - SAXT
//  2)	Azimuth or Vertical Axis Tracker - AVAT
//  3)	Two-Axis Tracking - TAXT
//  4)  No Axis Tracking - NOAT
//
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
// Solar geometry for fixed and tracking surfaces - Braun and Mitchell
//          Solar Energy, Vol. 31, No. 5, pp. 439-444, 1983
//
// Rotation Angle for the Optimum Tracking of One-Axis Trackers - Marion and Dobos
//          NREL, http://www.nrel.gov/docs/fy13osti/58891.pdf
//
///////////////////////////////////////////////////////////////////////////////
// Notes
///////////////////////////////////////////////////////////////////////////////
// Note 1: NB
// The 360 degree (2PI) correction for the surface azimuth of the tilt and roll
// tracker is due to a new set of equations that were used for the tilt and roll
// trackers. These equations were from the NREL paper. In this paper the azimuth
// was defined between 0 and 360 degrees clockwise from north, whereas CASSYS
// uses an azimuth input of 0 being true south and the angle is negative in the
// east and positive in the west, with +/-180 being north. The equations work
// well except for when the tracker azimuth is set to greater than the absolute
// value of 90 degrees, as in some circumstances the equations output values
// either greater than 180 degrees when the axis azimuth is greater than 90 or
// less than -180 degrees when the axis azimuth is less than -90. This is due to
// the difference between the way the paper defines azimuth and the way CASSYS
// defines azimuth. To correct this 360 degrees is either added to or subtracted
// from the azimuth if the azimuth is greater than 180 or less than -180 to put
// the given angle into a quadrant useable by CASSYS while keeping the angle
// equivalent to the angle output by the equation.
//
///////////////////////////////////////////////////////////////////////////////



using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace CASSYS
{
    public enum TrackMode { NOAT, SAXT, AVAT, TAXT }             // See description in header for expanded form.


    class Tracker
    {
        // Local variables
        // Tracker Variables (Made public to write to output file)
        TrackMode itsTrackMode;
        public double itsTrackerSlope;
        public double itsTrackerAzimuth;
        

        // Operational Limits (as they apply to the surface, typically)
        double itsMinTilt;
        double itsMaxTilt;
        double itsMinRotationAngle;
        double itsMaxRotationAngle;
        double itsMinAzimuth;
        double itsMaxAzimuth;

        // Output Variables
        public double SurfSlope;
        public double SurfAzimuth;
        public double IncidenceAngle;
        public double RotAngle;

        // Constructor for the tracker
        public Tracker()
        {
        }


        // Calculate the tracker slope, azimuth and incidence angle using
        public void Calculate(double SunZenith, double SunAzimuth)
        {
            switch (itsTrackMode)
            {
                case TrackMode.SAXT:
                    // Surface stays parallel to the ground.
                    if (itsTrackerSlope == 0.0)
                    {
                        // For east-west tracking, the absolute value of the sun-azimuth is checked against the tracker azimuth
                        // This is from Duffie and Beckman Page 22.
                        if (itsTrackerAzimuth == Math.PI / 2)
                        {
                            if (Math.Abs(SunAzimuth) >= itsTrackerAzimuth)
                            {
                                SurfAzimuth = itsTrackerAzimuth + Math.PI / 2;
                            }
                            else
                            {
                                SurfAzimuth = itsTrackerAzimuth - Math.PI / 2;
                            }
                        }
                        else if (itsTrackerAzimuth == 0)
                        {
                            // For north-south tracking, the sign of the sun-azimuth is checked against the tracker azimuth
                            // This is from Duffie and Beckman Page 22.
                            if (SunAzimuth >= itsTrackerAzimuth)
                            {
                                SurfAzimuth = itsTrackerAzimuth + Math.PI / 2;
                            }
                            else
                            {
                                SurfAzimuth = itsTrackerAzimuth - Math.PI / 2;
                            }
                        }

                        SurfSlope = Math.Atan(Math.Tan(SunZenith) * Math.Cos(SurfAzimuth - SunAzimuth));

                        if (SurfSlope < 0.0)
                        {
                            SurfSlope += Math.PI;
                        }
                    }
                    else
                    {
                        // Tilt and Roll.
                        double aux = Tilt.GetCosIncidenceAngle(SunZenith, SunAzimuth, itsTrackerSlope, itsTrackerAzimuth);
                        // Equation given by NREL paper
                        RotAngle =  Math.Atan2((Math.Sin(SunZenith) * Math.Sin(SunAzimuth - itsTrackerAzimuth)), aux);
                        
                        // Slope from NREL paper
                        SurfSlope = Math.Acos(Math.Cos(RotAngle) * Math.Cos(itsTrackerSlope));

                        // Surface Azimuth from NREL paper
                        if (SurfSlope != 0)
                        {
                            if ((-Math.PI <= RotAngle) && (RotAngle < -Math.PI/2))
                            {
                                SurfAzimuth = itsTrackerAzimuth - Math.Asin(Math.Sin(RotAngle) / Math.Sin(SurfSlope)) - Math.PI;
                            }
                            else if ((Math.PI / 2 < RotAngle) && (RotAngle <= Math.PI))
                            {
                                SurfAzimuth = itsTrackerAzimuth - Math.Asin(Math.Sin(RotAngle) / Math.Sin(SurfSlope)) + Math.PI;
                            }
                            else
                            {
                                SurfAzimuth = itsTrackerAzimuth + Math.Asin(Math.Sin(RotAngle) / Math.Sin(SurfSlope));
                            }
                        }
                        //NB: 360 degree correction to put Surface Azimuth into the correct quadrant, see Note 1
                        if (SurfAzimuth > Math.PI)
                        {
                            SurfAzimuth -= (Math.PI) * 2;
                        }
                        else if (SurfAzimuth < -Math.PI)
                        {
                            SurfAzimuth += (Math.PI) * 2;
                        }
                    }
                    break;

                // Two Axis Tracking
                case TrackMode.TAXT:
                    // Forcing the Slope to be within tilt limits
                    SurfSlope = SunZenith;
                    SurfSlope = Math.Max(itsMinTilt, SurfSlope);
                    SurfSlope = Math.Min(itsMaxTilt, SurfSlope);

                    // Forcing the Azimuth to be within the limits
                    SurfAzimuth = SunAzimuth;
                    SurfAzimuth = Math.Max(itsMinAzimuth, SurfAzimuth);
                    SurfAzimuth = Math.Min(itsMaxAzimuth, SurfAzimuth);
                    break;

                // Azimuth Vertical Axis Tracking
                case TrackMode.AVAT:
                    // Slope is constant.
                    // Forcing the Azimuth to be within the limits
                    SurfAzimuth = SunAzimuth;
                    SurfAzimuth = Math.Max(itsMinAzimuth, SurfAzimuth);
                    SurfAzimuth = Math.Min(itsMaxAzimuth, SurfAzimuth);
                    break;

                case TrackMode.NOAT:

                    break;

                // Throw error to user if there is an issue with the tracker.
                default:
                    ErrorLogger.Log("Tracking Parameters were incorrectly defined. Please check your input file.", ErrLevel.FATAL);
                    break;
            }

            IncidenceAngle = Tilt.GetIncidenceAngle(SunZenith, SunAzimuth, SurfSlope, SurfAzimuth);
        }

        // Gathering the tracker mode, and relevant operational limits, and tracking axis characteristics.
        public void Config()
        {
            switch (ReadFarmSettings.GetAttribute("O&S", "ArrayType", ErrLevel.FATAL))
            {
                case "Fixed Tilted Plane":
                    itsTrackMode = TrackMode.NOAT;
                    // Defining all the parameters for the shading of a unlimited row array configuration
                    SurfSlope = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PlaneTiltFix", ErrLevel.FATAL));
                    SurfAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "AzimuthFix", ErrLevel.FATAL));
                    break;

                case "Unlimited Rows":
                    itsTrackMode = TrackMode.NOAT;
                    // Defining all the parameters for the shading of a unlimited row array configuration
                    SurfSlope = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PlaneTilt", ErrLevel.FATAL));
                    SurfAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "Azimuth", ErrLevel.FATAL));
                    break;

                case "Single Axis Elevation Tracking (E-W)":
                    // Tracker Parameters
                    itsTrackMode = TrackMode.SAXT;
                    itsTrackerAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "AxisAzimuthSAET", ErrLevel.FATAL));
                    itsTrackerSlope = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "AxisTiltSAET", ErrLevel.FATAL));

                    // Operational Limits
                    itsMinTilt = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "MinTiltSAET", ErrLevel.FATAL));
                    itsMaxTilt = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "MaxTiltSAET", ErrLevel.FATAL));
                    break;

                case "Single Axis Horizontal Tracking (N-S)":
                    // Tracker Parameters
                    itsTrackMode = TrackMode.SAXT;
                    itsTrackerSlope = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "AxisTiltSAST", ErrLevel.FATAL));
                    itsTrackerAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "AxisAzimuthSAST", ErrLevel.FATAL));

                    // Operational Limits
                    itsMinRotationAngle = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "RotationMinSAST", ErrLevel.FATAL));
                    itsMaxRotationAngle = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "RotationMaxSAST", ErrLevel.FATAL));
                    break;

                
                case "Tilt and Roll Tracking":
                    // Tracker Parameters
                    itsTrackMode = TrackMode.SAXT;
                    itsTrackerSlope = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "AxisTiltTART", ErrLevel.FATAL));
                    itsTrackerAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "AxisAzimuthTART", ErrLevel.FATAL));

                    // Operational Limits
                    itsMinRotationAngle = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "RotationMinTART", ErrLevel.FATAL));
                    itsMaxRotationAngle = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "RotationMaxTART", ErrLevel.FATAL));
                    break;
                
                case "Two Axis Tracking":
                    itsTrackMode = TrackMode.TAXT;
                    // Operational Limits
                    itsMinTilt = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "MinTiltTAXT", ErrLevel.FATAL));
                    itsMaxTilt = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "MaxTiltTAXT", ErrLevel.FATAL));
                    itsMinAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "MinAzimuthTAXT", ErrLevel.FATAL));
                    itsMaxAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "MaxAzimuthTAXT", ErrLevel.FATAL));
                    break;

                case "Azimuth (Vertical Axis) Tracking":
                    itsTrackMode = TrackMode.AVAT;
                    // Surface Parameters
                    SurfSlope = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PlaneTiltAVAT", ErrLevel.FATAL));
                    // Operational Limits
                    itsMinAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "MinAzimuthAVAT", ErrLevel.FATAL));
                    itsMaxAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "MaxAzimuthAVAT", ErrLevel.FATAL));
                    break;

                default:
                    ErrorLogger.Log("No orientation and shading was specified by the user.", ErrLevel.FATAL);
                    break;

            }

        }
    }
}
