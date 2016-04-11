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
// NB - 2016-02-24: Addition of backtracking options for horizontal axis cases
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
// Tracking and back-tracking - Lorenzo, Narvarte, and Munoz
//          Progress in Photovoltaics: Research and Applications, Vol. 19,
//          Issue 6, pp. 747-753, 2011
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
        double itsAzimuthRef;
        public Boolean useBackTracking;
        public double itsTrackerPitch;
        public double itsTrackerBW;
       


        // Output Variables
        public double SurfSlope;
        public double SurfAzimuth;
        public double IncidenceAngle;
        public double RotAngle;
        public double AngleCorrection;
       


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
                        if (itsTrackerAzimuth == Math.PI / 2 || itsTrackerAzimuth == -Math.PI/2)
                        {
                            // NB: If the user inputs a minimum tilt less than 0, the tracker is able to face the non-dominant direction, so the surface azimuth will change based on the sun azimuth.
                            // However, if the minimum tilt is greater than zero, the tracker can only face the dominant direction.
                            if (itsMinTilt <= 0)
                            {
                                // TODO: simplify surface azimuth
                                // NB: Math.Abs is used so that the surface azimuth is set to 0 degrees if the sun azimuth is between -90 and 90, and set to 180 degrees if the sun azimuth is between -180 and -90 or between 90 and 180
                                if (Math.Abs(SunAzimuth) >= Math.Abs(itsTrackerAzimuth))
                                {
                                    SurfAzimuth = Math.PI;
                                }
                                else
                                {
                                    SurfAzimuth = 0;
                                }
                            }

                            else
                            {
                                SurfAzimuth = itsTrackerAzimuth - Math.PI / 2;
                            }
                        }
                        else if (itsTrackerAzimuth == 0)
                        {
                            // TODO: simplify surface azimuth equations
                            // For north-south tracking, the sign of the sun-azimuth is checked against the tracker azimuth
                            // This is from Duffie and Beckman Page 22.
                            if (SunAzimuth >= itsTrackerAzimuth)
                            {
                                SurfAzimuth = Math.PI / 2;                            
                            }
                            else
                            {
                                SurfAzimuth = -Math.PI / 2;
                            }
                        }

                        // TODO: see if atan2() can be used
                        SurfSlope = Math.Atan2(Math.Sin(SunZenith) * Math.Cos(SurfAzimuth - SunAzimuth),Math.Cos(SunZenith));

                        // NB: to put surface slope into the correct quadrant
                        // Correction should not be needed if atan2() works in statement above
                        //if (SurfSlope < 0.0)
                        //{
                        //    SurfSlope += Math.PI;
                        //}

 
                        // If the shadow is greater than the Pitch and backtracking is selected
                        if (useBackTracking)
                        {
                            if (itsTrackerBW / (Math.Cos(SurfSlope)) > itsTrackerPitch)
                            {
                                // NB: From Lorenzo, Narvarte, and Munoz
                                AngleCorrection = Math.Acos((itsTrackerPitch * (Math.Cos(SurfSlope))) / itsTrackerBW);
                                SurfSlope = SurfSlope - AngleCorrection;
                            }
                        }

                        // NB: Adjusting limits for elevation tracking, so if positive min tilt, the tracker operates within limits properly
                        if (itsTrackerAzimuth == Math.PI / 2 || itsTrackerAzimuth == -Math.PI / 2)
                        {
                            if (itsMinTilt <= 0)
                            {
                                if (Math.Abs(SunAzimuth) <= itsTrackerAzimuth)
                                {
                                    SurfSlope = Math.Min(itsMaxTilt, SurfSlope);
                                }
                                else if (Math.Abs(SunAzimuth) > itsTrackerAzimuth)
                                {
                                    SurfSlope = Math.Min(Math.Abs(itsMinTilt), SurfSlope);
                                }
                            }

                            else if (itsMinTilt > 0)
                            {
                                SurfSlope = Math.Min(SurfSlope, itsMaxTilt);
                                SurfSlope = Math.Max(SurfSlope, itsMinTilt);
                            }
                        }

                        else if (itsTrackerAzimuth == 0)
                        {
                            SurfSlope = Math.Min(itsMaxTilt, SurfSlope);
                        }


                       
                    }
                    else
                    {
                        // Tilt and Roll.
                        double aux = Tilt.GetCosIncidenceAngle(SunZenith, SunAzimuth, itsTrackerSlope, itsTrackerAzimuth);
                        // Equation (7) from Marion and Dobos
                        RotAngle =  Math.Atan2((Math.Sin(SunZenith) * Math.Sin(SunAzimuth - itsTrackerAzimuth)), aux);

                        //NB: enforcing rotation limits on tracker
                        RotAngle = Math.Min(itsMaxRotationAngle, RotAngle);
                        RotAngle = Math.Max(itsMinRotationAngle, RotAngle);


                        // Slope from equation (1) in Marion and Dobos
                        SurfSlope = Math.Acos(Math.Cos(RotAngle) * Math.Cos(itsTrackerSlope));

                        // Surface Azimuth from NREL paper
                        if (SurfSlope != 0)
                        {
                            // Equation (3) in Marion and Dobos
                            if ((-Math.PI <= RotAngle) && (RotAngle < -Math.PI/2))
                            {
                                SurfAzimuth = itsTrackerAzimuth - Math.Asin(Math.Sin(RotAngle) / Math.Sin(SurfSlope)) - Math.PI;
                            }
                            // Equation (4) in Marion and Dobos
                            else if ((Math.PI / 2 < RotAngle) && (RotAngle <= Math.PI))
                            {
                                SurfAzimuth = itsTrackerAzimuth - Math.Asin(Math.Sin(RotAngle) / Math.Sin(SurfSlope)) + Math.PI;
                            }
                            // Equation (2) in Marion and Dobos
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
                    // Defining the surface slope
                    SurfSlope = SunZenith;
                    SurfSlope = Math.Max(itsMinTilt, SurfSlope);
                    SurfSlope = Math.Min(itsMaxTilt, SurfSlope);

                    // Defining the surface azimuth
                    SurfAzimuth = SunAzimuth;
                    
                    // Changes the reference frame to be with respect to the reference azimuth
                    if (SurfAzimuth >= 0)
                    {
                        SurfAzimuth -= itsAzimuthRef;
                    }
                    else
                    {
                        SurfAzimuth += itsAzimuthRef;
                    }

                    // Enforcing the rotation limits with respect to the reference azimuth
                    SurfAzimuth = Math.Max(itsMinAzimuth, SurfAzimuth);
                    SurfAzimuth = Math.Min(itsMaxAzimuth, SurfAzimuth);

                    // Moving the surface azimuth back into the azimuth variable convention
                    if (SurfAzimuth >= 0)
                    {
                        SurfAzimuth -= itsAzimuthRef;
                    }

                    else
                    {
                        SurfAzimuth += itsAzimuthRef;
                    }
                    break;

                // Azimuth Vertical Axis Tracking
                case TrackMode.AVAT:
                    // Slope is constant.
                    // Defining the surface azimuth
                    SurfAzimuth = SunAzimuth;


                    // Changes the reference frame to be with respect to the reference azimuth
                    if (SurfAzimuth >= 0)
                    {
                        SurfAzimuth -= itsAzimuthRef;
                    }
                    else
                    {
                        SurfAzimuth += itsAzimuthRef;
                    }

                    // Enforcing the rotation limits with respect to the reference azimuth
                    SurfAzimuth = Math.Max(itsMinAzimuth, SurfAzimuth);
                    SurfAzimuth = Math.Min(itsMaxAzimuth, SurfAzimuth);

                    // Moving the surface azimuth back into the azimuth variable convention
                    if (SurfAzimuth >= 0)
                    {
                        SurfAzimuth -= itsAzimuthRef;
                    }

                    else
                    {
                        SurfAzimuth += itsAzimuthRef;
                    }

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
                    if (ReadFarmSettings.CASSYSCSYXVersion == "0.9.3")
                    {
                        SurfSlope = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PlaneTiltFix", ErrLevel.FATAL));
                        SurfAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "AzimuthFix", ErrLevel.FATAL));
                    }
                    else
                    {
                        SurfSlope = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PlaneTilt", ErrLevel.FATAL));
                        SurfAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "Azimuth", ErrLevel.FATAL));
                    }
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

                    
                    // Backtracking Options
                    useBackTracking = Convert.ToBoolean(ReadFarmSettings.GetInnerText("O&S", "BacktrackOptSAET", ErrLevel.WARNING, _default: "false"));
                    if (useBackTracking)
                    {
                        itsTrackerPitch = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PitchSAET", ErrLevel.FATAL));
                        itsTrackerBW = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "WActiveSAET", ErrLevel.FATAL));
                    }
                    break;

                case "Single Axis Horizontal Tracking (N-S)":
                    // Tracker Parameters
                    itsTrackMode = TrackMode.SAXT;
                    itsTrackerSlope = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "AxisTiltSAST", ErrLevel.FATAL));
                    itsTrackerAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "AxisAzimuthSAST", ErrLevel.FATAL));

                    // Operational Limits
                    itsMaxTilt = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "RotationMaxSAST", ErrLevel.FATAL));

                    
                    // Backtracking Options
                    useBackTracking = Convert.ToBoolean(ReadFarmSettings.GetInnerText("O&S", "BacktrackOptSAST", ErrLevel.WARNING, _default: "false"));
                    if (useBackTracking)
                    {
                        itsTrackerPitch = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PitchSAST", ErrLevel.FATAL));
                        itsTrackerBW = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "WActiveSAST", ErrLevel.FATAL));
                        
                    }
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
                    itsAzimuthRef = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "AzimuthRefTAXT", ErrLevel.FATAL));
                    itsMinAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "MinAzimuthTAXT", ErrLevel.FATAL));
                    itsMaxAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "MaxAzimuthTAXT", ErrLevel.FATAL));
                    break;

                case "Azimuth (Vertical Axis) Tracking":
                    itsTrackMode = TrackMode.AVAT;
                    // Surface Parameters
                    SurfSlope = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PlaneTiltAVAT", ErrLevel.FATAL));
                    // Operational Limits
                    itsAzimuthRef = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "AzimuthRefAVAT", ErrLevel.FATAL));
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
