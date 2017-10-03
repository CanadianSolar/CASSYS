// CASSYS - Grid connected PV system modelling software 
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: HorizonShading.cs
//
// Revision History:
//
// Description:
// This class is responsible for the simulation of the shading effects of a 
// horizon or other far-off objects
// 
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
// Notes
///////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;

namespace CASSYS
{
    class HorizonShading
    {
        // Input variables
        string HorizonDefinitionAzimStr;                         // String containing the azimuthal Horizon Profile points from the .csyx file [degrees]
        string HorizonDefinitionElevStr;                         // String containing the elevation Horizon Profile points from the .csyz file [degrees]

        // Horizon Shading local variables/arrays and intermediate calculation variables and arrays
        double[] LimitingAngle;                                  // Array containing the limiting angles caused by the tilted solar array [radians]
        double[] HorizonAzim;                                    // Array containing the comma-separated azimuth values from HorizonDefinitionAzimStr [radians]
        double[] HorizonElev;                                    // Array containing the comma-separated elevation values from HorizonDefinitionElevStr [radians]
        double[] HorizonAzimExtended;                            // Extended version of HorizonAzim with added dummy values for interpolation [radians]
        double[] HorizonElevExtended;                            // Extended version of HorizonElev with added dummy values for interpolation [radians]
        double[] HorizonAzimInterpolated;                        // Array containing a full circle's worth of azimuthal values centred about PanelAzimuth [radians]
        double[] HorizonElevInterpolated;                        // Array containing the interpolated values of elevation horizon shading created from the Horizon Profile [radians]
        double[] RefLimitingAngle;                               // A dummy array containing limiting angle array for the ground, all values are PI/2 [radians]
        double RefDiffFactor;                                    // View factor for the ground 

        // Settings used in the class
        Boolean horizonDefined;                                  // Used to determine whether or not the user defined a horizon

        // Output variables
        public double DiffFactor;                                // Diffuse global factor caused by the far-shading of the sky [#]
        public double GRefFactor;                                // Ground reflected global factor caused by the far-shading of the sky [#]
        public double BeamFactor;                                // Beam factor caused by horizon. 1 or 0 based on whether sun is above or below horizon [#]
        public double TDir;                                      // Beam irradiance allowing for horizon effects [W/m^2]
        public double TDif;                                      // Diffuse irradiance allowing for horizon effects [W/m^2]
        public double TRef;                                      // Ground reflected irradiance allowing for horizon effects [W/m^2]
        public double TGlo;                                      // Total irradiance allowing for horizon effects [W/m^2]

        // Horizon shading constructor
        public HorizonShading()
        {
        }

        // Config takes values from the xml as well as manages calculations that need only to be run once
        public void Config
            (
              double PanelTiltFixed                              // The tilt of the panel, used in config if no tracking is selected [radians]
            , double PanelAzimFixed                              // The azimuth position of the panel, used in config if no tracking is selected [radians]
            , TrackMode ArrayTrackMode                           // The tracking mode, used to determine whether the diffuse fraction can be calculated in config
            )
        {

            // Loads the Horizon Profile from the .csyx document
            horizonDefined = Convert.ToBoolean(ReadFarmSettings.GetInnerText("O&S", "DefineHorizonProfile", ErrLevel.WARNING, _default: "false"));
            if (horizonDefined == true)
            {
                // Getting the horizon information from the .csyx file
                HorizonDefinitionAzimStr = ReadFarmSettings.GetInnerText("O&S", "HorizonAzi", ErrLevel.WARNING, "0.9", _default: "0");
                HorizonDefinitionElevStr = ReadFarmSettings.GetInnerText("O&S", "HorizonElev", ErrLevel.WARNING, "0.9", _default: "0");

                // Converts the Horizon Profile imported from the.csyx document into an array of doubles
                HorizonAzim = HorizonCSVStringtoArray(HorizonDefinitionAzimStr);
                HorizonElev = HorizonCSVStringtoArray(HorizonDefinitionElevStr);

                // If user inputs horizon azimuth/elevation data of two different lengths
                if (HorizonAzim.Length != HorizonElev.Length)
                {
                    ErrorLogger.Log("The number of horizon azimuth values was not equal to the number of horizon elevation values.", ErrLevel.FATAL);
                }

                // Extends the Horizon Profile by duplicating the first and last values and transposing them 360 degrees forward and backwards, respectively
                CalcExtendedHorizon(HorizonAzim, HorizonElev, out HorizonAzimExtended, out HorizonElevExtended);

                // The full azimuthal range of the Horizon Profile is filled in using Interpolation
                // The horizon profile must be calculated here both for the non-tracking case, as well as for the ground reflected diffuse factor for all tracking cases
                // It is calculated separately in the Calculate method for tracking cases
                HorizonAzimInterpolated = InitializeHorizonProfile(PanelAzimFixed);
                
                HorizonElevInterpolated = GetInterpolatedElevationProfile(HorizonAzimExtended, HorizonElevExtended, HorizonAzimInterpolated);

                // If there is no tracking implemented the diffuse horizon factor only needs to be calculated once, it is otherwise calculated in the Calculate method
                if (ArrayTrackMode == TrackMode.NOAT && ReadFarmSettings.UsePOA != true)
                {
                    // Calculates the limiting angle array
                    LimitingAngle = GetLimitingAngleArray(PanelTiltFixed, PanelAzimFixed, HorizonAzimInterpolated);

                    // The shading factor is calculated via numerical computation using the mathematical models described in the CASSYS documentation
                    DiffFactor = GetHorizonDiffuseFactor(PanelTiltFixed, PanelAzimFixed, HorizonAzimInterpolated, HorizonElevInterpolated, LimitingAngle);
                }

                // Creating the limiting angle array for the ground, all values are PI/2
                RefLimitingAngle = new double[361];
                for (int i = 0; i < RefLimitingAngle.Length; i++)
                {
                    RefLimitingAngle[i] = Math.PI / 2;
                }

                // Diffuse part of ground reflected factor only needs to be calculated once. 0s are used for azimuth and tilt of surface, as it represents the ground.
                RefDiffFactor = GetHorizonDiffuseFactor(0, 0, HorizonAzimInterpolated, HorizonElevInterpolated, RefLimitingAngle);
            }
        }
            
        // Calculate manages calculations that need to be run for each time step
        public void Calculate
            (
              double SunZenith                                  // The Zenith position of the sun with 0 being normal to the earth [radians]
            , double SunAzimuth                                 // The Azimuth position of the sun relative to 0 being true south. Positive if west, negative if east [radians]
            , double PanelTilt                                  // The angle between the surface tilt of the module and the ground [radians]
            , double PanelAzimuth                               // The azimuth direction in which the surface is facing. Positive if west, negative if east [radians]
            , double POABeam                                    // The amount of direct irradiance incident on the panel [W/m^2]
            , double POADiff                                    // The amount of diffuse irradiance incident on the panel [W/m^2]
            , double POAGRef                                    // The amount of ground reflected irradianc incident on the panel [W/m^2]
            , double HBeam                                      // The amount of direct irradiance incident on a horizontally oriented detector [W/m^2]
            , double HDiff                                      // The amount of diffuse irradiance incident on a horizontally oriented detector [W/m^2]
            , TrackMode ArrayTrackMode                          // The tracking mode, used to determine whether to calculate the diffuse factor in Config or Calculate
            )
        {
            // If the user inputs POA irradiance data, it is assumed to already account for horizon
            // If the horizon is either not defined or is 0, the horizon calculations do not need to be carried out
            // The horizon shaded irradiance factors are then assumed to be 1, therefore the horizon shaded irradiance is the same as that output by the tilter class
            if (ReadFarmSettings.UsePOA == true || horizonDefined != true)
            {
                TDir = POABeam;
                TDif = POADiff;
                TRef = POAGRef;
                TGlo = TDir + TDif + TRef;
            }
            else
            {
                // If there is no tracking, these calculations are taken care of in Config
                if (ArrayTrackMode != TrackMode.NOAT)
                {
                    // The full azimuthal range of the Horizon Profile is filled in using Interpolation
                    HorizonAzimInterpolated = InitializeHorizonProfile(PanelAzimuth);
                    HorizonElevInterpolated = GetInterpolatedElevationProfile(HorizonAzimExtended, HorizonElevExtended, HorizonAzimInterpolated);

                    // Calculates the limiting angle array
                    LimitingAngle = GetLimitingAngleArray(PanelTilt, PanelAzimuth, HorizonAzimInterpolated);

                    // The shading factor is calculated via numerical computation using the mathematical models described in the CASSYS documentation
                    DiffFactor = GetHorizonDiffuseFactor(PanelTilt, PanelAzimuth, HorizonAzimInterpolated, HorizonElevInterpolated, LimitingAngle);
                }

                BeamFactor = GetHorizonBeamFactor(SunAzimuth, SunZenith, HorizonAzimInterpolated, HorizonElevInterpolated, PanelAzimuth);

                GRefFactor = GetHorizonGRefFactor(BeamFactor, RefDiffFactor, HBeam, HDiff);

                // The horizon shaded irradiance values are those multiplied by the horizon shading factors
                TDir = BeamFactor * POABeam;
                TDif = DiffFactor * POADiff;
                TRef = GRefFactor * POAGRef;
                TGlo = TDir + TDif + TRef;
            }
        }

        // This function creates a 361-element length array that extends 180 degrees above and below the Panel Azimuth value
        public static double[] InitializeHorizonProfile(double PanelAzimuth)
        {
            double[] HorizonProfileFormat = new double[361];
            for (int Azi = -180; Azi < 181; Azi++)
            {
                HorizonProfileFormat[Azi + 180] = PanelAzimuth + Util.DTOR * Azi;
            }
            return HorizonProfileFormat;
        }

        // This function creates an array that holds the effective "shading" values that a panel casts upon itself
        public static double[] GetLimitingAngleArray
            (
              double PanelTilt                                   // The tilt of the panel surface [radians]
            , double PanelAzimuth                                // The azimuth position of the panel surface [radians]
            , double[] AzimuthAngleArray                         // The 361 element array of azimuth angles [radians]
            )
        {
            double[] LimitingAngleArray = new double[361];

            for (int i = 0; i < LimitingAngleArray.Length; i++)
            {
                // For azimuth values in front of the panel the limiting angle will always be PI/2
                if (i > 90 && i < 270)
                {
                    LimitingAngleArray[i] = Math.PI / 2;
                }
                else
                {
                    // Absolute value ensures the limiting angle is always positive
                    LimitingAngleArray[i] = Math.Abs(Math.Atan(1 / (Math.Tan(PanelTilt) * Math.Cos(PanelAzimuth - AzimuthAngleArray[i]))));
                }
            } 
            return LimitingAngleArray;
        }

        // Converts string array into an array of doubles that can be used by the program
        public static double[] HorizonCSVStringtoArray(string HorizonString)
        {
            string[] tempArray = HorizonString.Split(',');
            int arrayLength = tempArray.Length;
            double[] HorizonArray = new double[arrayLength];

            for (int arrayIndex = 0; arrayIndex < arrayLength; arrayIndex++)
            {
                HorizonArray[arrayIndex] = Util.DTOR * double.Parse(tempArray[arrayIndex]);
            }

            return HorizonArray;
        }

        // Adds azimuth value 2PI greater than the minimum azimuth and 2PI less than the maximum value, each with their respective elevations.
        // Necessary to get the interpolated elevation array
        static void CalcExtendedHorizon
            (
              double[] InputAzim                                   // The array of input horizon azimuth angles [radians]
            , double[] InputElev                                   // The array of input horizon elevation angles [radians]
            , out double[] ExtendedAzim                            // The array of horizon azimuth angles extended to complete the 360 degree azimuth profile [radians]
            , out double[] ExtendedElev                            // The array of horizon elevation angles extended to complete the profile [radians]
            )
        {
            ExtendedAzim = new double[InputAzim.Length + 2];
            ExtendedElev = new double[InputElev.Length + 2];

            // Adding values to complete the azimuthal range of the horizon
            ExtendedAzim[0] = InputAzim[InputAzim.Length - 1] - 2 * Math.PI;
            ExtendedAzim[ExtendedAzim.Length - 1] = InputAzim[0] + 2 * Math.PI;

            // Giving the extended horizon the correct elevation
            ExtendedElev[0] = InputElev[InputElev.Length - 1];
            ExtendedElev[ExtendedElev.Length - 1] = InputElev[0];

            for (int i = 1; i <= InputElev.Length; i++)
            {
                ExtendedAzim[i] = InputAzim[i - 1];
                ExtendedElev[i] = InputElev[i - 1];
            }
        }

        // Using the extended horizon profiles interpolates the horizon elevation for every degree value and creates a 361 element array
        public static double[] GetInterpolatedElevationProfile
            (
              double[] Azim                                      // The extended azimuth profile [radians]
            , double[] Elev                                      // The extended elevation profile [radians]
            , double[] HorizonProfileAzi                         // The 361 element azimuth profile [radians]
            ) 
        {
            double[] HorizonProfileElev = new double[361];       // The 361 element elevation profile to be filled [radians]

            // Using the Interpolate class to find the horizon elevation for every degree value of the azimuth
            for (int i = 0; i < HorizonProfileElev.Length; i++)
            {
                HorizonProfileElev[i] = Interpolate.Linear(Azim, Elev, HorizonProfileAzi[i]);
            }

            return HorizonProfileElev;
        }

        // Calculates the horizon diffuse factor using the view factor equation from the Far Shading Documentation
        public static double GetHorizonDiffuseFactor
            (
              double PanelTilt                                   // The tilt of of the panel surface [radians]
            , double PanelAzimuth                                // The azimuth position of the panel surface [radians]
            , double[] HorizonProfileAzim                        // The 361 element horizon azimuth profile [radians]
            , double[] HorizonProfileElev                        // The 361 element horizon elevation profile [radians]
            , double[] LimitingAngle                             // The 361 element array of limiting angles for every azimuth [radians]
            )
        {
            double DiffuseFactor = 0;                            // Initializing shading factor value
            double cosThetaTerm = 0;                             // Initializing the CosThetaTerm value
            double sinThetaTerm = 0;                             // Initializing the SinThetaTerm value
            double thetaShade;
            // The step value is used to convert the integration step value to the equivalent value in radians
            double step = Math.PI / 180;

            // Numerical integration of the horizon elevation effects on the view factor
            for (int i = 0; i < HorizonProfileAzim.Length - 1; i++)
            {
                thetaShade = Math.PI / 2 - HorizonProfileElev[i];               
                // If the panel tilt defines the horizon, it is not included in the integration so the panel effect is not considered twice
                if (LimitingAngle[i] < thetaShade)
                {
                    continue;
                }

                cosThetaTerm += (Math.Cos(2 * thetaShade) - Math.Cos(2 * LimitingAngle[i]));
                sinThetaTerm += Math.Cos(HorizonProfileAzim[i] - PanelAzimuth) * (2 * (LimitingAngle[i] - thetaShade) + Math.Sin(2 * thetaShade) - Math.Sin(2 * LimitingAngle[i]));
            }

            DiffuseFactor = (1 / (4 * Math.PI)) * Math.Cos(PanelTilt) * step * cosThetaTerm;           
            DiffuseFactor += (1 / (4 * Math.PI)) * Math.Sin(PanelTilt) * step * sinThetaTerm;

            // Shading factor is a fraction of the view factor left from the already accounted for panel tilt
            DiffuseFactor = (((1 + Math.Cos(PanelTilt)) / 2) - DiffuseFactor) / ((1 + Math.Cos(PanelTilt)) / 2);
            return DiffuseFactor;
        }

        // Calculates the ground reflected horizon factor. Applies the diffuse and beam horizon factors to the horizontal irradiance to get the total factor of irradiance incident on the ground
        public static double GetHorizonGRefFactor
            (
              double RefBeamFactor                               // The beam horizon factor calculated prior in the calculate method [#]
            , double RefDiffFactor                               // The diffuse horizon factor of the ground, calculated in config [#]
            , double HBeam                                       // The beam irradiance incident on a horizontal detector [W/m^2]
            , double HDiff                                       // The diffuse irradiance incident on a horizontal detector [W/m^2]
            )
        {
            double GRefFactor = 0;                               // The ground reflected horizon factor. Defaulted to 0, calculation run if there is irradiance on a horizontal detector [#]

            if (HDiff + HBeam != 0)
            {
                GRefFactor = (RefDiffFactor * HDiff + RefBeamFactor * HBeam) / (HDiff + HBeam);
            }
            
            return GRefFactor;
        }

        // Returns either 1 or 0 for the beam horizon factor depending on whether the sun is above or below the horizon
        public static double GetHorizonBeamFactor
            (
              double SunAzimuth                                  // The azimuth position of the sun [radians]
            , double SunZenith                                   // The zenith position of the sun [radians]
            , double[] HorizonProfileAzim                        // The 361 element azimuth profile of the horizon [radians]
            , double[] HorizonProfileElev                        // The 361 element elevation profile of the horizon [radians]
            , double PanelAzimuth                                // The azimuth direction in which the panel is facing [radians]
            )                               
        {
            // Defining the elevation values
            double TempSunAzim;  // New value for sun azimuth if the value is more than 180 degrees greater or less than the panel azimuth
            double HorizonElev;
            double BeamSF;
            double SunElev = Math.PI / 2 - SunZenith;

            if (SunAzimuth < PanelAzimuth - Math.PI)
            {
                TempSunAzim = SunAzimuth + 2 * Math.PI;
            }
            else if (SunAzimuth > PanelAzimuth + Math.PI)
            {
                TempSunAzim = SunAzimuth - 2 * Math.PI;
            }
            else
            {
                TempSunAzim = SunAzimuth;
            }

            HorizonElev = Interpolate.Linear(HorizonProfileAzim, HorizonProfileElev, TempSunAzim);            
            
            // Whether the sun is above or below the defined horizon
            // Tolerance of 1/4 a degree, as that is the approximate angular diameter of the solar disk
            if (SunElev > (HorizonElev - Math.PI / 180 / 4 ))
            {
                BeamSF = 1;
            }
            else
            {
                BeamSF = 0;
            }

            return BeamSF;
        }
    }
}
