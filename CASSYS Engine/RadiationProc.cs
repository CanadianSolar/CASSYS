// CASSYS - Grid connected PV system modelling software 
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: RadiationProc.cs
// 
// Revision History:
// NA - 2017-06-09: First release - Modularized the simulation class
//
// Description 
// This class is used to deal with radiation related processes within the simulation.
// This class configures/initializes radiation related classes, performes radiation calculations,
// and assigns readiation outputs.
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
//
//
///////////////////////////////////////////////////////////////////////////////

using System;

namespace CASSYS
{
    class RadiationProc
    {
        // Accessible by other classes that require this information.
        public Sun SimSun = new Sun();                                  // Instance used for Radiation calculations
        public Tracker SimTracker = new Tracker();                      // Instance used for use in systems with tracking
        public HorizonShading SimHorizonShading = new HorizonShading(); // Instance used for Horizon shading calculations
        public Tilter SimTilter = new Tilter();                         // instance used for titler calculations
        DateTime TimeStampAnalyzed;                                     // The time-stamp that used for Sun position calculations [yyyy-mm-dd hh:mm:ss]
        double HourOfDay;                                               // Hour of day specific to radiation calculations

        Splitter SimSplitter = new Splitter();
        Tilter pyranoTilter = new Tilter(TiltAlgorithm.HAY);
        bool negativeIrradFlag = false;                                 // Negative Irradiance Warning Flag

        // Blank Constructor for the Radiation Processor Object
        public RadiationProc()
        {

        }

        public void Calculate(
            
            SimMeteo SimMet                                             // Meteological data from inputfile

            )
        {
            // Calculating Sun position
            // Calculate the Solar Azimuth, and Zenith angles [radians]
            SimSun.itsSurfaceSlope = SimTracker.SurfSlope;
            SimSun.Calculate(SimMet.DayOfYear, SimMet.HourOfDay);
            
            HourOfDay = SimMet.HourOfDay;

            // The time stamp must be adjusted for sunset and sunrise hours such that the position of the sun is only calculated
            // for the middle of the interval where the sun is above the horizon.
            if ((SimMet.TimeStepEnd > SimSun.TrueSunSetHour) && (SimMet.TimeStepBeg < SimSun.TrueSunSetHour))
            {
                HourOfDay = SimMet.TimeStepBeg + (SimSun.TrueSunSetHour - SimMet.TimeStepBeg) / 2;
            }
            else if ((SimMet.TimeStepBeg < SimSun.TrueSunRiseHour) && (SimMet.TimeStepEnd > SimSun.TrueSunRiseHour))
            {
                HourOfDay = SimSun.TrueSunRiseHour + (SimMet.TimeStepEnd - SimSun.TrueSunRiseHour) / 2;
            }

            // Based on the definition of Input file, use Tilted irradiance or transpose the horizontal irradiance
            if (ReadFarmSettings.UsePOA == true)
            {
                // Check if the meter tilt and surface tilt are equal, if not detranspose the pyranometer
                if (string.Compare(ReadFarmSettings.CASSYSCSYXVersion, "0.9.2") >=0)
                {
                    // Checking if the Meter and Panel Tilt are different:
                    if ((pyranoTilter.itsSurfaceAzimuth != SimTracker.SurfAzimuth) || (pyranoTilter.itsSurfaceSlope != SimTracker.SurfSlope))
                    {
                        if (SimMet.TGlo < 0)
                        {
                            SimMet.TGlo = 0;

                            if (negativeIrradFlag == false)
                            {
                                ErrorLogger.Log("Global Plane of Array Irradiance contains negative values. CASSYS will set the value to 0.", ErrLevel.WARNING);
                                negativeIrradFlag = true;
                            }
                        }
                        PyranoDetranspose(SimMet);
                    }
                    else
                    {
                        if (SimMet.TGlo < 0)
                        {
                            SimMet.TGlo = 0;

                            if (negativeIrradFlag == false)
                            {
                                ErrorLogger.Log("Global Plane of Array Irradiance contains negative values. CASSYS will set the value to 0.", ErrLevel.WARNING);
                                negativeIrradFlag = true;
                            }
                        }
                        Detranspose(SimMet);
                    }
                }
                else
                {
                    if (SimMet.TGlo < 0)
                    {
                        SimMet.TGlo = 0;

                        if (negativeIrradFlag == false)
                        {
                            ErrorLogger.Log("Global Plane of Array Irradiance contains negative values. CASSYS will the value to 0.", ErrLevel.WARNING);
                            negativeIrradFlag = true;
                        }
                    }
                    Detranspose(SimMet);
                }
            }
            else
            {
                if (SimMet.HGlo < 0)
                {
                    SimMet.HGlo = 0;

                    if (negativeIrradFlag == false)
                    {
                        ErrorLogger.Log("Global Horizontal Irradiance is negative. CASSYS set the value to 0.", ErrLevel.WARNING);
                        negativeIrradFlag = true;
                    }
                }
                if (ReadFarmSettings.UseDiffMeasured == true)
                {
                    if (SimMet.HDiff < 0)
                    {
                        if (negativeIrradFlag == false)
                        {
                            SimMet.HDiff = 0;
                            ErrorLogger.Log("Horizontal Diffuse Irradiance is negative. CASSYS set the value to 0.", ErrLevel.WARNING);
                            negativeIrradFlag = true;
                        }
                    }
                }
                else
                {
                    SimMet.HDiff = double.NaN;
                }

                Transpose(SimMet);

            }
            // Calculate horizon shading effects
            SimHorizonShading.Calculate(SimSun.Zenith, SimSun.Azimuth, SimTracker.SurfSlope, SimTracker.SurfAzimuth, SimTilter.TDir, SimTilter.TDif, SimTilter.TRef, SimSplitter.HDir, SimSplitter.HDif, SimTracker.itsTrackMode);

            // Assigning outputs
            AssignOutputs();
        }

        // Transposition of the global horizontal irradiance values to the transposed values
        void Transpose(SimMeteo SimMet)
        {
            SimSun.Calculate(SimMet.DayOfYear, HourOfDay);

            // Calculating the Surface Slope and Azimuth based on the Tracker Chosen
            SimTracker.Calculate(SimSun.Zenith, SimSun.Azimuth, SimMet.Year, SimMet.DayOfYear);
            SimTilter.itsSurfaceSlope = SimTracker.SurfSlope;
            SimTilter.itsSurfaceAzimuth = SimTracker.SurfAzimuth;

            if (double.IsNaN(SimMet.HDiff))
            {
                // Split global into direct
                SimSplitter.Calculate(SimSun.Zenith, SimMet.HGlo, NExtra: SimSun.NExtra);
            }
            else
            {
                // Split global into direct and diffuse
                SimSplitter.Calculate(SimSun.Zenith, SimMet.HGlo, _HDif: SimMet.HDiff, NExtra: SimSun.NExtra);
            }

            // Calculate tilted irradiance
            SimTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, SimMet.MonthOfYear);

        }

        // De-transposition of the titled irradiance values to the global horizontal values
        void Detranspose(SimMeteo SimMet)
        {
            // Lower bound of bisection
            double HGloLo = 0;

            // Higher bound of bisection
            double HGloHi = SimSun.NExtra;

            // Calculating the Surface Slope and Azimuth based on the Tracker Chosen
            SimTracker.Calculate(SimSun.Zenith, SimSun.Azimuth, SimMet.Year, SimMet.DayOfYear);
            SimTilter.itsSurfaceSlope = SimTracker.SurfSlope;
            SimTilter.itsSurfaceAzimuth = SimTracker.SurfAzimuth;
            SimTilter.IncidenceAngle = SimTracker.IncidenceAngle;

            // Calculating the Incidence Angle for the current setup
            double cosInc = Tilt.GetCosIncidenceAngle(SimSun.Zenith, SimSun.Azimuth, SimTilter.itsSurfaceSlope, SimTilter.itsSurfaceAzimuth);

            // Trivial case
            if (SimMet.TGlo <= 0)
            {
                SimSplitter.Calculate(SimSun.Zenith, 0, NExtra: SimSun.NExtra);
                SimTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, SimMet.MonthOfYear);
            }
            else if ((SimSun.Zenith > 87.5 * Util.DTOR) || (cosInc <= Math.Cos(87.5 * Util.DTOR)))
            {
                SimMet.HGlo = SimMet.TGlo / ((1 + Math.Cos(SimTilter.itsSurfaceSlope)) / 2 + SimTilter.itsMonthlyAlbedo[SimMet.MonthOfYear] * (1 - Math.Cos(SimTilter.itsSurfaceSlope)) / 2);

                // Forcing the horizontal irradiance to be composed entirely of diffuse irradiance
                SimSplitter.HGlo = SimMet.HGlo;
                SimSplitter.HDif = SimMet.HGlo;
                SimSplitter.NDir = 0;
                SimSplitter.HDir = 0;

                SimTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, SimMet.MonthOfYear);
            }
            // Otherwise, bisection loop
            else
            {
                // Bisection loop
                while (Math.Abs(HGloHi - HGloLo) > 0.01)
                {
                    // Use the central value between the domain to start the bisection, and then solve for TGlo,
                    double HGloAv = (HGloLo + HGloHi) / 2;
                    SimSplitter.Calculate(SimSun.Zenith, _HGlo: HGloAv, NExtra: SimSun.NExtra);
                    SimTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, SimMet.MonthOfYear);
                    double TGloAv = SimTilter.TGlo;

                    // Compare the TGloAv calculated from the Horizontal guess to the acutal TGlo and change the bounds for analysis
                    // comparing the TGloAv and TGlo
                    if (TGloAv < SimMet.TGlo)
                    {
                        HGloLo = HGloAv;
                    }
                    else
                    {
                        HGloHi = HGloAv;
                    }
                }
            }

            SimMet.TGlo = SimTilter.TGlo;
            SimMet.HGlo = SimSplitter.HGlo;
        }

        // De-transposition method to the be used if the meter and panel tilt do not match
        void PyranoDetranspose(SimMeteo SimMet)
        {
            if (pyranoTilter.NoPyranoAnglesDefined)
            {
                SimTracker.Calculate(SimSun.Zenith, SimSun.Azimuth, SimMet.Year, SimMet.DayOfYear);
                pyranoTilter.itsSurfaceAzimuth = SimTracker.SurfAzimuth;
                pyranoTilter.itsSurfaceSlope = SimTracker.SurfSlope;
                pyranoTilter.IncidenceAngle = SimTracker.IncidenceAngle;
            }

            // Lower bound of bisection
            double HGloLo = 0;

            // Higher bound of bisection
            double HGloHi = SimSun.NExtra;

            // Calculating the Incidence Angle for the current setup
            double cosInc = Tilt.GetCosIncidenceAngle(SimSun.Zenith, SimSun.Azimuth, pyranoTilter.itsSurfaceSlope, pyranoTilter.itsSurfaceAzimuth);

            // Trivial case
            if (SimMet.TGlo <= 0)
            {
                SimSplitter.Calculate(SimSun.Zenith, 0, NExtra: SimSun.NExtra);
                pyranoTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, SimMet.MonthOfYear);
            }
            else if ((SimSun.Zenith > 87.5 * Util.DTOR) || (cosInc <= Math.Cos(87.5 * Util.DTOR)))
            {
                SimMet.HGlo = SimMet.TGlo / ((1 + Math.Cos(pyranoTilter.itsSurfaceSlope)) / 2 + pyranoTilter.itsMonthlyAlbedo[SimMet.MonthOfYear] * (1 - Math.Cos(pyranoTilter.itsSurfaceSlope)) / 2);

                // Forcing the horizontal irradiance to be composed entirely of diffuse irradiance
                SimSplitter.HGlo = SimMet.HGlo;
                SimSplitter.HDif = SimMet.HGlo;
                SimSplitter.NDir = 0;
                SimSplitter.HDir = 0;

                //SimSplitter.Calculate(SimSun.Zenith, HGlo, NExtra: SimSun.NExtra);
                pyranoTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, SimMet.MonthOfYear);
            }
            // Otherwise, bisection loop
            else
            {
                // Bisection loop
                while (Math.Abs(HGloHi - HGloLo) > 0.01)
                {
                    // Use the central value between the domain to start the bisection, and then solve for TGlo,
                    double HGloAv = (HGloLo + HGloHi) / 2;
                    SimSplitter.Calculate(SimSun.Zenith, _HGlo: HGloAv, NExtra: SimSun.NExtra);
                    pyranoTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, SimMet.MonthOfYear);
                    double TGloAv = pyranoTilter.TGlo;

                    // Compare the TGloAv calculated from the Horizontal guess to the acutal TGlo and change the bounds for analysis
                    // comparing the TGloAv and TGlo
                    if (TGloAv < SimMet.TGlo)
                    {
                        HGloLo = HGloAv;
                    }
                    else
                    {
                        HGloHi = HGloAv;
                    }
                }
            }

            SimMet.HGlo = SimSplitter.HGlo;

            // This value of the horizontal global should now be transposed to the tilt value from the array. 
            Transpose(SimMet);
        }

        public void AssignOutputs()
        {
            // Using the TimeSpan function to assemble the String for the modified Time Stamp of calculation
            TimeSpan thisHour = TimeSpan.FromHours(HourOfDay);
            TimeStampAnalyzed = new DateTime(Utilities.CurrentTimeStamp.Year, Utilities.CurrentTimeStamp.Month, Utilities.CurrentTimeStamp.Day, thisHour.Hours, thisHour.Minutes, thisHour.Seconds);
            ReadFarmSettings.Outputlist["Timestamp_Used_for_Simulation"] = String.Format("{0:u}", TimeStampAnalyzed).Replace('Z', ' ');
            ReadFarmSettings.Outputlist["Sun_Zenith_Angle"] = Util.RTOD * SimSun.Zenith;
            ReadFarmSettings.Outputlist["Sun_Azimuth_Angle"] = Util.RTOD * SimSun.Azimuth;
            ReadFarmSettings.Outputlist["ET_Irrad"] = SimSun.NExtra;
            ReadFarmSettings.Outputlist["Air_Mass"] = SimSun.AirMass;
            ReadFarmSettings.Outputlist["Albedo"] = SimTilter.itsMonthlyAlbedo[Utilities.CurrentTimeStamp.Month];
            ReadFarmSettings.Outputlist["Normal_beam_irradiance"] = SimSplitter.NDir;
            ReadFarmSettings.Outputlist["Horizontal_Global_Irradiance"] = SimSplitter.HGlo;
            ReadFarmSettings.Outputlist["Horizontal_diffuse_irradiance"] = SimSplitter.HDif;
            ReadFarmSettings.Outputlist["Horizontal_beam_irradiance"] = SimSplitter.HDir;
            ReadFarmSettings.Outputlist["Global_Irradiance_in_Array_Plane"] = SimTilter.TGlo;
            ReadFarmSettings.Outputlist["Beam_Irradiance_in_Array_Plane"] = SimTilter.TDir;
            ReadFarmSettings.Outputlist["Diffuse_Irradiance_in_Array_Plane"] = SimTilter.TDif;
            ReadFarmSettings.Outputlist["Ground_Reflected_Irradiance_in_Array_Plane"] = SimTilter.TRef;
            ReadFarmSettings.Outputlist["Tracker_Slope"] = SimTracker.itsTrackerSlope * Util.RTOD;
            ReadFarmSettings.Outputlist["Tracker_Azimuth"] = SimTracker.itsTrackerAzimuth * Util.RTOD;
            ReadFarmSettings.Outputlist["Tracker_Rotation_Angle"] = SimTracker.RotAngle * Util.RTOD;
            ReadFarmSettings.Outputlist["Collector_Surface_Slope"] = SimTilter.itsSurfaceSlope * Util.RTOD;
            ReadFarmSettings.Outputlist["Collector_Surface_Azimuth"] = SimTilter.itsSurfaceAzimuth * Util.RTOD;
            ReadFarmSettings.Outputlist["Incidence_Angle"] = Math.Min(Util.RTOD * SimTilter.IncidenceAngle, 90);
            ReadFarmSettings.Outputlist["FarShading_Global_Loss"] = SimHorizonShading.LossGlo;
            ReadFarmSettings.Outputlist["FarShading_Beam_Loss"] = SimHorizonShading.LossDir;
            ReadFarmSettings.Outputlist["FarShading_Diffuse_Loss"] = SimHorizonShading.LossDif;
            ReadFarmSettings.Outputlist["FarShading_Ground_Reflected_Loss"] = SimHorizonShading.LossRef;
        }

        public void Config()
        {
            // Weather and radiation related objects.
            // The sun class requires the configuration of the surface slope to calculate the apparent sunset and sunrise hours.
            SimTracker.Config();
            SimHorizonShading.Config(SimTracker.SurfSlope, SimTracker.SurfAzimuth, SimTracker.itsTrackMode);
            SimSun.Config();

            if (ReadFarmSettings.UsePOA)
            {
                pyranoTilter.ConfigPyranometer();
            }

            SimTilter.Config();
        }
    }
}
