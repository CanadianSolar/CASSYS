// CASSYS - Grid connected PV system modelling software 
// Version 0.9  
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Shading Class
// 
// Revision History:
// AP - 2014-11-10: Version 0.9
//
// NB - 2016-03-17: Version 0.9.3
//
// Description:
// This class calculates the shading factors on the beam, diffuse and ground-reflected 
// components of incident irradiance based on the sun position throughout the day 
// resulting from a near shading model. The shading models available are panels 
// arranged in an unlimited rows or a fixed tilt configuration. If the unlimited 
// row model is to be used, the model can be further customized to use a linear 
// shading model or a cell based (step-wise) shading model.
//                             
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
// Ref 1: Duffie JA and Beckman WA (1991) Solar Engineering of Thermal
//     Processes, Second Edition. John Wiley & Sons. Specific Example 1.9.3
//
// 
/////////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Text;
using CASSYS;

namespace CASSYS
{
    public enum ShadModel { FT, UR, TE, TS, None };                 // Different Array Types

    class Shading
    {
        // Parameters for the shading class
        double itsCollTilt;                         // Tilt of the collector [radians]
        double itsCollAzimuth;                      // Collector Azimuth [radians]
        public double itsShadingLimitAngle;         // The shading limit angle [radians]
        double itsCollBW;                           // Collector Distance [m]
        double itsPitch;                            // The distance between the rows [m]
        double itsRowsBlock;                        // The number of rows used in the farm set-up [#]
        double itsRowBlockFactor;                   // The factor applied to shading factors depending on the number of rows [#]
        ShadModel itsShadModel;                     // The shading model used based on the type of installation 
        bool useCellBasedShading;                   // Boolean used to determine if cell bases shading should be used [false, not used]
        int itsNumModTransverseStrings;             // The number of modules in a string (as they appear in the transverse direction)
        double[] cellSetup;                         // The number of cells in the transverse direction of the entire table
        double[] shadingPC;                         // The different shading percentages that occur
        double CellSize;                            // The size of a cell in the module [user defined]

        // Outputs of the shading class 
        public double DiffuseSF;                    // Diffuse shading fraction [%]
        public double BeamSF;                       // Beam shading fraction  [%]
        public double ReflectedSF;                  // Shading fraction applied to the Ground reflected component [%]
        public double ProfileAng;                   // Profile Angle [radians] (see Ref 1.)
        public double ShadTGlo = 0;                 // Post shading irradiance [W/m^2]
        public double ShadTDir = 0;                 // Tilted Beam Irradiance [W/m^2]
        public double ShadTDif = 0;                 // Tilted Diffuse Irradiance [W/m^2]
        public double ShadTRef = 0;                 // Tilted Ground Reflected Irradiance [W/m^2]

        // Blank constructor for the Shading calculations
        public Shading()
        {
        }


        // Calculate Method will solve for the shading factor as it applies to the Beam component of incident irradiance
        public void Calculate
            (
              double SunZenith              // Zenith angle of sun [radians] 
            , double SunAzimuth             // Azimuth angle of sun [radians]
            , double TDir                   // Tilted direct irradiance [W/m^2]
            , double TDif                   // Tilted diffuse irradiance [W/m^2]
            , double TRef                   // Ground Reflected Irradiance [W/m^2]
            , double CollectorTilt          // Tilt of the collector [radians]
            , double CollectorAzimuth       // Azimuth of the collector [radians]
            )
        {
            // Redefining the collector position (dynamic for Trackers, static for other types)
            itsCollTilt = CollectorTilt;
            itsCollAzimuth = CollectorAzimuth;

            switch (itsShadModel)
            {
                case ShadModel.UR:
                    GetBeamShadingFraction(SunZenith, SunAzimuth, CollectorTilt);
                    GetDiffuseShadingFraction();
                    GetGroundReflectedShadingFraction();
                    break;

                case ShadModel.TE:
                    GetBeamShadingFraction(SunZenith, SunAzimuth, CollectorTilt);
                    GetDiffuseShadingFraction();
                    GetGroundReflectedShadingFraction();
                    break;

                case ShadModel.TS:
                    GetBeamShadingFraction(SunZenith, SunAzimuth, CollectorTilt);
                    GetDiffuseShadingFraction();
                    GetGroundReflectedShadingFraction();
                    break;

                case ShadModel.FT:
                    // Calculating the shading fractions that apply to the beam irradiance (Diffuse and ground reflected component remain constant - calculated in Config)
                    BeamSF = 1;
                    break;
            }

            // Shading each component of the Tilted irradiance
            ShadTDir = TDir * BeamSF;
            ShadTDif = TDif * DiffuseSF;
            ShadTRef = TRef * ReflectedSF;

            // Combining for an effective Irradiance after shading value is applied
            ShadTGlo = ShadTDif + ShadTDir + ShadTRef;
        }

        // Returns the shading fraction that must be applied to the diffuse component of POA irradiance
        public void GetDiffuseShadingFraction
            (
            )
        {
            DiffuseSF = itsRowBlockFactor*(1 + Math.Cos(itsShadingLimitAngle)) / 2;
        }

        // Returns the shading limit angle which is needed to calculate beam and diffuse shading fractions
        public void GetShadingLimitAngle
        (
        double CollectorTilt
        )
        {
            if (CollectorTilt == 0)
            {
                itsShadingLimitAngle = 0;
            }

            else if (itsPitch == 0)
            {
                itsShadingLimitAngle = 0;
            }
            else
            {
                double aux = Math.Sqrt(Math.Pow(itsPitch, 2.0) + Math.Pow(itsCollBW, 2.0) - 2.0 * (itsCollBW * itsPitch * Math.Cos(CollectorTilt)));
                itsShadingLimitAngle = Math.Acos((Math.Pow(itsPitch, 2.0) + Math.Pow(aux, 2.0) - Math.Pow(itsCollBW, 2.0)) / (2 * itsPitch * aux));

                itsShadingLimitAngle = Math.Max(0, itsShadingLimitAngle);
                itsShadingLimitAngle = Math.Min(Math.PI, itsShadingLimitAngle);
            }
        }

        // Returns the shading fraction that must be applied to the albedo or ground-reflected component of POA irradiance
        public void GetGroundReflectedShadingFraction
            (
            )
        {
            // The ground reflected component is assumed to be seen only by the first shed; the following value will return the shading factor for all sheds, except 1
            ReflectedSF = 1 - itsRowBlockFactor;
        }


        // Returns the shading fraction that must be applied to the beam component of the POA irradiance
        public void GetBeamShadingFraction
            (
            double SunZenith                // Zenith angle of sun [radians] 
            , double SunAzimuth             // Azimuth angle of sun [radians]
            , double CollectorTilt          // Tilt of the module [radians]
            )
        {
            // NB: getting shading limit angle for shading calculations
            GetShadingLimitAngle(CollectorTilt);

            // Returns the shading factor for Beam using the GetShadedRow method
            BeamSF = 1 - GetShadedFraction(SunZenith, SunAzimuth)*itsRowBlockFactor;
        }

        // Computes the fraction of collectors arranged in rows that will be shaded
        // at a particular sun position.  Example 1.9.3 (Duffie and Beckman, 1991) 
        public double GetShadedFraction
            (
            double SunZenith                // Zenith angle of sun [radians] 
            , double SunAzimuth             // Azimuth angle of sun [radians]
            )
        {
            if (SunZenith > Math.PI / 2)
            {
                ProfileAng = 0;
                return 0;                  // No shading possible as Sun is set or not risen
            }
            else if (Math.Abs(SunAzimuth - itsCollAzimuth) > Util.DTOR * 90)
            {
                ProfileAng = 0;
                return 0;                 // No shading possible as Sun is behind the collectors
            }
            else
            {
                // Compute profile angle (see Ref 1.)
                ProfileAng = Tilt.GetProfileAngle(SunZenith, SunAzimuth, itsCollAzimuth);

                // NB: Added small tolerance since shading limit angle and profile angle are found through different methods and could have a small difference
                if (itsShadingLimitAngle - ProfileAng <= 0.000001)
                {
                    return 0; // No shading possible as the light reaching the panel behind the row is not limited by the preceding row
                }
                else
                {
                    // Computes the fraction of collectors arranged in rows that will be shaded at a particular sun position.  
                    double AC = Math.Sin(itsCollTilt) * itsCollBW / Math.Sin(itsShadingLimitAngle);
                    double CAAp = Math.PI - itsShadingLimitAngle - itsCollTilt;
                    double CApA = Math.PI - CAAp - (itsShadingLimitAngle - ProfileAng);
                    double ACAp = Math.PI - CAAp - CApA;

                    // Length of shaded section
                    double AAp = AC * Math.Sin(ACAp) / Math.Sin(CApA);

                    // Using the Cell based shading model
                    if (useCellBasedShading)
                    {
                        double cellShaded = AAp / (CellSize);               // The number of cells shaded
                        double SF = 1;                                      // The resultant shading factor initialized to 1, modified later// Calculate the Shading fraction based on which cell numbers they are between

                        for (int i = 1; i <= itsNumModTransverseStrings; i++)
                        {
                            if ((cellShaded > cellSetup[i - 1]) && (cellShaded < cellSetup[i]))
                            {
                                SF = Math.Min(((shadingPC[i - 1] + (shadingPC[i] * (cellShaded - cellSetup[i - 1])))), shadingPC[i]);
                            }
                        }
                        return SF;
                    }
                    else
                    {
                        // Return shaded fraction of the collector bandwidth
                        return (AAp / itsCollBW);
                    }
                }
            }
        }

        public void Config()
        {
            // Determine the type of Array layout
            switch (ReadFarmSettings.GetAttribute("O&S","ArrayType", ErrLevel.FATAL))
            {
                case "Unlimited Rows":
                    itsShadModel = ShadModel.UR;
                    break;

                case "Fixed Tilted Plane":
                    itsShadModel = ShadModel.FT;
                    break;

                case "Single Axis Elevation Tracking (E-W)":
                    itsShadModel = ShadModel.TE;
                    break;

                case "Single Axis Horizontal Tracking (N-S)":
                    itsShadModel = ShadModel.TS;
                    break;
                
                default:
                    itsShadModel = ShadModel.None;
                    break;
            }

            // reading and assigning values to the parameters for the Shading class based on the Array layout
            switch (itsShadModel)
            {
                case ShadModel.UR:

                    // Defining all the parameters for the shading of a unlimited row array configuration
                    itsCollTilt = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PlaneTilt", ErrLevel.FATAL));
                    itsCollAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "Azimuth", ErrLevel.FATAL));
                    itsPitch = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "Pitch", ErrLevel.FATAL));
                    itsCollBW = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "CollBandWidth", ErrLevel.FATAL));
                    
                    itsRowsBlock = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "RowsBlock", ErrLevel.FATAL));
                    itsRowBlockFactor = (itsRowsBlock - 1) / itsRowsBlock;

                    //// Running one-time only methods - the shading factors applied to diffuse and ground reflected component are constant throughout the simulation
                    //GetShadingLimitAngle();
                    //GetDiffuseShadingFraction();
                    //GetGroundReflectedShadingFraction();

                    // Collecting definitions for cell based shading models or preparing for its absence
                    useCellBasedShading = Convert.ToBoolean(ReadFarmSettings.GetInnerText("O&S","UseCellVal", ErrLevel.WARNING, _default: "false"));

                    // Set up the arrays to allow for shading calculations according to electrical effect
                    if (useCellBasedShading)
                    {
                        CellSize = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "CellSize", ErrLevel.FATAL)) / 100D;
                        itsNumModTransverseStrings = int.Parse(ReadFarmSettings.GetInnerText("O&S", "StrInWid", ErrLevel.FATAL));
                        itsRowBlockFactor = 1; // No row related shading adjustments should be applied.

                        // Use cell based shading to calculate the effect on the beam shading factor
                        // The shading factor gets worse in steps based on how much of the collector bandwidth is currently under shadowed length
                        cellSetup = new double[itsNumModTransverseStrings + 1];
                        shadingPC = new double[itsNumModTransverseStrings + 1];

                        // Defining the arrays needed for Number of cells in each string (transverse) and shading %
                        for (int i = 1; i <= itsNumModTransverseStrings; i++)
                        {
                            cellSetup[i] = (double)i / (double)itsNumModTransverseStrings * (itsCollBW / CellSize);
                            shadingPC[i] = (double)i / (double)itsNumModTransverseStrings;
                        }
                    }
                    break;

                case ShadModel.TE:
                    itsPitch = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PitchSAET", ErrLevel.FATAL));
                    itsCollBW = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "WActiveSAET", ErrLevel.FATAL));
                    itsRowsBlock = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "RowsBlockSAET", ErrLevel.FATAL));
                    
                    
                    itsRowBlockFactor = (itsRowsBlock - 1) / itsRowsBlock;

                    

                    // NB: Using same formula as Unlimited Rows, with SAET added to variable names
                    // Collecting definitions for cell based shading models or preparing for its absence
                    useCellBasedShading = Convert.ToBoolean(ReadFarmSettings.GetInnerText("O&S", "UseCellValSAET", ErrLevel.WARNING, _default: "false"));

                    // Set up the arrays to allow for shading calculations according to electrical effect
                    if (useCellBasedShading)
                    {
                        CellSize = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "CellSizeSAET", ErrLevel.FATAL)) / 100D;
                        itsNumModTransverseStrings = int.Parse(ReadFarmSettings.GetInnerText("O&S", "StrInWidSAET", ErrLevel.FATAL));
                        itsRowBlockFactor = 1; // No row related shading adjustments should be applied.

                        // Use cell based shading to calculate the effect on the beam shading factor
                        // The shading factor gets worse in steps based on how much of the collector bandwidth is currently under shadowed length
                        cellSetup = new double[itsNumModTransverseStrings + 1];
                        shadingPC = new double[itsNumModTransverseStrings + 1];

                        // Defining the arrays needed for Number of cells in each string (transverse) and shading %
                        for (int i = 1; i <= itsNumModTransverseStrings; i++)
                        {
                            cellSetup[i] = (double)i / (double)itsNumModTransverseStrings * (itsCollBW / CellSize);
                            shadingPC[i] = (double)i / (double)itsNumModTransverseStrings;
                        }
                    }
                    
                    break;

                case ShadModel.TS:
                    itsPitch = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "PitchSAST", ErrLevel.FATAL));
                    itsCollBW = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "WActiveSAST", ErrLevel.FATAL));
                    itsRowsBlock = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "RowsBlockSAST", ErrLevel.FATAL));
                    
                    itsRowBlockFactor = (itsRowsBlock - 1) / itsRowsBlock;

                    
                    // NB: Using same formula as Unlimited Rows, with SAST added to variable names
                    // Collecting definitions for cell based shading models or preparing for its absence
                    useCellBasedShading = Convert.ToBoolean(ReadFarmSettings.GetInnerText("O&S", "UseCellValSAST", ErrLevel.WARNING, _default: "false"));
                    
                    // Set up the arrays to allow for shading calculations according to electrical effect
                    if (useCellBasedShading)
                    {
                        CellSize = Convert.ToDouble(ReadFarmSettings.GetInnerText("O&S", "CellSizeSAST", ErrLevel.FATAL)) / 100D;
                        itsNumModTransverseStrings = int.Parse(ReadFarmSettings.GetInnerText("O&S", "StrInWidSAST", ErrLevel.FATAL));
                        itsRowBlockFactor = 1; // No row related shading adjustments should be applied.

                        // Use cell based shading to calculate the effect on the beam shading factor
                        // The shading factor gets worse in steps based on how much of the collector bandwidth is currently under shadowed length
                        cellSetup = new double[itsNumModTransverseStrings + 1];
                        shadingPC = new double[itsNumModTransverseStrings + 1];
                        
                        // Defining the arrays needed for Number of cells in each string (transverse) and shading %
                        for (int i = 1; i <= itsNumModTransverseStrings; i++)
                        {
                            cellSetup[i] = (double)i / (double)itsNumModTransverseStrings * (itsCollBW / CellSize);
                            shadingPC[i] = (double)i / (double)itsNumModTransverseStrings;
                        }
                    }
                    
                    break;

                case ShadModel.FT:

                    // Defining the parameters for the shading for a fixed tilt configuration 
                    itsShadingLimitAngle = 0;
                    

                    // Running one-time only methods - the shading factors applied to diffuse and ground reflected component are constant and 1.
                    DiffuseSF = 1;
                    ReflectedSF = 1;
                    break;
            
                case ShadModel.None:
                    // No shading applies.
                    DiffuseSF = 1;
                    ReflectedSF = 1;
                    BeamSF = 1;
                    break;


            }
        }
    }
}