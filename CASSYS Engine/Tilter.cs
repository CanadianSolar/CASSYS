// CASSYS - Grid connected PV system modelling software 
// Version 0.9  
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Tilter Class
// 
// Revision History:
// DT - 2014-11-11: Version 0.1
//
// Description: 
// Implementation of the Hay and Perez transposition models
///////////////////////////////////////////////////////////////////////////////
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
using System.Text;
using System.Xml;
using System.Xml.Linq;
using CASSYS;

namespace CASSYS
{

    // Constants
    // This enumeration is used to define the type of tilt algorithm to use
    public enum TiltAlgorithm { HAY = 0, PEREZ = 1 };

    // Tilted class
    class Tilter
    {
        // Tilter definition variables
        TiltAlgorithm itsTiltAlgorithm;            // Tilt algorithm used (Hay or Perez)

        // Parameters for the class (made public for availability to other classes)
        public double itsSurfaceSlope;             // Surface slope [radians]
        public double itsSurfaceAzimuth;           // Surface azimuth [radians]
        public double[] itsMonthlyAlbedo;          // Array to store Monthly Albedo values

        // Output variables calculated
        public double TGlo;                        // Tilted global irradiance [W/m2]
        public double TDir;                        // Beam irradiance in tilted plane [W/m2]
        public double TDif;                        // Diffuse irradiance in tilted plane [W/m2]
        public double TRef;                        // Reflected irradiance in tilted plane [W/m2]
        public double IncidenceAngle;              // Incidence angle [radians]
        public bool NoPyranoAnglesDefined;         // Boolean to track if the angles for the pyranometer are defined.            

        // Blank constructor
        public Tilter()
        {
        }

        // Constructor with definition variables
        public Tilter
        (
            TiltAlgorithm aTiltAlgorithm
        )
        {
            itsTiltAlgorithm = aTiltAlgorithm;
        }

        // Calculation method
        public void Calculate
            (
              double NDir            // normal direct irradiance [W/m2]
            , double HDif            // horizontal diffuse irradiance [W/m2]
            , double NExtra          // normal extraterrestrial irradiance [W/m2]
            , double SunZenith       // zenith angle of sun [radians]
            , double SunAzimuth      // azimuth angle of sun [radians]
            , double AirMass         // air mass [#]
            , int MonthNum           // The month number of the time stamp [1->12]
            )
        {
            // Calculate direct horizontal if direct normal is provided
            double HDir = NDir * Math.Cos(SunZenith);

            // call Perez et al. algorithm or Hay algorithm
                switch (itsTiltAlgorithm)
                {
                    case TiltAlgorithm.PEREZ:
                        TGlo = GetTiltCompIrradPerez(out TDir, out TDif, out TRef, HDir, HDif, NExtra, SunZenith, SunAzimuth, AirMass, MonthNum);
                        break;
                    case TiltAlgorithm.HAY:
                        TGlo = GetTiltCompIrradHay(out TDir, out TDif, out TRef, HDir, HDif, NExtra, SunZenith, SunAzimuth, MonthNum);
                        break;
                    default:
                        itsTiltAlgorithm = TiltAlgorithm.HAY;
                        break;
                }

            // Compute the incidence angle from the Tilt class
            IncidenceAngle = Tilt.GetIncidenceAngle(SunZenith, SunAzimuth, itsSurfaceSlope, itsSurfaceAzimuth);
        }

        // Calculation of global irradiance on a tilted surface using the Perez et al.
        // model (Perez et al, 1990).

        // Conversion for direct irradiance is geometric, for diffuse
        // irradiance an empirical model is used, and for reflected the
        // isotropic assumption is employed.

        double GetTiltCompIrradPerez        // (o) global irradiance on tilted surface [W/m2] 
            (
              out double TDir               // (o) beam //   on tilted surface [W/m2] 
            , out double TDif               // (o) diffuse irradiance on tilted surface [W/m2] 
            , out double TRef               // (o) reflected irradiance on tilted surface [W/m2] 
            , double HDir                   // (i) direct irradiance on horizontal surface [W/m2] 
            , double HDif                   // (i) diffuse irradiance on horizontal surface [W/m2] 
            , double NExtra                 // (i) normal extraterrestrial irradiance [W/m2] 
            , double SunZenith              // (i) zenith angle of sun [radians] 
            , double SunAzimuth             // (i) azimuth angle of sun [radians] 
            , double AirMass                // (i) air mass []
            , int MonthNum                  // (i) The month number 1 -> 12 used to allow monthly albedo
            )
        {
            // Declarations 
            int ibin;
            double delta, eps;
            double a, b, fone, ftwo;

            // Define constants
            //    f = circumsolar and horizon brightening coefficients
            //    epsbin: bins for sky's clearness
            //    kappa:
            double[, ,] f = new double[2, 3, 8]
                { { { -0.008, 0.130, 0.330, 0.568, 0.873, 1.132, 1.060, 0.678 }
                , {  0.588, 0.683, 0.487, 0.187,-0.392,-1.237,-1.600,-0.327 }
                , { -0.062,-0.151,-0.221,-0.295,-0.362,-0.412,-0.359,-0.250 } }
                , { { -0.060,-0.019, 0.055, 0.109, 0.226, 0.288, 0.264, 0.156 }
                , {  0.072, 0.066,-0.064,-0.152,-0.462,-0.823,-1.127,-1.377 }
                , { -0.022,-0.029,-0.026,-0.014, 0.001, 0.056, 0.131, 0.251 } } };
            double[] epsbin = new double[7] { 1.065, 1.23, 1.5, 1.95, 2.8, 4.5, 6.2 };
            double kappa = 1.041;

            // auxiliary quantities
            double cosZenith;                       // cos of zenith angle
            double cosInc;                          // cos incidence angle on slope
            double HGlo;                            // global irradiance on horizontal
            double rd1;                             // rd1, rd2: auxiliary quantities equal to (1+cos(slope))/2
            double rd2;                             // and (1-cos(slope))/2

            // initialize values 
            TGlo = TDif = TDir = TRef = 0;

            // Check arguments
            try
           {
                if (SunZenith < 0 || SunZenith > Math.PI || SunAzimuth < -Math.PI || SunAzimuth > Math.PI ||
                AirMass < 1 || itsSurfaceSlope < 0 || itsSurfaceSlope > Math.PI || itsSurfaceAzimuth < -Math.PI || itsSurfaceAzimuth > Math.PI || itsMonthlyAlbedo[MonthNum] < 0 || itsMonthlyAlbedo[MonthNum] > 1)
                {
                    throw new CASSYSException("GetTiltCompIrradiance: Arguments were out of range.");
               }
                cosZenith = Math.Cos(SunZenith);
                cosInc = Math.Cos(SunZenith) * Math.Cos(itsSurfaceSlope)
                         + Math.Sin(SunZenith) * Math.Sin(itsSurfaceSlope) * Math.Cos(itsSurfaceAzimuth - SunAzimuth);
                HGlo = HDif + HDir;
                rd1 = (1.0 + Math.Cos(itsSurfaceSlope)) / 2;
                rd2 = (1.0 - Math.Cos(itsSurfaceSlope)) / 2;


                // negative values or sun below horizon 
                if (HDif <= 0 && HDir <= 0) return 0.0;
                if (SunZenith > Math.PI / 2 || NExtra < 0) return HGlo;

                //  Compute delta, eps, and bin number
                //  delta = parametrization of sky's brightness
                //  eps = parametrization of sky's clearness
                //  ibin: bin number
                //  normally delta is in the range 0.08-0.048 (see Perez et al., 1990, fig. 5)
                //  but if the input data is wrong the values could be much higher, which can
                //  then cause problems in the calculation of tilted diffuse irradiance (TDif).
                //  Therefore we limit delta to 1.
                delta = Math.Min(HDif * AirMass / NExtra, 1);
                eps = 1 + HDir / cosZenith / HGlo / (1 + kappa * Math.Pow(SunZenith, 3));
                for (ibin = 0; ibin < 7 && eps > epsbin[ibin]; ibin++) ;

                // calculation of empirical coefficients a and b 
                a = Math.Max(0.0, cosInc);
                b = Math.Max(Math.Cos(85.0 * Util.DTOR), cosZenith);

                // calculation of empirical coefficient fone and ftwo 
                fone = Math.Max(0.0, f[0, 0, ibin] + delta * f[0, 1, ibin] + SunZenith * f[0, 2, ibin]);
                ftwo = f[1, 0, ibin] + delta * f[1, 1, ibin] + SunZenith * f[1, 2, ibin];

                // calculation of diffuse irradiance on the sloping surface 
                TDif = HDif * ((1 - fone) * rd1 + fone * a / b + ftwo * Math.Sin(itsSurfaceSlope));

                // calculation of direct irradiance on a sloping surface
                //  this is just a trigonometric transformation
                //  note: to avoid problems at low sun angles, HDir/cosZenith is limited
                //  to 90% of solar constant 
                TDir = Math.Max(Math.Min(HDir / cosZenith, 0.9 * Util.SOLAR_CONST) * cosInc, 0);

                // calculation of reflected irradiance onto the slope
                //  from Ineichen.  P.et al., Solar Energy, 41(4), 371-377, 1988
                //  a simple assumption of isotropic reflection is used 
                TRef = (HDir + HDif) * itsMonthlyAlbedo[MonthNum] * rd2;

                // summation for global irradiance on slope 
                TGlo = TDir + TDif + TRef;

                // end of subroutine 
                return TGlo;
            }
            catch (CASSYSException cs)
            {
                ErrorLogger.Log(cs, ErrLevel.WARNING);
                return Util.BADDATA;
            }
        }

        ///////////////////////////////////////////////////////////////////////////////
        // Calculation of global irradiance on a tilted surface using the Perez et al.
        // model (Perez et al, 1990) - short argument list
        double GetTiltIrradPerez        // (o) global irradiance on tilted surface [W/m2] 
            (double HDir                // (i) direct irradiance on horizontal surface [W/m2] 
            , double HDif               // (i) diffuse irradiance on horizontal surface [W/m2] 
            , double NExtra             // (i) normal extraterrestrial irradiance [W/m2] 
            , double SunZenith          // (i) zenith angle of sun [radians] 
            , double SunAzimuth         // (i) azimuth angle of sun [radians] 
            , double AirMass            // (i) air mass [] 
            , int MonthNum              // The month of the year number [1->12]
            )
        {
            double TDir;
            double TDif;
            double TRef;
            return GetTiltCompIrradPerez(out TDir, out TDif, out TRef, HDir, HDif, NExtra, AirMass, itsSurfaceSlope, itsSurfaceAzimuth, MonthNum);
        }

        ///////////////////////////////////////////////////////////////////////////////
        // Calculation of global irradiance on a tilted surface using the Hay and
        // Davies model (Hay and Davies, 1978).
        //
        // Conversion for direct irradiance is geometric, for diffuse
        // irradiance an empirical model is used, and for reflected the
        // isotropic assumption is employed.

        double GetTiltCompIrradHay      // (o) global irradiance on tilted surface [W/m2] 
            (
            out double TDir             // (o) beam irradiance on tilted surface [W/m2] 
            , out double TDif           // (o) diffuse irradiance on tilted surface [W/m2] 
            , out double TRef           // (o) reflected irradiance on tilted surface [W/m2] 
            , double HDir               // (i) direct irradiance on horizontal surface [W/m2] 
            , double HDif               // (i) diffuse irradiance on horizontal surface [W/m2] 
            , double NExtra             // (i) normal extraterrestrial irradiance [W/m2] 
            , double SunZenith          // (i) zenith angle of sun [radians] 
            , double SunAzimuth         // (i) azimuth angle of sun [radians] 
            , int MonthNum              // The month of the year number [1->12]
            )
        {
            // Declarations 
            double cosInc;              // The cosine of the incidence angle 
            double cosZenith;           // The cosine of the zenith angle
            double Rb;                  // Rb is the ratio of beam radiation on the tilted surface to that on a horizontal surface eqn 1.8.1 
            double AI;                  // Anisotropy Index eqn. 2.16.2 and 2.16.3 
            double HGlo = HDir + HDif;  // Global on horizontal 
            double TGlo;                // The global radiation in the plane of the array

            //  Initialize values 
            TGlo = TDif = TDir = TRef = 0;

            //  Negative values or sun below horizon 
            if (HDif <= 0 && HDir <= 0) return 0.0;

            //  Check arguments
            if (NExtra < 0 || SunZenith < 0 || SunZenith > Math.PI || SunAzimuth < -Math.PI || SunAzimuth > Math.PI ||
                itsSurfaceSlope < 0 || itsSurfaceSlope > Math.PI || itsSurfaceAzimuth < -Math.PI || itsSurfaceAzimuth > Math.PI || itsMonthlyAlbedo[MonthNum] < 0 || itsMonthlyAlbedo[MonthNum] > 1)

            {
                ErrorLogger.Log("GetTiltCompIrradHay: out of range arguments.", ErrLevel.FATAL);
            }

            //  Negative values or sun below horizon 
            if (HDif <= 0 && HDir <= 0) return 0.0;

            //  Compute cosine of incidence angle and cosine of zenith angle
            //  cos(Zenith) is bound by cos(86.25 degrees) to avoid large values
            //  near sunrise and sunset. 
            cosInc = Tilt.GetCosIncidenceAngle(SunZenith, SunAzimuth,
                itsSurfaceSlope, itsSurfaceAzimuth);

            cosZenith = Math.Max(Math.Cos(SunZenith), Math.Cos(89.0 * Util.DTOR));

            //  Compute tilted beam irradiance
            //  Rb is the ratio of beam radiation on the tilted surface to that on
            //  a horizontal surface. Duffie and Beckman (1991) eqn 1.8.1 
            Rb = 0;
            if (SunZenith < Math.PI / 2 && cosInc > 0)
                Rb = cosInc / cosZenith;
            TDir = HDir * Rb;

            // Compute anisotropy index AI and diffuse radiation
            //  Duffie and Beckman (1991) eqn. 2.16.2 and 2.16.3 
            AI = Math.Min(HDir / NExtra / cosZenith, 1.0);


			// Calculate diffuse irradiance
            // If sun below horizon or sun behind panel, treat all irradiance as diffuse isotropic
            if (SunZenith >= Math.PI / 2 || cosInc <= 0)
            {
                TDif = HGlo * (1 + Math.Cos(itsSurfaceSlope)) / 2;
            }
            else
            {
                TDif = HDif * (AI * Rb + (1 - AI) * (1 + Math.Cos(itsSurfaceSlope)) / 2);
            }

            // Compute ground-reflected irradiance 
            TRef = HGlo * itsMonthlyAlbedo[MonthNum] * (1 - Math.Cos(itsSurfaceSlope)) / 2;

            // Compute titled global irradiance 
            TGlo = TDir + TDif + TRef;

            // Normal end of subroutine 
            return TGlo;
        }

        ///////////////////////////////////////////////////////////////////////////////
        // Calculation of global irradiance on a tilted surface using the Hay and
        // Davies model (Hay and Davies, 1978) - short argument list
        double GetTiltIrradHay          // (o) global irradiance on tilted surface [W/m2] 
            (
              double HDir               // (i) direct irradiance on horizontal surface [W/m2] 
            , double HDif               // (i) diffuse irradiance on horizontal surface [W/m2] 
            , double NExtra             // (i) normal extraterrestrial irradiance [W/m2] 
            , double SunZenith          // (i) zenith angle of sun [radians] 
            , double SunAzimuth         // (i) azimuth angle of sun [radians]
            , int MonthNum              // (i) The month of the year number [1->12]
            )
        {
            return GetTiltCompIrradHay(out TDir, out TDif, out TRef, HDir, HDif, NExtra, SunZenith, SunAzimuth, MonthNum);
        }

        // Config will assign parameter variables their values as obtained from the XML file            
        public void Config()
        {
            // Getting the Tilt Algorithm for the Simulation
            if (ReadFarmSettings.GetInnerText("Site", "TransEnum", ErrLevel.WARNING) == "0")
            {
                itsTiltAlgorithm = TiltAlgorithm.HAY;
            }
            else if (ReadFarmSettings.GetInnerText("Site", "TransEnum", ErrLevel.WARNING) == "1")
            {
                itsTiltAlgorithm = TiltAlgorithm.PEREZ;
            }
            else
            {
                ErrorLogger.Log("Tilter: Invalid tilt algorithm chosen by User. CASSYS uses Hay as default.", ErrLevel.WARNING);
                itsTiltAlgorithm = TiltAlgorithm.HAY;
            }

            // Assign the albedo parameters from the .CSYX file
            ConfigAlbedo();
        }

        // Config will assign parameter variables their values as obtained from the .CSYX file
        public void ConfigPyranometer()
        {
            try
            {
                // Getting the parameter values
                itsSurfaceSlope = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("InputFile", "MeterTilt", ErrLevel.FATAL));
                itsSurfaceAzimuth = Util.DTOR * Convert.ToDouble(ReadFarmSettings.GetInnerText("InputFile", "MeterAzimuth", ErrLevel.FATAL));
                NoPyranoAnglesDefined = false;
            }
            catch
            {
                ErrorLogger.Log("Irradiance Meter Tilt or Azimuth were not specified. CASSYS will assume these values are the same as the Array Tilt and Azimuth. This can be changed in the Climate File Sheet in the interface.", ErrLevel.WARNING);
                NoPyranoAnglesDefined = true;
            }

            // Assign the albedo parameters from the .CSYX file
            ConfigAlbedo();
        }
        
        // Config the albedo value based on monthly/yearly values defined on file
        public void ConfigAlbedo()
        {
            // Getting the Albedo values either at a monthly or yearly level
            if (ReadFarmSettings.GetAttribute("Albedo", "Frequency", ErrLevel.WARNING) == "Monthly")
            {
                // Initializing the expected list
                itsMonthlyAlbedo = new double[13];

                // Using the month number as the index, populate the Albedo vales from each correspodning node
                itsMonthlyAlbedo[1] = double.Parse(ReadFarmSettings.GetInnerText("Albedo", "Jan", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[2] = double.Parse(ReadFarmSettings.GetInnerText("Albedo", "Feb", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[3] = double.Parse(ReadFarmSettings.GetInnerText("Albedo", "Mar", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[4] = double.Parse(ReadFarmSettings.GetInnerText("Albedo", "Apr", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[5] = double.Parse(ReadFarmSettings.GetInnerText("Albedo", "May", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[6] = double.Parse(ReadFarmSettings.GetInnerText("Albedo", "Jun", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[7] = double.Parse(ReadFarmSettings.GetInnerText("Albedo", "Jul", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[8] = double.Parse(ReadFarmSettings.GetInnerText("Albedo", "Aug", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[9] = double.Parse(ReadFarmSettings.GetInnerText("Albedo", "Sep", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[10] = double.Parse(ReadFarmSettings.GetInnerText("Albedo", "Oct", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[11] = double.Parse(ReadFarmSettings.GetInnerText("Albedo", "Nov", ErrLevel.WARNING, _default: "0.2"));
                itsMonthlyAlbedo[12] = double.Parse(ReadFarmSettings.GetInnerText("Albedo", "Dec", ErrLevel.WARNING, _default: "0.2"));
            }
            else
            {
                // Initializing the expected list
                itsMonthlyAlbedo = new double[13];
                itsMonthlyAlbedo[1] = Convert.ToDouble(ReadFarmSettings.GetInnerText("Albedo", "Yearly", ErrLevel.WARNING, _default: "0.2"));

                // Initializing the expected list
                for (int i = 3; i < itsMonthlyAlbedo.Length + 1; i++)
                {
                    itsMonthlyAlbedo[i - 1] = itsMonthlyAlbedo[1];
                }
            }
        }
    }
}
