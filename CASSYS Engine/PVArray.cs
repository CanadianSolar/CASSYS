// CASSYS - Grid connected PV system modelling software 
// Version 0.9 
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: PV Array Class
// 
// Revision History:
// AP - 2014-10-14: Version 0.9
// AP - 2015-04-23: Version 0.9.1 - Added the User defined IAM Profile Option.
// AP - 2015-06-10: Version 0.9.2 - Added the gamma temp coefficient determination method.
//
// Description:
// The PV Array Class evaluates the performance of a solar module using the 
// "standard" or one-diode model as described in Ref 1. The STC conditions 
// for the module are obtained from module data-sheets or a PAN file. Module
// behaviour is calculated for a number of non-STC operating conditions such
// as open circuit, fixed voltage, and maximum point tracking.
// Values are then converted from Module to Array level and losses are applied 
// in accordance with user input values.
//                             
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
// Ref 1: Mermoud, Lejeune (2010): Performance Assessment of a Simulation Model for PV Modules of Any Available Technology. 
//        Proc 25th European Photovoltaic Solar Energy Conference –Valencia, Spain, 6-10 September 2010 
// Ref 2: Mermoud (2013): PVSyst - Parameter determination PV Array Model at Sandia PV Performance Modelling Workshop.
//        Presented at the 2013 Sandia PV Performance Modeling Workshop, Santa Clara, CA. May 1-2, 2013
// Ref 3: Quaschning & Hanitsch (1996): Numerical Simulation of I-V characteristics Solar Energy 
//        Solar Energy 56, 513-520
// Ref 4: (Website, Accessed 2014-09) PV Modelling Collaborative: http://pvpmc.sandia.gov/modeling-steps/2-dc-module-iv/cell-temperature/pvsyst-cell-temperature-model/
// Ref 5: (Website, Accessed 2014-10) PV Modelling Collaborative: http://pvpmc.org/modeling-steps/shading-soiling-and-reflection-losses/incident-angle-reflection-losses/ashre-model/
///////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Text;
using System.IO;
using CASSYS;

namespace CASSYS
{
    class PVArray
    {
        // Parameters of the PV Module
        double itsPNom;                       // Nominal power of the module [W], typically at STC  
        double itsIscref;                     // Reference short circuit current [A]
        double itsVocref;                     // Reference open circuit voltage [V]
        double itsImppref;                    // Reference maximum power point current [A]
        double itsVmppref;                    // Reference maximum power point voltage [V]
        double itsEfficiencyRef;              // Reference module efficiency [%]
        public double itsArea;                // Area of individual module [m2]
        double itsIrsRef;                     // Evaluated using the GetGammaIrs method [A]
        int itsCellCount;                     // Number of cells in the module [#]
        double itsIPhiRef;                    // Reference photo-current determined from the Isc Point [A]
        double itsGammaRef;                   // Diode ideality factor, evaluated using CalcGammaIPhhirefIrsref method []
        double itsGamma;                      // Temperature adjusted gamma value [Unit-less]
        double itsGammaCoeff;                 // The gamma coefficient for a c-Si module [1/C] 
        double itsTref;                       // Reference temperature [C]
        public double itsHref;                // Reference irradiance level [W/m2]
        double itsRs;                         // Series resistance [ohms];
        double itsRshExp;                     // Rshunt exponential parameter follows PVSyst Model (default at 5.5 for c-Si modules)
        double itsRsh;                        // Unadjusted reference Rshunt [ohms]
        double itsRshZero;                    // Rshunt Parameter that follows the PVSyst Model
        double itsRpRef;                      // Shunt resistance [ohms];
        double itsRw;                         // Global wiring loss as seen from Inverter for entire Sub-Array [ohms]
        double itsSubArrayNum;                // The sub-array number that the PVArray belongs to. [#]
        

        // Module Temperature and Irradiance Coefficients
        double itsBo;                                       // ASHARE Parameter used for IAM calculation (see Ref 5)
        bool userIAMModel;                                  // If a userIAM profile is defined or the ASHRAE model is used.
        double[][] itsUserIAMProfile = new double[2][];     // The IAM profile entered by the user.
        double itsTCoefIsc;                                 // Temperature coefficient for Isc [1/C]
        double itsTCoefVoc;                                 // Temperature coefficient for Voc [1/C]
        double itsTCoefP;                                   // Temperature coefficient for Power [%/C]
        bool useMeasuredTemp;                               // Override boolean, true uses measured values, false calculates values using the Faiman Module Temperature Model

        // Array losses related variables
        double itsMismatchLossSTCPC;          // Mismatch Loss for Sub Array [%] 
        double itsMismatchFixedVLoss;         // Mismatch Loss during Fixed Voltage Operation [%]
        double itsModuleQualityLossPC;        // Module Quality Loss [%]
        double[] itsMonthlySoilingPC;         // Soiling loss Percentage Array [Month #, %]
        double itsSoilingLossPC;              // Soiling Loss Percentage [%]
        double itsConstHTC;                   // Constant Heat Transfer Coefficient (CHTC)
        double itsConvHTC;                    // Convective Heat Transfer Coefficient (ConvHTC)
        double itsAdsorp = 0.9;               // Adsorption, default 0.9 for PVSyst

        // PV Array local & intermediate calculation variables
        double lossLessTGlo;                  // Tilted Global Irradiance (used for efficiency calculations)
        double Irs;                           // Reverse saturation current [A]
        double mVoc;                          // Module open circuit voltage [V]
        double mIPhi;                         // Module short-circuit current [A]
        double mVout;                         // Module voltage at max power point [V]
        double mIout;                         // Module current at max power point [A]
        double mPower;                        // Module Power produced [W]      

        // Array Related Variables
        public double itsNSeries;             // Number of modules in series [#]
        public double itsNParallel;           // Number of series groups in parallel [#]
        public double itsNumModules;          // Number of modules in the block [#]

        // Output variables calculated
        public double IAMTGlo;                // IAM TGlo [W/m^2]
        public double IAMDir;                 // IAM Factor applied to Beam [#]
        public double IAMDif;                 // IAM Factor applied to Diffuse [#]
        public double IAMRef;                // IAM factor applied to ground reflected component of tilted irradiance [#]
        public double VOut;                   // PV array voltage at maximum power [V] 
        public double IOut;                   // PV array current at maximum power [A] 
        public double POut;                   // PV array power produced [W]  
        public double Voc;                    // PV array Voc [V]
        public double TModule;                // Temperature of module [C]
        public double SoilingLoss;            // Losses due to soiling of the PV Array [W]
        public double MismatchLoss;           // Losses due to mismatch of modules in the PV array [W]
        public double ModuleQualityLoss;      // Losses due to module quality [W]
        public double OhmicLosses;            // Losses due to Wiring between PV Array and Inverter [ohms]
        public double Efficiency;             // PV array efficiency [%] 
        public double TGloEff;                // Irradiance adjusted for incidence angle and soiling [W/m^2]
        public double itsPNomDCArray;         // The Nominal DC Power of the Array [W]
        public double itsRoughArea;           // The rough area of the DC Array [m^2]
        public double cellArea;               // The Area occupied by Cells of the DC Array [m^2]

        // PVArray Blank Constructor
        public PVArray()
        {
        }

        // Calculates the PV Array performance based on MPP or fixed voltage operation
        public void Calculate
            (
               bool isMPPT                          // Determines the operating mode of the PV Array as Fixed Voltage or MPP 
            , double InvVoltage                     // Voltage set by inverter [V], only used if MPPT status is false                                      
            )
        {

            if (isMPPT)
            {
                // If the Inverter is MPPT tracking allow PV array to produce MPP
                CalcAtMaximumPowerPoint();
            }
            else
            {
                // If the Inverter is fixing/raising the Array to an MPPT Limit/Clipping Voltage, use the inverter voltage and calculate new power point 
                mVout = InvVoltage / itsNSeries;
                CalcAtGivenVoltage(mVout, out mIout, out mPower);
            }

            // Assigning result to the output, i.e. multiplying the currents and voltages by Series and Parallel resp.
            CalcModuleToArray(isMPPT);
        }

        // Calculates the parameters required by the I-V curve equations given environmental conditions
        public void CalcIVCurveParameters
            (
              double TGlo                            // Tilted Irradiance without Shading & IAM Losses (used for efficiency calculation) 
            , double TDir                            // Tilted Beam Irradiance - Post Shading, if Shading model is defined [W/m^2]
            , double TDif                            // Tilted Diffuse Irradiance - Post Shading, if Shading model is defined [W/m^2]
            , double TRef                            // Tilted Ground Reflected Irradiance - Post Shading, if Shading model is defined [W/m^2]
            , double InciAng                         // Incidence Angle [radians]
            , double TAmbient                        // Ambient temperature [C]    
            , double WindSpeed                       // Wind speed [m/s]
            , double TModMeasured                    // Measured Module Temperature [C]
            , int MonthNumber                        // Month of the year [#, 1->12]
            )
        {
            // Assigning the Tilted Global value to a local holder (used for efficiency calculation)
            lossLessTGlo = TGlo;

            // Calculation of effective irradiance reaching the cell (Soiling and IAM accounted for)
            CalcEffectiveIrradiance(TDir, TDif, TRef, InciAng, MonthNumber);

            // Calculation of temperature GetTemperature Method used (see below)
            CalcTemperature(TAmbient, TGloEff, WindSpeed, TModMeasured);  // Using method to obtain the temperature [C]

            // Calculation of the Gamma value for given temperature (Ref 2 - Page 5) 
            itsGamma = itsGammaRef + itsGammaCoeff * (TModule - itsTref);

            //  Calculation of reverse saturation current (Ref 1 - Eq 3)
            double TModuleK = Utilities.ConvertCtoK(TModule);
            double itsTrefK = Utilities.ConvertCtoK(itsTref);
            Irs = itsIrsRef * Math.Pow(TModuleK / itsTrefK, 3) * Math.Exp(Util.ELEMCHARGE * Util.SiBANDGAP / itsGamma / Util.BOLTZMANNCONST * (1 / itsTrefK - 1 / TModuleK));

            // Calculation of the variable Rshunt based on the PVSyst model (Ref 2 - Page 6) 
            double itsRshBase = (itsRpRef - itsRshZero * Math.Exp(-itsRshExp)) / (1 - Math.Exp(-itsRshExp));     // Parametrization value introduced by PVSyst
            itsRsh = itsRshBase + (itsRshZero - itsRshBase) * Math.Exp(-itsRshExp * (TGloEff / itsHref));        // Determining the effective Rshunt based on irradiance change and a fitting parameter [ohms]

            // Module Isc calculation based on irradiance and temperature (Ref 1 - Eq 2) 
            mIPhi = TGloEff / itsHref * (itsIPhiRef + itsTCoefIsc * (TModule - itsTref));

            // Adjust Voc based on temperature (similar to current adjustment above)
            mVoc = itsVocref + itsTCoefVoc * (TModule - itsTref);
        }

        // Calculates the effective irradiance available for electricity conversion, based on IAM and Soiling Losses incurred
        void CalcEffectiveIrradiance
            (
              double TDir                            // Tilted Beam Irradiance [W/m^2]
            , double TDif                            // Tilted Diffuse Irradiance [W/m^2]
            , double TRef                            // Tilted Ground Reflected Irradiance [W/m^2]
            , double InciAng                         // Incidence Angle [radians]
            , int MonthNumber                        // Month of the Year [#]
            )
        {
            // Computing the Incidence Angle Modifier for Beam, Diffuse and Albedo Component (Calculated using ASHRAE Parameter, see Ref 5 in PV Array Class)
            if (userIAMModel)
            {
                InciAng = Math.Max(0, InciAng);
                InciAng = Math.Min(InciAng, Math.PI / 2);

                IAMDir = Interpolate.Bezier(itsUserIAMProfile[0], itsUserIAMProfile[1], InciAng * Util.RTOD, itsUserIAMProfile[0].Length);
                IAMDif = Interpolate.Bezier(itsUserIAMProfile[0], itsUserIAMProfile[1], Util.DiffInciAng * Util.RTOD, itsUserIAMProfile[0].Length);
                IAMRef = IAMDif;
            }
            else
            {
                IAMDir = Math.Cos(InciAng) > itsBo / (1 + itsBo) ? Math.Max((1 - itsBo * (1 / Math.Cos(InciAng) - 1)), 0) : 0;
                IAMDif = (1 - itsBo * (1 / Math.Cos(Util.DiffInciAng) - 1));
                IAMRef = IAMDif;
            }

            // Calculating the IAM Modified Tilted Irradiance [W/m^2]
            IAMTGlo = TDir * IAMDir + TDif * IAMDif + TRef * IAMRef;

            // Determing soiling loss based on month number specified from Time Stamp [%]
            itsSoilingLossPC = itsMonthlySoilingPC[MonthNumber];

            // Modified TGlo based on irradiance based on soiling and Incidence Angle modifier
            TGloEff = IAMTGlo * (1 - itsSoilingLossPC);
        }

        // Calculates the MPPT operating point of the module
        void CalcAtMaximumPowerPoint()
        {
            // Determining the Maximum Power Point and Voc from the IV curve (see method below)
            double iHolder = 0;                                             // Holds the value of current for any CalcAtGivenVoltage (Current is irrelevant to calculation)
            double tol = 0.0001;                                            // Tolerance on voltage solution [V]
            double lowBound = 0;                                            // Lower bound of domain to search for maximum [V]
            double highBound = mVoc;                                        // High bound of domain to search for maximum [V]
            double xL = lowBound + ((1 - Util.GOLDEN) * (highBound - lowBound)); // Section ~ 34% from lower bound [V]
            double xH = lowBound + (Util.GOLDEN * (highBound - lowBound));       // Section ~ 68% from lower bound [V]   
            double pL = 0;                                                  // Power at lower bound [W] Initializing
            CalcAtGivenVoltage(xL, out iHolder, out pL);                    // Power at lower section boundary [W]
            double pH = 0;                                                  // Power at higher bound [W] Initializing
            CalcAtGivenVoltage(xH, out iHolder, out pH);                    // Power at higher section boundary [W]

            // Begin Golden Section Search for Maximum at the lower and higher bound only if the powers are not equal
            if ((pH == 0) && (pL == 0))
            {
                // Assign outputs to 0 since the whole curve is essentially 0
                mVout = 0;
                mIout = 0;
                mPower = 0;
            }
            else
            {
                while ((xH - xL) > tol)
                {
                    if (pL > pH)
                    {
                        highBound = xH;                                    // Higher-bound will move up to older Golden Section higher-bound
                    }
                    else
                    {
                        lowBound = xL;                                     // Lower-bound will move up to Golden Section lower-bound
                    }

                    xL = lowBound + ((1 - Util.GOLDEN) * (highBound - lowBound));
                    xH = lowBound + (Util.GOLDEN * (highBound - lowBound));
                    CalcAtGivenVoltage(xL, out iHolder, out pL);
                    CalcAtGivenVoltage(xH, out iHolder, out pH);
                }

                mVout = (xL + xH) / 2;                                         // Assigning output, Vmpp
                CalcAtGivenVoltage(mVout, out mIout, out mPower);          // Assigning output, Pmpp and Impp calculated at module Vmpp
            }
        }

        // Calculates Module to Array and applies losses at the Array Level, i.e. Multiplying the currents and voltages by Series and Parallel resp.
        void CalcModuleToArray(bool isMPPT)
        {
            // Output variables are assigned their values, module voltages add in series, module currents add when parallel, powers add in both cases
            VOut = mVout * itsNSeries;
            IOut = mIout * itsNParallel * (1 - itsModuleQualityLossPC);

            // Module Quality Losses, Soiling Loses, etc are first adjusted to their before loss value, then losses are calculated.
            ModuleQualityLoss = mIout * itsNParallel * itsModuleQualityLossPC * VOut;
            SoilingLoss = mIout * itsNParallel * itsSoilingLossPC * VOut / (1 - itsSoilingLossPC);

            // The loss percentage applied is different if the array is in MPP Mode or in Fixed Operation Mode
            if (isMPPT)
            {
                MismatchLoss = IOut * itsMismatchLossSTCPC * VOut;
                IOut *= (1 - itsMismatchLossSTCPC);
            }
            else
            {
                MismatchLoss = IOut * itsMismatchFixedVLoss * VOut;
                IOut *= (1 - itsMismatchFixedVLoss);
            }

            // Calculating Ohmic Loss and assigning power out and efficiency values
            OhmicLosses = Math.Pow(IOut, 2) * itsRw;
            POut = VOut * IOut - OhmicLosses;

            // Calculating DC Efficiency
            Efficiency = lossLessTGlo > 0 ? mPower / lossLessTGlo / itsArea : 0;
        }

        // Calculate the Voc for the Module
        public void CalcAtOpenCircuit()
        {
            // Finding Voc (Using N-R method)
            double TModuleK = Utilities.ConvertCtoK(TModule);                                           // Converting the Temperature from C to K [K]
            double expArg = Util.ELEMCHARGE / (itsCellCount * itsGamma * Util.BOLTZMANNCONST * TModuleK);         // Argument for the exponential - constant
            double vocNew = mVoc;                                                                       // First guess for Voc [V]
            double vocGuess = 0;                                                                        // Iteration variable for Voc - Initialized [V]
            double vocTol = 0.0001;                                                                     // Tolerated Voc Error [V]
            double fVoc = 0;                                                                            // Voc function initialized
            double fpVoc = 0;                                                                           // Voc function derivative initialized
            int iterCounter = 0;                                                                        // Counter for Number of iterations used

            // Begin N-R iterations for the Voc
            try
            {

                while (Math.Abs(vocGuess - vocNew) > vocTol)
                {
                    vocGuess = vocNew;
                    fVoc = mIPhi - Irs * (Math.Exp(expArg * vocGuess) - 1) - vocGuess / itsRsh;
                    fpVoc = -Irs * expArg * Math.Exp(expArg * vocGuess) - 1 / itsRsh;
                    vocNew = vocGuess - fVoc / fpVoc;
                    iterCounter++;
                    if (iterCounter > Util.NRLIMIT)
                    {

                        throw new NRException("Open-circuit voltage.");
                    }
                }

                // Assigning Outputs
                Voc = vocNew * itsNSeries;
            }
            catch (NRException nr)
            {
                ErrorLogger.Log(nr, ErrLevel.WARNING);
            }
        }

        // Calculates the I and P based on the one-diode model (for a voltage value for the module)
        public void CalcAtGivenVoltage
            (
              double v                                // Voltage at which the I-V and P-V Curves should be evaluated [V]
            , out double moduleI                      // Evaluated Module Current [A]
            , out double moduleP                      // Evaluated Module Power [W]
            )
        {

            double tolerance = 0.001;                                                 // Tolerance of error [A]
            double iNew = 0;                                                          // Initializing variable
            double iGuess = itsImppref * (mIPhi / itsIscref);                         // Best guess scenario is an Isc-adjusted Impp from ref conditions
            double TModuleK = Utilities.ConvertCtoK(TModule);                         // Converting the temperature of the module from C to K [K]
            double vThermal = (Util.BOLTZMANNCONST * TModuleK) / (Util.ELEMCHARGE);   // Thermal voltage at current module temperature 
            double fI = 0;                                                            // The one-diode model solar cell equation (modified for N-R Method)
            double fPI = 0;                                                           // The derivative of the ODM (modified for N-R Method)        
            double expArg = itsCellCount * itsGamma * vThermal;                       // Arguments in the exponential which remain constant
            int iterCounter = 0;                                                      // Counter for the number of iterations used

            //  Starting the N-R method to iteratively solve for the solar cell current
            try
            {
                while (Math.Abs(iNew - iGuess) > tolerance)                           // Compare two values and check if within tolerance  
                {
                    iGuess = iNew;                                                    // Calculating the new current value
                    fI = mIPhi - iGuess - Irs * (Math.Exp((v + (iGuess * (itsRs + itsRw))) / expArg) - 1) - (v + iGuess * (itsRs + itsRw)) / itsRsh;
                    fPI = -((Irs * itsRs) / expArg) * Math.Exp((v + (iGuess * (itsRs + itsRw))) / expArg) - (itsRs + itsRw) / itsRsh - 1;
                    iNew = iGuess - fI / fPI;
                    iterCounter++;
                    if (iterCounter > Util.NRLIMIT)
                    {

                        throw new NRException("Current at a given voltage.");
                    }
                }
            }
            catch (NRException nr)
            {
                ErrorLogger.Log(nr, ErrLevel.WARNING);
                iNew = 0;
            }

            // Assigning Outputs to the variables
            moduleI = iNew;
            moduleP = (moduleI * v);
        }

        // Calculates the Module Temperature (in C, in K or keep Measured based on user preferences found from .CSYX)
        void CalcTemperature
            (
              double TAmbient                         // Ambient temperature [C]
            , double TGloAkt                          // Effective Irradiance Value [W/m^2]               
            , double WindSpeed                        // Wind speed [m/s]
            , double TModMeasured                     // Measured temperature of the module [C]  
            )
        {

            // Beginning of Temperature calculation model (see Ref 4)
            // Enables user to choose the temperature model, or use measured values
            if (useMeasuredTemp)
            {
                // Checks if the Measured Value of Panel Temperature is available.
                if (double.IsNaN(TModMeasured))
                {
                    throw new CASSYSException("CASSYS: Measured module temperature is not defined. This is a required value for the selected temperature model.");
                }
                else
                {
                    // Use measured values if available.
                    TModule = TModMeasured;
                }
            }
            else
            {


                // NB: throw an error so temp. model does not error
                if (itsConstHTC <= 0)
                {
                    throw new CASSYSException("CASSYS: Constant Heat Loss Factor is negative or is not specified for Temperature Model. This is required as a positive value for the selected temperature model.");
                }
                
                // Checks if WindSpeed is available, forgives if the Convective Heat Transfer Constant is 0.
                else if (double.IsNaN(WindSpeed) && (itsConvHTC != 0))
                {
                    throw new CASSYSException("CASSYS: Wind Speed is not specified for Temperature Model. This is a required value for the selected temperature model.");
                }
                else if (itsConvHTC == 0 && double.IsNaN(WindSpeed))
                {
                    // Forgives the fact that no Wind Speed was specified.
                    WindSpeed = 0;
                }
                
                
                // Calculate temperature based on values provided by the User
                TModule = TAmbient + itsAdsorp * TGloAkt * (1 - itsEfficiencyRef) / (itsConstHTC + itsConvHTC * WindSpeed); // Faiman's module temperature model

             }
        }

        // Config will assign parameter variables their values as obtained from the .CSYX file
        public void Config
            (
              int ArrayNum                            // SubArray Number as determined in the Main Program [#]
            )
        {
            // Gathering all the parameters for the PV Array
            itsSubArrayNum = ArrayNum;
            itsNSeries = int.Parse(ReadFarmSettings.GetInnerText("PV", "ModulesInString", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            itsNParallel = int.Parse(ReadFarmSettings.GetInnerText("PV", "NumStrings", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            itsArea = double.Parse(ReadFarmSettings.GetInnerText("PV", "AreaM", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            itsPNom = double.Parse(ReadFarmSettings.GetInnerText("PV", "Pnom", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            itsVmppref = double.Parse(ReadFarmSettings.GetInnerText("PV", "Vmpp", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            itsImppref = double.Parse(ReadFarmSettings.GetInnerText("PV", "Impp", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            itsVocref = double.Parse(ReadFarmSettings.GetInnerText("PV", "Voc", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            itsIscref = double.Parse(ReadFarmSettings.GetInnerText("PV", "Isc", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            itsTref = double.Parse(ReadFarmSettings.GetInnerText("PV", "Tref", _ArrayNum: ArrayNum, _Error: ErrLevel.WARNING, _default: "25"));
            itsHref = double.Parse(ReadFarmSettings.GetInnerText("PV", "Gref", _ArrayNum: ArrayNum, _Error: ErrLevel.WARNING, _default: "1000"));
            itsCellCount = int.Parse(ReadFarmSettings.GetInnerText("PV", "CellsinS", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            itsTCoefIsc = double.Parse(ReadFarmSettings.GetInnerText("PV", "mIsc", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL)) / 1000;
            itsTCoefVoc = double.Parse(ReadFarmSettings.GetInnerText("PV", "mVco", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL)) / 1000;
            itsRpRef = double.Parse(ReadFarmSettings.GetInnerText("PV", "Rshunt", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            itsRs = double.Parse(ReadFarmSettings.GetInnerText("PV", "Rserie", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            itsRshZero = double.Parse(ReadFarmSettings.GetInnerText("PV", "Rsh0", _ArrayNum: ArrayNum, _Error: ErrLevel.WARNING, _default: (4*itsRpRef).ToString()));
            itsRshExp = double.Parse(ReadFarmSettings.GetInnerText("PV", "Rshexp", _ArrayNum: ArrayNum, _Error: ErrLevel.WARNING, _default: "5.5"));
            itsTCoefP = double.Parse(ReadFarmSettings.GetInnerText("PV", "mPmpp", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            cellArea = itsNParallel * itsNSeries * itsCellCount * double.Parse(ReadFarmSettings.GetInnerText("PV", "Cellarea", _ArrayNum: ArrayNum, _Error: ErrLevel.WARNING, _default: "0.01")) / 10000;

            // Defining all thermal loss variables for the PV Array
            useMeasuredTemp = Convert.ToBoolean(ReadFarmSettings.GetInnerText("Losses", "ThermalLosses/UseMeasuredValues", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            if (!useMeasuredTemp)
            {
                itsConstHTC = double.Parse(ReadFarmSettings.GetInnerText("Losses", "ThermalLosses/ConsHLF", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
                itsConvHTC = double.Parse(ReadFarmSettings.GetInnerText("Losses", "ThermalLosses/ConvHLF", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL));
            }

            // Defining the Module Quality Losses
            itsModuleQualityLossPC = double.Parse(ReadFarmSettings.GetInnerText("Losses", "ModuleQualityLosses/EfficiencyLoss", _ArrayNum: ArrayNum, _Error: ErrLevel.WARNING, _default: "0"));

            // Defining the Mismatch Losses
            itsMismatchLossSTCPC = double.Parse(ReadFarmSettings.GetInnerText("Losses", "ModuleMismatchLosses/PowerLoss", _ArrayNum: ArrayNum, _Error: ErrLevel.WARNING, _default: "0"));
            itsMismatchFixedVLoss = double.Parse(ReadFarmSettings.GetInnerText("Losses", "ModuleMismatchLosses/LossFixedVoltage", _ArrayNum: ArrayNum, _Error: ErrLevel.WARNING, _default: "0"));

            // Defining the Incidence Angle Modifier
            if (ReadFarmSettings.CASSYSCSYXVersion == "0.9")
            {
                itsBo = double.Parse(ReadFarmSettings.GetInnerText("Losses", "IncidenceAngleModifier/bNaught", _ArrayNum: ArrayNum, _Error: ErrLevel.WARNING, _default: "0.05"));
            }
            else
            {
                if ((ReadFarmSettings.GetAttribute("Losses", "IAMSelection", _Adder: "/IncidenceAngleModifier", _VersionNum: "0.9.1", _ArrayNum: ArrayNum) == "ASHRAE"))
                {
                    itsBo = double.Parse(ReadFarmSettings.GetInnerText("Losses", "IncidenceAngleModifier/bNaught", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL, _default: "0.05"));
                }
                else
                {
                    List<double> UserAOI = new List<double>();
                    List<double> UserIAM = new List<double>();

                    double placeholder;

                    // Using the Angle of incidence as an iteration variable
                    // Gettings values from the IAM list
                    for (int AOI = 0; AOI <= 90; AOI +=5)
                    {
                        userIAMModel = double.TryParse(ReadFarmSettings.GetInnerText("Losses", "IncidenceAngleModifier/IAM_" + AOI.ToString(), _VersionNum: "0.9.1", _ArrayNum: ArrayNum, _default: null), out placeholder);

                        if (userIAMModel)
                        {
                            UserAOI.Add(Convert.ToDouble(AOI));
                            UserIAM.Add(placeholder);
                        }
                    }

                    itsUserIAMProfile[0] = UserAOI.ToArray();
                    itsUserIAMProfile[1] = UserIAM.ToArray();
                }
            }

            // If soiling losses are defined on a monthly basis, then populate an array with values for each month
            // Based on month number
            if (ReadFarmSettings.GetAttribute("Losses", "Frequency", _Adder: "/SoilingLosses") == "Monthly")
            {
                // Initializing the array for the soiling losses; index numbers correspond to the months (therefore index 0 must be included in array size but is not used for assigning values
                itsMonthlySoilingPC = new double[13];

                // Using the month number as the index, populate the Soiling Loss values from each corresponding node
                itsMonthlySoilingPC[1] = double.Parse(ReadFarmSettings.GetInnerText("Losses", "SoilingLosses/Jan", ErrLevel.WARNING, _default: "0.01"));
                itsMonthlySoilingPC[2] = double.Parse(ReadFarmSettings.GetInnerText("Losses", "SoilingLosses/Feb", ErrLevel.WARNING, _default: "0.01"));
                itsMonthlySoilingPC[3] = double.Parse(ReadFarmSettings.GetInnerText("Losses", "SoilingLosses/Mar", ErrLevel.WARNING, _default: "0.01"));
                itsMonthlySoilingPC[4] = double.Parse(ReadFarmSettings.GetInnerText("Losses", "SoilingLosses/Apr", ErrLevel.WARNING, _default: "0.01"));
                itsMonthlySoilingPC[5] = double.Parse(ReadFarmSettings.GetInnerText("Losses", "SoilingLosses/May", ErrLevel.WARNING, _default: "0.01"));
                itsMonthlySoilingPC[6] = double.Parse(ReadFarmSettings.GetInnerText("Losses", "SoilingLosses/Jun", ErrLevel.WARNING, _default: "0.01"));
                itsMonthlySoilingPC[7] = double.Parse(ReadFarmSettings.GetInnerText("Losses", "SoilingLosses/Jul", ErrLevel.WARNING, _default: "0.01"));
                itsMonthlySoilingPC[8] = double.Parse(ReadFarmSettings.GetInnerText("Losses", "SoilingLosses/Aug", ErrLevel.WARNING, _default: "0.01"));
                itsMonthlySoilingPC[9] = double.Parse(ReadFarmSettings.GetInnerText("Losses", "SoilingLosses/Sep", ErrLevel.WARNING, _default: "0.01"));
                itsMonthlySoilingPC[10] = double.Parse(ReadFarmSettings.GetInnerText("Losses", "SoilingLosses/Oct", ErrLevel.WARNING, _default: "0.01"));
                itsMonthlySoilingPC[11] = double.Parse(ReadFarmSettings.GetInnerText("Losses", "SoilingLosses/Nov", ErrLevel.WARNING, _default: "0.01"));
                itsMonthlySoilingPC[12] = double.Parse(ReadFarmSettings.GetInnerText("Losses", "SoilingLosses/Dec", ErrLevel.WARNING, _default: "0.01"));
            }
            // If soiling losses are defined with one value for the whole year then use that value for each month
            else
            {
                // Initializing the expected list
                itsMonthlySoilingPC = new double[13];
                itsMonthlySoilingPC[1] = double.Parse(ReadFarmSettings.GetInnerText("Losses", "SoilingLosses/Yearly", ErrLevel.WARNING, _default: "0.01"));

                // Apply the same soiling percentage to all months
                for (int i = 3; i < itsMonthlySoilingPC.Length + 1; i++)
                {
                    itsMonthlySoilingPC[i - 1] = itsMonthlySoilingPC[1];
                }
            }

            // Calculating and defining the efficiency for the module at Ref conditions
            itsEfficiencyRef = (itsImppref * itsVmppref) / (itsArea * itsHref);
            itsPNomDCArray = itsPNom * itsNParallel * itsNSeries;
            itsRoughArea = itsArea * itsNSeries * itsNParallel;
            itsNumModules = itsNParallel * itsNSeries;
            CalcGammaIPhiIrsRef();

            // Gathering wiring losses from the CSYX File
            itsRw = double.Parse(ReadFarmSettings.GetInnerText("PV", "GlobWireResist", _ArrayNum: ArrayNum, _Error: ErrLevel.WARNING, _default: "1")) / 1000;
        }

        // Calculate Gamma, IrsRef, IphiRef for the module provided using equations for Impp condition and Voc condition with N-R method to calculate Gamma
        void CalcGammaIPhiIrsRef()
        {

            // Constants (cX - X=1,2,3 - 1 - Isc condition, 2 - Pmpp Condition, 3 - Voc Condition)
            double c1 = -itsIscref - (itsIscref * (itsRs / itsRpRef));
            double c2 = -itsImppref - (itsVmppref + (itsImppref * itsRs)) / itsRpRef;
            double c3 = -itsVocref / itsRpRef;

            // Exponential arguments (dX - X=1,2,3 - 1 - Isc condition, 2 - Pmpp Condition, 3 - Voc Condition)
            double itsTrefK = Utilities.ConvertCtoK(itsTref);
            double d1 = Util.ELEMCHARGE * itsIscref * itsRs / (Util.BOLTZMANNCONST * itsCellCount * itsTrefK);
            double d2 = Util.ELEMCHARGE * (itsVmppref + (itsImppref * itsRs)) / (Util.BOLTZMANNCONST * itsCellCount * itsTrefK);
            double d3 = Util.ELEMCHARGE * itsVocref / (Util.BOLTZMANNCONST * itsCellCount * itsTrefK);

            // Using N-R method to solve for GammaRef
            double gInit = 1;                                                           // Initial guess for GAMMA (To track changes on non-convergence)  
            double gammaTol = 0.0001;                                                   // Tolerance for the Gamma value
            double gNew = 1;                                                            // First guess for the new Gamma value
            double gGuess = 0;                                                          // Guess and evaluation variable
            double fGamma = 0;                                                          // Function of Gamma initialized        
            double fPGamma = 0;                                                         // Derivative of fGamma
            double counter = 0;                                                         // Iteration counter to check for divergence

            // Begin N-R method to solve for gamma by defining its function and derivative of the function
            try
            {
                do
                {
                    gGuess = gNew;
                    fGamma = (c2 - c3) * (Math.Exp(d1 / gGuess))
                            + (c3 - c1) * (Math.Exp(d2 / gGuess))
                            + (c1 - c2) * (Math.Exp(d3 / gGuess));

                    fPGamma = (c2 - c3) * -d1 * (Math.Exp(d1 / gGuess)) / Math.Pow(gGuess, 2)
                            + (c3 - c1) * -d2 * (Math.Exp(d2 / gGuess)) / Math.Pow(gGuess, 2)
                            + (c1 - c2) * -d3 * (Math.Exp(d3 / gGuess)) / Math.Pow(gGuess, 2);
                    gNew = gGuess - (fGamma / fPGamma);
                    counter++;
                    if (counter > Util.NRLIMIT)
                    {
                        // If the first time Gamma does not converge, change the initial guess to try another method.
                        // Try gInit = 0.5
                        if (gInit == 1)
                        {
                            gNew = 0.5;
                            gInit = 0.5;
                            gGuess = 0;
                            fGamma = 0;
                            fPGamma = 0;
                            counter = 0;

                        }
                        // Try gInit = 1.5
                        else if (gInit == 0.5)
                        {
                            gNew = 1.5;
                            gInit = 1.5;
                            gGuess = 0;
                            fGamma = 0;
                            fPGamma = 0;
                            counter = 0;
                        }
                        // Try gInit = 0
                        else if (gInit == 1.5)
                        {
                            gNew = 0;
                            gInit = 0;
                            gGuess = 0;
                            fGamma = 0;
                            fPGamma = 0;
                            counter = 0;
                        }
                        // If it does not converge in any of these cases, Log the error and end simulation.
                        else
                        {
                            ErrorLogger.Log("Newton-Raphson Method did not converge for Gamma. Simulation has ended.", ErrLevel.FATAL);
                        }
                    }
                }
                while (Math.Abs(gNew - gGuess) > gammaTol);
            }
            catch (ArithmeticException ae)
            {
                ErrorLogger.Log(ae, ErrLevel.FATAL);
            }
            catch (NRException nr)
            {
                ErrorLogger.Log(nr, ErrLevel.FATAL);
            }

            // Assigning output value to Gamma, and calculating the value for IPhiRef, and IrsRef using Case 1 and Case 2
            itsGammaRef = Math.Round(gNew, 3);
            itsIrsRef = Utilities.Truncate((c1 - c2) / (Math.Exp(d1 / itsGammaRef) - Math.Exp(d2 / itsGammaRef)), 4);
            itsIPhiRef = Math.Round(itsIrsRef * (Math.Exp(d1 / itsGammaRef) - 1) - c1, 3);

            // Calculate the Gamma Temperature Coefficient Parameter for the module
            CalcGammaCoeff();
        }

        // Calculate GammaCoefficient Parameters for the module using the Temperature Coefficient for Power provided in the database. 
        // This is done using the Bisection method for Gamma. For every change in Gamma, the reverse saturation current must be calculated as well 
        // as it is dependent on Gamma.
        // Calculation of power is done at reference Irradiance but 2 * Reference Temperature.
        void CalcGammaCoeff()
        {
            // The temperature for which Gamma must be calculated.
            double TrialTModule = itsTref + 25;
            double tolerance = 1 / 10000D;

            // Converting to Kelvin and assigning TModule to the correct value.
            double TrialTModuleK = Utilities.ConvertCtoK(TrialTModule);
            double itsTrefK = Utilities.ConvertCtoK(itsTref);
            TModule = TrialTModule;

            // Module Iphi calculation based on irradiance and temperature (Ref 1 - Eq 2) 
            mIPhi = itsIPhiRef + itsTCoefIsc * (TModule - itsTref);

            // Adjust Voc based on temperature (similar to current adjustment above)
            mVoc = itsVocref + itsTCoefVoc * (TModule - itsTref);

            // The shunt resistance will be at reference levels.
            itsRsh = itsRpRef;

            // Variables required to use the bisection method to match gamma to the target Pnom
            double targetPNom = itsPNom * (1 + itsTCoefP/100 * (TrialTModule - itsTref));         // The target power is Pnom corrected to TrialTModule using uPMPP
            double gammaL = itsGammaRef - 0.5D;                                                     // Lower bound: The uGamma is a small adjustment for TModule, so Gamma should lie within +/- 0.5 of the refGamma    
            double gammaH = itsGammaRef + 0.5D;                                                     // Higher bound: The uGamma is a small adjustment for TModule, so Gamma should lie within +/- 0.5 of the refGamma

            // Set the value of gamma L to 0.1 if the gammaL is negative. This avoids any unrealistic values for GammaL and allows user to proceed with analysis.
            if (gammaL < 0)
            {
                gammaL = 0.1;
            }
            
            double powGammaL = 0;                                                                   // Power at lower value of Gamma
            double powGammaH = 0;                                                                   // Power at higher value of Gamma
            double gammaTrial = 0;                                                                  // Gamma value at bisection
            double powGammaTrial = 0;                                                               // Power at bisection value of Gamma

            // Calculating the Lower Gamma boundary power value and calculation of reverse saturation current (Ref 1 - Eq 3)
            itsGamma = gammaL;
            Irs = itsIrsRef * Math.Pow(TrialTModuleK / itsTrefK, 3) * Math.Exp(Util.ELEMCHARGE * Util.SiBANDGAP / itsGamma / Util.BOLTZMANNCONST * (1 / itsTrefK - 1 / TrialTModuleK));
            CalcAtMaximumPowerPoint();
            powGammaL = mPower;

            // Calculating the higher Gamma boundary power value and calculation of reverse saturation current (Ref 1 - Eq 3)
            itsGamma = gammaH;
            Irs = itsIrsRef * Math.Pow(TrialTModuleK / itsTrefK, 3) * Math.Exp(Util.ELEMCHARGE * Util.SiBANDGAP / itsGamma / Util.BOLTZMANNCONST * (1 / itsTrefK - 1 / TrialTModuleK));
            CalcAtMaximumPowerPoint();
            powGammaH = mPower;

            // Checking to make sure bisection actual takes place by implementing 
            if ((powGammaL - itsPNom) * (powGammaH - itsPNom) < 0)
            {
                // Iterate to find where the two gamma match the target power as long as the difference is above the tolerance level
                while (Math.Abs(gammaH - gammaL) > tolerance)
                {
                    // Assigning the new trial value
                    gammaTrial = (gammaL + gammaH) / 2;

                    // Determining the maximum power obtained the new trial Gamma
                    itsGamma = gammaTrial;
                    Irs = itsIrsRef * Math.Pow(TrialTModuleK / itsTrefK, 3) * Math.Exp(Util.ELEMCHARGE * Util.SiBANDGAP / itsGamma / Util.BOLTZMANNCONST * (1 / itsTrefK - 1 / TrialTModuleK));
                    CalcAtMaximumPowerPoint();
                    powGammaTrial = mPower;

                    // A higher than target Pnom value of powGammaTrial implies your gamma must be reduced,
                    // therefore the boundary must move to a lower window.
                    if (powGammaTrial < targetPNom)
                    {
                        gammaL = gammaTrial;
                    }
                    else
                    {
                        gammaH = gammaTrial;
                    }
                }
            }
            else
            {
                ErrorLogger.Log("The calculation for Gamma did not have the correct boundries. CASSYS cannot configure the module, and has stopped.", ErrLevel.FATAL);
            }

            // Determining the resulting coefficient. Linear assumption allows a calculation with one temperature change to determine the coefficient.
            itsGammaCoeff = (gammaTrial - itsGammaRef) / (TrialTModule - itsTref);
        }        
    }
}