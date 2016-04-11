// CASSYS - Grid connected PV system modelling software 
// Version 0.9  
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Inverter Class
// 
// Revision History:
// AP - 2014-10-14: Version 0.9
//
// Description: 
// The Inverter class uses one or three efficiency curve(s) to determine the output
// of an inverter given power (DC, W) and voltage (DC, V) from a PV Array. 
// The inverter can be bipolar or unipolar, and its status pertaining to MPPT,  
// clipDCPwrIng, etc are checked using methods and conditions pertaining to each case
// in the main program.                          
//   
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
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace CASSYS
{
    class Inverter
    {
        // Inverter definition variables
        bool itsMPPTracking;                                // MppTracking flag [TRUE, FALSE]
        double itsMppWindowMin;                             // Minimum value for Mppt window [V]
        double itsMppWindowMax;                             // Maximum value for Mppt window [V]
        public int itsNumInverters;                         // Number of same inverters supporting array [#]
        public double itsThresholdPwr;                      // Power below which inverter cuts off [W]     
        public double itsMinVoltage;                        // Minimum voltage required for the Inverter to turn ON [V]
        public double itsMaxVoltage;                        // Maximumum voltage allowed for the Inverter to work, if the Inverter crosses this limit, the Inverter will cut or limit
        public double itsOutputVoltage;                     // Output voltage value for Inverter [V]
        public int outputPhases;                            // The number of phases at the inverter output [#]

        // Efficiency curves related variables
        bool threeCurves;                                   // True if Inverter has three efficiency curves, false if Inverter has one efficiency curve
        double itsLowVoltage;                               // The voltage threshold for low voltage efficiency curve [V]
        double itsMedVoltage;                               // The voltage threshold for medium voltage efficiency curve [V]
        double itsHighVoltage;                              // The voltage threshold for high voltage efficiency curve [V]
        double[][] itsPresentEfficiencies = new double[2][];// Current Efficiency curve chosen calculated for different voltage values [%]
        public double[][] itsLowEff = new double[2][];      // Low voltage efficiency curve values [P DC in, %]
        double[][] itsMedEff = new double[2][];             // Medium voltage efficiency curve values [P DC in, %]
        double[][] itsHighEff = new double[2][];            // High voltage efficiency curve values [P DC in, %]
        public double[][] itsOnlyEff = new double[2][];     // Its only efficiency curve [P DC in, %]

        // Inverter wiring losses related variable
        double itsACWiringLossPC;                           // The AC wiring loss specified as a percentage [%]
        double itsACWiringRes;                              // The AC wiring loss translated from a percentage to a Resistance [ohms]

        // Inverter control variables (Used to determine the ClipDCPwrIng & ON status)
        public double itsNomOutputPwr;                      // Nominal output power [W AC]          
        public bool hasMinVoltage = false;                  // Determines if the Inverter has enough voltage to turn ON [default false]
        public bool isClipping = false;                     // Determines if the Inverter is ClipDCPwrIng or not [default false]
        public bool isON = false;                           // ON/OFF state of the inverter [default OFF]
        public bool inMPPTWindow = false;                   // Boolean to define if it is in MPPT Window [default false]
        public bool isBipolar;                              // Determines if the Inverter is BiPolar or not

        // Output variables calculated
        public double ACPwrOut;                             // AC Power delivered by Inverter [W]  
        public double Losses;                               // Losses from inverter [W]
        public double Efficiency;                           // Actual Efficiency of the inverter [%]
        public double VInDC;                                // Voltage, DC side of the Inverter
        public double IOut;                                 // Current Output of the Inverter [A, AC Single Phase]
        public double ACWiringLoss;                         // AC Wiring Loss incurred [W]
        public double itsPNomArrayAC;                       // Nominal AC Production of the Array [W]
        public double LossPMinThreshold;                    // Loss when the power of the array is not sufficient for starting the inverter. 
        public double LossClipping;                         // Produced power before reduction by Inverter (clipping) [W]

        // Inverter constructor
        public Inverter
        (
        )
        {
        }

        // Calculation for inverter output power, using efficiency curve
        public void Calculate
            (
              double DCPwrIn             // DC Power in from PVArray
            , double Vin                 // Inverter's Voltage (as determined during GetInverterStatus)
            )
        {
            // To combat negative DCPwrIn and Voltage Values in case they ever occur
            DCPwrIn = Math.Max(DCPwrIn, 0);
            Vin = Math.Max(Vin, 0);

            // Determining the Efficiency of the Inverter based on whether three or one single curve is used
            // to specify the efficiency of the Inverter
            if (threeCurves)
            {
                // Changing the voltage for evaluating the Efficiency curve if the Inverter is Bipolar
                if (isBipolar)
                {
                    Vin /= 2;
                }

                // Calculating the Efficiency Value for given DCPwrIn and Efficiency
                // Obtaining the Efficiency value for each efficiency curve provided on the basis of Power In (DC)          
                itsPresentEfficiencies[1][0] = Interpolate.Linear(itsLowEff[0], itsLowEff[1], DCPwrIn);
                itsPresentEfficiencies[1][1] = Interpolate.Linear(itsMedEff[0], itsMedEff[1], DCPwrIn);
                itsPresentEfficiencies[1][2] = Interpolate.Linear(itsHighEff[0], itsHighEff[1], DCPwrIn);

                // Obtaining the efficiency given the VOut of the Inverter
                Efficiency = Interpolate.Quadratic(itsPresentEfficiencies[0], itsPresentEfficiencies[1], Vin);
            }
            else
            {
                // Calculating the Efficiency Value for given DCPwrIn and Efficiency
                // Obtaining the Efficiency value using the Efficiency Curve on the basis of Power In (DC)
                Efficiency = Interpolate.Linear(itsOnlyEff[0], itsOnlyEff[1], DCPwrIn);
            }

            // Calculating AC Power Out of the Inverter
            ACPwrOut = Efficiency * DCPwrIn;

            // Check and set clipDCPwrIng status of the Inverter using the calculated AC Output Power
            isClipping = (ACPwrOut >= (itsPNomArrayAC));

            // Losses are difference between input value and output value
            Losses = (DCPwrIn - ACPwrOut);

            // Calculating the current to determine the wiring losses that follow (three phase output assumed)
            IOut = ACPwrOut / itsOutputVoltage / Math.Sqrt(outputPhases);

            // AC wiring losses using the current caculated (three phase output assumed)
            ACWiringLoss = Math.Sqrt(outputPhases) * Math.Pow(IOut, 2) * itsACWiringRes;
        }

        // Check if the PV Array voltage is within the MPPT Window of the Inverter
        public void GetMPPTStatus(double arrayV, out bool MPPTStatus)
        {
            // Default value for if the Inverter is in the MPPT Window
            MPPTStatus = false;

            if (itsMPPTracking)
            {
                if (arrayV < itsMppWindowMin)
                {
                    if (isBipolar)
                    {
                        VInDC = itsMppWindowMin * 2;     // Less than or equal to voltage window minimum, fix to voltage minimum
                    }
                    else
                    {
                        VInDC = itsMppWindowMin;        // Less than or equal to voltage window minimum, fix to voltage minimum
                    }

                    MPPTStatus = false;
                }
                else if (arrayV > itsMppWindowMax)
                {
                    if (isBipolar)
                    {
                        VInDC = itsMppWindowMax * 2;    // Less than or equal to voltage window minimum, fix to voltage minimum
                    }
                    else
                    {
                        VInDC = itsMppWindowMax;        // Less than or equal to voltage window minimum, fix to voltage minimum
                    }

                    MPPTStatus = false;
                }
                // If in between the voltage window, the voltage stays as received
                else
                {
                    if (isBipolar)
                    {
                        VInDC = arrayV * 2;
                    }
                    else
                    {
                        VInDC = arrayV;
                    }
                    MPPTStatus = true;
                }
            }
        }

        // Config will assign parameter variables their values as obtained from the .CSYX file              
        public void Config(int ArrayNum, XmlDocument doc)
        {
            itsNomOutputPwr = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "PNomAC", _ArrayNum: ArrayNum)) * 1000;
            itsMinVoltage = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Min.V", _ArrayNum: ArrayNum));
            itsMaxVoltage = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Max.V", _ArrayNum: ArrayNum, _default: itsMppWindowMax.ToString()));
            itsNumInverters = int.Parse(ReadFarmSettings.GetInnerText("Inverter", "NumInverters", _ArrayNum: ArrayNum));
            itsACWiringLossPC = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "LossFraction", _ArrayNum: ArrayNum));
            itsOutputVoltage = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Output", _ArrayNum: ArrayNum));
            itsPNomArrayAC = itsNomOutputPwr * itsNumInverters;
            
            // Assigning the number of output phases;
            if (ReadFarmSettings.GetInnerText("Inverter", "Type", _ArrayNum: ArrayNum) == "Tri")
            {
                // Tri-phase Inverter
                outputPhases = 3;
            }
            else if (ReadFarmSettings.GetInnerText("Inverter", "Type", _ArrayNum: ArrayNum) == "Bi")
            {
                // Bi-phase inverter
                outputPhases = 2;
            }
            else if (ReadFarmSettings.GetInnerText("Inverter", "Type", _ArrayNum: ArrayNum) == "Mono")
            {
                outputPhases = 1;
            }
            else
            {
                ErrorLogger.Log("The Inverter output phase definition was not found. Please check the Inverter in the Inverter database.", ErrLevel.FATAL);
            }

            // Assigning the operation type of the Inverter
            if (ReadFarmSettings.GetInnerText("Inverter", "Oper.", _ArrayNum: ArrayNum) == "MPPT")
            {
                itsMPPTracking = true;
                itsMppWindowMin = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "MinMPP", _ArrayNum: ArrayNum));
                itsMppWindowMax = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "MaxMPP", _ArrayNum: ArrayNum));
            }
            else
            {
                ErrorLogger.Log("The Inverter does not operate in MPPT according to the configurations. CASSYS does not support these inverters at this time. Simulation has ended.", ErrLevel.FATAL);
                itsMPPTracking = false;
            }

            // Assigning if the Inverter is Bipolar or not
            if ((ReadFarmSettings.GetInnerText("Inverter", "BipolarInput", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL) == "Yes") || (ReadFarmSettings.GetInnerText("Inverter", "BipolarInput", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL) == "True") || (ReadFarmSettings.GetInnerText("Inverter", "BipolarInput", _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL) == "Bipolar inputs"))
            {
                isBipolar = true;
                itsThresholdPwr = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Threshold", _ArrayNum: ArrayNum));
            }
            else
            {
                isBipolar = false;
                itsThresholdPwr = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Threshold", _ArrayNum: ArrayNum));
            }

            // Inverter Efficiency Curve Configuration
            threeCurves = Convert.ToBoolean(ReadFarmSettings.GetInnerText("Inverter", "MultiCurve", _ArrayNum: ArrayNum));

            // Initialization of efficiency curve arrays
            itsLowEff[0] = new double[8];
            itsLowEff[1] = new double[8];
            itsMedEff[0] = new double[8];
            itsMedEff[1] = new double[8];
            itsHighEff[0] = new double[8];
            itsHighEff[1] = new double[8];

            // Initialization of efficiency curve arrays
            itsOnlyEff[0] = new double[8];
            itsOnlyEff[1] = new double[8];

            // Configuration of the Efficiency Curves for the Inverter
            ConfigEffCurves(doc, ArrayNum, threeCurves);
        }

        // Obtaining and Setting efficiency curve values from .CSYX file
        public void ConfigEffCurves(XmlDocument doc, int ArrayNum, bool threeCurves)
        {
            // Configuration begins with a check if the Inverter has three efficiency curves or just one
            // Once determined, relevant values are collected from the .CSYX file
            if (threeCurves)
            {
                // Initiating an Array to hold Interpolated Values and the three voltage levels
                itsPresentEfficiencies[0] = new double[3];             // To hold the three voltage levels [Array, V]
                itsPresentEfficiencies[1] = new double[3];             // To hold the three efficiencies from interpolation [Array, %, Computed in Calculate Method]

                // Getting the Low, Medium and High Voltage Values used in the Curve
                itsLowVoltage = double.Parse(ReadFarmSettings.GetAttribute("Inverter", "Voltage", _Adder: "/Efficiency/Low", _ArrayNum: ArrayNum));
                itsPresentEfficiencies[0][0] = itsLowVoltage;
                itsMedVoltage = double.Parse(ReadFarmSettings.GetAttribute("Inverter", "Voltage", _Adder: "/Efficiency/Med", _ArrayNum: ArrayNum));
                itsPresentEfficiencies[0][1] = itsMedVoltage;
                itsHighVoltage = double.Parse(ReadFarmSettings.GetAttribute("Inverter", "Voltage", _Adder: "/Efficiency/High", _ArrayNum: ArrayNum));
                itsPresentEfficiencies[0][2] = itsHighVoltage;

                // Setting the lowest efficiency (i.e. at Threshold Power)
                itsLowEff[0][0] = itsThresholdPwr * itsNumInverters;
                itsLowEff[1][0] = 0;
                itsMedEff[0][0] = itsThresholdPwr * itsNumInverters;
                itsMedEff[1][0] = 0;
                itsHighEff[0][0] = itsThresholdPwr * itsNumInverters;
                itsHighEff[1][0] = 0;

                // Get Low Voltage Efficiency Curve
                for (int eff = 1; eff < itsLowEff[0].Length; eff++)
                {
                    ErrorLogger.Assert("The Inverter Efficiency Curve is Incorrectly defined. Check Inverter in Sub-Array " + ArrayNum + ".", double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Efficiency/Low/Effic" + (eff + 1).ToString(), _ArrayNum: ArrayNum)) < 100, ErrLevel.FATAL);
                    itsLowEff[0][eff] = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Efficiency/Low/IN" + (eff + 1).ToString(), _ArrayNum: ArrayNum)) * itsNumInverters * 1000;
                    itsLowEff[1][eff] = itsLowEff[0][eff]*double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Efficiency/Low/Effic" + (eff + 1).ToString(), _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL)) / 100;
                }

                // Get Med Voltage Efficiency Curve
                for (int eff = 1; eff < itsMedEff[0].Length; eff++)
                {
                    ErrorLogger.Assert("The Inverter Efficiency Curve is Incorrectly defined. Check Inverter in Sub-Array " + ArrayNum + ".", double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Efficiency/Med/Effic" + (eff + 1).ToString(), _ArrayNum: ArrayNum)) < 100, ErrLevel.FATAL);
                    itsMedEff[0][eff] = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Efficiency/Med/IN" + (eff + 1).ToString(), _ArrayNum: ArrayNum)) * itsNumInverters * 1000;
                    itsMedEff[1][eff] = itsMedEff[0][eff] * double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Efficiency/Med/Effic" + (eff + 1).ToString(), _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL)) / 100;
                }

                // Get High Voltage Efficiency Curve
                for (int eff = 1; eff < itsHighEff[0].Length; eff++)
                {
                    ErrorLogger.Assert("The Inverter Efficiency Curve is Incorrectly defined. Check Inverter in Sub-Array " + ArrayNum + ".", double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Efficiency/High/Effic" + (eff + 1).ToString(), _ArrayNum: ArrayNum)) < 100, ErrLevel.FATAL);
                    itsHighEff[0][eff] = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Efficiency/High/IN" + (eff + 1).ToString(), _ArrayNum: ArrayNum)) * itsNumInverters * 1000;
                    itsHighEff[1][eff] = itsHighEff[0][eff] * double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Efficiency/High/Effic" + (eff + 1).ToString(), _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL)) / 100;
                }
            }
            else
            {
                // The Inverter only has one Efficiency Curve
                // Setting the lowest efficiency (i.e. at Threshold Power)
                itsOnlyEff[0][0] = itsThresholdPwr * itsNumInverters;
                itsOnlyEff[1][0] = 0;

                // Go through all .CSYX Nodes with Efficiency values and assign them into Array
                for (int eff = 1; eff < itsOnlyEff[0].Length; eff++)
                {
                    ErrorLogger.Assert("The Inverter Efficiency Curve is Incorrectly defined. Check Inverter in Sub-Array " + ArrayNum + ".", double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Efficiency/EffCurve/Effic." + (eff + 1).ToString(), _ArrayNum: ArrayNum)) <100, ErrLevel.FATAL);
                    itsOnlyEff[0][eff] = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Efficiency/EffCurve/IN" + (eff + 1).ToString(), _ArrayNum: ArrayNum, _default: itsNomOutputPwr.ToString())) * itsNumInverters * 1000;
                    itsOnlyEff[1][eff] = itsOnlyEff[0][eff]*double.Parse(ReadFarmSettings.GetInnerText("Inverter", "Efficiency/EffCurve/Effic." + (eff + 1).ToString(), _ArrayNum: ArrayNum, _Error: ErrLevel.FATAL)) / 100;
                }
            }
        }

        // Obtaining the AC wiring Resistance to calculate the AC Wiring Losses, this is dependent on the DC array it is connected to.
        public void ConfigACWiring(double PNomArrayDC)
        {
            double itsMaxSubArrayAC;            // Initializing the value of the maximum AC power produced by this Sub-Array

            if (threeCurves)
            {
                itsMaxSubArrayAC = PNomArrayDC * itsMedEff[1][itsMedEff[1].Length - 1]/itsMedEff[0][itsMedEff[0].Length - 1];
            }
            else
            {
                itsMaxSubArrayAC = PNomArrayDC * itsOnlyEff[1][itsOnlyEff[1].Length - 1] / itsOnlyEff[0][itsOnlyEff[0].Length - 1];
            }

            itsACWiringRes = itsACWiringLossPC * itsOutputVoltage / (itsMaxSubArrayAC / (itsOutputVoltage * Math.Sqrt(outputPhases)));
        }
    }
}
