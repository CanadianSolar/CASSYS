// CASSYS - Grid connected PV system modelling software 
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: AC Wiring
// 
// Revision History:
// NA - 2017-06-09: First release
//
// Description:
// The AC Wiring class calculates the ohmic power loss through resistance in the system wiring
// on the AC side of the inverter.
//                             
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
// None
///////////////////////////////////////////////////////////////////////////////

using System;

namespace CASSYS
{
    class ACWiring
    {
        // Inverter wiring losses related variable
        double itsACWiringLossPC;                           // The AC wiring loss specified as a percentage [%]
        double itsACWiringRes;                              // The AC wiring loss translated from a percentage to a Resistance [ohms]

        // Output variables calculated
        public double ACWiringLoss;                         // AC Wiring Loss incurred [W]

        // Blank constructor
        public ACWiring()
        {
        }

        public void Calculate
            (
            Inverter SimInverter
            )
        {
            // AC wiring losses using the current caculated (three phase output assumed)
            ACWiringLoss = Math.Sqrt(SimInverter.outputPhases) * Math.Pow(SimInverter.IOut, 2) * itsACWiringRes;
        }


        // Config will assign parameter variables their values as obtained from the .CSYX file
        public void Config
            (
              int ArrayNum
            , double OperatingVoltage
            , double Phases
            , double MaximumInputPower
            )
        {
            itsACWiringLossPC = double.Parse(ReadFarmSettings.GetInnerText("Inverter", "LossFraction", _ArrayNum: ArrayNum));
            itsACWiringRes = itsACWiringLossPC * OperatingVoltage / (MaximumInputPower / (OperatingVoltage * Math.Sqrt(Phases)));
        }
    }
}
