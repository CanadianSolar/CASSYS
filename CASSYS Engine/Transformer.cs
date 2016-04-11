// CASSYS - Grid connected PV system modelling software 
// Version 0.9  
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Transformer Class
// 
// Revision History:
// AP - 2015-01-20: Version 0.9
//
// Description: 
// The Transformer class is used to model a transformer with two types of losses
// Steady state losses such as Iron Loss, and input-dependent 
// Quadratic Resistive Losses.
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
using System.Xml;
using System.Xml.Linq;


namespace CASSYS
{
    class Transformer
    {
        // Parameters of the transformer       
        public double itsPNom;                          // The nominal power of the transformer [W]
        double itsPGlobLoss;                            // The global loss of the transformer [W]
        double itsPIronLoss;                            // The resultant iron loss of the transformer [W]
        bool isNightlyDisconnected;                     // Determines if the transformer is disconnected at night [true for nightly disconnect, false otherwise]

        // Output variables
        public bool isDisconnectedNow = true;           // Determines if the transformer is disconnected at the current time stamp
        public double POut;                             // The transformer power output [W]
        public double EnergyToGrid;                     // Energy supplied to the grid [Wh]
        public double itsPResLss;                       // The resistive loss of the transformer [W] 
        public double Losses;                           // The losses experienced at the transformer [W]

        // Blank constructor for the Transformer
        public Transformer()
        {
        }

        // Calculates the Transformer output
        public void Calculate
            (
            double InputPwr                              // The Inverter power fed into the transformer
            )
        {
            // Check to see if the Total InvAC Pwr is negative, if yes, fix the value to 0
            if (InputPwr < 0)
            {
                InputPwr = 0;
            }

            // Calculating the Resistive Losses that result from the Input Power to the Transformer. 
            if (itsPNom > 0 && itsPResLss > 0)
            {
                itsPResLss = (itsPGlobLoss - itsPIronLoss) * Math.Pow((InputPwr / itsPNom), 2);
            }
            else
            {
                itsPResLss = 0;
            }

            // If the incoming power is more than 0 calculate the output using that and the losses
            if (isNightlyDisconnected)
            {
                // If the Transformer is disconnected now, the Power should be 0, else both constant 
                // and ohmic losses are applied and the output is determined.
                if (isDisconnectedNow)
                {
                    POut = 0;
                }
                else
                {
                    POut = InputPwr - itsPIronLoss - itsPResLss;
                }
            }
            else
            {
                // If the Transformer receives input power, both losses are applied to determine the output
                // If not, only the Iron loss is applied to the transformer. 
                if (InputPwr > 0)
                {
                    POut = InputPwr - itsPIronLoss - itsPResLss;
                }
                else
                {
                    POut = -itsPIronLoss;
                }
            }

            // Calculating Energy to Grid.
            if (POut > 0)
            {
                EnergyToGrid = POut * Util.timeStep / 60;
            }
            else
            {
                EnergyToGrid = 0;
            }

            Losses = itsPIronLoss + itsPResLss;
        }

        //Config will determine and assign values for the losses at the transformer using an .CSYX file
        public void Config()
        {
            // Config will find the Iron Losses, and Global losses from the file. The resistive loss, etc are calculated by the program from these two values.
            itsPIronLoss = Convert.ToDouble(ReadFarmSettings.GetInnerText("Transformer", "PIronLoss", ErrLevel.WARNING)) * 1000;
            itsPGlobLoss = Convert.ToDouble(ReadFarmSettings.GetInnerText("Transformer", "PGlobLossTrf", ErrLevel.WARNING)) * 1000;
            itsPNom = Convert.ToDouble(ReadFarmSettings.GetInnerText("Transformer", "PNomTrf", ErrLevel.WARNING)) * 1000;
            itsPResLss = Convert.ToDouble(ReadFarmSettings.GetInnerText("Transformer", "PResLssTrf", ErrLevel.WARNING)) * 1000;
            
            // Parameters that determine if the transformer remains ON at night, and initializing the disconnection of the transformer. 
            isNightlyDisconnected = Convert.ToBoolean(ReadFarmSettings.GetInnerText("Transformer", "NightlyDisconnect", ErrLevel.WARNING, _default: "False"));
            if (isNightlyDisconnected)
            {
                isDisconnectedNow = true;
            }
            else
            {
                isDisconnectedNow = false;
            }

        }
    }
}
