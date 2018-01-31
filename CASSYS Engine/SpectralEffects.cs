// CASSYS - Grid connected PV system modelling software
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: SpectralEffects.cs
//
// Revision History:
//
// Description:
// This class is responsible for the simulation of 'spectral effects', albeit
// with a relatively simple model. PV array production is adjusted based on a
// clearness index curve given by the user.
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
    class SpectralEffects
    {
        // Input variables
        string SpectralClearnessIndexStr;                           // String containing the Clearness Index points from the .csyx file [unitless]
        string SpectralClearnessCorrectionStr;                      // String containing the Clearness Correction points from the .csyx file [unitless]

        // Spectral Effects local variables/arrays and intermediate calculation variables and arrays
        double[] ClearnessIndexArr;                                 // Array containing the comma-separated index values from SpectralClearnessIndexStr [unitless]
        double[] ClearnessCorrectionArr;                            // Array containing the comma-separated correction values from SpectralClearnessCorrectionStr [unitless]
        double clearnessIndex;                                      // Clearness index

        // Settings used in the class
        bool spectralModelUsed;                                     // Used to determine whether or not the user is using a spectral effects model

        // Output variables
        public double clearnessCorrection;                          // Clearness correction

        // Spectral Effects constructor
        public SpectralEffects()
        {
        }

        // Config takes values from the xml, which only needs to be done once
        public void Config
            (
            )
        {
            // Loads the Spectral Model from the .csyx document
            spectralModelUsed = Convert.ToBoolean(ReadFarmSettings.GetInnerText("Spectral", "UseSpectralModel", ErrLevel.WARNING, _default: "false"));
            if (spectralModelUsed == true)
            {
                // Getting the Spectral Model information from the .csyx file
                SpectralClearnessIndexStr = ReadFarmSettings.GetInnerText("Spectral", "ClearnessIndex/kt", ErrLevel.WARNING, "0.9", _default: "1");
                SpectralClearnessCorrectionStr = ReadFarmSettings.GetInnerText("Spectral", "ClearnessIndex/ktCorrection", ErrLevel.WARNING, "0.9", _default: "0");

                // Converts the Spectral Model imported from the .csyx file into an array of doubles
                ClearnessIndexArr = SpectralCSVStringtoArray(SpectralClearnessIndexStr);
                ClearnessCorrectionArr = SpectralCSVStringtoArray(SpectralClearnessCorrectionStr);

                // If user inputs spectral index/correction data of two different lengths
                if (ClearnessIndexArr.Length != ClearnessCorrectionArr.Length)
                {
                    ErrorLogger.Log("The number of clearness correction values was not equal to the number of clearness index values.", ErrLevel.FATAL);
                }
            }
            else
            {
                ClearnessIndexArr = new double[] { 1 };
                ClearnessCorrectionArr = new double[] { 0 };
            }
            // TODO: remove later
            for (int i = 0; i < ClearnessIndexArr.Length; i++)
            {
                Console.Write(ClearnessIndexArr[i]);
                Console.Write(" ");
            }
            Console.WriteLine();
            for (int i = 0; i < ClearnessCorrectionArr.Length; i++)
            {
                Console.Write(ClearnessCorrectionArr[i]);
                Console.Write(" ");
            }
            Console.WriteLine();
        }

        // Calculate manages calculations that need to be run for each time step
        public void Calculate
            (
              double HGlo                                           // Horizontal Global Irradiance [W/m2]
            , double NExtra                                         // Extraterrestrial Normal Irradiance [W/m2]
            , double SunZenith                                      // The Zenith position of the sun with 0 being normal to the earth [radians]
            )
        {
            clearnessIndex = Sun.GetClearnessIndex(HGlo, NExtra, SunZenith);
            clearnessCorrection = Interpolate.Linear(ClearnessIndexArr, ClearnessCorrectionArr, clearnessIndex);
        }

        // Converts string array into an array of doubles that can be used by the program
        public static double[] SpectralCSVStringtoArray(string SpectralString)
        {
            string[] tempArray = SpectralString.Split(',');
            int arrayLength = tempArray.Length;
            double[] SpectralArray = new double[arrayLength];

            for (int arrayIndex = 0; arrayIndex < arrayLength; arrayIndex++)
            {
                SpectralArray[arrayIndex] = double.Parse(tempArray[arrayIndex]);
            }

            return SpectralArray;
        }
    }
}
