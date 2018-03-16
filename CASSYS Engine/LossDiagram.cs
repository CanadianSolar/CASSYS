// CASSYS - Grid connected PV system modelling software  
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: LossDiagram
// 
// Revision History:
// TS - 2017-11-01: Version 1.3.0 First release
// 
//
// Description:
// This class calculates the total losses of the system to write them to a temp 
// file for the interface to read and update the losses diagram
//                        
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
// https://www.dotnetperls.com/xmlwriter: XmlWriter
// 
///////////////////////////////////////////////////////////////////////////////
using System;
using System.Xml;
using System.Collections.Generic;
using System.Windows.Forms;

namespace CASSYS
{
    public class LossDiagram
    {
        // Parameters for config and variables used for calculations
        public static Dictionary<String, dynamic> LossOutputs;         // Creating a dictionary to hold all output values for the loss diagram;

        // Irradiance values [W/m^2]
        double horizontalGlobRad = 0;       // Sum of horizontal global irradiance 
        double globRadInPOA = 0;            // Sum of global radiation in POA 
        double horizonShadeLosses = 0;      // Sum of radiation loss due to horizon shading factors
        double nearShadeLosses = 0;         // Sum of radiation loss due to row to row shading fatctors 
        double soilingLoss = 0;             // Sum of radiation loss due to panel soiling
        double incidenceAngleLoss = 0;      // Sum of radiation loss due to incidence angle
        double bifacialGain = 0;            // Sum of radiation gain due to bifacial panels
        double spectralLoss = 0;            // Sum of radiation loss due to spectral effects
        double effectivePOARad = 0;         // Sum of effective radiation available for power conversion

        // PV array level values [kWh]
        double arrayNomEnergy = 0;          // Nominal power of the farm
        double temperatureLoss = 0;         // Energy loss due to module temperature efficiency 
        double irradianceLevelLoss = 0;     // Energy loss due to irradiance level 
        double qualityLoss = 0;             // Energy loss due to module quality
        double mismatchLoss = 0;            // Energy loss due to module mismatch
        double wiringLossDC = 0;            // Energy loss due to DC wiring

        // Inverter level variables [kWh]
        double virtInverterInEnergy = 0;    // Virtual energy available to the inverter at the inverter input 
        double powLowLoss = 0;              // Energy loss due to below threshold power at inverter input
        double powHighLoss = 0;             // Energy loss due to above nominal power at inverter input (Clipping) 
        double voltageLowLoss = 0;          // Energy loss due to below threshold voltage at inverter input 
        double voltageHighloss = 0;         // Energy loss due to above maximum voltage at inverter input
        double actualInverterInEnergy = 0;  // Actual energy available to the inverter at the inverter input
        double DCAC_ConversionLoss = 0;     // Energy loss due to power conversion from DC to AC
        double inverterOutput = 0;          // AC Energy available at the inverter output 

        // Transformer level variables [kWh]
        double wiringLossAC = 0;            // Energy loss due to AC wiring
        double nightEnergizeLoss = 0;       // Energy loss due to transformer energization at night
        double transformerLoss = 0;         // Energy loss due to transformer energization during day //-----------------------------
        double gridEnergy = 0;               // Energy injected into the grid

        // Blank constructor
        public LossDiagram()
        {
        }

        // Calculate will configure and calculate configure necessary values for loss diagram and assign them to dictionary for output
        public void Calculate()
        {
            // Sum loss values for every individual simulation
            // Radiation level values
            horizontalGlobRad += ReadFarmSettings.Outputlist["Horizontal_Global_Irradiance"];
            globRadInPOA += ReadFarmSettings.Outputlist["Global_Irradiance_in_Array_Plane"];
            horizonShadeLosses += ReadFarmSettings.Outputlist["FarShading_Global_Loss"];
            nearShadeLosses += ReadFarmSettings.Outputlist["Near_Shading_Loss_for_Global"];
            soilingLoss += ReadFarmSettings.Outputlist["Radiation_Soiling_Loss"];
            incidenceAngleLoss += ReadFarmSettings.Outputlist["Incidence_Loss_for_Global"];
            bifacialGain += ReadFarmSettings.Outputlist["Bifacial_Gain"];
            spectralLoss += ReadFarmSettings.Outputlist["Radiation_Spectral_Loss"];
            effectivePOARad += ReadFarmSettings.Outputlist["Effective_Irradiance_in_POA"];

            // PV array level values [kWh]
            arrayNomEnergy += ReadFarmSettings.Outputlist["Array_Nominal_Power"];
            temperatureLoss += ReadFarmSettings.Outputlist["Power_Loss_Due_to_Temperature"]; 
            irradianceLevelLoss += ReadFarmSettings.Outputlist["Energy_Loss_Due_to_Irradiance"]; 
            qualityLoss += ReadFarmSettings.Outputlist["Module_Quality_Loss"]; 
            mismatchLoss += ReadFarmSettings.Outputlist["Modules_Array_Mismatch_Loss"]; 
            wiringLossDC += ReadFarmSettings.Outputlist["Ohmic_Wiring_Loss"]; 

            //Inverter level values [kWh]
            virtInverterInEnergy += ReadFarmSettings.Outputlist["Virtual_Inverter_Input_Energy"]; 
            powLowLoss += ReadFarmSettings.Outputlist["Inverter_Loss_Due_to_Low_Power_Threshold"]; 
            powHighLoss += ReadFarmSettings.Outputlist["Inverter_Loss_Due_to_High_Power_Threshold"]; 
            voltageLowLoss += ReadFarmSettings.Outputlist["Inverter_Loss_Due_to_Low_Voltage_Threshold"]; 
            voltageHighloss += ReadFarmSettings.Outputlist["Inverter_Loss_Due_to_High_Voltage_Threshold"]; 
            actualInverterInEnergy += ReadFarmSettings.Outputlist["Effective_Energy_at_the_Output_of_the_Array"]; 
            DCAC_ConversionLoss += ReadFarmSettings.Outputlist["DCAC_Conversion_Losses"]; 
            inverterOutput += ReadFarmSettings.Outputlist["Available_Energy_at_Inverter_Output"]; 
            
            // Transformer level values [kWh]
            wiringLossAC += ReadFarmSettings.Outputlist["AC_Ohmic_Loss"]; 
            nightEnergizeLoss += ReadFarmSettings.Outputlist["NightTime_Energizing_Loss"]; 
            transformerLoss += ReadFarmSettings.Outputlist["External_transformer_loss"];
            gridEnergy += ReadFarmSettings.Outputlist["Power_Injected_into_Grid"];
        }
        
        // Write dictionary values to temp output file
        public void AssignLossOutputs()
        {
            LossOutputs = new Dictionary<String, dynamic>();
            // Setting up the xml writer for xml output
            XmlWriterSettings writeSettings = new XmlWriterSettings();
            writeSettings.Indent = true;
            writeSettings.NewLineOnAttributes = true;
            XmlWriter OutputFileWriter = XmlWriter.Create(Application.StartupPath + "/LossDiagramOutputs.xml", writeSettings);  

            // Losses due to radiation
            LossOutputs["Horizontal_Global_Radiation"] = horizontalGlobRad * Util.timeStep / 60;
            LossOutputs["Global_Radiation_in_POA"] = globRadInPOA * Util.timeStep / 60;
            LossOutputs["Horizon_Shading_Losses"] = horizonShadeLosses * Util.timeStep / 60;
            LossOutputs["Near_Shading_Losses"] = nearShadeLosses * Util.timeStep / 60;
            LossOutputs["Soiling_Losses"] = soilingLoss * Util.timeStep / 60;
            LossOutputs["Incidence_Angle_Losses"] = incidenceAngleLoss * Util.timeStep / 60;
            LossOutputs["Bifacial_Gain"] = -1 * bifacialGain * Util.timeStep / 60;                  // Multiply by -1 to convert gains to losses
            LossOutputs["Spectral_Losses"] = spectralLoss * Util.timeStep / 60;
            LossOutputs["Effective_POA_Radiation"] = effectivePOARad * Util.timeStep / 60;
            // PV CONVERSION takes place here
            // Losses at PV modules
            LossOutputs["PV_Array_Nominal_Energy"] = arrayNomEnergy * Util.timeStep / 60;
            LossOutputs["Energy_Loss_Due_to_Temperature"] = temperatureLoss * Util.timeStep / 60;
            LossOutputs["Energy_Loss_Due_to_Irradiance_Level"] = irradianceLevelLoss * Util.timeStep / 60;
            LossOutputs["Module_Quality_Losses"] = qualityLoss * Util.timeStep / 60;
            LossOutputs["Mismatch_Losses"] = mismatchLoss * Util.timeStep / 60;
            LossOutputs["DC_Wiring_Losses"] = wiringLossDC * Util.timeStep / 60;
            // Losses at inverter
            LossOutputs["Virtual_Inverter_Input"] = virtInverterInEnergy * Util.timeStep / 60;
            LossOutputs["Energy_Lost_to_Input_Voltage_too_Low"] = voltageLowLoss * Util.timeStep / 60;
            LossOutputs["Energy_Lost_to_Input_Voltage_too_High"] = voltageHighloss * Util.timeStep / 60;
            LossOutputs["Energy_Lost_to_Input_Power_too_Low"] = powLowLoss * Util.timeStep / 60;
            LossOutputs["Energy_Lost_to_Input_Power_too_High"] = powHighLoss * Util.timeStep / 60;
            LossOutputs["Actual_Inverter_Input_Energy"] = actualInverterInEnergy * Util.timeStep / 60;
            LossOutputs["DCAC_Conversion_Losses"] = DCAC_ConversionLoss * Util.timeStep / 60;
            LossOutputs["Inverter_Output"] = inverterOutput * Util.timeStep / 60;
            // Losses at transformer
            LossOutputs["AC_Wiring_Losses"] = wiringLossAC * Util.timeStep / 60;
            LossOutputs["Night-Time_Energization_Losses"] = nightEnergizeLoss * Util.timeStep / 60;
            LossOutputs["External_Transformer_Losses"] = transformerLoss * Util.timeStep / 60;
            LossOutputs["Power_Injected_into_Grid"] = gridEnergy * Util.timeStep / 60;

            // Write to temp file
            OutputFileWriter.WriteStartDocument();
            OutputFileWriter.WriteStartElement("Losses");

            foreach (KeyValuePair<string, dynamic> entry in LossOutputs)
            {
                // write to xml document
                OutputFileWriter.WriteElementString(entry.Key, entry.Value.ToString());
            }

            OutputFileWriter.WriteEndElement();
            OutputFileWriter.WriteEndDocument();
            OutputFileWriter.Flush();
        }
    }
}
