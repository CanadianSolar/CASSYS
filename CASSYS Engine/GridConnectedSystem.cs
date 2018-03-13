// CASSYS - Grid connected PV system modelling software  
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: GridConnectedSystem.cs
// 
// Revision History:
// NA - 2017-06-09: First release - Modularized the simulation class
//
// Description 
// This class is used to deal with GridConnected related processes within the simulation.
// This class configures/initializes grid-connection related classes, performes grid-connected calculations,
// and assigns grid-connected outputs.
///////////////////////////////////////////////////////////////////////////////
// References and Supporting Documentation or Links
///////////////////////////////////////////////////////////////////////////////
//
//
///////////////////////////////////////////////////////////////////////////////

using System;

namespace CASSYS
{
    class GridConnectedSystem
    {
        // Creating Array of PV Array and Inverter Objects
        PVArray[] SimPVA;                                       // Array of Photovoltaic Arrays within farm
        Inverter[] SimInv;                                      // Array of Inverters within farm
        ACWiring[] SimACWiring;                                 // Array of Wires used in calculating AC wiring loss
        Transformer SimTransformer = new Transformer();         // Transformer instance used in calculations

        SpectralEffects SimSpectral = new SpectralEffects();    // Used to calculate spectral correction relative to a given model

        // Shading Related variables
        HorizonShading SimHorizon = new HorizonShading();       // Used to calculate solar panel shading relative to a given horizon
        Shading SimShading = new Shading();                     // Used to calculate solar panel shading (row to row)
        double ShadGloLoss;                                     // Shading Losses in POA Global
        double ShadGloFactor;                                   // Shading factor on the POA Global
        double ShadBeamLoss;                                    // Shading Losses to Beam
        double ShadDiffLoss;                                    // Shading Losses to Diffuse
        double ShadRefLoss;                                     // Shading Losses to Albedo

        // Output variables Summation or Averages from Sub-Arrays 
        // PV Array related:
        double farmDC = 0;                                      // Farm/PVArray DC Output [W]
        double farmDCModuleQualityLoss = 0;                     // Farm/PVArray DC Module Quality Loss (Sum for all sub-arrays) [W]
        double farmDCMismatchLoss = 0;                          // Farm/PVArray DC Module Mismatch Loss (Sum for all sub-arrays) [W]
        double farmDCOhmicLoss = 0;                             // Farm/PVArray DC Ohmic Loss (Sum for all sub-arrays) [W]
        double farmDCSoilingLoss = 0;                           // Farm/PVArray DC Soiling Loss (Sum for all sub-arrays) [W]
        double farmDCCurrent = 0;                               // Farm/PVArray DC Current Values [A]
        double farmDCTemp = 0;                                  // Average temperature of all PV Arrays [deg C]
        double farmPNomDC = 0;                                  // The nominal Pnom DC for the Farm [kW]
        double farmPNomAC = 0;                                  // The nominal Pnom AC for the Farm [kW]
        double farmTotalModules = 0;                            // The total number of modules in the farm [#]
        double farmArea = 0;                                    // Rough farm area (based on PV Array * Number of Modules in each Sub Array)
        double farmModuleTempAndAmbientTempDiff = 0;            // Difference between the array temperature and the ambient temperature [C]
        double farmDCEfficiency = 0;                            // DC-side efficiency of the farm [%]
        double farmOverAllEff = 0;                              // Overall efficiency of farm
        double farmPR = 0;                                      // Farm Performance Ratio
        double farmSysIER = 0;
        double farmPnom;                                        // Nominal Power of farm [W]
        double farmTempLoss;                                    // Farm/PVArray Loss due to temperature [W]
        double farmRadLoss;                                     // Farm/PVArray Loss due to irradiance level [W]

        // Inverter related calculation variables:
        double farmACOutput = 0;                                // Farm inverter output [W AC]
        double farmACOhmicLoss = 0;                             // Farm/Inverter to Transformer AC Ohmic Loss (Sum for all sub-arrays) [W]
        double farmACPMinThreshLoss = 0;                        // Loss when the power of the array is not sufficient for starting the inverter. [W]
        double farmACClippingPower = 0;                         // Produced power before reduction by Inverter (clipping) [W]
        double farmACMaxVoltageLoss = 0;                        // Loss of power when voltage of the array is too large and forces the inverters to 'shut off' and when inverter is not operating at MPP [W]
        double farmACMinVoltageLoss = 0;                        // Loss of power when voltage of the array is too small and forces the inverters to 'shut off' and when inverter is not operating at MPP [W]

        // Calculate method
        public void Calculate(
                RadiationProc RadProc,                          // Radiation related data
                SimMeteo SimMet                                 // Meteological data from inputfile
            )
        {
            // Reset Losses
            for (int i = 0; i < SimPVA.Length; i++)
            {
                SimInv[i].LossPMinThreshold = 0;
                SimInv[i].LossClipping = 0;
                SimInv[i].LossLowVoltage = 0;
                SimInv[i].LossHighVoltage = 0;
            }

            // Calculating solar panel shading
            SimShading.Calculate(RadProc.SimSun.Zenith, RadProc.SimSun.Azimuth, RadProc.SimHorizonShading.TDir, RadProc.SimHorizonShading.TDif, RadProc.SimHorizonShading.TRef, RadProc.SimTracker.SurfSlope, RadProc.SimTracker.SurfAzimuth);

            // Calculating spectral model effects
            SimSpectral.Calculate(SimMet.HGlo, RadProc.SimSun.NExtra, RadProc.SimSun.Zenith);

            try
            {
                // Calculate PV Array Output for inputs read in this loop
                for (int j = 0; j < ReadFarmSettings.SubArrayCount; j++)
                {
                    // Adjust the IV Curve based on based on Temperature and Irradiance
                    SimPVA[j].CalcIVCurveParameters(SimMet.TGlo, SimShading.ShadTDir, SimShading.ShadTDif, SimShading.ShadTRef, RadProc.SimTilter.IncidenceAngle, SimMet.TAmbient, SimMet.WindSpeed, SimMet.TModMeasured, SimMet.MonthOfYear, SimSpectral.clearnessCorrection);

                    // Check Inverter status to determine if the Inverter is ON or OFF
                    GetInverterStatus(j);

                    // If inverter is off set appropriate variables to 0 and recalculate array in open circuit voltage
                    if (!SimInv[j].isON)
                    {
                        SimInv[j].ACPwrOut = 0;
                        SimInv[j].IOut = 0;
                        SimPVA[j].CalcAtOpenCircuit();
                        SimPVA[j].Calculate(false, SimPVA[j].Voc);
                    }

                    //performing AC wiring calculations
                    SimACWiring[j].Calculate(SimInv[j]);

                    // Assigning the outputs to the dictionary
                    ReadFarmSettings.Outputlist["SubArray_Current" + (j + 1).ToString()] = SimPVA[j].IOut;
                    ReadFarmSettings.Outputlist["SubArray_Voltage" + (j + 1).ToString()] = SimPVA[j].VOut;
                    ReadFarmSettings.Outputlist["SubArray_Power" + (j + 1).ToString()] = SimPVA[j].POut / 1000;
                    ReadFarmSettings.Outputlist["SubArray_Current_Inv" + (j + 1).ToString()] = SimInv[j].IOut;
                    ReadFarmSettings.Outputlist["SubArray_Voltage_Inv" + (j + 1).ToString()] = SimInv[j].itsOutputVoltage;
                    ReadFarmSettings.Outputlist["SubArray_Power_Inv" + (j + 1).ToString()] = SimInv[j].ACPwrOut / 1000;
                }
                
                //Calculating total farm output and total ohmic loss
                farmACOutput = 0;
                farmACOhmicLoss = 0;
                for (int i = 0; i < SimInv.Length; i++)
                {
                    farmACOutput += SimInv[i].ACPwrOut;
                    farmACOhmicLoss += SimACWiring[i].ACWiringLoss;
                }

                SimTransformer.Calculate(farmACOutput - farmACOhmicLoss);

                // Calculating outputs that will be assigned for this interval
                // Shading each component of the Tilted radiaton
                // Using horizon affected tilted radiation
                ShadGloLoss = (RadProc.SimTilter.TGlo - SimShading.ShadTGlo) - RadProc.SimHorizonShading.LossGlo;
                ShadGloFactor = (RadProc.SimTilter.TGlo > 0 ? SimShading.ShadTGlo / RadProc.SimTilter.TGlo : 1);
                ShadBeamLoss = RadProc.SimHorizonShading.TDir - SimShading.ShadTDir;
                ShadDiffLoss = RadProc.SimTilter.TDif > 0 ? RadProc.SimHorizonShading.TDif - SimShading.ShadTDif : 0;
                ShadRefLoss = RadProc.SimTilter.TRef > 0 ? RadProc.SimHorizonShading.TRef - SimShading.ShadTRef : 0;

                //Calculating total farm level variables. Cleaning them so they are non-cumulative.
                farmDC = 0;
                farmDCCurrent = 0;
                farmDCMismatchLoss = 0;
                farmDCModuleQualityLoss = 0;
                farmDCOhmicLoss = 0;
                farmDCSoilingLoss = 0;
                farmDCTemp = 0;
                farmTotalModules = 0;
                farmPNomDC = 0;
                farmPNomAC = 0;
                farmACPMinThreshLoss = 0;
                farmACClippingPower = 0;
                farmACMaxVoltageLoss = 0;
                farmACMinVoltageLoss = 0;
                farmPnom = 0;
                farmTempLoss = 0;
                farmRadLoss = 0;

                for (int i = 0; i < SimPVA.Length; i++)
                {
                    farmDC += SimPVA[i].POut;
                    farmDCCurrent += SimPVA[i].IOut;
                    farmDCMismatchLoss += Math.Max(0, SimPVA[i].MismatchLoss);
                    farmDCModuleQualityLoss += SimPVA[i].ModuleQualityLoss;
                    farmDCOhmicLoss += SimPVA[i].OhmicLosses;
                    farmDCSoilingLoss += SimPVA[i].SoilingLoss;
                    farmDCTemp += SimPVA[i].TModule * SimPVA[i].itsNumModules;
                    farmTotalModules += SimPVA[i].itsNumModules;
                    farmPNomDC += SimPVA[i].itsPNomDCArray;
                    farmPNomAC += SimInv[i].itsPNomArrayAC;
                    farmACPMinThreshLoss += SimInv[i].LossPMinThreshold;
                    farmACClippingPower += SimInv[i].LossClipping;
                    farmACMaxVoltageLoss += SimInv[i].LossHighVoltage;
                    farmACMinVoltageLoss += SimInv[i].LossLowVoltage;
                    farmPnom += (SimPVA[i].itsPNom * SimPVA[i].itsNumModules) * SimPVA[i].TGloEff / 1000;
                    farmTempLoss += SimPVA[i].tempLoss;
                    farmRadLoss += SimPVA[i].radLoss;
                }

                // Averages all PV Array temperature values
                farmDCTemp /= farmTotalModules;
                farmModuleTempAndAmbientTempDiff = farmDCTemp - SimMet.TAmbient;
                farmDCEfficiency = (RadProc.SimTilter.TGlo > 0 ? farmDC / (RadProc.SimTilter.TGlo * farmArea) : 0) * 100;
                farmPNomDC = Utilities.ConvertWtokW(farmPNomDC);
                farmPNomAC = Utilities.ConvertWtokW(farmPNomAC);

                farmOverAllEff = (RadProc.SimTilter.TGlo > 0 && SimTransformer.POut > 0 ? SimTransformer.POut / (RadProc.SimTilter.TGlo * farmArea) : 0) * 100;
                farmPR = RadProc.SimTilter.TGlo > 0 && farmPNomDC > 0 && SimTransformer.POut > 0 ? SimTransformer.POut / RadProc.SimTilter.TGlo / farmPNomDC : 0;
                farmSysIER = (SimTransformer.itsPNom - SimTransformer.POut) / (RadProc.SimTilter.TGlo * 1000);
            }
            catch (Exception ce)
            {
                ErrorLogger.Log(ce, ErrLevel.FATAL);
            }

            // Assigning Outputs for this class.
            AssignOutputs();
        }
        
        public void AssignOutputs()
        {
            ReadFarmSettings.Outputlist["Global_POA_Irradiance_Corrected_for_Shading"] = SimShading.ShadTGlo;
            ReadFarmSettings.Outputlist["Near_Shading_Loss_for_Global"] = ShadGloLoss;
            ReadFarmSettings.Outputlist["Near_Shading_Loss_for_Beam"] = ShadBeamLoss;
            ReadFarmSettings.Outputlist["Near_Shading_Loss_for_Diffuse"] = ShadDiffLoss;
            ReadFarmSettings.Outputlist["Near_Shading_Loss_for_Ground_Reflected"] = ShadRefLoss;
            ReadFarmSettings.Outputlist["Global_POA_Irradiance_Corrected_for_Incidence"] = SimPVA[0].IAMTGlo;
            ReadFarmSettings.Outputlist["Radiation_Soiling_Loss"] = SimPVA[0].RadSoilingLoss;
            ReadFarmSettings.Outputlist["Radiation_Spectral_Loss"] = SimPVA[0].RadSpectralLoss;
            ReadFarmSettings.Outputlist["Incidence_Loss_for_Global"] = SimShading.ShadTGlo - SimPVA[0].IAMTGlo;
            ReadFarmSettings.Outputlist["Incidence_Loss_for_Beam"] = SimShading.ShadTDir * (1 - SimPVA[0].IAMDir);
            ReadFarmSettings.Outputlist["Incidence_Loss_for_Diffuse"] = SimShading.ShadTDif * (1 - SimPVA[0].IAMDif);
            ReadFarmSettings.Outputlist["Incidence_Loss_for_Ground_Reflected"] = SimShading.ShadTRef * (1 - SimPVA[0].IAMRef);
            ReadFarmSettings.Outputlist["Profile_Angle"] = Util.RTOD * SimShading.ProfileAng;
            ReadFarmSettings.Outputlist["Near_Shading_Factor_on_Global"] = ShadGloFactor;
            ReadFarmSettings.Outputlist["Near_Shading_Factor_on_Beam"] = SimShading.BeamSF;
            ReadFarmSettings.Outputlist["Near_Shading_Factor_on__Diffuse"] = SimShading.DiffuseSF;
            ReadFarmSettings.Outputlist["Near_Shading_Factor_on_Ground_Reflected"] = SimShading.ReflectedSF;
            ReadFarmSettings.Outputlist["IAM_Factor_on_Global"] = (SimShading.ShadTGlo > 0 ? SimPVA[0].IAMTGlo / SimShading.ShadTGlo : 1);
            ReadFarmSettings.Outputlist["IAM_Factor_on_Beam"] = SimPVA[0].IAMDir;
            ReadFarmSettings.Outputlist["IAM_Factor_on__Diffuse"] = SimPVA[0].IAMDif;
            ReadFarmSettings.Outputlist["IAM_Factor_on_Ground_Reflected"] = SimPVA[0].IAMRef;
            ReadFarmSettings.Outputlist["Effective_Irradiance_in_POA"] = SimPVA[0].TGloEff;
            ReadFarmSettings.Outputlist["Array_Nominal_Power"] = farmPnom / 1000;
            ReadFarmSettings.Outputlist["Array_Soiling_Loss"] = farmDCSoilingLoss / 1000;
            ReadFarmSettings.Outputlist["Modules_Array_Mismatch_Loss"] = farmDCMismatchLoss / 1000;
            ReadFarmSettings.Outputlist["Ohmic_Wiring_Loss"] = farmDCOhmicLoss / 1000;
            ReadFarmSettings.Outputlist["Module_Quality_Loss"] = farmDCModuleQualityLoss / 1000;
            ReadFarmSettings.Outputlist["Effective_Energy_at_the_Output_of_the_Array"] = farmDC / 1000;
            ReadFarmSettings.Outputlist["Calculated_Module_Temperature__deg_C_"] = farmDCTemp;
            ReadFarmSettings.Outputlist["Difference_between_Module_and_Ambient_Temp.__deg_C_"] = farmModuleTempAndAmbientTempDiff;
            ReadFarmSettings.Outputlist["PV_Array_Current"] = farmDCCurrent;
            ReadFarmSettings.Outputlist["PV_Array_Voltage"] = (farmDCCurrent > 0 ? farmDC / farmDCCurrent : 0);
            ReadFarmSettings.Outputlist["Available_Energy_at_Inverter_Output"] = farmACOutput / 1000;
            ReadFarmSettings.Outputlist["AC_Ohmic_Loss"] = farmACOhmicLoss / 1000;
            ReadFarmSettings.Outputlist["Inverter_Efficiency"] = (farmACOutput > 0 ? farmACOutput / farmDC : 0) * 100;      
            ReadFarmSettings.Outputlist["DCAC_Conversion_Losses"] = (farmACOutput > 0 ? farmDC - farmACOutput:0)/1000;
            ReadFarmSettings.Outputlist["Inverter_Loss_Due_to_Low_Power_Threshold"] = farmACPMinThreshLoss / 1000;
            ReadFarmSettings.Outputlist["Inverter_Loss_Due_to_Low_Voltage_Threshold"] = farmACMinVoltageLoss / 1000;
            ReadFarmSettings.Outputlist["Inverter_Loss_Due_to_High_Power_Threshold"] = farmACClippingPower > 0 ? farmACClippingPower / 1000 : 0;
            ReadFarmSettings.Outputlist["Inverter_Loss_Due_to_High_Voltage_Threshold"] = farmACMaxVoltageLoss / 1000;
            ReadFarmSettings.Outputlist["Virtual_Inverter_Input_Energy"] = (farmDC + farmACPMinThreshLoss + farmACMinVoltageLoss + (farmACClippingPower > 0 ? farmACClippingPower : 0) + farmACMaxVoltageLoss) /1000;
            ReadFarmSettings.Outputlist["External_transformer_loss"] = (SimTransformer.POut < 0) ? 0 : (SimTransformer.Losses / 1000);
            ReadFarmSettings.Outputlist["Power_Injected_into_Grid"] = SimTransformer.POut / 1000;
            ReadFarmSettings.Outputlist["Energy_Injected_into_Grid"] = SimTransformer.EnergyToGrid / 1000;
            ReadFarmSettings.Outputlist["PV_Array_Efficiency"] = farmDCEfficiency;
            ReadFarmSettings.Outputlist["AC_side_Efficiency"] = (farmACOutput > 0 && SimTransformer.POut > 0 ? SimTransformer.POut / farmACOutput : 0) * 100;
            ReadFarmSettings.Outputlist["Overall_System_Efficiency"] = farmOverAllEff;
            ReadFarmSettings.Outputlist["Normalized_System_Production"] = SimTransformer.POut > 0 ? SimTransformer.POut / (farmPNomDC * 1000) : 0;
            ReadFarmSettings.Outputlist["Array_losses_ratio"] = SimTransformer.POut > 0 ? (farmDCMismatchLoss + farmDCModuleQualityLoss + farmDCOhmicLoss + farmDCSoilingLoss) / SimTransformer.POut : 0;
            ReadFarmSettings.Outputlist["Inverter_losses_ratio"] = SimTransformer.POut > 0 ? farmACOhmicLoss / SimTransformer.POut : 0;
            ReadFarmSettings.Outputlist["AC_losses_ratio"] = SimTransformer.Losses / SimTransformer.POut < 0 ? 0 : SimTransformer.Losses / SimTransformer.POut;
            ReadFarmSettings.Outputlist["Performance_Ratio"] = farmPR;
            ReadFarmSettings.Outputlist["System_Loss_Incident_Energy_Ratio"] = farmSysIER;
            ReadFarmSettings.Outputlist["Power_Loss_Due_to_Temperature"] = farmTempLoss / 1000;
            ReadFarmSettings.Outputlist["Energy_Loss_Due_to_Irradiance"] = farmRadLoss / 1000;
            ReadFarmSettings.Outputlist["NightTime_Energizing_Loss"] = (SimTransformer.POut < 0) ? SimTransformer.itsPIronLoss / 1000 : 0;

            // Get the power for individual Sub-Arrays
            ReadFarmSettings.Outputlist["Sub_Array_Performance"] = "";
            ReadFarmSettings.Outputlist["ShowSubInv"] = "";
            ReadFarmSettings.Outputlist["ShowSubInvV"] = "";
            ReadFarmSettings.Outputlist["ShowSubInvC"] = "";

            for (int subNum = 1; subNum < SimPVA.Length + 1; subNum++)
            {
                ReadFarmSettings.Outputlist["Sub_Array_Performance"] += ReadFarmSettings.Outputlist["SubArray_Voltage" + subNum].ToString() + "," + ReadFarmSettings.Outputlist["SubArray_Current" + subNum].ToString() + "," + ReadFarmSettings.Outputlist["SubArray_Power" + subNum].ToString() + (subNum != SimPVA.Length ? "," : "");
                ReadFarmSettings.Outputlist["ShowSubInv"] += ReadFarmSettings.Outputlist["SubArray_Power_Inv" + subNum] + (subNum != SimPVA.Length ? "," : "");
                ReadFarmSettings.Outputlist["ShowSubInvV"] += ReadFarmSettings.Outputlist["SubArray_Voltage_Inv" + subNum].ToString() + (subNum != SimPVA.Length ? ",": "");
                ReadFarmSettings.Outputlist["ShowSubInvC"] += ReadFarmSettings.Outputlist["SubArray_Current_Inv" + subNum].ToString() + (subNum != SimPVA.Length ? "," : "");
            }
        }

        // Calculates the Voltage at which the Inverter will produce Nom AC Power (when Clipping) using Bisection Method.
        void GetClippingVoltage(int j)
        {
            SimPVA[j].CalcAtOpenCircuit();                                  // Calculating Open Circuit characteristics to determine upper and lower bound of interpolation

            double InvVR = SimPVA[j].Voc;                                   // The higher bound of the Voltage Range [V]
            double InvVL = SimPVA[j].VOut;                                  // The lower bound of the Voltage Range  [V]
            double trialInvV = (InvVR + InvVL) / 2;                         // Search variable                       [V] 
            double tolerance = 0.0001;                                      // The tolerance value value at which the bounds are close enough [V]

            // Beginning Bisection Method to find the voltage at which the Inverter will produce Nominal AC Power Out
            do
            {
                SimPVA[j].Calculate(false, trialInvV);                      // Calculate the PV Array Power at given voltage
                SimInv[j].Calculate(SimPVA[j].POut, trialInvV);             // Calculate the Inverter AC Out and Determine the Clipping Status

                if (SimInv[j].isClipping)
                {
                    InvVL = trialInvV;                                      // If Clipping, lower bound moves to Search Variable   
                }
                else
                {
                    InvVR = trialInvV;                                      // If not clipping, higher bound moves to Search Variable
                }

                trialInvV = (InvVR + InvVL) / 2;                            // Calculate new search variable [V]
            }
            while (Math.Abs(InvVR - InvVL) > tolerance);
            
            SimInv[j].VInDC = trialInvV;
        }


        // Checks the status of the Inverter (ON, MPPT tracking, etc) and configures its operation based on the PV Array's characteristics. 
        void GetInverterStatus(int j)
        {
           // Determine MPPT from the array
            SimPVA[j].Calculate(true, 0);
            double arrayVMPP = SimPVA[j].VOut;
            double arrayPMPP = SimPVA[j].PMPP;
            
            // If the Inverter is off, check if the Open Circuit Voltage of the Array is sufficient to turn the Inverter ON
            if (!SimInv[j].isON)
            {
                // Determining Array Voltage 
                SimPVA[j].CalcAtOpenCircuit();
                double vOpenC = SimPVA[j].Voc;

                //(if BiPolar open circuit voltage is divided by 2)
                if (SimInv[j].isBipolar)
                {
                    vOpenC = vOpenC / 2;
                }

                if (vOpenC < SimInv[j].itsMinVoltage)
                {
                    SimInv[j].hasMinVoltage = false;
                    SimInv[j].isON = false;
                    SimInv[j].ACPwrOut = 0;
                    SimInv[j].IOut = 0;
                    // Calculate energy loss due to voltage too low
                    SimInv[j].LossLowVoltage = arrayPMPP;
                    // Inverter is off, nothing else to do
                    return;
                }
                else
                {
                    SimInv[j].hasMinVoltage = true;
                    SimInv[j].isON = true;
                }
            }
            
            // If the Inverter turns on because of sufficient voltage, or if the Inverter was already ON
            // MPPT check, if true then use voltage window to determine voltage out of Inverter and if it is in the MPPT Window
            if (SimInv[j].isBipolar)
            {
                arrayVMPP /= 2;
            }
            
            // SimInv[j].GetMPPTStatus(arrayVMPP, out SimInv[j].inMPPTWindow);
            SimInv[j].GetMPPTStatus(arrayVMPP);

            // Recalculate array operating point now that the voltage has been determined
            SimPVA[j].Calculate(false, SimInv[j].VInDC);
            SimInv[j].Calculate(SimPVA[j].POut, SimInv[j].VInDC); 

            if (SimInv[j].VInDC == SimInv[j].itsMppWindowMin)
            {
                // Calculate energy loss due to voltage too low
                SimInv[j].LossLowVoltage = arrayPMPP;
            }

            if (SimInv[j].VInDC == SimInv[j].itsMppWindowMax)
            {
                // Calculate energy loss due to voltage too high
                SimInv[j].LossHighVoltage = arrayPMPP - SimPVA[j].POutNoLoss; 
            }

            // Check if the inverter has sufficient power to stay ON 
            if (SimPVA[j].POut < (SimInv[j].itsThresholdPwr * SimInv[j].itsNumInverters) || SimInv[j].ACPwrOut < 0)
            { 
                // Inverter must be turned off.
                SimInv[j].isON = false;
                SimInv[j].ACPwrOut = 0;
                SimInv[j].IOut = 0;
                SimInv[j].LossLowVoltage = 0;
                SimInv[j].LossHighVoltage = 0;
                SimInv[j].LossPMinThreshold = SimPVA[j].PMPP;
                return; 
            }
            
            // If the Inverter is Clipping, the voltage is increased till the Inverter will not Clip anymore. 
            if (SimInv[j].isClipping)
            {
                double NoClippingPwr = SimPVA[j].POutNoLoss;                      // Output power of the array
                GetClippingVoltage(j);
                
                SimPVA[j].Calculate(false, SimInv[j].VInDC);
                SimInv[j].Calculate(SimPVA[j].POut, SimInv[j].VInDC);
 
                SimInv[j].LossClipping = NoClippingPwr - SimPVA[j].POutNoLoss;
            }
            // Check that max voltage is not exceeded
            if (SimInv[j].VInDC > SimInv[j].itsMaxVoltage) 
            {
                SimInv[j].isON = false;
                SimInv[j].ACPwrOut = 0;
                SimInv[j].IOut = 0;
                SimInv[j].LossLowVoltage = 0;
                SimInv[j].LossClipping = 0;
                // Calculate energy lost
                SimInv[j].LossHighVoltage = SimPVA[j].PMPP;
            }
        }

        // Configuring the required elements for a grid connected system
        public void Config()
        {
            SimShading.Config();
            SimTransformer.Config();
            SimSpectral.Config();

            // Array of PVArray, Inverter and Wiring objects based on the number of Sub-Arrays 
            SimPVA = new PVArray[ReadFarmSettings.SubArrayCount];
            SimInv = new Inverter[ReadFarmSettings.SubArrayCount];
            SimACWiring = new ACWiring[ReadFarmSettings.SubArrayCount];

            // Initialize and Configure PVArray and Inverter Objects through their .CSYX file
            for (int SubArrayCount = 0; SubArrayCount < ReadFarmSettings.SubArrayCount; SubArrayCount++)
            {
                SimInv[SubArrayCount] = new Inverter();
                SimInv[SubArrayCount].Config(SubArrayCount + 1);
                SimPVA[SubArrayCount] = new PVArray();
                SimPVA[SubArrayCount].Config(SubArrayCount + 1);
                SimACWiring[SubArrayCount] = new ACWiring();

                // If 'at Pnom' specified for AC Loss Fraction in version 1.2.0
                if (string.Compare(ReadFarmSettings.CASSYSCSYXVersion, "1.2.0") >= 0 && ReadFarmSettings.GetInnerText("System", "ACWiringLossAtSTC", _Error: ErrLevel.WARNING) == "False")
                {
                    SimACWiring[SubArrayCount].Config(SubArrayCount + 1, SimInv[SubArrayCount].itsOutputVoltage, SimInv[SubArrayCount].outputPhases, SimInv[SubArrayCount].itsPNomArrayAC);
                }
                else
                {
                    SimACWiring[SubArrayCount].Config(SubArrayCount + 1, SimInv[SubArrayCount].itsOutputVoltage, SimInv[SubArrayCount].outputPhases, SimInv[SubArrayCount].itsMaxSubArrayACEff * SimPVA[SubArrayCount].itsPNomDCArray);
                }
                farmArea += SimPVA[SubArrayCount].itsRoughArea;
            }
        }
    }
}
