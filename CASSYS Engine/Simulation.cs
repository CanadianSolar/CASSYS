// CASSYS - Grid connected PV system modelling software 
// Version 0.9  
// (c) Canadian Solar Solutions Inc.
///////////////////////////////////////////////////////////////////////////////
//
// Title: Simulation.cs
// 
// Revision History:
// AP - 2014-10-14: Version 0.9
// AP - 2015-04-20: Veresion 0.9.1 Added user defined IAM profile for PV Module
// AP - 2015-05-22: Version 0.9.2 Added ability to perform irradiance only calculation
// AP - 2015-06-12: Version 0.9.2 Added separate detranspose method if meter and panel tilt do not match
// 
// Description 
// The main program reads the input file and output file and interacts with 
// all other classes that pertain to a grid connected solar farm 
// (PV Array, Inverter, Transformers, etc.)
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
using System.Xml;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using CASSYS;
using System.Xml.XPath;
using Microsoft.VisualBasic.FileIO;

namespace CASSYS
{
    class Simulation
    {
        #region List and definition of variables used in the program file
        // Instance of all one-time objects, i.e. weather and radiation related objects
        Sun SimSun = new Sun();
        Splitter SimSplitter = new Splitter();
        Tracker SimTracker = new Tracker();
        Tilter SimTilter = new Tilter();
        Tilter pyranoTilter = new Tilter(TiltAlgorithm.HAY);
        Shading SimShading = new Shading();
        Transformer SimTransformer = new Transformer();
        SimMeteo SimMet = new SimMeteo();

        // Creating Array of PV Array and Inverter Objects
        PVArray[] SimPVA;
        Inverter[] SimInv;

        // Calculations for Writing to Output
        // Local variable declarations
        double ShadBeamLoss;                            // Shading Losses to Beam
        double ShadDiffLoss;                            // Shading Losses to Diffuse
        double ShadRefLoss;                             // Shading Losses to Albedo
        bool negativeIrradFlag = false;                 // Negative Irradiance Warning Flag 


        // Output variables Summation or Averages from Sub-Arrays 
        // PV Array related:
        double farmDC;                                      // Farm/PVArray DC Output [W]
        double farmDCModuleQualityLoss;                     // Farm/PVArray DC Module Quality Loss (Sum for all sub-arrays) [W]
        double farmDCMismatchLoss;                          // Farm/PVArray DC Module Mismatch Loss (Sum for all sub-arrays) [W]
        double farmDCOhmicLoss;                             // Farm/PVArray DC Ohmic Loss (Sum for all sub-arrays) [W]
        double farmDCSoilingLoss;                           // Farm/PVArray DC Soiling Loss (Sum for all sub-arrays) [W]
        double farmDCCurrent;                               // Farm/PVArray DC Current Values [A]
        double farmDCTemp;                                  // Average temperature of all PV Arrays [deg C]
        double farmPNomDC;                                  // The nominal Pnom DC for the Farm [kW]
        double farmPNomAC;                                  // The nominal Pnom AC for the Farm [kW]
        double farmTotalModules;                            // The total number of modules in the farm [#]

        // Inverter related calculation variables:
        double farmACOutput;                                // Farm output [W AC]
        double farmACOhmicLoss;                             // Farm/Inverter to Transformer AC Ohmic Loss (Sum for all sub-arrays) [W]
        double farmACPMinThreshLoss;                        // Loss when the power of the array is not sufficient for starting the inverter. [W]
        double farmACClippingPower;                         // Produced power before reduction by Inverter (clipping) [W]

        // Transformer related variables:
        double pGrid;                                       // Power exported to the grid [W]

        // Intermediate calculation variables
        int DayOfYear;                                      // Day of the Year [#]
        int MonthOfYear;                                    // Month of the year [#]
        double HourOfDay;                                   // Hour of Day [#]   
        double TimeStepEnd;                                 // Next Hour of Day [#]
        double TimeStepBeg;                                 // The Hour of the timeStamp from the Input file [#]
        DateTime TimeStampAnalyzed;                         // The time-stamp that used for Sun position calculations [yyyy-mm-dd hh:mm:ss]

        #endregion

        // Blank constructor for the CASSYS Program
        public Simulation()
        {
        }

        // Method to start the simulation software run.
        public void Simulate(String XMLFileName, String _Input = null, String _Output = null)
        {
            #region SITE CONFIGURATION BEINGS HERE

            // Adding timer to calculate elapsed time for the simulation.
            Stopwatch timeTaken = new Stopwatch();
            timeTaken.Start();

            // Get .CSYX file name from CMDPrompt, and Load document
            XmlDocument SiteSettings = new XmlDocument();
            try
            {
                SiteSettings.Load(XMLFileName);
                ErrorLogger.RunFileName = XMLFileName;
                ErrorLogger.Clean();
            }
            catch (Exception)
            {
                ErrorLogger.Log("CASSYS was unable to access or load the Site XML file. Simulation has ended.", ErrLevel.FATAL);
            }

            // ReadFarmSettings reads the overall configuration of the farm
            ReadFarmSettings.doc = SiteSettings;
            ReadFarmSettings.CheckCSYXVersion();


            // Inform user about site configuration:
            Console.WriteLine("Status: Configuring parameters");

            // Configuration of singular instance objects (weather, farm transformer)
            try
            {
                // Weather and radiation related objects.
                // The sun class requires the configuration of the surface slope to calculate the apparent sunset and sunrise hours.
                SimTracker.Config();
                SimSun.itsSurfaceSlope = SimTracker.SurfSlope;
                SimSun.Config();
                
                pyranoTilter.ConfigPyranometer();
                SimTilter.Config();
            }
            catch (XPathException ex)
            {
                ErrorLogger.Log(ex, ErrLevel.FATAL);
            }
            catch (FormatException ex)
            {
                ErrorLogger.Log(ex, ErrLevel.FATAL);
            }

            // Obtain IO File names, either from the command prompt, or from .CSYX
            if ((_Input == null) && (_Output == null))
            {
                ReadFarmSettings.AssignIOFileNames();
            }
            else
            {
                ReadFarmSettings.SimInputFile = _Input;
                ReadFarmSettings.SimOutputFile = _Output;
            }

            // Assign the Sub-Array Count
            ReadFarmSettings.AssignSubArrayCount();

            // collecting input and output file locations and Input/Output file location
            ReadFarmSettings.AssignOutputFileSchema();
            ReadFarmSettings.AssignInputFileSchema();

            if (!ReadFarmSettings.NoSystemDefined)
            {
                // Configuration of shading object
                SimShading.Config();

                // Configuring the transformer object
                SimTransformer.Config();

                // Array of PVArray, Inverter and Wiring objects based on the number of Sub-Arrays 
                SimPVA = new PVArray[ReadFarmSettings.SubArrayCount];
                SimInv = new Inverter[ReadFarmSettings.SubArrayCount];

                // Initialize and Configure PVArray and Inverter Objects through their .CSYX file
                for (int SubArrayCount = 0; SubArrayCount < ReadFarmSettings.SubArrayCount; SubArrayCount++)
                {
                    try
                    {
                        SimInv[SubArrayCount] = new Inverter();
                        SimInv[SubArrayCount].Config(SubArrayCount + 1, SiteSettings);
                        SimPVA[SubArrayCount] = new PVArray();
                        SimPVA[SubArrayCount].Config(SubArrayCount + 1);
                        SimInv[SubArrayCount].ConfigACWiring(SimPVA[SubArrayCount].itsPNomDCArray);
                    }
                    catch (XPathException ex)
                    {
                        ErrorLogger.Log(ex, ErrLevel.FATAL);
                    }
                    catch (FormatException ex)
                    {
                        ErrorLogger.Log(ex, ErrLevel.FATAL);
                    }
                }
            }
            else
            {
                Console.WriteLine("Status: No system was defined but irradiance specifications were found.");
                Console.WriteLine("Status: Only irradiance calculations will be performed in this simulation.");
            }

            // Creating the StreamReader and StreamWriter to Read the Input File and Write to the Output File
            try
            {

                TextFieldParser InputFileReader = new TextFieldParser(ReadFarmSettings.SimInputFile);
                StreamWriter OutputFileWriter = new StreamWriter(ReadFarmSettings.SimOutputFile);

                // if in batch mode, use the appending overload of StreamWriter
                if (ReadFarmSettings.batchMode)
                {
                    OutputFileWriter.Close();
                    OutputFileWriter = new StreamWriter(ReadFarmSettings.SimOutputFile, true);
                }

                // Notifying user of the Configuration status
                if (ErrorLogger.numWarnings == 0)
                {
                    Console.WriteLine("Status: Configuration OK.");
                }
                else
                {
                    Console.WriteLine("Status: There were problems encountered during configuration.");
                    Console.WriteLine("        Please see the error log file for details.");
                    Console.WriteLine("        CASSYS has configured parameters with default values.");
                }

                // Reading the first line and Writing the first line of the output file (Headers of columns)
                try
                {
                    // Skip # of rows that are defined by the User in the .CSYX
                    for (int i = 0; i < ReadFarmSettings.ClimateFileRowsToSkip; i++)
                    {
                        InputFileReader.ReadLine();
                    }

                    // Creating the header of the output file
                    OutputFileWriter.WriteLine(ReadFarmSettings.OutputHeader);
                }
                catch (IOException)
                {
                    ErrorLogger.Log("There was a problem accessing the Climate or Output File. Please check if the file is open or if the path to the file is valid.", ErrLevel.FATAL);
                }
            #endregion

                #region SIMULATION PROCESSES HERE

                #region Read Input File, Perform: Irradiance and Weather Calculations
                // Read through the input file and perform calculations
                Console.Write("Status: Simulation Running");


                // Setting the delimiter or fixed width type to read the Input file (Only TM2 is treated differently)
                if (ReadFarmSettings.SimInputFile.Contains(".tm2"))
                {
                    InputFileReader.TextFieldType = FieldType.FixedWidth;
                    InputFileReader.SetFieldWidths(3, 2, 2, 2, 4, 4, 4, 1, 1, 4, 1, 1, 4, 1, 1, 4, 1, 1, 4, 1, 1, 4, 1, 1, 4, 1, 1, 2, 1, 1, 2, 1, 1, 4, 1, 1, 4, 1, 1, 3, 1, 1, 4, 1, 1, 3, 1, 1, 3, 1, 1, 4, 1, 1, 5, 1, 1, 1, 3, 1, 1, 3, 1, 1, 3, 1, 1, 2, 1, 1);
                }
                else
                {
                    InputFileReader.SetDelimiters(ReadFarmSettings.delim);
                }


                // Read through the Input File 
                while (!InputFileReader.EndOfData)
                {
                    // Iteration counter is used to display a running status dot to the user every 1000 Input lines processed
                    if (ErrorLogger.iterationCount % 1000 == 0)
                    {
                        Console.Write('.');
                    }
                    ErrorLogger.iterationCount++;

                    // Read a TMY or CSV File. The TMY Type 2 is a TM2, and TMY Type 3 is a TM3. If no type is found, a user defined file is used.  
                    if (ReadFarmSettings.TMYType == 2)
                    {
                        MetReader.ParseTM2Line(InputFileReader, SimMet);
                    }
                    else if (ReadFarmSettings.TMYType == 3)
                    {
                        MetReader.ParseTM3Line(InputFileReader, SimMet);
                    }
                    else
                    {
                        MetReader.ParseCSVLine(InputFileReader, SimMet);
                    }



                    // Analyse the TimeStamp for use in Solar Calculations.
                    // Get the value for the Day of the Year, and Hour from Time Stamp [# 1->365,# 0->24, # 1->12]
                    Utilities.TSBreak(SimMet.TimeStamp, out DayOfYear, out HourOfDay, out MonthOfYear, out TimeStepEnd, out TimeStepBeg);

                    // Calculating Sun position
                    // Calculate the Solar Azimuth, and Zenith angles [radians]
                    SimSun.itsSurfaceSlope = SimTracker.SurfSlope;
                    SimSun.Calculate(DayOfYear, HourOfDay);

                    // The time stamp must be adjusted for sunset and sunrise hours such that the position of the sun is only calculated
                    // for the middle of the interval where the sun is above the horizon.
                    if ((TimeStepEnd > SimSun.TrueSunSetHour) && (TimeStepBeg < SimSun.TrueSunSetHour))
                    {
                        HourOfDay = TimeStepBeg + (SimSun.TrueSunSetHour - TimeStepBeg) / 2;

                        // The sun has set, so the transformer should now be disconnected (only used if transformer is disconnected at night)
                        SimTransformer.isDisconnectedNow = true;
                    }
                    else if ((TimeStepBeg < SimSun.TrueSunRiseHour) && (TimeStepEnd > SimSun.TrueSunRiseHour))
                    {
                        HourOfDay = SimSun.TrueSunRiseHour + (TimeStepEnd - SimSun.TrueSunRiseHour) / 2;

                        // The sun has risen, so the transformer should now be Connected (only used if transformer is disconnected at night)
                        SimTransformer.isDisconnectedNow = false;
                    }
                    
                    // Based on the definition of Input file, use Tilted irradiance or transpose the horizontal irradiance
                    if (ReadFarmSettings.UsePOA == true)
                    {
                        // Check if the meter tilt and surface tilt are equal, if not detranspose the pyranometer
                        if (ReadFarmSettings.CASSYSCSYXVersion == "0.9.2" || ReadFarmSettings.CASSYSCSYXVersion == "0.9.3")
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
                                PyranoDetranspose();
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
                                Detranspose();
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
                            Detranspose();
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

                        Transpose();
                    }

                #endregion

                    // If Irradiance is the only item required by the user do not do the calculations.
                    if (!ReadFarmSettings.NoSystemDefined)
                    {
                        // Calculate shading and determine the values of tilted radiation components based on shading factors
                        SimShading.Calculate(SimSun.Zenith, SimSun.Azimuth, SimTilter.TDir, SimTilter.TDif, SimTilter.TRef, SimTracker.SurfSlope, SimTracker.SurfAzimuth);

                        #region PV Array and Inverter calculations
                        try
                        {
                            // Calculate PV Array Output for inputs read in this loop
                            for (int j = 0; j < ReadFarmSettings.SubArrayCount; j++)
                            {
                                // Adjust the IV Curve based on based on Temperature and Irradiance
                                SimPVA[j].CalcIVCurveParameters(SimMet.TGlo, SimShading.ShadTDir, SimShading.ShadTDif, SimShading.ShadTRef, SimTilter.IncidenceAngle, SimMet.TAmbient, SimMet.WindSpeed, SimMet.TModMeasured, MonthOfYear);

                                // Check Inverter status to determine if the Inverter is ON or OFF
                                GetInverterStatus(j);

                                if (SimInv[j].isON)
                                {
                                    // If ON and If the PVArray Voltage in the MPPT Window, calculate the Inverter Output
                                    if (SimInv[j].inMPPTWindow)
                                    {
                                        SimPVA[j].Calculate(true, 0);
                                        SimInv[j].Calculate(SimPVA[j].POut, SimInv[j].VInDC);
                                        // If the Inverter is Clipping, the voltage is increased till the Inverter will not Clip anymore. 
                                        if (SimInv[j].isClipping)
                                        {
                                            GetClippingVoltage(j);
                                            SimPVA[j].Calculate(false, SimInv[j].VInDC);
                                            SimInv[j].Calculate(SimPVA[j].POut, SimInv[j].VInDC);
                                        }
                                    }
                                    else
                                    {
                                        // If ON and if the PV Array Voltage is NOT in the MPPT Window, re-calculate with the PV Array at Fixed Voltage Mode
                                        SimPVA[j].Calculate(false, SimInv[j].VInDC);
                                        GetInverterStatus(j);

                                        if ((SimInv[j].isON == true) && (SimInv[j].inMPPTWindow == false))
                                        {
                                            SimInv[j].Calculate(SimPVA[j].POut, SimInv[j].VInDC);

                                            if (SimInv[j].isClipping)
                                            {
                                                GetClippingVoltage(j);
                                                SimPVA[j].Calculate(false, SimInv[j].VInDC);
                                                SimInv[j].Calculate(SimPVA[j].POut, SimInv[j].VInDC);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    // If the Inverter is OFF, everything should be 0.
                                    SimPVA[j].VOut = SimInv[j].VInDC;
                                    SimPVA[j].IOut = 0;
                                    SimPVA[j].POut = 0;
                                    SimPVA[j].OhmicLosses = 0;
                                    SimPVA[j].MismatchLoss = 0;
                                    SimPVA[j].ModuleQualityLoss = 0;
                                    SimPVA[j].SoilingLoss = 0;
                                    SimInv[j].IOut = 0;
                                    SimInv[j].ACWiringLoss = 0;
                                }

                                // Assigning the outputs to the dictionary
                                ReadFarmSettings.Outputlist["SubArray_Current" + (j + 1).ToString()] = SimPVA[j].IOut;
                                ReadFarmSettings.Outputlist["SubArray_Voltage" + (j + 1).ToString()] = SimPVA[j].VOut;
                                ReadFarmSettings.Outputlist["SubArray_Power" + (j + 1).ToString()] = SimPVA[j].POut / 1000;
                                ReadFarmSettings.Outputlist["SubArray_Current_Inv" + (j + 1).ToString()] = SimInv[j].IOut;
                                ReadFarmSettings.Outputlist["SubArray_Voltage_Inv" + (j + 1).ToString()] = SimInv[j].itsOutputVoltage;
                                ReadFarmSettings.Outputlist["SubArray_Power_Inv" + (j + 1).ToString()] = SimInv[j].ACPwrOut / 1000;
                            }
                        }
                        catch (CASSYSException ce)
                        {
                            ErrorLogger.Log(ce, ErrLevel.FATAL);
                        }

                        #endregion


                        #region Transformer and Net AC Side Calculations
                        // Calculate total Farm output (AC, W) to Grid
                        try
                        {
                            farmACOutput = 0;
                            farmACOhmicLoss = 0;
                            for (int i = 0; i < SimInv.Length; i++)
                            {
                                farmACOutput += SimInv[i].ACPwrOut;
                                farmACOhmicLoss += SimInv[i].ACWiringLoss;
                            }

                            SimTransformer.Calculate(farmACOutput - farmACOhmicLoss);
                        }
                        catch (ArithmeticException AE)
                        {
                            ErrorLogger.Log(AE, ErrLevel.WARNING);
                        }

                        #endregion

                    }

                    #region Writing resultant values to the file and Console Window
                    try
                    {
                        // Assembling and writing the line containing all output values
                        OutputFileWriter.WriteLine(String.Join(",", GetOutputLine()));
                    }
                    catch (IOException ex)
                    {
                        ErrorLogger.Log(ex, ErrLevel.WARNING);
                    }
                }

                // Clean out the buffer of the writer to ensure all entries are written to the output file.
                try
                {
                    OutputFileWriter.Flush();
                }
                catch (IOException ex)
                {
                    ErrorLogger.Log(ex, ErrLevel.WARNING);
                }
                InputFileReader.Dispose();
            }

            catch (ArithmeticException AE)
            {
                ErrorLogger.Log(AE, ErrLevel.WARNING);
                // Mark all outputs with bad data tag
                farmDC = Util.BADDATA;
                farmDCModuleQualityLoss = Util.BADDATA;
                farmDCMismatchLoss = Util.BADDATA;
                farmDCOhmicLoss = Util.BADDATA;
                farmDCSoilingLoss = Util.BADDATA;
                farmDCTemp = Util.BADDATA;
                farmACOutput = Util.BADDATA;
                farmACOhmicLoss = Util.BADDATA;
                pGrid = Util.BADDATA;
            }
            catch (IOException ex)
            {
                if (ex.ToString().Contains("Writer") == true)
                {
                    ErrorLogger.Log("The output file name " + ReadFarmSettings.SimOutputFile + " or volume label syntax is incorrect.", ErrLevel.FATAL);
                }
                else
                {
                    ErrorLogger.Log("The input file name " + ReadFarmSettings.SimInputFile + " or volume label syntax is incorrect.", ErrLevel.FATAL);
                }
            }

            timeTaken.Stop();

            Console.WriteLine("");
            Console.WriteLine("Status: Complete. Simulation took " + timeTaken.ElapsedMilliseconds / 1000D + " seconds.");

                    #endregion

                #endregion
        }

        #region Methods used in this program

        // Calculates the Voltage at which the Inverter will produce Nom AC Power (when Clipping) using Bisection Method.
        void GetClippingVoltage(int j)
        {
            SimInv[j].LossClipping = SimPVA[j].POut;                        // The input power that begins the clipping

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

            SimInv[j].LossClipping -= SimInv[j].ACPwrOut;
            SimInv[j].VInDC = trialInvV;
        }

        // Checks the status of the Inverter (ON, MPPT tracking, etc) and configures its operation based on the PV Array's characteristics.
        void GetInverterStatus(int j)
        {
            // If the Inverter is off, check if the Open Circuit Voltage of the Array is sufficient to turn the Inverter ON
            if (SimInv[j].isON == false)
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
                    SimInv[j].VInDC = 0;
                    SimInv[j].ACPwrOut = 0;
                    SimInv[j].inMPPTWindow = false;
                }
                else
                {
                    SimInv[j].hasMinVoltage = true;
                    SimInv[j].isON = true;
                    SimInv[j].inMPPTWindow = false;
                }
            }

            // If the Inverter turns on because of sufficient voltage, or if the Inverter was already ON  
            // Check if the Incoming Array Power with MPP Operation is sufficient to keep it ON
            if (SimInv[j].isON)
            {
                SimPVA[j].Calculate(true, 0);
                double arrayVMPP = SimPVA[j].VOut;
                double arrayPMPP = SimPVA[j].POut;

                // MPPT check, if true then use voltage window to determine voltage out of Inverter and if it is in the MPPT Window
                if (SimInv[j].isBipolar)
                {
                    arrayVMPP /= 2;
                }

                SimInv[j].GetMPPTStatus(arrayVMPP, out SimInv[j].inMPPTWindow);

                if (SimInv[j].inMPPTWindow)
                {
                    // Check if the inverter has sufficient power to stay ON 
                    if (arrayPMPP > (SimInv[j].itsThresholdPwr * SimInv[j].itsNumInverters))
                    {
                        SimInv[j].isON = true;
                    }
                    else
                    {
                        // Inverter must be turned off.
                        SimInv[j].LossPMinThreshold = SimPVA[j].POut;
                        SimInv[j].hasMinVoltage = false;
                        SimInv[j].isON = false;
                        SimInv[j].VInDC = SimPVA[j].VOut;
                        SimInv[j].ACPwrOut = 0;
                        SimInv[j].inMPPTWindow = false;
                    }
                }
                else
                {
                    SimPVA[j].Calculate(false, SimInv[j].VInDC);
                    // Check if the inverter has sufficient power to stay ON after the voltage has been pinned to the Min/Max PT Voltage level
                    if (SimPVA[j].POut > (SimInv[j].itsThresholdPwr * SimInv[j].itsNumInverters))
                    {
                        SimInv[j].isON = true;
                    }
                    else
                    {
                        // Inverter must be turned off.
                        SimInv[j].LossPMinThreshold = SimPVA[j].POut > 0 ? SimPVA[j].POut : 0;
                        SimInv[j].hasMinVoltage = false;
                        SimInv[j].isON = false;
                        SimInv[j].ACPwrOut = 0;
                        SimInv[j].inMPPTWindow = false;
                    }

                }
            }
        }

        // Transposition of the global horizontal irradiance values to the transposed values
        void Transpose()
        {
            SimSun.Calculate(DayOfYear, HourOfDay);
            
            // Calculating the Surface Slope and Azimuth based on the Tracker Chosen
            SimTracker.Calculate(SimSun.Zenith, SimSun.Azimuth);
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
            SimTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, MonthOfYear);
        }

        // De-transposition of the titled irradiance values to the global horizontal values
        void Detranspose()
        {
            // Lower bound of bisection
            double HGloLo = 0;

            // Higher bound of bisection
            double HGloHi = SimSun.NExtra;

            // Calculating the Surface Slope and Azimuth based on the Tracker Chosen
            SimTracker.Calculate(SimSun.Zenith, SimSun.Azimuth);
            SimTilter.itsSurfaceSlope = SimTracker.SurfSlope;
            SimTilter.itsSurfaceAzimuth = SimTracker.SurfAzimuth;
            SimTilter.IncidenceAngle = SimTracker.IncidenceAngle;

            // Calculating the Incidence Angle for the current setup
            double cosInc = Tilt.GetCosIncidenceAngle(SimSun.Zenith, SimSun.Azimuth, SimTilter.itsSurfaceSlope, SimTilter.itsSurfaceAzimuth);

            // Trivial case
            if (SimMet.TGlo <= 0)
            {
                SimSplitter.Calculate(SimSun.Zenith, 0, NExtra: SimSun.NExtra);
                SimTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, MonthOfYear);
            }
            else if ((SimSun.Zenith > 87.5 * Util.DTOR) || (cosInc <= Math.Cos(87.5 * Util.DTOR)))
            {
                SimMet.HGlo = SimMet.TGlo / ((1 + Math.Cos(SimTilter.itsSurfaceSlope)) / 2 + SimTilter.itsMonthlyAlbedo[MonthOfYear] * (1 - Math.Cos(SimTilter.itsSurfaceSlope)) / 2);

                // Forcing the horizontal irradiance to be composed entirely of diffuse irradiance
                SimSplitter.HGlo = SimMet.HGlo;
                SimSplitter.HDif = SimMet.HGlo;
                SimSplitter.NDir = 0;
                SimSplitter.HDir = 0;

                //SimSplitter.Calculate(SimSun.Zenith, HGlo, NExtra: SimSun.NExtra);
                SimTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, MonthOfYear);
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
                    SimTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, MonthOfYear);
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
        void PyranoDetranspose()
        {
            if (pyranoTilter.NoPyranoAnglesDefined)
            {
                SimTracker.Calculate(SimSun.Zenith, SimSun.Azimuth);
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
                pyranoTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, MonthOfYear);
            }
            else if ((SimSun.Zenith > 87.5 * Util.DTOR) || (cosInc <= Math.Cos(87.5 * Util.DTOR)))
            {
                SimMet.HGlo = SimMet.TGlo / ((1 + Math.Cos(pyranoTilter.itsSurfaceSlope)) / 2 + pyranoTilter.itsMonthlyAlbedo[MonthOfYear] * (1 - Math.Cos(pyranoTilter.itsSurfaceSlope)) / 2);

                // Forcing the horizontal irradiance to be composed entirely of diffuse irradiance
                SimSplitter.HGlo = SimMet.HGlo;
                SimSplitter.HDif = SimMet.HGlo;
                SimSplitter.NDir = 0;
                SimSplitter.HDir = 0;

                //SimSplitter.Calculate(SimSun.Zenith, HGlo, NExtra: SimSun.NExtra);
                pyranoTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, MonthOfYear);
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
                    pyranoTilter.Calculate(SimSplitter.NDir, SimSplitter.HDif, SimSun.NExtra, SimSun.Zenith, SimSun.Azimuth, SimSun.AirMass, MonthOfYear);
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
            Transpose();
        }
        
        // Gather the information for the output as per user selection.
        String GetOutputLine()
        {
            // Shading each component of the Tilted radiaton
            ShadBeamLoss = SimTilter.TDir - SimShading.ShadTDir;
            ShadDiffLoss = SimTilter.TDif > 0 ? SimTilter.TDif - SimShading.ShadTDif : 0;
            ShadRefLoss = SimTilter.TRef > 0 ? SimTilter.TRef - SimShading.ShadTRef : 0;

            // Setting all output variables to 0 to enable calculations
            farmDC = 0;                                  // Farm/PVArray DC Output [W]
            farmDCModuleQualityLoss = 0;                 // Farm/PVArray DC Module Quality Loss (Sum for all sub-arrays) [W]
            farmDCMismatchLoss = 0;                      // Farm/PVArray DC Module Mismatch Loss (Sum for all sub-arrays) [W]
            farmDCOhmicLoss = 0;                         // Farm/PVArray DC Ohmic Loss (Sum for all sub-arrays) [W]
            farmDCSoilingLoss = 0;                       // Farm/PVArray DC Soiling Loss (Sum for all sub-arrays) [W]
            farmDCCurrent = 0;                           // Farm/PVArray DC Current Values [A]
            farmDCTemp = 0;                              // Average temperature of all PV Arrays [deg C]
            farmPNomDC = 0;                              // Farm/PVArray DC Nominal Output [W] for normalized calculations
            farmPNomAC = 0;                              // Farm/Inverter AC Nominal Output [W] for normalized calculations
            double farmArea = 0;                         // Rough farm area (based on PV Array * Number of Modules in each Sub Array)
            farmTotalModules = 0;                        // Total number of modules in the farm [#]
            farmACPMinThreshLoss = 0;                    // Total loss when the power of the array is not sufficient for starting the inverter. 
            farmACClippingPower = 0;                     // Produced power before reduction by Inverter (clipping) [W]

            // Calculate and assign values to the outputs required by the program.
            if (!ReadFarmSettings.NoSystemDefined)
            {
                for (int i = 0; i < SimPVA.Length; i++)
                {
                    farmDC += SimPVA[i].POut;
                    farmDCCurrent += SimPVA[i].IOut;
                    farmDCMismatchLoss += SimPVA[i].MismatchLoss;
                    farmDCModuleQualityLoss += SimPVA[i].ModuleQualityLoss;
                    farmDCOhmicLoss += SimPVA[i].OhmicLosses;
                    farmDCSoilingLoss += SimPVA[i].SoilingLoss;
                    farmDCTemp += SimPVA[i].TModule * SimPVA[i].itsNumModules;
                    farmTotalModules += SimPVA[i].itsNumModules;
                    farmPNomDC += SimPVA[i].itsPNomDCArray;
                    farmPNomAC += SimInv[i].itsPNomArrayAC;
                    farmArea += SimPVA[i].itsRoughArea;
                    farmACPMinThreshLoss += SimInv[i].LossPMinThreshold;
                    farmACClippingPower += SimInv[i].LossClipping;
                    SimInv[i].LossPMinThreshold = 0;
                    SimInv[i].LossClipping = 0;
                }


                // Averages all PV Array temperature values
                farmDCTemp /= farmTotalModules;
                farmPNomDC = Utilities.ConvertWtokW(farmPNomDC);
                farmPNomAC = Utilities.ConvertWtokW(farmPNomAC);
            }

            // Using the TimeSpan function to assemble the String for the modified Time Stamp of calculation
            TimeSpan thisHour = TimeSpan.FromHours(HourOfDay);
            TimeStampAnalyzed = new DateTime(Utilities.CurrentTimeStamp.Year, Utilities.CurrentTimeStamp.Month, Utilities.CurrentTimeStamp.Day, thisHour.Hours, thisHour.Minutes, thisHour.Seconds);

            //  Assigning all outputs their corresponding values;
            ReadFarmSettings.Outputlist["Input_Timestamp"] = SimMet.TimeStamp;
            ReadFarmSettings.Outputlist["Timestamp_Used_for_Simulation"] = String.Format("{0:u}", TimeStampAnalyzed).Replace('Z', ' ');
            ReadFarmSettings.Outputlist["Sun_Zenith_Angle"] = Util.RTOD * SimSun.Zenith;
            ReadFarmSettings.Outputlist["Sun_Azimuth_Angle"] = Util.RTOD * SimSun.Azimuth;
            ReadFarmSettings.Outputlist["ET_Irrad"] = SimSun.NExtra;
            ReadFarmSettings.Outputlist["Albedo"] = SimTilter.itsMonthlyAlbedo[Utilities.CurrentTimeStamp.Month];
            ReadFarmSettings.Outputlist["Normal_beam_irradiance"] = SimSplitter.NDir;
            ReadFarmSettings.Outputlist["Horizontal_Global_Irradiance"] = SimSplitter.HGlo;
            ReadFarmSettings.Outputlist["Horizontal_diffuse_irradiance"] = SimSplitter.HDif;
            ReadFarmSettings.Outputlist["Horizontal_beam_irradiance"] = SimSplitter.HDir;
            ReadFarmSettings.Outputlist["Ambient_Temperature"] = SimMet.TAmbient;
            ReadFarmSettings.Outputlist["Wind_Velocity"] = SimMet.WindSpeed;
            ReadFarmSettings.Outputlist["Global_Irradiance_in_Array_Plane"] = SimTilter.TGlo;
            ReadFarmSettings.Outputlist["Beam_Irradiance_in_Array_Plane"] = SimTilter.TDir;
            ReadFarmSettings.Outputlist["Diffuse_Irradiance_in_Array_Plane"] = SimTilter.TDif;
            ReadFarmSettings.Outputlist["Ground_Reflected_Irradiance_in_Array_Plane"] = SimTilter.TRef;
            // NB: moved these outputs so they are processed with or without a system
            ReadFarmSettings.Outputlist["Tracker_Slope"] = SimTracker.itsTrackerSlope * Util.RTOD;
            ReadFarmSettings.Outputlist["Tracker_Azimuth"] = SimTracker.itsTrackerAzimuth * Util.RTOD;
            ReadFarmSettings.Outputlist["Tracker_Rotation_Angle"] = SimTracker.RotAngle * Util.RTOD;
            ReadFarmSettings.Outputlist["Collector_Surface_Slope"] = SimTilter.itsSurfaceSlope * Util.RTOD;
            ReadFarmSettings.Outputlist["Collector_Surface_Azimuth"] = SimTilter.itsSurfaceAzimuth * Util.RTOD;
            ReadFarmSettings.Outputlist["Incidence_Angle"] = Math.Min(Util.RTOD * SimTilter.IncidenceAngle, 90);
            

            // If a system is defined get all the other performance characteristics needed
            if (!ReadFarmSettings.NoSystemDefined)
            {
                ReadFarmSettings.Outputlist["Global_POA_Irradiance_Corrected_for_Shading"] = SimShading.ShadTGlo;                
                ReadFarmSettings.Outputlist["Near_Shading_Loss_for_Global"] = SimTilter.TGlo - SimShading.ShadTGlo;
                ReadFarmSettings.Outputlist["Near_Shading_Loss_for_Beam"] = ShadBeamLoss;
                ReadFarmSettings.Outputlist["Near_Shading_Loss_for_Diffuse"] = ShadDiffLoss;
                ReadFarmSettings.Outputlist["Near_Shading_Loss_for_Ground_Reflected"] = ShadRefLoss;
                ReadFarmSettings.Outputlist["Global_POA_Irradiance_Corrected_for_Incidence"] = SimPVA[0].IAMTGlo;
                ReadFarmSettings.Outputlist["Incidence_Loss_for_Global"] = SimShading.ShadTGlo - SimPVA[0].IAMTGlo;
                ReadFarmSettings.Outputlist["Incidence_Loss_for_Beam"] = SimShading.ShadTDir * (1 - SimPVA[0].IAMDir);
                ReadFarmSettings.Outputlist["Incidence_Loss_for_Diffuse"] = SimShading.ShadTDif * (1 - SimPVA[0].IAMDif);
                ReadFarmSettings.Outputlist["Incidence_Loss_for_Ground_Reflected"] = SimShading.ShadTRef * (1 - SimPVA[0].IAMRef);
                ReadFarmSettings.Outputlist["Profile_Angle"] = Util.RTOD * SimShading.ProfileAng;
                ReadFarmSettings.Outputlist["Near_Shading_Factor_on_Global"] = (SimTilter.TGlo > 0 ? SimShading.ShadTGlo / SimTilter.TGlo : 1);
                ReadFarmSettings.Outputlist["Near_Shading_Factor_on_Beam"] = SimShading.BeamSF;
                ReadFarmSettings.Outputlist["Near_Shading_Factor_on__Diffuse"] = SimShading.DiffuseSF;
                ReadFarmSettings.Outputlist["Near_Shading_Factor_on_Ground_Reflected"] = SimShading.ReflectedSF;
                ReadFarmSettings.Outputlist["IAM_Factor_on_Global"] = (SimShading.ShadTGlo > 0 ? SimPVA[0].IAMTGlo / SimShading.ShadTGlo : 1);
                ReadFarmSettings.Outputlist["IAM_Factor_on_Beam"] = SimPVA[0].IAMDir;
                ReadFarmSettings.Outputlist["IAM_Factor_on__Diffuse"] = SimPVA[0].IAMDif;
                ReadFarmSettings.Outputlist["IAM_Factor_on_Ground_Reflected"] = SimPVA[0].IAMRef;
                ReadFarmSettings.Outputlist["Array_Soiling_Loss"] = farmDCSoilingLoss / 1000;
                ReadFarmSettings.Outputlist["Modules_Array_Mismatch_Loss"] = farmDCMismatchLoss / 1000;
                ReadFarmSettings.Outputlist["Ohmic_Wiring_Loss"] = farmDCOhmicLoss / 1000;
                ReadFarmSettings.Outputlist["Module_Quality_Loss"] = farmDCModuleQualityLoss / 1000;
                ReadFarmSettings.Outputlist["Effective_Energy_at_the_Output_of_the_Array"] = farmDC / 1000;
                ReadFarmSettings.Outputlist["Average_Ambient_Temperature_deg_C_"] = SimMet.TAmbient;
                ReadFarmSettings.Outputlist["Calculated_Module_Temperature__deg_C_"] = farmDCTemp;
                ReadFarmSettings.Outputlist["Measured_Module_Temperature__deg_C_"] = SimMet.TModMeasured;
                ReadFarmSettings.Outputlist["Difference_between_Module_and_Ambient_Temp.__deg_C_"] = farmDCTemp - SimMet.TAmbient;
                ReadFarmSettings.Outputlist["PV_Array_Current"] = farmDCCurrent;
                ReadFarmSettings.Outputlist["PV_Array_Voltage"] = (farmDCCurrent > 0 ? farmDC / farmDCCurrent : 0);
                ReadFarmSettings.Outputlist["Available_Energy_at_Inverter_Output"] = farmACOutput / 1000;
                ReadFarmSettings.Outputlist["AC_Ohmic_Loss"] = farmACOhmicLoss / 1000;
                ReadFarmSettings.Outputlist["Inverter_Efficiency"] = (farmACOutput > 0 ? farmACOutput / farmDC : 0) * 100;
                ReadFarmSettings.Outputlist["Inverter_Loss_Due_to_Power_Threshold"] = farmACPMinThreshLoss / 1000;
                ReadFarmSettings.Outputlist["Inverter_Loss_Due_to_Voltage_Threshold"] = 0;
                ReadFarmSettings.Outputlist["Inverter_Loss_Due_to_Nominal_Inv._Power"] = farmACClippingPower > 0 ? farmACClippingPower / 1000 : 0;
                ReadFarmSettings.Outputlist["Inverter_Loss_Due_to_Nominal_Inv._Voltage"] = 0;
                ReadFarmSettings.Outputlist["External_transformer_loss"] = SimTransformer.Losses / 1000;
                ReadFarmSettings.Outputlist["Power_Injected_into_Grid"] = SimTransformer.POut / 1000;
                ReadFarmSettings.Outputlist["Energy_Injected_into_Grid"] = SimTransformer.EnergyToGrid / 1000;
                ReadFarmSettings.Outputlist["PV_Array_Efficiency"] = (SimTilter.TGlo > 0 ? farmDC / (SimTilter.TGlo * farmArea) : 0) * 100;
                ReadFarmSettings.Outputlist["AC_side_Efficiency"] = (farmACOutput > 0 && SimTransformer.POut > 0 ? SimTransformer.POut / farmACOutput : 0) * 100;
                ReadFarmSettings.Outputlist["Overall_System_Efficiency"] = (SimTilter.TGlo > 0 && SimTransformer.POut > 0 ? SimTransformer.POut / (SimTilter.TGlo * farmArea) : 0) * 100;
                ReadFarmSettings.Outputlist["Normalized_System_Production"] = SimTransformer.POut > 0 ? SimTransformer.POut / (farmPNomDC * 1000) : 0;
                ReadFarmSettings.Outputlist["Array_losses_ratio"] = SimTransformer.POut > 0 ? (farmDCMismatchLoss + farmDCModuleQualityLoss + farmDCOhmicLoss + farmDCSoilingLoss) / SimTransformer.POut : 0;
                ReadFarmSettings.Outputlist["Inverter_losses_ratio"] = SimTransformer.POut > 0 ? farmACOhmicLoss / SimTransformer.POut : 0;
                ReadFarmSettings.Outputlist["AC_losses_ratio"] = SimTransformer.Losses / SimTransformer.POut < 0 ? 0 : SimTransformer.Losses / SimTransformer.POut;
                ReadFarmSettings.Outputlist["Performance_Ratio"] = SimTilter.TGlo > 0 && farmPNomDC > 0 && SimTransformer.POut > 0 ? SimTransformer.POut / SimTilter.TGlo / farmPNomDC : 0;
                ReadFarmSettings.Outputlist["System_Loss_Incident_Energy_Ratio"] = (SimTransformer.itsPNom - SimTransformer.POut) / (SimTilter.TGlo * 1000);
            }

            // Constructing the OutputLine to be written to string;
            string OutputLine = null;
            foreach (String required in ReadFarmSettings.OutputScheme)
            {
                // Check if the required Output Value exists in the dictionary and then print                
                if (required == "Sub_Array_Performance" && !ReadFarmSettings.NoSystemDefined)
                {
                    // Get the power for individual Sub-Arrays
                    for (int subNum = 1; subNum < SimPVA.Length + 1; subNum++)
                    {
                        string temp = ReadFarmSettings.Outputlist["SubArray_Voltage" + subNum].ToString() + "," + ReadFarmSettings.Outputlist["SubArray_Current" + subNum].ToString() + "," + ReadFarmSettings.Outputlist["SubArray_Power" + subNum].ToString() + ",";
                        OutputLine += temp;
                    }
                }
                else if (required == "ShowSubInv" && !ReadFarmSettings.NoSystemDefined)
                {
                    // Get the power for individual Sub-Arrays
                    for (int subNum = 1; subNum < SimInv.Length + 1; subNum++)
                    {
                        string temp = ReadFarmSettings.Outputlist["SubArray_Power_Inv" + subNum].ToString() + ",";
                        OutputLine += temp;
                    }
                }
                else if (required == "ShowSubInvV" && !ReadFarmSettings.NoSystemDefined)
                {
                    // Get the voltage for individual Sub-Arrays
                    for (int subNum = 1; subNum < SimInv.Length + 1; subNum++)
                    {
                        string temp = ReadFarmSettings.Outputlist["SubArray_Voltage_Inv" + subNum].ToString() + ",";
                        OutputLine += temp;
                    }
                }
                else if (required == "ShowSubInvC" && !ReadFarmSettings.NoSystemDefined)
                {
                    // Get the current for individual Sub-Arrays
                    for (int subNum = 1; subNum < SimInv.Length + 1; subNum++)
                    {
                        string temp = ReadFarmSettings.Outputlist["SubArray_Current_Inv" + subNum].ToString() + ",";
                        OutputLine += temp;
                    }
                }
                else
                {
                    if (ReadFarmSettings.Outputlist[required] != null)
                    {
                        string temp = ReadFarmSettings.Outputlist[required].ToString() + ",";
                        OutputLine += temp;
                    }
                    else
                    {
                        OutputLine += "Not Available,";
                    }
                }
            }

            return OutputLine;
        }
        #endregion
    }
}
