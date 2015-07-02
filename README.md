# CASSYS - Canadian Solar System Simulator#
*A Simulation Tool for Grid-Connected Photovoltaic Systems*

*Current Version: 0.9.2*

<p align="center">
<img src =https://github.com/CanadianSolar/CASSYS/blob/gh-pages/images/CASSYS-logo.png alt="CASSYS Logo"></img>
</p>
 
### Program goals and description ##
---------------------------------------
 
CASSYS is a computer program used to simulate the performance of photovoltaic grid-connected systems. Using a detailed description of the system provided by the user (arrays, inverters, and balance of system components), site location, and weather conditions (irradiance, temperature) at arbitrary time steps, it calculates the state of the system at each step and provides a detailed estimate of the energy flows and losses in the system.

The goal of CASSYS is to provide a reliable, flexible and user-friendly way to simulate solar grid-connected system performance for operational purposes.  

CASSYS is developed and maintained by [Canadian Solar O&M Inc. (Ontario)](http://www.canadiansolar.com/ "Canadian Solar O&M Inc. (Ontario)"). 

### Software components and languages ##
----------------------------------------
CASSYS is composed of two main software components: 

 1. User Interface: A macro-enabled Microsoft Excel Workbook called CASSYS.XLSM. This is your main tool to interact with the program. With the workbook you can define the system, run the simulation, and retrieve and analyze the results.
 2. Simulation Engine: A C# program called CASSYS.exe which performs the actual simulation. Do not run CASSYS.exe directly; use the interface to launch it.
 
### Installing and running the program ##
-----------------------------------------
The easiest way to install the program is to download the  [Installer](https://github.com/CanadianSolar/CASSYS/blob/master/CASSYS%20Installer.exe?raw=true "Installer").

The installer lets you specify the directory in which you want the program to be installed. After installation, go to the selected directory, open the CASSYS.XLSM file in MS Excel, and start interacting with the program.

### Documentation ##
--------------------
[Documents and Help](https://github.com/CanadianSolar/CASSYS/tree/master/Documents%20and%20Help "Documents and Help")
contains a user's manual and various documents to better understand the models used in the simulation engine. See [Release Notes](https://github.com/CanadianSolar/CASSYS/wiki/Release-Notes "Release Notes") for a list of changes made to newer versions of CASSYS.

### Licensing ##
----------------
That's the best part: CASSYS is free and open source software. You are free to use it and explore the code, subject to the conditions expressed in the  [Licensing Agreement](https://github.com/CanadianSolar/CASSYS/blob/master/LICENSE "Licensing Agreement"). Feel free to send your comments (positive or negative) or provide suggestions using the 
[Report Issues](https://github.com/CanadianSolar/CASSYS/issues "Report Issues") Link.

 



