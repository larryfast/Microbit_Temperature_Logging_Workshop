# Microbit_Temperature_Logging_Workshop
### User Setup instructions for Temperature Logging Workshop
If you're new to this project, check out the video introduction (TODO: LOCATION??). This workshop is suitable for experienced Microbit users who want to try something more challenging. 

The remainder of this README are detailed instructions for downloading and running this project

####  Table of Contents
- [Project Requirements](#Project-Requirements)
- [Download the project code to a folder](#Download-the-project-code-to-a-folder)
- [Program your Microbits](#Program-your-microbits)
- [Install Python](#Install-Python)
- [Install the Python Packages used by the code](#Install-the-Python-Packages-used-by-the-code)
- [Start your engines!](#start-your-engines)


- [start_serial2csv with radio2serial microbit connected](#start_serial2csv-with-radio2serial-microbit-connected)
- Power up a Temperature sensing Microbit
- Open live_data.xlsm
- Enable Developer Tab 
- Run the Macro
- Data => Click Refresh
- See your temperature readings in the graph
- Power up more Temperature Microbits
- Hit Refresh and these will start appearing on the graph

### Project Requirements
* Currently this project runs on Windows only. As soon as someone asks, we can set it up to run on Linux and Mac.
- PC running Windows
- Excel (future: Libre Calc)
- USB port for connecting to Microbits - will not work with Microbits connected via Wifi
- Microbits! (of course)
	- At least one Microbit must be version 2.0 or higher
	- [finding your Microbit version number](https://support.microbit.org/support/solutions/articles/19000119162-how-to-identify-the-version-number-of-your-micro-bit#:~:text=20%2C%20V2%2C%20V1.,beneath%20the%20black%20silkscreen%20paint.)

### Download the project code to a folder

In a browser, open the Github project repository
https://github.com/larryfast/Microbit_Temperature_Logging_Workshop

Download a ZIP file of all the code
![](Readme_files/github_download_zip.png)

Pick a location on your computer. Unzip the project files into that folder.

### Program your microbits
The project folder contains 2 .hex files
- microbit-radio2serial.hex 
	- Program one Microbit with this code
- microbit-2s-TemperatureSender.hex
	- Program this code into as many Microbits as desired

If preferred you can use the microbit editor to review the code and program your microbits. Here are the Microbit Project links
- https://makecode.microbit.org/_AEkKMPWjEKss = radio2serial
- https://makecode.microbit.org/_P7JURDf6vU8J = Temperature Sender 2s

#### Optional: Test your microbits 
You can test the microbits without the rest of the project code
- Connect the radio2serial Microbit to your USB port
- Open the Microbit Editor
- Connect the editor to your device
- Power up at least one Temperature Sender
	- Each of these will read the Temperature every 2 seconds and send the result to the radio2serial Microbit
- Display the output graphically
![](Readme_files/MicrobitEditor_ShowData.png)

Note: The funky strings you see (zitaz & tugog in this example) are the Microbit IDs. The IDs are used to uniquely identify each data source.

### Install Python
Download the Python installer from here:
Recommended version: [Windows Installer (64-bit)](https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe)
Installation Page: [Python Releases for Windows | Python.org](https://www.python.org/downloads/windows/)
![](Readme_files/PythonVersionSelected.png)

NOTE: When running the installer byou can review the microbit editor and program your microbits from these projectsoth options should be checked.
![](Readme_files/InstallPython_OptionsChecked.png)

### Install the Python Packages used by the code
The project uses Python code to collect the Temperature readings from radio2serial Microbit and relay them to the spreadsheet. It needs a couple Python packages that are not part of the standard release.

The video introduction has a segment on installing these packages. 
- Open a Powershell window in project folder
- Run this command:   .\install_python_packages.bat

### Configure the spreadsheet for your computer
The short story is run a macro that does the final configuration of the spreadsheet.  Instructions for the following steps are in the video.
- Open the spreadsheet in Excel
- Enable the Developer Menu
- Open Developer Tab
- Select Macros => Workbook Macros
- Run the TODO: MACRO NAME
- Exit Excel
- Save a copy of your configured Excel file

## Start your engines!
The remaining steps power up all the bits and get them talking to each other

### Open spreadsheet in Excel

### start_serial2csv with radio2serial microbit connected
The video introduction has a segment on starting the Python program.
- Open a Powershell window in project folder
- Run this command:   .\start_serial2csv.bat

### Power up one or more Temperature sensing Microbits

### Spreadsheet should show updated Temperatures every minute
If you're in a hurry... 
- open the Data Tab
- Click Refresh
