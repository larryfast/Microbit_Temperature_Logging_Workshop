# Microbit_Temperature_Logging_Workshop
### User Setup instructions for Temperature Logging Workshop
test
If you're new to this project, check out the video introduction (TODO: LOCATION??). This workshop is suitable for experienced Microbit users who want to try something more challenging. 

####  TOC
- [[#Download the project code to a folder]]
- [[#Program your microbits]]
- [[#Install Python]]
- [[#Install the Python Packages used by the code]]
- 
- start_serial2csv with radio2serial microbit connected
- Power up a Temperature sensing Microbit
- Open live_data.xlsm
- Enable Developer Tab 
- Run the Macro
- Data => Click Refresh
- See your temperature readings in the graph
- Power up more Temperature Microbits
- Hit Refresh and these will start appearing on the graph


### Download the project code to a folder

In a browser, open to the Github project repository
https://github.com/larryfast/Microbit_Temperature_Logging_Workshop

Download a ZIP file of all the code
![[github_download_zip.png]]

Pick a location on your computer. Unzip the project files into that folder.


### Program your microbits
The project folder contains 2 .hex files
- microbit-radio2serial.hex 
	- Program one Microbit with this code
- microbit-2s-TemperatureSender.hex
	- Program this code into as many Microbits as desired

If preferred you can review the microbit editor and program your microbits from these projects
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
![[MicrobitEditor_ShowData.png]]

Note: The funky strings (zitaz & tugog in this example) are the Microbit IDs. The IDs are used to uniquely identify each data source.

### Install Python
Download the Python installer from here:
Recommended version: [Windows Installer (64-bit)](https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe)
Installation Page: [Python Releases for Windows | Python.org](https://www.python.org/downloads/windows/)
![[Pasted image 20230601155107.png]]

NOTE: When running the installer both options should be checked.
![[InstallPython_OptionsChecked.png]]

### Install the Python Packages used by the code
The project uses Python code to collect the Temperature readings from radio2serial Microbit and relay them to the spreadsheet. It needs a couple Python packages that are not part of the standard release.

The video introduction has a segment on installing these packages. 
- Open a Powershell window in project folder
- Run this command:   .\install_python_packages.bat

### start_serial2csv with radio2serial microbit connected
