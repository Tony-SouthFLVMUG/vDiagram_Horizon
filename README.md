# vDiagram Horizon 1.0
vDiagram Horizon 1.0 based off Alan Renouf's vDiagram

## SYNOPSIS
vDiagram Visio Drawing Tool

## DESCRIPTION
Powershell script that will capture and draw in Visio a VMware Horizon Infrastructure.

## REQUIREMENTS
	1. PowerShell - Download Windows Management Framework 5.1 available here and install. https://www.microsoft.com/en-us/download/details.aspx?id=54616
	2. PowerCLI Modules - To install PowerCLI Modules, open Powershell (installed in step above) and run the following command "Install-Module -Name VMware.PowerCLI –Scope CurrentUser"
	3. Visio - Microsoft Visio must be installed in order for the draw feature to work.

## HOW TO RUN SCRIPT
	1. From within Windows, click on the start button.
	2. Type Powershell and right click on the search results and select "Run as administrator".
	3. At the Powershell command prompt navigate to the the directory where you have unzipped the vDiagram files. Example: "cd c:\Users\<your user name>\Downloads\vDiagram_Horizon_1.0.X"
	4. Type the name of the Powershell script. Example: "vDiagram_Horizon_1.0.X.ps1"
	5. Follow directions listed below in "Usage Notes".

## NOTES
	File Name	: vDiagram_Horizon_1.0.1.ps1
	Author		: Tony Gonzalez
	Author		: Jason Hopkins
	Based on	: vDiagram by Alan Renouf
	Version		: 1.0.1

## USAGE NOTES
	Directions:
	1. Ensure to unblock file before unzipping within file properties
	2. Ensure to run as administrator
	3. Required Files:
            PowerCLI or PowerShell 5.0 with PowerCLI Modules installed
            Active connection to vCenter to capture data
            MS Visio
	    
	Prerequisites Tab:
	1. Verify that prerequisites are met on the "Prerequisites" tab.
	2. If not please install needed requirements.
	
	vCenter Info Tab:
	1. Click on "vCenter Info" tab.
	2. Enter name of vCenter.
	3. Enter User Name and Password (password will be hashed and not plain text).
	4. Click on "Connect to vCenter" button.
	
	Capture CSVs for Visio Tab:
	1. Click on "Capture CSVs for Visio" tab.
	2. Click on "Select Output Folder" button and select folder where you would like to output the CSVs to.
	3. Select items you wish to grab data on.
	4. Click on "Collect CSV Data" button.
	
	Draw Visio Tab:
	1. Click on "Select Input Folder" button and select location where CSVs can be found.
	2. Click on "Check for CSVs" button to validate presence of required files.
	3. Click on "Select Output Folder" button and select where location where you would like to save the Visio drawing.
	4. Select drawing that you would like to produce.
	5. Click on "Draw Visio" button.
	6. Click on "Open Visio Drawing" button once "Draw Visio" button says it has completed.

## CHANGE LOG
	- 09/12/2021 - v1.0.1
		Initial release
