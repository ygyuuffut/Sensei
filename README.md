# SENSEI - A Clustered Applets for Excel
![License](https://img.shields.io/badge/License-AGPL%203-BB8FCE?style=flat-square)
![Version](https://img.shields.io/badge/Version-1.4.0-76D754?style=flat-square)
![Version Type](https://img.shields.io/badge/Type-Release-16A085?style=flat-square)
![Language](https://img.shields.io/badge/Language-Virtual%20Basic-EB984E?style=flat-square)

A very much scatter-mind driven programming that develops as demands emerges

## Table of Contents
- [SENSEI - A Clustered Applets for Excel](#sensei---a-clustered-applets-for-excel)
	- [Table of Contents](#table-of-contents)
	- [Structure Guidances](#structure-guidances)
	- [Installations](#installations)
		- [Manual Composition](#manual-composition)
		- [Alternative Method](#alternative-method)
	- [Todo Lists](#todo-lists)


## Structure Guidances

**Cask** -  Contains the Source document in .xlsx, which is non-Macro Enabled structure for the program to operate upon.

**Forms** - Contains the Main Operating Body of the Device, which requires some import actions to have them fully setup.

**Modules** - The Rest of Core functions that are isolated from the actual Operating Body for modular purposes.


## Installations
There are two methods which this could be setup and become operational on any Windows device, with one significantly easier than the other.

### Manual Composition
> 1. Download the "SENSEI - dev.xlsx" located within "Cask" folder
> 2. Open downloaded file, in Excel (Microsoft 365) save as "SENSEI - dev.xlsm" to enable macro holding
> 3. Download Respective Modules and Forms from "Modules" and "Forms" folder, import them through VBA interface within Excel 
> 4. VBA interface can be open thru Keybind: Alt + F11; or right click ribbon to customize it so it display Developer tab.
> 5. Import all the Modules and Forms one by one till done

### Alternative Method
> 1. Use the Release tab get the composed Zip file that contains "*.xlsm" file and start there instead.


## Todo Lists
- [x] [1.0.2] SENSEI LINK - WHICH IS THE INTEGRATION OF SENSEI'S EXTENSIVE ABILITY TO CONTROL ANOTHER WORKBOOK
  - [x] [1.0.2] MISSING SUPPORT FOR LINKING 114 (STORAGE ACCESS ON SENSEI.DATA)
- [ ] [ N/A ] SENSEI LINK MASTER 114 - A BASIC DRIVER FOR 114 INFINITE EDITION
	- [ ] APPEND ABILITY TO LOOKUP SSN BASED ON DODID, CLIP TO CLIPBOARD
	- [ ] APPEND ABILITY TO FILL DODID OR SSAN FOR LOOKUP PURPOSES

- [ ] [1.0.0] SENSEI Form Distller - DEBT FORM WHICH WILL FUNCTION BASED ON FORM 110 INCLUDED
	- [ ] [1.0.0] Project Form 110
		- [ ] [1.0.0] FILL THE 110
			- [x] [1.0.0] Load DATA
			- [x] [1.0.0] Write DATA
			- [x] [1.0.0] trigger Update
			- [x] [1.0.0] OPTIONAL CHANGE PAGE FUNCTION
		- [x] [1.0.0] PRINT THE 110
			- [x] [1.0.0] FILE NAME = LASTNAME.DC.AMOUNT.PDF
			- [x] [1.0.0] REMEBER STORAGE PATH OPTION (STORAGE ACCESS ON SENSEI.DATA)
		- [x] [1.0.0] CLEAR THE 110

- [ ] [1.3.9] SENSEI Main - Main Body FUNCTION
	- [x] [1.3.9] Add clean up function
		- [x] [1.3.9] Why there are free-floating cycle number when it was not attached to IQID?
	- [x] [1.4.0] Enable Update Inquiry in two ways
	 	- [x] [1.3.9] By import Update
		- [x] [1.4.0] By Reminder Date Expiration
	- [x] [1.4.0] Applied AGPL License

- [ ] [0.4.0] SENSEI CL - Reference Manual
	- [x] [0.4.0] Allow General Eligibility Look-up
		- [x] [0.4.0] Laydown GUI
		- [x] [0.4.0] Laydown Dictionary in forward and backward
		- [x] [0.4.0] Allow basic lookup
	- [x] [0.4.0] Allow Specific HDP LCTN Lookup
	- [ ] Allow Update from DJMS TABLE

- [ ] [ N/A ] SENSEI CE - CODING ENGINE THAT NEED TO SOON REPLACE 114 INFINITE
