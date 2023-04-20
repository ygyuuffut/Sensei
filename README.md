# SENSEI - A Clustered Applets for Excel
![License](https://img.shields.io/badge/License-AGPL%203-BB8FCE?style=flat-square)
![Version](https://img.shields.io/badge/Version-1.7.0-76D754?style=flat-square)
![Version Type](https://img.shields.io/badge/Type-Release--01-16A085?style=flat-square)
![Language](https://img.shields.io/badge/Language-Virtual%20Basic-EB984E?style=flat-square)
![Size](https://img.shields.io/github/languages/code-size/ygyuuffut/Sensei?style=flat-square)

A very much scatter-mind driven programming that develops as demands emerges

## Table of Contents
- [SENSEI - A Clustered Applets for Excel](#sensei---a-clustered-applets-for-excel)
	- [Table of Contents](#table-of-contents)
	- [Supported Platforms](#supported-platforms)
	- [Structure Guidances](#structure-guidances)
	- [Installations](#installations)
		- [Manual Composition](#manual-composition)
		- [Alternative Method](#alternative-method)
	- [Update History](#update-history)
		- [Sensei Main UI](#sensei-main--)
		- [Sensei Link - Document Network](#sensei-link---document-network-)
		- [Sensei FD - Form Distiller](#sensei-fd---form-distiller-)
		- [Sensei DS - Deployee Scantron](#sensei-ds---deployee-scantron-)
		- [Sensei REJC - Rejection Report](#sensei-rejc---reject-report-cl)
		- [Sensei CL - Common Library](#sensei-cl---common-library-)
		- [Sensei CE - Coding Engine](#sensei-ce---coding-engine-)


## Supported Platforms
**Microsoft Office** - 2016 (64 bit) or later, recommended for Office 365 (64 bit) on Windows 7 or later

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
> 2. Then Migrate the Data as needed (if the original is over written, then it is done for)


## Update History

### Sensei Main  ![UI](https://img.shields.io/badge/1.7.0-Release-76D754)

<details><summary>SENSEI 1.7 - Reporter</summary>
<p>

![](https://img.shields.io/badge/1.7.0-424949?style=flat-square)
- Added Clarification to Country lookup
- Added 2424 JPBB Input Field for Console Input
- Added Legacy 114 Port Option 'Soon to Remove'
- Fixed Scantron not printing correctly
- Fixed Incomplete Deletion (need to verify)
- Fixed Leakage when editing after resort
- Fixed Errorneous Erase after resort
- Fixed Shadow entry when archiving while stage 5 entry is in display
- Patch Leakeage on Append Page sort
- Patch Record Page Entry Miccounting
- Patch AutoScroll infinite loop on MISC entries
- Patch Inadequate Update in Form 110 Indicator
- Modified Record Page percentage display method
- Modified Layout in Edit pane
- Modified New Entry Appending information
- Accessibility Updates
- The Email Contacting with Logging
- Font Obtainer
- Updated User Agreement
- Alternative Method to copy SSN
- Added Auto Logging for today's date in description
- Amended function for travel in edit tab - Auto "Profile"
- Application title modification function
- Isolated Export on REJC
- Experiment Entry Flagging
- Improved readability on Amending Description Box
- Display Clarity improvements on form display
- Fixed inaccurate date format in Form 110
- EXPANDED Capacity to 500
- Fixed Notification Leak
- Complete 117 Runner
- Imporved readibility on Amending Description Box
- Display Clearity improvements on form display
- Fixed inaccurate date format in Form 110
- EXPANDED Capacity to 500
- Fixed Notification Leak
- Integrate Rejection Report


</p>
</details>

<details><summary>SENSEI 1.6 - Historian</summary>
<p>

![](https://img.shields.io/badge/1.6.0-424949?style=flat-square)
- Added limited search based on stage
- Amended Limted search into compound search
- Changed how MISC type is recorded
- Modified Edit Panel Commentary display method
- Repaired where Edit Panel Loads SSID incorrectly (0 header)
- Repaired Capacity Display
- Repaired Config Incorrect Loading
- Repaired Incomplete Erase
- Repaired Formula Self-Fixuture issue
- Repaired Document Export Leakage issue
- Add Dash Board to view historical Performances when disgnated a time frame (or all time if blank)
- Separate CSP and CMS Benchmarks on dashboards
- Allow Re-count based on Current Archive
- Enable other pages tap into auto save data

</p>
</details>

<details><summary>SENSEI 1.5 - Scantrons</summary>
<p>

![](https://img.shields.io/badge/1.5.7-424949?style=flat-square)
- Additional Support to display Private information upon request
- Added Support to Misc. entry type

![](https://img.shields.io/badge/1.5.6-424949?style=flat-square)
- Depreciated Reset Function and Replaced with Save and Quit

![](https://img.shields.io/badge/1.5.5-424949?style=flat-square)
- Introduced primitive auto-save function
- Nuke Module is updated to accomodate latest configurations
- Introduced primitive blocking for trust warning upon saving

![](https://img.shields.io/badge/1.5.4-424949?style=flat-square)
- Eliminated Ambiguity in Form tab, altered naming convention for accessibility

![](https://img.shields.io/badge/1.5.3-424949?style=flat-square)
- Main Holder Capacity increased to 300
- Added 1 Additional Amending Type (CMS)

![](https://img.shields.io/badge/1.5.2-424949?style=flat-square)
- Bumped due to Distiller Update
	
![](https://img.shields.io/badge/1.5.1-424949?style=flat-square)
- Optimized Scantron and enable Omit function
  
![](https://img.shields.io/badge/1.5.0-424949?style=flat-square)
- Added Scantron Function
</p>
</details>

<details><summary>SENSEI 1.4 - Distill Forms</summary>
<p>

![](https://img.shields.io/badge/1.4.1-424949?style=flat-square)
- Embedded Links
- Optimized Nuke Function
  
![](https://img.shields.io/badge/1.4.0-424949?style=flat-square)
- Dual Method Data Update
- Data Update by Reminder
- Applied AGPL v3 License
</p>
</details>

<details><summary>SENSEI 1.3 - The Library</summary>
<p>

![](https://img.shields.io/badge/1.3.9-424949?style=flat-square) 
- Data Update by Expiration
- Add clean up function
- free-floating cycle resolution
- Handle free-floating data not associated with ID
- Update Data by import
</p>
</details>

### Sensei Link - Document Network ![CL](https://img.shields.io/badge/1.0.3-Release-76D754)
<details><summary>LINK 1.0 - Bridge</summary>
<p>

![](https://img.shields.io/badge/1.0.3-424949?style=flat-square)
- Embedded Link for Quick Access

![](https://img.shields.io/badge/1.0.2-424949?style=flat-square)
- SENSEI LINK - Bridge between files
  - Link to Modified 114
  - Link to 3R Report
</p>
</details>

### Sensei FD - Form Distiller ![CL](https://img.shields.io/badge/1.2.0-Release-76D754)
<details><summary>FD 1.0 - Distilling Forms</summary>
<p>

![](https://img.shields.io/badge/1.2.0-424949?style=flat-square)
- Added Support to Form 2424

![](https://img.shields.io/badge/1.1.0-424949?style=flat-square)
- Additional Controlls for Form 110

![](https://img.shields.io/badge/1.0.0-424949?style=flat-square)
- SENSEI Form Distller introduction
- Project Form 110
  - trigger Update
  - Page Change
  - Print and Clear the Form
</p>
</details>

### Sensei DS - Data Scantron ![CL](https://img.shields.io/badge/1.2.0-Release-76D754)
<details><summary>DS 1.0 - Data Handler</summary>
<p>

![](https://img.shields.io/badge/1.2.0-424949?style=flat-square)
- Overhaul of elements for future expandibility

![](https://img.shields.io/badge/1.1.0-424949?style=flat-square)
- Changed naming convention for expandibility
- Added Total Count whenever completed data loading
- Enabled Global Printing Config

![](https://img.shields.io/badge/1.0.1-424949?style=flat-square)
- Optimized Function handling speed
- Added Omit function

![](https://img.shields.io/badge/1.0.0-424949?style=flat-square)
- SENSEI Deployee Scantron Introduction
- Full iteration and Data recognition Logic

</p>
</details>


### Sensei REJC - Reject Report ![CL](https://img.shields.io/badge/1.1.0-Release-76D754)
<details><summary>REJC 1.0 - Data Handler</summary>
<p>

![](https://img.shields.io/badge/1.1.0-424949?style=flat-square)
- Improved Stability while generating HTML based email and web report

![](https://img.shields.io/badge/1.0.0-424949?style=flat-square)
- SENSEI Rejection Report Introduction

</p>
</details>

### Sensei CL - Common Library ![CL](https://img.shields.io/badge/0.4.0-Develop-FF4545)
<details><summary>CL 0.4 - In Development</summary>
<p>

![](https://img.shields.io/badge/0.4.0-424949?style=flat-square)
- Allow General Eligibility Look-up
- Laydown GUI
- Laydown Dictionary in forward and backward
- Allow basic lookup
- Allow Specific HDP LCTN Lookup
- Allow Update from DJMS TABLE
</p>
</details>


### Sensei CE - Coding Engine ![CL](https://img.shields.io/badge/0.0.1-Un--Started-888895)
<details><summary>CE 0.0 - Un-started</summary>
<p>

![](https://img.shields.io/badge/Yet%20to%20Start-424949?style=flat-square)
- SENSEI CE - CODING ENGINE THAT NEED TO SOON REPLACE 114 INFINITE
</p>
</details>
