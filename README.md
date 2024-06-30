# Documentation

This repo contains VBA modules that help with automation of Excel spreadsheets. To be able to use the VBA modules globally across multiple spreadsheets, the modules must be imported into "Personal Macro Workbook" (Personal.xlsb). Refer to this [resource](https://support.microsoft.com/en-us/office/create-and-save-all-your-macros-in-a-single-workbook-66c97ab3-11c2-44db-b021-ae005a9bc790) for more information.

## How to import modules?
Two ways to import:
1. Manual import (refer to online [documentations](https://support.esri.com/en-us/knowledge-base/import-modules-to-vba-1462478931981-000003483) on how to do this)
2. Automated import using PowerShell script

### For automated import via PowerShell script 
Right click on [ImportVBAModules.ps1](ImportVBAModules.ps1) and Click on "Run with PowerShell". However, for powershell automation to run properly, some security settings have to be enabled. Refer to pre-requisites below.

## Pre-requisites
1. Excel Application of course
2. Enable Trust access to VBA project model (optional, only required if you are importing using PowerShell script)

### How to enable "Trust access to the VBA project object model"?

For programatic access to VBA project, security settings below have to be enabled. However, it is **recommended to disable** trust access to VBA project object model after running powershell script to prevent malicious code from accessing and rewriting VBA modules.

**Steps:**
1. Start Microsoft Excel.
2. Open a workbook.
3. Click File and then Options.
4. In the navigation pane, select Trust Center.
5. Click Trust Center Settings....
6. In the navigation pane, select Macro Settings.
7. Ensure that Trust access to the VBA project object model is checked.
8. Click OK and close Excel.

## How to run macros or VBA modules?
Go to Developer Tab (navigation pane) > Macros > Select the macro you want to run and select "Run".