# Finance/Accounting Excel Toolbox (FAXT)
A collection of tools developed to help Finance and Accounting professionals to collect, transform, validate, analyze, and export data. This toolbox is designed to work in a local enviornment and all data will be contained in users' local machines.
# Tool Description
## Dropdown Helper:
This tool is useful when working with workbook with data validation. It provides additional functionalities such as Search As You Type (work with substring searching), sorting, and toggles to auto show/hide the window and move cell selection to any direction after input (you can hover on each button to view the description).
![](Images/Dropdown%20Helper%20Demo.gif)
## XML Importer:
This generic XML Importer will work with XMLs with any schema so you won't need to work with the poorly-designed XML Import Wizard built in Excel.
![](Images/XML%20Importer.PNG)
## PDF Table Grabber/Tabula:
This is a direct port of the amazing open-source Tabula project. To use it, you must have [Java](https://www.java.com/download/) installed.
![](Images/Tabula%201.PNG)
![](Images/Tabula%202.PNG)
For more information, you can head over to https://github.com/tabulapdf/tabula.
## External Link Breaker:
External links can be painful to deal with as they can hide in named ranges and conditional formatting. This tool can help breaking those normally unbreakable links through native Excel "Edit Link" window. Notes: Links within charts and other special objects cannot be broken through this too.
![](Images/External%20Link%20Breaker.PNG)
# Installation
The toolbox supports auto-update upon the ribbon loading. To install it properly, you have to follow the steps below:
1. Download the latest version (.zip file) from Release page
2. Unzip it to any directory
3. Right-click on "setup.exe" - Properties
4. Open the "Digital Signatures" tab - highlight the signature - "Details"
5. In the new "Digital Signature Details" window, click "View Certificate"
6. In the new "Certificate" window, click "Install Certificate"
![](Images/Certificate%201.PNG)
(reference for step 1 - 6)
7. In the popped-up wizard, click "Next" -> Select "Place all certificates in the following store" -> click "Browse..."
8. In the new "Select Certificate Store" window, highlight "Trusted Root Certification Authorities", click "OK"
![](Images/Certificate%202.PNG)
(reference for step 7-8)
9. This will return you back to the wizard. Click "Next" then "Finish". You should see a message box saying the import was successful.
10. Double-click "setup.exe" and you should be able to install the toolbox successfully. The next time you open Excel, the toolbox will be updated to the latest version, and you should see a new ribbon tab named "FAXT"
