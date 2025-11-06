<img src="./krtrnsprnt.png">
<br>
# MAGIC IMPORT
---
<br>
VBA Add-In Macro for Financial Reports excel spreadsheet. Imports manager flash and journal by transaction code reports in .xml format to spreadsheet, selects date row by parsing report name for yyyy/mm/dd, enters data according to transaction code coloumn (sums mc+vs for single column)
<br>
<br>
## Instructions
---
**Installable Add-In**<br>
As an Add-in (.xlam)<br>
<br>
Download .xlam from this Repository<br>  
<br>
Distribute the .xlam file to users.<br>
<br>
---
<br>
You may have to "unblock" the .xlam add-in file by right clicking the file in file explorer, Properties > General tab > check "Unblock"<br>
<img src="./unblock.png">
<br>
Users install the add-in to Excel Spreadshet (_saved as .xlsm format_) via File > Options > Add-ins > Go... > browse > select downloaded .xlam file<br>
<br>
In "Developer" tab > "Macro Security", enable all macros & trust access to the VBA object model
<img src="./macro_security.png"> 
<br>
**Trust adjust Center Settings:**<br>
In Excel, go to File > Options<br>
Click Trust Center (left sidebar)<br>
Click Trust Center Settings button<br>
Go to Protected View (left sidebar)<br>
Uncheck these options:<br>
"Enable Protected View for files originating from the Internet"<br>
"Enable Protected View for files located in potentially unsafe locations"<br>
<img src="./protected.png"> <br>
**Add to Trusted Locations **(the download directory where th4e .xlam file is located):<br>
File > Options > Trust Center > Trust Center Settings<br>
Click Trusted Locations<br>
Click Add new location<br>
Browse to the folder containing your .xlam file<br>
Check "Subfolders of this location are also trusted" (if needed)<br>
Click OK<br>
<br>
---
<br>
**Naming format for Reports Scheduler**<br>
<br>
-**Manager Flash:**<br>
<br>
manager<P_FROM_DATE_YEAR><P_FROM_DATE_MONTH><P_FROM_DATE_DAY>_
<br>
-**Journal by Cashier & Transaction Code:** <br>
journal<P_FROM_DATE_YEAR><P_FROM_DATE_MONTH><P_FROM_DATE_DAY>_<br>
<br>
(they both end with  with ".xml", so ensure the underscore_ is before the ".xml" because that will separate<br> 
the 8 digit date text from the random generated numerical text that follows and make it more easily readable<br> 
while still retaining the date in the correct format for the script to parse which row to insert the data.<br> 
".xml" is auto-added to the name, outside of the custom text naming field)
**
Excel Revenue Spreadhseet Format**<br>
<img src="./format.png">
