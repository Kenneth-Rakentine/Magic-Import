<img src="./krtrnsprnt.png">
**Installable Add-In**<br>
As an Add-in (.xlam)<br>
<br>
Save your workbook as .xlam (File > Save As > Excel Add-In (*.xlam)).<br>
<br>
Distribute the .xlam file to users.<br>
<br>
Users install the add-in via File > Options > Add-ins > Go... > browse<br>
<br>
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
