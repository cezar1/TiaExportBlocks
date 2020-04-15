# TiaExportBlocks
Tia Openness Export Blocks
This application connects to the TIA portal (if there is one active) and then to the TIA project (if there is one opened). It saves the following elements:
+ SCL functions in .scl format
+ DB data blocks in .db format
+ UDT user data types in .udt format
+ PLC tag tables in .xml format

The export location must be specified as an input argument otherwise the program will not run.

Additional comments:
+ If there are multiple projects in TIA portal instances opened, all the elements will be exported in same location - with overwriting based on whichever is the last element processed.
+ The application has been tested on Windows 10 Pro x64 version 1909 with TIA V15_1 UPD2
+ It is required to have .NET Framework 4.7.2 installed

The release 0.1 includes the exe file compiled in Visual Studio 2019 with the Siemens.Engineering DLL of version 15.1. This DLL can be obtained from https://support.industry.siemens.com/cs/document/108716692/tia-portal-openness%3A-introduction-and-demo-application?dti=0&lc=en-US (Visual Studio project "TiaPortalOpennessDemo" (660,4 KB) https://support.industry.siemens.com/cs/attachments/108716692/108716692_TiaPortalOpennessDemo_V15_1.zip ) and it should be placed in the same folder as the exe file.
