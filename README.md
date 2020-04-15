# TiaExportBlocks
This console application connects to all the TIA project instances and exports the SCL functions, DBs, UDTs and user data types in an output folder that is supplied as argument

Example
TiaExportBlocks.exe C:\Test\MyExportedCode

Will create in C:\Test\MyExportedCode the following structure:

+ scl\
+ db\
+ udt\
+ tag_tables\

Additional comments/clarifications:
+ The tag tables are exported in xml format.
+ The block groups are parsed recursively.

The application has been tested for following platforms:
+ Win 10 x64 Version 1909
+ Installed SW TIA V15.1 Upd2

The release included is not bundled with the Siemens.Engineering.dll file. The dll file must be downloaded from Siemens website:
https://support.industry.siemens.com/cs/document/108716692/tia-portal-openness%3A-introduction-and-demo-application?dti=0&lc=en-US

The demo application V15.1 contains it. This file must be copied to the same location as the exe file.
