# TiaExportBlocks
This console application connects to all the TIA project instances and exports the SCL functions, DBs, UDTs and user data types in an output folder that is supplied as argument

Example
TiaExportBlocks.exe C:\Test\MyExportedCode

Will create in C:\Test\MyExportedCode the following structure:

-- scl\
-- db\
-- udt\
-- tag_tables\

The tag tables are exported in xml format.

The application has been tested for following platforms:
Win 10 x64 Version 1909
Installed SW TIA V15.1 Upd2
