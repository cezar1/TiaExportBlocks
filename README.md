# TiaExportBlocks
Tia Openness Export Blocks
This application connects to the TIA portal (if there is one active) and then to the TIA project (if there is one opened). It saves the following elements:
+ SCL functions in .scl format
+ DB data blocks in .db format
+ UDT user data types in .udt format
+ PLC tag tables in .xml format

The export location must be specified as an input argument otherwise the program will not run.

By default the export is incremental: only blocks, types and tag tables whose modification date in TIA changed since the last export (or whose file changed or is missing) are exported again, and files of elements no longer in the project are removed. The state is kept in `%LOCALAPPDATA%\TiaExportBlocks`, outside the export location. Add `--full` to delete the export location and export everything, as before:

    TiaExportBlocks.exe C:\Test [--full] [--batch N]

Blocks and types are generated in batches of N per Openness call (default 50) and split back into one file per block, which is much faster than one call per block. A batch that fails or does not split into exactly its blocks is exported block by block. `--batch 1` exports block by block as before.

Additional comments:
+ If there are multiple projects in TIA portal instances opened, all the elements will be exported in same location - with overwriting based on whichever is the last element processed.
+ The application has been tested on Windows 10 Pro x64 version 1909 with TIA V15_1 UPD2
+ It is required to have .NET Framework 4.7.2 installed

The release 0.1 includes the exe file compiled in Visual Studio 2019 with the Siemens.Engineering DLL of version 15.1. This DLL can be obtained from https://support.industry.siemens.com/cs/attachments/108716692/108716692_TIA_PortalOpenness_Demo_V17.zip and it should be placed in the same folder as the exe file.

# Troubleshooting

- Make sure the user is part of the windows group 'Siemens TIA Openness'
