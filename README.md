# TiaExportBlocks
Tia Openness Export Blocks
This application connects to the TIA portal (if there is one active) and then to the TIA project (if there is one opened). It saves the following elements:
+ SCL functions in .scl format
+ DB data blocks in .db format
+ UDT user data types in .udt format
+ PLC tag tables in .xml format

The export location must be specified as an input argument otherwise the program will not run.

By default the export location is deleted and everything is exported, as before. Add `--incremental` to keep the location and export again only the blocks and types whose modification date in TIA changed since the last export (or whose file changed or is missing); tag tables and text lists are always exported. Afterwards the files of elements no longer in the project are deleted (PLC `scl`, `db`, `stl`, `udt`, `tag_tables\xml`, `hmi_tag_tables\xml`, `hmi_text_lists\xml`), so no stale files remain. The state is kept in `%LOCALAPPDATA%\TiaExportBlocks`, outside the export location.

    TiaExportBlocks.exe C:\Test [--incremental] [--batch N]

By default blocks and types are exported one Openness call per block, as before. Add `--batch N` (e.g. 50) to generate N blocks per call and split them back into one file per block, which is much faster. A batch that fails or does not split into exactly its blocks is exported block by block. Both `--incremental` and `--batch` are opt-in; without them the tool behaves exactly as before.

Additional comments:
+ If there are multiple projects in TIA portal instances opened, all the elements will be exported in same location - with overwriting based on whichever is the last element processed.
+ The application has been tested on Windows 10 Pro x64 version 1909 with TIA V15_1 UPD2
+ It is required to have .NET Framework 4.7.2 installed

The release 0.1 includes the exe file compiled in Visual Studio 2019 with the Siemens.Engineering DLL of version 15.1. This DLL can be obtained from https://support.industry.siemens.com/cs/attachments/108716692/108716692_TIA_PortalOpenness_Demo_V17.zip and it should be placed in the same folder as the exe file.

# Troubleshooting

- Make sure the user is part of the windows group 'Siemens TIA Openness'
