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

The release 0.1 includes the exe file compiled in Visual Studio 2019 with the Siemens.Engineering DLL of version 15.1. This DLL can be obtained from https://support.industry.siemens.com/cs/attachments/108716692/108716692_TIA_PortalOpenness_Demo_V17.zip and it should be placed in the same folder as the exe file.

# Troubleshooting

- Make sure the user is part of the windows group 'Siemens TIA Openness'

# CI build

GitHub Actions (`.github/workflows/build.yml`) builds the exe on a Windows runner on every push/PR, uploads it as a workflow artifact, and attaches a zip to the GitHub Release when a `v*` tag is pushed.

Because `Siemens.Engineering.dll` is not in this repo, the workflow downloads it from the private repo `cezar1/siemens_tia_engineering_dlls` using a fine-grained token stored in the secret `TIA_DLLS_TOKEN` (read-only "Contents" access to that repo). The DLL for each TIA version goes in its own folder there (`v15_1/Siemens.Engineering.dll`, `v16/...`) and run the workflow manually with that version. The DLL is used only for compiling and is not included in the published zip.

# Unit tests

`tests/TiaExportBlocks.Tests` (xUnit, .NET Framework 4.8) covers everything that does not need TIA Portal: export folder layout and file names, file name sanitizing, which block languages are exported, PLC/HMI device classification, argument handling and the file replace steps around each export. That logic lives in `ExportLogic.cs`; `Program.cs` keeps only the calls into the Siemens Openness API, which can only be exercised against a running TIA Portal.

Run locally with `dotnet test tests/TiaExportBlocks.Tests` (or Test Explorer in Visual Studio). CI runs them on every push and PR before the exe is built.
