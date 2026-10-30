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

# Usage

    TiaExportBlocks.exe C:\Test [--incremental] [--batch N] [--benchmark [--benchmark-csv file.csv]] [--pause-at-end]

Without options: full export. The export location is emptied and everything is exported, one block per Openness call.

## --batch N

Generates N blocks (or UDTs) per `GenerateSource` call and splits the result back into one file per block, which is much faster (≈330 s → ≈50 s for 500 blocks with `--batch 50`). The split keeps the generated text byte for byte. The first batch of each file type is checked against a single-block export of its first block; if the bytes differ (e.g. another TIA version formats multi-block output differently) batching is switched off for that file type for the rest of the run. If TIA refuses a batch call, usually because one block in it is not pure SCL/STL, the batch is split in halves and retried, narrowing down to the refused block in about 2·log2(N) calls; all other blocks are still batched, and only that block is reported (`Could not generate [name]: ...`). Batching is not switched off by such errors. A batch whose output does not split cleanly into exactly its blocks is exported block by block. Know-how protected blocks are skipped before asking TIA. N is 1 to 1000.

## --incremental

Keeps the export location and generates a block or UDT again only when needed. A block is skipped only if **all** of these hold:

- TIA reports the same `ModifiedDate` as at its last successful export
- its file still has exactly the content written then (SHA-256; write times are ignored, so sync tools or git checkouts that only touch files don't cause re-exports)
- no block, UDT or tag table of that PLC was added, removed or renamed since the last export
- none of that PLC's tag tables changed (a renamed tag changes the source of blocks whose own `ModifiedDate` does not change). Tag tables are compared without their `<DocumentInfo>` header, which only holds the export time.

Tag tables and HMI tag tables/text lists are always exported, but an existing XML file is only replaced when its content changed (ignoring the export time), so unchanged files keep their date and show no diff in git. The log says per PLC how many blocks/types are exported and, when all of them are, why (which elements were added/removed, which tag tables changed). A failed export keeps the previous file and is retried on the next run. Afterwards the tool deletes the files it exported earlier whose element is gone (deleted or renamed blocks, UDTs, tag tables, text lists, removed or renamed PLCs/HMIs), and folders left empty. It never deletes files it did not write (README, `.git`, hand-written files) or files another project in the same export location also wrote. If part of a project cannot be read from TIA, nothing is deleted in that run.

State is kept per export location and per project in `%LOCALAPPDATA%\TiaExportBlocks`, outside the export location. A missing or unreadable state simply means everything is exported. Changes that alter generated sources without changing any of the above (e.g. switching the project's editing language) are not detected; run a full export occasionally.

## --pause-at-end

Waits for Enter before the program exits, also after an error, so the output stays readable when the tool is started from a shortcut, a double-click or a TIA Portal external tool entry. Ignored when input is redirected (scripts, CI).

## --benchmark

Every run prints its total time (`Execution Time`, from emptying the export location to the end of the export) and the time per project. With `--benchmark` it also prints where the time went and appends the run to a CSV history, so different settings (e.g. `--batch` sizes, full vs incremental) can be compared over time:

```
Benchmark: total 47.3 s
  Phase                                  Time  Share  Calls  Items  Avg/call
  Prepare export location              412 ms    1 %      1                412 ms
  Attach to TIA Portal                  2.1 s    4 %      1                  2.1 s
  Generate sources (blocks, UDTs)      38.0 s   80 %     12    520         3.2 s
  Batch verification                    0.9 s    2 %      4                236 ms
  Export XML (tag tables, text lists)   4.1 s    9 %     30               137 ms
  Cleanup and state                    180 ms    0 %      2                 90 ms
  Other (reading TIA, files)            1.6 s    3 %
```

"Other" is everything outside the measured Openness calls: walking the project tree in TIA, moving and hashing files. The CSV goes to `%LOCALAPPDATA%\TiaExportBlocks\benchmark.csv`, or to the file given with `--benchmark-csv` (which implies `--benchmark`). Columns: `timestamp_utc, export_location, mode, batch_size, total_ms, prepare_ms, attach_ms, generate_ms, generate_calls, generated_items, verify_ms, export_xml_ms, export_xml_calls, cleanup_ms, other_ms, exported, unchanged, failed, removed`.

# Clean export guarantee

Every full (non-incremental) run starts by deleting **everything** inside the export location (the folder itself is kept): read-only, hidden and system files included, files locked by another program retried for about a second. Links (junctions/symlinks) inside it are removed without touching what they point to. If anything still cannot be deleted, the tool stops with exit code 1 and lists the files, without exporting anything. After the export, anything in the export location that the run did not write itself is removed too. The export location therefore always holds exactly the current project state. Drive and share roots (`C:\`, `\\server\share`) are refused as export location.

Exit codes: 0 done, 1 export location could not be emptied, 2 invalid arguments, 3 unexpected error.

# Export layout

```
<export>\<PLC device item>\scl|db|stl\<block group path>\<block>.scl|.db|.awl
<export>\<PLC device item>\udt\<type>.udt
<export>\<PLC device item>\tag_tables\xml\<tag table>.xml
<export>\<HMI device item>\hmi_tag_tables\xml\<tag table>.xml
<export>\<HMI device item>\hmi_text_lists\xml\<text list>.xml
```

Names are made file-system safe (invalid characters become `_`). Only S7-1500 stations are exported as PLCs; other devices are scanned for HMI targets.

# Code layout and tests

| File | Role |
|---|---|
| `Program.cs` | Entry point: argument check, timing, wiring |
| `Exporter.cs` | The export run: walks instances → projects → devices → blocks/types/tag tables/HMI items |
| `ExportLogic.cs` | Rules, output paths, file name sanitizing, file replace steps |
| `Benchmark.cs` | Total and per-phase timing, report table, CSV history |
| `TiaModel.cs` | Interfaces for the slice of the Openness object model the exporter uses |
| `SiemensAdapters.cs` | Thin wrappers exposing the real `Siemens.Engineering` objects through those interfaces |

`tests/TiaExportBlocks.Tests` (xUnit, .NET Framework 4.8) compiles everything except `Program.cs` and `SiemensAdapters.cs`, so it needs neither the Siemens DLL nor TIA Portal. `Fakes/FakeTia.cs` simulates TIA Portal: it builds fake projects (stations, PLC software, block groups, UDTs, tag tables, HMI targets), writes realistic source/XML files on `GenerateSource`/`Export`, refuses to overwrite existing files or write into missing folders, can inject failures (know-how protected blocks, aborted generation, disposed objects, failing exports) and records every Openness call in order so tests can assert on it.

Run locally with `dotnet test tests/TiaExportBlocks.Tests` (or Test Explorer in Visual Studio). CI runs them on every push and PR before the exe is built.
