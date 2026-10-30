using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace TiaExportBlocks
{
    /// <summary>Options of an export run.</summary>
    public sealed class ExporterSettings
    {
        /// <summary>
        /// Keep the export location and export a block or type only when it changed (TIA ModifiedDate) or its file
        /// was changed/removed on disk. Files of elements that disappeared from the project are deleted.
        /// </summary>
        public bool Incremental;

        /// <summary>Blocks/types per GenerateSource call. 1 = one call per block (original behaviour).</summary>
        public int BatchSize = 1;

        /// <summary>Where the incremental state is kept. Null: %LOCALAPPDATA%\TiaExportBlocks.</summary>
        public string StateDirectory;

        /// <summary>Parent of the per-run scratch folder TIA writes into. Null: %TEMP%.</summary>
        public string TempDirectory;

        /// <summary>Time source for <see cref="Exporter.Benchmark"/>. Null: a Stopwatch.</summary>
        public IClock Clock;
    }

    /// <summary>Counters for one exported project.</summary>
    public sealed class ProjectSummary
    {
        public string Project;
        public int Exported;
        public int Unchanged;
        public int Failed;
        public int Removed;

        /// <summary>Time spent on this project.</summary>
        public TimeSpan Elapsed;

        public override string ToString()
        {
            return "Exported " + Exported + ", unchanged " + Unchanged + ", failed " + Failed + ", removed " + Removed
                + " in " + Benchmark.FormatDuration(Elapsed);
        }
    }

    /// <summary>The export location could not be emptied; exporting on top of it could leave stale files.</summary>
    public sealed class ExportLocationNotCleanException : IOException
    {
        public ExportLocationNotCleanException(string exportLocation, IList<string> remaining)
            : base("Could not delete " + remaining.Count + " item(s) in " + exportLocation + " (in use or access denied); nothing was exported. "
                + "Close programs using them (Explorer, editors, git tools) and run again:" + Environment.NewLine + "  "
                + string.Join(Environment.NewLine + "  ", remaining.Take(20)) + (remaining.Count > 20 ? Environment.NewLine + "  ..." : ""))
        {
            Remaining = remaining;
        }

        public IList<string> Remaining { get; private set; }
    }

    /// <summary>
    /// Walks TIA Portal instances / projects / devices and exports PLC blocks, UDTs, PLC tag tables,
    /// HMI tag tables and HMI text lists. Works on the <see cref="ITiaPortalProcess"/> abstraction so it
    /// runs against the real Openness API (SiemensAdapters.cs) or against test fakes.
    /// </summary>
    public sealed class Exporter
    {
        private readonly string _exportLocation;
        private readonly TextWriter _log;
        private readonly ExporterSettings _settings;
        private readonly List<ProjectSummary> _summaries = new List<ProjectSummary>();
        private readonly Benchmark _benchmark;
        private readonly HashSet<string> _written = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Per run
        private ExportState _state;
        private string _runTempDirectory;
        private readonly HashSet<string> _batchVerified = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _batchDisabled = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Per project
        private string _projectKey;
        private Dictionary<string, StateEntry> _previous;
        private Dictionary<string, StateEntry> _current;
        private bool _cleanupSafe;
        private ProjectSummary _summary;

        public Exporter(string exportLocation, TextWriter log) : this(exportLocation, log, null) { }

        public Exporter(string exportLocation, TextWriter log, ExporterSettings settings)
        {
            if (exportLocation == null) throw new ArgumentNullException("exportLocation");
            if (log == null) throw new ArgumentNullException("log");
            _exportLocation = exportLocation;
            _log = log;
            _settings = settings ?? new ExporterSettings();
            _benchmark = new Benchmark(_settings.Clock ?? new StopwatchClock());
            if (_settings.BatchSize < 1) throw new ArgumentOutOfRangeException("settings", "BatchSize must be at least 1");
        }

        public string ExportLocation { get { return _exportLocation; } }

        /// <summary>
        /// Timing of this exporter's work: total from <see cref="PrepareExportLocation"/> (or <see cref="Run"/>)
        /// to the end of <see cref="Run"/>, and per phase.
        /// </summary>
        public Benchmark Benchmark { get { return _benchmark; } }

        /// <summary>Per-project counters of the last <see cref="Run"/>.</summary>
        public IList<ProjectSummary> Summaries { get { return _summaries; } }

        public string StateFile
        {
            get { return ExportState.FileFor(_settings.StateDirectory ?? ExportState.DefaultDirectory, _exportLocation); }
        }

        private bool Incremental { get { return _settings.Incremental; } }

        /// <summary>
        /// Full export: deletes everything inside the export location (read-only and hidden files included, locked
        /// files retried) and checks that it is really empty; throws <see cref="ExportLocationNotCleanException"/>
        /// if anything is left, so an export never lands on top of stale files.
        /// Incremental export: keeps the previous export. Either way the folder exists afterwards.
        /// </summary>
        public void PrepareExportLocation()
        {
            _benchmark.Start();
            using (_benchmark.Measure(Benchmark.Prepare))
            {
                PrepareExportLocationCore();
            }
        }

        private void PrepareExportLocationCore()
        {
            if (Directory.Exists(_exportLocation))
            {
                if (!Incremental)
                {
                    _log.WriteLine("Deleting location " + _exportLocation + ".");
                    List<string> remaining = FileOps.WipeDirectoryContents(_exportLocation);
                    if (remaining.Count > 0)
                    {
                        throw new ExportLocationNotCleanException(_exportLocation, remaining);
                    }
                }
            }
            else
            {
                _log.WriteLine("Export location " + _exportLocation + " does not exist");
            }
            FileOps.EnsureDirectory(_exportLocation);
        }

        /// <summary>Files written by the last <see cref="Run"/> (unchanged files skipped by an incremental run are not included).</summary>
        public ICollection<string> WrittenFiles { get { return _written; } }

        /// <summary>
        /// After a run: removes anything in the export location that this run did not write (e.g. created by
        /// another program during the export), plus folders left empty. The export location then holds
        /// exactly this run's output.
        /// </summary>
        public int RemoveFilesNotWrittenThisRun()
        {
            int removed = 0;
            List<string> entries = FileOps.ListEntries(_exportLocation);
            foreach (string entry in entries.Where(e => !Directory.Exists(e) || (File.GetAttributes(e) & FileAttributes.ReparsePoint) != 0))
            {
                if (_written.Contains(Path.GetFullPath(entry))) continue;
                try
                {
                    var attributes = File.GetAttributes(entry);
                    File.SetAttributes(entry, attributes & ~(FileAttributes.ReadOnly | FileAttributes.Hidden | FileAttributes.System));
                    if ((attributes & FileAttributes.Directory) != 0) Directory.Delete(entry, false); else File.Delete(entry);
                    _log.WriteLine("Removed " + entry + " (not written by this export)");
                    removed++;
                }
                catch (Exception exc)
                {
                    _log.WriteLine("WARNING: could not remove " + entry + " (not written by this export): " + exc.Message);
                }
            }
            foreach (string dir in entries.Where(Directory.Exists).OrderByDescending(d => d.Length))
            {
                try
                {
                    if (Directory.GetFileSystemEntries(dir).Length == 0) Directory.Delete(dir, false);
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
            return removed;
        }

        /// <summary>Exports every open project of every TIA Portal instance.</summary>
        public void Run(IEnumerable<ITiaPortalProcess> processes)
        {
            _benchmark.Start();
            try
            {
                RunCore(processes);
            }
            finally
            {
                _benchmark.Stop();
            }
        }

        private void RunCore(IEnumerable<ITiaPortalProcess> processes)
        {
            _summaries.Clear();
            _written.Clear();
            _state = Incremental ? ExportState.Load(StateFile, _log) : new ExportState();
            _runTempDirectory = Path.Combine(_settings.TempDirectory ?? Path.GetTempPath(), "TiaExportBlocks_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_runTempDirectory);
            try
            {
                RunProcesses(processes);
                if (!Incremental)
                {
                    // Full export: the folder must hold exactly what this run wrote.
                    using (_benchmark.Measure(Benchmark.Cleanup))
                    {
                        RemoveFilesNotWrittenThisRun();
                    }
                }
            }
            finally
            {
                using (_benchmark.Measure(Benchmark.Cleanup))
                {
                    SaveState();
                }
                try { FileOps.RemoveDirectory(_runTempDirectory); }
                catch (Exception exc) { _log.WriteLine("Could not remove temporary folder " + _runTempDirectory + ": " + exc.Message); }
            }
        }

        private void RunProcesses(IEnumerable<ITiaPortalProcess> processes)
        {
            _log.WriteLine("Enumerating TIA processes..");
            foreach (ITiaPortalProcess process in processes)
            {
                _log.WriteLine("Process ID " + process.Id);
                _log.WriteLine("Project PATH " + process.ProjectPath);
                ITiaPortal tiaPortal;
                using (_benchmark.Measure(Benchmark.Attach))
                {
                    tiaPortal = process.Attach();
                }
                if (tiaPortal == null)
                {
                    _log.WriteLine("Failed to attach to TIA Portal.");
                    return;
                }
                IList<ITiaProject> projects = tiaPortal.Projects;
                if (projects == null || projects.Count == 0)
                {
                    _log.WriteLine("No projects found in TIA Portal.");
                    continue;
                }
                foreach (ITiaProject project in projects)
                {
                    ExportProject(project);
                }
            }
        }

        private void SaveState()
        {
            try
            {
                _state.Save(StateFile);
            }
            catch (Exception exc)
            {
                _log.WriteLine("Could not save incremental state to " + StateFile + ": " + exc.Message + " (next incremental export will export everything)");
            }
        }

        public void ExportProject(ITiaProject project)
        {
            _log.WriteLine("Handling project " + project.Name);
            _projectKey = string.IsNullOrEmpty(project.Path) ? project.Name : project.Path;
            _previous = _state.GetProject(_projectKey);
            _current = new Dictionary<string, StateEntry>(StringComparer.OrdinalIgnoreCase);
            _cleanupSafe = true;
            _summary = new ProjectSummary { Project = project.Name };
            TimeSpan projectStart = _benchmark.Total;

            foreach (ITiaDevice device in project.Devices)
            {
                ExportDevice(device);
            }

            if (Incremental)
            {
                if (_cleanupSafe)
                {
                    using (_benchmark.Measure(Benchmark.Cleanup))
                    {
                        RemoveStaleFiles();
                    }
                }
                else
                {
                    // Something could not be read from TIA: its files are unknown this run, so keep everything we had.
                    _log.WriteLine("Not removing any files: part of the project could not be read");
                    foreach (var entry in _previous)
                    {
                        if (!_current.ContainsKey(entry.Key)) _current[entry.Key] = entry.Value;
                    }
                }
            }
            _state.SetProject(_projectKey, _current);
            _summary.Elapsed = _benchmark.Total - projectStart;
            _summaries.Add(_summary);
            _log.WriteLine(_summary.ToString());
        }

        public void ExportDevice(ITiaDevice device)
        {
            _log.WriteLine("Handling device " + device.Name + " of type [" + device.TypeIdentifier + "]");
            bool isPlcDevice = ExportRules.IsS71500(device.TypeIdentifier);
            foreach (ITiaDeviceItem deviceItem in device.DeviceItems)
            {
                try
                {
                    if (isPlcDevice)
                        ExportPlcDeviceItem(deviceItem);
                    else
                        ExportHmiDeviceItem(deviceItem);
                }
                catch (Exception exc)
                {
                    // One broken device item must not stop the export of everything else.
                    _cleanupSafe = false;
                    _log.WriteLine("Error handling device item [" + deviceItem.Name + "]: " + exc);
                }
            }
        }

        private void ExportPlcDeviceItem(ITiaDeviceItem deviceItem)
        {
            _log.WriteLine("Handling device item " + deviceItem.Name + " of type " + deviceItem.TypeIdentifier);
            if (!ExportRules.IsPlcDeviceItem(deviceItem.Name)) return;

            _log.WriteLine("Handling PLC device item");
            ITiaSoftware software = deviceItem.GetSoftware();
            if (software == null) return;
            IPlcSoftware plcSoftware = software as IPlcSoftware;
            if (plcSoftware == null)
            {
                _log.WriteLine("Device item [" + deviceItem.Name + "] has no PLC software, skipped");
                return;
            }
            ExportPlc(plcSoftware, deviceItem.Name);
        }

        private void ExportHmiDeviceItem(ITiaDeviceItem deviceItem)
        {
            _log.WriteLine("Handling device item [" + deviceItem.Name + "] of type [" + deviceItem.TypeIdentifier + "]");
            _log.WriteLine("Handling HMI device item");
            ITiaSoftware software = deviceItem.GetSoftware();
            if (software == null)
            {
                _log.WriteLine("SW Container is null");
                return;
            }
            IHmiTarget hmiTarget = software as IHmiTarget;
            if (hmiTarget == null)
            {
                // e.g. the CPU of an S7-1200 / S7-300 station: PLC software, but not an S7-1500 device.
                _log.WriteLine("Device item [" + deviceItem.Name + "] software is not an HMI target, skipped");
                return;
            }
            ExportHmi(hmiTarget, deviceItem.Name);
        }

        // ---------------- PLC ----------------

        private sealed class Pending<T>
        {
            public T Item;
            public string Name;
            public string Extension;
            public string TargetPath;
            public long Stamp;
        }

        public void ExportPlc(IPlcSoftware software, string rootPrefix)
        {
            _log.WriteLine(software.Name);

            // Tag tables first: they are always exported, and a change in them (e.g. a renamed tag)
            // changes the source text of blocks whose own ModifiedDate stays the same.
            var changedTagTables = new List<string>();
            ExportPlcTagTables(software.TagTableGroup, rootPrefix, changedTagTables);

            var types = CollectTypes(software, rootPrefix);
            var blocks = new List<Pending<IPlcBlock>>();
            CollectBlocks(software.BlockGroup, string.Empty, rootPrefix, blocks);

            bool exportAll = !Incremental;
            if (Incremental)
            {
                string plcFolder = PlcFolder(rootPrefix);
                var previousNames = new HashSet<string>(_previous.Keys.Where(k => IsUnder(k, plcFolder)), StringComparer.OrdinalIgnoreCase);
                var currentNames = new HashSet<string>(_current.Keys.Where(k => IsUnder(k, plcFolder)), StringComparer.OrdinalIgnoreCase);
                currentNames.UnionWith(types.Select(t => t.TargetPath));
                currentNames.UnionWith(blocks.Select(b => b.TargetPath));
                if (!previousNames.SetEquals(currentNames))
                {
                    var added = currentNames.Except(previousNames).Select(Path.GetFileName).ToList();
                    var removed = previousNames.Except(currentNames).Select(Path.GetFileName).ToList();
                    _log.WriteLine("Incremental: blocks, types or tag tables of " + rootPrefix + " were added, removed or renamed since the last export"
                        + Describe(" (new: ", added) + Describe(" (gone: ", removed) + ": exporting all its blocks and types");
                    exportAll = true;
                }
                else if (changedTagTables.Count > 0)
                {
                    _log.WriteLine("Incremental: tag tables of " + rootPrefix + " changed since the last export" + Describe(" (", changedTagTables)
                        + ": exporting all its blocks and types (tag names in their sources may have changed)");
                    exportAll = true;
                }
            }

            int candidates = types.Count + blocks.Count;
            int unchangedBefore = _summary.Unchanged;
            var typesToExport = types.Where(t => exportAll || NeedsExport(t)).ToList();
            var blocksToExport = blocks.GroupBy(b => b.Extension).Select(g => g.Where(b => exportAll || NeedsExport(b)).ToList()).ToList();
            if (Incremental)
            {
                int toExport = typesToExport.Count + blocksToExport.Sum(g => g.Count);
                _log.WriteLine("Incremental: " + rootPrefix + ": " + toExport + " of " + candidates + " blocks/types to export, "
                    + (_summary.Unchanged - unchangedBefore) + " unchanged");
                int noStamp = types.Count(t => t.Stamp == StateEntry.NoStamp) + blocks.Count(b => b.Stamp == StateEntry.NoStamp);
                if (noStamp > 0)
                {
                    _log.WriteLine("Incremental: " + noStamp + " blocks/types of " + rootPrefix + " have no readable ModifiedDate and are always exported");
                }
            }

            ExportSources(typesToExport, (list, file) => software.GenerateSource(list, file));
            foreach (var group in blocksToExport)
            {
                ExportSources(group, (list, file) => software.GenerateSource(list, file));
            }
        }

        private static string Describe(string prefix, List<string> names)
        {
            if (names.Count == 0) return "";
            return prefix + string.Join(", ", names.Take(5)) + (names.Count > 5 ? ", +" + (names.Count - 5) + " more" : "") + ")";
        }

        private string PlcFolder(string rootPrefix)
        {
            return Path.GetFullPath(_exportLocation + @"\" + ExportPaths.MakeValidFileName(rootPrefix)) + @"\";
        }

        private static bool IsUnder(string path, string folderWithSeparator)
        {
            return path.StartsWith(folderWithSeparator, StringComparison.OrdinalIgnoreCase);
        }

        private List<Pending<IPlcType>> CollectTypes(IPlcSoftware software, string rootPrefix)
        {
            var result = new List<Pending<IPlcType>>();
            try
            {
                foreach (IPlcType plcType in software.Types)
                {
                    _log.WriteLine("Handling type at prefix " + rootPrefix + " named " + plcType.Name);
                    result.Add(new Pending<IPlcType>
                    {
                        Item = plcType,
                        Name = plcType.Name,
                        Extension = ".udt",
                        TargetPath = Path.GetFullPath(ExportPaths.TypeFile(_exportLocation, rootPrefix, plcType.Name)),
                        Stamp = Stamp(plcType.ModifiedDate),
                    });
                }
            }
            catch (Exception exc)
            {
                _cleanupSafe = false;
                _log.WriteLine("Error reading PLC data types of " + rootPrefix + ": " + exc.Message);
            }
            return result;
        }

        private void CollectBlocks(IPlcBlockGroup blockGroup, string currentPath, string rootPrefix, List<Pending<IPlcBlock>> result)
        {
            string newPath = ExportPaths.GroupPath(currentPath, blockGroup.Name);
            try
            {
                _log.WriteLine("Handling block group: " + blockGroup.Name);
                foreach (IPlcBlock block in blockGroup.Blocks)
                {
                    string language = block.ProgrammingLanguage;
                    string extension;
                    string folderPrefix;
                    if (!ExportRules.TryGetBlockFormat(language, out extension, out folderPrefix))
                    {
                        _log.WriteLine(block.Name + " programming language " + language + " not supported");
                        continue;
                    }
                    if (block.IsKnowHowProtected)
                    {
                        _log.WriteLine(block.Name + " is know-how protected: TIA cannot generate its source, skipped");
                        continue;
                    }
                    result.Add(new Pending<IPlcBlock>
                    {
                        Item = block,
                        Name = block.Name,
                        Extension = extension,
                        TargetPath = Path.GetFullPath(ExportPaths.BlockFile(_exportLocation, rootPrefix, folderPrefix, newPath, block.Name, extension)),
                        Stamp = Stamp(block.ModifiedDate),
                    });
                }
                foreach (IPlcBlockGroup nestedGroup in blockGroup.Groups)
                {
                    CollectBlocks(nestedGroup, newPath, rootPrefix, result);
                }
            }
            catch (Exception ex)
            {
                _cleanupSafe = false;
                _log.WriteLine("Error handling block group: " + blockGroup.Name + " - " + ex.Message);
            }
        }

        private static long Stamp(DateTime? modified)
        {
            return modified.HasValue ? modified.Value.ToUniversalTime().Ticks : StateEntry.NoStamp;
        }

        /// <summary>
        /// Incremental check: unchanged when TIA reports the same ModifiedDate as at the last successful export
        /// and the file on disk still has the content written then.
        /// </summary>
        private bool NeedsExport<T>(Pending<T> item)
        {
            StateEntry previous;
            if (item.Stamp != StateEntry.NoStamp
                && _previous.TryGetValue(item.TargetPath, out previous)
                && previous.Stamp == item.Stamp
                && previous.Hash != null
                && previous.Hash == FileOps.Sha256(item.TargetPath))
            {
                _current[item.TargetPath] = previous;
                _summary.Unchanged++;
                return false;
            }
            return true;
        }

        private enum BatchResult
        {
            /// <summary>All blocks of the batch were exported.</summary>
            Done,

            /// <summary>TIA refused the call, usually because of one block in it (know-how protected, not pure SCL/STL).</summary>
            CallFailed,

            /// <summary>TIA generated the batch but its output can't be used as is.</summary>
            NotUsable,
        }

        private void ExportSources<T>(List<Pending<T>> items, Action<IList<T>, FileInfo> generate)
        {
            if (items.Count == 0) return;
            for (int start = 0; start < items.Count; start += _settings.BatchSize)
            {
                ExportChunk(items.Skip(start).Take(_settings.BatchSize).ToList(), generate);
            }
        }

        /// <summary>
        /// Exports a chunk in one call. If TIA refuses the call, the chunk is split in halves and each half is tried
        /// again, so one bad block costs about 2·log2(N) calls instead of N single-block calls, and the other blocks
        /// still get batched. Unusable output (split or verification mismatch) falls back to one call per block.
        /// </summary>
        private void ExportChunk<T>(List<Pending<T>> chunk, Action<IList<T>, FileInfo> generate)
        {
            string extension = chunk[0].Extension;
            if (chunk.Count > 1 && !_batchDisabled.Contains(extension))
            {
                BatchResult result = TryExportBatch(chunk, extension, generate);
                if (result == BatchResult.Done) return;
                if (result == BatchResult.CallFailed)
                {
                    int half = chunk.Count / 2;
                    ExportChunk(chunk.Take(half).ToList(), generate);
                    ExportChunk(chunk.Skip(half).ToList(), generate);
                    return;
                }
            }
            foreach (var item in chunk)
            {
                ExportSingle(item, generate);
            }
        }

        /// <summary>First line of an exception message, for one-line log entries.</summary>
        private static string FirstLine(Exception exc)
        {
            string message = (exc.Message ?? "").Trim();
            int newline = message.IndexOfAny(new[] { '\r', '\n' });
            return newline < 0 ? message : message.Substring(0, newline).Trim();
        }

        private void ExportSingle<T>(Pending<T> item, Action<IList<T>, FileInfo> generate)
        {
            string temp = Path.Combine(_runTempDirectory, ExportPaths.MakeValidFileName(item.Name) + item.Extension);
            try
            {
                FileOps.DeleteIfExists(temp);
                _log.WriteLine("Generating [" + item.Name + "] to " + item.TargetPath);
                using (_benchmark.Measure(Benchmark.GenerateSource, 1))
                {
                    generate(new List<T> { item.Item }, new FileInfo(temp));
                }
                FileOps.MoveReplacing(temp, item.TargetPath);
                RecordSuccess(item.TargetPath, item.Stamp);
            }
            catch (Exception exc)
            {
                // TIA's message says why (know-how protected, not pure SCL/STL, inconsistent); its stack trace adds nothing.
                _log.WriteLine("Could not generate [" + item.Name + "]: " + (exc.Message ?? "").Trim());
                TryDelete(temp);
                RecordFailure(item.TargetPath);
            }
        }

        /// <summary>
        /// Generates several blocks in one Openness call and splits the result per block. Returns false (caller
        /// falls back to one call per block) when the call fails or the output can't be split into exactly these
        /// blocks. The first batch per file type is checked against a single-block export of its first block; if
        /// the bytes differ, batching is switched off for that file type for the rest of the run.
        /// </summary>
        private BatchResult TryExportBatch<T>(List<Pending<T>> chunk, string extension, Action<IList<T>, FileInfo> generate)
        {
            string batchFile = Path.Combine(_runTempDirectory, "batch_" + Guid.NewGuid().ToString("N") + extension);
            try
            {
                _log.WriteLine("Generating " + chunk.Count + " " + extension + " sources in one call");
                try
                {
                    using (_benchmark.Measure(Benchmark.GenerateSource, chunk.Count))
                    {
                        generate(chunk.Select(c => c.Item).ToList(), new FileInfo(batchFile));
                    }
                }
                catch (Exception exc)
                {
                    _log.WriteLine("Batch of " + chunk.Count + " " + extension + " sources failed (" + FirstLine(exc) + "): "
                        + (chunk.Count > 1 ? "splitting it to find the block(s) TIA refuses" : ""));
                    return BatchResult.CallFailed;
                }
                bool hasBom;
                string text = FileOps.ReadUtf8(batchFile, out hasBom);
                List<SourcePart> parts = SourceSplitter.Split(text);
                if (parts == null || parts.Count != chunk.Count
                    || !chunk.All(c => parts.Any(p => p.Name == c.Name)))
                {
                    _log.WriteLine("Batch output did not split into its " + chunk.Count + " blocks: exporting them one by one");
                    return BatchResult.NotUsable;
                }

                if (!_batchVerified.Contains(extension) && !VerifyBatchMatchesSingle(chunk[0], parts, hasBom, generate))
                {
                    _batchDisabled.Add(extension);
                    _log.WriteLine("Batch output for " + extension + " differs from single-block output: batching switched off for " + extension + " in this run");
                    return BatchResult.NotUsable;
                }
                _batchVerified.Add(extension);

                foreach (var item in chunk)
                {
                    string part = parts.First(p => p.Name == item.Name).Text;
                    string temp = Path.Combine(_runTempDirectory, ExportPaths.MakeValidFileName(item.Name) + item.Extension);
                    File.WriteAllBytes(temp, FileOps.EncodeUtf8(part, hasBom));
                    FileOps.MoveReplacing(temp, item.TargetPath);
                    RecordSuccess(item.TargetPath, item.Stamp);
                }
                return BatchResult.Done;
            }
            catch (Exception exc)
            {
                _log.WriteLine("Batch export failed (" + FirstLine(exc) + "): exporting its blocks one by one");
                return BatchResult.NotUsable;
            }
            finally
            {
                TryDelete(batchFile);
            }
        }

        private bool VerifyBatchMatchesSingle<T>(Pending<T> first, List<SourcePart> parts, bool hasBom, Action<IList<T>, FileInfo> generate)
        {
            string single = Path.Combine(_runTempDirectory, "verify_" + Guid.NewGuid().ToString("N") + first.Extension);
            try
            {
                using (_benchmark.Measure(Benchmark.BatchVerification, 1))
                {
                    generate(new List<T> { first.Item }, new FileInfo(single));
                }
                byte[] expected = File.ReadAllBytes(single);
                byte[] fromBatch = FileOps.EncodeUtf8(parts.First(p => p.Name == first.Name).Text, hasBom);
                return expected.SequenceEqual(fromBatch);
            }
            finally
            {
                TryDelete(single);
            }
        }

        private void RecordSuccess(string targetPath, long stamp)
        {
            _current[targetPath] = new StateEntry(stamp, FileOps.Sha256(targetPath));
            _written.Add(Path.GetFullPath(targetPath));
            _summary.Exported++;
        }

        /// <summary>
        /// The export failed. An older export of the element, if any, stays on disk (it is still the element's
        /// file and must not be cleaned up) but is recorded without a stamp so it is retried next time.
        /// </summary>
        private void RecordFailure(string targetPath)
        {
            _summary.Failed++;
            if (File.Exists(targetPath))
            {
                _current[targetPath] = new StateEntry(StateEntry.NoStamp, FileOps.Sha256(targetPath));
                _log.WriteLine("Kept the previous export " + targetPath + "; it may be outdated and is retried on the next run");
            }
        }

        /// <summary>Exports all PLC tag tables. Returns true if any of them is new, failed, or differs from the last export.</summary>
        private void ExportPlcTagTables(IPlcTagTableGroup group, string rootPrefix, List<string> changed)
        {
            foreach (IExportableItem table in group.TagTables)
            {
                if (ExportXml(table, ExportPaths.PlcTagTableFile(_exportLocation, rootPrefix, table.Name)))
                {
                    changed.Add(table.Name);
                }
            }
            foreach (IPlcTagTableGroup userGroup in group.Groups)
            {
                ExportPlcTagTables(userGroup, rootPrefix, changed);
            }
        }

        // ---------------- HMI ----------------

        public void ExportHmi(IHmiTarget hmiTarget, string hmiPrefix)
        {
            IHmiTagFolder systemFolder = hmiTarget.TagFolder;
            // First the tables in the user folders, then the tables in the system folder.
            foreach (IHmiTagFolder userFolder in systemFolder.Folders)
            {
                ExportHmiUserFolderDeep(userFolder, hmiPrefix);
            }
            ExportHmiTagTables(systemFolder, hmiPrefix);
            ExportTextLists(hmiTarget, hmiPrefix);
        }

        private void ExportHmiUserFolderDeep(IHmiTagFolder folder, string hmiPrefix)
        {
            foreach (IHmiTagFolder subFolder in folder.Folders)
            {
                ExportHmiUserFolderDeep(subFolder, hmiPrefix);
            }
            ExportHmiTagTables(folder, hmiPrefix);
        }

        private void ExportHmiTagTables(IHmiTagFolder folder, string hmiPrefix)
        {
            foreach (IExportableItem table in folder.TagTables)
            {
                ExportXml(table, ExportPaths.HmiTagTableFile(_exportLocation, hmiPrefix, table.Name));
            }
        }

        private void ExportTextLists(IHmiTarget hmiTarget, string hmiPrefix)
        {
            foreach (IExportableItem textList in hmiTarget.TextLists)
            {
                ExportXml(textList, ExportPaths.HmiTextListFile(_exportLocation, hmiPrefix, textList.Name));
            }
        }

        /// <summary>
        /// Exports an XML element (always; XML exports have no stamp) into the scratch folder and compares it with
        /// the file already in place, ignoring the &lt;DocumentInfo&gt; header (export time). An unchanged file is
        /// left untouched, so its date and content stay stable for git and sync tools.
        /// Returns true if the content is new, changed, failed, or differs from the last export.
        /// </summary>
        private bool ExportXml(IExportableItem item, string path)
        {
            string target = Path.GetFullPath(path);
            string temp = ExportPaths.PlcTagTableTempFile(_runTempDirectory, item.Name);
            _log.WriteLine(item.Name);
            try
            {
                FileOps.DeleteIfExists(temp);
                using (_benchmark.Measure(Benchmark.ExportXml, 1))
                {
                    item.Export(new FileInfo(temp));
                }
                string hash = FileOps.XmlContentHash(temp);
                if (File.Exists(target) && FileOps.XmlContentHash(target) == hash)
                {
                    FileOps.DeleteIfExists(temp);
                    _current[target] = new StateEntry(StateEntry.NoStamp, hash);
                    _written.Add(target);
                    _summary.Unchanged++;
                }
                else
                {
                    _log.WriteLine(item.Name + " to " + target);
                    FileOps.MoveReplacing(temp, target);
                    _current[target] = new StateEntry(StateEntry.NoStamp, hash);
                    _written.Add(target);
                    _summary.Exported++;
                }
                StateEntry previous;
                return !_previous.TryGetValue(target, out previous) || previous.Hash != hash;
            }
            catch (Exception exc)
            {
                _log.WriteLine(exc.ToString());
                TryDelete(temp);
                RecordFailure(target);
                return true;
            }
        }

        // ---------------- cleanup ----------------

        /// <summary>
        /// Deletes files this tool wrote for this project on an earlier run whose element is gone (deleted or
        /// renamed block, type, tag table, text list, or a removed/renamed PLC or HMI). Files the tool did not
        /// write, and files another project in the same export location also wrote, are never touched.
        /// </summary>
        private void RemoveStaleFiles()
        {
            string root = Path.GetFullPath(_exportLocation).TrimEnd('\\') + @"\";
            foreach (string file in _previous.Keys.Where(k => !_current.ContainsKey(k)).OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                if (!IsUnder(file, root)) continue;
                if (_state.IsClaimedByOtherProject(file, _projectKey)) continue;
                try
                {
                    if (FileOps.DeleteIfExists(file))
                    {
                        _log.WriteLine("Removing " + file + " (no longer in the project)");
                        _summary.Removed++;
                    }
                    FileOps.RemoveEmptyParents(file, _exportLocation);
                }
                catch (Exception exc)
                {
                    _log.WriteLine("Could not remove " + file + ": " + exc.Message);
                    _current[file] = _previous[file];
                }
            }
        }

        private static void TryDelete(string path)
        {
            try { FileOps.DeleteIfExists(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
