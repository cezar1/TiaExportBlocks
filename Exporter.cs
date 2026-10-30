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
            bool tagTablesChanged = ExportPlcTagTables(software.TagTableGroup, rootPrefix);

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
                    _log.WriteLine("Blocks, types or tag tables of " + rootPrefix + " were added, removed or renamed since the last export: exporting all its blocks and types");
                    exportAll = true;
                }
                else if (tagTablesChanged)
                {
                    _log.WriteLine("Tag tables of " + rootPrefix + " changed since the last export: exporting all its blocks and types");
                    exportAll = true;
                }
            }

            ExportSources(types.Where(t => exportAll || NeedsExport(t)).ToList(), (list, file) => software.GenerateSource(list, file));
            foreach (var byExtension in blocks.GroupBy(b => b.Extension))
            {
                ExportSources(byExtension.Where(b => exportAll || NeedsExport(b)).ToList(), (list, file) => software.GenerateSource(list, file));
            }
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

        private void ExportSources<T>(List<Pending<T>> items, Action<IList<T>, FileInfo> generate)
        {
            if (items.Count == 0) return;
            string extension = items[0].Extension;
            for (int start = 0; start < items.Count; start += _settings.BatchSize)
            {
                var chunk = items.Skip(start).Take(_settings.BatchSize).ToList();
                if (chunk.Count > 1 && !_batchDisabled.Contains(extension) && TryExportBatch(chunk, extension, generate))
                {
                    continue;
                }
                foreach (var item in chunk)
                {
                    ExportSingle(item, generate);
                }
            }
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
                _log.WriteLine(exc.ToString());
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
        private bool TryExportBatch<T>(List<Pending<T>> chunk, string extension, Action<IList<T>, FileInfo> generate)
        {
            string batchFile = Path.Combine(_runTempDirectory, "batch_" + Guid.NewGuid().ToString("N") + extension);
            try
            {
                _log.WriteLine("Generating " + chunk.Count + " " + extension + " sources in one call");
                using (_benchmark.Measure(Benchmark.GenerateSource, chunk.Count))
                {
                    generate(chunk.Select(c => c.Item).ToList(), new FileInfo(batchFile));
                }
                bool hasBom;
                string text = FileOps.ReadUtf8(batchFile, out hasBom);
                List<SourcePart> parts = SourceSplitter.Split(text);
                if (parts == null || parts.Count != chunk.Count
                    || !chunk.All(c => parts.Any(p => p.Name == c.Name)))
                {
                    _log.WriteLine("Batch output did not split into its " + chunk.Count + " blocks: exporting them one by one");
                    return false;
                }

                if (!_batchVerified.Contains(extension) && !VerifyBatchMatchesSingle(chunk[0], parts, hasBom, generate))
                {
                    _batchDisabled.Add(extension);
                    _log.WriteLine("Batch output for " + extension + " differs from single-block output: batching switched off for " + extension + " in this run");
                    return false;
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
                return true;
            }
            catch (Exception exc)
            {
                _log.WriteLine("Batch export failed (" + exc.Message + "): exporting its blocks one by one");
                return false;
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
        private bool ExportPlcTagTables(IPlcTagTableGroup group, string rootPrefix)
        {
            bool changed = false;
            foreach (IExportableItem table in group.TagTables)
            {
                changed |= ExportXml(table, ExportPaths.PlcTagTableFile(_exportLocation, rootPrefix, table.Name));
            }
            foreach (IPlcTagTableGroup userGroup in group.Groups)
            {
                changed |= ExportPlcTagTables(userGroup, rootPrefix);
            }
            return changed;
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

        /// <summary>Exports an XML element (always, no stamp). Returns true if the content is new, failed, or differs from the last export.</summary>
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
                _log.WriteLine(item.Name + " to " + target);
                FileOps.MoveReplacing(temp, target);
                RecordSuccess(target, StateEntry.NoStamp);
                StateEntry previous;
                return !_previous.TryGetValue(target, out previous) || previous.Hash != _current[target].Hash;
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
