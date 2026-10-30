using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TiaExportBlocks.Tests.Fakes;

namespace TiaExportBlocks.Tests
{
    /// <summary>A sandbox (export folder, state folder, scratch folder) plus a fake TIA, for multi-run export tests.</summary>
    public sealed class ExportHarness : IDisposable
    {
        public readonly string Sandbox;
        public readonly string Export;
        public readonly string StateDir;
        public readonly string TempDir;
        public readonly FakeTia Tia = new FakeTia();
        private StringWriter _log = new StringWriter();

        public ExportHarness()
        {
            Sandbox = Path.Combine(Path.GetTempPath(), "TiaHarness_" + Guid.NewGuid().ToString("N"));
            Export = Path.Combine(Sandbox, "export");
            StateDir = Path.Combine(Sandbox, "state");
            TempDir = Path.Combine(Sandbox, "tmp");
            Directory.CreateDirectory(Sandbox);
        }

        public void Dispose()
        {
            if (!Directory.Exists(Sandbox)) return;
            foreach (var f in Directory.GetFiles(Sandbox, "*", SearchOption.AllDirectories)) File.SetAttributes(f, FileAttributes.Normal);
            Directory.Delete(Sandbox, true);
        }

        /// <summary>Log of the last run only.</summary>
        public string Log { get; private set; }

        public Exporter LastExporter { get; private set; }

        public ProjectSummary Summary(string project)
        {
            return LastExporter.Summaries.Single(s => s.Project == project);
        }

        /// <summary>Runs Prepare + Run like Program.Main. Clears the fake's call log first.</summary>
        public Exporter Run(bool incremental, int batch, params ITiaPortalProcess[] processes)
        {
            Tia.ClearCalls();
            _log = new StringWriter();
            var exporter = new Exporter(Export, _log, new ExporterSettings
            {
                Incremental = incremental,
                BatchSize = batch,
                StateDirectory = StateDir,
                TempDirectory = TempDir,
            });
            exporter.PrepareExportLocation();
            exporter.Run(processes);
            Log = _log.ToString();
            LastExporter = exporter;
            return exporter;
        }

        public Exporter Full(params ITiaPortalProcess[] processes) { return Run(false, 1, processes); }
        public Exporter Incremental(params ITiaPortalProcess[] processes) { return Run(true, 1, processes); }

        public List<string> ExportedFiles()
        {
            if (!Directory.Exists(Export)) return new List<string>();
            return Directory.GetFiles(Export, "*", SearchOption.AllDirectories)
                .Select(f => f.Substring(Export.Length + 1))
                .OrderBy(f => f, StringComparer.Ordinal)
                .ToList();
        }

        public string PathOf(string relative) { return Path.Combine(Export, relative); }
        public string Read(string relative) { return File.ReadAllText(PathOf(relative)); }
        public byte[] Bytes(string relative) { return File.ReadAllBytes(PathOf(relative)); }

        /// <summary>Names of blocks/types sent to GenerateSource in the last run (batch calls split up, verification calls included).</summary>
        public List<string> Generated()
        {
            return Tia.CallsStartingWith("GenerateSource ")
                .Select(c => c.Substring("GenerateSource ".Length).Split(new[] { " -> " }, StringSplitOptions.None)[0])
                .SelectMany(n => n.Split(','))
                .ToList();
        }
    }
}
