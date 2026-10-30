using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace TiaExportBlocks
{
    /// <summary>Monotonic time source; replaceable in tests.</summary>
    public interface IClock
    {
        TimeSpan Elapsed { get; }
    }

    public sealed class StopwatchClock : IClock
    {
        private readonly Stopwatch _watch = Stopwatch.StartNew();
        public TimeSpan Elapsed { get { return _watch.Elapsed; } }
    }

    /// <summary>Time, number of calls and number of items for one phase of an export.</summary>
    public sealed class PhaseStats
    {
        public TimeSpan Time;
        public int Calls;
        public int Items;

        public TimeSpan AveragePerCall
        {
            get { return Calls == 0 ? TimeSpan.Zero : TimeSpan.FromTicks(Time.Ticks / Calls); }
        }
    }

    /// <summary>
    /// Measures how long an export takes: in total and per phase (wiping the export location, attaching to TIA,
    /// generating sources, exporting XML, ...). Everything not inside a measured phase (reading the project
    /// tree from TIA, moving and hashing files) is reported as "Other".
    /// </summary>
    public sealed class Benchmark
    {
        public const string Prepare = "Prepare export location";
        public const string Attach = "Attach to TIA Portal";
        public const string GenerateSource = "Generate sources (blocks, UDTs)";
        public const string BatchVerification = "Batch verification";
        public const string ExportXml = "Export XML (tag tables, text lists)";
        public const string Cleanup = "Cleanup and state";
        public const string Other = "Other (reading TIA, files)";

        /// <summary>Order in reports and CSV.</summary>
        public static readonly string[] Phases = { Prepare, Attach, GenerateSource, BatchVerification, ExportXml, Cleanup };

        private readonly IClock _clock;
        private readonly Dictionary<string, PhaseStats> _phases = new Dictionary<string, PhaseStats>();
        private TimeSpan? _start;
        private TimeSpan? _stop;

        public Benchmark() : this(new StopwatchClock()) { }

        public Benchmark(IClock clock)
        {
            if (clock == null) throw new ArgumentNullException("clock");
            _clock = clock;
            foreach (string phase in Phases) _phases[phase] = new PhaseStats();
        }

        /// <summary>Starts the total time (only the first call counts).</summary>
        public void Start()
        {
            if (!_start.HasValue) _start = _clock.Elapsed;
        }

        /// <summary>Stops the total time (the last call counts).</summary>
        public void Stop()
        {
            Start();
            _stop = _clock.Elapsed;
        }

        public bool IsRunning { get { return _start.HasValue && !_stop.HasValue; } }

        /// <summary>Total time from <see cref="Start"/> to <see cref="Stop"/> (or now, while running).</summary>
        public TimeSpan Total
        {
            get
            {
                if (!_start.HasValue) return TimeSpan.Zero;
                return (_stop ?? _clock.Elapsed) - _start.Value;
            }
        }

        public PhaseStats this[string phase]
        {
            get
            {
                if (phase == Other) return new PhaseStats { Time = OtherTime };
                return _phases[phase];
            }
        }

        public TimeSpan OtherTime
        {
            get
            {
                TimeSpan measured = TimeSpan.FromTicks(_phases.Values.Sum(p => p.Time.Ticks));
                TimeSpan other = Total - measured;
                return other < TimeSpan.Zero ? TimeSpan.Zero : other;
            }
        }

        /// <summary>Times one call of a phase: <c>using (benchmark.Measure(Benchmark.ExportXml)) { ... }</c>. Also counted if the call throws.</summary>
        public IDisposable Measure(string phase, int items = 0)
        {
            Start();
            return new Scope(this, phase, items, _clock.Elapsed);
        }

        private sealed class Scope : IDisposable
        {
            private readonly Benchmark _owner;
            private readonly string _phase;
            private readonly int _items;
            private readonly TimeSpan _started;
            private bool _done;

            public Scope(Benchmark owner, string phase, int items, TimeSpan started)
            {
                _owner = owner;
                _phase = phase;
                _items = items;
                _started = started;
            }

            public void Dispose()
            {
                if (_done) return;
                _done = true;
                PhaseStats stats = _owner._phases[_phase];
                stats.Time += _owner._clock.Elapsed - _started;
                stats.Calls++;
                stats.Items += _items;
            }
        }

        // ---------------- reporting ----------------

        public static string FormatDuration(TimeSpan time)
        {
            if (time.TotalSeconds < 1) return ((int)Math.Round(time.TotalMilliseconds)).ToString(CultureInfo.InvariantCulture) + " ms";
            if (time.TotalMinutes < 1) return time.TotalSeconds.ToString("0.0", CultureInfo.InvariantCulture) + " s";
            return ((int)time.TotalMinutes).ToString(CultureInfo.InvariantCulture) + " min " + time.Seconds.ToString("00", CultureInfo.InvariantCulture) + " s";
        }

        /// <summary>Human-readable table of the phases.</summary>
        public string FormatReport()
        {
            var rows = new List<string[]> { new[] { "Phase", "Time", "Share", "Calls", "Items", "Avg/call" } };
            double total = Math.Max(Total.TotalMilliseconds, 1e-9);
            foreach (string phase in Phases.Concat(new[] { Other }))
            {
                PhaseStats s = this[phase];
                if (phase != Other && s.Calls == 0) continue;
                rows.Add(new[]
                {
                    phase,
                    FormatDuration(s.Time),
                    (s.Time.TotalMilliseconds / total * 100).ToString("0", CultureInfo.InvariantCulture) + " %",
                    s.Calls == 0 ? "" : s.Calls.ToString(CultureInfo.InvariantCulture),
                    s.Items == 0 ? "" : s.Items.ToString(CultureInfo.InvariantCulture),
                    s.Calls == 0 ? "" : FormatDuration(s.AveragePerCall),
                });
            }
            int[] widths = Enumerable.Range(0, 6).Select(c => rows.Max(r => r[c].Length)).ToArray();
            var sb = new StringBuilder();
            sb.AppendLine("Benchmark: total " + FormatDuration(Total));
            foreach (var row in rows)
            {
                sb.Append("  ").Append(row[0].PadRight(widths[0]));
                for (int c = 1; c < 6; c++) sb.Append("  ").Append(row[c].PadLeft(widths[c]));
                sb.AppendLine();
            }
            return sb.ToString();
        }

        // ---------------- CSV history ----------------

        public static string DefaultCsvPath
        {
            get { return Path.Combine(ExportState.DefaultDirectory, "benchmark.csv"); }
        }

        public static string CsvHeader
        {
            get
            {
                return "timestamp_utc,export_location,mode,batch_size,total_ms,prepare_ms,attach_ms,generate_ms,generate_calls,generated_items,"
                    + "verify_ms,export_xml_ms,export_xml_calls,cleanup_ms,other_ms,exported,unchanged,failed,removed";
            }
        }

        /// <summary>One CSV line for this run (invariant culture, comma separated, quoted where needed).</summary>
        public string ToCsvLine(DateTime timestampUtc, string exportLocation, bool incremental, int batchSize, IEnumerable<ProjectSummary> summaries)
        {
            var list = (summaries ?? Enumerable.Empty<ProjectSummary>()).ToList();
            Func<TimeSpan, string> ms = t => ((long)Math.Round(t.TotalMilliseconds)).ToString(CultureInfo.InvariantCulture);
            Func<int, string> n = i => i.ToString(CultureInfo.InvariantCulture);
            var fields = new[]
            {
                timestampUtc.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture),
                Quote(exportLocation),
                incremental ? "incremental" : "full",
                n(batchSize),
                ms(Total),
                ms(this[Prepare].Time),
                ms(this[Attach].Time),
                ms(this[GenerateSource].Time),
                n(this[GenerateSource].Calls),
                n(this[GenerateSource].Items),
                ms(this[BatchVerification].Time),
                ms(this[ExportXml].Time),
                n(this[ExportXml].Calls),
                ms(this[Cleanup].Time),
                ms(OtherTime),
                n(list.Sum(s => s.Exported)),
                n(list.Sum(s => s.Unchanged)),
                n(list.Sum(s => s.Failed)),
                n(list.Sum(s => s.Removed)),
            };
            return string.Join(",", fields);
        }

        /// <summary>Appends a line to the CSV file, writing the header first if the file is new.</summary>
        public static void AppendCsv(string path, string line)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
            bool isNew = !File.Exists(path) || new FileInfo(path).Length == 0;
            using (var writer = new StreamWriter(path, true, new UTF8Encoding(false)))
            {
                if (isNew) writer.WriteLine(CsvHeader);
                writer.WriteLine(line);
            }
        }

        private static string Quote(string value)
        {
            if (value == null) return "";
            if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) < 0) return value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}
