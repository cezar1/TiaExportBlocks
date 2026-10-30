using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using TiaExportBlocks.Tests.Fakes;
using Xunit;

namespace TiaExportBlocks.Tests
{
    public class BenchmarkTests : IDisposable
    {
        private readonly FakeClock _clock = new FakeClock();
        private readonly ExportHarness _h = new ExportHarness();

        public void Dispose()
        {
            _h.Dispose();
        }

        private static TimeSpan S(double seconds) { return TimeSpan.FromSeconds(seconds); }

        // ---------------- Benchmark itself ----------------

        [Fact]
        public void Measure_AddsTimeCallsAndItemsPerPhase()
        {
            var b = new Benchmark(_clock);
            b.Start();
            using (b.Measure(Benchmark.GenerateSource, 50)) _clock.Advance(S(3));
            using (b.Measure(Benchmark.GenerateSource, 20)) _clock.Advance(S(2));
            using (b.Measure(Benchmark.ExportXml, 1)) _clock.Advance(S(0.5));
            _clock.Advance(S(1)); // not in any phase
            b.Stop();

            Assert.Equal(S(6.5), b.Total);
            Assert.Equal(S(5), b[Benchmark.GenerateSource].Time);
            Assert.Equal(2, b[Benchmark.GenerateSource].Calls);
            Assert.Equal(70, b[Benchmark.GenerateSource].Items);
            Assert.Equal(S(2.5), b[Benchmark.GenerateSource].AveragePerCall);
            Assert.Equal(S(1), b.OtherTime);
            Assert.Equal(S(1), b[Benchmark.Other].Time);
            Assert.Equal(0, b[Benchmark.Attach].Calls);
        }

        [Fact]
        public void Measure_CountsCallsThatThrow()
        {
            var b = new Benchmark(_clock);

            Assert.Throws<InvalidOperationException>(() =>
            {
                using (b.Measure(Benchmark.ExportXml, 1))
                {
                    _clock.Advance(S(2));
                    throw new InvalidOperationException();
                }
            });

            Assert.Equal(S(2), b[Benchmark.ExportXml].Time);
            Assert.Equal(1, b[Benchmark.ExportXml].Calls);
        }

        [Fact]
        public void Total_RunsUntilStop_AndFirstStartCounts()
        {
            var b = new Benchmark(_clock);
            Assert.Equal(TimeSpan.Zero, b.Total);

            b.Start();
            _clock.Advance(S(1));
            b.Start(); // ignored
            _clock.Advance(S(1));
            Assert.True(b.IsRunning);
            Assert.Equal(S(2), b.Total);

            b.Stop();
            _clock.Advance(S(10));
            Assert.False(b.IsRunning);
            Assert.Equal(S(2), b.Total);
        }

        [Theory]
        [InlineData(0.0004, "0 ms")]
        [InlineData(0.250, "250 ms")]
        [InlineData(1.0, "1.0 s")]
        [InlineData(47.26, "47.3 s")]
        [InlineData(329.0, "5 min 29 s")]
        [InlineData(3725.0, "62 min 05 s")]
        public void FormatDuration_IsReadable(double seconds, string expected)
        {
            Assert.Equal(expected, Benchmark.FormatDuration(S(seconds)));
        }

        [Fact]
        public void FormatReport_ListsUsedPhasesWithShares()
        {
            var b = new Benchmark(_clock);
            b.Start();
            using (b.Measure(Benchmark.GenerateSource, 10)) _clock.Advance(S(8));
            _clock.Advance(S(2));
            b.Stop();

            string report = b.FormatReport();

            Assert.StartsWith("Benchmark: total 10.0 s", report);
            Assert.Contains(Benchmark.GenerateSource, report);
            Assert.Contains("80 %", report);
            Assert.Contains(Benchmark.Other, report);
            Assert.Contains("20 %", report);
            Assert.DoesNotContain(Benchmark.Attach, report); // never called
        }

        [Fact]
        public void Csv_LineMatchesHeader_AndUsesInvariantCulture()
        {
            var previous = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
            try
            {
                var b = new Benchmark(_clock);
                b.Start();
                using (b.Measure(Benchmark.GenerateSource, 500)) _clock.Advance(TimeSpan.FromMilliseconds(46999.6));
                b.Stop();
                var summaries = new[] { new ProjectSummary { Exported = 520, Unchanged = 3, Failed = 1, Removed = 2 } };

                string line = b.ToCsvLine(new DateTime(2026, 10, 30, 20, 0, 0, DateTimeKind.Utc), @"C:\Exports\Line, 3", true, 50, summaries);

                string[] header = Benchmark.CsvHeader.Split(',');
                string[] fields = ParseCsv(line);
                Assert.Equal(header.Length, fields.Length);
                var row = header.Zip(fields, (h, f) => new { h, f }).ToDictionary(x => x.h, x => x.f);
                Assert.Equal("2026-10-30T20:00:00Z", row["timestamp_utc"]);
                Assert.Equal(@"C:\Exports\Line, 3", row["export_location"]);
                Assert.Equal("incremental", row["mode"]);
                Assert.Equal("50", row["batch_size"]);
                Assert.Equal("47000", row["total_ms"]);
                Assert.Equal("47000", row["generate_ms"]);
                Assert.Equal("500", row["generated_items"]);
                Assert.Equal("520", row["exported"]);
                Assert.Equal("2", row["removed"]);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = previous;
            }
        }

        [Fact]
        public void AppendCsv_WritesHeaderOnceThenAppends()
        {
            string csv = Path.Combine(_h.Sandbox, "bench", "history.csv");

            Benchmark.AppendCsv(csv, "a");
            Benchmark.AppendCsv(csv, "b");

            Assert.Equal(new[] { Benchmark.CsvHeader, "a", "b" }, File.ReadAllLines(csv));
        }

        private static string[] ParseCsv(string line)
        {
            var fields = new System.Collections.Generic.List<string>();
            var current = new System.Text.StringBuilder();
            bool quoted = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (quoted)
                {
                    if (c == '"' && i + 1 < line.Length && line[i + 1] == '"') { current.Append('"'); i++; }
                    else if (c == '"') quoted = false;
                    else current.Append(c);
                }
                else if (c == '"') quoted = true;
                else if (c == ',') { fields.Add(current.ToString()); current.Clear(); }
                else current.Append(c);
            }
            fields.Add(current.ToString());
            return fields.ToArray();
        }

        // ---------------- Export runs with simulated TIA timings ----------------

        private FakeProcess Project(int blocks)
        {
            var tia = _h.Tia;
            var plc = tia.Plc("PLC_1")
                .WithBlocks(Enumerable.Range(1, blocks).Select(i => tia.Scl("FB_" + i)).ToArray())
                .WithTagTables(tia.PlcTagTable("Tags"), tia.PlcTagTable("IO"));
            return tia.Process(1, tia.Project("Line3",
                tia.S71500Station("S", "PLC_1", plc),
                tia.HmiStation("H", "HMI_1", tia.Hmi("HMI_1").WithTextLists(tia.TextList("Alarms")))));
        }

        private void SimulateTiaTimings()
        {
            _h.Tia.AttachCost = S(2);
            _h.Tia.ReadSoftwareCost = S(0.5);
            _h.Tia.GenerateCallCost = S(0.6);
            _h.Tia.GeneratePerItemCost = TimeSpan.FromMilliseconds(30);
            _h.Tia.ExportCost = TimeSpan.FromMilliseconds(100);
        }

        [Fact]
        public void FullExport_OneBlockPerCall_PhasesAddUp()
        {
            SimulateTiaTimings();

            Benchmark b = _h.Full(Project(20)).Benchmark;

            Assert.Equal(S(2), b[Benchmark.Attach].Time);
            Assert.Equal(1, b[Benchmark.Attach].Calls);
            Assert.Equal(20, b[Benchmark.GenerateSource].Calls);
            Assert.Equal(20, b[Benchmark.GenerateSource].Items);
            Assert.Equal(S(20 * 0.63), b[Benchmark.GenerateSource].Time);
            Assert.Equal(3, b[Benchmark.ExportXml].Calls);
            Assert.Equal(S(0.3), b[Benchmark.ExportXml].Time);
            Assert.Equal(1, b[Benchmark.Prepare].Calls);
            Assert.Equal(S(1.0), b.OtherTime); // reading software of PLC_1 and HMI_1
            Assert.Equal(S(2 + 12.6 + 0.3 + 1.0), b.Total);
            Assert.False(b.IsRunning);
        }

        [Fact]
        public void Batch_ShowsUpAsFewerAndCheaperGenerateCalls()
        {
            SimulateTiaTimings();

            Benchmark b = _h.Run(false, 50, Project(20)).Benchmark;

            Assert.Equal(1, b[Benchmark.GenerateSource].Calls);
            Assert.Equal(20, b[Benchmark.GenerateSource].Items);
            Assert.Equal(S(0.6 + 20 * 0.03), b[Benchmark.GenerateSource].Time);
            Assert.Equal(1, b[Benchmark.BatchVerification].Calls);
            Assert.Equal(S(0.63), b[Benchmark.BatchVerification].Time);
        }

        [Fact]
        public void Incremental_NothingChanged_NoGenerateTime()
        {
            SimulateTiaTimings();
            _h.Incremental(Project(20));

            Benchmark b = _h.Incremental(Project(20)).Benchmark;

            Assert.Equal(0, b[Benchmark.GenerateSource].Calls);
            Assert.Equal(TimeSpan.Zero, b[Benchmark.GenerateSource].Time);
            Assert.Equal(3, b[Benchmark.ExportXml].Calls);
        }

        [Fact]
        public void ProjectSummary_HasTheProjectTime()
        {
            SimulateTiaTimings();
            var tia = _h.Tia;
            var a = tia.Project("A", tia.S71500Station("SA", "PLC_A", tia.Plc("PLC_A").WithBlocks(tia.Scl("A1"))));
            var bProject = tia.Project("B", tia.S71500Station("SB", "PLC_B", tia.Plc("PLC_B").WithBlocks(tia.Scl("B1"), tia.Scl("B2"))));

            _h.Full(tia.Process(1, a, bProject));

            Assert.Equal(S(0.5 + 0.63), _h.Summary("A").Elapsed);
            Assert.Equal(S(0.5 + 2 * 0.63), _h.Summary("B").Elapsed);
            Assert.Contains("in 1.1 s", _h.Log);
        }

        [Fact]
        public void PrepareIsPartOfTheTotal()
        {
            Directory.CreateDirectory(_h.Export);
            File.WriteAllText(Path.Combine(_h.Export, "old.scl"), "x");

            Benchmark b = _h.Full(Project(1)).Benchmark;

            Assert.Equal(1, b[Benchmark.Prepare].Calls);
            Assert.True(b.Total >= b[Benchmark.Prepare].Time);
        }

        [Fact]
        public void Report_ForARealRun_ListsTheMeasuredPhases()
        {
            SimulateTiaTimings();

            string report = _h.Run(false, 50, Project(20)).Benchmark.FormatReport();

            foreach (var phase in new[] { Benchmark.Prepare, Benchmark.Attach, Benchmark.GenerateSource, Benchmark.BatchVerification, Benchmark.ExportXml, Benchmark.Cleanup, Benchmark.Other })
            {
                Assert.Contains(phase, report);
            }
        }
    }
}
