using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TiaExportBlocks.Tests.Fakes;
using Xunit;

namespace TiaExportBlocks.Tests
{
    /// <summary>
    /// Runs the real <see cref="Exporter"/> against a simulated TIA Portal (<see cref="FakeTia"/>)
    /// in an isolated sandbox folder, and checks the files produced, the console log and the
    /// exact sequence of Openness calls.
    /// </summary>
    public class ExporterTests : IDisposable
    {
        private readonly string _sandbox;
        private readonly string _export;
        private readonly StringWriter _log = new StringWriter();
        private readonly FakeTia _tia = new FakeTia();

        public ExporterTests()
        {
            _sandbox = Path.Combine(Path.GetTempPath(), "TiaExporterTests_" + Guid.NewGuid().ToString("N"));
            _export = Path.Combine(_sandbox, "export");
            Directory.CreateDirectory(_sandbox);
        }

        public void Dispose()
        {
            if (Directory.Exists(_sandbox)) Directory.Delete(_sandbox, true);
        }

        // ---------------- helpers ----------------

        private Exporter NewExporter()
        {
            return new Exporter(_export, _log);
        }

        private void Run(params ITiaPortalProcess[] processes)
        {
            Exporter exporter = NewExporter();
            exporter.PrepareExportLocation();
            exporter.Run(processes);
        }

        private string Log { get { return _log.ToString(); } }

        /// <summary>All files under the export folder, relative, '\'-separated, sorted.</summary>
        private List<string> ExportedFiles()
        {
            if (!Directory.Exists(_export)) return new List<string>();
            return Directory.GetFiles(_export, "*", SearchOption.AllDirectories)
                .Select(f => f.Substring(_export.Length + 1))
                .OrderBy(f => f, StringComparer.Ordinal)
                .ToList();
        }

        private string Read(string relativePath)
        {
            return File.ReadAllText(Path.Combine(_export, relativePath));
        }

        private FakeProcess OneProject(params FakeDevice[] devices)
        {
            return _tia.Process(1, _tia.Project("Line3", devices));
        }

        // ---------------- construction / preparation ----------------

        [Fact]
        public void Constructor_NullArguments_Throw()
        {
            Assert.Throws<ArgumentNullException>(() => new Exporter(null, _log));
            Assert.Throws<ArgumentNullException>(() => new Exporter(_export, null));
        }

        [Fact]
        public void PrepareExportLocation_DeletesPreviousExportAndRecreatesFolder()
        {
            Directory.CreateDirectory(Path.Combine(_export, "old", "stuff"));
            File.WriteAllText(Path.Combine(_export, "old", "stuff", "stale.scl"), "x");

            NewExporter().PrepareExportLocation();

            Assert.True(Directory.Exists(_export));
            Assert.Empty(Directory.GetFileSystemEntries(_export));
            Assert.Contains("Deleting location " + _export + ".", Log);
        }

        [Fact]
        public void PrepareExportLocation_RemovesReadOnlyHiddenAndSystemFiles()
        {
            string ro = Path.Combine(_export, "PLC_1", "scl", "ReadOnly.scl");
            string hidden = Path.Combine(_export, "hidden.xml");
            string system = Path.Combine(_export, "PLC_1", "udt", "sys.udt");
            foreach (var f in new[] { ro, hidden, system })
            {
                Directory.CreateDirectory(Path.GetDirectoryName(f));
                File.WriteAllText(f, "stale");
            }
            File.SetAttributes(ro, FileAttributes.ReadOnly);
            File.SetAttributes(hidden, FileAttributes.Hidden);
            File.SetAttributes(system, FileAttributes.System | FileAttributes.ReadOnly);
            var roDir = new DirectoryInfo(Path.Combine(_export, "PLC_1"));
            roDir.Attributes |= FileAttributes.ReadOnly;

            NewExporter().PrepareExportLocation();

            Assert.True(Directory.Exists(_export));
            Assert.Empty(Directory.GetFileSystemEntries(_export));
        }

        [Fact]
        public void PrepareExportLocation_LockedFile_ThrowsAndNamesIt()
        {
            string locked = Path.Combine(_export, "PLC_1", "scl", "Locked.scl");
            Directory.CreateDirectory(Path.GetDirectoryName(locked));
            File.WriteAllText(locked, "stale");

            using (new FileStream(locked, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                var exc = Assert.Throws<ExportLocationNotCleanException>(() => NewExporter().PrepareExportLocation());
                Assert.Contains(locked, exc.Remaining);
                Assert.Contains("nothing was exported", exc.Message);
            }
        }

        [Fact]
        public void Run_FilesAppearingDuringExport_AreRemovedAfterwards()
        {
            string stray = Path.Combine(_export, "stray.txt");
            string strayScl = Path.Combine(_export, "PLC_1", "scl", "Old", "Removed.scl");
            var block = _tia.Scl("Main");
            block.OnGenerate = () =>
            {
                File.WriteAllText(stray, "x");
                Directory.CreateDirectory(Path.GetDirectoryName(strayScl));
                File.WriteAllText(strayScl, "x");
                File.SetAttributes(strayScl, FileAttributes.ReadOnly);
            };

            Run(OneProject(_tia.S71500Station("S", "PLC_1", _tia.Plc("PLC_1").WithBlocks(block))));

            Assert.Equal(new[] { @"PLC_1\scl\Program blocks\Main.scl" }, ExportedFiles());
            Assert.False(Directory.Exists(Path.Combine(_export, "PLC_1", "scl", "Old")));
            Assert.Contains("Removed " + stray + " (not written by this export)", Log);
        }

        [Fact]
        public void Run_ExportLocationHoldsExactlyTheFilesWrittenThisRun()
        {
            var plc = _tia.Plc("PLC_1")
                .WithBlocks(_tia.Scl("Main"), _tia.Db("DB1"), _tia.Scl("Broken").Failing("inconsistent", writePartialFile: true))
                .WithTypes(_tia.Udt("UDT1"))
                .WithTagTables(_tia.PlcTagTable("Tags"));
            var hmi = _tia.Hmi("HMI_1").WithTagTables(_tia.HmiTagTable("T")).WithTextLists(_tia.TextList("L"));
            var exporter = NewExporter();
            exporter.PrepareExportLocation();

            exporter.Run(new[] { OneProject(_tia.S71500Station("S", "PLC_1", plc), _tia.HmiStation("H", "HMI_1", hmi)) });

            var onDisk = Directory.GetFiles(_export, "*", SearchOption.AllDirectories).OrderBy(f => f, StringComparer.OrdinalIgnoreCase);
            Assert.Equal(exporter.WrittenFiles.OrderBy(f => f, StringComparer.OrdinalIgnoreCase), onDisk);
            Assert.Equal(6, exporter.WrittenFiles.Count);
        }

        [Fact]
        public void PrepareExportLocation_FolderMissing_CreatesItWithoutDeleteMessage()
        {
            NewExporter().PrepareExportLocation();

            Assert.True(Directory.Exists(_export));
            Assert.DoesNotContain("Deleting location", Log);
        }

        // ---------------- processes / projects ----------------

        [Fact]
        public void Run_NoTiaRunning_ExportsNothing()
        {
            Run();

            Assert.Empty(ExportedFiles());
            Assert.Empty(_tia.Calls);
            Assert.Contains("Enumerating TIA processes..", Log);
        }

        [Fact]
        public void Run_AttachFails_StopsWithoutTouchingLaterInstances()
        {
            var locked = _tia.ProcessThatCannotAttach(10);
            var healthy = _tia.Process(11, _tia.Project("P", _tia.S71500Station("S", "PLC_1", _tia.Plc("PLC_1").WithBlocks(_tia.Scl("Main")))));

            Run(locked, healthy);

            Assert.Contains("Failed to attach to TIA Portal.", Log);
            Assert.Equal(new[] { "Attach 10" }, _tia.Calls);
            Assert.Empty(ExportedFiles());
        }

        [Fact]
        public void Run_InstanceWithoutProjects_IsSkippedAndNextInstanceExported()
        {
            var empty = _tia.ProcessWithoutProjects(1);
            var full = _tia.Process(2, _tia.Project("P", _tia.S71500Station("S", "PLC_1", _tia.Plc("PLC_1").WithBlocks(_tia.Scl("Main")))));

            Run(empty, full);

            Assert.Contains("No projects found in TIA Portal.", Log);
            Assert.Equal(new[] { @"PLC_1\scl\Program blocks\Main.scl" }, ExportedFiles());
        }

        [Fact]
        public void Run_InstanceWithNullProjectList_IsSkipped()
        {
            var process = new FakeProcess_NullProjects();

            Run(process);

            Assert.Contains("No projects found in TIA Portal.", Log);
        }

        private sealed class FakeProcess_NullProjects : ITiaPortalProcess, ITiaPortal
        {
            public int Id { get { return 99; } }
            public string ProjectPath { get { return null; } }
            public ITiaPortal Attach() { return this; }
            public IList<ITiaProject> Projects { get { return null; } }
        }

        [Fact]
        public void Run_SeveralInstancesAndProjects_ExportsAll()
        {
            var p1 = _tia.Process(1,
                _tia.Project("A", _tia.S71500Station("S1", "PLC_A", _tia.Plc("PLC_A").WithBlocks(_tia.Scl("MainA")))),
                _tia.Project("B", _tia.S71500Station("S2", "PLC_B", _tia.Plc("PLC_B").WithBlocks(_tia.Scl("MainB")))));
            var p2 = _tia.Process(2,
                _tia.Project("C", _tia.S71500Station("S3", "PLC_C", _tia.Plc("PLC_C").WithBlocks(_tia.Scl("MainC")))));

            Run(p1, p2);

            Assert.Equal(new[]
            {
                @"PLC_A\scl\Program blocks\MainA.scl",
                @"PLC_B\scl\Program blocks\MainB.scl",
                @"PLC_C\scl\Program blocks\MainC.scl",
            }, ExportedFiles());
            Assert.Contains("Handling project A", Log);
            Assert.Contains("Handling project B", Log);
            Assert.Contains("Handling project C", Log);
        }

        // ---------------- PLC: full layout ----------------

        [Fact]
        public void Run_TypicalPlc_ProducesExpectedLayoutAndContent()
        {
            var main = _tia.Scl("Main");
            var recipe = _tia.Db("DB_Recipe");
            var legacy = _tia.Stl("FC_Legacy");
            var plc = _tia.Plc("PLC_1")
                .WithBlocks(main, _tia.Lad("OB_Cyclic"))
                .WithGroups(
                    _tia.Group("Data").With(recipe),
                    _tia.Group("Legacy").With(_tia.Group("Line 2").With(legacy)))
                .WithTypes(_tia.Udt("UDT_Motor"))
                .WithTagTables(_tia.PlcTagTable("Default tag table"))
                .WithTagGroups(_tia.TagGroup()
                    .With(_tia.PlcTagTable("Inputs"))
                    .With(_tia.TagGroup().With(_tia.PlcTagTable("Outputs"))));

            Run(OneProject(_tia.S71500Station("Station_1", "PLC_1", plc)));

            Assert.Equal(new[]
            {
                @"PLC_1\db\Program blocks\Data\DB_Recipe.db",
                @"PLC_1\scl\Program blocks\Main.scl",
                @"PLC_1\stl\Program blocks\Legacy\Line 2\FC_Legacy.awl",
                @"PLC_1\tag_tables\xml\Default tag table.xml",
                @"PLC_1\tag_tables\xml\Inputs.xml",
                @"PLC_1\tag_tables\xml\Outputs.xml",
                @"PLC_1\udt\UDT_Motor.udt",
            }, ExportedFiles());
            Assert.Equal(main.ExpectedSource, Read(@"PLC_1\scl\Program blocks\Main.scl"));
            Assert.Equal(recipe.ExpectedSource, Read(@"PLC_1\db\Program blocks\Data\DB_Recipe.db"));
            Assert.Equal(legacy.ExpectedSource, Read(@"PLC_1\stl\Program blocks\Legacy\Line 2\FC_Legacy.awl"));
            Assert.StartsWith("TYPE \"UDT_Motor\"", Read(@"PLC_1\udt\UDT_Motor.udt"));
            Assert.Contains("<Name>Outputs</Name>", Read(@"PLC_1\tag_tables\xml\Outputs.xml"));
        }

        [Fact]
        public void Run_TypicalPlc_CallsOpennessInExpectedOrder()
        {
            var plc = _tia.Plc("PLC_1")
                .WithBlocks(_tia.Scl("Main"))
                .WithGroups(_tia.Group("Sub").With(_tia.Db("DB1")))
                .WithTypes(_tia.Udt("UDT1"))
                .WithTagTables(_tia.PlcTagTable("Tags"));

            Run(OneProject(_tia.S71500Station("S", "PLC_1", plc)));

            var shortCalls = _tia.Calls.Select(c => c.Split(new[] { " -> " }, StringSplitOptions.None)[0]).ToList();
            Assert.Equal(new[]
            {
                "Attach 1",
                "GetSoftware PLC_1",
                "GenerateSource Main",
                "GenerateSource DB1",
                "GenerateSource UDT1",
                "Export Tags",
            }, shortCalls);
        }

        [Fact]
        public void Run_BlocksAreGeneratedIntoTempFileInsideExportFolder_ThenMoved()
        {
            Run(OneProject(_tia.S71500Station("S", "PLC_1", _tia.Plc("PLC_1").WithBlocks(_tia.Scl("Main")))));

            string generate = _tia.CallsStartingWith("GenerateSource Main").Single();
            Assert.EndsWith(" -> " + Path.Combine(_export, "Main.scl"), generate);
            Assert.False(File.Exists(Path.Combine(_export, "Main.scl")), "temp file must be moved away");
            Assert.Equal(new[] { "export" }, Directory.GetFileSystemEntries(_sandbox).Select(Path.GetFileName));
        }

        [Fact]
        public void Run_TagTablesAreExportedToTempFileThenMoved()
        {
            Run(OneProject(_tia.S71500Station("S", "PLC_1", _tia.Plc("PLC_1").WithTagTables(_tia.PlcTagTable("Tags")))));

            Assert.EndsWith(" -> " + Path.Combine(_export, "Tags.xml"), _tia.CallsStartingWith("Export Tags").Single());
            Assert.Equal(new[] { @"PLC_1\tag_tables\xml\Tags.xml" }, ExportedFiles());
        }

        [Fact]
        public void Run_NothingIsLeftAtExportRoot()
        {
            var plc = _tia.Plc("PLC_1")
                .WithBlocks(_tia.Scl("A"), _tia.Db("B"), _tia.Stl("C"), _tia.Scl("Broken").Failing("inconsistent", writePartialFile: true))
                .WithTagTables(_tia.PlcTagTable("T1"), _tia.PlcTagTable("T2"));

            Run(OneProject(_tia.S71500Station("S", "PLC_1", plc)));

            Assert.Empty(Directory.GetFiles(_export));
        }

        // ---------------- PLC: what is and isn't exported ----------------

        [Theory]
        [InlineData("LAD")]
        [InlineData("FBD")]
        [InlineData("GRAPH")]
        [InlineData("CEM")]
        public void Run_GraphicalBlocks_AreSkippedAndNeverSentToTia(string language)
        {
            Run(OneProject(_tia.S71500Station("S", "PLC_1", _tia.Plc("PLC_1").WithBlocks(_tia.Block("OB_Main", language)))));

            Assert.Contains("OB_Main programming language " + language + " not supported", Log);
            Assert.Empty(_tia.CallsStartingWith("GenerateSource"));
            Assert.Empty(ExportedFiles());
        }

        [Fact]
        public void Run_S71500ItemsWithoutPlcInName_AreNotAskedForSoftware()
        {
            var station = _tia.Device("S", FakeTia.S71500,
                _tia.Item("Rail_0", null),
                _tia.Item("CPU 1516", _tia.Plc("x").WithBlocks(_tia.Scl("Hidden"))),
                _tia.Item("PLC_1", _tia.Plc("PLC_1").WithBlocks(_tia.Scl("Main"))));

            Run(OneProject(station));

            Assert.Equal(new[] { "GetSoftware PLC_1" }, _tia.CallsStartingWith("GetSoftware"));
            Assert.Equal(new[] { @"PLC_1\scl\Program blocks\Main.scl" }, ExportedFiles());
        }

        [Fact]
        public void Run_PlcItemWithoutSoftwareContainer_IsSkipped()
        {
            Run(OneProject(_tia.Device("S", FakeTia.S71500, _tia.Item("PLC_1", null))));

            Assert.Empty(ExportedFiles());
            Assert.Contains("Handling PLC device item", Log);
        }

        [Fact]
        public void Run_PlcItemWithNonPlcSoftware_IsSkippedWithMessage()
        {
            Run(OneProject(_tia.Device("S", FakeTia.S71500, _tia.Item("PLC_1", new FakeOtherSoftware("odd")))));

            Assert.Contains("Device item [PLC_1] has no PLC software, skipped", Log);
            Assert.Empty(ExportedFiles());
        }

        [Fact]
        public void Run_S71200Cpu_IsSkippedWithoutCrashing_AndLaterDevicesStillExported()
        {
            // Before: non-S7-1500 devices were treated as HMI and "software as HmiTarget" gave null
            // for an S7-1200 CPU, which crashed the whole run with a NullReferenceException.
            var s71200 = _tia.Device("Station_1200", FakeTia.S71200, _tia.Item("PLC_2", _tia.Plc("PLC_2").WithBlocks(_tia.Scl("Main"))));
            var hmi = _tia.HmiStation("HMI_Station", "HMI_RT_1", _tia.Hmi("HMI_RT_1").WithTagTables(_tia.HmiTagTable("Tags")));

            Run(OneProject(s71200, hmi));

            Assert.Contains("Device item [PLC_2] software is not an HMI target, skipped", Log);
            Assert.Equal(new[] { @"HMI_RT_1\hmi_tag_tables\xml\Tags.xml" }, ExportedFiles());
        }

        // ---------------- PLC: failures in TIA ----------------

        [Fact]
        public void Run_BlockGenerationFails_ErrorLoggedAndOtherBlocksExported()
        {
            var plc = _tia.Plc("PLC_1").WithBlocks(
                _tia.Scl("A"),
                _tia.Scl("Protected").Failing("Block 'Protected' is know-how protected."),
                _tia.Scl("C"));

            Run(OneProject(_tia.S71500Station("S", "PLC_1", plc)));

            Assert.Contains("know-how protected", Log);
            Assert.Equal(new[] { @"PLC_1\scl\Program blocks\A.scl", @"PLC_1\scl\Program blocks\C.scl" }, ExportedFiles());
        }

        [Fact]
        public void Run_BlockGenerationAbortsHalfway_PartialTempFileIsRemoved()
        {
            var plc = _tia.Plc("PLC_1").WithBlocks(_tia.Scl("Broken").Failing("Inconsistent block", writePartialFile: true));

            Run(OneProject(_tia.S71500Station("S", "PLC_1", plc)));

            Assert.False(File.Exists(Path.Combine(_export, "Broken.scl")));
            Assert.Empty(ExportedFiles());
        }

        [Fact]
        public void Run_BlockGroupFails_ErrorLoggedAndSiblingGroupsExported()
        {
            var broken = _tia.Group("Broken").With(_tia.Scl("Lost"));
            broken.FailEnumeration = new FakeOpennessException("Object has been disposed");
            var plc = _tia.Plc("PLC_1").WithGroups(broken, _tia.Group("Fine").With(_tia.Scl("Kept")));

            Run(OneProject(_tia.S71500Station("S", "PLC_1", plc)));

            Assert.Contains("Error handling block group: Broken - Object has been disposed", Log);
            Assert.Equal(new[] { @"PLC_1\scl\Program blocks\Fine\Kept.scl" }, ExportedFiles());
        }

        [Fact]
        public void Run_TypeGenerationFails_ErrorLoggedAndOtherTypesExported()
        {
            var plc = _tia.Plc("PLC_1").WithTypes(_tia.Udt("UDT_Bad").Failing("UDT is inconsistent"), _tia.Udt("UDT_Good"));

            Run(OneProject(_tia.S71500Station("S", "PLC_1", plc)));

            Assert.Contains("UDT is inconsistent", Log);
            Assert.Equal(new[] { @"PLC_1\udt\UDT_Good.udt" }, ExportedFiles());
        }

        [Fact]
        public void Run_PlcTagTableExportFails_ErrorLoggedAndRunContinues()
        {
            // Before: an exception here was not caught and aborted the whole program.
            var plc = _tia.Plc("PLC_1")
                .WithTagTables(_tia.PlcTagTable("Bad").Failing("Export failed"), _tia.PlcTagTable("Good"));
            var hmi = _tia.HmiStation("H", "HMI_1", _tia.Hmi("HMI_1").WithTextLists(_tia.TextList("Texts")));

            Run(OneProject(_tia.S71500Station("S", "PLC_1", plc), hmi));

            Assert.Contains("Export failed", Log);
            Assert.Equal(new[]
            {
                @"HMI_1\hmi_text_lists\xml\Texts.xml",
                @"PLC_1\tag_tables\xml\Good.xml",
            }, ExportedFiles());
        }

        [Fact]
        public void Run_DeviceItemThrows_ErrorLoggedAndOtherItemsExported()
        {
            var bad = _tia.Item("PLC_Bad", null);
            bad.FailGetSoftware = new FakeOpennessException("Access denied");
            var station = _tia.Device("S", FakeTia.S71500, bad, _tia.Item("PLC_Good", _tia.Plc("PLC_Good").WithBlocks(_tia.Scl("Main"))));

            Run(OneProject(station));

            Assert.Contains("Error handling device item [PLC_Bad]", Log);
            Assert.Equal(new[] { @"PLC_Good\scl\Program blocks\Main.scl" }, ExportedFiles());
        }

        // ---------------- PLC: name clashes and odd names ----------------

        [Fact]
        public void Run_SameBlockNameInTwoPlcs_BothExported()
        {
            // The temp file name is reused, and TIA refuses to overwrite: the exporter must clean up between blocks.
            var plc1 = _tia.Plc("PLC_1").WithBlocks(_tia.Scl("Main"));
            var plc2 = _tia.Plc("PLC_2").WithBlocks(_tia.Scl("Main"));

            Run(OneProject(_tia.S71500Station("S1", "PLC_1", plc1), _tia.S71500Station("S2", "PLC_2", plc2)));

            Assert.Equal(new[] { @"PLC_1\scl\Program blocks\Main.scl", @"PLC_2\scl\Program blocks\Main.scl" }, ExportedFiles());
        }

        [Fact]
        public void Run_SameBlockNameInTwoGroups_BothExported()
        {
            var plc = _tia.Plc("PLC_1").WithGroups(
                _tia.Group("Line1").With(_tia.Scl("FB_Conveyor")),
                _tia.Group("Line2").With(_tia.Scl("FB_Conveyor")));

            Run(OneProject(_tia.S71500Station("S", "PLC_1", plc)));

            Assert.Equal(new[]
            {
                @"PLC_1\scl\Program blocks\Line1\FB_Conveyor.scl",
                @"PLC_1\scl\Program blocks\Line2\FB_Conveyor.scl",
            }, ExportedFiles());
        }

        [Fact]
        public void Run_StaleTempFileFromEarlierCrash_DoesNotBlockExport()
        {
            var exporter = NewExporter();
            exporter.PrepareExportLocation();
            File.WriteAllText(Path.Combine(_export, "Main.scl"), "left over");

            exporter.Run(new[] { OneProject(_tia.S71500Station("S", "PLC_1", _tia.Plc("PLC_1").WithBlocks(_tia.Scl("Main")))) });

            Assert.Equal(new[] { @"PLC_1\scl\Program blocks\Main.scl" }, ExportedFiles());
        }

        [Fact]
        public void Run_NamesWithInvalidCharacters_AreExportedUnderSanitizedNames()
        {
            var plc = _tia.Plc("PLC_1")
                .WithGroups(_tia.Group("In/Out").With(_tia.Scl("Motor:Pump")))
                .WithTypes(_tia.Udt("UDT/Valve"))
                .WithTagTables(_tia.PlcTagTable("IO|Map"));
            var hmi = _tia.Hmi("HMI_1")
                .WithTagTables(_tia.HmiTagTable("Tags?"))
                .WithTextLists(_tia.TextList("Status*Texts"));

            Run(OneProject(_tia.S71500Station("S", "PLC_1", plc), _tia.HmiStation("H", "HMI_1", hmi)));

            Assert.Equal(new[]
            {
                @"HMI_1\hmi_tag_tables\xml\Tags_.xml",
                @"HMI_1\hmi_text_lists\xml\Status_Texts.xml",
                @"PLC_1\scl\Program blocks\In_Out\Motor_Pump.scl",
                @"PLC_1\tag_tables\xml\IO_Map.xml",
                @"PLC_1\udt\UDT_Valve.udt",
            }, ExportedFiles());
        }

        // ---------------- HMI ----------------

        [Fact]
        public void Run_Hmi_ExportsTagTablesFromAllFoldersAndTextLists()
        {
            var hmi = _tia.Hmi("HMI_RT_1")
                .WithTagTables(_tia.HmiTagTable("Default tag table"))
                .WithTagFolders(_tia.TagFolder()
                    .With(_tia.HmiTagTable("Motors"))
                    .With(_tia.TagFolder().With(_tia.HmiTagTable("Valves"))))
                .WithTextLists(_tia.TextList("Alarm_Texts"), _tia.TextList("Modes"));

            Run(OneProject(_tia.HmiStation("HMI_Station", "HMI_RT_1", hmi)));

            Assert.Equal(new[]
            {
                @"HMI_RT_1\hmi_tag_tables\xml\Default tag table.xml",
                @"HMI_RT_1\hmi_tag_tables\xml\Motors.xml",
                @"HMI_RT_1\hmi_tag_tables\xml\Valves.xml",
                @"HMI_RT_1\hmi_text_lists\xml\Alarm_Texts.xml",
                @"HMI_RT_1\hmi_text_lists\xml\Modes.xml",
            }, ExportedFiles());
            Assert.Contains("<Name>Valves</Name>", Read(@"HMI_RT_1\hmi_tag_tables\xml\Valves.xml"));
        }

        [Fact]
        public void Run_Hmi_ExportOrderIsDeepestUserFoldersFirstThenSystemFolderThenTextLists()
        {
            var hmi = _tia.Hmi("HMI_1")
                .WithTagTables(_tia.HmiTagTable("System"))
                .WithTagFolders(_tia.TagFolder()
                    .With(_tia.HmiTagTable("Parent"))
                    .With(_tia.TagFolder().With(_tia.HmiTagTable("Child"))))
                .WithTextLists(_tia.TextList("Texts"));

            Run(OneProject(_tia.HmiStation("H", "HMI_1", hmi)));

            var exports = _tia.CallsStartingWith("Export ").Select(c => c.Split(' ')[1]).ToList();
            Assert.Equal(new[] { "Child", "Parent", "System", "Texts" }, exports);
        }

        [Fact]
        public void Run_Hmi_CreatesOutputFoldersBeforeExporting()
        {
            // The fake refuses to write into a missing folder (strict mode); the exporter must create it.
            _tia.RequireExistingDirectory = true;
            var hmi = _tia.Hmi("HMI_1").WithTagTables(_tia.HmiTagTable("Tags")).WithTextLists(_tia.TextList("Texts"));

            Run(OneProject(_tia.HmiStation("H", "HMI_1", hmi)));

            Assert.DoesNotContain("does not exist.", Log);
            Assert.Equal(2, ExportedFiles().Count);
        }

        [Fact]
        public void Run_TwoHmisWithSameTableNames_DoNotOverwriteEachOther()
        {
            var hmi1 = _tia.Hmi("HMI_1").WithTagTables(_tia.HmiTagTable("Tags")).WithTextLists(_tia.TextList("Texts"));
            var hmi2 = _tia.Hmi("HMI_2").WithTagTables(_tia.HmiTagTable("Tags")).WithTextLists(_tia.TextList("Texts"));

            Run(OneProject(_tia.HmiStation("H1", "HMI_1", hmi1), _tia.HmiStation("H2", "HMI_2", hmi2)));

            Assert.Equal(new[]
            {
                @"HMI_1\hmi_tag_tables\xml\Tags.xml",
                @"HMI_1\hmi_text_lists\xml\Texts.xml",
                @"HMI_2\hmi_tag_tables\xml\Tags.xml",
                @"HMI_2\hmi_text_lists\xml\Texts.xml",
            }, ExportedFiles());
        }

        [Fact]
        public void Run_HmiTagTableFails_OthersStillExported()
        {
            var hmi = _tia.Hmi("HMI_1").WithTagTables(_tia.HmiTagTable("Bad").Failing("Export failed"), _tia.HmiTagTable("Good"));

            Run(OneProject(_tia.HmiStation("H", "HMI_1", hmi)));

            Assert.Contains("Export failed", Log);
            Assert.Equal(new[] { @"HMI_1\hmi_tag_tables\xml\Good.xml" }, ExportedFiles());
        }

        [Fact]
        public void Run_NonS71500DeviceItemWithoutSoftware_LogsSwContainerNull()
        {
            Run(OneProject(_tia.Device("Switch", "System:Device.Scalance", _tia.Item("Port_1", null))));

            Assert.Contains("SW Container is null", Log);
            Assert.Empty(ExportedFiles());
        }

        // ---------------- repeatability ----------------

        [Fact]
        public void Run_Twice_GivesIdenticalExport()
        {
            Func<FakeProcess> project = () => OneProject(
                _tia.S71500Station("S", "PLC_1", _tia.Plc("PLC_1")
                    .WithBlocks(_tia.Scl("Main"))
                    .WithTypes(_tia.Udt("UDT1"))
                    .WithTagTables(_tia.PlcTagTable("Tags"))),
                _tia.HmiStation("H", "HMI_1", _tia.Hmi("HMI_1").WithTagTables(_tia.HmiTagTable("Tags"))));

            Run(project());
            var first = ExportedFiles().ToDictionary(f => f, Read);
            Run(project());
            var second = ExportedFiles().ToDictionary(f => f, Read);

            Assert.Equal(first, second);
            Assert.Equal(4, second.Count);
        }
    }
}
