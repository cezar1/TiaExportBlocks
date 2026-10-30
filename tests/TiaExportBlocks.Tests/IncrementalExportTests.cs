using System;
using System.IO;
using System.Linq;
using TiaExportBlocks.Tests.Fakes;
using Xunit;

namespace TiaExportBlocks.Tests
{
    /// <summary>
    /// --incremental: several export runs against the same fake project, changing it in between the way a
    /// user would in TIA (or on disk), and checking what gets regenerated, kept and removed.
    /// </summary>
    public class IncrementalExportTests : IDisposable
    {
        private readonly ExportHarness _h = new ExportHarness();
        private FakeTia Tia { get { return _h.Tia; } }

        // A small project: PLC_1 with Main (SCL), Data\DB_Recipe (DB), UDT_Motor, two tag tables; HMI_1 with a tag table and a text list.
        private readonly FakeBlock _main;
        private readonly FakeBlock _recipe;
        private readonly FakeBlockGroup _data;
        private readonly FakeType _udt;
        private readonly FakeExportable _tags;
        private readonly FakePlcSoftware _plc;
        private readonly FakeHmiTarget _hmi;
        private readonly FakeExportable _alarms;
        private string _plcItemName = "PLC_1";

        private const string MainFile = @"PLC_1\scl\Program blocks\Main.scl";
        private const string RecipeFile = @"PLC_1\db\Program blocks\Data\DB_Recipe.db";
        private const string UdtFile = @"PLC_1\udt\UDT_Motor.udt";

        public IncrementalExportTests()
        {
            _main = Tia.Scl("Main");
            _recipe = Tia.Db("DB_Recipe");
            _data = Tia.Group("Data").With(_recipe);
            _udt = Tia.Udt("UDT_Motor");
            _tags = Tia.PlcTagTable("Default tag table");
            _plc = Tia.Plc("PLC_1").WithBlocks(_main).WithGroups(_data).WithTypes(_udt)
                .WithTagTables(_tags).WithTagGroups(Tia.TagGroup().With(Tia.PlcTagTable("IO")));
            _alarms = Tia.TextList("Alarms");
            _hmi = Tia.Hmi("HMI_1").WithTagTables(Tia.HmiTagTable("HMI tags")).WithTextLists(_alarms);
        }

        public void Dispose()
        {
            _h.Dispose();
        }

        private FakeProcess Project()
        {
            return Tia.Process(1, Tia.Project("Line3",
                Tia.S71500Station("Station_1", _plcItemName, _plc),
                Tia.HmiStation("HMI_Station", "HMI_1", _hmi)));
        }

        private string[] AllFiles
        {
            get
            {
                return new[]
                {
                    @"HMI_1\hmi_tag_tables\xml\HMI tags.xml",
                    @"HMI_1\hmi_text_lists\xml\Alarms.xml",
                    RecipeFile,
                    MainFile,
                    @"PLC_1\tag_tables\xml\Default tag table.xml",
                    @"PLC_1\tag_tables\xml\IO.xml",
                    UdtFile,
                }.OrderBy(f => f, StringComparer.Ordinal).ToArray();
            }
        }

        // ---------------- basics ----------------

        [Fact]
        public void FirstIncrementalRun_WithoutState_ExportsEverything()
        {
            _h.Incremental(Project());

            Assert.Equal(AllFiles, _h.ExportedFiles());
            Assert.Equal(new[] { "UDT_Motor", "Main", "DB_Recipe" }, _h.Generated());
        }

        [Fact]
        public void SecondRun_NothingChanged_GeneratesNoSources()
        {
            _h.Incremental(Project());

            _h.Incremental(Project());

            Assert.Empty(_h.Generated());
            Assert.Equal(AllFiles, _h.ExportedFiles());
            var summary = _h.Summary("Line3");
            Assert.Equal(7, summary.Unchanged); // 3 blocks/types + 4 XML exports identical apart from their export time
            Assert.Equal(0, summary.Removed);
            Assert.Equal(0, summary.Exported);
        }

        // ---------------- regression (0.2): XML export time made every incremental run export everything ----------------

        [Fact]
        public void XmlExportTimeAlone_DoesNotCountAsTagTableChange()
        {
            _h.Incremental(Project());

            _h.Incremental(Project());

            Assert.DoesNotContain("tag tables of PLC_1 changed", _h.Log);
            Assert.Contains("Incremental: PLC_1: 0 of 3 blocks/types to export, 3 unchanged", _h.Log);
            Assert.Empty(_h.Generated());
        }

        [Fact]
        public void UnchangedXml_IsLeftUntouchedOnDisk()
        {
            _h.Incremental(Project());
            string file = _h.PathOf(@"PLC_1\tag_tables\xml\Default tag table.xml");
            string before = File.ReadAllText(file);
            var stamp = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            File.SetLastWriteTimeUtc(file, stamp);

            _h.Incremental(Project());

            Assert.Equal(before, File.ReadAllText(file)); // old <Created> kept: no diff in git
            Assert.Equal(stamp, File.GetLastWriteTimeUtc(file));
        }

        [Fact]
        public void ChangedXml_IsReplaced_AndNamedInTheLog()
        {
            _h.Incremental(Project());
            _tags.Content = "<Tag>Renamed</Tag>";

            _h.Incremental(Project());

            Assert.Contains("<Tag>Renamed</Tag>", _h.Read(@"PLC_1\tag_tables\xml\Default tag table.xml"));
            Assert.Contains("Incremental: tag tables of PLC_1 changed since the last export (Default tag table)", _h.Log);
        }

        [Fact]
        public void Log_SaysHowManyBlocksAreExportedAndWhy()
        {
            _h.Incremental(Project());
            _main.Edit("// edited");

            _h.Incremental(Project());

            Assert.Contains("Incremental: PLC_1: 1 of 3 blocks/types to export, 2 unchanged", _h.Log);
        }

        [Fact]
        public void Log_NamesAddedAndRemovedElements()
        {
            _h.Incremental(Project());
            _data.Remove(_recipe);
            _data.With(Tia.Db("DB_Recipes"));

            _h.Incremental(Project());

            Assert.Contains("(new: DB_Recipes.db) (gone: DB_Recipe.db)", _h.Log);
        }

        [Fact]
        public void Log_WarnsAboutUnreadableModifiedDates()
        {
            _main.ModifiedDate = null;

            _h.Incremental(Project());

            Assert.Contains("Incremental: 1 blocks/types of PLC_1 have no readable ModifiedDate and are always exported", _h.Log);
        }

        [Fact]
        public void FullExport_ThenIncremental_SkipsEverything()
        {
            _h.Full(Project());

            _h.Incremental(Project());

            Assert.Empty(_h.Generated());
        }

        [Fact]
        public void IncrementalRun_KeepsTheExportFolder()
        {
            _h.Incremental(Project());
            File.WriteAllText(_h.PathOf("README.md"), "mine");
            Directory.CreateDirectory(_h.PathOf(".git"));
            File.WriteAllText(_h.PathOf(@".git\HEAD"), "ref: refs/heads/main");

            _h.Incremental(Project());

            Assert.Equal("mine", _h.Read("README.md"));
            Assert.Equal("ref: refs/heads/main", _h.Read(@".git\HEAD"));
        }

        [Fact]
        public void FullRun_StillWipesEverything()
        {
            _h.Incremental(Project());
            File.WriteAllText(_h.PathOf("README.md"), "mine");

            _h.Full(Project());

            Assert.Equal(AllFiles, _h.ExportedFiles());
        }

        // ---------------- what triggers a regeneration ----------------

        [Fact]
        public void EditedBlock_OnlyThatBlockIsRegenerated()
        {
            _h.Incremental(Project());
            _main.Edit("#Enable := NOT #Enable;");

            _h.Incremental(Project());

            Assert.Equal(new[] { "Main" }, _h.Generated());
            Assert.Contains("#Enable := NOT #Enable;", _h.Read(MainFile));
        }

        [Fact]
        public void EditedType_OnlyThatTypeIsRegenerated()
        {
            _h.Incremental(Project());
            _udt.Edit("Speed : LReal;");

            _h.Incremental(Project());

            Assert.Equal(new[] { "UDT_Motor" }, _h.Generated());
            Assert.Contains("Speed : LReal;", _h.Read(UdtFile));
        }

        [Fact]
        public void ModifiedDateGoingBackwards_IsAChangeToo()
        {
            // e.g. an older archived version of the project retrieved to the same path
            _h.Incremental(Project());
            _main.ModifiedDate = FakeTia.InitialDate.AddDays(-30);
            _main.Body = "// older version";

            _h.Incremental(Project());

            Assert.Equal(new[] { "Main" }, _h.Generated());
            Assert.Contains("// older version", _h.Read(MainFile));
        }

        [Fact]
        public void UnreadableModifiedDate_IsAlwaysRegenerated()
        {
            _main.ModifiedDate = null;
            _h.Incremental(Project());

            _h.Incremental(Project());

            Assert.Equal(new[] { "Main" }, _h.Generated());
        }

        [Fact]
        public void FileDeletedOnDisk_IsRegenerated()
        {
            _h.Incremental(Project());
            File.Delete(_h.PathOf(MainFile));

            _h.Incremental(Project());

            Assert.Equal(new[] { "Main" }, _h.Generated());
            Assert.True(File.Exists(_h.PathOf(MainFile)));
        }

        [Fact]
        public void FileChangedOnDisk_IsRegenerated()
        {
            // e.g. a git branch switch or a manual edit in the export folder
            _h.Incremental(Project());
            File.WriteAllText(_h.PathOf(MainFile), "edited by hand");

            _h.Incremental(Project());

            Assert.Equal(new[] { "Main" }, _h.Generated());
            Assert.Equal(_main.ExpectedSource, _h.Read(MainFile));
        }

        [Fact]
        public void FileOnlyTouchedOnDisk_IsNotRegenerated()
        {
            // Sync tools, antivirus, copies and git checkouts change write times without changing content.
            _h.Incremental(Project());
            File.SetLastWriteTimeUtc(_h.PathOf(MainFile), DateTime.UtcNow.AddHours(3));
            File.SetLastWriteTimeUtc(_h.PathOf(UdtFile), new DateTime(2001, 1, 1, 0, 0, 0, DateTimeKind.Utc));

            _h.Incremental(Project());

            Assert.Empty(_h.Generated());
        }

        // ---------------- review finding 1: failed export must not be skipped forever ----------------

        [Fact]
        public void FailedExport_KeepsPreviousFile_AndIsRetriedOnNextRun()
        {
            _h.Incremental(Project());
            string before = _h.Read(MainFile);
            _main.Edit("#Enable := TRUE;");
            _main.Failing("Block 'Main' is inconsistent");

            _h.Incremental(Project());

            Assert.Equal(before, _h.Read(MainFile));
            Assert.Contains("Kept the previous export", _h.Log);
            Assert.Equal(1, _h.Summary("Line3").Failed);

            _main.Fail = null; // fixed in TIA without touching the block again
            _h.Incremental(Project());

            Assert.Equal(new[] { "Main" }, _h.Generated());
            Assert.Contains("#Enable := TRUE;", _h.Read(MainFile));
        }

        [Fact]
        public void FailedExportOfNewBlock_LeavesNoFile_AndIsRetried()
        {
            _h.Incremental(Project());
            var added = Tia.Scl("Added").Failing("inconsistent");
            _plc.Root.With(added);

            _h.Incremental(Project());
            Assert.False(File.Exists(_h.PathOf(@"PLC_1\scl\Program blocks\Added.scl")));

            added.Fail = null;
            _h.Incremental(Project());
            Assert.True(File.Exists(_h.PathOf(@"PLC_1\scl\Program blocks\Added.scl")));
        }

        // ---------------- review finding 2: changes in other elements ----------------

        [Fact]
        public void RenamedTag_ChangesTagTable_AllBlocksAndTypesRegenerated()
        {
            // TIA shows the new tag name in Main's source although Main's ModifiedDate does not change.
            _h.Incremental(Project());
            _tags.Content = "<Tag>StartButton</Tag>";
            _main.Body = "\"StartButton\" := TRUE;";

            _h.Incremental(Project());

            Assert.Contains("Incremental: tag tables of PLC_1 changed", _h.Log);
            Assert.Equal(new[] { "UDT_Motor", "Main", "DB_Recipe" }, _h.Generated());
            Assert.Contains("\"StartButton\" := TRUE;", _h.Read(MainFile));
        }

        [Fact]
        public void RenamedBlock_AllBlocksRegenerated_AndOldFileRemoved()
        {
            // Callers of a renamed FB show the new name in their source with unchanged ModifiedDate.
            _h.Incremental(Project());
            _data.Remove(_recipe);
            _data.With(Tia.Db("DB_Recipes"));

            _h.Incremental(Project());

            Assert.Contains("added, removed or renamed", _h.Log);
            Assert.Equal(new[] { "UDT_Motor", "Main", "DB_Recipes" }, _h.Generated());
            Assert.False(File.Exists(_h.PathOf(RecipeFile)));
            Assert.True(File.Exists(_h.PathOf(@"PLC_1\db\Program blocks\Data\DB_Recipes.db")));
            Assert.Equal(1, _h.Summary("Line3").Removed);
        }

        // ---------------- cleanup ----------------

        [Fact]
        public void DeletedBlock_FileAndEmptyFolderAreRemoved()
        {
            _h.Incremental(Project());
            _data.Remove(_recipe);

            _h.Incremental(Project());

            Assert.False(File.Exists(_h.PathOf(RecipeFile)));
            Assert.False(Directory.Exists(_h.PathOf(@"PLC_1\db")));
            Assert.Contains("Removing " + _h.PathOf(RecipeFile) + " (no longer in the project)", _h.Log);
        }

        [Fact]
        public void DeletedType_FileIsRemoved()
        {
            _h.Incremental(Project());
            _plc.TypeList.Clear();

            _h.Incremental(Project());

            Assert.False(File.Exists(_h.PathOf(UdtFile)));
        }

        [Fact]
        public void DeletedHmiTextList_FileIsRemoved()
        {
            _h.Incremental(Project());
            _hmi.RemoveTextList(_alarms);

            _h.Incremental(Project());

            Assert.False(File.Exists(_h.PathOf(@"HMI_1\hmi_text_lists\xml\Alarms.xml")));
            Assert.True(File.Exists(_h.PathOf(@"HMI_1\hmi_tag_tables\xml\HMI tags.xml")));
        }

        [Fact]
        public void RenamedPlc_OldPlcFolderIsRemovedCompletely()
        {
            // review finding 5: folders of PLCs no longer in the project were never cleaned
            _h.Incremental(Project());
            _plcItemName = "PLC_Main";

            _h.Incremental(Project());

            Assert.False(Directory.Exists(_h.PathOf("PLC_1")));
            Assert.True(File.Exists(_h.PathOf(@"PLC_Main\scl\Program blocks\Main.scl")));
        }

        [Fact]
        public void FilesNotWrittenByTheTool_AreNeverRemoved()
        {
            _h.Incremental(Project());
            File.WriteAllText(_h.PathOf(@"PLC_1\scl\Program blocks\Handwritten.scl"), "mine");
            File.WriteAllText(_h.PathOf(@"PLC_1\tag_tables\xml\notes.xml"), "mine");
            _data.Remove(_recipe);

            _h.Incremental(Project());

            Assert.Equal("mine", _h.Read(@"PLC_1\scl\Program blocks\Handwritten.scl"));
            Assert.Equal("mine", _h.Read(@"PLC_1\tag_tables\xml\notes.xml"));
        }

        [Fact]
        public void ReadErrorInTia_NothingIsRemoved_AndFilesAreKeptForNextRun()
        {
            _h.Incremental(Project());
            _data.FailEnumeration = new FakeOpennessException("Object has been disposed");

            _h.Incremental(Project());

            Assert.True(File.Exists(_h.PathOf(RecipeFile)));
            Assert.Contains("Not removing any files", _h.Log);

            _data.FailEnumeration = null;
            _h.Incremental(Project());

            Assert.True(File.Exists(_h.PathOf(RecipeFile)));
            Assert.DoesNotContain("Removing", _h.Log);
        }

        // ---------------- review finding 4: several projects in one export location ----------------

        [Fact]
        public void TwoProjects_SameLocation_DoNotRemoveEachOthersFiles()
        {
            var a = Tia.Plc("PLC_A").WithBlocks(Tia.Scl("MainA"));
            var aBlock2 = Tia.Scl("ExtraA");
            a.Root.With(aBlock2);
            var b = Tia.Plc("PLC_B").WithBlocks(Tia.Scl("MainB"));
            Func<FakeProcess> processes = () => Tia.Process(1,
                Tia.Project("A", Tia.S71500Station("SA", "PLC_A", a), Tia.HmiStation("HA", "HMI_A", Tia.Hmi("HMI_A").WithTagTables(Tia.HmiTagTable("Tags")))),
                Tia.Project("B", Tia.S71500Station("SB", "PLC_B", b), Tia.HmiStation("HB", "HMI_B", Tia.Hmi("HMI_B").WithTagTables(Tia.HmiTagTable("Tags")))));

            _h.Incremental(processes());
            _h.Incremental(processes());

            Assert.Empty(_h.Generated());
            Assert.Equal(5, _h.ExportedFiles().Count);

            a.Root.Remove(aBlock2);
            _h.Incremental(processes());

            Assert.False(File.Exists(_h.PathOf(@"PLC_A\scl\Program blocks\ExtraA.scl")));
            Assert.Equal(4, _h.ExportedFiles().Count);
            Assert.True(File.Exists(_h.PathOf(@"HMI_A\hmi_tag_tables\xml\Tags.xml")));
            Assert.True(File.Exists(_h.PathOf(@"HMI_B\hmi_tag_tables\xml\Tags.xml")));
        }

        [Fact]
        public void TwoProjects_WritingTheSameFile_ItIsKeptWhileOneStillHasIt()
        {
            var mainA = Tia.Scl("Main");
            var a = Tia.Plc("PLC_1").WithBlocks(mainA);
            var b = Tia.Plc("PLC_1").WithBlocks(Tia.Scl("Main"));
            Func<FakeProcess> processes = () => Tia.Process(1,
                Tia.Project("A", Tia.S71500Station("SA", "PLC_1", a)),
                Tia.Project("B", Tia.S71500Station("SB", "PLC_1", b)));
            _h.Incremental(processes());

            a.Root.Remove(mainA);
            _h.Incremental(processes());

            Assert.True(File.Exists(_h.PathOf(MainFile)));
        }

        // ---------------- state robustness ----------------

        [Fact]
        public void CorruptState_ExportsEverything()
        {
            _h.Incremental(Project());
            string stateFile = _h.LastExporter.StateFile;
            File.WriteAllText(stateFile, "garbage");

            _h.Incremental(Project());

            Assert.Contains("unknown format: exporting everything", _h.Log);
            Assert.Equal(new[] { "UDT_Motor", "Main", "DB_Recipe" }, _h.Generated());
        }

        [Fact]
        public void State_IsKeptOutsideTheExportLocation()
        {
            _h.Incremental(Project());

            Assert.StartsWith(_h.StateDir, _h.LastExporter.StateFile);
            Assert.True(File.Exists(_h.LastExporter.StateFile));
            Assert.Equal(AllFiles, _h.ExportedFiles());
        }

        [Fact]
        public void SummaryCounters_ArePerProject()
        {
            var p1 = Tia.Plc("PLC_A").WithBlocks(Tia.Scl("A1"), Tia.Scl("A2"));
            var p2 = Tia.Plc("PLC_B").WithBlocks(Tia.Scl("B1"));

            _h.Incremental(Tia.Process(1,
                Tia.Project("A", Tia.S71500Station("SA", "PLC_A", p1)),
                Tia.Project("B", Tia.S71500Station("SB", "PLC_B", p2))));

            Assert.Equal(2, _h.Summary("A").Exported);
            Assert.Equal(1, _h.Summary("B").Exported);
        }
    }
}
