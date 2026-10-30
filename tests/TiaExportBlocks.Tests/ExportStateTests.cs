using System;
using System.IO;
using System.Linq;
using Xunit;

namespace TiaExportBlocks.Tests
{
    public class ExportStateTests : IDisposable
    {
        private readonly string _dir = Path.Combine(Path.GetTempPath(), "TiaStateTests_" + Guid.NewGuid().ToString("N"));
        private readonly StringWriter _log = new StringWriter();

        public void Dispose()
        {
            if (Directory.Exists(_dir)) Directory.Delete(_dir, true);
        }

        private string StateFile { get { return Path.Combine(_dir, "x.state"); } }

        [Fact]
        public void SaveAndLoad_RoundTripsSeveralProjects()
        {
            var state = new ExportState();
            state.SetProject(@"C:\P\A.ap15_1", new System.Collections.Generic.Dictionary<string, StateEntry>
            {
                { @"C:\E\PLC_1\scl\Main.scl", new StateEntry(638000000000000000, "ABC") },
                { @"C:\E\PLC_1\tag_tables\xml\Tags.xml", new StateEntry(StateEntry.NoStamp, "DEF") },
            });
            state.SetProject(@"C:\P\B.ap15_1", new System.Collections.Generic.Dictionary<string, StateEntry>
            {
                { @"C:\E\PLC_2\udt\U.udt", new StateEntry(5, null) },
            });

            state.Save(StateFile);
            var loaded = ExportState.Load(StateFile, _log);

            Assert.Equal(2, loaded.ProjectKeys.Count());
            var a = loaded.GetProject(@"c:\p\a.ap15_1");
            Assert.Equal(638000000000000000, a[@"C:\E\PLC_1\scl\Main.scl"].Stamp);
            Assert.Equal("ABC", a[@"c:\e\plc_1\SCL\main.scl"].Hash);
            Assert.Equal(StateEntry.NoStamp, a[@"C:\E\PLC_1\tag_tables\xml\Tags.xml"].Stamp);
            Assert.Null(loaded.GetProject(@"C:\P\B.ap15_1")[@"C:\E\PLC_2\udt\U.udt"].Hash);
            Assert.Empty(_log.ToString());
        }

        [Fact]
        public void Save_LeavesNoTemporaryFile()
        {
            new ExportState().Save(StateFile);
            new ExportState().Save(StateFile);

            Assert.Equal(new[] { "x.state" }, Directory.GetFiles(_dir).Select(Path.GetFileName));
        }

        [Fact]
        public void Load_MissingFile_GivesEmptyStateSilently()
        {
            Assert.Empty(ExportState.Load(StateFile, _log).ProjectKeys);
            Assert.Empty(_log.ToString());
        }

        [Fact]
        public void Load_OldFormatFromPr2_GivesEmptyStateWithMessage()
        {
            Directory.CreateDirectory(_dir);
            File.WriteAllLines(StateFile, new[] { @"C:\P\A.ap15_1", "123\t456\tC:\\E\\x.scl" });

            Assert.Empty(ExportState.Load(StateFile, _log).ProjectKeys);
            Assert.Contains("unknown format", _log.ToString());
        }

        [Fact]
        public void Load_CorruptLine_GivesEmptyStateWithMessage()
        {
            Directory.CreateDirectory(_dir);
            File.WriteAllLines(StateFile, new[] { "TiaExportBlocks state v2", "P\tA", "E\tnot-a-number\tH\tC:\\x.scl" });

            Assert.Empty(ExportState.Load(StateFile, _log).ProjectKeys);
            Assert.Contains("could not be read", _log.ToString());
        }

        [Fact]
        public void GetProject_ReturnsACopy()
        {
            var state = new ExportState();
            state.SetProject("A", new System.Collections.Generic.Dictionary<string, StateEntry> { { "f", new StateEntry(1, "h") } });

            state.GetProject("A").Clear();

            Assert.Single(state.GetProject("A"));
            Assert.Empty(state.GetProject("unknown"));
        }

        [Fact]
        public void IsClaimedByOtherProject()
        {
            var state = new ExportState();
            state.SetProject("A", new System.Collections.Generic.Dictionary<string, StateEntry> { { @"C:\E\x.scl", new StateEntry(1, "h") } });
            state.SetProject("B", new System.Collections.Generic.Dictionary<string, StateEntry> { { @"C:\E\y.scl", new StateEntry(1, "h") } });

            Assert.True(state.IsClaimedByOtherProject(@"c:\e\X.SCL", "B"));
            Assert.False(state.IsClaimedByOtherProject(@"C:\E\x.scl", "A"));
            Assert.False(state.IsClaimedByOtherProject(@"C:\E\z.scl", "A"));
        }

        [Theory]
        [InlineData(@"C:\Export", @"c:\export")]
        [InlineData(@"C:\Export", @"C:\Export\")]
        [InlineData(@"C:\Export", @"C:\Export\.\")]
        public void FileFor_SameLocationWrittenDifferently_SameFile(string a, string b)
        {
            Assert.Equal(ExportState.FileFor(_dir, a), ExportState.FileFor(_dir, b));
        }

        [Fact]
        public void FileFor_DifferentLocations_DifferentFiles()
        {
            Assert.NotEqual(ExportState.FileFor(_dir, @"C:\Export1"), ExportState.FileFor(_dir, @"C:\Export2"));
            Assert.StartsWith(_dir, ExportState.FileFor(_dir, @"C:\Export1"));
        }
    }
}
