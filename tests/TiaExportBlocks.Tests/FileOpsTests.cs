using System;
using System.IO;
using Xunit;

namespace TiaExportBlocks.Tests
{
    /// <summary>Each test gets its own empty temp folder, deleted afterwards.</summary>
    public class FileOpsTests : IDisposable
    {
        private readonly string _root;

        public FileOpsTests()
        {
            _root = Path.Combine(Path.GetTempPath(), "TiaExportBlocksTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
        }

        public void Dispose()
        {
            if (Directory.Exists(_root)) Directory.Delete(_root, true);
        }

        private string In(params string[] parts)
        {
            return Path.Combine(_root, Path.Combine(parts));
        }

        private static void Write(string path, string content)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, content);
        }

        // ---- EnsureDirectoryExists ----

        [Fact]
        public void EnsureDirectoryExists_CreatesMissingParentFolders()
        {
            string file = In("a", "b", "c", "file.xml");

            FileOps.EnsureDirectoryExists(file);

            Assert.True(Directory.Exists(In("a", "b", "c")));
            Assert.False(File.Exists(file));
        }

        [Fact]
        public void EnsureDirectoryExists_ExistingFolder_LeavesContentAlone()
        {
            string existing = In("a", "keep.txt");
            Write(existing, "keep");

            FileOps.EnsureDirectoryExists(In("a", "file.xml"));

            Assert.Equal("keep", File.ReadAllText(existing));
        }

        [Fact]
        public void EnsureDirectoryExists_FileNameWithoutFolder_DoesNothing()
        {
            FileOps.EnsureDirectoryExists("file.xml");
        }

        [Fact]
        public void EnsureDirectoryExists_HandlesDoubledSeparators()
        {
            // PlcTagTableFile produces "...\xml\\name.xml"
            string file = _root + @"\PLC_1\tag_tables\xml\" + @"\" + "Tags.xml";

            FileOps.EnsureDirectoryExists(file);

            Assert.True(Directory.Exists(In("PLC_1", "tag_tables", "xml")));
        }

        // ---- EnsureDirectory ----

        [Fact]
        public void EnsureDirectory_Missing_CreatesAndReturnsTrue()
        {
            string dir = In("export");

            Assert.True(FileOps.EnsureDirectory(dir));
            Assert.True(Directory.Exists(dir));
        }

        [Fact]
        public void EnsureDirectory_MissingNested_CreatesAllLevels()
        {
            string dir = In("x", "y", "z");

            Assert.True(FileOps.EnsureDirectory(dir));
            Assert.True(Directory.Exists(dir));
        }

        [Fact]
        public void EnsureDirectory_Existing_ReturnsFalseAndKeepsContent()
        {
            string dir = In("export");
            Write(Path.Combine(dir, "keep.txt"), "keep");

            Assert.False(FileOps.EnsureDirectory(dir));
            Assert.True(File.Exists(Path.Combine(dir, "keep.txt")));
        }

        // ---- RemoveDirectory ----

        [Fact]
        public void RemoveDirectory_Missing_ReturnsFalse()
        {
            Assert.False(FileOps.RemoveDirectory(In("nope")));
        }

        [Fact]
        public void RemoveDirectory_Empty_RemovesAndReturnsTrue()
        {
            string dir = In("empty");
            Directory.CreateDirectory(dir);

            Assert.True(FileOps.RemoveDirectory(dir));
            Assert.False(Directory.Exists(dir));
        }

        [Fact]
        public void RemoveDirectory_WithNestedContent_RemovesEverything()
        {
            string dir = In("export");
            Write(Path.Combine(dir, "PLC_1", "scl", "Main.scl"), "x");
            Write(Path.Combine(dir, "hmi_tag_tables", "xml", "Tags.xml"), "y");

            Assert.True(FileOps.RemoveDirectory(dir));
            Assert.False(Directory.Exists(dir));
        }

        [Fact]
        public void RemoveDirectory_DoesNotTouchSiblings()
        {
            string dir = In("export");
            string sibling = In("other", "keep.txt");
            Write(Path.Combine(dir, "a.txt"), "a");
            Write(sibling, "keep");

            FileOps.RemoveDirectory(dir);

            Assert.True(File.Exists(sibling));
        }

        // ---- DeleteIfExists ----

        [Fact]
        public void DeleteIfExists_Existing_DeletesAndReturnsTrue()
        {
            string file = In("f.txt");
            Write(file, "x");

            Assert.True(FileOps.DeleteIfExists(file));
            Assert.False(File.Exists(file));
        }

        [Fact]
        public void DeleteIfExists_Missing_ReturnsFalse()
        {
            Assert.False(FileOps.DeleteIfExists(In("missing.txt")));
        }

        [Fact]
        public void DeleteIfExists_MissingFolder_ReturnsFalse()
        {
            Assert.False(FileOps.DeleteIfExists(In("no", "such", "folder", "f.txt")));
        }

        // ---- MoveReplacing ----

        [Fact]
        public void MoveReplacing_DestinationMissing_MovesFile()
        {
            string src = In("tmp.scl");
            string dst = In("out", "Main.scl");
            Write(src, "FUNCTION_BLOCK Main");
            Directory.CreateDirectory(In("out"));

            FileOps.MoveReplacing(src, dst);

            Assert.False(File.Exists(src));
            Assert.Equal("FUNCTION_BLOCK Main", File.ReadAllText(dst));
        }

        [Fact]
        public void MoveReplacing_DestinationExists_IsOverwritten()
        {
            string src = In("tmp.scl");
            string dst = In("out", "Main.scl");
            Write(src, "new");
            Write(dst, "old");

            FileOps.MoveReplacing(src, dst);

            Assert.False(File.Exists(src));
            Assert.Equal("new", File.ReadAllText(dst));
        }

        [Fact]
        public void MoveReplacing_DestinationFolderMissing_IsCreated()
        {
            string src = In("tmp.xml");
            string dst = In("PLC_1", "tag_tables", "xml", "Tags.xml");
            Write(src, "<xml/>");

            FileOps.MoveReplacing(src, dst);

            Assert.Equal("<xml/>", File.ReadAllText(dst));
        }

        [Fact]
        public void MoveReplacing_SourceMissing_ThrowsAndLeavesDestinationIntact()
        {
            string dst = In("out", "Main.scl");
            Write(dst, "old");

            Assert.Throws<FileNotFoundException>(() => FileOps.MoveReplacing(In("missing.scl"), dst));
            Assert.Equal("old", File.ReadAllText(dst));
        }

        [Fact]
        public void MoveReplacing_PreservesBinaryContent()
        {
            string src = In("tmp.bin");
            string dst = In("out.bin");
            var bytes = new byte[] { 0xEF, 0xBB, 0xBF, 0x00, 0xFF, 0x10, 0x0D, 0x0A };
            File.WriteAllBytes(src, bytes);

            FileOps.MoveReplacing(src, dst);

            Assert.Equal(bytes, File.ReadAllBytes(dst));
        }

        // ---- End-to-end layout (what Program does around each TIA export call) ----

        [Fact]
        public void BlockExportFlow_ProducesFileAtBlockPath()
        {
            string export = In("export");
            FileOps.RemoveDirectory(export);
            FileOps.EnsureDirectory(export);
            string extension, folder;
            ExportRules.TryGetBlockFormat("SCL", out extension, out folder);

            var target = new FileInfo(ExportPaths.BlockFile(export, "PLC_1", folder, @"Program blocks\Conveyors", "FB_Belt", extension));
            var temp = new FileInfo(ExportPaths.BlockTempFile(export + @"\", "FB_Belt", extension));
            Directory.CreateDirectory(target.Directory.FullName);
            FileOps.DeleteIfExists(target.FullName);
            File.WriteAllText(temp.FullName, "FUNCTION_BLOCK FB_Belt"); // stands in for GenerateSource
            FileOps.MoveReplacing(temp.FullName, target.FullName);

            Assert.Equal("FUNCTION_BLOCK FB_Belt",
                File.ReadAllText(Path.Combine(export, "PLC_1", "scl", "Program blocks", "Conveyors", "FB_Belt.scl")));
            Assert.False(File.Exists(temp.FullName));
        }

        [Fact]
        public void TagTableExportFlow_RepeatedRun_ReplacesPreviousExport()
        {
            string export = In("export");
            FileOps.EnsureDirectory(export);

            for (int run = 1; run <= 2; run++)
            {
                var temp = new FileInfo(ExportPaths.PlcTagTableTempFile(export, "Default tag table"));
                FileOps.DeleteIfExists(temp.FullName);
                File.WriteAllText(temp.FullName, "run " + run); // stands in for table.Export
                var target = new FileInfo(ExportPaths.PlcTagTableFile(export, "PLC_1", "Default tag table"));
                FileOps.MoveReplacing(temp.FullName, target.FullName);
            }

            Assert.Equal("run 2",
                File.ReadAllText(Path.Combine(export, "PLC_1", "tag_tables", "xml", "Default tag table.xml")));
            Assert.False(File.Exists(Path.Combine(export, "Default tag table.xml")));
        }
    }
}
