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
            // Junction-safe: removes links without following them (Directory.Delete(recursive) can fail on them).
            if (!Directory.Exists(_root)) return;
            FileOps.WipeDirectoryContents(_root, attempts: 3, retryDelayMs: 50);
            Directory.Delete(_root, false);
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
            // Doubled separators can come from user-supplied export locations with a trailing '\'
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

        // ---- WipeDirectoryContents ----

        [Fact]
        public void Wipe_MissingFolder_ReturnsNothingLeft()
        {
            Assert.Empty(FileOps.WipeDirectoryContents(In("missing")));
        }

        [Fact]
        public void Wipe_RemovesEverythingButKeepsTheFolder()
        {
            string export = In("export");
            Write(Path.Combine(export, "PLC_1", "scl", "Program blocks", "Main.scl"), "x");
            Write(Path.Combine(export, "HMI_1", "hmi_tag_tables", "xml", "Tags.xml"), "x");
            Write(Path.Combine(export, "notes.txt"), "x");
            Directory.CreateDirectory(Path.Combine(export, "empty", "nested"));

            Assert.Empty(FileOps.WipeDirectoryContents(export));
            Assert.True(Directory.Exists(export));
            Assert.Empty(Directory.GetFileSystemEntries(export));
        }

        [Fact]
        public void Wipe_RemovesReadOnlyHiddenAndSystemEntries()
        {
            string export = In("export");
            string file = Path.Combine(export, "a", "ro.scl");
            Write(file, "x");
            File.SetAttributes(file, FileAttributes.ReadOnly | FileAttributes.Hidden | FileAttributes.System);
            var dir = new DirectoryInfo(Path.Combine(export, "a"));
            dir.Attributes |= FileAttributes.ReadOnly | FileAttributes.Hidden;

            Assert.Empty(FileOps.WipeDirectoryContents(export));
            Assert.Empty(Directory.GetFileSystemEntries(export));
        }

        [Fact]
        public void Wipe_LockedFile_IsReportedAndOthersAreDeleted()
        {
            string export = In("export");
            string locked = Path.Combine(export, "locked.scl");
            Write(locked, "x");
            Write(Path.Combine(export, "free.scl"), "x");

            using (new FileStream(locked, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                var remaining = FileOps.WipeDirectoryContents(export, attempts: 2, retryDelayMs: 10);

                Assert.Equal(new[] { locked }, remaining);
            }
            Assert.False(File.Exists(Path.Combine(export, "free.scl")));
        }

        [Fact]
        public void Wipe_FileUnlockedWhileRetrying_IsDeleted()
        {
            string export = In("export");
            string locked = Path.Combine(export, "busy.scl");
            Write(locked, "x");
            var stream = new FileStream(locked, FileMode.Open, FileAccess.Read, FileShare.None);
            // Dedicated thread (not the thread pool, which can be busy with parallel tests) releases the lock
            // shortly after the first delete attempt; the retry window is generous.
            var release = new System.Threading.Thread(() => { System.Threading.Thread.Sleep(100); stream.Dispose(); });
            release.Start();

            var remaining = FileOps.WipeDirectoryContents(export, attempts: 100, retryDelayMs: 50);

            release.Join();
            Assert.Empty(remaining);
            Assert.Empty(Directory.GetFileSystemEntries(export));
        }

        [Fact]
        public void Wipe_Junction_RemovesTheLinkButNeverTheTarget()
        {
            string export = In("export");
            string outside = In("outside");
            Write(Path.Combine(outside, "precious.txt"), "keep me");
            Directory.CreateDirectory(export);
            string link = Path.Combine(export, "link");
            CreateJunction(link, outside);

            Assert.Empty(FileOps.WipeDirectoryContents(export));
            Assert.False(Directory.Exists(link));
            Assert.Equal("keep me", File.ReadAllText(Path.Combine(outside, "precious.txt")));
        }

        [Fact]
        public void ListEntries_DoesNotFollowJunctions()
        {
            string export = In("export");
            string outside = In("outside");
            Write(Path.Combine(outside, "precious.txt"), "x");
            Write(Path.Combine(export, "a.scl"), "x");
            string link = Path.Combine(export, "link");
            CreateJunction(link, outside);

            var entries = FileOps.ListEntries(export);

            Assert.Equal(2, entries.Count);
            Assert.Contains(link, entries);
        }

        private static void CreateJunction(string link, string target)
        {
            var psi = new System.Diagnostics.ProcessStartInfo("cmd.exe", "/c mklink /J \"" + link + "\" \"" + target + "\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            using (var p = System.Diagnostics.Process.Start(psi))
            {
                p.WaitForExit();
                Assert.Equal(0, p.ExitCode);
            }
            Assert.True((File.GetAttributes(link) & FileAttributes.ReparsePoint) != 0);
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
            var temp = new FileInfo(ExportPaths.BlockTempFile(export, "FB_Belt", extension));
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
