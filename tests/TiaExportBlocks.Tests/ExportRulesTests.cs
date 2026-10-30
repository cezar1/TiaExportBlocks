using System.Linq;
using Xunit;

namespace TiaExportBlocks.Tests
{
    public class ExportRulesTests
    {
        // ---- Block format (language -> extension / folder) ----

        [Theory]
        [InlineData("SCL", ".scl", @"\scl\")]
        [InlineData("DB", ".db", @"\db\")]
        [InlineData("STL", ".awl", @"\stl\")]
        public void TryGetBlockFormat_SupportedLanguage_ReturnsExtensionAndFolder(string language, string expectedExtension, string expectedFolder)
        {
            string extension, folder;

            bool supported = ExportRules.TryGetBlockFormat(language, out extension, out folder);

            Assert.True(supported);
            Assert.Equal(expectedExtension, extension);
            Assert.Equal(expectedFolder, folder);
        }

        [Theory]
        [InlineData("LAD")]
        [InlineData("FBD")]
        [InlineData("GRAPH")]
        [InlineData("CEM")]
        [InlineData("ProDiag")]
        [InlineData("")]
        [InlineData(" ")]
        public void TryGetBlockFormat_UnsupportedLanguage_ReturnsFalse(string language)
        {
            string extension, folder;

            Assert.False(ExportRules.TryGetBlockFormat(language, out extension, out folder));
            Assert.Null(extension);
            Assert.Null(folder);
        }

        [Theory]
        [InlineData("scl")]
        [InlineData("Scl")]
        [InlineData("db")]
        [InlineData("stl")]
        [InlineData(" SCL")]
        [InlineData("SCL ")]
        public void TryGetBlockFormat_IsCaseAndWhitespaceSensitive(string language)
        {
            string extension, folder;

            Assert.False(ExportRules.TryGetBlockFormat(language, out extension, out folder));
        }

        [Fact]
        public void TryGetBlockFormat_Null_ReturnsFalseWithoutThrowing()
        {
            string extension, folder;

            Assert.False(ExportRules.TryGetBlockFormat(null, out extension, out folder));
            Assert.Null(extension);
            Assert.Null(folder);
        }

        [Fact]
        public void LanguageTables_CoverTheSameLanguages()
        {
            var extensionKeys = ExportRules.LanguageToExtension.Keys.OrderBy(k => k);
            var folderKeys = ExportRules.LanguageToFolderPrefix.Keys.OrderBy(k => k);

            Assert.Equal(extensionKeys, folderKeys);
        }

        [Fact]
        public void LanguageTables_ExtensionsAreUniqueAndStartWithDot()
        {
            var extensions = ExportRules.LanguageToExtension.Values.ToList();

            Assert.Equal(extensions.Count, extensions.Distinct().Count());
            Assert.All(extensions, e => Assert.StartsWith(".", e));
        }

        [Fact]
        public void LanguageTables_FoldersAreUniqueAndWrappedInBackslashes()
        {
            var folders = ExportRules.LanguageToFolderPrefix.Values.ToList();

            Assert.Equal(folders.Count, folders.Distinct().Count());
            Assert.All(folders, f =>
            {
                Assert.StartsWith(@"\", f);
                Assert.EndsWith(@"\", f);
            });
        }

        // ---- Device classification ----

        [Fact]
        public void IsS71500_ExactTypeIdentifier_ReturnsTrue()
        {
            Assert.True(ExportRules.IsS71500("System:Device.S71500"));
            Assert.Equal("System:Device.S71500", ExportRules.S71500DeviceType);
        }

        [Theory]
        [InlineData("System:Device.S71200")]
        [InlineData("System:Device.S7300")]
        [InlineData("System:Device.S71500 ")]
        [InlineData("system:device.s71500")]
        [InlineData("OrderNumber:6AV2 124-0MC01-0AX0/16.0.0.0")]
        [InlineData("")]
        [InlineData(null)]
        public void IsS71500_AnythingElse_ReturnsFalse(string typeIdentifier)
        {
            Assert.False(ExportRules.IsS71500(typeIdentifier));
        }

        [Theory]
        [InlineData("PLC_1")]
        [InlineData("PLC")]
        [InlineData("Line3_PLC")]
        [InlineData("MyPLCController")]
        public void IsPlcDeviceItem_NameContainsPLC_ReturnsTrue(string name)
        {
            Assert.True(ExportRules.IsPlcDeviceItem(name));
        }

        [Theory]
        [InlineData("plc_1")]
        [InlineData("Plc_1")]
        [InlineData("CPU 1516-3 PN/DP")]
        [InlineData("Rail_0")]
        [InlineData("PROFINET interface_1")]
        [InlineData("")]
        [InlineData(null)]
        public void IsPlcDeviceItem_OtherNames_ReturnsFalse(string name)
        {
            Assert.False(ExportRules.IsPlcDeviceItem(name));
        }

        // ---- Export location safety ----

        [Theory]
        [InlineData(@"C:\")]
        [InlineData(@"D:\")]
        [InlineData(@"\\server\share")]
        [InlineData(@"\\server\share\")]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData(null)]
        [InlineData(@"C:\a|b")]
        public void TryNormalizeExportLocation_RootsEmptyOrInvalid_AreRefused(string path)
        {
            string normalized, error;

            Assert.False(ExportRules.TryNormalizeExportLocation(path, out normalized, out error));
            Assert.Null(normalized);
            Assert.False(string.IsNullOrEmpty(error));
        }

        [Theory]
        [InlineData(@"C:\Export", @"C:\Export")]
        [InlineData(@"C:\Export\", @"C:\Export")]
        [InlineData(@"C:\Export\\", @"C:\Export")]
        [InlineData(@"C:\Projects\Line 3\export", @"C:\Projects\Line 3\export")]
        [InlineData(@"\\server\share\export", @"\\server\share\export")]
        public void TryNormalizeExportLocation_Folders_AreAcceptedWithoutTrailingSeparator(string path, string expected)
        {
            string normalized, error;

            Assert.True(ExportRules.TryNormalizeExportLocation(path, out normalized, out error));
            Assert.Equal(expected, normalized);
            Assert.Null(error);
        }

        [Fact]
        public void TryNormalizeExportLocation_RelativePath_IsMadeAbsolute()
        {
            string normalized, error;

            Assert.True(ExportRules.TryNormalizeExportLocation("out", out normalized, out error));
            Assert.Equal(System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "out"), normalized);
        }

        // ---- Command line ----

        [Fact]
        public void TryParseArgs_SingleArgument_ReturnsExportLocation()
        {
            string location;

            Assert.True(ExportRules.TryParseArgs(new[] { @"C:\Export" }, out location));
            Assert.Equal(@"C:\Export", location);
        }

        [Fact]
        public void TryParseArgs_PathWithSpaces_IsTakenAsIs()
        {
            string location;

            Assert.True(ExportRules.TryParseArgs(new[] { @"C:\My Exports\Line 3" }, out location));
            Assert.Equal(@"C:\My Exports\Line 3", location);
        }

        [Fact]
        public void TryParseArgs_Null_ReturnsFalse()
        {
            string location;

            Assert.False(ExportRules.TryParseArgs(null, out location));
            Assert.Null(location);
        }

        [Fact]
        public void TryParseArgs_NoArguments_ReturnsFalse()
        {
            string location;

            Assert.False(ExportRules.TryParseArgs(new string[0], out location));
            Assert.Null(location);
        }

        [Theory]
        [InlineData(2)]
        [InlineData(3)]
        [InlineData(10)]
        public void TryParseArgs_MoreThanOneArgument_ReturnsFalse(int count)
        {
            string location;
            var args = Enumerable.Range(0, count).Select(i => @"C:\Export" + i).ToArray();

            Assert.False(ExportRules.TryParseArgs(args, out location));
            Assert.Null(location);
        }
    }
}
