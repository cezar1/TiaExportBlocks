using System.IO;
using System.Linq;
using Xunit;

namespace TiaExportBlocks.Tests
{
    public class ExportPathsTests
    {
        private const string Export = @"C:\Export";

        // Program wraps every path in a FileInfo, which normalizes separators; compare the same way.
        private static string Full(string path)
        {
            return new FileInfo(path).FullName;
        }

        // ---- MakeValidFileName ----

        [Theory]
        [InlineData("Main")]
        [InlineData("FB_Conveyor_Control")]
        [InlineData("Program blocks")]
        [InlineData("Motor.v2")]
        [InlineData("Ünïcödé_Ș_ț")]
        [InlineData("")]
        public void MakeValidFileName_ValidName_IsUnchanged(string name)
        {
            Assert.Equal(name, ExportPaths.MakeValidFileName(name));
        }

        [Theory]
        [InlineData('<')]
        [InlineData('>')]
        [InlineData(':')]
        [InlineData('"')]
        [InlineData('/')]
        [InlineData('\\')]
        [InlineData('|')]
        [InlineData('?')]
        [InlineData('*')]
        [InlineData('\0')]
        [InlineData('\t')]
        [InlineData('\n')]
        public void MakeValidFileName_InvalidCharacter_IsReplacedWithUnderscore(char invalid)
        {
            Assert.Equal("a_b", ExportPaths.MakeValidFileName("a" + invalid + "b"));
        }

        [Fact]
        public void MakeValidFileName_EveryInvalidCharacter_IsReplaced()
        {
            string allInvalid = new string(Path.GetInvalidFileNameChars());

            string result = ExportPaths.MakeValidFileName(allInvalid);

            Assert.Equal(new string('_', allInvalid.Length), result);
        }

        [Fact]
        public void MakeValidFileName_KeepsLength()
        {
            const string name = @"Line 3: Motor/Pump ""A"" <old>|?*";

            Assert.Equal(name.Length, ExportPaths.MakeValidFileName(name).Length);
        }

        [Fact]
        public void MakeValidFileName_ResultHasNoInvalidCharacters()
        {
            string result = ExportPaths.MakeValidFileName(@"x:\y/z<1>""2""|3?4*");

            Assert.Equal(-1, result.IndexOfAny(Path.GetInvalidFileNameChars()));
            Assert.Equal("x__y_z_1__2__3_4_", result);
        }

        // ---- SanitizeRelativePath / GroupPath ----

        [Fact]
        public void SanitizeRelativePath_Empty_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, ExportPaths.SanitizeRelativePath(string.Empty));
        }

        [Fact]
        public void SanitizeRelativePath_KeepsFolderStructure()
        {
            Assert.Equal(Path.Combine("Program blocks", "Conveyors", "Line 3"),
                ExportPaths.SanitizeRelativePath(@"Program blocks\Conveyors\Line 3"));
        }

        [Fact]
        public void SanitizeRelativePath_SanitizesEachSegment()
        {
            Assert.Equal(Path.Combine("Program blocks", "A_B", "C_D"),
                ExportPaths.SanitizeRelativePath(@"Program blocks\A/B\C:D"));
        }

        [Fact]
        public void GroupPath_FromRoot_IsJustTheGroupName()
        {
            Assert.Equal("Program blocks", ExportPaths.GroupPath(string.Empty, "Program blocks"));
        }

        [Fact]
        public void GroupPath_Nested_AppendsGroupName()
        {
            Assert.Equal(Path.Combine("Program blocks", "Conveyors"),
                ExportPaths.GroupPath("Program blocks", "Conveyors"));
        }

        [Fact]
        public void GroupPath_SanitizesGroupName()
        {
            Assert.Equal(Path.Combine("Program blocks", "In_Out"),
                ExportPaths.GroupPath("Program blocks", "In/Out"));
        }

        [Fact]
        public void GroupPath_ThreeLevelsDeep_BuildsNestedPath()
        {
            string level1 = ExportPaths.GroupPath(string.Empty, "Program blocks");
            string level2 = ExportPaths.GroupPath(level1, "Area 1");
            string level3 = ExportPaths.GroupPath(level2, "Station:5");

            Assert.Equal(Path.Combine("Program blocks", "Area 1", "Station_5"), level3);
        }

        // ---- BlockFile ----

        [Fact]
        public void BlockFile_SclBlockInRootGroup()
        {
            string path = ExportPaths.BlockFile(Export, "PLC_1", @"\scl\", "Program blocks", "Main", ".scl");

            Assert.Equal(@"C:\Export\PLC_1\scl\Program blocks\Main.scl", Full(path));
        }

        [Fact]
        public void BlockFile_DataBlockInNestedGroup()
        {
            string path = ExportPaths.BlockFile(Export, "PLC_1", @"\db\", @"Program blocks\Data\Recipes", "DB_Recipe", ".db");

            Assert.Equal(@"C:\Export\PLC_1\db\Program blocks\Data\Recipes\DB_Recipe.db", Full(path));
        }

        [Fact]
        public void BlockFile_StlBlockUsesAwlExtension()
        {
            string extension, folder;
            ExportRules.TryGetBlockFormat("STL", out extension, out folder);

            string path = ExportPaths.BlockFile(Export, "PLC_1", folder, "Program blocks", "FC_Legacy", extension);

            Assert.Equal(@"C:\Export\PLC_1\stl\Program blocks\FC_Legacy.awl", Full(path));
        }

        [Fact]
        public void BlockFile_EmptyGroupPath_PutsBlockDirectlyInLanguageFolder()
        {
            string path = ExportPaths.BlockFile(Export, "PLC_1", @"\scl\", string.Empty, "Main", ".scl");

            Assert.Equal(@"C:\Export\PLC_1\scl\Main.scl", Full(path));
        }

        [Fact]
        public void BlockFile_SanitizesBlockName()
        {
            string path = ExportPaths.BlockFile(Export, "PLC_1", @"\scl\", "Program blocks", "Motor/Pump?", ".scl");

            Assert.Equal(@"C:\Export\PLC_1\scl\Program blocks\Motor_Pump_.scl", Full(path));
        }

        [Fact]
        public void BlockFile_SanitizesGroupSegments()
        {
            string path = ExportPaths.BlockFile(Export, "PLC_1", @"\scl\", @"Program blocks\In|Out", "Main", ".scl");

            Assert.Equal(@"C:\Export\PLC_1\scl\Program blocks\In_Out\Main.scl", Full(path));
        }

        [Fact]
        public void BlockFile_TrailingSeparatorOnExportLocation_GivesSamePath()
        {
            string without = ExportPaths.BlockFile(@"C:\Export", "PLC_1", @"\scl\", "Program blocks", "Main", ".scl");
            string with = ExportPaths.BlockFile(@"C:\Export\", "PLC_1", @"\scl\", "Program blocks", "Main", ".scl");

            Assert.Equal(Full(without), Full(with));
        }

        [Fact]
        public void BlockFile_DifferentPlcs_DoNotCollide()
        {
            string plc1 = ExportPaths.BlockFile(Export, "PLC_1", @"\scl\", "Program blocks", "Main", ".scl");
            string plc2 = ExportPaths.BlockFile(Export, "PLC_2", @"\scl\", "Program blocks", "Main", ".scl");

            Assert.NotEqual(Full(plc1), Full(plc2));
        }

        [Fact]
        public void BlockFile_SameNameDifferentLanguages_DoNotCollide()
        {
            var paths = ExportRules.LanguageToExtension.Keys
                .Select(lang => ExportPaths.BlockFile(Export, "PLC_1", ExportRules.LanguageToFolderPrefix[lang],
                    "Program blocks", "Main", ExportRules.LanguageToExtension[lang]))
                .Select(Full)
                .ToList();

            Assert.Equal(paths.Count, paths.Distinct().Count());
        }

        [Fact]
        public void BlockFile_IsAlwaysInsideExportLocationAndPlcFolder()
        {
            string path = Full(ExportPaths.BlockFile(Export, "PLC_1", @"\scl\", @"a\b\c", @"..\..\evil", ".scl"));

            Assert.StartsWith(@"C:\Export\PLC_1\", path);
        }

        // ---- BlockTempFile ----

        [Fact]
        public void BlockTempFile_WithTrailingSeparator_IsInsideExportLocation()
        {
            Assert.Equal(@"C:\Export\Main.scl", Full(ExportPaths.BlockTempFile(@"C:\Export\", "Main", ".scl")));
        }

        [Fact]
        public void BlockTempFile_WithoutTrailingSeparator_IsGluedToFolderName()
        {
            // Current behaviour, kept as-is in this refactor: no separator is inserted.
            Assert.Equal(@"C:\ExportMain.scl", Full(ExportPaths.BlockTempFile(@"C:\Export", "Main", ".scl")));
        }

        [Fact]
        public void BlockTempFile_SanitizesBlockName()
        {
            Assert.Equal(@"C:\Export\A_B.db", Full(ExportPaths.BlockTempFile(@"C:\Export\", "A:B", ".db")));
        }

        // ---- UDTs ----

        [Fact]
        public void TypeFile_IsInUdtFolderOfPlc()
        {
            Assert.Equal(@"C:\Export\PLC_1\udt\UDT_Motor.udt", Full(ExportPaths.TypeFile(Export, "PLC_1", "UDT_Motor")));
        }

        [Fact]
        public void TypeFile_DifferentPlcs_DoNotCollide()
        {
            Assert.NotEqual(Full(ExportPaths.TypeFile(Export, "PLC_1", "UDT_Motor")),
                Full(ExportPaths.TypeFile(Export, "PLC_2", "UDT_Motor")));
        }

        // ---- PLC tag tables ----

        [Fact]
        public void PlcTagTableTempFile_IsDirectlyInExportLocation()
        {
            Assert.Equal(@"C:\Export\Default tag table.xml", Full(ExportPaths.PlcTagTableTempFile(Export, "Default tag table")));
        }

        [Fact]
        public void PlcTagTableTempFile_SanitizesName()
        {
            Assert.Equal(@"C:\Export\IO_Inputs.xml", Full(ExportPaths.PlcTagTableTempFile(Export, "IO/Inputs")));
        }

        [Fact]
        public void PlcTagTableFile_IsInTagTablesXmlFolderOfPlc()
        {
            Assert.Equal(@"C:\Export\PLC_1\tag_tables\xml\Default tag table.xml",
                Full(ExportPaths.PlcTagTableFile(Export, "PLC_1", "Default tag table")));
        }

        [Fact]
        public void PlcTagTableFile_DiffersFromTempFile()
        {
            Assert.NotEqual(Full(ExportPaths.PlcTagTableTempFile(Export, "Tags")),
                Full(ExportPaths.PlcTagTableFile(Export, "PLC_1", "Tags")));
        }

        // ---- HMI ----

        [Fact]
        public void HmiTagTableFile_IsInHmiTagTablesFolder()
        {
            Assert.Equal(@"C:\Export\hmi_tag_tables\xml\Default tag table.xml",
                Full(ExportPaths.HmiTagTableFile(Export, "Default tag table")));
        }

        [Fact]
        public void HmiTextListFile_IsInHmiTextListsFolder()
        {
            Assert.Equal(@"C:\Export\hmi_text_lists\xml\Alarm_Texts.xml",
                Full(ExportPaths.HmiTextListFile(Export, "Alarm_Texts")));
        }

        [Fact]
        public void HmiTagTableAndTextList_WithSameName_DoNotCollide()
        {
            Assert.NotEqual(Full(ExportPaths.HmiTagTableFile(Export, "X")), Full(ExportPaths.HmiTextListFile(Export, "X")));
        }

        // ---- Layout as a whole ----

        [Fact]
        public void AllExportPaths_AreUnderExportLocation()
        {
            var paths = new[]
            {
                ExportPaths.BlockFile(Export, "PLC_1", @"\scl\", "Program blocks", "Main", ".scl"),
                ExportPaths.TypeFile(Export, "PLC_1", "UDT_Motor"),
                ExportPaths.PlcTagTableTempFile(Export, "Tags"),
                ExportPaths.PlcTagTableFile(Export, "PLC_1", "Tags"),
                ExportPaths.HmiTagTableFile(Export, "Tags"),
                ExportPaths.HmiTextListFile(Export, "Texts"),
            };

            Assert.All(paths.Select(Full), p => Assert.StartsWith(@"C:\Export\", p));
        }

        [Fact]
        public void AllExportPaths_WorkWithRelativeExportLocation()
        {
            string cwd = Directory.GetCurrentDirectory();

            string path = Full(ExportPaths.TypeFile("out", "PLC_1", "UDT_Motor"));

            Assert.Equal(Path.Combine(cwd, "out", "PLC_1", "udt", "UDT_Motor.udt"), path);
        }
    }
}
