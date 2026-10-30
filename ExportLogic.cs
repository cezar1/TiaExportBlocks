using System;
using System.Collections.Generic;
using System.IO;

// Logic that does not depend on Siemens.Engineering, kept separate so it can be unit tested
// without TIA Portal. The test project compiles this file directly (see tests/).
namespace TiaExportBlocks
{
    /// <summary>Decisions about what gets exported and in which format.</summary>
    public static class ExportRules
    {
        public const string S71500DeviceType = "System:Device.S71500";

        /// <summary>Block programming language (ProgrammingLanguage.ToString()) to file extension.</summary>
        public static readonly IReadOnlyDictionary<string, string> LanguageToExtension =
            new Dictionary<string, string> { { "SCL", ".scl" }, { "DB", ".db" }, { "STL", ".awl" } };

        /// <summary>Block programming language to the sub folder it is exported into.</summary>
        public static readonly IReadOnlyDictionary<string, string> LanguageToFolderPrefix =
            new Dictionary<string, string> { { "SCL", @"\scl\" }, { "DB", @"\db\" }, { "STL", @"\stl\" } };

        /// <summary>Returns false for languages that cannot be exported as source (LAD, FBD, GRAPH, ...).</summary>
        public static bool TryGetBlockFormat(string programmingLanguage, out string extension, out string folderPrefix)
        {
            extension = null;
            folderPrefix = null;
            if (programmingLanguage == null) return false;
            return LanguageToExtension.TryGetValue(programmingLanguage, out extension)
                && LanguageToFolderPrefix.TryGetValue(programmingLanguage, out folderPrefix);
        }

        /// <summary>Devices of this type are treated as PLCs, everything else as HMI.</summary>
        public static bool IsS71500(string deviceTypeIdentifier)
        {
            return deviceTypeIdentifier == S71500DeviceType;
        }

        /// <summary>Inside an S7-1500 device, only device items whose name contains "PLC" are exported.</summary>
        public static bool IsPlcDeviceItem(string deviceItemName)
        {
            return deviceItemName != null && deviceItemName.Contains("PLC");
        }

        /// <summary>The program takes exactly one argument: the export location.</summary>
        public static bool TryParseArgs(string[] args, out string exportLocation)
        {
            exportLocation = null;
            if (args == null || args.Length != 1) return false;
            exportLocation = args[0];
            return true;
        }
    }

    /// <summary>Where each exported element is written.</summary>
    public static class ExportPaths
    {
        /// <summary>Replaces every character that is invalid in a Windows file name with '_'.</summary>
        public static string MakeValidFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }

        /// <summary>Sanitizes each '\'-separated segment of a block group path.</summary>
        public static string SanitizeRelativePath(string currentPath)
        {
            string[] folderStructure = currentPath.Split('\\');
            for (int i = 0; i < folderStructure.Length; i++)
            {
                folderStructure[i] = MakeValidFileName(folderStructure[i]);
            }
            return Path.Combine(folderStructure);
        }

        /// <summary>Path of a nested block group, relative to the PLC's block root.</summary>
        public static string GroupPath(string currentPath, string groupName)
        {
            return Path.Combine(currentPath, MakeValidFileName(groupName));
        }

        /// <summary>Final location of an exported block, e.g. &lt;export&gt;\PLC_1\scl\Program blocks\Main.scl</summary>
        public static string BlockFile(string exportLocation, string rootPrefix, string folderPrefix,
            string currentPath, string blockName, string extension)
        {
            return exportLocation + @"\" + rootPrefix + @"\"
                + Path.Combine(folderPrefix, SanitizeRelativePath(currentPath), MakeValidFileName(blockName) + extension);
        }

        /// <summary>
        /// File TIA generates the block source into before it is moved to <see cref="BlockFile"/>.
        /// Note: no separator is inserted, so without a trailing '\' on exportLocation the file lands
        /// next to the export folder (C:\Test + Main.scl = C:\TestMain.scl).
        /// </summary>
        public static string BlockTempFile(string exportLocation, string blockName, string extension)
        {
            return exportLocation + MakeValidFileName(blockName) + extension;
        }

        /// <summary>Location of an exported user data type, e.g. &lt;export&gt;\PLC_1\udt\Motor.udt</summary>
        public static string TypeFile(string exportLocation, string rootPrefix, string typeName)
        {
            return exportLocation + @"\" + rootPrefix + @"\udt\" + typeName + ".udt";
        }

        /// <summary>File a PLC tag table is exported into before it is moved to <see cref="PlcTagTableFile"/>.</summary>
        public static string PlcTagTableTempFile(string exportLocation, string tableName)
        {
            return exportLocation + @"\" + MakeValidFileName(tableName) + ".xml";
        }

        /// <summary>Location of an exported PLC tag table, e.g. &lt;export&gt;\PLC_1\tag_tables\xml\Default tag table.xml</summary>
        public static string PlcTagTableFile(string exportLocation, string rootPrefix, string tableName)
        {
            return exportLocation + @"\" + rootPrefix + @"\tag_tables\xml\" + @"\" + tableName + ".xml";
        }

        /// <summary>Location of an exported HMI tag table, e.g. &lt;export&gt;\hmi_tag_tables\xml\Tags.xml</summary>
        public static string HmiTagTableFile(string exportLocation, string tableName)
        {
            return exportLocation + @"\hmi_tag_tables\xml\" + tableName + ".xml";
        }

        /// <summary>Location of an exported HMI text list, e.g. &lt;export&gt;\hmi_text_lists\xml\Alarms.xml</summary>
        public static string HmiTextListFile(string exportLocation, string textListName)
        {
            return exportLocation + @"\hmi_text_lists\xml\" + textListName + ".xml";
        }
    }

    /// <summary>File system steps used around the TIA export calls.</summary>
    public static class FileOps
    {
        /// <summary>Creates the directory part of a file path if it is missing.</summary>
        public static void EnsureDirectoryExists(string filePath)
        {
            string directoryPath = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
        }

        /// <summary>Creates the directory if missing. Returns true if it had to be created.</summary>
        public static bool EnsureDirectory(string path)
        {
            if (Directory.Exists(path)) return false;
            Directory.CreateDirectory(path);
            return true;
        }

        /// <summary>Deletes the directory and everything in it. Returns true if it existed.</summary>
        public static bool RemoveDirectory(string path)
        {
            if (!Directory.Exists(path)) return false;
            Directory.Delete(path, true);
            return true;
        }

        /// <summary>Deletes the file if present. Returns true if it existed.</summary>
        public static bool DeleteIfExists(string path)
        {
            if (!File.Exists(path)) return false;
            File.Delete(path);
            return true;
        }

        /// <summary>Moves source to destination, replacing any existing destination file.</summary>
        public static void MoveReplacing(string sourcePath, string destinationPath)
        {
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException("Exported file not found", sourcePath);
            EnsureDirectoryExists(destinationPath);
            DeleteIfExists(destinationPath);
            File.Move(sourcePath, destinationPath);
        }
    }
}
