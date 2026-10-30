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

        /// <summary>
        /// Absolute export location without trailing separator. Refuses empty/invalid paths and drive or share
        /// roots: everything inside the export location is deleted on every full export.
        /// </summary>
        public static bool TryNormalizeExportLocation(string path, out string normalized, out string error)
        {
            normalized = null;
            error = null;
            if (string.IsNullOrWhiteSpace(path))
            {
                error = "Export location is empty";
                return false;
            }
            string full;
            try
            {
                full = Path.GetFullPath(path);
            }
            catch (Exception exc)
            {
                error = "Export location '" + path + "' is not a valid path: " + exc.Message;
                return false;
            }
            string root = Path.GetPathRoot(full) ?? string.Empty;
            string trimmed = full.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            if (root.Length == 0 || trimmed.Length <= root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).Length)
            {
                error = "Export location '" + path + "' is a drive or share root. Use a dedicated folder: its whole content is deleted on every export.";
                return false;
            }
            normalized = trimmed;
            return true;
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
            return exportLocation + @"\" + MakeValidFileName(rootPrefix) + @"\"
                + Path.Combine(folderPrefix, SanitizeRelativePath(currentPath), MakeValidFileName(blockName) + extension);
        }

        /// <summary>File TIA generates the block source into (inside the export folder) before it is moved to <see cref="BlockFile"/>.</summary>
        public static string BlockTempFile(string exportLocation, string blockName, string extension)
        {
            return Path.Combine(exportLocation, MakeValidFileName(blockName) + extension);
        }

        /// <summary>Location of an exported user data type, e.g. &lt;export&gt;\PLC_1\udt\Motor.udt</summary>
        public static string TypeFile(string exportLocation, string rootPrefix, string typeName)
        {
            return exportLocation + @"\" + MakeValidFileName(rootPrefix) + @"\udt\" + MakeValidFileName(typeName) + ".udt";
        }

        /// <summary>File a PLC tag table is exported into before it is moved to <see cref="PlcTagTableFile"/>.</summary>
        public static string PlcTagTableTempFile(string exportLocation, string tableName)
        {
            return exportLocation + @"\" + MakeValidFileName(tableName) + ".xml";
        }

        /// <summary>Location of an exported PLC tag table, e.g. &lt;export&gt;\PLC_1\tag_tables\xml\Default tag table.xml</summary>
        public static string PlcTagTableFile(string exportLocation, string rootPrefix, string tableName)
        {
            return exportLocation + @"\" + MakeValidFileName(rootPrefix) + @"\tag_tables\xml\" + MakeValidFileName(tableName) + ".xml";
        }

        /// <summary>Location of an exported HMI tag table, e.g. &lt;export&gt;\HMI_1\hmi_tag_tables\xml\Tags.xml</summary>
        public static string HmiTagTableFile(string exportLocation, string hmiPrefix, string tableName)
        {
            return exportLocation + @"\" + MakeValidFileName(hmiPrefix) + @"\hmi_tag_tables\xml\" + MakeValidFileName(tableName) + ".xml";
        }

        /// <summary>Location of an exported HMI text list, e.g. &lt;export&gt;\HMI_1\hmi_text_lists\xml\Alarms.xml</summary>
        public static string HmiTextListFile(string exportLocation, string hmiPrefix, string textListName)
        {
            return exportLocation + @"\" + MakeValidFileName(hmiPrefix) + @"\hmi_text_lists\xml\" + MakeValidFileName(textListName) + ".xml";
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

        /// <summary>
        /// Deletes everything inside <paramref name="path"/> but keeps the folder itself. Read-only, hidden and
        /// system attributes are cleared first. Links (junctions, symlinks) are removed without touching what they
        /// point to. Locked entries are retried. Returns what could not be deleted (empty when the folder is clean).
        /// </summary>
        public static List<string> WipeDirectoryContents(string path, int attempts = 5, int retryDelayMs = 250)
        {
            var remaining = new List<string>();
            if (!Directory.Exists(path)) return remaining;
            for (int attempt = 1; attempt <= attempts; attempt++)
            {
                foreach (string entry in Directory.GetFileSystemEntries(path))
                {
                    try { DeleteEntry(entry); }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
                remaining = ListEntries(path);
                if (remaining.Count == 0 || attempt == attempts) break;
                System.Threading.Thread.Sleep(retryDelayMs);
            }
            return remaining;
        }

        /// <summary>All files and folders under <paramref name="path"/>, without following links.</summary>
        public static List<string> ListEntries(string path)
        {
            var result = new List<string>();
            if (!Directory.Exists(path)) return result;
            foreach (string entry in Directory.GetFileSystemEntries(path))
            {
                result.Add(entry);
                FileAttributes attributes = File.GetAttributes(entry);
                if ((attributes & FileAttributes.Directory) != 0 && (attributes & FileAttributes.ReparsePoint) == 0)
                {
                    result.AddRange(ListEntries(entry));
                }
            }
            return result;
        }

        private static void DeleteEntry(string entry)
        {
            FileAttributes attributes = File.GetAttributes(entry);
            bool isDirectory = (attributes & FileAttributes.Directory) != 0;
            bool isLink = (attributes & FileAttributes.ReparsePoint) != 0;
            ClearProtection(entry, attributes);
            if (!isDirectory)
            {
                File.Delete(entry);
                return;
            }
            if (!isLink)
            {
                foreach (string child in Directory.GetFileSystemEntries(entry))
                {
                    DeleteEntry(child);
                }
            }
            // Non-recursive: for a junction/symlink this removes only the link, never the target.
            Directory.Delete(entry, false);
        }

        private static void ClearProtection(string entry, FileAttributes attributes)
        {
            const FileAttributes protection = FileAttributes.ReadOnly | FileAttributes.Hidden | FileAttributes.System;
            if ((attributes & protection) != 0)
            {
                File.SetAttributes(entry, attributes & ~protection);
            }
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
