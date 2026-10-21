using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Siemens.Engineering;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.Hmi.Tag;
using Siemens.Engineering.Hmi.TextGraphicList;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.ExternalSources;
using Siemens.Engineering.SW.Tags;
using Siemens.Engineering.SW.Types;

namespace TiaExportBlocks
{
    class Program
    {
        static Dictionary<string, string> programLanguageToExtension = new Dictionary<string, string> { { "SCL", ".scl" }, { "DB", ".db" }, { "STL", ".awl" }};
        static Dictionary<string, string> programLanguageToFolderPrefixExtension = new Dictionary<string, string> { { "SCL", @"\scl\" }, { "DB", @"\db\" }, { "STL", @"\stl\" } };
        static string exportLocation;
        // Incremental export: a block, type or tag table is exported again only
        // when its TIA modification stamp differs from the last export or its
        // file is missing. --full restores the old behaviour (wipe and export all).
        static bool fullExport;
        static Dictionary<string, long> previousStamps = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        // File write times at the last export: a file changed on disk since (e.g. a
        // git branch switch) is exported again even if TIA reports no change.
        static Dictionary<string, long> previousWrites = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        static Dictionary<string, long> currentStamps = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        static HashSet<string> expectedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        static int exportedCount;
        static int skippedCount;

        // The TIA modification stamp of a block, type or tag table, or -1 when
        // it cannot be read (then the element is always exported).
        static long ModifiedStamp(IEngineeringObject obj)
        {
            try
            {
                object value = obj.GetAttribute("ModifiedDate");
                if (value is DateTime) return ((DateTime)value).Ticks;
            }
            catch (Exception)
            {
            }
            return -1;
        }

        // True when the element at targetPath must be exported again.
        static bool NeedsExport(string targetPath, long stamp)
        {
            expectedFiles.Add(targetPath);
            currentStamps[targetPath] = stamp;
            if (fullExport || stamp < 0 || !File.Exists(targetPath)) return true;
            long previous;
            long previousWrite;
            if (previousStamps.TryGetValue(targetPath, out previous) && previous == stamp
                && previousWrites.TryGetValue(targetPath, out previousWrite)
                && previousWrite == File.GetLastWriteTimeUtc(targetPath).Ticks)
            {
                skippedCount++;
                return false;
            }
            return true;
        }

        // Openness writes into the local temp folder; the result is then copied
        // to the (possibly shared) export location in one step.
        static string LocalTempFile(string name)
        {
            return Path.Combine(Path.GetTempPath(), "TiaExportBlocks_" + Guid.NewGuid().ToString("N") + "_" + name);
        }

        static void PlaceFile(string tempPath, string targetPath)
        {
            EnsureDirectoryExists(targetPath);
            File.Copy(tempPath, targetPath, true);
            File.Delete(tempPath);
            exportedCount++;
        }

        // Batch export: each Openness GenerateSource call costs about the same
        // whether it carries one block or many, so blocks are generated in
        // batches (--batch N, default 50) into one file and split back into the
        // usual one-file-per-block layout. A batch that fails or does not split
        // into exactly its blocks is exported block by block instead.
        static int batchSize = 50;
        static Dictionary<string, List<KeyValuePair<PlcBlock, string>>> pendingBlocks = new Dictionary<string, List<KeyValuePair<PlcBlock, string>>>();
        static List<KeyValuePair<PlcType, string>> pendingTypes = new List<KeyValuePair<PlcType, string>>();
        static Regex blockStart = new Regex(@"^(FUNCTION_BLOCK|FUNCTION|ORGANIZATION_BLOCK|DATA_BLOCK|TYPE)\s+""([^""]+)""");
        static UTF8Encoding utf8WithBom = new UTF8Encoding(true);

        // Splits a generated source into its top-level blocks, each as written by a
        // single-block export: the block's lines, its END_ line and one empty line.
        static Dictionary<string, string> SplitSource(string text)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            string[] lines = text.Replace("\r\n", "\n").Split('\n');
            List<string> current = null;
            string name = null;
            string endKeyword = null;
            foreach (string line in lines)
            {
                if (current == null)
                {
                    Match m = blockStart.Match(line);
                    if (!m.Success) continue;
                    current = new List<string>();
                    name = m.Groups[2].Value;
                    endKeyword = "END_" + m.Groups[1].Value;
                }
                current.Add(line);
                if (line.TrimEnd() == endKeyword)
                {
                    result[name] = string.Join("\r\n", current) + "\r\n\r\n";
                    current = null;
                }
            }
            return result;
        }

        static void ExportBatched<T>(List<KeyValuePair<T, string>> items, string extension, Func<T, string> nameOf, Action<List<T>, FileInfo> generate)
        {
            for (int start = 0; start < items.Count; start += batchSize)
            {
                var batch = items.Skip(start).Take(batchSize).ToList();
                bool done = false;
                if (batch.Count > 1)
                {
                    var temp = new FileInfo(LocalTempFile("batch" + extension));
                    try
                    {
                        Console.WriteLine("Generating " + batch.Count + " " + extension + " blocks in one batch");
                        generate(batch.Select(b => b.Key).ToList(), temp);
                        var parts = SplitSource(File.ReadAllText(temp.FullName, Encoding.UTF8));
                        if (parts.Count == batch.Count && batch.All(b => parts.ContainsKey(nameOf(b.Key))))
                        {
                            foreach (var b in batch)
                            {
                                EnsureDirectoryExists(b.Value);
                                File.WriteAllText(b.Value, parts[nameOf(b.Key)], utf8WithBom);
                                exportedCount++;
                            }
                            done = true;
                        }
                        else
                        {
                            Console.WriteLine("Batch did not split into its " + batch.Count + " blocks (" + parts.Count + " found): exporting them one by one");
                        }
                    }
                    catch (Exception exc)
                    {
                        Console.WriteLine("Batch export failed (" + exc.Message + "): exporting its blocks one by one");
                    }
                    finally
                    {
                        if (File.Exists(temp.FullName)) File.Delete(temp.FullName);
                    }
                }
                if (done) continue;
                foreach (var b in batch)
                {
                    var temp = new FileInfo(LocalTempFile(MakeValidFileName(nameOf(b.Key)) + extension));
                    try
                    {
                        Console.WriteLine("Block [" + nameOf(b.Key) + "] to " + b.Value);
                        generate(new List<T>() { b.Key }, temp);
                        PlaceFile(temp.FullName, b.Value);
                    }
                    catch (Exception exc)
                    {
                        Console.WriteLine(exc.ToString());
                    }
                }
            }
        }

        static void FlushBlocks(PlcSoftware software)
        {
            PlcExternalSourceSystemGroup group = software.ExternalSourceGroup;
            foreach (var entry in pendingBlocks)
            {
                ExportBatched(entry.Value, entry.Key, b => b.Name, (list, file) => group.GenerateSource(list, file, GenerateOptions.None));
            }
            pendingBlocks.Clear();
        }

        static void FlushTypes(PlcSoftware software)
        {
            PlcExternalSourceSystemGroup group = software.ExternalSourceGroup;
            ExportBatched(pendingTypes, ".udt", t => t.Name, (list, file) => group.GenerateSource(list, file, GenerateOptions.None));
            pendingTypes.Clear();
        }
        static void HandleBlock(PlcBlock block,PlcSoftware software, string currentPath, string rootPrefix)
        {
            //Console.WriteLine(block.Name + " " + block.GetType()+ " " +block.ProgrammingLanguage);
            string block_programming_language=block.ProgrammingLanguage.ToString();
            if (programLanguageToExtension.ContainsKey(block_programming_language))
            {
                string extension;
                programLanguageToExtension.TryGetValue(block.ProgrammingLanguage.ToString(),out extension);
                string folder_prefix;
                programLanguageToFolderPrefixExtension.TryGetValue(block.ProgrammingLanguage.ToString(), out folder_prefix);
                // Convert the folder structure to a valid path
                string[] folderStructure = currentPath.Split('\\');
                for (int i = 0; i < folderStructure.Length; i++)
                {
                    folderStructure[i] = MakeValidFileName(folderStructure[i]);
                }
                string sanitizedCurrentPath= Path.Combine(folderStructure);
                var fileInfo = new FileInfo(exportLocation+@"\"+ rootPrefix+@"\"+Path.Combine(folder_prefix, sanitizedCurrentPath, MakeValidFileName(block.Name)+ extension));
                if (!NeedsExport(fileInfo.FullName, ModifiedStamp(block))) return;
                List<KeyValuePair<PlcBlock, string>> queue;
                if (!pendingBlocks.TryGetValue(extension, out queue))
                {
                    queue = new List<KeyValuePair<PlcBlock, string>>();
                    pendingBlocks[extension] = queue;
                }
                queue.Add(new KeyValuePair<PlcBlock, string>(block, fileInfo.FullName));
            }
            else
            {
                Console.WriteLine(block.Name + " programming language " + block_programming_language + " not supported");
            }

        }
        static void HandleBlockGroup(PlcBlockGroup blockGroup, PlcSoftware software, string currentPath, string rootPrefix)
        {
            string newPath = Path.Combine(currentPath, MakeValidFileName(blockGroup.Name));
            try
            {
                Console.WriteLine("Handling block group: " + blockGroup.Name);

                // Handle blocks in the current block group
                foreach (PlcBlock block in blockGroup.Blocks)
                {
                    HandleBlock(block, software, newPath, rootPrefix);
                }

                // Recursively handle nested block groups
                foreach (PlcBlockGroup nestedGroup in blockGroup.Groups)
                {
                    HandleBlockGroup(nestedGroup, software, newPath, rootPrefix);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error handling block group: " + blockGroup.Name + " - " + ex.Message);
            }
        }
        static string MakeValidFileName(string name)
        {
            // Replace invalid characters with an underscore
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
        static void HandleType(PlcType plcType, PlcSoftware software,string rootPrefix)
        {
            string extension=".udt";
            var fileInfo = new FileInfo(exportLocation + @"\" + rootPrefix + @"\udt\" + plcType.Name + extension);
            if (!NeedsExport(fileInfo.FullName, ModifiedStamp(plcType))) return;
            pendingTypes.Add(new KeyValuePair<PlcType, string>(plcType, fileInfo.FullName));

        }
        private static void ExportAllTagTables(PlcSoftware plcSoftware, string rootPrefix)
        {
            PlcTagTableSystemGroup plcTagTableSystemGroup = plcSoftware.TagTableGroup;
            // Export all tables in the system group
            ExportTagTables(plcTagTableSystemGroup.TagTables, rootPrefix);
            // Export the tables in underlying user groups
            foreach (PlcTagTableUserGroup userGroup in plcTagTableSystemGroup.Groups)
            {
                ExportUserGroupDeep(userGroup, rootPrefix);
            }
        }
        private static void ExportBlocks(PlcSoftware software, string rootPrefix)
        {
            string name = software.Name;
            Console.WriteLine(name);
            // Recursively handle block groups and blocks
            HandleBlockGroup(software.BlockGroup, software, string.Empty, rootPrefix);
            FlushBlocks(software);

        }
        private static void ExportTypes(PlcSoftware software, string rootPrefix)
        {
            foreach (PlcType plcType in software.TypeGroup.Types)
            {
                Console.WriteLine("Handling type at prefix " + rootPrefix + " named "+ plcType.Name);
                HandleType(plcType, software, rootPrefix);
            }
            FlushTypes(software);
        }
        public static void EnsureDirectoryExists(string filePath)
        {
            // Extract the directory part from the file path
            string directoryPath = Path.GetDirectoryName(filePath);

            if (!string.IsNullOrEmpty(directoryPath))
            {
                // Ensure the directory exists
                Directory.CreateDirectory(directoryPath);
            }
        }
        private static void ExportTagTables(PlcTagTableComposition tagTables, string rootPrefix)
        {
            foreach (PlcTagTable table in tagTables)
            {
                Console.WriteLine(table.Name);
                string filePath = exportLocation +@"\"+ rootPrefix+@"\tag_tables\xml\"  + @"\" + table.Name + ".xml";
                var fileInfo = new FileInfo(filePath);
                // Tag tables are few and also hold the user constants: always exported.
                expectedFiles.Add(fileInfo.FullName);
                var fileInfoTemp = new FileInfo(LocalTempFile(MakeValidFileName(table.Name) + ".xml"));
                table.Export(fileInfoTemp, ExportOptions.WithDefaults);
                Console.WriteLine(table.Name + " to " + fileInfo.FullName);
                PlaceFile(fileInfoTemp.FullName, fileInfo.FullName);
            }
        }
        private static void ExportUserGroupDeep(PlcTagTableUserGroup group, string rootPrefix)
        {
            ExportTagTables(group.TagTables, rootPrefix);
            foreach (PlcTagTableUserGroup userGroup in group.Groups)
            {
                ExportUserGroupDeep(userGroup, rootPrefix);
            }
        }
        //Exports all tag tables from an HMI device 
        private static void ExportAllTagTablesFromHMITarget(HmiTarget hmitarget) 
        {    
            TagSystemFolder sysFolder = hmitarget.TagFolder;        
            //First export the tables in underlying user folder    
            foreach (TagUserFolder userFolder in sysFolder.Folders)    
            {            
                ExportUserFolderDeep(userFolder);    
            }        
            //then, export all tables in the system folder    
            ExportTablesInSystemFolder(sysFolder); 
        }    
        private static void ExportUserFolderDeep(TagUserFolder rootUserFolder) 
        {            
            foreach (TagUserFolder userFolder in rootUserFolder.Folders)            
            {                    
                ExportUserFolderDeep(userFolder);            
            }            
            ExportTablesInUserFolder(rootUserFolder); 
        }
        private static void ExportTablesInUserFolder(TagUserFolder folderToExport) 
        { 
            TagTableComposition tables = folderToExport.TagTables; 
            foreach (TagTable table in tables) 
            {
                string extension = ".xml";
                var fileInfo = new FileInfo(exportLocation + @"\hmi_tag_tables\xml\" + table.Name + extension);
                try
                {
                    if (File.Exists(fileInfo.FullName)) File.Delete(fileInfo.FullName);
                    Console.WriteLine(table.Name + " to " + fileInfo.FullName);
                    table.Export(fileInfo, ExportOptions.WithDefaults);
                }
                catch (Exception exc)
                {
                    Console.WriteLine(exc.ToString());
                }
            } 
        }
        private static void ExportTablesInSystemFolder(TagSystemFolder folderToExport) 
        { 
            TagTableComposition tables = folderToExport.TagTables; 
            foreach (TagTable table in tables) 
            {
                string extension = ".xml";
                var fileInfo = new FileInfo(exportLocation + @"\hmi_tag_tables\xml\" + table.Name + extension);
                try
                {
                    if (File.Exists(fileInfo.FullName)) File.Delete(fileInfo.FullName);
                    Console.WriteLine(table.Name + " to " + fileInfo.FullName);
                    table.Export(fileInfo, ExportOptions.WithDefaults);
                }
                catch (Exception exc)
                {
                    Console.WriteLine(exc.ToString());
                }
            } 
        }
        //Export TextLists 
        private static void ExportTextLists(HmiTarget hmitarget) 
        {    
            TextListComposition text = hmitarget.TextLists;    
            foreach (TextList textList in text)    
            {
                string extension = ".xml";
                var fileInfo = new FileInfo(exportLocation + @"\hmi_text_lists\xml\" + textList.Name + extension);
                try
                {
                    if (File.Exists(fileInfo.FullName)) File.Delete(fileInfo.FullName);
                    Console.WriteLine(textList.Name + " to " + fileInfo.FullName);
                    textList.Export(fileInfo, ExportOptions.WithDefaults);
                }
                catch (Exception exc)
                {
                    Console.WriteLine(exc.ToString());
                }
            } 
        }

        private static void RemoveDirectory(string path)
        {
            if (Directory.Exists(path))
            {
                Console.WriteLine("Deleting location " + path + ".");
                Directory.Delete(path,true);
            }
        }
        private static void CheckDirectory(string path)
        {
            if (!Directory.Exists(path))
            {
                Console.WriteLine("Export location "+path+" does not exist");
                Directory.CreateDirectory(path);
            }
        }                
        // Kept outside the export location so it never ends up in a repository:
        // %LOCALAPPDATA%\TiaExportBlocks\<hash of the export location>.state
        static string StatePath()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TiaExportBlocks");
            Directory.CreateDirectory(dir);
            using (var sha = System.Security.Cryptography.SHA1.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(exportLocation.ToLowerInvariant()));
                return Path.Combine(dir, BitConverter.ToString(hash).Replace("-", "").Substring(0, 16) + ".state");
            }
        }

        // First line: the TIA project path; then "stamp<TAB>file write ticks<TAB>file" per element.
        static void LoadState(string projectPath)
        {
            previousStamps.Clear();
            previousWrites.Clear();
            string path = StatePath();
            if (!File.Exists(path)) return;
            string[] lines = File.ReadAllLines(path);
            if (lines.Length == 0 || !string.Equals(lines[0], projectPath, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Last export was of another project: exporting everything");
                return;
            }
            for (int i = 1; i < lines.Length; i++)
            {
                string[] parts = lines[i].Split('\t');
                long stamp;
                long written;
                if (parts.Length == 3 && long.TryParse(parts[0], out stamp) && long.TryParse(parts[1], out written))
                {
                    previousStamps[parts[2]] = stamp;
                    previousWrites[parts[2]] = written;
                }
            }
        }

        static void SaveState(string projectPath)
        {
            var lines = new List<string>() { projectPath };
            foreach (var entry in currentStamps)
            {
                long written = File.Exists(entry.Key) ? File.GetLastWriteTimeUtc(entry.Key).Ticks : -1;
                lines.Add(entry.Value + "\t" + written + "\t" + entry.Key);
            }
            File.WriteAllLines(StatePath(), lines);
        }

        // Removes exported files whose block, type or tag table is no longer in the project.
        static int RemoveStaleFiles(string rootPrefix)
        {
            int removed = 0;
            string root = Path.Combine(exportLocation, rootPrefix);
            if (!Directory.Exists(root)) return 0;
            var extensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".scl", ".db", ".awl", ".udt", ".xml" };
            foreach (string sub in new[] { "scl", "db", "stl", "udt", @"tag_tables\xml" })
            {
                string dir = Path.Combine(root, sub);
                if (!Directory.Exists(dir)) continue;
                foreach (string file in Directory.GetFiles(dir, "*", SearchOption.AllDirectories))
                {
                    if (!extensions.Contains(Path.GetExtension(file))) continue;
                    if (expectedFiles.Contains(Path.GetFullPath(file))) continue;
                    Console.WriteLine("Removing " + file + " (no longer in the project)");
                    File.Delete(file);
                    removed++;
                }
            }
            return removed;
        }

            static void Main(string[] args)
        {

            fullExport = args != null && args.Any(a => string.Equals(a, "--full", StringComparison.OrdinalIgnoreCase));
            var paths = new List<string>();
            for (int i = 0; args != null && i < args.Length; i++)
            {
                int n;
                if (string.Equals(args[i], "--batch", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length && int.TryParse(args[i + 1], out n) && n > 0)
                {
                    batchSize = n;
                    i++;
                }
                else if (!args[i].StartsWith("--"))
                {
                    paths.Add(args[i]);
                }
            }
            if (paths.Count != 1)
            {
                Console.WriteLine("No arguments provided");
                Console.WriteLine(@"Usage TiaExportBlocks.exe C:\Test [--full] [--batch N]");
                Console.ReadLine();
            }
            else
            {
                exportLocation = Path.GetFullPath(paths[0]).TrimEnd('\\');
                Console.WriteLine("Export location is " + exportLocation + (fullExport ? " (full export)" : " (incremental export)") + ", batch size " + batchSize);
                if (fullExport) RemoveDirectory(exportLocation);
                CheckDirectory(exportLocation);
                var watch = new System.Diagnostics.Stopwatch();
                watch.Start();
                Console.WriteLine("Enumerating TIA processes..");
                foreach (TiaPortalProcess tiaPortalProcess in TiaPortal.GetProcesses())
                {
                    Console.WriteLine("Process ID " + tiaPortalProcess.Id);
                    Console.WriteLine("Project PATH " + tiaPortalProcess.ProjectPath);
                    TiaPortal tiaPortal = tiaPortalProcess.Attach();
                    // Check if attachment was successful
                    if (tiaPortal == null)
                    {
                        Console.WriteLine("Failed to attach to TIA Portal.");
                        return;
                    }
                    var projects = tiaPortal.Projects;
                    if (projects == null || projects.Count == 0)
                    {
                        Console.WriteLine("No projects found in TIA Portal.");
                    }
                    foreach (Project project in projects)
                    {
                        Console.WriteLine("Handling project " + project.Name);
                        string projectPath = project.Path.FullName;
                        LoadState(projectPath);
                        currentStamps.Clear();
                        expectedFiles.Clear();
                        var plcPrefixes = new List<string>();
                        foreach (Siemens.Engineering.HW.Device device in project.Devices)
                        {
                            Console.WriteLine("Handling device " + device.Name + " of type [" + device.TypeIdentifier+"]");
                            
                            if (device.TypeIdentifier == "System:Device.S71500")
                            {
                                foreach (Siemens.Engineering.HW.DeviceItem deviceItem in device.DeviceItems)
                                {
                                    Console.WriteLine("Handling device item " + deviceItem.Name + " of type " + deviceItem.TypeIdentifier);
                                    if (deviceItem.Name.Contains("PLC"))
                                    {
                                        Console.WriteLine("Handling PLC device item");
                                        Siemens.Engineering.HW.Features.SoftwareContainer softwareContainer = ((IEngineeringServiceProvider)deviceItem).GetService<SoftwareContainer>();
                                        string rootPrefix = deviceItem.Name;
                                        if (softwareContainer != null)
                                        {
                                            PlcSoftware software = softwareContainer.Software as PlcSoftware;
                                            plcPrefixes.Add(rootPrefix);
                                            ExportBlocks(software,rootPrefix);
                                            ExportTypes(software, rootPrefix);
                                            ExportAllTagTables(software, rootPrefix);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                foreach (Siemens.Engineering.HW.DeviceItem deviceItem in device.DeviceItems)
                                {
                                    Console.WriteLine("Handling device item [" + deviceItem.Name + "] of type [" + deviceItem.TypeIdentifier+"]");
                                    
                                        Console.WriteLine("Handling HMI device item");
                                        DeviceItem deviceItemToGetService = deviceItem as DeviceItem; 
                                        SoftwareContainer softwareContainer = deviceItemToGetService.GetService<SoftwareContainer>();
                                        if (softwareContainer != null)
                                        {
                                            HmiTarget software = softwareContainer.Software as HmiTarget;
                                            ExportAllTagTablesFromHMITarget(software);
                                            ExportTextLists(software);
                                        }
                                        else 
                                        { Console.WriteLine("SW Container is null"); 
                                        }
                                    
                                }
                            }
                        }
                        int removed = 0;
                        foreach (string prefix in plcPrefixes) removed += RemoveStaleFiles(prefix);
                        SaveState(projectPath);
                        Console.WriteLine("Exported " + exportedCount + ", unchanged " + skippedCount + ", removed " + removed);
                    }
                }
                watch.Stop();
                Console.WriteLine($"Execution Time: {watch.ElapsedMilliseconds} ms");
                Console.WriteLine("Done");
                //Console.ReadLine();
                return;
            }
        }
    }
}