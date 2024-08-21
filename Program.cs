using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
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
        static void HandleBlock(PlcBlock block,PlcSoftware software, string currentPath, string rootPrefix)
        {
            PlcExternalSourceSystemGroup externalSourceGroup = software.ExternalSourceGroup;
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
                var fileInfoTemp = new FileInfo(exportLocation + MakeValidFileName(block.Name) + extension);
                //var fileInfo = new FileInfo(exportLocation + folder_prefix+ '\\' + sanitizedFolderPrefix + '\\' +MakeValidFileName(block.Name) + extension);
                if (!fileInfo.Directory.Exists)
                {
                    Console.WriteLine("Creating [" + fileInfo.FullName + "]");
                    Directory.CreateDirectory(fileInfo.Directory.FullName);
                }
                var blocks = new List<PlcBlock>() { block };
                try
                {
                    if (File.Exists(fileInfo.FullName)) File.Delete(fileInfo.FullName);
                    Console.WriteLine("RP ["+ rootPrefix + "] Block [" + block.Name + "] prefix [" + currentPath + "] to " + fileInfo.FullName);
                    externalSourceGroup.GenerateSource(blocks, fileInfoTemp, GenerateOptions.None);
                    File.Move(fileInfoTemp.FullName, fileInfo.FullName);
                }
                catch (Exception exc)
                {
                    Console.WriteLine(exc.ToString());
                }
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
            PlcExternalSourceSystemGroup externalSourceGroup = software.ExternalSourceGroup;
            string extension=".udt";
            var fileInfo = new FileInfo(exportLocation + @"\" + rootPrefix + @"\udt\" + plcType.Name + extension);
            var blocks = new List<PlcType>() { plcType };
            try
            {
                if (!fileInfo.Directory.Exists)
                {
                    Directory.CreateDirectory(fileInfo.Directory.FullName);
                }
                if (File.Exists(fileInfo.FullName))
                {
                    Console.WriteLine("Deleting " + fileInfo.FullName);
                    File.Delete(fileInfo.FullName);
                }
                externalSourceGroup.GenerateSource(blocks, fileInfo, GenerateOptions.None);
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.ToString());
            }

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

        }
        private static void ExportTypes(PlcSoftware software, string rootPrefix)
        {
            foreach (PlcType plcType in software.TypeGroup.Types)
            {
                Console.WriteLine("Handling type at prefix " + rootPrefix + " named "+ plcType.Name);
                HandleType(plcType, software, rootPrefix);
            }
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
                var fileInfoTemp = new FileInfo(exportLocation +@"\"+ MakeValidFileName(table.Name) + ".xml");
                if (File.Exists(fileInfoTemp.FullName)) File.Delete(fileInfoTemp.FullName);
                table.Export(fileInfoTemp, ExportOptions.WithDefaults);
                string filePath = exportLocation +@"\"+ rootPrefix+@"\tag_tables\xml\"  + @"\" + table.Name + ".xml";
                var fileInfo = new FileInfo(filePath);
                Console.WriteLine(table.Name + " to " + fileInfo.FullName);
                EnsureDirectoryExists(filePath);
                if (File.Exists(fileInfo.FullName)) File.Delete(fileInfo.FullName);
                File.Move(fileInfoTemp.FullName, fileInfo.FullName);
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
            static void Main(string[] args)
        {

            if (args == null || args.Length == 0 || args.Length>1)
            {
                Console.WriteLine("No arguments provided");
                Console.WriteLine(@"Usage TiaExportBlocks.exe C:\Test");
                Console.ReadLine();
            }
            else
            {
                exportLocation = args[0];
                Console.WriteLine("Export location is " + exportLocation);
                RemoveDirectory(exportLocation);
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