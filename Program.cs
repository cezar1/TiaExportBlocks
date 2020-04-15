﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Siemens.Engineering;
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
        static Dictionary<string, string> programLanguageToExtension = new Dictionary<string, string> { { "SCL", ".scl" }, { "DB", ".db" } };
        static Dictionary<string, string> programLanguageToFolderPrefixExtension = new Dictionary<string, string> { { "SCL", @"\scl\" }, { "DB", @"\db\" } };
        static string exportLocation;
        static void HandleBlock(PlcBlock block,PlcSoftware software)
        {
            PlcExternalSourceSystemGroup externalSourceGroup = software.ExternalSourceGroup;
            //Console.WriteLine(block.Name + " " + block.GetType()+ " " +block.ProgrammingLanguage);
            if (programLanguageToExtension.ContainsKey(block.ProgrammingLanguage.ToString()))
            {
                string extension;
                programLanguageToExtension.TryGetValue(block.ProgrammingLanguage.ToString(),out extension);
                string folder_prefix;
                programLanguageToFolderPrefixExtension.TryGetValue(block.ProgrammingLanguage.ToString(), out folder_prefix);
                var fileInfo = new FileInfo(exportLocation + folder_prefix+block.Name + extension);
                var blocks = new List<PlcBlock>() { block };
                try
                {
                    if (File.Exists(fileInfo.FullName)) File.Delete(fileInfo.FullName);
                    Console.WriteLine(block.Name + " to " + fileInfo.FullName);
                    externalSourceGroup.GenerateSource(blocks, fileInfo, GenerateOptions.None);
                }
                catch (Exception exc)
                {
                    Console.WriteLine(exc.ToString());
                }
            }

        }
        static void HandleType(PlcType plcType, PlcSoftware software)
        {
            PlcExternalSourceSystemGroup externalSourceGroup = software.ExternalSourceGroup;
            string extension=".udt";
            var fileInfo = new FileInfo(exportLocation + @"\udt\"+plcType.Name + extension);
            var blocks = new List<PlcType>() { plcType };
            try
            {
                if (File.Exists(fileInfo.FullName)) File.Delete(fileInfo.FullName);
                externalSourceGroup.GenerateSource(blocks, fileInfo, GenerateOptions.None);
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.ToString());
            }

        }
        private static void ExportAllTagTables(PlcSoftware plcSoftware)
        {
            PlcTagTableSystemGroup plcTagTableSystemGroup = plcSoftware.TagTableGroup;
            // Export all tables in the system group
            ExportTagTables(plcTagTableSystemGroup.TagTables);
            // Export the tables in underlying user groups
            foreach (PlcTagTableUserGroup userGroup in plcTagTableSystemGroup.Groups)
            {
                ExportUserGroupDeep(userGroup);
            }
        }
        private static void ExportTagTables(PlcTagTableComposition tagTables)
        {
            foreach (PlcTagTable table in tagTables)
            {
                string filePath = exportLocation + @"\tag_tables\"+table.Name + ".xml";
                var fileInfo = new FileInfo(filePath);
                Console.WriteLine(table.Name+" to "+ fileInfo.FullName);
                if (File.Exists(fileInfo.FullName)) File.Delete(fileInfo.FullName);
                table.Export(fileInfo, ExportOptions.WithDefaults);
            }
        }
        private static void ExportUserGroupDeep(PlcTagTableUserGroup group)
        {
            ExportTagTables(group.TagTables);
            foreach (PlcTagTableUserGroup userGroup in group.Groups)
            {
                ExportUserGroupDeep(userGroup);
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
                CheckDirectory(exportLocation);
                CheckDirectory(exportLocation + @"\scl");
                CheckDirectory(exportLocation + @"\db");
                CheckDirectory(exportLocation + @"\udt");
                CheckDirectory(exportLocation + @"\tag_tables");
                var watch = new System.Diagnostics.Stopwatch();
                watch.Start();
                Console.WriteLine("Enumerating TIA processes..");
                foreach (TiaPortalProcess tiaPortalProcess in TiaPortal.GetProcesses())
                {
                    Console.WriteLine("Process ID " + tiaPortalProcess.Id);
                    Console.WriteLine("Project PATH " + tiaPortalProcess.ProjectPath);
                    TiaPortal tiaPortal = tiaPortalProcess.Attach();
                    foreach (Project project in tiaPortal.Projects)
                    {
                        Console.WriteLine("Handling project " + project.Name);
                        foreach (Siemens.Engineering.HW.Device device in project.Devices)
                        {
                            Console.WriteLine("Handling device " + device.Name + " of type " + device.TypeIdentifier);
                            if (device.TypeIdentifier == "System:Device.S71500")
                            {
                                foreach (Siemens.Engineering.HW.DeviceItem deviceItem in device.DeviceItems)
                                {
                                    Console.WriteLine("Handling device item " + deviceItem.Name + " of type " + deviceItem.TypeIdentifier);
                                    if (deviceItem.Name.Contains("PLC"))
                                    {
                                        Console.WriteLine("Handling PLC device item");
                                        Siemens.Engineering.HW.Features.SoftwareContainer softwareContainer = ((IEngineeringServiceProvider)deviceItem).GetService<SoftwareContainer>();
                                        if (softwareContainer != null)
                                        {
                                            PlcSoftware software = softwareContainer.Software as PlcSoftware;
                                            string name = software.Name;
                                            Console.WriteLine(name);
                                            foreach (PlcBlock block in software.BlockGroup.Blocks)
                                            {
                                                HandleBlock(block, software);
                                            }
                                            foreach (PlcBlockGroup blockGroup in software.BlockGroup.Groups)
                                            {
                                                Console.WriteLine("Handling block group " + blockGroup.Name);
                                                foreach (PlcBlock block in blockGroup.Blocks)
                                                {
                                                    HandleBlock(block, software);
                                                }
                                            }

                                            foreach (PlcType plcType in software.TypeGroup.Types)
                                            {
                                                Console.WriteLine("Handling type " + plcType.Name);
                                                HandleType(plcType, software);
                                            }
                                            ExportAllTagTables(software);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                watch.Stop();
                Console.WriteLine($"Execution Time: {watch.ElapsedMilliseconds} ms");
                Console.WriteLine("Done");
                Console.ReadLine();
                return;
            }
        }
    }
}