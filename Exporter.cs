using System;
using System.Collections.Generic;
using System.IO;

namespace TiaExportBlocks
{
    /// <summary>
    /// Walks TIA Portal instances / projects / devices and exports PLC blocks, UDTs, PLC tag tables,
    /// HMI tag tables and HMI text lists. Works on the <see cref="ITiaPortalProcess"/> abstraction so it
    /// runs against the real Openness API (SiemensAdapters.cs) or against test fakes.
    /// </summary>
    public sealed class Exporter
    {
        private readonly string _exportLocation;
        private readonly TextWriter _log;

        public Exporter(string exportLocation, TextWriter log)
        {
            if (exportLocation == null) throw new ArgumentNullException("exportLocation");
            if (log == null) throw new ArgumentNullException("log");
            _exportLocation = exportLocation;
            _log = log;
        }

        public string ExportLocation { get { return _exportLocation; } }

        /// <summary>Deletes any previous export and recreates an empty export folder.</summary>
        public void PrepareExportLocation()
        {
            if (Directory.Exists(_exportLocation))
            {
                _log.WriteLine("Deleting location " + _exportLocation + ".");
            }
            FileOps.RemoveDirectory(_exportLocation);
            if (!Directory.Exists(_exportLocation))
            {
                _log.WriteLine("Export location " + _exportLocation + " does not exist");
            }
            FileOps.EnsureDirectory(_exportLocation);
        }

        /// <summary>Exports every open project of every TIA Portal instance.</summary>
        public void Run(IEnumerable<ITiaPortalProcess> processes)
        {
            _log.WriteLine("Enumerating TIA processes..");
            foreach (ITiaPortalProcess process in processes)
            {
                _log.WriteLine("Process ID " + process.Id);
                _log.WriteLine("Project PATH " + process.ProjectPath);
                ITiaPortal tiaPortal = process.Attach();
                if (tiaPortal == null)
                {
                    _log.WriteLine("Failed to attach to TIA Portal.");
                    return;
                }
                IList<ITiaProject> projects = tiaPortal.Projects;
                if (projects == null || projects.Count == 0)
                {
                    _log.WriteLine("No projects found in TIA Portal.");
                    continue;
                }
                foreach (ITiaProject project in projects)
                {
                    ExportProject(project);
                }
            }
        }

        public void ExportProject(ITiaProject project)
        {
            _log.WriteLine("Handling project " + project.Name);
            foreach (ITiaDevice device in project.Devices)
            {
                ExportDevice(device);
            }
        }

        public void ExportDevice(ITiaDevice device)
        {
            _log.WriteLine("Handling device " + device.Name + " of type [" + device.TypeIdentifier + "]");
            bool isPlcDevice = ExportRules.IsS71500(device.TypeIdentifier);
            foreach (ITiaDeviceItem deviceItem in device.DeviceItems)
            {
                try
                {
                    if (isPlcDevice)
                        ExportPlcDeviceItem(deviceItem);
                    else
                        ExportHmiDeviceItem(deviceItem);
                }
                catch (Exception exc)
                {
                    // One broken device item must not stop the export of everything else.
                    _log.WriteLine("Error handling device item [" + deviceItem.Name + "]: " + exc);
                }
            }
        }

        private void ExportPlcDeviceItem(ITiaDeviceItem deviceItem)
        {
            _log.WriteLine("Handling device item " + deviceItem.Name + " of type " + deviceItem.TypeIdentifier);
            if (!ExportRules.IsPlcDeviceItem(deviceItem.Name)) return;

            _log.WriteLine("Handling PLC device item");
            ITiaSoftware software = deviceItem.GetSoftware();
            if (software == null) return;
            IPlcSoftware plcSoftware = software as IPlcSoftware;
            if (plcSoftware == null)
            {
                _log.WriteLine("Device item [" + deviceItem.Name + "] has no PLC software, skipped");
                return;
            }
            ExportPlc(plcSoftware, deviceItem.Name);
        }

        private void ExportHmiDeviceItem(ITiaDeviceItem deviceItem)
        {
            _log.WriteLine("Handling device item [" + deviceItem.Name + "] of type [" + deviceItem.TypeIdentifier + "]");
            _log.WriteLine("Handling HMI device item");
            ITiaSoftware software = deviceItem.GetSoftware();
            if (software == null)
            {
                _log.WriteLine("SW Container is null");
                return;
            }
            IHmiTarget hmiTarget = software as IHmiTarget;
            if (hmiTarget == null)
            {
                // e.g. the CPU of an S7-1200 / S7-300 station: PLC software, but not an S7-1500 device.
                _log.WriteLine("Device item [" + deviceItem.Name + "] software is not an HMI target, skipped");
                return;
            }
            ExportHmi(hmiTarget, deviceItem.Name);
        }

        // ---------------- PLC ----------------

        public void ExportPlc(IPlcSoftware software, string rootPrefix)
        {
            _log.WriteLine(software.Name);
            HandleBlockGroup(software.BlockGroup, string.Empty, rootPrefix);
            ExportTypes(software, rootPrefix);
            ExportPlcTagTables(software.TagTableGroup, rootPrefix);
        }

        private void HandleBlockGroup(IPlcBlockGroup blockGroup, string currentPath, string rootPrefix)
        {
            string newPath = ExportPaths.GroupPath(currentPath, blockGroup.Name);
            try
            {
                _log.WriteLine("Handling block group: " + blockGroup.Name);
                foreach (IPlcBlock block in blockGroup.Blocks)
                {
                    HandleBlock(block, newPath, rootPrefix);
                }
                foreach (IPlcBlockGroup nestedGroup in blockGroup.Groups)
                {
                    HandleBlockGroup(nestedGroup, newPath, rootPrefix);
                }
            }
            catch (Exception ex)
            {
                _log.WriteLine("Error handling block group: " + blockGroup.Name + " - " + ex.Message);
            }
        }

        private void HandleBlock(IPlcBlock block, string currentPath, string rootPrefix)
        {
            string language = block.ProgrammingLanguage;
            string extension;
            string folderPrefix;
            if (!ExportRules.TryGetBlockFormat(language, out extension, out folderPrefix))
            {
                _log.WriteLine(block.Name + " programming language " + language + " not supported");
                return;
            }

            var fileInfo = new FileInfo(ExportPaths.BlockFile(_exportLocation, rootPrefix, folderPrefix, currentPath, block.Name, extension));
            var fileInfoTemp = new FileInfo(ExportPaths.BlockTempFile(_exportLocation, block.Name, extension));
            if (!fileInfo.Directory.Exists)
            {
                _log.WriteLine("Creating [" + fileInfo.FullName + "]");
                Directory.CreateDirectory(fileInfo.Directory.FullName);
            }
            try
            {
                FileOps.DeleteIfExists(fileInfo.FullName);
                FileOps.DeleteIfExists(fileInfoTemp.FullName);
                _log.WriteLine("RP [" + rootPrefix + "] Block [" + block.Name + "] prefix [" + currentPath + "] to " + fileInfo.FullName);
                block.GenerateSource(fileInfoTemp);
                FileOps.MoveReplacing(fileInfoTemp.FullName, fileInfo.FullName);
            }
            catch (Exception exc)
            {
                _log.WriteLine(exc.ToString());
                TryDelete(fileInfoTemp.FullName);
            }
        }

        private void ExportTypes(IPlcSoftware software, string rootPrefix)
        {
            foreach (IPlcType plcType in software.Types)
            {
                _log.WriteLine("Handling type at prefix " + rootPrefix + " named " + plcType.Name);
                HandleType(plcType, rootPrefix);
            }
        }

        private void HandleType(IPlcType plcType, string rootPrefix)
        {
            var fileInfo = new FileInfo(ExportPaths.TypeFile(_exportLocation, rootPrefix, plcType.Name));
            try
            {
                if (!fileInfo.Directory.Exists)
                {
                    Directory.CreateDirectory(fileInfo.Directory.FullName);
                }
                if (File.Exists(fileInfo.FullName))
                {
                    _log.WriteLine("Deleting " + fileInfo.FullName);
                    File.Delete(fileInfo.FullName);
                }
                plcType.GenerateSource(fileInfo);
            }
            catch (Exception exc)
            {
                _log.WriteLine(exc.ToString());
            }
        }

        private void ExportPlcTagTables(IPlcTagTableGroup group, string rootPrefix)
        {
            foreach (IExportableItem table in group.TagTables)
            {
                ExportPlcTagTable(table, rootPrefix);
            }
            foreach (IPlcTagTableGroup userGroup in group.Groups)
            {
                ExportPlcTagTables(userGroup, rootPrefix);
            }
        }

        private void ExportPlcTagTable(IExportableItem table, string rootPrefix)
        {
            _log.WriteLine(table.Name);
            var fileInfoTemp = new FileInfo(ExportPaths.PlcTagTableTempFile(_exportLocation, table.Name));
            try
            {
                FileOps.DeleteIfExists(fileInfoTemp.FullName);
                table.Export(fileInfoTemp);
                var fileInfo = new FileInfo(ExportPaths.PlcTagTableFile(_exportLocation, rootPrefix, table.Name));
                _log.WriteLine(table.Name + " to " + fileInfo.FullName);
                FileOps.MoveReplacing(fileInfoTemp.FullName, fileInfo.FullName);
            }
            catch (Exception exc)
            {
                _log.WriteLine(exc.ToString());
                TryDelete(fileInfoTemp.FullName);
            }
        }

        // ---------------- HMI ----------------

        public void ExportHmi(IHmiTarget hmiTarget, string hmiPrefix)
        {
            IHmiTagFolder systemFolder = hmiTarget.TagFolder;
            // First the tables in the user folders, then the tables in the system folder.
            foreach (IHmiTagFolder userFolder in systemFolder.Folders)
            {
                ExportHmiUserFolderDeep(userFolder, hmiPrefix);
            }
            ExportHmiTagTables(systemFolder, hmiPrefix);
            ExportTextLists(hmiTarget, hmiPrefix);
        }

        private void ExportHmiUserFolderDeep(IHmiTagFolder folder, string hmiPrefix)
        {
            foreach (IHmiTagFolder subFolder in folder.Folders)
            {
                ExportHmiUserFolderDeep(subFolder, hmiPrefix);
            }
            ExportHmiTagTables(folder, hmiPrefix);
        }

        private void ExportHmiTagTables(IHmiTagFolder folder, string hmiPrefix)
        {
            foreach (IExportableItem table in folder.TagTables)
            {
                ExportXml(table, ExportPaths.HmiTagTableFile(_exportLocation, hmiPrefix, table.Name));
            }
        }

        private void ExportTextLists(IHmiTarget hmiTarget, string hmiPrefix)
        {
            foreach (IExportableItem textList in hmiTarget.TextLists)
            {
                ExportXml(textList, ExportPaths.HmiTextListFile(_exportLocation, hmiPrefix, textList.Name));
            }
        }

        private void ExportXml(IExportableItem item, string path)
        {
            var fileInfo = new FileInfo(path);
            try
            {
                FileOps.EnsureDirectoryExists(fileInfo.FullName);
                FileOps.DeleteIfExists(fileInfo.FullName);
                _log.WriteLine(item.Name + " to " + fileInfo.FullName);
                item.Export(fileInfo);
            }
            catch (Exception exc)
            {
                _log.WriteLine(exc.ToString());
            }
        }

        private static void TryDelete(string path)
        {
            try { FileOps.DeleteIfExists(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
