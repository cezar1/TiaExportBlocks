using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

// Thin wrappers that expose the real TIA Portal Openness objects through the interfaces in TiaModel.cs.
// This is the only file (besides Program.Main) that touches Siemens.Engineering; it holds no logic.
namespace TiaExportBlocks
{
    internal static class SiemensTia
    {
        /// <summary>All running TIA Portal instances. Enumerated lazily, like the original loop.</summary>
        public static IEnumerable<ITiaPortalProcess> GetProcesses()
        {
            foreach (TiaPortalProcess process in TiaPortal.GetProcesses())
            {
                yield return new TiaPortalProcessAdapter(process);
            }
        }
    }

    internal sealed class TiaPortalProcessAdapter : ITiaPortalProcess
    {
        private readonly TiaPortalProcess _process;
        public TiaPortalProcessAdapter(TiaPortalProcess process) { _process = process; }

        public int Id { get { return _process.Id; } }
        public string ProjectPath { get { return Convert.ToString(_process.ProjectPath); } }

        public ITiaPortal Attach()
        {
            TiaPortal tiaPortal = _process.Attach();
            return tiaPortal == null ? null : new TiaPortalAdapter(tiaPortal);
        }
    }

    internal sealed class TiaPortalAdapter : ITiaPortal
    {
        private readonly TiaPortal _tiaPortal;
        public TiaPortalAdapter(TiaPortal tiaPortal) { _tiaPortal = tiaPortal; }

        public IList<ITiaProject> Projects
        {
            get
            {
                var projects = _tiaPortal.Projects;
                if (projects == null) return null;
                return projects.Select(p => (ITiaProject)new ProjectAdapter(p)).ToList();
            }
        }
    }

    internal sealed class ProjectAdapter : ITiaProject
    {
        private readonly Project _project;
        public ProjectAdapter(Project project) { _project = project; }

        public string Name { get { return _project.Name; } }
        public string Path { get { return _project.Path == null ? null : _project.Path.FullName; } }
        public IEnumerable<ITiaDevice> Devices
        {
            get { return _project.Devices.Select(d => (ITiaDevice)new DeviceAdapter(d)); }
        }
    }

    internal sealed class DeviceAdapter : ITiaDevice
    {
        private readonly Siemens.Engineering.HW.Device _device;
        public DeviceAdapter(Siemens.Engineering.HW.Device device) { _device = device; }

        public string Name { get { return _device.Name; } }
        public string TypeIdentifier { get { return _device.TypeIdentifier; } }
        public IEnumerable<ITiaDeviceItem> DeviceItems
        {
            get { return _device.DeviceItems.Select(i => (ITiaDeviceItem)new DeviceItemAdapter(i)); }
        }
    }

    internal sealed class DeviceItemAdapter : ITiaDeviceItem
    {
        private readonly Siemens.Engineering.HW.DeviceItem _item;
        public DeviceItemAdapter(Siemens.Engineering.HW.DeviceItem item) { _item = item; }

        public string Name { get { return _item.Name; } }
        public string TypeIdentifier { get { return _item.TypeIdentifier; } }

        public ITiaSoftware GetSoftware()
        {
            SoftwareContainer softwareContainer = _item.GetService<SoftwareContainer>();
            if (softwareContainer == null) return null;
            object software = softwareContainer.Software;
            PlcSoftware plcSoftware = software as PlcSoftware;
            if (plcSoftware != null) return new PlcSoftwareAdapter(plcSoftware);
            HmiTarget hmiTarget = software as HmiTarget;
            if (hmiTarget != null) return new HmiTargetAdapter(hmiTarget, _item.Name);
            return new OtherSoftwareAdapter(software == null ? "(none)" : software.GetType().Name);
        }
    }

    internal sealed class OtherSoftwareAdapter : ITiaSoftware
    {
        public OtherSoftwareAdapter(string name) { Name = name; }
        public string Name { get; private set; }
    }

    // ---------------- PLC ----------------

    internal sealed class PlcSoftwareAdapter : IPlcSoftware
    {
        private readonly PlcSoftware _software;
        public PlcSoftwareAdapter(PlcSoftware software) { _software = software; }

        public string Name { get { return _software.Name; } }
        public IPlcBlockGroup BlockGroup { get { return new PlcBlockGroupAdapter(_software.BlockGroup, _software); } }
        public IEnumerable<IPlcType> Types
        {
            get { return _software.TypeGroup.Types.Select(t => (IPlcType)new PlcTypeAdapter(t, _software)); }
        }
        public IPlcTagTableGroup TagTableGroup { get { return new PlcTagTableSystemGroupAdapter(_software.TagTableGroup); } }

        public void GenerateSource(IList<IPlcBlock> blocks, FileInfo target)
        {
            List<PlcBlock> list = blocks.Select(b => ((PlcBlockAdapter)b).Block).ToList();
            _software.ExternalSourceGroup.GenerateSource(list, target, GenerateOptions.None);
        }

        public void GenerateSource(IList<IPlcType> types, FileInfo target)
        {
            List<PlcType> list = types.Select(t => ((PlcTypeAdapter)t).Type).ToList();
            _software.ExternalSourceGroup.GenerateSource(list, target, GenerateOptions.None);
        }
    }

    internal static class SiemensAttributes
    {
        /// <summary>The "ModifiedDate" attribute, or null when it cannot be read (the element is then always exported).</summary>
        public static DateTime? ModifiedDate(IEngineeringObject engineeringObject)
        {
            try
            {
                object value = engineeringObject.GetAttribute("ModifiedDate");
                if (value is DateTime) return (DateTime)value;
            }
            catch (Exception)
            {
            }
            return null;
        }
    }

    internal sealed class PlcBlockGroupAdapter : IPlcBlockGroup
    {
        private readonly PlcBlockGroup _group;
        private readonly PlcSoftware _software;
        public PlcBlockGroupAdapter(PlcBlockGroup group, PlcSoftware software) { _group = group; _software = software; }

        public string Name { get { return _group.Name; } }
        public IEnumerable<IPlcBlock> Blocks
        {
            get { return _group.Blocks.Select(b => (IPlcBlock)new PlcBlockAdapter(b, _software)); }
        }
        public IEnumerable<IPlcBlockGroup> Groups
        {
            get { return _group.Groups.Select(g => (IPlcBlockGroup)new PlcBlockGroupAdapter(g, _software)); }
        }
    }

    internal sealed class PlcBlockAdapter : IPlcBlock
    {
        private readonly PlcBlock _block;
        public PlcBlockAdapter(PlcBlock block, PlcSoftware software) { _block = block; }

        public PlcBlock Block { get { return _block; } }
        public string Name { get { return _block.Name; } }
        public string ProgrammingLanguage { get { return _block.ProgrammingLanguage.ToString(); } }
        public DateTime? ModifiedDate { get { return SiemensAttributes.ModifiedDate(_block); } }
    }

    internal sealed class PlcTypeAdapter : IPlcType
    {
        private readonly PlcType _type;
        public PlcTypeAdapter(PlcType type, PlcSoftware software) { _type = type; }

        public PlcType Type { get { return _type; } }
        public string Name { get { return _type.Name; } }
        public DateTime? ModifiedDate { get { return SiemensAttributes.ModifiedDate(_type); } }
    }

    internal sealed class PlcTagTableSystemGroupAdapter : IPlcTagTableGroup
    {
        private readonly PlcTagTableSystemGroup _group;
        public PlcTagTableSystemGroupAdapter(PlcTagTableSystemGroup group) { _group = group; }

        public IEnumerable<IExportableItem> TagTables
        {
            get { return _group.TagTables.Select(t => (IExportableItem)new PlcTagTableAdapter(t)); }
        }
        public IEnumerable<IPlcTagTableGroup> Groups
        {
            get { return _group.Groups.Select(g => (IPlcTagTableGroup)new PlcTagTableUserGroupAdapter(g)); }
        }
    }

    internal sealed class PlcTagTableUserGroupAdapter : IPlcTagTableGroup
    {
        private readonly PlcTagTableUserGroup _group;
        public PlcTagTableUserGroupAdapter(PlcTagTableUserGroup group) { _group = group; }

        public IEnumerable<IExportableItem> TagTables
        {
            get { return _group.TagTables.Select(t => (IExportableItem)new PlcTagTableAdapter(t)); }
        }
        public IEnumerable<IPlcTagTableGroup> Groups
        {
            get { return _group.Groups.Select(g => (IPlcTagTableGroup)new PlcTagTableUserGroupAdapter(g)); }
        }
    }

    internal sealed class PlcTagTableAdapter : IExportableItem
    {
        private readonly PlcTagTable _table;
        public PlcTagTableAdapter(PlcTagTable table) { _table = table; }

        public string Name { get { return _table.Name; } }
        public void Export(FileInfo target) { _table.Export(target, ExportOptions.WithDefaults); }
    }

    // ---------------- HMI ----------------

    internal sealed class HmiTargetAdapter : IHmiTarget
    {
        private readonly HmiTarget _target;
        public HmiTargetAdapter(HmiTarget target, string name) { _target = target; Name = name; }

        public string Name { get; private set; }
        public IHmiTagFolder TagFolder { get { return new HmiTagSystemFolderAdapter(_target.TagFolder); } }
        public IEnumerable<IExportableItem> TextLists
        {
            get { return _target.TextLists.Select(l => (IExportableItem)new HmiTextListAdapter(l)); }
        }
    }

    internal sealed class HmiTagSystemFolderAdapter : IHmiTagFolder
    {
        private readonly TagSystemFolder _folder;
        public HmiTagSystemFolderAdapter(TagSystemFolder folder) { _folder = folder; }

        public IEnumerable<IExportableItem> TagTables
        {
            get { return _folder.TagTables.Select(t => (IExportableItem)new HmiTagTableAdapter(t)); }
        }
        public IEnumerable<IHmiTagFolder> Folders
        {
            get { return _folder.Folders.Select(f => (IHmiTagFolder)new HmiTagUserFolderAdapter(f)); }
        }
    }

    internal sealed class HmiTagUserFolderAdapter : IHmiTagFolder
    {
        private readonly TagUserFolder _folder;
        public HmiTagUserFolderAdapter(TagUserFolder folder) { _folder = folder; }

        public IEnumerable<IExportableItem> TagTables
        {
            get { return _folder.TagTables.Select(t => (IExportableItem)new HmiTagTableAdapter(t)); }
        }
        public IEnumerable<IHmiTagFolder> Folders
        {
            get { return _folder.Folders.Select(f => (IHmiTagFolder)new HmiTagUserFolderAdapter(f)); }
        }
    }

    internal sealed class HmiTagTableAdapter : IExportableItem
    {
        private readonly TagTable _table;
        public HmiTagTableAdapter(TagTable table) { _table = table; }

        public string Name { get { return _table.Name; } }
        public void Export(FileInfo target) { _table.Export(target, ExportOptions.WithDefaults); }
    }

    internal sealed class HmiTextListAdapter : IExportableItem
    {
        private readonly TextList _textList;
        public HmiTextListAdapter(TextList textList) { _textList = textList; }

        public string Name { get { return _textList.Name; } }
        public void Export(FileInfo target) { _textList.Export(target, ExportOptions.WithDefaults); }
    }
}
