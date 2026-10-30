using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace TiaExportBlocks.Tests.Fakes
{
    /// <summary>Stands in for Siemens.Engineering.EngineeringTargetInvocationException and friends.</summary>
    public sealed class FakeOpennessException : Exception
    {
        public FakeOpennessException(string message) : base(message) { }
    }

    /// <summary>
    /// A simulated TIA Portal. Builds a fake project tree, records every Openness call in
    /// <see cref="Calls"/> (in order), and writes realistic source/XML files when asked to
    /// generate or export.
    /// </summary>
    public sealed class FakeTia
    {
        public const string S71500 = "System:Device.S71500";
        public const string S71200 = "System:Device.S71200";
        public const string HmiPanel = "System:Device.HMI";

        private readonly List<string> _calls = new List<string>();

        /// <summary>Like TIA, GenerateSource/Export fail if the target file already exists.</summary>
        public bool FailIfTargetExists = true;

        /// <summary>Strict mode: GenerateSource/Export fail if the target folder does not exist.</summary>
        public bool RequireExistingDirectory = true;

        /// <summary>Every Openness call made, in order, e.g. "GenerateSource Main -> C:\...\Main.scl".</summary>
        public IReadOnlyList<string> Calls { get { return _calls; } }

        public IEnumerable<string> CallsStartingWith(string prefix)
        {
            return _calls.Where(c => c.StartsWith(prefix, StringComparison.Ordinal));
        }

        internal void Record(string call)
        {
            _calls.Add(call);
        }

        /// <summary>Common behaviour of GenerateSource / Export.</summary>
        internal void WriteTarget(string call, string itemName, FileInfo target, string content, Exception fail, bool writeBeforeFailing)
        {
            Record(call + " " + itemName + " -> " + target.FullName);
            if (RequireExistingDirectory && !Directory.Exists(target.DirectoryName))
                throw new FakeOpennessException("Directory '" + target.DirectoryName + "' does not exist.");
            if (FailIfTargetExists && File.Exists(target.FullName))
                throw new FakeOpennessException("The file '" + target.FullName + "' already exists.");
            if (fail != null)
            {
                if (writeBeforeFailing) File.WriteAllText(target.FullName, content.Substring(0, content.Length / 2));
                throw fail;
            }
            File.WriteAllText(target.FullName, content);
        }

        // ---------------- builders ----------------

        public FakeProcess Process(int id, params FakeProject[] projects)
        {
            return new FakeProcess(this, id, @"C:\Projects\Project" + id + ".ap15_1", new FakePortal(projects.Cast<ITiaProject>().ToList()));
        }

        public FakeProcess ProcessThatCannotAttach(int id)
        {
            return new FakeProcess(this, id, @"C:\Projects\Locked.ap15_1", null);
        }

        public FakeProcess ProcessWithoutProjects(int id)
        {
            return new FakeProcess(this, id, null, new FakePortal(new List<ITiaProject>()));
        }

        public FakeProject Project(string name, params FakeDevice[] devices)
        {
            return new FakeProject(name, devices);
        }

        /// <summary>An S7-1500 station whose CPU item (named <paramref name="plcItemName"/>) holds the software.</summary>
        public FakeDevice S71500Station(string stationName, string plcItemName, FakePlcSoftware software)
        {
            return new FakeDevice(stationName, S71500,
                new FakeDeviceItem(this, "Rail_0", "System:Rail", null),
                new FakeDeviceItem(this, plcItemName, "OrderNumber:6ES7 516-3AN01-0AB0/V2.6", software));
        }

        public FakeDevice Device(string name, string typeIdentifier, params FakeDeviceItem[] items)
        {
            return new FakeDevice(name, typeIdentifier, items);
        }

        public FakeDeviceItem Item(string name, ITiaSoftware software)
        {
            return new FakeDeviceItem(this, name, "OrderNumber:fake", software);
        }

        /// <summary>An HMI station with one item holding the HMI target.</summary>
        public FakeDevice HmiStation(string stationName, string hmiItemName, FakeHmiTarget target)
        {
            return new FakeDevice(stationName, HmiPanel,
                new FakeDeviceItem(this, hmiItemName, "OrderNumber:6AV2 124-0MC01-0AX0/15.1.0.0", target));
        }

        public FakePlcSoftware Plc(string name)
        {
            return new FakePlcSoftware(this, name);
        }

        public FakeBlockGroup Group(string name)
        {
            return new FakeBlockGroup(name);
        }

        public FakeBlock Block(string name, string language)
        {
            return new FakeBlock(this, name, language);
        }

        public FakeBlock Scl(string name) { return Block(name, "SCL"); }
        public FakeBlock Db(string name) { return Block(name, "DB"); }
        public FakeBlock Stl(string name) { return Block(name, "STL"); }
        public FakeBlock Lad(string name) { return Block(name, "LAD"); }

        public FakeType Udt(string name)
        {
            return new FakeType(this, name);
        }

        public FakeTagTableGroup TagGroup()
        {
            return new FakeTagTableGroup();
        }

        public FakeExportable PlcTagTable(string name) { return new FakeExportable(this, name, "SW.Tags.PlcTagTable"); }
        public FakeExportable HmiTagTable(string name) { return new FakeExportable(this, name, "Hmi.Tag.TagTable"); }
        public FakeExportable TextList(string name) { return new FakeExportable(this, name, "Hmi.TextGraphicList.TextList"); }

        public FakeHmiTarget Hmi(string name)
        {
            return new FakeHmiTarget(name);
        }

        public FakeHmiTagFolder TagFolder()
        {
            return new FakeHmiTagFolder();
        }
    }

    public sealed class FakeProcess : ITiaPortalProcess
    {
        private readonly FakeTia _tia;
        private readonly ITiaPortal _portal;

        internal FakeProcess(FakeTia tia, int id, string projectPath, ITiaPortal portal)
        {
            _tia = tia;
            Id = id;
            ProjectPath = projectPath;
            _portal = portal;
        }

        public int Id { get; private set; }
        public string ProjectPath { get; private set; }

        public ITiaPortal Attach()
        {
            _tia.Record("Attach " + Id);
            return _portal;
        }
    }

    public sealed class FakePortal : ITiaPortal
    {
        public FakePortal(IList<ITiaProject> projects) { Projects = projects; }
        public IList<ITiaProject> Projects { get; private set; }
    }

    public sealed class FakeProject : ITiaProject
    {
        public FakeProject(string name, IEnumerable<FakeDevice> devices)
        {
            Name = name;
            DeviceList = devices.ToList();
        }

        public string Name { get; private set; }
        public List<FakeDevice> DeviceList { get; private set; }
        public IEnumerable<ITiaDevice> Devices { get { return DeviceList; } }
    }

    public sealed class FakeDevice : ITiaDevice
    {
        public FakeDevice(string name, string typeIdentifier, params FakeDeviceItem[] items)
        {
            Name = name;
            TypeIdentifier = typeIdentifier;
            Items = items.ToList();
        }

        public string Name { get; private set; }
        public string TypeIdentifier { get; private set; }
        public List<FakeDeviceItem> Items { get; private set; }
        public IEnumerable<ITiaDeviceItem> DeviceItems { get { return Items; } }
    }

    public sealed class FakeDeviceItem : ITiaDeviceItem
    {
        private readonly FakeTia _tia;
        private readonly ITiaSoftware _software;

        internal FakeDeviceItem(FakeTia tia, string name, string typeIdentifier, ITiaSoftware software)
        {
            _tia = tia;
            Name = name;
            TypeIdentifier = typeIdentifier;
            _software = software;
        }

        public string Name { get; private set; }
        public string TypeIdentifier { get; private set; }

        /// <summary>Makes GetService&lt;SoftwareContainer&gt; throw.</summary>
        public Exception FailGetSoftware;

        public ITiaSoftware GetSoftware()
        {
            _tia.Record("GetSoftware " + Name);
            if (FailGetSoftware != null) throw FailGetSoftware;
            return _software;
        }
    }

    /// <summary>Software that is neither PLC nor HMI (or PLC software reached through a non-S7-1500 device).</summary>
    public sealed class FakeOtherSoftware : ITiaSoftware
    {
        public FakeOtherSoftware(string name) { Name = name; }
        public string Name { get; private set; }
    }

    // ---------------- PLC ----------------

    public sealed class FakePlcSoftware : IPlcSoftware
    {
        internal FakePlcSoftware(FakeTia tia, string name)
        {
            Name = name;
            Root = new FakeBlockGroup("Program blocks");
            TypeList = new List<FakeType>();
            Tags = new FakeTagTableGroup();
        }

        public string Name { get; private set; }

        /// <summary>The "Program blocks" system group.</summary>
        public FakeBlockGroup Root { get; private set; }
        public List<FakeType> TypeList { get; private set; }
        public FakeTagTableGroup Tags { get; private set; }

        public IPlcBlockGroup BlockGroup { get { return Root; } }
        public IEnumerable<IPlcType> Types { get { return TypeList; } }
        public IPlcTagTableGroup TagTableGroup { get { return Tags; } }

        public FakePlcSoftware WithBlocks(params FakeBlock[] blocks) { Root.With(blocks); return this; }
        public FakePlcSoftware WithGroups(params FakeBlockGroup[] groups) { Root.With(groups); return this; }
        public FakePlcSoftware WithTypes(params FakeType[] types) { TypeList.AddRange(types); return this; }
        public FakePlcSoftware WithTagTables(params FakeExportable[] tables) { Tags.With(tables); return this; }
        public FakePlcSoftware WithTagGroups(params FakeTagTableGroup[] groups) { Tags.With(groups); return this; }
    }

    public sealed class FakeBlockGroup : IPlcBlockGroup
    {
        private readonly List<FakeBlock> _blocks = new List<FakeBlock>();
        private readonly List<FakeBlockGroup> _groups = new List<FakeBlockGroup>();

        internal FakeBlockGroup(string name) { Name = name; }

        public string Name { get; private set; }

        /// <summary>Makes enumerating this group's blocks throw (e.g. object disposed while exporting).</summary>
        public Exception FailEnumeration;

        public IEnumerable<IPlcBlock> Blocks
        {
            get
            {
                if (FailEnumeration != null) throw FailEnumeration;
                return _blocks;
            }
        }

        public IEnumerable<IPlcBlockGroup> Groups { get { return _groups; } }

        public FakeBlockGroup With(params FakeBlock[] blocks) { _blocks.AddRange(blocks); return this; }
        public FakeBlockGroup With(params FakeBlockGroup[] groups) { _groups.AddRange(groups); return this; }
    }

    public sealed class FakeBlock : IPlcBlock
    {
        private readonly FakeTia _tia;

        internal FakeBlock(FakeTia tia, string name, string language)
        {
            _tia = tia;
            Name = name;
            ProgrammingLanguage = language;
        }

        public string Name { get; private set; }
        public string ProgrammingLanguage { get; private set; }

        /// <summary>Makes GenerateSource throw, e.g. know-how protected or inconsistent block.</summary>
        public Exception Fail;

        /// <summary>When failing, first write half a file (TIA aborting mid-way).</summary>
        public bool WritePartialFileBeforeFailing;

        public FakeBlock Failing(string message, bool writePartialFile = false)
        {
            Fail = new FakeOpennessException(message);
            WritePartialFileBeforeFailing = writePartialFile;
            return this;
        }

        public string ExpectedSource
        {
            get
            {
                switch (ProgrammingLanguage)
                {
                    case "DB":
                        return "DATA_BLOCK \"" + Name + "\"\r\n{ S7_Optimized_Access := 'TRUE' }\r\nVERSION : 0.1\r\nNON_RETAIN\r\n   VAR \r\n      Counter : Int;\r\n   END_VAR\r\n\r\nBEGIN\r\n\r\nEND_DATA_BLOCK\r\n\r\n";
                    case "STL":
                        return "FUNCTION \"" + Name + "\" : Void\r\nVERSION : 0.1\r\n\r\nBEGIN\r\nNETWORK\r\nTITLE = \r\n      A     \"Start\";\r\n      =     \"Run\";\r\nEND_FUNCTION\r\n\r\n";
                    default:
                        return "FUNCTION_BLOCK \"" + Name + "\"\r\n{ S7_Optimized_Access := 'TRUE' }\r\nVERSION : 0.1\r\n   VAR_INPUT \r\n      Enable : Bool;\r\n   END_VAR\r\n\r\nBEGIN\r\n\t#Enable := #Enable;\r\nEND_FUNCTION_BLOCK\r\n\r\n";
                }
            }
        }

        public void GenerateSource(FileInfo target)
        {
            if (!(ProgrammingLanguage == "SCL" || ProgrammingLanguage == "DB" || ProgrammingLanguage == "STL"))
            {
                _tia.Record("GenerateSource " + Name + " -> " + target.FullName);
                throw new FakeOpennessException("Source generation is not supported for " + ProgrammingLanguage + " blocks.");
            }
            _tia.WriteTarget("GenerateSource", Name, target, ExpectedSource, Fail, WritePartialFileBeforeFailing);
        }
    }

    public sealed class FakeType : IPlcType
    {
        private readonly FakeTia _tia;

        internal FakeType(FakeTia tia, string name)
        {
            _tia = tia;
            Name = name;
        }

        public string Name { get; private set; }
        public Exception Fail;

        public FakeType Failing(string message) { Fail = new FakeOpennessException(message); return this; }

        public string ExpectedSource
        {
            get { return "TYPE \"" + Name + "\"\r\nVERSION : 0.1\r\n   STRUCT\r\n      Speed : Real;\r\n      Running : Bool;\r\n   END_STRUCT;\r\n\r\nEND_TYPE\r\n\r\n"; }
        }

        public void GenerateSource(FileInfo target)
        {
            _tia.WriteTarget("GenerateSource", Name, target, ExpectedSource, Fail, false);
        }
    }

    public sealed class FakeTagTableGroup : IPlcTagTableGroup
    {
        private readonly List<FakeExportable> _tables = new List<FakeExportable>();
        private readonly List<FakeTagTableGroup> _groups = new List<FakeTagTableGroup>();

        public IEnumerable<IExportableItem> TagTables { get { return _tables; } }
        public IEnumerable<IPlcTagTableGroup> Groups { get { return _groups; } }

        public FakeTagTableGroup With(params FakeExportable[] tables) { _tables.AddRange(tables); return this; }
        public FakeTagTableGroup With(params FakeTagTableGroup[] groups) { _groups.AddRange(groups); return this; }
    }

    /// <summary>PLC tag table, HMI tag table or HMI text list.</summary>
    public sealed class FakeExportable : IExportableItem
    {
        private readonly FakeTia _tia;
        private readonly string _opennessType;

        internal FakeExportable(FakeTia tia, string name, string opennessType)
        {
            _tia = tia;
            Name = name;
            _opennessType = opennessType;
        }

        public string Name { get; private set; }
        public Exception Fail;

        public FakeExportable Failing(string message) { Fail = new FakeOpennessException(message); return this; }

        public string ExpectedXml
        {
            get
            {
                return "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<Document>\r\n  <Engineering version=\"V15.1\" />\r\n"
                    + "  <" + _opennessType + " ID=\"0\">\r\n    <AttributeList>\r\n      <Name>" + Name + "</Name>\r\n    </AttributeList>\r\n"
                    + "  </" + _opennessType + ">\r\n</Document>";
            }
        }

        public void Export(FileInfo target)
        {
            _tia.WriteTarget("Export", Name, target, ExpectedXml, Fail, false);
        }
    }

    // ---------------- HMI ----------------

    public sealed class FakeHmiTarget : IHmiTarget
    {
        private readonly List<FakeExportable> _textLists = new List<FakeExportable>();

        internal FakeHmiTarget(string name)
        {
            Name = name;
            SystemFolder = new FakeHmiTagFolder();
        }

        public string Name { get; private set; }
        public FakeHmiTagFolder SystemFolder { get; private set; }

        public IHmiTagFolder TagFolder { get { return SystemFolder; } }
        public IEnumerable<IExportableItem> TextLists { get { return _textLists; } }

        public FakeHmiTarget WithTagTables(params FakeExportable[] tables) { SystemFolder.With(tables); return this; }
        public FakeHmiTarget WithTagFolders(params FakeHmiTagFolder[] folders) { SystemFolder.With(folders); return this; }
        public FakeHmiTarget WithTextLists(params FakeExportable[] textLists) { _textLists.AddRange(textLists); return this; }
    }

    public sealed class FakeHmiTagFolder : IHmiTagFolder
    {
        private readonly List<FakeExportable> _tables = new List<FakeExportable>();
        private readonly List<FakeHmiTagFolder> _folders = new List<FakeHmiTagFolder>();

        public IEnumerable<IExportableItem> TagTables { get { return _tables; } }
        public IEnumerable<IHmiTagFolder> Folders { get { return _folders; } }

        public FakeHmiTagFolder With(params FakeExportable[] tables) { _tables.AddRange(tables); return this; }
        public FakeHmiTagFolder With(params FakeHmiTagFolder[] folders) { _folders.AddRange(folders); return this; }
    }
}
