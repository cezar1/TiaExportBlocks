using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

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

        /// <summary>ModifiedDate of every fake block/type until it is edited.</summary>
        public static readonly DateTime InitialDate = new DateTime(2026, 10, 1, 8, 0, 0, DateTimeKind.Utc);

        private readonly List<string> _calls = new List<string>();

        /// <summary>Like TIA, GenerateSource/Export fail if the target file already exists.</summary>
        public bool FailIfTargetExists = true;

        /// <summary>Strict mode: GenerateSource/Export fail if the target folder does not exist.</summary>
        public bool RequireExistingDirectory = true;

        /// <summary>Write generated sources with a UTF-8 BOM, like TIA Portal.</summary>
        public bool WriteUtf8Bom = true;

        /// <summary>
        /// Applied to GenerateSource output for more than one block/type, to simulate a TIA version whose
        /// multi-block output is formatted differently from single-block output.
        /// </summary>
        public Func<string, string> BatchOutputQuirk;

        /// <summary>Number of GenerateSource calls made.</summary>
        public int GenerateSourceCalls { get { return _calls.Count(c => c.StartsWith("GenerateSource ", StringComparison.Ordinal)); } }

        /// <summary>Every Openness call made, in order, e.g. "GenerateSource Main -> C:\...\Main.scl".</summary>
        public IReadOnlyList<string> Calls { get { return _calls; } }

        public IEnumerable<string> CallsStartingWith(string prefix)
        {
            return _calls.Where(c => c.StartsWith(prefix, StringComparison.Ordinal));
        }

        public void ClearCalls()
        {
            _calls.Clear();
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
            Encoding encoding = new UTF8Encoding(WriteUtf8Bom && call == "GenerateSource");
            if (fail != null)
            {
                if (writeBeforeFailing) File.WriteAllText(target.FullName, content.Substring(0, content.Length / 2), encoding);
                throw fail;
            }
            File.WriteAllText(target.FullName, content, encoding);
        }

        /// <summary>GenerateSource for a list of blocks or types: one file with all of them, like TIA.</summary>
        internal void GenerateSources(IList<string> names, IList<string> sources, IList<Exception> failures, bool writePartial, FileInfo target)
        {
            string content = string.Concat(sources);
            if (names.Count > 1 && BatchOutputQuirk != null) content = BatchOutputQuirk(content);
            Exception fail = failures.FirstOrDefault(f => f != null);
            WriteTarget("GenerateSource", string.Join(",", names), target, content, fail, writePartial);
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
            return new FakeProject(name, @"C:\Projects\" + name + @"\" + name + ".ap15_1", devices);
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
        public FakeProject(string name, string path, IEnumerable<FakeDevice> devices)
        {
            Name = name;
            Path = path;
            DeviceList = devices.ToList();
        }

        public string Name { get; private set; }
        public string Path { get; private set; }
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
        private readonly FakeTia _tia;

        internal FakePlcSoftware(FakeTia tia, string name)
        {
            _tia = tia;
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

        public void GenerateSource(IList<IPlcBlock> blocks, FileInfo target)
        {
            var fakes = blocks.Cast<FakeBlock>().ToList();
            foreach (var b in fakes) if (b.OnGenerate != null) b.OnGenerate();
            var unsupported = fakes.FirstOrDefault(b => !(b.ProgrammingLanguage == "SCL" || b.ProgrammingLanguage == "DB" || b.ProgrammingLanguage == "STL"));
            if (unsupported != null)
            {
                _tia.Record("GenerateSource " + string.Join(",", fakes.Select(b => b.Name)) + " -> " + target.FullName);
                throw new FakeOpennessException("Source generation is not supported for " + unsupported.ProgrammingLanguage + " blocks.");
            }
            _tia.GenerateSources(fakes.Select(b => b.Name).ToList(), fakes.Select(b => b.ExpectedSource).ToList(),
                fakes.Select(b => b.Fail).ToList(), fakes.Any(b => b.Fail != null && b.WritePartialFileBeforeFailing), target);
        }

        public void GenerateSource(IList<IPlcType> types, FileInfo target)
        {
            var fakes = types.Cast<FakeType>().ToList();
            _tia.GenerateSources(fakes.Select(t => t.Name).ToList(), fakes.Select(t => t.ExpectedSource).ToList(),
                fakes.Select(t => t.Fail).ToList(), false, target);
        }

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
        public FakeBlockGroup Remove(FakeBlock block) { _blocks.Remove(block); return this; }
        public FakeBlockGroup With(params FakeBlockGroup[] groups) { _groups.AddRange(groups); return this; }
    }

    public sealed class FakeBlock : IPlcBlock
    {
        internal FakeBlock(FakeTia tia, string name, string language)
        {
            Name = name;
            ProgrammingLanguage = language;
            ModifiedDate = FakeTia.InitialDate;
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
                        return "FUNCTION_BLOCK \"" + Name + "\"\r\n{ S7_Optimized_Access := 'TRUE' }\r\nVERSION : 0.1\r\n   VAR_INPUT \r\n      Enable : Bool;\r\n   END_VAR\r\n\r\nBEGIN\r\n\t" + Body + "\r\nEND_FUNCTION_BLOCK\r\n\r\n";
                }
            }
        }

        /// <summary>Runs during GenerateSource, e.g. to simulate another program writing into the export folder.</summary>
        public Action OnGenerate;

        /// <summary>TIA "ModifiedDate"; null simulates an attribute that cannot be read.</summary>
        public DateTime? ModifiedDate { get; set; }

        /// <summary>Extra line in the generated source, e.g. a tag name TIA shows (changes without touching ModifiedDate).</summary>
        public string Body = "#Enable := #Enable;";

        /// <summary>Simulates editing the block in TIA: new content and a newer ModifiedDate.</summary>
        public FakeBlock Edit(string newBody)
        {
            Body = newBody;
            ModifiedDate = (ModifiedDate ?? FakeTia.InitialDate).AddMinutes(1);
            return this;
        }
    }

    public sealed class FakeType : IPlcType
    {
        internal FakeType(FakeTia tia, string name)
        {
            Name = name;
            ModifiedDate = FakeTia.InitialDate;
        }

        public DateTime? ModifiedDate { get; set; }

        /// <summary>Extra member line; see <see cref="FakeBlock.Body"/>.</summary>
        public string Members = "Speed : Real;";

        public FakeType Edit(string members)
        {
            Members = members;
            ModifiedDate = (ModifiedDate ?? FakeTia.InitialDate).AddMinutes(1);
            return this;
        }

        public string Name { get; private set; }
        public Exception Fail;

        public FakeType Failing(string message) { Fail = new FakeOpennessException(message); return this; }

        public string ExpectedSource
        {
            get { return "TYPE \"" + Name + "\"\r\nVERSION : 0.1\r\n   STRUCT\r\n      " + Members + "\r\n      Running : Bool;\r\n   END_STRUCT;\r\n\r\nEND_TYPE\r\n\r\n"; }
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

        /// <summary>Tag list content; change it to simulate e.g. a renamed tag.</summary>
        public string Content = "<Tag>Start</Tag>";

        public FakeExportable Failing(string message) { Fail = new FakeOpennessException(message); return this; }

        public string ExpectedXml
        {
            get
            {
                return "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<Document>\r\n  <Engineering version=\"V15.1\" />\r\n"
                    + "  <" + _opennessType + " ID=\"0\">\r\n    <AttributeList>\r\n      <Name>" + Name + "</Name>\r\n    </AttributeList>\r\n    " + Content + "\r\n"
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
        public FakeHmiTarget RemoveTextList(FakeExportable textList) { _textLists.Remove(textList); return this; }
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
