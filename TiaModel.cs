using System;
using System.Collections.Generic;
using System.IO;

// The slice of the TIA Portal Openness object model this tool uses, as plain interfaces.
// Production code wraps the real Siemens.Engineering objects (SiemensAdapters.cs);
// tests plug in fakes that simulate TIA, so the whole export run can be tested without TIA Portal.
namespace TiaExportBlocks
{
    /// <summary>A running TIA Portal instance (TiaPortalProcess).</summary>
    public interface ITiaPortalProcess
    {
        int Id { get; }
        string ProjectPath { get; }

        /// <summary>Attaches to the instance. Returns null if attaching failed.</summary>
        ITiaPortal Attach();
    }

    /// <summary>An attached TIA Portal (TiaPortal).</summary>
    public interface ITiaPortal
    {
        /// <summary>Open projects; may be null or empty.</summary>
        IList<ITiaProject> Projects { get; }
    }

    public interface ITiaProject
    {
        string Name { get; }

        /// <summary>Full path of the project file; identifies the project in the incremental export state.</summary>
        string Path { get; }
        IEnumerable<ITiaDevice> Devices { get; }
    }

    public interface ITiaDevice
    {
        string Name { get; }
        string TypeIdentifier { get; }
        IEnumerable<ITiaDeviceItem> DeviceItems { get; }
    }

    public interface ITiaDeviceItem
    {
        string Name { get; }
        string TypeIdentifier { get; }

        /// <summary>
        /// Software of the item's SoftwareContainer: an <see cref="IPlcSoftware"/>, an <see cref="IHmiTarget"/>,
        /// some other <see cref="ITiaSoftware"/>, or null when the item has no software container.
        /// </summary>
        ITiaSoftware GetSoftware();
    }

    /// <summary>Base for anything a SoftwareContainer can hold.</summary>
    public interface ITiaSoftware
    {
        string Name { get; }
    }

    public interface IPlcSoftware : ITiaSoftware
    {
        IPlcBlockGroup BlockGroup { get; }
        IEnumerable<IPlcType> Types { get; }
        IPlcTagTableGroup TagTableGroup { get; }

        /// <summary>ExternalSourceGroup.GenerateSource for one or more blocks into a single file.</summary>
        void GenerateSource(IList<IPlcBlock> blocks, FileInfo target);

        /// <summary>ExternalSourceGroup.GenerateSource for one or more PLC data types into a single file.</summary>
        void GenerateSource(IList<IPlcType> types, FileInfo target);
    }

    public interface IPlcBlockGroup
    {
        string Name { get; }
        IEnumerable<IPlcBlock> Blocks { get; }
        IEnumerable<IPlcBlockGroup> Groups { get; }
    }

    public interface IPlcBlock
    {
        string Name { get; }

        /// <summary>ProgrammingLanguage.ToString(), e.g. "SCL", "DB", "LAD".</summary>
        string ProgrammingLanguage { get; }

        /// <summary>The "ModifiedDate" attribute, or null if it cannot be read.</summary>
        DateTime? ModifiedDate { get; }

        /// <summary>Know-how protected blocks can't be generated as source; false if unknown.</summary>
        bool IsKnowHowProtected { get; }
    }

    /// <summary>A PLC data type (UDT).</summary>
    public interface IPlcType
    {
        string Name { get; }

        /// <summary>The "ModifiedDate" attribute, or null if it cannot be read.</summary>
        DateTime? ModifiedDate { get; }
    }

    /// <summary>PLC tag table system group or user group.</summary>
    public interface IPlcTagTableGroup
    {
        IEnumerable<IExportableItem> TagTables { get; }
        IEnumerable<IPlcTagTableGroup> Groups { get; }
    }

    public interface IHmiTarget : ITiaSoftware
    {
        IHmiTagFolder TagFolder { get; }
        IEnumerable<IExportableItem> TextLists { get; }
    }

    /// <summary>HMI tag system folder or user folder.</summary>
    public interface IHmiTagFolder
    {
        IEnumerable<IExportableItem> TagTables { get; }
        IEnumerable<IHmiTagFolder> Folders { get; }
    }

    /// <summary>Anything exported as Openness XML (PLC tag table, HMI tag table, text list).</summary>
    public interface IExportableItem
    {
        string Name { get; }

        /// <summary>Export(file, ExportOptions.WithDefaults).</summary>
        void Export(FileInfo target);
    }
}
