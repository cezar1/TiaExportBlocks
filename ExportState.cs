using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace TiaExportBlocks
{
    /// <summary>What was exported to one file: the TIA modification stamp of its element and a hash of the file content.</summary>
    public sealed class StateEntry
    {
        /// <summary>No TIA stamp: element without ModifiedDate (tag tables, text lists) or a failed export to retry.</summary>
        public const long NoStamp = -1;

        public StateEntry(long stamp, string hash)
        {
            Stamp = stamp;
            Hash = hash;
        }

        public long Stamp { get; private set; }
        public string Hash { get; private set; }
    }

    /// <summary>
    /// Incremental export state for one export location: per TIA project, every file the tool wrote with
    /// the stamp and content hash at that time. Kept outside the export location
    /// (%LOCALAPPDATA%\TiaExportBlocks\&lt;hash of export location&gt;.state) so it never ends up in a repository.
    /// </summary>
    public sealed class ExportState
    {
        private const string Header = "TiaExportBlocks state v2";

        private readonly Dictionary<string, Dictionary<string, StateEntry>> _projects =
            new Dictionary<string, Dictionary<string, StateEntry>>(StringComparer.OrdinalIgnoreCase);

        public static string DefaultDirectory
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TiaExportBlocks"); }
        }

        /// <summary>State file for an export location; one file per location.</summary>
        public static string FileFor(string stateDirectory, string exportLocation)
        {
            string key = Path.GetFullPath(exportLocation).TrimEnd('\\', '/').ToLowerInvariant();
            using (var sha = SHA1.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(key));
                return Path.Combine(stateDirectory, BitConverter.ToString(hash).Replace("-", "").Substring(0, 16) + ".state");
            }
        }

        public IEnumerable<string> ProjectKeys { get { return _projects.Keys; } }

        /// <summary>Copy of the entries of a project (empty if unknown).</summary>
        public Dictionary<string, StateEntry> GetProject(string projectKey)
        {
            Dictionary<string, StateEntry> entries;
            var copy = new Dictionary<string, StateEntry>(StringComparer.OrdinalIgnoreCase);
            if (_projects.TryGetValue(projectKey, out entries))
            {
                foreach (var e in entries) copy[e.Key] = e.Value;
            }
            return copy;
        }

        public void SetProject(string projectKey, IDictionary<string, StateEntry> entries)
        {
            _projects[projectKey] = new Dictionary<string, StateEntry>(entries, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>True when another project in the same export location also wrote this file.</summary>
        public bool IsClaimedByOtherProject(string file, string projectKey)
        {
            return _projects.Any(p => !string.Equals(p.Key, projectKey, StringComparison.OrdinalIgnoreCase) && p.Value.ContainsKey(file));
        }

        /// <summary>Loads the state; a missing, unreadable or foreign file gives an empty state (everything is exported).</summary>
        public static ExportState Load(string path, TextWriter log)
        {
            var state = new ExportState();
            if (!File.Exists(path)) return state;
            try
            {
                string[] lines = File.ReadAllLines(path, Encoding.UTF8);
                if (lines.Length == 0 || lines[0] != Header)
                {
                    log.WriteLine("Incremental state " + path + " has an unknown format: exporting everything");
                    return state;
                }
                Dictionary<string, StateEntry> current = null;
                for (int i = 1; i < lines.Length; i++)
                {
                    string[] parts = lines[i].Split('\t');
                    if (parts.Length == 2 && parts[0] == "P")
                    {
                        current = new Dictionary<string, StateEntry>(StringComparer.OrdinalIgnoreCase);
                        state._projects[parts[1]] = current;
                    }
                    else if (parts.Length == 4 && parts[0] == "E" && current != null)
                    {
                        long stamp;
                        if (!long.TryParse(parts[1], out stamp)) throw new FormatException("bad stamp on line " + (i + 1));
                        current[parts[3]] = new StateEntry(stamp, parts[2].Length == 0 ? null : parts[2]);
                    }
                    else if (lines[i].Length != 0)
                    {
                        throw new FormatException("unexpected line " + (i + 1));
                    }
                }
            }
            catch (Exception exc)
            {
                log.WriteLine("Incremental state " + path + " could not be read (" + exc.Message + "): exporting everything");
                return new ExportState();
            }
            return state;
        }

        /// <summary>Writes the state via a temporary file so a crash never leaves a half-written state.</summary>
        public void Save(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            var lines = new List<string> { Header };
            foreach (var project in _projects.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase))
            {
                lines.Add("P\t" + project.Key);
                foreach (var entry in project.Value.OrderBy(e => e.Key, StringComparer.OrdinalIgnoreCase))
                {
                    lines.Add("E\t" + entry.Value.Stamp + "\t" + (entry.Value.Hash ?? "") + "\t" + entry.Key);
                }
            }
            string temp = path + ".tmp";
            File.WriteAllLines(temp, lines, new UTF8Encoding(false));
            FileOps.MoveReplacing(temp, path);
        }
    }
}
