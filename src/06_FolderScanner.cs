using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;

namespace WatchBox
{
    // Unified folder scanner: copies files when source_folder is set,
    // otherwise scans output_root in-place (manifest only).
    // Supports since (date filter), keywords (filename match), and flat_output.
    public class FolderScanner : SourceScanner
    {
        DateTime _sinceDate;
        string _filterMode;
        List<string> _filterWords;
        bool _flatOutput;
        bool _autoUnzip;

        void ParseFilters(Dictionary<string, string> config)
        {
            _sinceDate = DateTime.MinValue;
            string since = config.ContainsKey("since") ? config["since"] : "";
            if (!string.IsNullOrEmpty(since))
            {
                DateTime dt;
                string[] fmts = { "yyyy-MM-dd", "yyyy/MM/dd", "yyyy/M/d", "M/d/yyyy" };
                if (DateTime.TryParseExact(since, fmts, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out dt))
                    _sinceDate = dt;
                else
                    System.Diagnostics.Debug.WriteLine(
                        "FolderScanner: failed to parse since date: " + since);
            }

            _filterMode = "or";
            if (config.ContainsKey("filter_mode") && config["filter_mode"] == "and")
                _filterMode = "and";

            _filterWords = new List<string>();
            string keywords = config.ContainsKey("filters") ? config["filters"] : "";
            if (!string.IsNullOrEmpty(keywords))
                foreach (var kw in keywords.Split(';'))
                    if (kw.Trim().Length > 0) _filterWords.Add(kw.Trim().ToLower());

            _flatOutput = config.ContainsKey("flat_output") && config["flat_output"] == "1";
            _autoUnzip = config.ContainsKey("auto_unzip") && config["auto_unzip"] == "1";
        }

        bool PassesFilter(FileInfo fi)
        {
            // Date filter
            if (_sinceDate > DateTime.MinValue && fi.LastWriteTime < _sinceDate)
                return false;

            // Keyword filter on filename
            if (_filterWords.Count > 0)
            {
                string name = fi.Name.ToLower();
                if (_filterMode == "and")
                {
                    foreach (var kw in _filterWords)
                        if (!name.Contains(kw)) return false;
                }
                else
                {
                    bool any = false;
                    foreach (var kw in _filterWords)
                        if (name.Contains(kw)) { any = true; break; }
                    if (!any) return false;
                }
            }
            return true;
        }

        public override List<ScanResult> Scan(
            Dictionary<string, string> config, HashSet<string> knownIds)
        {
            ParseFilters(config);
            string sourceFolder = config.ContainsKey("source_folder") ? config["source_folder"] : "";
            bool hasSource = !string.IsNullOrEmpty(sourceFolder) && Directory.Exists(sourceFolder);
            return hasSource ? ScanWithCopy(config, sourceFolder) : ScanInPlace(config);
        }

        List<ScanResult> ScanWithCopy(Dictionary<string, string> config, string sourceFolder)
        {
            var results = new List<ScanResult>();
            string outputRoot = config.ContainsKey("output_root") ? config["output_root"].Trim() : "";
            bool recurse = !config.ContainsKey("recurse") || config["recurse"] != "0";

            if (string.IsNullOrEmpty(outputRoot)) return results;
            Directory.CreateDirectory(outputRoot);

            var existing = ManifestIO.LoadFolderRows(outputRoot);
            var files = EnumFiles(sourceFolder, recurse);

            int count = 0;
            foreach (var filePath in files)
            {
                if (CancelRequested) break;
                try
                {
                    var fi = new FileInfo(filePath);
                    if (!PassesFilter(fi)) continue;

                    string relativePath = filePath.Substring(sourceFolder.Length).TrimStart('\\', '/');
                    string itemId = ManifestIO.ComputeItemId(relativePath);

                    // Compute destination path for re-copy detection
                    string destPath;
                    if (_flatOutput)
                        destPath = Path.Combine(outputRoot, fi.Name);
                    else
                        destPath = Path.Combine(outputRoot, relativePath);

                    if (!IsNewOrModified(itemId, fi, existing, destPath)) continue;
                    string destDir = Path.GetDirectoryName(destPath);
                    Directory.CreateDirectory(destDir);
                    File.Copy(filePath, destPath, true);

                    // Auto-extract zip files: extract then remove the zip
                    if (_autoUnzip && fi.Extension.ToLower() == ".zip")
                    {
                        if (TryExtractZip(destPath, destDir))
                        {
                            try { File.Delete(destPath); } catch { }
                            continue; // skip adding zip to results
                        }
                    }

                    count++;
                    results.Add(new ScanResult {
                        ItemId = itemId,
                        Name = fi.Name,
                        SourcePath = filePath,
                        Size = fi.Length,
                        ModifiedAt = fi.LastWriteTime,
                        CreatedAt = fi.CreationTime,
                        ItemFolder = destDir
                    });
                    OnProgress(count, fi.Name);
                }
                catch { }
            }
            return results;
        }

        List<ScanResult> ScanInPlace(Dictionary<string, string> config)
        {
            var results = new List<ScanResult>();
            string outputRoot = config.ContainsKey("output_root") ? config["output_root"].Trim() : "";
            bool recurse = !config.ContainsKey("recurse") || config["recurse"] != "0";

            if (string.IsNullOrEmpty(outputRoot) || !Directory.Exists(outputRoot)) return results;

            var existing = ManifestIO.LoadFolderRows(outputRoot);
            var files = EnumFiles(outputRoot, recurse);

            int count = 0;
            foreach (var filePath in files)
            {
                if (CancelRequested) break;
                string fn = Path.GetFileName(filePath);
                if (fn == ".manifest.csv" || fn == "manifest.csv" || fn == "log.csv") continue;

                try
                {
                    var fi = new FileInfo(filePath);
                    if (!PassesFilter(fi)) continue;

                    string relativePath = filePath.Substring(outputRoot.Length).TrimStart('\\', '/');
                    string itemId = ManifestIO.ComputeItemId(relativePath);

                    if (!IsNewOrModified(itemId, fi, existing)) continue;

                    count++;
                    results.Add(new ScanResult {
                        ItemId = itemId,
                        Name = fi.Name,
                        SourcePath = filePath,
                        Size = fi.Length,
                        ModifiedAt = fi.LastWriteTime,
                        CreatedAt = fi.CreationTime,
                        ItemFolder = Path.GetDirectoryName(filePath)
                    });
                    OnProgress(count, fi.Name);
                }
                catch { }
            }
            return results;
        }

        // --- Zip extraction ---

        static string LongPath(string path)
        {
            if (path.StartsWith(@"\\?\")) return path;
            return @"\\?\" + Path.GetFullPath(path);
        }

        static bool TryExtractZip(string zipPath, string destDir)
        {
            string extractDir = Path.Combine(destDir,
                Path.GetFileNameWithoutExtension(zipPath));
            try
            {
                using (var archive = ZipFile.OpenRead(zipPath))
                {
                    // Check first entry for encryption (password protection)
                    foreach (var entry in archive.Entries)
                    {
                        if (entry.Length == 0) continue;
                        try
                        {
                            using (var stream = entry.Open())
                                stream.ReadByte();
                        }
                        catch (InvalidDataException)
                        {
                            return false; // password-protected
                        }
                        break;
                    }

                    foreach (var entry in archive.Entries)
                    {
                        if (string.IsNullOrEmpty(entry.Name)) continue;
                        string entryDest = Path.Combine(extractDir, entry.FullName);
                        // Use long path prefix to handle paths > 260 chars
                        string entryDestLong = LongPath(entryDest);
                        string entryDir = Path.GetDirectoryName(entryDestLong);
                        Directory.CreateDirectory(entryDir);
                        try
                        {
                            using (var src = entry.Open())
                            using (var dst = new FileStream(entryDestLong,
                                FileMode.Create, FileAccess.Write, FileShare.None))
                            {
                                src.CopyTo(dst);
                            }
                        }
                        catch { }
                    }
                }
                return true;
            }
            catch { return false; }
        }

        // --- Common helpers ---

        static string[] EnumFiles(string root, bool recurse)
        {
            var option = recurse ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            try { return Directory.GetFiles(root, "*", option); }
            catch { return new string[0]; }
        }

        static bool IsNewOrModified(string itemId, FileInfo fi,
            Dictionary<string, FolderManifestRow> existing, string destPath = null)
        {
            FolderManifestRow row;
            if (!existing.TryGetValue(itemId, out row)) return true;
            // Re-copy if destination file was manually deleted
            if (destPath != null && !File.Exists(destPath)) return true;
            return row.FileSize != fi.Length.ToString() ||
                   row.ModifiedAt != fi.LastWriteTime.ToString("yyyy-MM-dd\\THH:mm:ss");
        }

        // --- Removal detection ---

        public override List<string> DetectRemoved(
            Dictionary<string, string> config, HashSet<string> knownIds)
        {
            ParseFilters(config);
            string sourceFolder = config.ContainsKey("source_folder") ? config["source_folder"] : "";
            bool hasSource = !string.IsNullOrEmpty(sourceFolder) && Directory.Exists(sourceFolder);
            return hasSource ? DetectRemovedFromSource(config, sourceFolder)
                             : DetectRemovedInPlace(config);
        }

        List<string> DetectRemovedFromSource(Dictionary<string, string> config, string sourceFolder)
        {
            var removed = new List<string>();
            string outputRoot = config.ContainsKey("output_root") ? config["output_root"] : "";
            var existing = ManifestIO.LoadFolderRows(outputRoot);
            foreach (var row in existing.Values)
            {
                string sourcePath = Path.Combine(sourceFolder, row.RelativePath.Replace('/', '\\'));
                if (!File.Exists(sourcePath))
                {
                    removed.Add(row.ItemId);
                    continue;
                }
                // Also remove entries that no longer pass filters
                try
                {
                    var fi = new FileInfo(sourcePath);
                    if (!PassesFilter(fi))
                        removed.Add(row.ItemId);
                }
                catch { }
            }
            return removed;
        }

        List<string> DetectRemovedInPlace(Dictionary<string, string> config)
        {
            var removed = new List<string>();
            string outputRoot = config.ContainsKey("output_root") ? config["output_root"] : "";
            var existing = ManifestIO.LoadFolderRows(outputRoot);
            foreach (var row in existing.Values)
            {
                if (!File.Exists(row.FilePath))
                {
                    removed.Add(row.ItemId);
                    continue;
                }
                // Also remove entries that no longer pass filters
                try
                {
                    var fi = new FileInfo(row.FilePath);
                    if (!PassesFilter(fi))
                        removed.Add(row.ItemId);
                }
                catch { }
            }
            return removed;
        }
    }
}
