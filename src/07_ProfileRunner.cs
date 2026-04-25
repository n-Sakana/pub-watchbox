using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace WatchBox
{
    public static class ProfileRunner
    {
        // Run a single profile: scan, write manifest, log changes.
        // If scanner is provided, caller controls the instance (for cancel support).
        public static RunResult Run(int profileIndex, SourceScanner scanner = null,
            Action<int, string> onProgress = null)
        {
            string type = Config.PGet(profileIndex, "type", "mail");
            string outputRoot = (Config.PGet(profileIndex, "output_root") ?? "").Trim();
            if (string.IsNullOrEmpty(outputRoot)) return new RunResult();

            bool logEnabled = Config.PGet(profileIndex, "log_enabled", "1") == "1";
            bool hideManifest = Config.PGet(profileIndex, "manifest_hidden", "1") == "1";
            var config = BuildConfig(profileIndex);
            bool hasSource = !string.IsNullOrEmpty(
                type == "mail" ? config["account"] : config["source_folder"]);

            bool ownScanner = scanner == null;
            if (ownScanner)
            {
                switch (type)
                {
                    case "folder": scanner = new FolderScanner(); break;
                    default: scanner = new MailScanner(); break;
                }
            }
            if (onProgress != null) scanner.ProgressChanged += onProgress;

            List<ScanResult> newItems;
            List<string> removedIds;
            var knownIds = ManifestIO.LoadIds(outputRoot);
            try
            {
                newItems = scanner.Scan(config, knownIds);
                removedIds = scanner.DetectRemoved(config, knownIds);
            }
            finally
            {
                if (onProgress != null) scanner.ProgressChanged -= onProgress;
                // Release COM objects when we created the scanner
                if (ownScanner && scanner is MailScanner)
                    ((MailScanner)scanner).Cleanup();
            }

            // Process new/changed items
            int added = 0, modified = 0;
            foreach (var item in newItems)
            {
                bool wasKnown = knownIds.Contains(item.ItemId);
                WriteManifestRow(type, outputRoot, item, config, hideManifest);
                if (logEnabled)
                    ChangeLog.Append(outputRoot, wasKnown ? "modified" : "added",
                        item.ItemId, item.Name);
                if (wasKnown) modified++; else added++;
            }

            // Process removals (mirror: delete from output too)
            if (removedIds.Count > 0)
            {
                if (hasSource && type != "mail")
                    DeleteMirroredFiles(outputRoot, removedIds);
                if (logEnabled)
                    LogRemovals(outputRoot, removedIds);
                ManifestIO.RemoveRows(outputRoot, new HashSet<string>(removedIds));
            }

            // Rewrite folder manifest every run to ensure filter/flat_output changes
            // are reflected even when no files were modified
            if (type != "mail")
                RewriteFolderManifest(outputRoot, config, type, hideManifest);

            // Record successful scan timestamp for incremental date filter
            if (type == "mail" && !scanner.CancelRequested)
                Config.PSet(profileIndex, "last_scan",
                    DateTime.UtcNow.ToString("yyyy-MM-dd"));

            return new RunResult { Added = added, Modified = modified, Removed = removedIds.Count };
        }

        // Run a mail profile using pre-cached folder data (for grouped scan).
        public static RunResult RunFromCache(int profileIndex, MailScanner scanner,
            List<CachedMailItem> cache, Action<int, string> onProgress = null)
        {
            string outputRoot = (Config.PGet(profileIndex, "output_root") ?? "").Trim();
            if (string.IsNullOrEmpty(outputRoot)) return new RunResult();

            bool logEnabled = Config.PGet(profileIndex, "log_enabled", "1") == "1";
            bool hideManifest = Config.PGet(profileIndex, "manifest_hidden", "1") == "1";
            var config = BuildConfig(profileIndex);

            if (onProgress != null) scanner.ProgressChanged += onProgress;

            var knownIds = ManifestIO.LoadIds(outputRoot);
            var newItems = scanner.ExportFromCache(cache, config, knownIds);

            int added = 0, modified = 0;
            foreach (var item in newItems)
            {
                bool wasKnown = knownIds.Contains(item.ItemId);
                WriteManifestRow("mail", outputRoot, item, config, hideManifest);
                if (logEnabled)
                    ChangeLog.Append(outputRoot, wasKnown ? "modified" : "added",
                        item.ItemId, item.Name);
                if (wasKnown) modified++; else added++;
            }

            if (!scanner.CancelRequested)
                Config.PSet(profileIndex, "last_scan",
                    DateTime.UtcNow.ToString("yyyy-MM-dd"));

            if (onProgress != null) scanner.ProgressChanged -= onProgress;
            return new RunResult { Added = added, Modified = modified, Removed = 0 };
        }

        static void WriteManifestRow(string type, string outputRoot,
            ScanResult item, Dictionary<string, string> config, bool hide)
        {
            if (type == "mail")
            {
                ManifestIO.AppendMailRow(outputRoot,
                    item.ItemId, item.SenderEmail, item.SenderName,
                    item.Subject, item.ReceivedAt, item.SourcePath,
                    item.BodyPath, item.MsgPath, item.AttachmentPaths,
                    item.ItemFolder, item.BodyText,
                    item.ToRecipients, item.CcRecipients, hide);
            }
            else
            {
                ManifestIO.AppendFolderRow(outputRoot,
                    item.ItemId, item.Name, item.SourcePath,
                    item.ItemFolder, GetRelativePath(outputRoot, item, config),
                    item.Size, item.ModifiedAt, hide);
            }
        }

        static void DeleteMirroredFiles(string outputRoot, List<string> removedIds)
        {
            var rows = ManifestIO.LoadFolderRows(outputRoot);
            foreach (var id in removedIds)
            {
                FolderManifestRow row;
                if (!rows.TryGetValue(id, out row)) continue;
                string destPath = Path.Combine(outputRoot,
                    row.RelativePath.Replace('/', '\\'));
                try
                {
                    if (File.Exists(destPath)) File.Delete(destPath);
                }
                catch { }
            }
        }

        static void LogRemovals(string outputRoot, List<string> removedIds)
        {
            var existingRows = ManifestIO.LoadRows(outputRoot);
            foreach (var id in removedIds)
            {
                string name = "";
                string[] row;
                if (existingRows.TryGetValue(id, out row) && row.Length > 1)
                    name = row[1];
                ChangeLog.Append(outputRoot, "removed", id, name);
            }
        }

        static Dictionary<string, string> BuildConfig(int profileIndex)
        {
            var config = new Dictionary<string, string>();
            string[] keys = { "output_root", "account", "outlook_folder", "since",
                "filter_mode", "filters", "flat_output", "source_folder",
                "recurse", "type", "short_dirname", "auto_unzip", "last_scan" };
            foreach (var k in keys)
                config[k] = Config.PGet(profileIndex, k);
            return config;
        }

        static string GetRelativePath(string outputRoot, ScanResult item,
            Dictionary<string, string> config)
        {
            bool flatOutput = config.ContainsKey("flat_output") && config["flat_output"] == "1";
            if (flatOutput) return item.Name;

            string source = config.ContainsKey("source_folder") ? config["source_folder"] : "";
            if (!string.IsNullOrEmpty(source) && item.SourcePath.StartsWith(source))
            {
                return item.SourcePath.Substring(source.Length).TrimStart('\\', '/');
            }
            else if (item.SourcePath.StartsWith(outputRoot))
            {
                return item.SourcePath.Substring(outputRoot.Length).TrimStart('\\', '/');
            }
            return item.Name;
        }

        // Rewrite folder manifest: re-scan and write fresh
        // Applies the same filters and flat_output logic as FolderScanner.Scan()
        static void RewriteFolderManifest(string outputRoot, Dictionary<string, string> config,
            string type, bool hide = true)
        {
            bool recurse = !config.ContainsKey("recurse") || config["recurse"] != "0";
            bool flatOutput = config.ContainsKey("flat_output") && config["flat_output"] == "1";

            // Parse filter settings (same logic as FolderScanner.ParseFilters)
            DateTime sinceDate = DateTime.MinValue;
            string since = config.ContainsKey("since") ? config["since"] : "";
            if (!string.IsNullOrEmpty(since))
            {
                DateTime dt;
                string[] fmts = { "yyyy-MM-dd", "yyyy/MM/dd", "yyyy/M/d", "M/d/yyyy" };
                if (DateTime.TryParseExact(since, fmts, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out dt))
                    sinceDate = dt;
                else
                    System.Diagnostics.Debug.WriteLine(
                        "RewriteFolderManifest: failed to parse since date: " + since);
            }
            string filterMode = config.ContainsKey("filter_mode") && config["filter_mode"] == "and"
                ? "and" : "or";
            var filterWords = new List<string>();
            string keywords = config.ContainsKey("filters") ? config["filters"] : "";
            if (!string.IsNullOrEmpty(keywords))
                foreach (var kw in keywords.Split(';'))
                    if (kw.Trim().Length > 0) filterWords.Add(kw.Trim().ToLower());

            string source = config.ContainsKey("source_folder") ? config["source_folder"] : "";
            string scanRoot = !string.IsNullOrEmpty(source) ? source : outputRoot;
            if (!Directory.Exists(scanRoot)) return;

            var option = recurse ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            string[] files;
            try { files = Directory.GetFiles(scanRoot, "*", option); }
            catch { return; }

            var rows = new List<FolderManifestRow>();
            foreach (var filePath in files)
            {
                string fn = Path.GetFileName(filePath);
                if (fn == ".manifest.csv" || fn == "manifest.csv" || fn == "log.csv") continue;
                try
                {
                    var fi = new FileInfo(filePath);

                    // Apply date filter
                    if (sinceDate > DateTime.MinValue && fi.LastWriteTime < sinceDate) continue;

                    // Apply keyword filter on filename
                    if (filterWords.Count > 0)
                    {
                        string nameLower = fi.Name.ToLower();
                        if (filterMode == "and")
                        {
                            bool allMatch = true;
                            foreach (var kw in filterWords)
                                if (!nameLower.Contains(kw)) { allMatch = false; break; }
                            if (!allMatch) continue;
                        }
                        else
                        {
                            bool anyMatch = false;
                            foreach (var kw in filterWords)
                                if (nameLower.Contains(kw)) { anyMatch = true; break; }
                            if (!anyMatch) continue;
                        }
                    }

                    // Compute relative path respecting flat_output
                    string relativePath = flatOutput
                        ? fi.Name
                        : filePath.Substring(scanRoot.Length).TrimStart('\\', '/');
                    rows.Add(new FolderManifestRow {
                        ItemId = ManifestIO.ComputeItemId(relativePath),
                        FileName = fi.Name,
                        FilePath = filePath,
                        FolderPath = Path.GetDirectoryName(filePath),
                        RelativePath = relativePath,
                        FileSize = fi.Length.ToString(),
                        ModifiedAt = fi.LastWriteTime.ToString("yyyy-MM-dd\\THH:mm:ss")
                    });
                }
                catch { }
            }
            ManifestIO.WriteFolderManifest(outputRoot, rows, hide);
        }
    }
}
