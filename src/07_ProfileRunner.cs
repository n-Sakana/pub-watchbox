using System;
using System.Collections.Generic;
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
            string outputRoot = Config.PGet(profileIndex, "output_root");
            if (string.IsNullOrEmpty(outputRoot)) return new RunResult();

            bool logEnabled = Config.PGet(profileIndex, "log_enabled", "1") == "1";
            bool hideManifest = Config.PGet(profileIndex, "manifest_hidden", "1") == "1";
            var config = BuildConfig(profileIndex);
            bool hasSource = !string.IsNullOrEmpty(
                type == "mail" ? config["account"] : config["source_folder"]);

            if (scanner == null)
            {
                switch (type)
                {
                    case "folder": scanner = new FolderScanner(); break;
                    default: scanner = new MailScanner(); break;
                }
            }
            if (onProgress != null) scanner.ProgressChanged += onProgress;

            var knownIds = ManifestIO.LoadIds(outputRoot);
            var newItems = scanner.Scan(config, knownIds);
            var removedIds = scanner.DetectRemoved(config, knownIds);

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

            // Rewrite folder manifest when items were modified (updated rows need replacing)
            if (modified > 0 && type != "mail")
                RewriteFolderManifest(outputRoot, config, type, hideManifest);

            return new RunResult { Added = added, Modified = modified, Removed = removedIds.Count };
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
                    item.ItemFolder, item.BodyText, hide);
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
                "recurse", "type" };
            foreach (var k in keys)
                config[k] = Config.PGet(profileIndex, k);
            return config;
        }

        static string GetRelativePath(string outputRoot, ScanResult item,
            Dictionary<string, string> config)
        {
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
        static void RewriteFolderManifest(string outputRoot, Dictionary<string, string> config,
            string type, bool hide = true)
        {
            bool recurse = !config.ContainsKey("recurse") || config["recurse"] != "0";

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
                    string relativePath = filePath.Substring(scanRoot.Length).TrimStart('\\', '/');
                    var fi = new FileInfo(filePath);
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
