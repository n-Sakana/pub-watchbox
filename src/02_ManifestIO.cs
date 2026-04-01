using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace WatchBox
{
    // Handles reading and writing manifest.csv in type-specific formats.
    // Mail format: entry_id,sender_email,sender_name,subject,received_at,folder_path,body_path,msg_path,attachment_paths,mail_folder,body_text
    // Folder format: item_id,file_name,file_path,folder_path,relative_path,file_size,modified_at
    public static class ManifestIO
    {
        const string MailHeader = "entry_id,sender_email,sender_name,subject,received_at,folder_path,body_path,msg_path,attachment_paths,mail_folder,body_text";
        const string FolderHeader = "item_id,file_name,file_path,folder_path,relative_path,file_size,modified_at";

        // Resolve manifest path for reading (check both names, prefer hidden)
        public static string ResolvePath(string outputRoot)
        {
            string hidden = Path.Combine(outputRoot, ".manifest.csv");
            if (File.Exists(hidden)) return hidden;
            string visible = Path.Combine(outputRoot, "manifest.csv");
            if (File.Exists(visible)) return visible;
            return hidden; // default for new files
        }

        // Get manifest path for writing (caller decides filename)
        public static string WritePath(string outputRoot, bool hide)
        {
            return Path.Combine(outputRoot, hide ? ".manifest.csv" : "manifest.csv");
        }

        // --- Load known IDs from manifest (for duplicate/change detection) ---

        public static HashSet<string> LoadIds(string outputRoot)
        {
            var ids = new HashSet<string>();
            string path = ResolvePath(outputRoot);
            if (!File.Exists(path)) return ids;
            foreach (var line in File.ReadAllLines(path, Encoding.UTF8))
            {
                if (string.IsNullOrEmpty(line)) continue;
                var cols = CsvSplit(line);
                if (cols.Length < 1) continue;
                string id = cols[0];
                if (id == "entry_id" || id == "item_id") continue;
                ids.Add(id);
            }
            return ids;
        }

        // --- Load full rows keyed by item ID ---

        public static Dictionary<string, string[]> LoadRows(string outputRoot)
        {
            var rows = new Dictionary<string, string[]>();
            string path = ResolvePath(outputRoot);
            if (!File.Exists(path)) return rows;
            foreach (var line in File.ReadAllLines(path, Encoding.UTF8))
            {
                if (string.IsNullOrEmpty(line)) continue;
                var cols = CsvSplit(line);
                if (cols.Length < 2) continue;
                if (cols[0] == "entry_id" || cols[0] == "item_id") continue;
                rows[cols[0]] = cols;
            }
            return rows;
        }

        // --- Load folder manifest rows keyed by item_id, returning size+mtime for change detection ---

        public static Dictionary<string, FolderManifestRow> LoadFolderRows(string outputRoot)
        {
            var rows = new Dictionary<string, FolderManifestRow>();
            string path = ResolvePath(outputRoot);
            if (!File.Exists(path)) return rows;
            foreach (var line in File.ReadAllLines(path, Encoding.UTF8))
            {
                if (string.IsNullOrEmpty(line)) continue;
                var cols = CsvSplit(line);
                if (cols.Length < 7 || cols[0] == "item_id") continue;
                rows[cols[0]] = new FolderManifestRow {
                    ItemId = cols[0],
                    FileName = cols[1],
                    FilePath = cols[2],
                    FolderPath = cols[3],
                    RelativePath = cols[4],
                    FileSize = cols[5],
                    ModifiedAt = cols[6]
                };
            }
            return rows;
        }

        // --- Append a mail row ---

        public static void AppendMailRow(string outputRoot, string entryId,
            string senderEmail, string senderName, string subject, DateTime receivedAt,
            string folderPath, string bodyPath, string msgPath, string attachmentPaths,
            string mailFolder, string bodyText, bool hide = true)
        {
            var csvPath = WritePath(outputRoot, hide);
            if (!File.Exists(csvPath))
                File.WriteAllText(csvPath, MailHeader + Environment.NewLine,
                    new UTF8Encoding(true));

            string line = string.Join(",", new[] {
                entryId,
                senderEmail,
                senderName,
                CsvQuote(subject),
                receivedAt.ToString("yyyy-MM-dd\\THH:mm:ss"),
                folderPath,
                bodyPath,
                msgPath,
                attachmentPaths,
                mailFolder,
                CsvQuote(bodyText)
            });
            File.AppendAllText(csvPath, line + Environment.NewLine, new UTF8Encoding(true));
        }

        // --- Append a folder row ---

        public static void AppendFolderRow(string outputRoot, string itemId,
            string fileName, string filePath, string folderPath, string relativePath,
            long fileSize, DateTime modifiedAt, bool hide = true)
        {
            var csvPath = WritePath(outputRoot, hide);
            if (!File.Exists(csvPath))
                File.WriteAllText(csvPath, FolderHeader + Environment.NewLine,
                    new UTF8Encoding(true));

            string line = string.Join(",", new[] {
                itemId,
                CsvQuote(fileName),
                filePath,
                folderPath,
                relativePath,
                fileSize.ToString(),
                modifiedAt.ToString("yyyy-MM-dd\\THH:mm:ss")
            });
            File.AppendAllText(csvPath, line + Environment.NewLine, new UTF8Encoding(true));
        }

        // --- Rewrite manifest removing specific IDs ---

        public static void RemoveRows(string outputRoot, HashSet<string> removeIds)
        {
            var csvPath = ResolvePath(outputRoot);
            if (!File.Exists(csvPath) || removeIds.Count == 0) return;

            var lines = File.ReadAllLines(csvPath, Encoding.UTF8);
            var kept = new List<string>();
            foreach (var line in lines)
            {
                if (string.IsNullOrEmpty(line)) continue;
                int sep = line.IndexOf(',');
                string id = sep > 0 ? line.Substring(0, sep) : line;
                if (id == "entry_id" || id == "item_id" || !removeIds.Contains(id))
                    kept.Add(line);
            }
            File.WriteAllLines(csvPath, kept.ToArray(), new UTF8Encoding(true));
        }

        // --- Rewrite entire folder manifest from scratch ---

        public static void WriteFolderManifest(string outputRoot, List<FolderManifestRow> rows, bool hide = true)
        {
            var csvPath = WritePath(outputRoot, hide);
            var lines = new List<string> { FolderHeader };
            foreach (var r in rows)
            {
                lines.Add(string.Join(",", new[] {
                    r.ItemId,
                    CsvQuote(r.FileName),
                    r.FilePath,
                    r.FolderPath,
                    r.RelativePath,
                    r.FileSize,
                    r.ModifiedAt
                }));
            }
            File.WriteAllLines(csvPath, lines.ToArray(), new UTF8Encoding(true));
        }

        // --- Search manifest (used by Viewer) ---

        public static List<string[]> SearchManifest(string outputRoot, string query)
        {
            var results = new List<string[]>();
            if (string.IsNullOrEmpty(outputRoot)) return results;
            string path = ResolvePath(outputRoot);
            if (!File.Exists(path)) return results;

            string q = (query ?? "").Trim().ToLower();
            if (q.Length == 0) return results;

            foreach (var line in File.ReadAllLines(path, Encoding.UTF8))
            {
                if (string.IsNullOrEmpty(line)) continue;
                if (line.ToLower().Contains(q))
                {
                    var cols = CsvSplit(line);
                    if (cols[0] == "entry_id" || cols[0] == "item_id") continue;
                    results.Add(cols);
                }
            }
            return results;
        }

        // --- Get latest received_at from mail manifest (for incremental scan) ---

        public static DateTime GetLatestMailDate(string outputRoot)
        {
            DateTime latest = DateTime.MinValue;
            string path = ResolvePath(outputRoot);
            if (!File.Exists(path)) return latest;
            foreach (var line in File.ReadAllLines(path, Encoding.UTF8))
            {
                if (string.IsNullOrEmpty(line)) continue;
                var cols = CsvSplit(line);
                if (cols.Length < 5 || cols[0] == "entry_id") continue;
                DateTime dt;
                if (DateTime.TryParse(cols[4], out dt) && dt > latest)
                    latest = dt;
            }
            return latest;
        }

        // --- Detect manifest type from header ---

        public static string DetectType(string outputRoot)
        {
            string path = ResolvePath(outputRoot);
            if (!File.Exists(path)) return "";
            using (var sr = new StreamReader(path, Encoding.UTF8))
            {
                string header = sr.ReadLine();
                if (header != null && header.StartsWith("item_id")) return "folder";
                if (header != null && header.StartsWith("entry_id")) return "mail";
            }
            return "";
        }

        // --- Hash helper for folder item IDs ---

        public static string ComputeItemId(string relativePath)
        {
            string normalized = relativePath.ToLower().Replace('\\', '/');
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(normalized));
                var sb = new StringBuilder(16);
                for (int i = 0; i < 8; i++)
                    sb.Append(hash[i].ToString("x2"));
                return sb.ToString();
            }
        }

        // Quote a CSV field if it contains comma, quote, or newline (RFC 4180)
        static string CsvQuote(string value)
        {
            if (value == null) return "";
            if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0)
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            return value;
        }

        // Split a CSV line respecting quoted fields (RFC 4180)
        public static string[] CsvSplit(string line)
        {
            var fields = new List<string>();
            int i = 0;
            while (i <= line.Length)
            {
                if (i == line.Length) { fields.Add(""); break; }
                if (line[i] == '"')
                {
                    i++;
                    var sb = new StringBuilder();
                    while (i < line.Length)
                    {
                        if (line[i] == '"')
                        {
                            if (i + 1 < line.Length && line[i + 1] == '"')
                                { sb.Append('"'); i += 2; }
                            else { i++; break; }
                        }
                        else { sb.Append(line[i]); i++; }
                    }
                    fields.Add(sb.ToString());
                    if (i < line.Length && line[i] == ',') i++;
                }
                else
                {
                    int sep = line.IndexOf(',', i);
                    if (sep < 0) { fields.Add(line.Substring(i)); break; }
                    fields.Add(line.Substring(i, sep - i));
                    i = sep + 1;
                }
            }
            return fields.ToArray();
        }
    }

    public class FolderManifestRow
    {
        public string ItemId;
        public string FileName;
        public string FilePath;
        public string FolderPath;
        public string RelativePath;
        public string FileSize;
        public string ModifiedAt;
    }
}
