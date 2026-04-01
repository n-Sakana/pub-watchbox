using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace WatchBox
{
    public class MailScanner : SourceScanner
    {
        const int OL_MSG_UNICODE = 9;
        const int OL_MAIL_CLASS = 43;

        dynamic _olApp;
        dynamic _olNs;
        HashSet<string> _exported;
        int _itemCount;
        string _filterMode;
        List<string> _filterWords;
        bool _flatOutput;
        bool _shortDirname;

        // --- Outlook connection ---

        public bool Connect()
        {
            if (_olApp != null) return true;
            try
            {
                try { _olApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application"); }
                catch { _olApp = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")); }
                _olNs = _olApp.GetNamespace("MAPI");
                return true;
            }
            catch { return false; }
        }

        // --- Accounts / Folders (used by SettingsForm) ---

        public List<string> GetAccounts()
        {
            var list = new List<string>();
            if (!Connect()) return list;
            try
            {
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (dynamic acct in _olNs.Accounts)
                {
                    try
                    {
                        string smtp = (string)acct.SmtpAddress;
                        if (!string.IsNullOrEmpty(smtp) && seen.Add(smtp))
                            list.Add(smtp);
                    }
                    catch { }
                }
                // Also enumerate stores not tied to any account (shared/delegate mailboxes)
                var accountStoreIds = new HashSet<string>();
                foreach (dynamic acct in _olNs.Accounts)
                {
                    try { accountStoreIds.Add((string)acct.DeliveryStore.StoreID); } catch { }
                }
                foreach (dynamic store in _olNs.Stores)
                {
                    try
                    {
                        if (accountStoreIds.Contains((string)store.StoreID)) continue;
                        string addr = "";
                        // Try PR_EMAIL_ADDRESS
                        try { addr = ((string)store.GetRootFolder().PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E")).ToLower(); } catch { }
                        // Try PR_SMTP_ADDRESS (alternative property for shared mailboxes)
                        if (string.IsNullOrEmpty(addr))
                            try { addr = ((string)store.GetRootFolder().PropertyAccessor.GetProperty(
                                "http://schemas.microsoft.com/mapi/proptag/0x39FE001F")).ToLower(); } catch { }
                        // Fallback to display name
                        if (string.IsNullOrEmpty(addr))
                            try { addr = ((string)store.DisplayName).ToLower(); } catch { }
                        if (!string.IsNullOrEmpty(addr) && seen.Add(addr))
                            list.Add(addr);
                    }
                    catch { }
                }
            }
            catch { }
            return list;
        }

        public List<string[]> GetFolders(string accountFilter)
        {
            var list = new List<string[]>();
            if (!Connect()) return list;
            try
            {
                if (string.IsNullOrEmpty(accountFilter))
                {
                    foreach (dynamic store in _olNs.Stores)
                    {
                        try
                        {
                            string smtp = GetStoreSmtp(store);
                            if (!string.IsNullOrEmpty(smtp))
                                CollectFolders(store.GetRootFolder(), 0, smtp + ": ", list);
                        }
                        catch { }
                    }
                }
                else
                {
                    // Try matching account first, then fall back to store lookup
                    bool found = false;
                    foreach (dynamic acct in _olNs.Accounts)
                    {
                        if (string.Equals((string)acct.SmtpAddress, accountFilter,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            CollectFolders(acct.DeliveryStore.GetRootFolder(), 0, "", list);
                            found = true;
                            break;
                        }
                    }
                    // Shared/delegate mailbox: match by store SMTP or display name
                    if (!found)
                    {
                        foreach (dynamic store in _olNs.Stores)
                        {
                            try
                            {
                                string smtp = GetStoreSmtp(store);
                                if (string.Equals(smtp, accountFilter, StringComparison.OrdinalIgnoreCase))
                                {
                                    CollectFolders(store.GetRootFolder(), 0, "", list);
                                    break;
                                }
                            }
                            catch { }
                        }
                    }
                }
            }
            catch { }
            return list;
        }

        void CollectFolders(dynamic folder, int depth, string prefix, List<string[]> list)
        {
            try
            {
                foreach (dynamic child in folder.Folders)
                {
                    try
                    {
                        // Only include mail folders (DefaultItemType == olMailItem == 0)
                        if ((int)child.DefaultItemType != 0) continue;
                        // Skip hidden system folders (PR_ATTR_HIDDEN)
                        try { if ((bool)child.PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x10F4000B")) continue; }
                        catch { }
                        string indent = new string(' ', depth * 2);
                        list.Add(new[] { prefix + indent + (string)child.Name, (string)child.FolderPath });
                        CollectFolders(child, depth + 1, prefix, list);
                    }
                    catch { }
                }
            }
            catch { }
        }

        // --- SourceScanner implementation ---

        public override List<ScanResult> Scan(
            Dictionary<string, string> config, HashSet<string> knownIds)
        {
            var results = new List<ScanResult>();
            string outputRoot = config.ContainsKey("output_root") ? config["output_root"].Trim() : "";
            if (!Connect() || string.IsNullOrEmpty(outputRoot)) return results;
            Directory.CreateDirectory(outputRoot);

            string filterAccount = config.ContainsKey("account") ? config["account"] : "";
            string filterFolder = config.ContainsKey("outlook_folder") ? config["outlook_folder"] : "";
            string sinceDate = config.ContainsKey("since") ? config["since"] : "";
            string filterMode = config.ContainsKey("filter_mode") ? config["filter_mode"] : "";
            string filterKeywords = config.ContainsKey("filters") ? config["filters"] : "";
            _flatOutput = config.ContainsKey("flat_output") && config["flat_output"] == "1";
            _shortDirname = config.ContainsKey("short_dirname") && config["short_dirname"] == "1";

            _filterMode = (filterMode ?? "").ToLower() == "and" ? "and" : "or";
            _filterWords = new List<string>();
            if (!string.IsNullOrEmpty(filterKeywords))
                foreach (var kw in filterKeywords.Split(';'))
                    if (kw.Trim().Length > 0) _filterWords.Add(kw.Trim().ToLower());

            // Remove IDs whose output folders were manually deleted
            _exported = new HashSet<string>(knownIds);
            PurgeStaleMailIds(outputRoot, _exported);

            // Build two separate filters:
            // 1. Date filter (Jet syntax, local time) — accurate timezone handling
            // 2. Keyword filter (DASL syntax) — server-side text search
            // Cannot mix Jet and DASL in one Restrict(), so chain two calls.
            string dateFilter = BuildDateFilter(sinceDate, config);
            string keywordFilter = BuildKeywordFilter();

            // When account is empty, scan the first account only to avoid
            // exporting every mailbox into a single profile's output_root
            if (string.IsNullOrEmpty(filterAccount))
            {
                try
                {
                    foreach (dynamic acct in _olNs.Accounts)
                    {
                        try { filterAccount = ((string)acct.SmtpAddress).ToLower(); break; }
                        catch { }
                    }
                }
                catch { }
                if (string.IsNullOrEmpty(filterAccount)) return results;
            }

            foreach (dynamic store in _olNs.Stores)
            {
                if (CancelRequested) break;
                try
                {
                    string smtp = GetStoreSmtp(store);
                    if (string.IsNullOrEmpty(smtp)) continue;
                    if (!string.Equals(smtp, filterAccount, StringComparison.OrdinalIgnoreCase)) continue;

                    if (!string.IsNullOrEmpty(filterFolder))
                    {
                        dynamic startFolder = FindFolder(store.GetRootFolder(), filterFolder);
                        if (startFolder != null)
                            ScanTree(startFolder, outputRoot, smtp, dateFilter, keywordFilter, results, false);
                    }
                    else
                        ScanTree(store.GetRootFolder(), outputRoot, smtp, dateFilter, keywordFilter, results, true);
                }
                catch { }
            }
            return results;
        }

        public override List<string> DetectRemoved(
            Dictionary<string, string> config, HashSet<string> knownIds)
        {
            // Mail items are append-only; removal detection is expensive and skipped for now
            return new List<string>();
        }

        // --- Internal tree scan ---

        void ScanTree(dynamic folder, string outputRoot, string smtp,
            string dateFilter, string keywordFilter,
            List<ScanResult> results, bool recurse = true)
        {
            try
            {
                string folderRoot;
                if (_flatOutput)
                    folderRoot = outputRoot;
                else
                    folderRoot = Path.Combine(outputRoot,
                        SafeName(smtp) + NormalizeFolderPath((string)folder.FolderPath));
                Directory.CreateDirectory(folderRoot);

                dynamic items = folder.Items;
                // Chain two Restrict() calls: date (Jet, local time) then keywords (DASL)
                if (dateFilter != null) items = items.Restrict(dateFilter);
                if (keywordFilter != null) items = items.Restrict(keywordFilter);

                dynamic item = items.GetFirst();
                while (item != null)
                {
                    try
                    {
                        if ((int)item.Class == OL_MAIL_CLASS)
                        {
                            string eid = (string)item.EntryID;
                            if (!_exported.Contains(eid))
                            {
                                // Export files immediately (msg, body, meta, attachments)
                                var sr = ExportAndBuildResult(item, folderRoot, outputRoot, smtp);
                                if (sr != null)
                                {
                                    results.Add(sr);
                                    _exported.Add(eid);
                                    OnProgress(results.Count, sr.Subject);
                                }
                            }
                        }
                    }
                    catch { }
                    if (CancelRequested) break;
                    _itemCount++;
                    try { item = items.GetNext(); } catch { break; }
                }

                if (recurse && !CancelRequested)
                    foreach (dynamic child in folder.Folders)
                    {
                        if (CancelRequested) break;
                        ScanTree(child, outputRoot, smtp, dateFilter, keywordFilter, results, true);
                    }
            }
            catch { }
        }

        ScanResult ExportAndBuildResult(dynamic mail, string folderRoot, string exportRoot, string smtp)
        {
            try
            {
                string mailDir = Path.Combine(folderRoot, BuildMailDirName(mail, folderRoot));
                if (File.Exists(Path.Combine(mailDir, "meta.json"))) return null;

                Directory.CreateDirectory(mailDir);
                mail.SaveAs(Path.Combine(mailDir, "mail.msg"), OL_MSG_UNICODE);
                string bodyText = (string)mail.Body ?? "";
                File.WriteAllText(Path.Combine(mailDir, "body.txt"), bodyText, Encoding.UTF8);

                var attNames = SaveAttachments(mail, mailDir);
                WriteMetaJson(Path.Combine(mailDir, "meta.json"), mail, attNames, smtp);

                string senderAddr = "";
                try { senderAddr = (string)mail.SenderEmailAddress; } catch { }

                string bodyFlat = bodyText.Replace(",", " ").Replace("\r", " ").Replace("\n", " ");
                if (bodyFlat.Length > 2000) bodyFlat = bodyFlat.Substring(0, 2000);

                return new ScanResult {
                    ItemId = (string)mail.EntryID,
                    Name = (string)mail.Subject,
                    SourcePath = (string)mail.Parent.FolderPath,
                    SenderEmail = senderAddr,
                    SenderName = (string)mail.SenderName,
                    Subject = (string)mail.Subject,
                    ReceivedAt = (DateTime)mail.ReceivedTime,
                    BodyText = bodyFlat,
                    BodyPath = Path.Combine(mailDir, "body.txt"),
                    MsgPath = Path.Combine(mailDir, "mail.msg"),
                    AttachmentPaths = BuildAttachPaths(attNames, mailDir),
                    AttachmentNames = attNames,
                    ItemFolder = mailDir
                };
            }
            catch { return null; }
        }

        // --- Filter construction ---

        // Build date filter using Jet syntax (uses local time, not UTC).
        // Kept separate from keyword filter because Jet and DASL cannot be
        // mixed in a single Restrict() call.
        string BuildDateFilter(string sinceDate, Dictionary<string, string> config)
        {
            DateTime dt;
            string[] dateFmts = { "yyyy-MM-dd", "yyyy/MM/dd", "yyyy/M/d", "M/d/yyyy" };

            // Priority 1: explicit user-configured "since" date
            if (!string.IsNullOrEmpty(sinceDate) && DateTime.TryParseExact(sinceDate, dateFmts,
                CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return string.Format("[ReceivedTime]>='{0:yyyy/MM/dd} 00:00'", dt);

            if (!string.IsNullOrEmpty(sinceDate))
                System.Diagnostics.Debug.WriteLine(
                    "MailScanner: failed to parse since date: " + sinceDate);

            // Priority 2: last successful scan date (only when prior exports exist)
            if (_exported.Count > 0)
            {
                string lastScan = config.ContainsKey("last_scan") ? config["last_scan"] : "";
                if (!string.IsNullOrEmpty(lastScan) && DateTime.TryParseExact(lastScan,
                    "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                    return string.Format("[ReceivedTime]>='{0:yyyy/MM/dd} 00:00'",
                        dt.AddDays(-1));
            }
            // Priority 3: no date filter
            return null;
        }

        // Build keyword filter using DASL syntax for server-side text search.
        // This avoids per-item COM calls for Body/Subject/Sender.
        string BuildKeywordFilter()
        {
            if (_filterWords.Count == 0) return null;

            const string subj = "\"urn:schemas:httpmail:subject\"";
            const string body = "\"urn:schemas:httpmail:textdescription\"";
            const string from = "\"urn:schemas:httpmail:fromemail\"";
            var parts = new List<string>();

            if (_filterMode == "and")
            {
                // Each keyword must appear somewhere in subject/body/sender
                foreach (var kw in _filterWords)
                {
                    string esc = DaslEscape(kw);
                    parts.Add(string.Format(
                        "({0} LIKE '%{1}%' OR {2} LIKE '%{1}%' OR {3} LIKE '%{1}%')",
                        subj, esc, body, from));
                }
            }
            else
            {
                // Any keyword matches in subject/body/sender
                var orClauses = new List<string>();
                foreach (var kw in _filterWords)
                {
                    string esc = DaslEscape(kw);
                    orClauses.Add(string.Format("{0} LIKE '%{1}%'", subj, esc));
                    orClauses.Add(string.Format("{0} LIKE '%{1}%'", body, esc));
                    orClauses.Add(string.Format("{0} LIKE '%{1}%'", from, esc));
                }
                parts.Add("(" + string.Join(" OR ", orClauses.ToArray()) + ")");
            }

            return "@SQL=" + string.Join(" AND ", parts.ToArray());
        }

        // Escape single quotes in DASL string values
        static string DaslEscape(string value)
        {
            return (value ?? "").Replace("'", "''");
        }

        // --- Helpers ---

        // Remove manifest entries whose output folder (mail_folder column) no longer exists
        static void PurgeStaleMailIds(string outputRoot, HashSet<string> exported)
        {
            var rows = ManifestIO.LoadRows(outputRoot);
            var staleIds = new HashSet<string>();
            foreach (var kv in rows)
            {
                // mail_folder is column index 9 (mail_folder)
                if (kv.Value.Length > 9)
                {
                    string mailFolder = kv.Value[9];
                    if (!string.IsNullOrEmpty(mailFolder) && !Directory.Exists(mailFolder))
                    {
                        staleIds.Add(kv.Key);
                        exported.Remove(kv.Key);
                    }
                }
            }
            if (staleIds.Count > 0)
                ManifestIO.RemoveRows(outputRoot, staleIds);
        }

        static string BuildAttachPaths(List<string> names, string dir)
        {
            var paths = new List<string>();
            foreach (var n in names) paths.Add(Path.Combine(dir, n));
            return string.Join("|", paths.ToArray());
        }

        dynamic FindFolder(dynamic root, string targetPath)
        {
            try
            {
                if (string.Equals((string)root.FolderPath, targetPath,
                    StringComparison.OrdinalIgnoreCase)) return root;
            }
            catch { }
            try
            {
                foreach (dynamic child in root.Folders)
                {
                    var found = FindFolder(child, targetPath);
                    if (found != null) return found;
                }
            }
            catch { }
            return null;
        }

        List<string> SaveAttachments(dynamic mail, string mailDir)
        {
            var names = new List<string>();
            try
            {
                for (int i = 1; i <= (int)mail.Attachments.Count; i++)
                {
                    try
                    {
                        dynamic att = mail.Attachments[i];
                        string safeFn = SafeName((string)att.FileName);
                        att.SaveAsFile(Path.Combine(mailDir, safeFn));
                        names.Add(safeFn);
                    }
                    catch { }
                }
            }
            catch { }
            return names;
        }

        void WriteMetaJson(string path, dynamic mail, List<string> attNames, string smtp)
        {
            var attJson = new StringBuilder("[");
            for (int i = 0; i < attNames.Count; i++)
            {
                if (i > 0) attJson.Append(", ");
                attJson.AppendFormat("{{\"path\": \"{0}\"}}", JsonEsc(attNames[i]));
            }
            attJson.Append("]");

            string senderAddr = "";
            try { senderAddr = (string)mail.SenderEmailAddress; } catch { }

            var json = string.Format(
                "{{\n  \"entry_id\": \"{0}\",\n  \"mailbox_address\": \"{1}\",\n" +
                "  \"folder_path\": \"{2}\",\n  \"sender_name\": \"{3}\",\n" +
                "  \"sender_email\": \"{4}\",\n  \"subject\": \"{5}\",\n" +
                "  \"received_at\": \"{6:yyyy-MM-dd\\THH:mm:ss}\",\n" +
                "  \"body_path\": \"body.txt\",\n  \"msg_path\": \"mail.msg\",\n" +
                "  \"attachments\": {7}\n}}",
                JsonEsc((string)mail.EntryID), JsonEsc(smtp),
                JsonEsc((string)mail.Parent.FolderPath), JsonEsc((string)mail.SenderName),
                JsonEsc(senderAddr), JsonEsc((string)mail.Subject),
                (DateTime)mail.ReceivedTime, attJson);

            File.WriteAllText(path, json, Encoding.UTF8);
        }

        // Fallback keyword filter for cases where DASL filtering cannot be used.
        // Normally keywords are pushed into Restrict() via BuildDaslFilter().
        bool MatchesFilter(dynamic mail)
        {
            if (_filterWords.Count == 0) return true;
            string subject = ""; string body = ""; string sender = "";
            try { subject = ((string)mail.Subject ?? "").ToLower(); } catch { }
            try { body = ((string)mail.Body ?? "").ToLower(); } catch { }
            try { sender = ((string)mail.SenderEmailAddress ?? "").ToLower(); } catch { }

            string text = subject + "\n" + body + "\n" + sender;
            if (_filterMode == "and")
            {
                foreach (var kw in _filterWords)
                    if (!text.Contains(kw)) return false;
                return true;
            }
            else
            {
                foreach (var kw in _filterWords)
                    if (text.Contains(kw)) return true;
                return false;
            }
        }

        string GetStoreSmtp(dynamic store)
        {
            try
            {
                foreach (dynamic acct in _olNs.Accounts)
                {
                    try
                    {
                        if ((string)acct.DeliveryStore.StoreID == (string)store.StoreID)
                            return ((string)acct.SmtpAddress).ToLower();
                    }
                    catch { }
                }
            }
            catch { }
            try
            {
                return ((string)store.GetRootFolder().PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001E")).ToLower();
            }
            catch { }
            try
            {
                return ((string)store.GetRootFolder().PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001F")).ToLower();
            }
            catch { }
            try { return ((string)store.DisplayName).ToLower(); } catch { return ""; }
        }

        string BuildMailDirName(dynamic mail, string parentDir)
        {
            string ts = ((DateTime)mail.ReceivedTime).ToString("yyyyMMdd_HHmmss");
            string baseName = _shortDirname ? ts : ts + "_" + SafeName((string)mail.Subject);
            // Disambiguate if directory already exists (same-second emails)
            string candidate = baseName;
            int suffix = 2;
            while (Directory.Exists(Path.Combine(parentDir, candidate))
                && File.Exists(Path.Combine(parentDir, candidate, "meta.json")))
            {
                candidate = baseName + "_" + suffix;
                suffix++;
            }
            return candidate;
        }

        string NormalizeFolderPath(string folderPath)
        {
            var parts = folderPath.Split('\\');
            var result = "";
            int skip = 0;
            foreach (var part in parts)
            {
                if (string.IsNullOrEmpty(part)) continue;
                skip++;
                if (skip <= 1) continue;
                result += "\\" + SafeName(part);
            }
            return result;
        }

        static string SafeName(string value)
        {
            var text = (value ?? "").Trim();
            if (text.Length == 0) text = "blank";
            var sb = new StringBuilder();
            foreach (char c in text)
            {
                if (c < 32 || char.IsSurrogate(c)) continue;
                if ("\\/:|*?\"<>".IndexOf(c) >= 0) sb.Append('_');
                else sb.Append(c);
            }
            var s = sb.ToString();
            if (s.Length == 0) s = "blank";
            if (s.Length > 80) s = s.Substring(0, 80);
            return s;
        }

        static string JsonEsc(string value)
        {
            return (value ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"")
                .Replace("\r\n", "\\n").Replace("\r", "\\n").Replace("\n", "\\n");
        }
    }
}
