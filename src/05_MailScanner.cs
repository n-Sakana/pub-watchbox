using System;
using System.Collections.Generic;
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
                foreach (dynamic acct in _olNs.Accounts)
                {
                    try { if (!string.IsNullOrEmpty((string)acct.SmtpAddress)) list.Add((string)acct.SmtpAddress); }
                    catch { }
                }
                var known = new HashSet<string>();
                foreach (dynamic acct in _olNs.Accounts)
                {
                    try { known.Add((string)acct.DeliveryStore.StoreID); } catch { }
                }
                foreach (dynamic store in _olNs.Stores)
                {
                    try
                    {
                        if (known.Contains((string)store.StoreID)) continue;
                        string addr = "";
                        try { addr = ((string)store.GetRootFolder().PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E")).ToLower(); } catch { }
                        if (string.IsNullOrEmpty(addr)) addr = ((string)store.DisplayName).ToLower();
                        if (!string.IsNullOrEmpty(addr)) list.Add(addr);
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
                    foreach (dynamic acct in _olNs.Accounts)
                    {
                        if (string.Equals((string)acct.SmtpAddress, accountFilter,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            CollectFolders(acct.DeliveryStore.GetRootFolder(), 0, "", list);
                            break;
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
            string outputRoot = config.ContainsKey("output_root") ? config["output_root"] : "";
            if (!Connect() || string.IsNullOrEmpty(outputRoot)) return results;
            Directory.CreateDirectory(outputRoot);

            string filterAccount = config.ContainsKey("account") ? config["account"] : "";
            string filterFolder = config.ContainsKey("outlook_folder") ? config["outlook_folder"] : "";
            string sinceDate = config.ContainsKey("since") ? config["since"] : "";
            string filterMode = config.ContainsKey("filter_mode") ? config["filter_mode"] : "";
            string filterKeywords = config.ContainsKey("filters") ? config["filters"] : "";
            _flatOutput = config.ContainsKey("flat_output") && config["flat_output"] == "1";

            _filterMode = (filterMode ?? "").ToLower() == "and" ? "and" : "or";
            _filterWords = new List<string>();
            if (!string.IsNullOrEmpty(filterKeywords))
                foreach (var kw in filterKeywords.Split(';'))
                    if (kw.Trim().Length > 0) _filterWords.Add(kw.Trim().ToLower());

            _exported = new HashSet<string>(knownIds);

            string filter = null;
            DateTime dt;
            if (!string.IsNullOrEmpty(sinceDate) && DateTime.TryParse(sinceDate, out dt))
                filter = string.Format("[ReceivedTime]>='{0:yyyy/MM/dd}'", dt);
            else if (knownIds.Count > 0)
            {
                // Already have data: use latest date from manifest to narrow scan
                DateTime latest = ManifestIO.GetLatestMailDate(outputRoot);
                if (latest > DateTime.MinValue)
                    filter = string.Format("[ReceivedTime]>='{0:yyyy/MM/dd}'",
                        latest.AddDays(-1));
            }

            foreach (dynamic store in _olNs.Stores)
            {
                if (CancelRequested) break;
                try
                {
                    string smtp = GetStoreSmtp(store);
                    if (string.IsNullOrEmpty(smtp)) continue;
                    if (!string.IsNullOrEmpty(filterAccount) &&
                        !string.Equals(smtp, filterAccount, StringComparison.OrdinalIgnoreCase)) continue;

                    if (!string.IsNullOrEmpty(filterFolder))
                    {
                        dynamic startFolder = FindFolder(store.GetRootFolder(), filterFolder);
                        if (startFolder != null)
                            ScanTree(startFolder, outputRoot, smtp, filter, results);
                    }
                    else
                        ScanTree(store.GetRootFolder(), outputRoot, smtp, filter, results);
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

        void ScanTree(dynamic folder, string outputRoot, string smtp, string filter,
            List<ScanResult> results)
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
                if (filter != null) items = items.Restrict(filter);

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
                                if (_filterWords.Count > 0 && !MatchesFilter(item))
                                    goto SkipItem;

                                // Export files immediately (msg, body, meta, attachments)
                                var sr = ExportAndBuildResult(item, folderRoot, outputRoot, smtp);
                                if (sr != null)
                                {
                                    results.Add(sr);
                                    _exported.Add(eid);
                                    OnProgress(results.Count, sr.Subject);
                                }
                                SkipItem:;
                            }
                        }
                    }
                    catch { }
                    if (CancelRequested) break;
                    _itemCount++;
                    if (_itemCount % 10 == 0) System.Threading.Thread.Sleep(1);
                    try { item = items.GetNext(); } catch { break; }
                }

                if (!CancelRequested)
                    foreach (dynamic child in folder.Folders)
                    {
                        if (CancelRequested) break;
                        ScanTree(child, outputRoot, smtp, filter, results);
                    }
            }
            catch { }
        }

        ScanResult ExportAndBuildResult(dynamic mail, string folderRoot, string exportRoot, string smtp)
        {
            try
            {
                string mailDir = Path.Combine(folderRoot, BuildMailDirName(mail));
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

        // --- Helpers ---

        static string BuildAttachPaths(List<string> names, string dir)
        {
            var paths = new List<string>();
            foreach (var n in names) paths.Add(Path.Combine(dir, n));
            return string.Join("|", paths.ToArray());
        }

        dynamic FindFolder(dynamic root, string targetPath)
        {
            try { if ((string)root.FolderPath == targetPath) return root; } catch { }
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
            try { return ((string)store.DisplayName).ToLower(); } catch { return ""; }
        }

        string BuildMailDirName(dynamic mail)
        {
            return ((DateTime)mail.ReceivedTime).ToString("yyyyMMdd_HHmmss") + "_" +
                SafeName((string)mail.Subject);
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
