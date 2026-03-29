using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace MailPull
{
    public class Exporter
    {
        const int OL_MSG_UNICODE = 9;
        const int OL_MAIL_CLASS = 43;

        dynamic _olApp;
        dynamic _olNs;

        public event Action<int, string> ProgressChanged;
        public volatile bool CancelRequested;
        HashSet<string> _exported;
        int _itemCount;
        string _filterMode;
        List<string> _filterWords;

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

        // --- Accounts / Folders ---

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

        // --- Export ---

        public int Export(string exportRoot, string sinceDate, string filterAccount, string filterFolder,
            string filterMode = "", string filterKeywords = "")
        {
            if (!Connect() || string.IsNullOrEmpty(exportRoot)) return 0;
            Directory.CreateDirectory(exportRoot);

            // Parse keyword filters
            _filterMode = (filterMode ?? "").ToLower() == "and" ? "and" : "or";
            _filterWords = new List<string>();
            if (!string.IsNullOrEmpty(filterKeywords))
                foreach (var kw in filterKeywords.Split(';'))
                    if (kw.Trim().Length > 0) _filterWords.Add(kw.Trim().ToLower());

            // Pre-load exported EntryIDs from manifest (skip without COM calls)
            var exported = new HashSet<string>();
            string manifestPath = Path.Combine(exportRoot, "manifest.csv");
            if (File.Exists(manifestPath))
            {
                foreach (var line in File.ReadAllLines(manifestPath, Encoding.UTF8))
                {
                    if (string.IsNullOrEmpty(line)) continue;
                    int sep = line.IndexOf(',');
                    exported.Add(sep > 0 ? line.Substring(0, sep) : line);
                }
            }
            _exported = exported;

            string filter = null;
            if (!string.IsNullOrEmpty(sinceDate))
            {
                DateTime dt;
                if (DateTime.TryParse(sinceDate, out dt))
                    filter = string.Format("[ReceivedTime]>='{0:yyyy/MM/dd}'", dt);
            }

            int total = 0;
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
                            total += ExportTree(startFolder, exportRoot, smtp, filter);
                    }
                    else
                        total += ExportTree(store.GetRootFolder(), exportRoot, smtp, filter);
                }
                catch { }
            }
            return total;
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

        int ExportTree(dynamic folder, string exportRoot, string smtp, string filter)
        {
            int total = 0;
            try
            {
                string folderRoot = Path.Combine(exportRoot,
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
                            // Fast skip: check EntryID against manifest (O(1), one COM call)
                            string eid = (string)item.EntryID;
                            if (!_exported.Contains(eid))
                            {
                                // Keyword filter check
                                if (_filterWords.Count > 0 && !MatchesFilter(item))
                                    goto SkipItem;

                                ExportMail(item, folderRoot, exportRoot, smtp);
                                _exported.Add(eid);
                                total++;
                                if (ProgressChanged != null)
                                    ProgressChanged(total, (string)item.Subject);
                            }
                            SkipItem:;
                        }
                    }
                    catch { }
                    if (CancelRequested) break;
                    // Yield every 10 items to reduce Outlook freezing
                    _itemCount++;
                    if (_itemCount % 10 == 0) System.Threading.Thread.Sleep(1);
                    try { item = items.GetNext(); } catch { break; }
                }

                if (!CancelRequested)
                    foreach (dynamic child in folder.Folders)
                    {
                        if (CancelRequested) break;
                        total += ExportTree(child, exportRoot, smtp, filter);
                    }
            }
            catch { }
            return total;
        }

        void ExportMail(dynamic mail, string folderRoot, string exportRoot, string smtp)
        {
            string mailDir = Path.Combine(folderRoot, BuildMailDirName(mail));
            if (File.Exists(Path.Combine(mailDir, "meta.json"))) return;

            Directory.CreateDirectory(mailDir);
            mail.SaveAs(Path.Combine(mailDir, "mail.msg"), OL_MSG_UNICODE);
            File.WriteAllText(Path.Combine(mailDir, "body.txt"), (string)mail.Body, Encoding.UTF8);

            var attNames = SaveAttachments(mail, mailDir);
            WriteMetaJson(Path.Combine(mailDir, "meta.json"), mail, attNames, smtp);
            AppendManifest(exportRoot, mail, mailDir, attNames);
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

        void AppendManifest(string exportRoot, dynamic mail, string mailDir, List<string> attNames)
        {
            try
            {
                string senderAddr = "";
                try { senderAddr = (string)mail.SenderEmailAddress; } catch { }

                // Flatten body to single line for CSV search
                string bodyText = "";
                try { bodyText = ((string)mail.Body ?? "").Replace(",", " ").Replace("\r", " ").Replace("\n", " "); }
                catch { }
                // Truncate to keep CSV manageable
                if (bodyText.Length > 2000) bodyText = bodyText.Substring(0, 2000);

                string line = string.Join(",", new[] {
                    (string)mail.EntryID, senderAddr, (string)mail.SenderName,
                    ((string)mail.Subject).Replace(",", " ").Replace("\n", " "),
                    ((DateTime)mail.ReceivedTime).ToString("yyyy-MM-dd\\THH:mm:ss"),
                    (string)mail.Parent.FolderPath,
                    Path.Combine(mailDir, "body.txt"), Path.Combine(mailDir, "mail.msg"),
                    string.Join("|", attNames.ConvertAll(a => Path.Combine(mailDir, a)).ToArray()),
                    mailDir,
                    bodyText
                });

                var csvPath = Path.Combine(exportRoot, "manifest.csv");
                // BOM on first write so Excel opens correctly
                if (!File.Exists(csvPath))
                    File.WriteAllText(csvPath,
                        "entry_id,sender_email,sender_name,subject,received_at,folder_path,body_path,msg_path,attachment_paths,mail_folder,body_text"
                        + Environment.NewLine, new UTF8Encoding(true));
                File.AppendAllText(csvPath, line + Environment.NewLine, new UTF8Encoding(true));
            }
            catch { }
        }

        // --- Keyword filter ---

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

        // --- Search manifest ---

        // Search manifest.csv by subject, sender_email, sender_name, attachment filenames.
        // CSV columns: 0=entry_id 1=sender_email 2=sender_name 3=subject 4=received_at
        //   5=folder_path 6=body_path 7=msg_path 8=attachment_paths 9=mail_folder
        public static List<string[]> SearchManifest(string query)
        {
            var results = new List<string[]>();
            string root = Config.Get("export_root");
            if (string.IsNullOrEmpty(root)) return results;
            string path = Path.Combine(root, "manifest.csv");
            if (!File.Exists(path)) return results;

            string q = (query ?? "").Trim().ToLower();
            if (q.Length == 0) return results;

            foreach (var line in File.ReadAllLines(path, Encoding.UTF8))
            {
                if (string.IsNullOrEmpty(line)) continue;
                var cols = line.Split(',');
                if (cols.Length < 4) continue;
                // skip header
                if (cols[0] == "entry_id") continue;

                string email   = cols.Length > 1 ? cols[1].ToLower() : "";
                string name    = cols.Length > 2 ? cols[2].ToLower() : "";
                string subject = cols.Length > 3 ? cols[3].ToLower() : "";
                string attach  = cols.Length > 8 ? cols[8].ToLower() : "";
                string body    = cols.Length > 10 ? cols[10].ToLower() : "";

                if (subject.Contains(q) || email.Contains(q)
                    || name.Contains(q) || attach.Contains(q)
                    || body.Contains(q))
                    results.Add(cols);
            }
            return results;
        }

        // --- Helpers ---

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
