using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace WatchBox
{
    // Lightweight cache of mail item data for cross-profile optimization.
    // Holds text data only (no COM references) so items can be filtered
    // in memory without COM overhead.
    public class CachedMailItem
    {
        public string EntryID;
        public string Subject;
        public string SenderEmail;
        public string SenderName;
        public string BodyLower;      // lowercased for keyword matching
        public string SubjectLower;
        public string SenderEmailLower;
        public DateTime ReceivedTime;
        public string FolderPath;
    }

    public class MailScanner : SourceScanner
    {
        const int OL_MSG_UNICODE = 9;
        const int OL_MAIL_CLASS = 43;

        dynamic _olApp;
        dynamic _olNs;
        HashSet<string> _exported;
        HashSet<string> _dedupKeys;
        int _itemCount;
        string _filterMode;
        List<string> _filterWords;
        bool _flatOutput;
        bool _shortDirname;
        bool _autoUnzip;

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

        public void Cleanup()
        {
            try { if (_olNs != null) Marshal.ReleaseComObject(_olNs); } catch { }
            try { if (_olApp != null) Marshal.ReleaseComObject(_olApp); } catch { }
            _olNs = null;
            _olApp = null;
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
                    finally { try { Marshal.ReleaseComObject(acct); } catch { } }
                }
                // Also enumerate stores not tied to any account (shared/delegate mailboxes)
                var accountStoreIds = new HashSet<string>();
                foreach (dynamic acct in _olNs.Accounts)
                {
                    try
                    {
                        dynamic ds = acct.DeliveryStore;
                        accountStoreIds.Add((string)ds.StoreID);
                        try { Marshal.ReleaseComObject(ds); } catch { }
                    }
                    catch { }
                    finally { try { Marshal.ReleaseComObject(acct); } catch { } }
                }
                foreach (dynamic store in _olNs.Stores)
                {
                    try
                    {
                        if (accountStoreIds.Contains((string)store.StoreID)) continue;
                        string addr = GetStoreSmtp(store);
                        if (string.IsNullOrEmpty(addr))
                            try { addr = ((string)store.DisplayName).ToLower(); } catch { }
                        if (!string.IsNullOrEmpty(addr) && seen.Add(addr))
                            list.Add(addr);
                    }
                    catch { }
                    finally { try { Marshal.ReleaseComObject(store); } catch { } }
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
                            {
                                dynamic rootF = store.GetRootFolder();
                                try { CollectFolders(rootF, 0, smtp + ": ", list); }
                                finally { try { Marshal.ReleaseComObject(rootF); } catch { } }
                            }
                        }
                        catch { }
                        finally { try { Marshal.ReleaseComObject(store); } catch { } }
                    }
                }
                else
                {
                    // Try matching account first, then fall back to store lookup
                    bool found = false;
                    foreach (dynamic acct in _olNs.Accounts)
                    {
                        try
                        {
                            if (string.Equals((string)acct.SmtpAddress, accountFilter,
                                StringComparison.OrdinalIgnoreCase))
                            {
                                dynamic ds = acct.DeliveryStore;
                                dynamic rootF = ds.GetRootFolder();
                                try { CollectFolders(rootF, 0, "", list); }
                                finally
                                {
                                    try { Marshal.ReleaseComObject(rootF); } catch { }
                                    try { Marshal.ReleaseComObject(ds); } catch { }
                                }
                                found = true;
                                break;
                            }
                        }
                        catch { }
                        finally { try { Marshal.ReleaseComObject(acct); } catch { } }
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
                                    dynamic rootF = store.GetRootFolder();
                                    try { CollectFolders(rootF, 0, "", list); }
                                    finally { try { Marshal.ReleaseComObject(rootF); } catch { } }
                                    break;
                                }
                            }
                            catch { }
                            finally { try { Marshal.ReleaseComObject(store); } catch { } }
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
                    finally { try { Marshal.ReleaseComObject(child); } catch { } }
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
            _autoUnzip = config.ContainsKey("auto_unzip") && config["auto_unzip"] == "1";

            _filterMode = (filterMode ?? "").ToLower() == "and" ? "and" : "or";
            _filterWords = new List<string>();
            if (!string.IsNullOrEmpty(filterKeywords))
                foreach (var kw in filterKeywords.Split(';'))
                    if (kw.Trim().Length > 0) _filterWords.Add(kw.Trim().ToLower());

            // Remove IDs whose output folders were manually deleted
            _exported = new HashSet<string>(knownIds);
            PurgeStaleMailIds(outputRoot, _exported);
            // Load dedup keys AFTER purge so stale entries are excluded
            _dedupKeys = ManifestIO.LoadDedupKeys(outputRoot);

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

            var stores = new List<dynamic>();
            foreach (dynamic s in _olNs.Stores) stores.Add(s);
            foreach (dynamic store in stores)
            {
                if (CancelRequested) break;
                try
                {
                    string smtp = GetStoreSmtp(store);
                    if (string.IsNullOrEmpty(smtp)) continue;
                    if (!string.Equals(smtp, filterAccount, StringComparison.OrdinalIgnoreCase)) continue;

                    dynamic rootFolder = store.GetRootFolder();
                    try
                    {
                        if (!string.IsNullOrEmpty(filterFolder))
                        {
                            dynamic startFolder = FindFolder(rootFolder, filterFolder);
                            if (startFolder != null)
                                ScanTree(startFolder, outputRoot, smtp, dateFilter, keywordFilter, results, false);
                        }
                        else
                            ScanTree(rootFolder, outputRoot, smtp, dateFilter, keywordFilter, results, true);
                    }
                    finally
                    {
                        try { Marshal.ReleaseComObject(rootFolder); } catch { }
                    }
                }
                catch { }
                finally
                {
                    try { Marshal.ReleaseComObject(store); } catch { }
                }
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

                dynamic rawItems = folder.Items;
                dynamic items = rawItems;
                dynamic items2 = null;
                // Chain two Restrict() calls: date (Jet, local time) then keywords (DASL)
                if (dateFilter != null) { items = items.Restrict(dateFilter); items2 = items; }
                if (keywordFilter != null) items = items.Restrict(keywordFilter);

                dynamic item = items.GetFirst();
                while (item != null)
                {
                    try
                    {
                        if ((int)item.Class == OL_MAIL_CLASS)
                        {
                            string eid = (string)item.EntryID;

                            // Cache mode: collect item data without exporting
                            if (_cacheTarget != null)
                            {
                                string subj = ""; string body = "";
                                string senderAddr = ""; string senderName = "";
                                try { subj = (string)item.Subject ?? ""; } catch { }
                                try { body = (string)item.Body ?? ""; } catch { }
                                try { senderAddr = (string)item.SenderEmailAddress ?? ""; } catch { }
                                try { senderName = (string)item.SenderName ?? ""; } catch { }
                                dynamic itemParent = item.Parent;
                                string itemFolderPath = (string)itemParent.FolderPath;
                                try { Marshal.ReleaseComObject(itemParent); } catch { }
                                _cacheTarget.Add(new CachedMailItem {
                                    EntryID = eid,
                                    Subject = subj,
                                    SenderEmail = senderAddr,
                                    SenderName = senderName,
                                    BodyLower = body.ToLower(),
                                    SubjectLower = subj.ToLower(),
                                    SenderEmailLower = senderAddr.ToLower(),
                                    ReceivedTime = (DateTime)item.ReceivedTime,
                                    FolderPath = itemFolderPath
                                });
                                OnProgress(_cacheTarget.Count, subj);
                            }
                            else if (!_exported.Contains(eid))
                            {
                                // Secondary dedup: skip if subject+sender+date match
                                string subj2 = ""; string saddr2 = "";
                                DateTime recv2 = DateTime.MinValue;
                                try { subj2 = ((string)item.Subject ?? "").ToLower(); } catch { }
                                try { saddr2 = ((string)item.SenderEmailAddress ?? "").ToLower(); } catch { }
                                try { recv2 = (DateTime)item.ReceivedTime; } catch { }
                                string dedupKey = subj2 + "|" + saddr2 + "|" +
                                    recv2.ToString("yyyy-MM-dd\\THH:mm:ss");
                                if (_dedupKeys != null && _dedupKeys.Contains(dedupKey))
                                    continue;

                                // Export files immediately (msg, body, meta, attachments)
                                var sr = ExportAndBuildResult(item, folderRoot, outputRoot, smtp);
                                if (sr != null)
                                {
                                    results.Add(sr);
                                    _exported.Add(eid);
                                    if (_dedupKeys != null) _dedupKeys.Add(dedupKey);
                                    OnProgress(results.Count, sr.Subject);
                                }
                            }
                        }
                    }
                    catch { }
                    finally
                    {
                        try { Marshal.ReleaseComObject(item); } catch { }
                    }
                    if (CancelRequested) break;
                    _itemCount++;
                    if (_itemCount % 5 == 0) Thread.Sleep(50);
                    try { item = items.GetNext(); } catch { break; }
                }
                try { Marshal.ReleaseComObject(items); } catch { }
                if (items2 != null && !object.ReferenceEquals(items2, items))
                    try { Marshal.ReleaseComObject(items2); } catch { }
                if (!object.ReferenceEquals(rawItems, items) &&
                    (items2 == null || !object.ReferenceEquals(rawItems, items2)))
                    try { Marshal.ReleaseComObject(rawItems); } catch { }

                if (recurse && !CancelRequested)
                    foreach (dynamic child in folder.Folders)
                    {
                        if (CancelRequested) break;
                        try { ScanTree(child, outputRoot, smtp, dateFilter, keywordFilter, results, true); }
                        finally { try { Marshal.ReleaseComObject(child); } catch { } }
                    }
            }
            catch { }
        }

        // --- Bulk scan: uses existing Scan() flow but caches instead of exporting ---

        // Scan an account+folder with date-only filter and cache all item data.
        // Runs through the proven Scan() code path with a special config that
        // triggers caching mode (_cacheTarget is set).
        List<CachedMailItem> _cacheTarget;

        public List<CachedMailItem> ScanBulk(
            Dictionary<string, string> config, string dateFilter)
        {
            // Build a minimal config for Scan() — no keywords, no export
            var scanConfig = new Dictionary<string, string>();
            scanConfig["account"] = config.ContainsKey("account") ? config["account"] : "";
            scanConfig["outlook_folder"] = config.ContainsKey("outlook_folder") ? config["outlook_folder"] : "";
            scanConfig["output_root"] = Path.GetTempPath();
            scanConfig["since"] = "";
            scanConfig["filter_mode"] = "or";
            scanConfig["filters"] = "";
            scanConfig["flat_output"] = "1";
            scanConfig["source_folder"] = "";
            scanConfig["recurse"] = "1";
            scanConfig["type"] = "mail";
            scanConfig["short_dirname"] = "0";
            scanConfig["auto_unzip"] = "0";
            // Override date filter via last_scan if dateFilter provided, else no filter
            scanConfig["last_scan"] = "";
            // Use since to control the date filter
            if (dateFilter != null)
            {
                // Extract date from Jet filter: [ReceivedTime]>='yyyy/MM/dd 00:00'
                var m = System.Text.RegularExpressions.Regex.Match(
                    dateFilter, @"(\d{4}/\d{2}/\d{2})");
                if (m.Success)
                    scanConfig["since"] = m.Groups[1].Value.Replace("/", "-");
            }

            _cacheTarget = new List<CachedMailItem>();
            Scan(scanConfig, new HashSet<string>());
            var result = _cacheTarget;
            _cacheTarget = null;
            return result;
        }

        // Export items from cache that match a profile's keyword filter.
        // Uses GetItemFromID to fetch COM objects only for items that need export.
        public List<ScanResult> ExportFromCache(
            List<CachedMailItem> cache,
            Dictionary<string, string> config,
            HashSet<string> knownIds)
        {
            var results = new List<ScanResult>();
            string outputRoot = config.ContainsKey("output_root") ? config["output_root"].Trim() : "";
            if (string.IsNullOrEmpty(outputRoot)) return results;
            Directory.CreateDirectory(outputRoot);

            string filterAccount = config.ContainsKey("account") ? config["account"] : "";
            string filterMode = config.ContainsKey("filter_mode") ? config["filter_mode"] : "";
            string filterKeywords = config.ContainsKey("filters") ? config["filters"] : "";
            _flatOutput = config.ContainsKey("flat_output") && config["flat_output"] == "1";
            _shortDirname = config.ContainsKey("short_dirname") && config["short_dirname"] == "1";
            _autoUnzip = config.ContainsKey("auto_unzip") && config["auto_unzip"] == "1";

            string mode = (filterMode ?? "").ToLower() == "and" ? "and" : "or";
            var words = new List<string>();
            if (!string.IsNullOrEmpty(filterKeywords))
                foreach (var kw in filterKeywords.Split(';'))
                    if (kw.Trim().Length > 0) words.Add(kw.Trim().ToLower());

            // Parse per-profile date range (the cache may have a broader range)
            DateTime sinceDate = DateTime.MinValue;
            string sinceStr = config.ContainsKey("since") ? config["since"] : "";
            string[] dateFmts = { "yyyy-MM-dd", "yyyy/MM/dd", "yyyy/M/d", "M/d/yyyy" };
            if (!string.IsNullOrEmpty(sinceStr))
                DateTime.TryParseExact(sinceStr, dateFmts,
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out sinceDate);

            _exported = new HashSet<string>(knownIds);
            PurgeStaleMailIds(outputRoot, _exported);
            // Load dedup keys AFTER purge so stale entries are excluded
            _dedupKeys = ManifestIO.LoadDedupKeys(outputRoot);

            string smtp = (filterAccount ?? "").ToLower();
            string filterFolder = config.ContainsKey("outlook_folder") ? config["outlook_folder"] : "";

            foreach (var ci in cache)
            {
                if (CancelRequested) break;

                // Folder path check (cache may contain items from multiple folders)
                if (!string.IsNullOrEmpty(filterFolder) &&
                    !string.Equals(ci.FolderPath, filterFolder, StringComparison.OrdinalIgnoreCase))
                    continue;

                // Per-profile date check
                if (sinceDate > DateTime.MinValue && ci.ReceivedTime < sinceDate) continue;

                // Already exported
                if (_exported.Contains(ci.EntryID)) continue;

                // Secondary dedup: subject+sender+date composite key
                string dedupKey = ci.SubjectLower + "|" + ci.SenderEmailLower + "|" +
                    ci.ReceivedTime.ToString("yyyy-MM-dd\\THH:mm:ss");
                if (_dedupKeys != null && _dedupKeys.Contains(dedupKey)) continue;

                // Keyword match in memory
                if (words.Count > 0)
                {
                    string text = ci.SubjectLower + "\n" + ci.BodyLower + "\n" + ci.SenderEmailLower;
                    if (mode == "and")
                    {
                        bool allMatch = true;
                        foreach (var kw in words)
                            if (!text.Contains(kw)) { allMatch = false; break; }
                        if (!allMatch) continue;
                    }
                    else
                    {
                        bool anyMatch = false;
                        foreach (var kw in words)
                            if (text.Contains(kw)) { anyMatch = true; break; }
                        if (!anyMatch) continue;
                    }
                }

                // Fetch live COM object and export
                dynamic mail = null;
                try
                {
                    mail = _olNs.GetItemFromID(ci.EntryID);
                    string folderRoot;
                    if (_flatOutput)
                        folderRoot = outputRoot;
                    else
                        folderRoot = Path.Combine(outputRoot,
                            SafeName(smtp) + NormalizeFolderPath(ci.FolderPath));
                    Directory.CreateDirectory(folderRoot);

                    var sr = ExportAndBuildResult(mail, folderRoot, outputRoot, smtp);
                    if (sr != null)
                    {
                        results.Add(sr);
                        _exported.Add(ci.EntryID);
                        if (_dedupKeys != null) _dedupKeys.Add(dedupKey);
                        OnProgress(results.Count, sr.Subject);
                    }
                }
                catch { }
                finally
                {
                    if (mail != null)
                        try { Marshal.ReleaseComObject(mail); } catch { }
                }
            }
            return results;
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

                string senderAddr = "";
                try { senderAddr = (string)mail.SenderEmailAddress; } catch { }
                string senderName = (string)mail.SenderName;
                string subject = (string)mail.Subject;
                DateTime receivedAt = (DateTime)mail.ReceivedTime;

                // Extract recipient addresses
                string toRecipients = "";
                string ccRecipients = "";
                try
                {
                    var toList = new List<string>();
                    var ccList = new List<string>();
                    dynamic recipients = mail.Recipients;
                    int count = (int)recipients.Count;
                    for (int ri = 1; ri <= count; ri++)
                    {
                        dynamic recip = null;
                        try
                        {
                            recip = recipients[ri];
                            int recipType = (int)recip.Type;
                            string addr = ResolveRecipientSmtp(recip);
                            if (!string.IsNullOrEmpty(addr))
                            {
                                if (recipType == 1) toList.Add(addr);
                                else if (recipType == 2) ccList.Add(addr);
                            }
                        }
                        catch { }
                        finally
                        {
                            if (recip != null)
                                try { Marshal.ReleaseComObject(recip); } catch { }
                        }
                    }
                    try { Marshal.ReleaseComObject(recipients); } catch { }
                    toRecipients = string.Join(";", toList.ToArray());
                    ccRecipients = string.Join(";", ccList.ToArray());
                }
                catch { }

                // Write body.txt with metadata header
                var bodyBuilder = new StringBuilder();
                bodyBuilder.AppendFormat("From: {0} <{1}>\r\n", senderName, senderAddr);
                if (!string.IsNullOrEmpty(toRecipients))
                    bodyBuilder.AppendFormat("To: {0}\r\n", toRecipients);
                if (!string.IsNullOrEmpty(ccRecipients))
                    bodyBuilder.AppendFormat("CC: {0}\r\n", ccRecipients);
                bodyBuilder.AppendFormat("Date: {0:yyyy-MM-dd HH:mm:ss}\r\n", receivedAt);
                bodyBuilder.AppendFormat("Subject: {0}\r\n", subject);
                bodyBuilder.Append("---\r\n");
                bodyBuilder.Append(bodyText);
                File.WriteAllText(Path.Combine(mailDir, "body.txt"),
                    bodyBuilder.ToString(), Encoding.UTF8);

                var attNames = SaveAttachments(mail, mailDir);
                WriteMetaJson(Path.Combine(mailDir, "meta.json"), mail, attNames, smtp,
                    toRecipients, ccRecipients);

                // body_text for manifest: original body without header (for search)
                string bodyFlat = bodyText.Replace(",", " ").Replace("\r", " ").Replace("\n", " ");
                if (bodyFlat.Length > 2000) bodyFlat = bodyFlat.Substring(0, 2000);

                dynamic parentFolder = mail.Parent;
                string folderPath = (string)parentFolder.FolderPath;
                try { Marshal.ReleaseComObject(parentFolder); } catch { }

                return new ScanResult {
                    ItemId = (string)mail.EntryID,
                    Name = subject,
                    SourcePath = folderPath,
                    SenderEmail = senderAddr,
                    SenderName = senderName,
                    Subject = subject,
                    ReceivedAt = receivedAt,
                    BodyText = bodyFlat,
                    BodyPath = Path.Combine(mailDir, "body.txt"),
                    MsgPath = Path.Combine(mailDir, "mail.msg"),
                    AttachmentPaths = BuildAttachPaths(attNames, mailDir),
                    AttachmentNames = attNames,
                    ItemFolder = mailDir,
                    ToRecipients = toRecipients,
                    CcRecipients = ccRecipients
                };
            }
            catch { return null; }
        }

        // Resolve SMTP address from a Recipient COM object.
        // Handles Exchange recipients via PropertyAccessor and GetExchangeUser fallback.
        static string ResolveRecipientSmtp(dynamic recip)
        {
            // Try PR_SMTP_ADDRESS (works for most recipients)
            try
            {
                string smtp = (string)recip.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001F");
                if (!string.IsNullOrEmpty(smtp)) return smtp;
            }
            catch { }
            // Fallback: Exchange user
            try
            {
                dynamic exchUser = recip.AddressEntry.GetExchangeUser();
                if (exchUser != null)
                {
                    string smtp = (string)exchUser.PrimarySmtpAddress;
                    try { Marshal.ReleaseComObject(exchUser); } catch { }
                    if (!string.IsNullOrEmpty(smtp)) return smtp;
                }
            }
            catch { }
            // Last resort: raw Address (may be X500 for Exchange)
            try { return (string)recip.Address; } catch { }
            return "";
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
                    if (found != null)
                    {
                        // Release child only if it is not the found folder itself
                        if (!object.ReferenceEquals(found, child))
                            try { Marshal.ReleaseComObject(child); } catch { }
                        return found;
                    }
                    try { Marshal.ReleaseComObject(child); } catch { }
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
                dynamic attachments = mail.Attachments;
                int attCount = (int)attachments.Count;
                for (int i = 1; i <= attCount; i++)
                {
                    dynamic att = null;
                    try
                    {
                        att = attachments[i];
                        string safeFn = SafeName((string)att.FileName);
                        string savePath = Path.Combine(mailDir, safeFn);
                        att.SaveAsFile(savePath);

                        // Auto-extract zip attachments
                        if (_autoUnzip && safeFn.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                        {
                            if (TryExtractZip(savePath, mailDir))
                            {
                                try { File.Delete(savePath); } catch { }
                                continue; // skip adding zip to attachment list
                            }
                        }
                        names.Add(safeFn);
                    }
                    catch { }
                    finally
                    {
                        if (att != null)
                            try { Marshal.ReleaseComObject(att); } catch { }
                    }
                }
                try { Marshal.ReleaseComObject(attachments); } catch { }
            }
            catch { }
            return names;
        }

        static bool TryExtractZip(string zipPath, string destDir)
        {
            string extractDir = Path.Combine(destDir,
                Path.GetFileNameWithoutExtension(zipPath));
            try
            {
                using (var archive = System.IO.Compression.ZipFile.OpenRead(zipPath))
                {
                    // Check for password protection
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

                    string extractDirFull = Path.GetFullPath(extractDir);
                    foreach (var entry in archive.Entries)
                    {
                        if (string.IsNullOrEmpty(entry.Name)) continue;
                        string entryDest = Path.GetFullPath(
                            Path.Combine(extractDir, entry.FullName));
                        // Zip slip guard: reject entries that escape the extract directory
                        if (!entryDest.StartsWith(extractDirFull + "\\") &&
                            !entryDest.Equals(extractDirFull))
                            continue;
                        string entryDir = Path.GetDirectoryName(entryDest);
                        Directory.CreateDirectory(entryDir);
                        try { entry.ExtractToFile(entryDest, true); }
                        catch { }
                    }
                }
                return true;
            }
            catch { return false; }
        }

        void WriteMetaJson(string path, dynamic mail, List<string> attNames, string smtp,
            string toRecipients = "", string ccRecipients = "")
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

            dynamic parentFolder = mail.Parent;
            string folderPathVal = (string)parentFolder.FolderPath;
            try { Marshal.ReleaseComObject(parentFolder); } catch { }

            var json = string.Format(
                "{{\n  \"entry_id\": \"{0}\",\n  \"mailbox_address\": \"{1}\",\n" +
                "  \"folder_path\": \"{2}\",\n  \"sender_name\": \"{3}\",\n" +
                "  \"sender_email\": \"{4}\",\n  \"subject\": \"{5}\",\n" +
                "  \"received_at\": \"{6:yyyy-MM-dd\\THH:mm:ss}\",\n" +
                "  \"to_recipients\": \"{7}\",\n  \"cc_recipients\": \"{8}\",\n" +
                "  \"body_path\": \"body.txt\",\n  \"msg_path\": \"mail.msg\",\n" +
                "  \"attachments\": {9}\n}}",
                JsonEsc((string)mail.EntryID), JsonEsc(smtp),
                JsonEsc(folderPathVal), JsonEsc((string)mail.SenderName),
                JsonEsc(senderAddr), JsonEsc((string)mail.Subject),
                (DateTime)mail.ReceivedTime,
                JsonEsc(toRecipients ?? ""), JsonEsc(ccRecipients ?? ""),
                attJson);

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
                        dynamic ds = acct.DeliveryStore;
                        bool match = (string)ds.StoreID == (string)store.StoreID;
                        try { Marshal.ReleaseComObject(ds); } catch { }
                        if (match)
                            return ((string)acct.SmtpAddress).ToLower();
                    }
                    catch { }
                    finally { try { Marshal.ReleaseComObject(acct); } catch { } }
                }
            }
            catch { }
            // Try PR_EMAIL_ADDRESS / PR_SMTP_ADDRESS via root folder
            dynamic rootF = null;
            try
            {
                rootF = store.GetRootFolder();
                try
                {
                    return ((string)rootF.PropertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x39FE001E")).ToLower();
                }
                catch { }
                try
                {
                    return ((string)rootF.PropertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x39FE001F")).ToLower();
                }
                catch { }
            }
            catch { }
            finally
            {
                if (rootF != null)
                    try { Marshal.ReleaseComObject(rootF); } catch { }
            }
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
