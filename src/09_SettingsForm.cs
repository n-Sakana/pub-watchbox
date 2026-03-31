using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows;

namespace WatchBox
{
    public class SettingsWindow : WebViewHost
    {
        MailScanner _cachedScanner;

        public SettingsWindow()
            : base("Settings", "settings.html", 460, 560)
        {
            MinWidth = 380; MinHeight = 480;
        }

        MailScanner GetScanner()
        {
            if (_cachedScanner == null) _cachedScanner = new MailScanner();
            return _cachedScanner;
        }

        protected override void HandleMessage(string action, string json)
        {
            switch (action)
            {
                case "getConfig": SendConfig(); break;
                case "saveConfig": SaveConfig(json); break;
                case "getAccounts": LoadAccounts(); break;
                case "getFolders": LoadFolders(json); break;
                case "browseFolder": BrowseFolder(json); break;
                case "importCsv": ImportCsv(); break;
                case "exportCsv": ExportCsv(json); break;
                case "close": Dispatcher.BeginInvoke(new Action(() => Close())); break;
            }
        }

        // --- Send full config to JS ---

        void SendConfig()
        {
            if (Config.ProfileCount == 0)
                Config.AddProfile("Default");

            var sb = new StringBuilder();
            sb.Append("{\"profiles\":[");
            for (int i = 0; i < Config.ProfileCount; i++)
            {
                if (i > 0) sb.Append(",");
                sb.Append("{");
                AppendProfileJson(sb, i);
                sb.Append("}");
            }
            sb.Append("]}");
            SendMessage("configLoaded", sb.ToString());
        }

        void AppendProfileJson(StringBuilder sb, int idx)
        {
            string[] keys = { "name", "type", "output_root", "account",
                "outlook_folder", "source_folder", "recurse", "auto_unzip",
                "since", "filter_mode", "filters", "flat_output", "short_dirname",
                "notify", "log_enabled", "manifest_hidden" };
            for (int k = 0; k < keys.Length; k++)
            {
                if (k > 0) sb.Append(",");
                sb.AppendFormat("\"{0}\":\"{1}\"",
                    keys[k], JsonEsc(Config.PGet(idx, keys[k])));
            }
        }

        // --- Save config from JS ---

        void SaveConfig(string json)
        {
            try
            {
                var profiles = ExtractJsonArray(json, "profiles");

                // Clear existing profiles
                while (Config.ProfileCount > 0)
                    Config.RemoveProfile(0);

                foreach (var pJson in profiles)
                {
                    int idx = Config.AddProfile(
                        ExtractJsonString(pJson, "name"));
                    string[] keys = { "type", "output_root", "account",
                        "outlook_folder", "source_folder", "recurse", "auto_unzip",
                        "since", "filter_mode", "filters", "flat_output",
                        "short_dirname", "notify", "log_enabled", "manifest_hidden" };
                    foreach (var k in keys)
                        Config.PSet(idx, k, ExtractJsonString(pJson, k));
                }
                Config.Save();
                SendMessage("saveResult", "{\"ok\":true}");
            }
            catch (Exception ex)
            {
                SendMessage("saveResult",
                    "{\"ok\":false,\"error\":\"" + JsonEsc(ex.Message) + "\"}");
            }
        }

        // --- Outlook accounts (STA thread) ---

        void LoadAccounts()
        {
            RunOnStaThread(() =>
            {
                var scanner = GetScanner();
                var accounts = scanner.GetAccounts();
                var sb = new StringBuilder();
                sb.Append("{\"accounts\":[");
                for (int i = 0; i < accounts.Count; i++)
                {
                    if (i > 0) sb.Append(",");
                    sb.AppendFormat("\"{0}\"", JsonEsc(accounts[i]));
                }
                sb.Append("]}");
                Dispatcher.BeginInvoke(new Action(() =>
                    SendMessage("accountsLoaded", sb.ToString())));
            });
        }

        void LoadFolders(string json)
        {
            string account = ExtractJsonString(json, "account");
            RunOnStaThread(() =>
            {
                var scanner = GetScanner();
                var folders = scanner.GetFolders(account);
                var sb = new StringBuilder();
                sb.Append("{\"folders\":[");
                for (int i = 0; i < folders.Count; i++)
                {
                    if (i > 0) sb.Append(",");
                    sb.AppendFormat("{{\"display\":\"{0}\",\"path\":\"{1}\"}}",
                        JsonEsc(folders[i][0]), JsonEsc(folders[i][1]));
                }
                sb.Append("]}");
                Dispatcher.BeginInvoke(new Action(() =>
                    SendMessage("foldersLoaded", sb.ToString())));
            });
        }

        // --- Browse folder (UI thread) ---

        void BrowseFolder(string json)
        {
            string field = ExtractJsonString(json, "field");
            string current = ExtractJsonString(json, "current");
            Dispatcher.BeginInvoke(new Action(() =>
            {
                string path = FolderPicker.Show(current);
                if (path != null)
                    SendMessage("folderSelected",
                        string.Format("{{\"field\":\"{0}\",\"path\":\"{1}\"}}",
                            JsonEsc(field), JsonEsc(path)));
            }));
        }

        // --- Import CSV ---

        void ImportCsv()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                var dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.Filter = "CSV files|*.csv|All files|*.*";
                if (dlg.ShowDialog() != true) return;

                try
                {
                    var lines = File.ReadAllLines(dlg.FileName, Encoding.UTF8);
                    var imported = new List<string>();
                    for (int row = 0; row < lines.Length; row++)
                    {
                        string line = lines[row].Trim();
                        if (string.IsNullOrEmpty(line)) continue;
                        var cols = line.Split(',');
                        if (cols[0].Trim().ToLower() == "name") continue;
                        if (cols.Length < 1) continue;

                        var sb = new StringBuilder();
                        sb.Append("{");
                        sb.AppendFormat("\"name\":\"{0}\"", JsonEsc(ColVal(cols, 0, "Imported")));
                        sb.AppendFormat(",\"type\":\"{0}\"", JsonEsc(ColVal(cols, 1, "mail")));
                        sb.AppendFormat(",\"output_root\":\"{0}\"", JsonEsc(ColVal(cols, 2, "")));
                        sb.AppendFormat(",\"account\":\"{0}\"", JsonEsc(ColVal(cols, 3, "")));
                        sb.AppendFormat(",\"outlook_folder\":\"{0}\"", JsonEsc(ColVal(cols, 4, "")));
                        sb.AppendFormat(",\"source_folder\":\"{0}\"", JsonEsc(ColVal(cols, 5, "")));
                        sb.AppendFormat(",\"manifest_hidden\":\"{0}\"", JsonEsc(ColVal(cols, 6, "1")));
                        sb.AppendFormat(",\"filters\":\"{0}\"", JsonEsc(ColVal(cols, 7, "")));
                        sb.AppendFormat(",\"filter_mode\":\"{0}\"", JsonEsc(ColVal(cols, 8, "or")));
                        sb.AppendFormat(",\"flat_output\":\"{0}\"", JsonEsc(ColVal(cols, 9, "0")));
                        sb.AppendFormat(",\"recurse\":\"{0}\"", JsonEsc(ColVal(cols, 10, "1")));
                        sb.AppendFormat(",\"since\":\"{0}\"", JsonEsc(ColVal(cols, 11, "")));
                        sb.AppendFormat(",\"short_dirname\":\"{0}\"", JsonEsc(ColVal(cols, 12, "0")));
                        sb.AppendFormat(",\"notify\":\"1\",\"log_enabled\":\"1\"");
                        sb.Append("}");
                        imported.Add(sb.ToString());
                    }

                    var result = new StringBuilder();
                    result.Append("{\"profiles\":[");
                    result.Append(string.Join(",", imported.ToArray()));
                    result.Append("]}");
                    SendMessage("importResult", result.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Import failed: " + ex.Message);
                }
            }));
        }

        // --- Export CSV ---

        void ExportCsv(string json)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                var dlg = new Microsoft.Win32.SaveFileDialog();
                dlg.Filter = "CSV files|*.csv|All files|*.*";
                dlg.FileName = "profiles.csv";
                if (dlg.ShowDialog() != true) return;

                try
                {
                    var profiles = ExtractJsonArray(json, "profiles");
                    var lines = new List<string>();
                    lines.Add("name,type,output_root,account,outlook_folder," +
                        "source_folder,manifest_hidden,filters,filter_mode," +
                        "flat_output,recurse,since,short_dirname");
                    foreach (var pJson in profiles)
                    {
                        lines.Add(string.Join(",", new[] {
                            ExtractJsonString(pJson, "name"),
                            ExtractJsonString(pJson, "type"),
                            ExtractJsonString(pJson, "output_root"),
                            ExtractJsonString(pJson, "account"),
                            ExtractJsonString(pJson, "outlook_folder"),
                            ExtractJsonString(pJson, "source_folder"),
                            ExtractJsonString(pJson, "manifest_hidden"),
                            ExtractJsonString(pJson, "filters"),
                            ExtractJsonString(pJson, "filter_mode"),
                            ExtractJsonString(pJson, "flat_output"),
                            ExtractJsonString(pJson, "recurse"),
                            ExtractJsonString(pJson, "since"),
                            ExtractJsonString(pJson, "short_dirname")
                        }));
                    }
                    File.WriteAllLines(dlg.FileName, lines.ToArray(), Encoding.UTF8);
                    MessageBox.Show(profiles.Count + " profile(s) exported.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Export failed: " + ex.Message);
                }
            }));
        }

        // --- Helpers ---

        static void RunOnStaThread(Action work)
        {
            var thread = new Thread(() => { try { work(); } catch { } });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }

        // Extract JSON array as list of raw JSON object strings
        static List<string> ExtractJsonArray(string json, string key)
        {
            var list = new List<string>();
            int idx = json.IndexOf("\"" + key + "\"");
            if (idx < 0) return list;
            idx = json.IndexOf('[', idx);
            if (idx < 0) return list;
            idx++; // skip [

            while (idx < json.Length)
            {
                while (idx < json.Length && json[idx] != '{' && json[idx] != ']') idx++;
                if (idx >= json.Length || json[idx] == ']') break;

                int start = idx; int depth = 1; idx++;
                bool inStr = false;
                while (idx < json.Length && depth > 0)
                {
                    char c = json[idx];
                    if (c == '\\' && inStr) { idx += 2; continue; }
                    if (c == '"') inStr = !inStr;
                    else if (!inStr && c == '{') depth++;
                    else if (!inStr && c == '}') depth--;
                    idx++;
                }
                list.Add(json.Substring(start, idx - start));
            }
            return list;
        }

        static string ColVal(string[] cols, int idx, string def)
        {
            if (idx >= cols.Length) return def;
            string v = cols[idx].Trim();
            return v.Length > 0 ? v : def;
        }
    }
}
