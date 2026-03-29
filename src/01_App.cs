using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace MailPull
{
    public static class App
    {
        public static void Run(string appDir)
        {
            Config.Load(Path.Combine(appDir, "config.json"));
            var app = new System.Windows.Application();
            app.ShutdownMode = System.Windows.ShutdownMode.OnMainWindowClose;
            app.Run(new MonitorWindow());
        }
    }

    public static class Config
    {
        static Dictionary<string, string> _data = new Dictionary<string, string>(
            StringComparer.OrdinalIgnoreCase);
        static string _path;

        public static void Load(string path)
        {
            _path = path;
            _data.Clear();
            if (!File.Exists(path)) return;
            foreach (var line in File.ReadAllLines(path))
            {
                var m = Regex.Match(line, @"""([\w]+)""\s*:\s*""((?:[^""\\]|\\.)*)""");
                if (m.Success)
                    _data[m.Groups[1].Value] = m.Groups[2].Value
                        .Replace("\\\\", "\x01").Replace("\\n", "\n").Replace("\x01", "\\");
            }
            Migrate();
        }

        public static void Save()
        {
            if (_path == null) return;
            var lines = new List<string> { "{" };
            var keys = new List<string>(_data.Keys);
            for (int i = 0; i < keys.Count; i++)
            {
                var comma = i < keys.Count - 1 ? "," : "";
                var val = _data[keys[i]].Replace("\\", "\\\\").Replace("\n", "\\n")
                    .Replace("\"", "\\\"");
                lines.Add(string.Format("  \"{0}\": \"{1}\"{2}", keys[i], val, comma));
            }
            lines.Add("}");
            File.WriteAllLines(_path, lines.ToArray(), System.Text.Encoding.UTF8);
        }

        public static string Get(string key, string def = "")
        {
            string v;
            return _data.TryGetValue(key, out v) ? v : def;
        }

        public static void Set(string key, string value) { _data[key] = value; }

        public static void Remove(string key) { _data.Remove(key); }

        // --- Profile support ---
        // Keys: p0_name, p0_account, p0_folder_path, p0_export_root, p0_export_since, p0_poll_seconds
        static readonly string[] ProfileKeys = { "name", "account", "folder_path", "export_root", "export_since", "poll_seconds" };

        public static int ProfileCount
        {
            get
            {
                int n;
                return int.TryParse(Get("profile_count", "0"), out n) ? n : 0;
            }
        }

        public static string PGet(int idx, string key, string def = "")
        {
            return Get(string.Format("p{0}_{1}", idx, key), def);
        }

        public static void PSet(int idx, string key, string value)
        {
            Set(string.Format("p{0}_{1}", idx, key), value);
        }

        public static int AddProfile(string name)
        {
            int idx = ProfileCount;
            Set("profile_count", (idx + 1).ToString());
            PSet(idx, "name", name);
            PSet(idx, "poll_seconds", "60");
            return idx;
        }

        public static void RemoveProfile(int idx)
        {
            int count = ProfileCount;
            if (idx < 0 || idx >= count) return;
            // Shift subsequent profiles down
            for (int i = idx; i < count - 1; i++)
                foreach (var k in ProfileKeys)
                    PSet(i, k, PGet(i + 1, k));
            // Remove last profile's keys
            for (int i = 0; i < ProfileKeys.Length; i++)
                Remove(string.Format("p{0}_{1}", count - 1, ProfileKeys[i]));
            Set("profile_count", (count - 1).ToString());
        }

        // Migrate old single-profile config to profile 0
        static void Migrate()
        {
            if (ProfileCount > 0) return;
            string root = Get("export_root");
            if (string.IsNullOrEmpty(root)) return;

            Set("profile_count", "1");
            PSet(0, "name", "Default");
            PSet(0, "account", Get("account"));
            PSet(0, "folder_path", Get("folder_path"));
            PSet(0, "export_root", root);
            PSet(0, "export_since", Get("export_since"));
            PSet(0, "poll_seconds", Get("poll_seconds", "60"));

            // Clean old keys
            Remove("account"); Remove("folder_path"); Remove("export_root");
            Remove("export_since"); Remove("export_days"); Remove("poll_seconds");
            Remove("poll_days");
        }
    }
}
