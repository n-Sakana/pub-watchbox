using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace WatchBox
{
    public static class App
    {
        public static void Run(string appDir)
        {
            Config.Load(Path.Combine(appDir, "config.json"));
            Config.Set("app_dir", appDir);

            // Resolve WebView2 DLLs from lib/ at runtime
            string libDir = Path.Combine(appDir, "lib");
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                string name = new System.Reflection.AssemblyName(args.Name).Name;
                string dllPath = Path.Combine(libDir, name + ".dll");
                if (File.Exists(dllPath))
                    return System.Reflection.Assembly.LoadFrom(dllPath);
                return null;
            };

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
        static readonly string[] ProfileKeys = {
            // Common: mail or folder
            "name", "type", "output_root", "notify", "log_enabled",
            // Mail-specific
            "account", "outlook_folder", "since", "filter_mode", "filters", "flat_output",
            // Folder-specific (source_folder optional: set = copy mode, empty = manifest-only)
            "source_folder", "recurse",
            // Manifest visibility
            "manifest_hidden",
            // Mail directory naming
            "short_dirname",
            // Folder: auto-extract zip files
            "auto_unzip",
            // Internal: last successful scan timestamp (managed by ProfileRunner)
            "last_scan"
        };

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
            PSet(idx, "type", "mail");
            PSet(idx, "output_root", "");
            PSet(idx, "notify", "1");
            PSet(idx, "log_enabled", "1");
            PSet(idx, "account", "");
            PSet(idx, "outlook_folder", "");
            PSet(idx, "since", "");
            PSet(idx, "filter_mode", "or");
            PSet(idx, "filters", "");
            PSet(idx, "flat_output", "0");
            PSet(idx, "source_folder", "");
            PSet(idx, "recurse", "1");
            PSet(idx, "manifest_hidden", "1");
            PSet(idx, "short_dirname", "0");
            PSet(idx, "auto_unzip", "0");
            return idx;
        }

        public static void RemoveProfile(int idx)
        {
            int count = ProfileCount;
            if (idx < 0 || idx >= count) return;
            for (int i = idx; i < count - 1; i++)
                foreach (var k in ProfileKeys)
                    PSet(i, k, PGet(i + 1, k));
            for (int i = 0; i < ProfileKeys.Length; i++)
                Remove(string.Format("p{0}_{1}", count - 1, ProfileKeys[i]));
            Set("profile_count", (count - 1).ToString());
        }

    }
}
