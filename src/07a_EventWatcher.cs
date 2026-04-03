using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace WatchBox
{
    // Event-driven watcher:
    // - Folder: FileSystemWatcher
    // - Mail: Outlook NewMailEx event
    public class EventWatcher
    {
        List<FileSystemWatcher> _fsWatchers = new List<FileSystemWatcher>();
        Thread _olThread;
        System.Windows.Threading.Dispatcher _olDispatcher;
        Action<int, string> _onEvent;
        int _debounceMs = 2000;
        Dictionary<int, DateTime> _lastEvent = new Dictionary<int, DateTime>();

        public void Start(Action<int, string> onEvent)
        {
            _onEvent = onEvent;
            Stop();

            bool hasMail = false;
            for (int i = 0; i < Config.ProfileCount; i++)
            {
                string type = Config.PGet(i, "type", "mail");
                if (type == "folder")
                    StartFolderWatcher(i);
                else
                    hasMail = true;
            }

            if (hasMail)
                StartMailWatcher();
        }

        public void Stop()
        {
            foreach (var w in _fsWatchers)
            {
                try { w.EnableRaisingEvents = false; w.Dispose(); } catch { }
            }
            _fsWatchers.Clear();

            if (_olDispatcher != null)
            {
                try { _olDispatcher.InvokeShutdown(); } catch { }
                _olDispatcher = null;
            }

            foreach (var items in _watchedItems)
            {
                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(items); } catch { }
            }
            _watchedItems.Clear();
        }

        // --- Folder: FileSystemWatcher ---

        void StartFolderWatcher(int profileIndex)
        {
            string sourceFolder = Config.PGet(profileIndex, "source_folder");
            string outputRoot = Config.PGet(profileIndex, "output_root");
            string watchPath = !string.IsNullOrEmpty(sourceFolder) ? sourceFolder : outputRoot;
            if (string.IsNullOrEmpty(watchPath) || !Directory.Exists(watchPath)) return;

            bool recurse = Config.PGet(profileIndex, "recurse", "1") != "0";

            var watcher = new FileSystemWatcher();
            watcher.Path = watchPath;
            watcher.IncludeSubdirectories = recurse;
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName
                | NotifyFilters.LastWrite | NotifyFilters.Size;

            int idx = profileIndex;
            FileSystemEventHandler handler = (s, e) => OnFolderEvent(idx, e.FullPath);
            RenamedEventHandler renameHandler = (s, e) => OnFolderEvent(idx, e.FullPath);

            watcher.Created += handler;
            watcher.Changed += handler;
            watcher.Deleted += handler;
            watcher.Renamed += renameHandler;
            watcher.EnableRaisingEvents = true;
            _fsWatchers.Add(watcher);
        }

        void OnFolderEvent(int profileIndex, string path)
        {
            string fn = Path.GetFileName(path);
            if (fn == ".manifest.csv" || fn == "manifest.csv" || fn == "log.csv") return;

            DateTime now = DateTime.Now;
            DateTime last;
            if (_lastEvent.TryGetValue(profileIndex, out last))
                if ((now - last).TotalMilliseconds < _debounceMs) return;
            _lastEvent[profileIndex] = now;

            if (_onEvent != null)
                _onEvent(profileIndex, fn);
        }

        // --- Mail: Outlook NewMailEx event on STA thread ---

        void StartMailWatcher()
        {
            if (_olThread != null && _olThread.IsAlive) return;

            _olThread = new Thread(OutlookEventLoop);
            _olThread.SetApartmentState(ApartmentState.STA);
            _olThread.IsBackground = true;
            _olThread.Start();
        }

        void OutlookEventLoop()
        {
            dynamic olApp = null;
            dynamic olNs = null;
            try
            {
                try { olApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application"); }
                catch { olApp = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")); }
                olNs = olApp.GetNamespace("MAPI");

                // Hook ItemAdd on each watched folder's Items collection
                // This works reliably with dynamic/late-binding unlike NewMailEx
                var watchedItems = new List<dynamic>();
                foreach (dynamic store in olNs.Stores)
                {
                    try
                    {
                        HookFolderTree(store.GetRootFolder(), watchedItems);
                    }
                    catch { }
                }

                _olDispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher;
                System.Windows.Threading.Dispatcher.Run();
            }
            catch { }
            finally
            {
                try { if (olNs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(olNs); } catch { }
                try { if (olApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(olApp); } catch { }
            }
        }

        // We need to keep references to Items collections alive (prevent GC)
        List<dynamic> _watchedItems = new List<dynamic>();

        void HookFolderTree(dynamic folder, List<dynamic> watchedItems)
        {
            try
            {
                dynamic items = folder.Items;
                items.ItemAdd += new Action<dynamic>(OnItemAdd);
                _watchedItems.Add(items); // prevent GC
            }
            catch { }

            // Only hook top-level folders (Inbox, Sent, etc.) to avoid too many hooks
            // Deep recursion on all folders is expensive
            try
            {
                foreach (dynamic child in folder.Folders)
                {
                    try
                    {
                        dynamic childItems = child.Items;
                        childItems.ItemAdd += new Action<dynamic>(OnItemAdd);
                        _watchedItems.Add(childItems);
                    }
                    catch { }
                }
            }
            catch { }
        }

        void OnItemAdd(dynamic item)
        {
            try { if ((int)item.Class != 43) return; } catch { return; }
            try
            {
                string subject = "";
                try { subject = (string)item.Subject; } catch { }

                for (int i = 0; i < Config.ProfileCount; i++)
                {
                    if (Config.PGet(i, "type", "mail") != "mail") continue;

                    DateTime now = DateTime.Now;
                    DateTime last;
                    if (_lastEvent.TryGetValue(i, out last))
                        if ((now - last).TotalMilliseconds < _debounceMs) continue;
                    _lastEvent[i] = now;

                    if (_onEvent != null)
                        _onEvent(i, subject);
                }
            }
            catch { }
        }
    }
}
