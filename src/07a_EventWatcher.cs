using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace WatchBox
{
    // Event-driven watcher:
    // - Folder: FileSystemWatcher
    // - Mail: Outlook ItemAdd on profile-configured folders only
    public class EventWatcher
    {
        List<FileSystemWatcher> _fsWatchers = new List<FileSystemWatcher>();
        Thread _olThread;
        System.Windows.Threading.Dispatcher _olDispatcher;
        Action<int, string> _onEvent;
        int _debounceMs = 2000;
        ConcurrentDictionary<int, DateTime> _lastEvent = new ConcurrentDictionary<int, DateTime>();
        readonly object _watchLock = new object();

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

            // Wait for the STA thread to finish before releasing COM objects
            if (_olThread != null && _olThread.IsAlive)
            {
                try { _olThread.Join(3000); } catch { }
            }
            _olThread = null;

            lock (_watchLock)
            {
                foreach (var items in _watchedItems)
                {
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(items); } catch { }
                }
                _watchedItems.Clear();
            }
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

        // --- Mail: Outlook ItemAdd on STA thread ---
        // Only hooks folders configured in mail profiles to minimize COM overhead.

        void StartMailWatcher()
        {
            if (_olThread != null && _olThread.IsAlive) return;

            _olThread = new Thread(OutlookEventLoop);
            _olThread.SetApartmentState(ApartmentState.STA);
            _olThread.IsBackground = true;
            _olThread.Start();
        }

        // Collect target folder paths from mail profiles
        List<MailWatchTarget> GetMailWatchTargets()
        {
            var targets = new List<MailWatchTarget>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < Config.ProfileCount; i++)
            {
                if (Config.PGet(i, "type", "mail") != "mail") continue;
                string account = Config.PGet(i, "account", "");
                string folder = Config.PGet(i, "outlook_folder", "");
                string key = account + "|" + folder;
                if (!seen.Add(key)) continue;
                targets.Add(new MailWatchTarget { Account = account, FolderPath = folder });
            }
            return targets;
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

                var targets = GetMailWatchTargets();
                foreach (var target in targets)
                {
                    try { HookTargetFolder(olNs, target); }
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

        void HookTargetFolder(dynamic olNs, MailWatchTarget target)
        {
            // Find the store that matches the account
            dynamic rootFolder = null;
            if (!string.IsNullOrEmpty(target.Account))
            {
                foreach (dynamic acct in olNs.Accounts)
                {
                    try
                    {
                        if (string.Equals((string)acct.SmtpAddress, target.Account,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            dynamic ds = acct.DeliveryStore;
                            rootFolder = ds.GetRootFolder();
                            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(ds); } catch { }
                            break;
                        }
                    }
                    catch { }
                    finally
                    {
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(acct); } catch { }
                    }
                }
                // Fallback: shared/delegate mailbox via store display name
                if (rootFolder == null)
                {
                    foreach (dynamic store in olNs.Stores)
                    {
                        try
                        {
                            string name = "";
                            try { name = (string)store.DisplayName; } catch { }
                            if (string.Equals(name, target.Account, StringComparison.OrdinalIgnoreCase))
                            {
                                rootFolder = store.GetRootFolder();
                                break;
                            }
                        }
                        catch { }
                        finally
                        {
                            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(store); } catch { }
                        }
                    }
                }
            }
            else
            {
                // No account specified: use default store
                try
                {
                    dynamic defStore = olNs.DefaultStore;
                    rootFolder = defStore.GetRootFolder();
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(defStore); } catch { }
                }
                catch { }
            }

            if (rootFolder == null) return;

            // Navigate to the specific folder path, or hook root + children if no path
            if (!string.IsNullOrEmpty(target.FolderPath))
            {
                dynamic folder = NavigateToFolder(rootFolder, target.FolderPath);
                if (folder != null)
                    HookSingleFolder(folder, true);
                // rootFolder is an intermediate — release if we navigated deeper
                if (folder != null && !object.ReferenceEquals(folder, rootFolder))
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(rootFolder); } catch { }
            }
            else
            {
                // No specific folder: hook root's immediate children only
                HookSingleFolder(rootFolder, true);
            }
        }

        dynamic NavigateToFolder(dynamic rootFolder, string folderPath)
        {
            string[] parts = folderPath.Split(new[] { '\\', '/' },
                StringSplitOptions.RemoveEmptyEntries);
            dynamic current = rootFolder;
            foreach (string part in parts)
            {
                bool found = false;
                dynamic prev = current;
                try
                {
                    foreach (dynamic child in current.Folders)
                    {
                        try
                        {
                            if (string.Equals((string)child.Name, part,
                                StringComparison.OrdinalIgnoreCase))
                            {
                                current = child;
                                found = true;
                            }
                            else
                            {
                                // Release folders we don't need
                                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(child); } catch { }
                            }
                        }
                        catch { }
                        if (found) break;
                    }
                }
                catch { }
                // Release intermediate folder (not root, not the final target)
                if (found && !object.ReferenceEquals(prev, rootFolder))
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(prev); } catch { }
                if (!found) return null;
            }
            return current;
        }

        void HookSingleFolder(dynamic folder, bool includeChildren)
        {
            try
            {
                dynamic items = folder.Items;
                items.ItemAdd += new Action<dynamic>(OnItemAdd);
                lock (_watchLock) { _watchedItems.Add(items); }
            }
            catch { }

            if (!includeChildren) return;
            try
            {
                foreach (dynamic child in folder.Folders)
                {
                    try
                    {
                        dynamic childItems = child.Items;
                        childItems.ItemAdd += new Action<dynamic>(OnItemAdd);
                        lock (_watchLock) { _watchedItems.Add(childItems); }
                        // Release folder object; Items ref keeps the hook alive
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(child); } catch { }
                    }
                    catch { }
                }
            }
            catch { }
        }

        void OnItemAdd(dynamic item)
        {
            try
            {
                try { if ((int)item.Class != 43) return; } catch { return; }

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
            finally
            {
                // Release the COM object passed by the event
                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(item); } catch { }
            }
        }

        class MailWatchTarget
        {
            public string Account;
            public string FolderPath;
        }
    }
}
