using System;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Threading;

namespace WatchBox
{
    public class MonitorWindow : Window
    {
        TextBlock _status;
        Button _btnWatch, _btnPull;
        EventWatcher _eventWatcher;
        DispatcherTimer _clockTimer;
        bool _watching;
        volatile bool _cancelRequested;
        volatile bool _pulling;
        volatile SourceScanner _activeScanner;

        static readonly Brush AccentBrush = new SolidColorBrush(Color.FromRgb(55, 120, 200));
        static readonly Brush AccentHover = new SolidColorBrush(Color.FromRgb(45, 100, 180));

        public MonitorWindow()
        {
            Title = "watchbox";
            Width = 380; Height = 240;
            ResizeMode = ResizeMode.NoResize;
            Background = Brushes.White;
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = SystemParameters.WorkArea.Right - 400;
            Top = SystemParameters.WorkArea.Top + 40;
            FontFamily = new FontFamily("Segoe UI");
            FontSize = 13;

            var dock = new DockPanel { LastChildFill = true };

            var statusBorder = new Border
            {
                BorderThickness = new Thickness(0, 1, 0, 0),
                BorderBrush = new SolidColorBrush(Color.FromRgb(230, 230, 230)),
                Background = new SolidColorBrush(Color.FromRgb(248, 248, 248)),
                Padding = new Thickness(16, 8, 16, 8)
            };
            _status = new TextBlock
            {
                Text = "Ready",
                Foreground = new SolidColorBrush(Color.FromRgb(128, 128, 128)),
                FontSize = 12
            };
            statusBorder.Child = _status;
            DockPanel.SetDock(statusBorder, Dock.Bottom);
            dock.Children.Add(statusBorder);

            var grid = new UniformGrid { Rows = 2, Columns = 2, Margin = new Thickness(16) };
            var btnSettings = MkBtn("\uE713", "Settings", false);
            var btnViewer   = MkBtn("\uE8A1", "Viewer", false);
            _btnPull        = MkBtn("\uE74B", "Pull", false);
            _btnWatch       = MkBtn("\uE768", "Watch", true);

            grid.Children.Add(btnSettings);
            grid.Children.Add(btnViewer);
            grid.Children.Add(_btnPull);
            grid.Children.Add(_btnWatch);

            dock.Children.Add(grid);
            Content = dock;

            btnSettings.Click += (s, e) => {
                var f = new SettingsWindow();
                f.Owner = this;
                f.ShowDialog();
                if (_watching)
                {
                    StopWatching();
                    StartWatching();
                }
            };
            btnViewer.Click += (s, e) => { var w = new SearchWindow(); w.Owner = this; w.Show(); };

            // Pre-initialize WebView2 environment in background for fast window open
            Loaded += async (s, e) => { try { await WebViewHost.WarmUpAsync(); } catch { } };
            _btnPull.Click += OnPullClick;
            _btnWatch.Click += OnWatchClick;
        }

        Button MkBtn(string icon, string text, bool accent)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(4) };
            sp.Children.Add(new TextBlock
            {
                Text = icon,
                FontFamily = new FontFamily("Segoe MDL2 Assets"),
                FontSize = 18,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            });
            sp.Children.Add(new TextBlock
            {
                Text = text, FontSize = 14,
                VerticalAlignment = VerticalAlignment.Center
            });

            var btn = new Button
            {
                Content = sp, Margin = new Thickness(4),
                Padding = new Thickness(16, 10, 16, 10),
                Cursor = System.Windows.Input.Cursors.Hand
            };

            if (accent)
            {
                btn.Background = AccentBrush;
                btn.Foreground = Brushes.White;
                btn.BorderBrush = AccentBrush;
            }

            return btn;
        }

        // --- Pull: run all profiles once (full sync) ---

        void OnPullClick(object sender, RoutedEventArgs e)
        {
            if (_watching) return;

            // If already pulling, this is a Cancel click
            if (_pulling)
            {
                _cancelRequested = true;
                if (_activeScanner != null) _activeScanner.CancelRequested = true;
                _status.Text = "Cancelling...";
                return;
            }

            if (Config.ProfileCount == 0)
            { MessageBox.Show("No profiles configured.", "watchbox"); return; }

            _cancelRequested = false;
            _pulling = true;
            SetPullButton(true);

            RunOnStaThread(() =>
            {
                int totalAdded = 0;
                int profileCount = Config.ProfileCount;
                var mailScanner = new MailScanner();
                _activeScanner = mailScanner;

                // Collect unique (account, folder) keys and their broadest date filter
                var scanKeys = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var mailProfiles = new List<int>();
                var folderProfiles = new List<int>();

                for (int i = 0; i < profileCount; i++)
                {
                    if (string.IsNullOrEmpty(Config.PGet(i, "output_root"))) continue;
                    if (Config.PGet(i, "type", "mail") == "folder")
                    { folderProfiles.Add(i); continue; }
                    mailProfiles.Add(i);

                    string key = Config.PGet(i, "account").ToLower() + "\t" +
                        Config.PGet(i, "outlook_folder");
                    if (!scanKeys.ContainsKey(key)) scanKeys[key] = null;
                }

                // For each unique key, compute broadest date filter
                foreach (var key in new List<string>(scanKeys.Keys))
                {
                    DateTime earliest = DateTime.MaxValue;
                    bool needAll = false;
                    foreach (int i in mailProfiles)
                    {
                        string k = Config.PGet(i, "account").ToLower() + "\t" +
                            Config.PGet(i, "outlook_folder");
                        if (!string.Equals(k, key, StringComparison.OrdinalIgnoreCase)) continue;

                        DateTime dt;
                        string since = Config.PGet(i, "since");
                        string lastScan = Config.PGet(i, "last_scan");
                        string[] fmts = { "yyyy-MM-dd", "yyyy/MM/dd", "yyyy/M/d", "M/d/yyyy" };
                        if (!string.IsNullOrEmpty(since) && DateTime.TryParseExact(since, fmts,
                            CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                        { if (dt < earliest) earliest = dt; }
                        else if (!string.IsNullOrEmpty(lastScan) && DateTime.TryParseExact(
                            lastScan, "yyyy-MM-dd", CultureInfo.InvariantCulture,
                            DateTimeStyles.None, out dt) &&
                            ManifestIO.LoadIds(Config.PGet(i, "output_root")).Count > 0)
                        { dt = dt.AddDays(-1); if (dt < earliest) earliest = dt; }
                        else needAll = true;
                    }
                    scanKeys[key] = needAll || earliest == DateTime.MaxValue ? null
                        : string.Format("[ReceivedTime]>='{0:yyyy/MM/dd} 00:00'", earliest);
                }

                // Phase 1: Scan each unique folder once, cache items
                var caches = new Dictionary<string, List<CachedMailItem>>(
                    StringComparer.OrdinalIgnoreCase);
                int scanIdx = 0;
                foreach (var kv in scanKeys)
                {
                    if (_cancelRequested) break;
                    scanIdx++;
                    var parts = kv.Key.Split('\t');
                    string label = parts.Length > 1 && parts[1].Length > 0
                        ? parts[1].Substring(parts[1].LastIndexOf('\\') + 1) : "All";
                    Dispatcher.BeginInvoke(new Action(() =>
                        _status.Text = string.Format("Scanning {0} ({1}/{2})...",
                            label, scanIdx, scanKeys.Count)));

                    var cfg = new Dictionary<string, string>();
                    cfg["account"] = parts[0];
                    cfg["outlook_folder"] = parts.Length > 1 ? parts[1] : "";
                    caches[kv.Key] = mailScanner.ScanBulk(cfg, kv.Value);
                }

                // Phase 2: Per-profile keyword matching and export from cache
                for (int p = 0; p < mailProfiles.Count; p++)
                {
                    if (_cancelRequested) break;
                    int i = mailProfiles[p];
                    string pname = Config.PGet(i, "name", "Profile " + i);
                    int idx = i;
                    Dispatcher.BeginInvoke(new Action(() =>
                        _status.Text = string.Format("({0}/{1}) [{2}]...",
                            idx + 1, profileCount, pname)));

                    string key = Config.PGet(i, "account").ToLower() + "\t" +
                        Config.PGet(i, "outlook_folder");
                    List<CachedMailItem> cache;
                    if (!caches.TryGetValue(key, out cache)) continue;

                    Action<int, string> progress = (c, s) =>
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            if (s != null && s.Length > 36) s = s.Substring(0, 36);
                            _status.Text = string.Format("({0}/{1}) [{2}] {3}: {4}",
                                idx + 1, profileCount, pname, c, s);
                        }));
                    totalAdded += ProfileRunner.RunFromCache(
                        i, mailScanner, cache, progress).Added;
                }

                // Phase 3: Folder-type profiles (individual scan)
                foreach (int i in folderProfiles)
                {
                    if (_cancelRequested) break;
                    string pname = Config.PGet(i, "name", "Profile " + i);
                    int idx = i;
                    Dispatcher.BeginInvoke(new Action(() =>
                        _status.Text = string.Format("({0}/{1}) {2}...",
                            idx + 1, profileCount, pname)));
                    var fs = new FolderScanner();
                    _activeScanner = fs;
                    Action<int, string> progress = (c, s) =>
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            if (s != null && s.Length > 36) s = s.Substring(0, 36);
                            _status.Text = string.Format("({0}/{1}) [{2}] {3}: {4}",
                                idx + 1, profileCount, pname, c, s);
                        }));
                    totalAdded += ProfileRunner.Run(i, fs, progress).Added;
                }
                _activeScanner = null;
                mailScanner.Cleanup();
                return totalAdded;
            },
            total =>
            {
                bool cancelled = _cancelRequested;
                _cancelRequested = false;
                _pulling = false;
                SetPullButton(false);
                _status.Text = cancelled
                    ? string.Format("Cancelled ({0} new)", total)
                    : string.Format("Done: {0} new", total);
                if (total > 0 && !cancelled)
                    ToastPopup.Show("watchbox", string.Format("{0} new item(s) found", total));
            });
        }

        void SetPullButton(bool running)
        {
            var sp = (StackPanel)_btnPull.Content;
            var icon = (TextBlock)sp.Children[0];
            var label = (TextBlock)sp.Children[1];
            if (running) { icon.Text = "\uE711"; label.Text = "Cancel"; }
            else { icon.Text = "\uE74B"; label.Text = "Pull"; }
        }

        // --- Watch: event-driven monitoring ---

        void OnWatchClick(object sender, RoutedEventArgs e)
        {
            if (_watching)
                StopWatching();
            else
            {
                if (Config.ProfileCount == 0)
                { MessageBox.Show("No profiles configured.", "watchbox"); return; }
                StartWatching();
            }
        }

        void StartWatching()
        {
            _watching = true;
            _btnPull.IsEnabled = false;
            _btnWatch.IsEnabled = false;
            _status.Text = "Initial sync...";

            // Initial full sync (same bulk strategy as Pull), then start event watch
            RunOnStaThread(() =>
            {
                int total = 0;
                int profileCount = Config.ProfileCount;
                var mailScanner = new MailScanner();

                var scanKeys = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var mailProfiles = new List<int>();
                var folderProfiles = new List<int>();

                for (int i = 0; i < profileCount; i++)
                {
                    if (string.IsNullOrEmpty(Config.PGet(i, "output_root"))) continue;
                    if (Config.PGet(i, "type", "mail") == "folder")
                    { folderProfiles.Add(i); continue; }
                    mailProfiles.Add(i);

                    string key = Config.PGet(i, "account").ToLower() + "\t" +
                        Config.PGet(i, "outlook_folder");
                    if (!scanKeys.ContainsKey(key)) scanKeys[key] = null;
                }

                foreach (var key in new List<string>(scanKeys.Keys))
                {
                    DateTime earliest = DateTime.MaxValue;
                    bool needAll = false;
                    foreach (int i in mailProfiles)
                    {
                        string k = Config.PGet(i, "account").ToLower() + "\t" +
                            Config.PGet(i, "outlook_folder");
                        if (!string.Equals(k, key, StringComparison.OrdinalIgnoreCase)) continue;

                        DateTime dt;
                        string since = Config.PGet(i, "since");
                        string lastScan = Config.PGet(i, "last_scan");
                        string[] fmts = { "yyyy-MM-dd", "yyyy/MM/dd", "yyyy/M/d", "M/d/yyyy" };
                        if (!string.IsNullOrEmpty(since) && DateTime.TryParseExact(since, fmts,
                            CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                        { if (dt < earliest) earliest = dt; }
                        else if (!string.IsNullOrEmpty(lastScan) && DateTime.TryParseExact(
                            lastScan, "yyyy-MM-dd", CultureInfo.InvariantCulture,
                            DateTimeStyles.None, out dt) &&
                            ManifestIO.LoadIds(Config.PGet(i, "output_root")).Count > 0)
                        { dt = dt.AddDays(-1); if (dt < earliest) earliest = dt; }
                        else needAll = true;
                    }
                    scanKeys[key] = needAll || earliest == DateTime.MaxValue ? null
                        : string.Format("[ReceivedTime]>='{0:yyyy/MM/dd} 00:00'", earliest);
                }

                // Phase 1: Scan each unique folder once, cache items
                var caches = new Dictionary<string, List<CachedMailItem>>(
                    StringComparer.OrdinalIgnoreCase);
                int scanIdx = 0;
                foreach (var kv in scanKeys)
                {
                    scanIdx++;
                    var parts = kv.Key.Split('\t');
                    string label = parts.Length > 1 && parts[1].Length > 0
                        ? parts[1].Substring(parts[1].LastIndexOf('\\') + 1) : "All";
                    Dispatcher.BeginInvoke(new Action(() =>
                        _status.Text = string.Format("Initial sync: scanning {0} ({1}/{2})...",
                            label, scanIdx, scanKeys.Count)));

                    var cfg = new Dictionary<string, string>();
                    cfg["account"] = parts[0];
                    cfg["outlook_folder"] = parts.Length > 1 ? parts[1] : "";
                    caches[kv.Key] = mailScanner.ScanBulk(cfg, kv.Value);
                }

                // Phase 2: Per-profile keyword matching and export from cache
                for (int p = 0; p < mailProfiles.Count; p++)
                {
                    int i = mailProfiles[p];
                    string pname = Config.PGet(i, "name", "Profile " + i);
                    int idx = i;
                    Dispatcher.BeginInvoke(new Action(() =>
                        _status.Text = string.Format("Initial sync: ({0}/{1}) [{2}]...",
                            idx + 1, profileCount, pname)));

                    string key = Config.PGet(i, "account").ToLower() + "\t" +
                        Config.PGet(i, "outlook_folder");
                    List<CachedMailItem> cache;
                    if (!caches.TryGetValue(key, out cache)) continue;

                    total += ProfileRunner.RunFromCache(i, mailScanner, cache, null).Added;
                }

                // Phase 3: Folder-type profiles
                foreach (int i in folderProfiles)
                {
                    string pname = Config.PGet(i, "name", "Profile " + i);
                    int idx = i;
                    Dispatcher.BeginInvoke(new Action(() =>
                        _status.Text = string.Format("Initial sync: ({0}/{1}) {2}...",
                            idx + 1, profileCount, pname)));
                    total += ProfileRunner.Run(i, new FolderScanner(), null).Added;
                }

                mailScanner.Cleanup();
                return total;
            },
            total =>
            {
                _eventWatcher = new EventWatcher();
                _eventWatcher.Start(OnWatchEvent);

                _clockTimer = new DispatcherTimer();
                _clockTimer.Interval = TimeSpan.FromSeconds(1);
                _clockTimer.Tick += (s2, e2) =>
                {
                    if (_watching && !_status.Text.StartsWith("["))
                        _status.Text = string.Format("Watching  {0:HH:mm:ss}", DateTime.Now);
                };
                _clockTimer.Start();

                _btnWatch.IsEnabled = true;
                _btnWatch.Background = AccentHover;
                _btnWatch.Foreground = Brushes.White;
                _btnWatch.BorderBrush = AccentHover;
                ((TextBlock)((StackPanel)_btnWatch.Content).Children[0]).Text = "\uE71A";
                ((TextBlock)((StackPanel)_btnWatch.Content).Children[1]).Text = "Stop";
                _status.Text = string.Format("Watching  {0:HH:mm:ss}", DateTime.Now);
            });
        }

        void StopWatching()
        {
            if (_eventWatcher != null) { _eventWatcher.Stop(); _eventWatcher = null; }
            if (_clockTimer != null) { _clockTimer.Stop(); _clockTimer = null; }
            _watching = false;
            _btnPull.IsEnabled = true;
            _btnWatch.Background = AccentBrush;
            _btnWatch.Foreground = Brushes.White;
            _btnWatch.BorderBrush = AccentBrush;
            ((TextBlock)((StackPanel)_btnWatch.Content).Children[0]).Text = "\uE768";
            ((TextBlock)((StackPanel)_btnWatch.Content).Children[1]).Text = "Watch";
            _status.Text = "Ready";
        }

        void OnWatchEvent(int profileIndex, string description)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                string pname = Config.PGet(profileIndex, "name", "Profile " + profileIndex);
                _status.Text = string.Format("[{0}] {1}", pname, description);

                RunOnStaThread(() =>
                {
                    var result = ProfileRunner.Run(profileIndex);
                    return result.Added + result.Modified;
                },
                total =>
                {
                    if (total > 0)
                    {
                        _status.Text = string.Format("{0:HH:mm:ss}  +{1} [{2}]",
                            DateTime.Now, total, pname);
                        ToastPopup.Show("watchbox",
                            string.Format("{0}: {1} new", pname, total));
                    }
                    else
                    {
                        _status.Text = string.Format("Watching  {0:HH:mm:ss}", DateTime.Now);
                    }
                });
            }));
        }

        void RunOnStaThread(Func<int> work, Action<int> onComplete)
        {
            var thread = new Thread(() =>
            {
                try
                {
                    int result = work();
                    Dispatcher.BeginInvoke(new Action(() => onComplete(result)));
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        "RunOnStaThread error: " + ex.ToString());
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        _status.Text = "Error: " + ex.Message;
                        if (_pulling) { _pulling = false; SetPullButton(false); }
                        if (_watching) StopWatching();
                    }));
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            StopWatching();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Config.Save();
            base.OnClosing(e);
        }
    }
}
