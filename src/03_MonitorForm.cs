using System;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Threading;

namespace MailPull
{
    public class MonitorWindow : Window
    {
        TextBlock _status;
        Button _btnAuto, _btnPull;
        DispatcherTimer _pollTimer;
        bool _polling;
        bool _busy;
        Exporter _runningExporter;

        static readonly Brush AccentBrush = new SolidColorBrush(Color.FromRgb(55, 120, 200));
        static readonly Brush AccentHover = new SolidColorBrush(Color.FromRgb(45, 100, 180));

        public MonitorWindow()
        {
            Title = "mailpull";
            Width = 380; Height = 240;
            ResizeMode = ResizeMode.NoResize;
            Background = Brushes.White;
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = SystemParameters.WorkArea.Right - 400;
            Top = SystemParameters.WorkArea.Top + 40;
            FontFamily = new FontFamily("Segoe UI");
            FontSize = 13;

            var dock = new DockPanel { LastChildFill = true };

            // Status bar
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

            // 2x2 button grid
            var grid = new UniformGrid { Rows = 2, Columns = 2, Margin = new Thickness(16) };

            var btnSettings = MkBtn("\uE713", "Settings", false);
            var btnViewer   = MkBtn("\uE8A1", "Viewer", false);
            _btnPull        = MkBtn("\uE74B", "Pull", false);
            _btnAuto        = MkBtn("\uE768", "Auto", true);

            grid.Children.Add(btnSettings);
            grid.Children.Add(btnViewer);
            grid.Children.Add(_btnPull);
            grid.Children.Add(_btnAuto);

            dock.Children.Add(grid);
            Content = dock;

            btnSettings.Click += (s, e) => {
                var f = new SettingsWindow();
                f.Owner = this;
                f.ShowDialog();
            };
            btnViewer.Click += (s, e) => { var w = new SearchWindow(); w.Owner = this; w.Show(); };
            _btnPull.Click += OnPullClick;
            _btnAuto.Click += OnAutoClick;
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

        // --- Pull: run all profiles once ---

        void OnPullClick(object sender, RoutedEventArgs e)
        {
            if (_runningExporter != null)
            {
                _runningExporter.CancelRequested = true;
                _status.Text = "Cancelling...";
                return;
            }

            int count = Config.ProfileCount;
            if (count == 0)
            { MessageBox.Show("No profiles configured.", "mailpull"); return; }

            SetPullButton(true);

            RunOnStaThread(() =>
            {
                int total = 0;
                for (int i = 0; i < Config.ProfileCount; i++)
                {
                    string root = Config.PGet(i, "export_root");
                    if (string.IsNullOrEmpty(root)) continue;
                    var ex = new Exporter();
                    _runningExporter = ex;
                    ex.ProgressChanged += (c, s) => Dispatcher.BeginInvoke(
                        new Action(() =>
                        {
                            if (s != null && s.Length > 36) s = s.Substring(0, 36);
                            _status.Text = string.Format("[{0}] {1}: {2}",
                                Config.PGet(i, "name", "Profile " + i), c, s);
                        }));
                    total += ex.Export(root, Config.PGet(i, "since"),
                        Config.PGet(i, "account"), Config.PGet(i, "folder_path"),
                        Config.PGet(i, "filter_mode"), Config.PGet(i, "filters"));
                    if (ex.CancelRequested) break;
                }
                return total;
            },
            total =>
            {
                bool cancelled = _runningExporter != null && _runningExporter.CancelRequested;
                _runningExporter = null;
                SetPullButton(false);
                _status.Text = cancelled
                    ? string.Format("Cancelled ({0} pulled)", total)
                    : string.Format("Done: {0} pulled", total);
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

        // --- Auto: poll all profiles ---

        void OnAutoClick(object sender, RoutedEventArgs e)
        {
            if (_polling)
            {
                _pollTimer.Stop();
                _pollTimer = null;
                _polling = false;
                _btnAuto.Background = SystemColors.ControlBrush;
                _btnAuto.Foreground = SystemColors.ControlTextBrush;
                _btnAuto.BorderBrush = SystemColors.ActiveBorderBrush;
                ((TextBlock)((StackPanel)_btnAuto.Content).Children[0]).Text = "\uE768";
                ((TextBlock)((StackPanel)_btnAuto.Content).Children[1]).Text = "Auto";
                _status.Text = "Ready";
            }
            else
            {
                if (Config.ProfileCount == 0)
                { MessageBox.Show("No profiles configured.", "mailpull"); return; }

                // Use shortest poll interval across profiles
                int minSec = 60;
                for (int i = 0; i < Config.ProfileCount; i++)
                {
                    int s = 60;
                    int.TryParse(Config.PGet(i, "poll_seconds", "60"), out s);
                    if (s < minSec) minSec = s;
                }
                if (minSec < 10) minSec = 10;

                _pollTimer = new DispatcherTimer();
                _pollTimer.Interval = TimeSpan.FromSeconds(minSec);
                _pollTimer.Tick += OnPollTick;
                _pollTimer.Start();
                _polling = true;

                _btnAuto.Background = AccentHover;
                _btnAuto.Foreground = Brushes.White;
                _btnAuto.BorderBrush = AccentHover;
                ((TextBlock)((StackPanel)_btnAuto.Content).Children[0]).Text = "\uE71A";
                ((TextBlock)((StackPanel)_btnAuto.Content).Children[1]).Text = "Stop";
                _status.Text = string.Format("Running  {0:HH:mm:ss}", DateTime.Now);

                OnPollTick(null, null);
            }
        }

        void OnPollTick(object sender, EventArgs e)
        {
            if (_busy) return;
            _busy = true;

            RunOnStaThread(() =>
            {
                int total = 0;
                for (int i = 0; i < Config.ProfileCount; i++)
                {
                    string root = Config.PGet(i, "export_root");
                    if (string.IsNullOrEmpty(root)) continue;
                    var ex = new Exporter();
                    total += ex.Export(root, Config.PGet(i, "since"),
                        Config.PGet(i, "account"), Config.PGet(i, "folder_path"),
                        Config.PGet(i, "filter_mode"), Config.PGet(i, "filters"));
                }
                return total;
            },
            total =>
            {
                _busy = false;
                _status.Text = total > 0
                    ? string.Format("{0:HH:mm:ss}  +{1} new", DateTime.Now, total)
                    : string.Format("Running  {0:HH:mm:ss}", DateTime.Now);
            });
        }

        void RunOnStaThread(Func<int> work, Action<int> onComplete)
        {
            var thread = new Thread(() =>
            {
                int result = work();
                Dispatcher.BeginInvoke(new Action(() => onComplete(result)));
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (_pollTimer != null) _pollTimer.Stop();
            Config.Save();
            base.OnClosing(e);
        }
    }
}
