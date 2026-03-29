using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace MailPull
{
    public class SettingsWindow : Window
    {
        ComboBox _cmbProfile;
        TextBox _txtName, _txtPath, _txtSince, _txtPollSec;
        ComboBox _cmbAccount, _cmbFolder;
        System.Collections.Generic.List<string> _folderPaths =
            new System.Collections.Generic.List<string>();
        int _currentIdx = -1;
        bool _loading;

        public SettingsWindow()
        {
            Title = "Settings";
            Width = 460; SizeToContent = SizeToContent.Height;
            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            Background = Brushes.White;
            FontFamily = new FontFamily("Segoe UI");
            FontSize = 13;

            var root = new StackPanel { Margin = new Thickness(20) };

            // Profile selector
            var profileBar = new DockPanel { Margin = new Thickness(0, 0, 0, 12) };
            var btnRemove = new Button { Content = " \u2212 ", Padding = new Thickness(6, 2, 6, 2) };
            var btnAdd = new Button { Content = " + ", Padding = new Thickness(6, 2, 6, 2), Margin = new Thickness(0, 0, 4, 0) };
            DockPanel.SetDock(btnRemove, Dock.Right);
            DockPanel.SetDock(btnAdd, Dock.Right);
            profileBar.Children.Add(btnRemove);
            profileBar.Children.Add(btnAdd);
            _cmbProfile = new ComboBox { Margin = new Thickness(0, 0, 8, 0) };
            _cmbProfile.SelectionChanged += OnProfileChanged;
            profileBar.Children.Add(_cmbProfile);
            root.Children.Add(profileBar);

            btnAdd.Click += OnAddProfile;
            btnRemove.Click += OnRemoveProfile;

            // Fields
            root.Children.Add(FieldRow("Name", _txtName = new TextBox()));
            root.Children.Add(new Separator { Margin = new Thickness(0, 8, 0, 8) });
            root.Children.Add(SectionHeader("EXPORT"));
            root.Children.Add(FieldRow("Folder", _txtPath = new TextBox(), MkBrowseBtn()));
            root.Children.Add(FieldRow("Account", _cmbAccount = MkCombo()));
            root.Children.Add(FieldRow("Outlook folder", _cmbFolder = MkCombo()));
            root.Children.Add(FieldRow("Since", _txtSince = new TextBox { Width = 100 },
                HintLabel("yyyy-mm-dd")));

            root.Children.Add(new Separator { Margin = new Thickness(0, 8, 0, 8) });
            root.Children.Add(SectionHeader("AUTO POLLING"));
            root.Children.Add(FieldRow("Interval", _txtPollSec = new TextBox { Width = 60 },
                HintLabel("seconds")));

            // Buttons
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 16, 0, 0)
            };
            var btnSave = new Button { Content = "Save", Width = 80, Padding = new Thickness(0, 6, 0, 6), IsDefault = true };
            var btnCancel = new Button { Content = "Cancel", Width = 80, Padding = new Thickness(0, 6, 0, 6),
                Margin = new Thickness(8, 0, 0, 0), IsCancel = true };
            btnSave.Click += OnSave;
            btnCancel.Click += (s, e) => Close();
            btnPanel.Children.Add(btnSave);
            btnPanel.Children.Add(btnCancel);
            root.Children.Add(btnPanel);

            Content = root;

            _cmbAccount.SelectionChanged += (s, e) => { if (!_loading) LoadFolders(); };

            LoadProfileList();
            if (_cmbProfile.Items.Count > 0)
                _cmbProfile.SelectedIndex = 0;

            Loaded += (s, e) => LoadOutlookData();
        }

        // --- Layout helpers ---

        TextBlock SectionHeader(string text)
        {
            return new TextBlock
            {
                Text = text, FontSize = 11, FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(55, 120, 200)),
                Margin = new Thickness(0, 0, 0, 8)
            };
        }

        Grid FieldRow(string label, FrameworkElement input, UIElement extra = null)
        {
            var grid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var lbl = new TextBlock
            {
                Text = label, FontSize = 12, VerticalAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(Color.FromRgb(100, 100, 100))
            };
            Grid.SetColumn(lbl, 0);
            grid.Children.Add(lbl);

            if (extra != null)
            {
                var sp = new StackPanel { Orientation = Orientation.Horizontal };
                sp.Children.Add(input);
                var fe = extra as FrameworkElement;
                if (fe != null) fe.Margin = new Thickness(6, 0, 0, 0);
                sp.Children.Add(extra);
                Grid.SetColumn(sp, 1);
                grid.Children.Add(sp);
            }
            else
            {
                Grid.SetColumn(input, 1);
                grid.Children.Add(input);
            }
            return grid;
        }

        ComboBox MkCombo()
        {
            return new ComboBox { IsEditable = false };
        }

        Button MkBrowseBtn()
        {
            var btn = new Button { Content = "...", Width = 30, Padding = new Thickness(0, 2, 0, 2) };
            btn.Click += (s, e) =>
            {
                string path = FolderPicker.Show(_txtPath.Text);
                if (path != null) _txtPath.Text = path;
            };
            return btn;
        }

        TextBlock HintLabel(string text)
        {
            return new TextBlock
            {
                Text = text, FontSize = 12, VerticalAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(Color.FromRgb(160, 160, 160))
            };
        }

        // --- Profile management ---

        void LoadProfileList()
        {
            _cmbProfile.Items.Clear();
            for (int i = 0; i < Config.ProfileCount; i++)
            {
                string name = Config.PGet(i, "name", "Profile " + (i + 1));
                _cmbProfile.Items.Add(name);
            }
        }

        void OnProfileChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_cmbProfile.SelectedIndex < 0) return;
            SaveCurrentProfile();
            _currentIdx = _cmbProfile.SelectedIndex;
            LoadCurrentProfile();
        }

        void OnAddProfile(object sender, RoutedEventArgs e)
        {
            SaveCurrentProfile();
            int idx = Config.AddProfile("Profile " + (Config.ProfileCount));
            LoadProfileList();
            _cmbProfile.SelectedIndex = idx;
        }

        void OnRemoveProfile(object sender, RoutedEventArgs e)
        {
            if (Config.ProfileCount <= 1)
            { MessageBox.Show("Cannot remove the last profile."); return; }
            int idx = _cmbProfile.SelectedIndex;
            if (idx < 0) return;
            _currentIdx = -1;
            Config.RemoveProfile(idx);
            LoadProfileList();
            _cmbProfile.SelectedIndex = Math.Min(idx, Config.ProfileCount - 1);
        }

        void LoadCurrentProfile()
        {
            _loading = true;
            int i = _currentIdx;
            if (i < 0) { _loading = false; return; }
            _txtName.Text = Config.PGet(i, "name");
            _txtPath.Text = Config.PGet(i, "export_root");
            _txtSince.Text = Config.PGet(i, "export_since");
            _txtPollSec.Text = Config.PGet(i, "poll_seconds", "60");

            string savedAcct = Config.PGet(i, "account");
            for (int j = 0; j < _cmbAccount.Items.Count; j++)
                if (string.Equals(_cmbAccount.Items[j].ToString(), savedAcct,
                    StringComparison.OrdinalIgnoreCase))
                { _cmbAccount.SelectedIndex = j; break; }

            LoadFolders();
            string savedFolder = Config.PGet(i, "folder_path");
            if (!string.IsNullOrEmpty(savedFolder))
                for (int j = 0; j < _folderPaths.Count; j++)
                    if (_folderPaths[j] == savedFolder)
                    { _cmbFolder.SelectedIndex = j; break; }
            _loading = false;
        }

        void SaveCurrentProfile()
        {
            int i = _currentIdx;
            if (i < 0) return;
            Config.PSet(i, "name", _txtName.Text);
            Config.PSet(i, "export_root", _txtPath.Text);
            Config.PSet(i, "export_since", _txtSince.Text);
            Config.PSet(i, "poll_seconds", _txtPollSec.Text);
            Config.PSet(i, "account",
                _cmbAccount.SelectedIndex > 0 ? _cmbAccount.SelectedItem.ToString() : "");
            Config.PSet(i, "folder_path",
                _cmbFolder.SelectedIndex > 0 ? _folderPaths[_cmbFolder.SelectedIndex] : "");
            // Update profile list display name
            if (_cmbProfile.SelectedIndex == i && _cmbProfile.Items.Count > i)
                _cmbProfile.Items[i] = _txtName.Text;
        }

        // --- Outlook data ---

        void LoadOutlookData()
        {
            Cursor = System.Windows.Input.Cursors.Wait;
            _loading = true;
            try
            {
                _cmbAccount.Items.Clear();
                _cmbAccount.Items.Add("(All)");
                _cmbAccount.SelectedIndex = 0;

                var exporter = new Exporter();
                foreach (var a in exporter.GetAccounts()) _cmbAccount.Items.Add(a);

                if (_currentIdx >= 0) LoadCurrentProfile();
            }
            finally
            {
                _loading = false;
                Cursor = System.Windows.Input.Cursors.Arrow;
            }
        }

        void LoadFolders()
        {
            _cmbFolder.Items.Clear();
            _folderPaths.Clear();
            _cmbFolder.Items.Add("(All)");
            _folderPaths.Add("");
            string acct = _cmbAccount.SelectedIndex > 0 ? _cmbAccount.SelectedItem.ToString() : "";
            var exporter = new Exporter();
            foreach (var f in exporter.GetFolders(acct))
            {
                _cmbFolder.Items.Add(f[0]);
                _folderPaths.Add(f[1]);
            }
            _cmbFolder.SelectedIndex = 0;
        }

        void OnSave(object sender, RoutedEventArgs e)
        {
            SaveCurrentProfile();
            // Validate all profiles have export_root
            for (int i = 0; i < Config.ProfileCount; i++)
            {
                if (string.IsNullOrWhiteSpace(Config.PGet(i, "export_root")))
                {
                    MessageBox.Show(string.Format("Profile \"{0}\" needs an export folder.",
                        Config.PGet(i, "name", "Profile " + (i + 1))));
                    return;
                }
            }
            Config.Save();
            Close();
        }
    }
}
