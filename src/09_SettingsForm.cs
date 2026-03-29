using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace WatchBox
{
    public class SettingsWindow : Window
    {
        ComboBox _cmbProfile;
        TextBox _txtName;
        ComboBox _cmbType;
        // Mail source (swapped in/out of SOURCE section)
        Grid _rowAccount, _rowOutlookFolder;
        ComboBox _cmbAccount, _cmbFolder;
        // Folder source (swapped in/out of SOURCE section)
        Grid _rowSourceFolder, _rowRecurse;
        TextBox _txtSourceFolder;
        CheckBox _chkRecurse;
        // SOURCE container
        StackPanel _sourceContent;
        // Common
        DatePicker _dpSince;
        RadioButton _rbAnd, _rbOr;
        ListBox _lstFilters;
        TextBox _txtNewFilter;
        TextBox _txtPath;
        CheckBox _chkFlat;
        CheckBox _chkNotify, _chkLog;

        List<string> _folderPaths = new List<string>();
        List<string> _currentFilters = new List<string>();
        int _currentIdx = -1;
        bool _loading;

        public SettingsWindow()
        {
            Title = "Settings";
            Width = 480; SizeToContent = SizeToContent.Height;
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

            // Name + Type
            root.Children.Add(FieldRow("Name", _txtName = new TextBox()));
            _cmbType = new ComboBox { IsEditable = false, Width = 160 };
            _cmbType.Items.Add("Mail");
            _cmbType.Items.Add("Folder");
            _cmbType.SelectionChanged += OnTypeChanged;
            root.Children.Add(FieldRow("Type", _cmbType));

            // --- SOURCE (single section, content swapped by type) ---
            root.Children.Add(new Separator { Margin = new Thickness(0, 8, 0, 8) });
            root.Children.Add(SectionHeader("SOURCE"));
            _sourceContent = new StackPanel();
            root.Children.Add(_sourceContent);

            // Build mail source rows (not added yet)
            _rowAccount = FieldRow("Account", _cmbAccount = MkCombo());
            _rowOutlookFolder = FieldRow("Outlook folder", _cmbFolder = MkCombo());

            // Build folder source rows (not added yet)
            _txtSourceFolder = new TextBox();
            var btnBrowseSource = new Button { Content = "...", Width = 30, Padding = new Thickness(0, 2, 0, 2) };
            btnBrowseSource.Click += (s, e) =>
            {
                string path = FolderPicker.Show(_txtSourceFolder.Text);
                if (path != null) _txtSourceFolder.Text = path;
            };
            _rowSourceFolder = FieldRow("Source folder", _txtSourceFolder, btnBrowseSource);
            _chkRecurse = new CheckBox { Content = "Include subfolders", FontSize = 12, IsChecked = true };
            _rowRecurse = FieldRow("", _chkRecurse);

            // --- FILTER (common) ---
            root.Children.Add(new Separator { Margin = new Thickness(0, 8, 0, 8) });
            root.Children.Add(SectionHeader("FILTER"));
            _dpSince = new DatePicker { Width = 140, SelectedDateFormat = DatePickerFormat.Short };
            root.Children.Add(FieldRow("Since", _dpSince));

            var modePanel = new StackPanel { Orientation = Orientation.Horizontal };
            _rbOr = new RadioButton { Content = "Any match (OR)", Margin = new Thickness(0, 0, 16, 0) };
            _rbAnd = new RadioButton { Content = "All match (AND)" };
            _rbOr.IsChecked = true;
            modePanel.Children.Add(_rbOr);
            modePanel.Children.Add(_rbAnd);
            root.Children.Add(FieldRow("Match", modePanel));

            _lstFilters = new ListBox { Height = 50, Margin = new Thickness(0, 4, 0, 4) };
            var filterAddBar = new DockPanel();
            var btnFilterRemove = new Button { Content = " \u2212 ", Padding = new Thickness(6, 1, 6, 1) };
            var btnFilterAdd = new Button { Content = " + ", Padding = new Thickness(6, 1, 6, 1), Margin = new Thickness(0, 0, 4, 0) };
            DockPanel.SetDock(btnFilterRemove, Dock.Right);
            DockPanel.SetDock(btnFilterAdd, Dock.Right);
            filterAddBar.Children.Add(btnFilterRemove);
            filterAddBar.Children.Add(btnFilterAdd);
            _txtNewFilter = new TextBox { Margin = new Thickness(0, 0, 4, 0) };
            _txtNewFilter.KeyDown += (s, e) => { if (e.Key == System.Windows.Input.Key.Enter) AddFilter(); };
            filterAddBar.Children.Add(_txtNewFilter);
            btnFilterAdd.Click += (s, e) => AddFilter();
            btnFilterRemove.Click += (s, e) => RemoveFilter();
            var filterContent = new StackPanel();
            filterContent.Children.Add(_lstFilters);
            filterContent.Children.Add(filterAddBar);
            root.Children.Add(FieldRow("Keywords", filterContent));

            // --- OUTPUT (common) ---
            root.Children.Add(new Separator { Margin = new Thickness(0, 8, 0, 8) });
            root.Children.Add(SectionHeader("OUTPUT"));
            root.Children.Add(FieldRow("Folder", _txtPath = new TextBox(), MkBrowseBtn()));
            _chkFlat = new CheckBox { Content = "Flat (no folder structure)", FontSize = 12 };
            root.Children.Add(FieldRow("", _chkFlat));

            // --- MONITORING (common) ---
            root.Children.Add(new Separator { Margin = new Thickness(0, 8, 0, 8) });
            root.Children.Add(SectionHeader("MONITORING"));
            _chkNotify = new CheckBox { Content = "Show notifications", FontSize = 12, IsChecked = true };
            _chkLog = new CheckBox { Content = "Write change log", FontSize = 12, IsChecked = true };
            root.Children.Add(FieldRow("", _chkNotify));
            root.Children.Add(FieldRow("", _chkLog));

            // Bottom buttons
            var btnBar = new DockPanel { Margin = new Thickness(0, 16, 0, 0) };
            var btnReset = new Button { Content = "Reset All", Padding = new Thickness(8, 6, 8, 6) };
            var btnImport = new Button { Content = "Import CSV", Padding = new Thickness(8, 6, 8, 6), Margin = new Thickness(0, 0, 4, 0) };
            DockPanel.SetDock(btnImport, Dock.Left);
            DockPanel.SetDock(btnReset, Dock.Left);
            btnBar.Children.Add(btnImport);
            btnBar.Children.Add(btnReset);
            var rightBtns = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var btnSave = new Button { Content = "Save", Width = 80, Padding = new Thickness(0, 6, 0, 6), IsDefault = true };
            var btnCancel = new Button { Content = "Cancel", Width = 80, Padding = new Thickness(0, 6, 0, 6), Margin = new Thickness(8, 0, 0, 0), IsCancel = true };
            rightBtns.Children.Add(btnSave);
            rightBtns.Children.Add(btnCancel);
            btnBar.Children.Add(rightBtns);
            root.Children.Add(btnBar);

            Content = root;

            _cmbAccount.SelectionChanged += (s, e) => { if (!_loading) LoadFolders(); };
            btnSave.Click += OnSave;
            btnCancel.Click += (s, e) => Close();
            btnReset.Click += OnResetAll;
            btnImport.Click += OnImportCsv;

            LoadProfileList();
            if (_cmbProfile.Items.Count > 0)
                _cmbProfile.SelectedIndex = 0;

            Loaded += (s, e) => LoadOutlookData();
        }

        // --- Type switching: swap source rows ---

        void OnTypeChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_loading) return;
            UpdateSourceContent();
        }

        void UpdateSourceContent()
        {
            _sourceContent.Children.Clear();
            string type = GetSelectedType();
            if (type == "mail")
            {
                _sourceContent.Children.Add(_rowAccount);
                _sourceContent.Children.Add(_rowOutlookFolder);
            }
            else
            {
                _sourceContent.Children.Add(_rowSourceFolder);
                _sourceContent.Children.Add(_rowRecurse);
            }
        }

        string GetSelectedType()
        {
            return _cmbType.SelectedIndex == 1 ? "folder" : "mail";
        }

        void SetTypeCombo(string type)
        {
            _cmbType.SelectedIndex = type == "folder" ? 1 : 0;
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
            var grid = new Grid { Margin = new Thickness(0, 0, 0, 6) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var lbl = new TextBlock
            {
                Text = label, FontSize = 12, VerticalAlignment = VerticalAlignment.Top,
                Foreground = new SolidColorBrush(Color.FromRgb(100, 100, 100)),
                Margin = new Thickness(0, 4, 0, 0)
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

        ComboBox MkCombo() { return new ComboBox { IsEditable = false }; }

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

        // --- Filter list ---

        void AddFilter()
        {
            string kw = _txtNewFilter.Text.Trim();
            if (string.IsNullOrEmpty(kw)) return;
            _currentFilters.Add(kw);
            _lstFilters.Items.Add(kw);
            _txtNewFilter.Text = "";
            _txtNewFilter.Focus();
        }

        void RemoveFilter()
        {
            int idx = _lstFilters.SelectedIndex;
            if (idx < 0) return;
            _currentFilters.RemoveAt(idx);
            _lstFilters.Items.RemoveAt(idx);
        }

        // --- Profile management ---

        void LoadProfileList()
        {
            _cmbProfile.Items.Clear();
            for (int i = 0; i < Config.ProfileCount; i++)
                _cmbProfile.Items.Add(Config.PGet(i, "name", "Profile " + (i + 1)));
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
            SetTypeCombo(Config.PGet(i, "type", "mail"));
            UpdateSourceContent();

            // Mail source
            string savedAcct = Config.PGet(i, "account");
            _cmbAccount.SelectedIndex = 0;
            for (int j = 0; j < _cmbAccount.Items.Count; j++)
                if (string.Equals(_cmbAccount.Items[j].ToString(), savedAcct,
                    StringComparison.OrdinalIgnoreCase))
                { _cmbAccount.SelectedIndex = j; break; }

            LoadFolders();
            string savedFolder = Config.PGet(i, "outlook_folder");
            if (!string.IsNullOrEmpty(savedFolder))
                for (int j = 0; j < _folderPaths.Count; j++)
                    if (_folderPaths[j] == savedFolder)
                    { _cmbFolder.SelectedIndex = j; break; }

            // Folder source
            _txtSourceFolder.Text = Config.PGet(i, "source_folder");
            _chkRecurse.IsChecked = Config.PGet(i, "recurse", "1") != "0";

            // Common filter
            string sinceStr = Config.PGet(i, "since");
            DateTime dt;
            _dpSince.SelectedDate = DateTime.TryParse(sinceStr, out dt) ? (DateTime?)dt : null;

            string mode = Config.PGet(i, "filter_mode", "or");
            _rbAnd.IsChecked = mode == "and";
            _rbOr.IsChecked = mode != "and";

            _currentFilters.Clear();
            _lstFilters.Items.Clear();
            string filtersStr = Config.PGet(i, "filters");
            if (!string.IsNullOrEmpty(filtersStr))
                foreach (var f in filtersStr.Split(';'))
                {
                    string kw = f.Trim();
                    if (kw.Length > 0) { _currentFilters.Add(kw); _lstFilters.Items.Add(kw); }
                }

            // Common output
            _txtPath.Text = Config.PGet(i, "output_root");
            _chkFlat.IsChecked = Config.PGet(i, "flat_output") == "1";

            // Common monitoring
            _chkNotify.IsChecked = Config.PGet(i, "notify", "1") != "0";
            _chkLog.IsChecked = Config.PGet(i, "log_enabled", "1") != "0";

            _loading = false;
        }

        void SaveCurrentProfile()
        {
            int i = _currentIdx;
            if (i < 0) return;

            Config.PSet(i, "name", _txtName.Text);
            Config.PSet(i, "type", GetSelectedType());

            // Mail source
            Config.PSet(i, "account",
                _cmbAccount.SelectedIndex > 0 ? _cmbAccount.SelectedItem.ToString() : "");
            Config.PSet(i, "outlook_folder",
                _cmbFolder.SelectedIndex > 0 ? _folderPaths[_cmbFolder.SelectedIndex] : "");

            // Folder source
            Config.PSet(i, "source_folder", _txtSourceFolder.Text);
            Config.PSet(i, "recurse", _chkRecurse.IsChecked == true ? "1" : "0");

            // Common filter
            Config.PSet(i, "since",
                _dpSince.SelectedDate.HasValue ? _dpSince.SelectedDate.Value.ToString("yyyy-MM-dd") : "");
            Config.PSet(i, "filter_mode", _rbAnd.IsChecked == true ? "and" : "or");
            Config.PSet(i, "filters", string.Join(";", _currentFilters.ToArray()));

            // Common output
            Config.PSet(i, "output_root", _txtPath.Text);
            Config.PSet(i, "flat_output", _chkFlat.IsChecked == true ? "1" : "0");

            // Common monitoring
            Config.PSet(i, "notify", _chkNotify.IsChecked == true ? "1" : "0");
            Config.PSet(i, "log_enabled", _chkLog.IsChecked == true ? "1" : "0");

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

                var scanner = new MailScanner();
                foreach (var a in scanner.GetAccounts()) _cmbAccount.Items.Add(a);

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
            var scanner = new MailScanner();
            foreach (var f in scanner.GetFolders(acct))
            {
                _cmbFolder.Items.Add(f[0]);
                _folderPaths.Add(f[1]);
            }
            _cmbFolder.SelectedIndex = 0;
        }

        // --- Save ---

        void OnSave(object sender, RoutedEventArgs e)
        {
            SaveCurrentProfile();
            for (int i = 0; i < Config.ProfileCount; i++)
            {
                if (string.IsNullOrWhiteSpace(Config.PGet(i, "output_root")))
                {
                    MessageBox.Show(string.Format("Profile \"{0}\" needs an output folder.",
                        Config.PGet(i, "name", "Profile " + (i + 1))));
                    return;
                }
            }
            Config.Save();
            Close();
        }

        // --- Reset All ---

        void OnResetAll(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "All profiles will be deleted. Continue?",
                "Reset All", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            if (result != MessageBoxResult.OK) return;

            _currentIdx = -1;
            while (Config.ProfileCount > 0)
                Config.RemoveProfile(0);
            Config.AddProfile("Default");
            LoadProfileList();
            _cmbProfile.SelectedIndex = 0;
        }

        // --- Import CSV ---

        void OnImportCsv(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "CSV files|*.csv|All files|*.*";
            if (dlg.ShowDialog() != true) return;

            try
            {
                var lines = File.ReadAllLines(dlg.FileName, System.Text.Encoding.UTF8);
                int added = 0;
                for (int row = 0; row < lines.Length; row++)
                {
                    string line = lines[row].Trim();
                    if (string.IsNullOrEmpty(line)) continue;
                    var cols = line.Split(',');
                    if (cols[0].Trim().ToLower() == "name") continue;
                    if (cols.Length < 1) continue;

                    int idx = Config.AddProfile(ColVal(cols, 0, "Imported " + (added + 1)));
                    Config.PSet(idx, "type", ColVal(cols, 1, "mail"));
                    Config.PSet(idx, "output_root", ColVal(cols, 2, ""));
                    Config.PSet(idx, "account", ColVal(cols, 3, ""));
                    Config.PSet(idx, "outlook_folder", ColVal(cols, 4, ""));
                    Config.PSet(idx, "source_folder", ColVal(cols, 5, ""));
                    Config.PSet(idx, "since", "");
                    Config.PSet(idx, "filter_mode", "or");
                    Config.PSet(idx, "filters", "");
                    Config.PSet(idx, "flat_output", "0");
                    Config.PSet(idx, "recurse", "1");
                    added++;
                }

                LoadProfileList();
                if (_cmbProfile.Items.Count > 0)
                    _cmbProfile.SelectedIndex = _cmbProfile.Items.Count - 1;
                MessageBox.Show(string.Format("{0} profile(s) imported.", added));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Import failed: " + ex.Message);
            }
        }

        static string ColVal(string[] cols, int idx, string def)
        {
            if (idx >= cols.Length) return def;
            string v = cols[idx].Trim();
            return v.Length > 0 ? v : def;
        }
    }
}
