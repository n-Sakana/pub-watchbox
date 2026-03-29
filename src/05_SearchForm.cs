using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace MailPull
{
    public class SearchWindow : Window
    {
        ComboBox _cmbProfile;
        TreeView _treeView;
        ListView _mailList;
        TextBlock _headerBlock;
        TextBox _bodyBox;
        ListBox _attachList;
        TextBox _txtSearch;
        UIElement _placeholder;
        TextBlock _statusText;

        // Records per profile index
        Dictionary<int, List<string[]>> _profileRecords = new Dictionary<int, List<string[]>>();
        // Active records (filtered by selected profile)
        List<string[]> _activeRecords = new List<string[]>();
        // Currently displayed (after folder + search filter)
        List<string[]> _currentRecords = new List<string[]>();
        string _selectedFolder = "";

        public SearchWindow()
        {
            Title = "Viewer";
            ResizeMode = ResizeMode.CanResizeWithGrip;
            Width = 960; Height = 620;
            MinWidth = 700; MinHeight = 400;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = Brushes.White;
            FontFamily = new FontFamily("Segoe UI");
            FontSize = 13;

            var root = new DockPanel();

            // -- Top: search bar --
            var searchBar = new Border
            {
                BorderThickness = new Thickness(0, 0, 0, 1),
                BorderBrush = new SolidColorBrush(Color.FromRgb(230, 230, 230)),
                Background = new SolidColorBrush(Color.FromRgb(248, 248, 248)),
                Padding = new Thickness(12, 8, 12, 8)
            };
            var searchGrid = new Grid();
            _txtSearch = new TextBox
            {
                Padding = new Thickness(22, 4, 6, 4),
                BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 200)),
                Background = Brushes.White
            };
            var phPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                IsHitTestVisible = false,
                Margin = new Thickness(8, 5, 0, 0)
            };
            phPanel.Children.Add(new TextBlock
            {
                Text = "\uE721",
                FontFamily = new FontFamily("Segoe MDL2 Assets"),
                Foreground = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
                FontSize = 13, Margin = new Thickness(0, 0, 6, 0)
            });
            phPanel.Children.Add(new TextBlock
            {
                Text = "Search mail",
                Foreground = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
                FontSize = 12
            });
            _placeholder = phPanel;
            _txtSearch.TextChanged += (s, e) =>
            {
                _placeholder.Visibility = string.IsNullOrEmpty(_txtSearch.Text)
                    ? Visibility.Visible : Visibility.Collapsed;
                ApplyFilter();
            };
            searchGrid.Children.Add(_txtSearch);
            searchGrid.Children.Add(_placeholder);
            searchBar.Child = searchGrid;
            DockPanel.SetDock(searchBar, Dock.Top);
            root.Children.Add(searchBar);

            // -- Bottom: status bar --
            var statusBar = new Border
            {
                BorderThickness = new Thickness(0, 1, 0, 0),
                BorderBrush = new SolidColorBrush(Color.FromRgb(230, 230, 230)),
                Background = new SolidColorBrush(Color.FromRgb(248, 248, 248)),
                Padding = new Thickness(12, 6, 12, 6)
            };
            _statusText = new TextBlock
            {
                Text = "Ready",
                Foreground = new SolidColorBrush(Color.FromRgb(128, 128, 128)),
                FontSize = 11
            };
            statusBar.Child = _statusText;
            DockPanel.SetDock(statusBar, Dock.Bottom);
            root.Children.Add(statusBar);

            // -- Main 3-pane --
            var mainGrid = new Grid();
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(220), MinWidth = 140 });
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(300), MinWidth = 180 });
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star), MinWidth = 200 });

            // Left: profile combo + folder tree
            var leftPanel = new DockPanel();
            _cmbProfile = new ComboBox { Margin = new Thickness(4), FontSize = 12 };
            _cmbProfile.SelectionChanged += OnProfileChanged;
            DockPanel.SetDock(_cmbProfile, Dock.Top);
            leftPanel.Children.Add(_cmbProfile);

            _treeView = new TreeView
            {
                BorderThickness = new Thickness(0),
                Background = new SolidColorBrush(Color.FromRgb(252, 252, 252)),
                FontSize = 12
            };
            ScrollViewer.SetHorizontalScrollBarVisibility(_treeView, ScrollBarVisibility.Auto);
            _treeView.SelectedItemChanged += OnFolderSelected;
            leftPanel.Children.Add(_treeView);

            Grid.SetColumn(leftPanel, 0);
            mainGrid.Children.Add(leftPanel);

            mainGrid.Children.Add(MkVSplitter(1));
            mainGrid.Children.Add(MkVSplitter(3));

            // Center: mail list
            _mailList = new ListView { BorderThickness = new Thickness(0) };
            var gridView = new GridView();
            gridView.Columns.Add(new GridViewColumn
            {
                Header = "Date", Width = 90,
                DisplayMemberBinding = new Binding("[0]"),
                HeaderContainerStyle = LeftAlignStyle()
            });
            gridView.Columns.Add(new GridViewColumn
            {
                Header = "Subject",
                DisplayMemberBinding = new Binding("[1]"),
                HeaderContainerStyle = LeftAlignStyle()
            });
            var attachCol = new GridViewColumn
            {
                Width = 32,
                DisplayMemberBinding = new Binding("[2]"),
                HeaderContainerStyle = LeftAlignStyle()
            };
            var attachHeader = new TextBlock
            {
                Text = "\uE723",
                FontFamily = new FontFamily("Segoe MDL2 Assets"),
                FontSize = 12
            };
            attachCol.Header = attachHeader;
            gridView.Columns.Add(attachCol);
            _mailList.View = gridView;
            _mailList.SelectionChanged += OnMailSelected;
            _mailList.SizeChanged += (s, e) =>
            {
                if (gridView.Columns.Count > 2)
                    gridView.Columns[1].Width = _mailList.ActualWidth - 132;
            };
            Grid.SetColumn(_mailList, 2);
            mainGrid.Children.Add(_mailList);

            // Right: header+body (top) + attachments (bottom)
            var rightGrid = new Grid();
            rightGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star), MinHeight = 100 });
            rightGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rightGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(100), MinHeight = 60 });

            var bodyPanel = new DockPanel();
            _headerBlock = new TextBlock
            {
                Padding = new Thickness(12, 10, 12, 10),
                Background = new SolidColorBrush(Color.FromRgb(250, 250, 250)),
                TextWrapping = TextWrapping.Wrap, FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(80, 80, 80))
            };
            DockPanel.SetDock(_headerBlock, Dock.Top);
            bodyPanel.Children.Add(_headerBlock);
            _bodyBox = new TextBox
            {
                IsReadOnly = true, TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(12, 8, 12, 8),
                FontSize = 13, Background = Brushes.White
            };
            bodyPanel.Children.Add(_bodyBox);
            Grid.SetRow(bodyPanel, 0);
            rightGrid.Children.Add(bodyPanel);

            var hSplitter = new GridSplitter
            {
                Height = 4, HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Center,
                Background = new SolidColorBrush(Color.FromRgb(230, 230, 230))
            };
            Grid.SetRow(hSplitter, 1);
            rightGrid.Children.Add(hSplitter);

            var attachPanel = new DockPanel();
            var attachHdr = new TextBlock
            {
                Text = " Attachments", FontSize = 11,
                Padding = new Thickness(12, 4, 0, 4),
                Foreground = new SolidColorBrush(Color.FromRgb(100, 100, 100)),
                Background = new SolidColorBrush(Color.FromRgb(248, 248, 248))
            };
            DockPanel.SetDock(attachHdr, Dock.Top);
            attachPanel.Children.Add(attachHdr);
            _attachList = new ListBox { BorderThickness = new Thickness(0), Background = Brushes.White };
            _attachList.MouseDoubleClick += OnAttachDoubleClick;
            attachPanel.Children.Add(_attachList);
            Grid.SetRow(attachPanel, 2);
            rightGrid.Children.Add(attachPanel);

            Grid.SetColumn(rightGrid, 4);
            mainGrid.Children.Add(rightGrid);

            root.Children.Add(mainGrid);
            Content = root;

            Loaded += (s, e) => LoadData();
        }

        // --- Helpers ---

        GridSplitter MkVSplitter(int col)
        {
            var sp = new GridSplitter
            {
                Width = 4, HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Stretch,
                Background = new SolidColorBrush(Color.FromRgb(230, 230, 230))
            };
            Grid.SetColumn(sp, col);
            return sp;
        }

        static Style LeftAlignStyle()
        {
            var style = new Style(typeof(GridViewColumnHeader));
            style.Setters.Add(new Setter(GridViewColumnHeader.HorizontalContentAlignmentProperty,
                HorizontalAlignment.Left));
            style.Setters.Add(new Setter(GridViewColumnHeader.BorderThicknessProperty,
                new Thickness(0)));
            return style;
        }

        // --- Load all profiles' manifest data ---

        void LoadData()
        {
            _profileRecords.Clear();
            _cmbProfile.Items.Clear();
            _cmbProfile.Items.Add("(All)");

            for (int p = 0; p < Config.ProfileCount; p++)
            {
                _cmbProfile.Items.Add(Config.PGet(p, "name", "Profile " + (p + 1)));
                var records = new List<string[]>();
                string root = Config.PGet(p, "export_root");
                if (!string.IsNullOrEmpty(root))
                {
                    string csvPath = Path.Combine(root, "manifest.csv");
                    if (File.Exists(csvPath))
                        foreach (var line in File.ReadAllLines(csvPath, System.Text.Encoding.UTF8))
                        {
                            if (string.IsNullOrEmpty(line)) continue;
                            var cols = line.Split(',');
                            if (cols.Length < 6 || cols[0] == "entry_id") continue;
                            records.Add(cols);
                        }
                }
                _profileRecords[p] = records;
            }

            _cmbProfile.SelectedIndex = 0;
        }

        // --- Profile selection ---

        void OnProfileChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_cmbProfile.SelectedIndex < 0) return;
            _selectedFolder = "";
            RebuildActive();
            RebuildTree();
            ApplyFilter();
        }

        void RebuildActive()
        {
            _activeRecords.Clear();
            int sel = _cmbProfile.SelectedIndex - 1; // -1 = all
            var seenIds = new HashSet<string>();

            if (sel < 0)
            {
                foreach (var kv in _profileRecords)
                    foreach (var r in kv.Value)
                    {
                        if (seenIds.Contains(r[0])) continue;
                        seenIds.Add(r[0]);
                        _activeRecords.Add(r);
                    }
            }
            else
            {
                List<string[]> recs;
                if (_profileRecords.TryGetValue(sel, out recs))
                    _activeRecords.AddRange(recs);
            }
        }

        void RebuildTree()
        {
            _treeView.Items.Clear();
            var folders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in _activeRecords)
                if (r.Length > 5) folders.Add(r[5]);

            var nodes = new Dictionary<string, TreeViewItem>(StringComparer.OrdinalIgnoreCase);
            var sorted = new List<string>(folders);
            sorted.Sort(StringComparer.OrdinalIgnoreCase);

            foreach (var fp in sorted)
            {
                var parts = new List<string>();
                foreach (var p in fp.Split('\\'))
                    if (p.Length > 0) parts.Add(p);
                if (parts.Count == 0) continue;

                string cum = "";
                TreeViewItem parent = null;
                for (int i = 0; i < parts.Count; i++)
                {
                    cum += "\\" + parts[i];
                    TreeViewItem existing;
                    if (nodes.TryGetValue(cum, out existing)) { parent = existing; continue; }
                    var node = new TreeViewItem
                    {
                        Header = parts[i], Tag = cum,
                        IsExpanded = i == 0,
                        FontWeight = i == 0 ? FontWeights.SemiBold : FontWeights.Normal
                    };
                    nodes[cum] = node;
                    if (parent != null) parent.Items.Add(node);
                    else _treeView.Items.Add(node);
                    parent = node;
                }
            }
        }

        // --- Folder selected ---

        void OnFolderSelected(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var item = _treeView.SelectedItem as TreeViewItem;
            if (item == null || item.Tag == null) return;
            _selectedFolder = (string)item.Tag;
            ApplyFilter();
        }

        // --- Filter: profile + folder + search ---

        void ApplyFilter()
        {
            _mailList.Items.Clear();
            _currentRecords.Clear();
            ClearDetail();

            string q = (_txtSearch.Text ?? "").Trim().ToLower();

            foreach (var r in _activeRecords)
            {
                if (_selectedFolder.Length > 0)
                {
                    string fp = r.Length > 5 ? r[5].TrimStart('\\') : "";
                    string sel = _selectedFolder.TrimStart('\\');
                    if (!fp.StartsWith(sel, StringComparison.OrdinalIgnoreCase))
                        continue;
                }

                if (q.Length > 0)
                {
                    string subject = r.Length > 3 ? r[3].ToLower() : "";
                    string email = r.Length > 1 ? r[1].ToLower() : "";
                    string name = r.Length > 2 ? r[2].ToLower() : "";
                    string attach = r.Length > 8 ? r[8].ToLower() : "";
                    string body = r.Length > 10 ? r[10].ToLower() : "";

                    if (!subject.Contains(q) && !email.Contains(q) && !name.Contains(q)
                        && !attach.Contains(q) && !body.Contains(q))
                        continue;
                }

                _currentRecords.Add(r);
                string date = r.Length > 4 ? (r[4].Length >= 10 ? r[4].Substring(0, 10) : r[4]) : "";
                string att = r.Length > 8 && !string.IsNullOrEmpty(r[8])
                    ? r[8].Split('|').Length.ToString() : "";
                _mailList.Items.Add(new[] { date, r.Length > 3 ? r[3] : "", att });
            }

            _statusText.Text = string.Format("{0} mail(s)", _currentRecords.Count);
        }

        // --- Mail selected ---

        void OnMailSelected(object sender, SelectionChangedEventArgs e)
        {
            int idx = _mailList.SelectedIndex;
            if (idx < 0 || idx >= _currentRecords.Count) return;

            var r = _currentRecords[idx];
            string name = r.Length > 2 ? r[2] : "";
            string email = r.Length > 1 ? r[1] : "";
            string date = r.Length > 4 ? r[4] : "";
            string subject = r.Length > 3 ? r[3] : "";
            string mailDir = r.Length > 9 ? r[9] : "";

            _headerBlock.Text = string.Format("From: {0} ({1})\nDate: {2}\nSubject: {3}",
                name, email, date, subject);

            string bodyPath = r.Length > 6 ? r[6] : "";
            try
            {
                if (!string.IsNullOrEmpty(bodyPath) && File.Exists(bodyPath))
                    _bodyBox.Text = File.ReadAllText(bodyPath, System.Text.Encoding.UTF8);
                else
                    _bodyBox.Text = r.Length > 10 ? r[10] : "";
            }
            catch { _bodyBox.Text = r.Length > 10 ? r[10] : ""; }

            _attachList.Items.Clear();
            if (!string.IsNullOrEmpty(mailDir))
            {
                try
                {
                    foreach (var f in Directory.GetFiles(mailDir))
                    {
                        string fn = Path.GetFileName(f);
                        if (fn == "meta.json" || fn == "body.txt" || fn == "mail.msg") continue;
                        _attachList.Items.Add(new ListBoxItem
                        {
                            Content = fn, Tag = f,
                            Padding = new Thickness(8, 3, 8, 3)
                        });
                    }
                }
                catch { }
            }
        }

        void OnAttachDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = _attachList.SelectedItem as ListBoxItem;
            if (item == null || item.Tag == null) return;
            try { System.Diagnostics.Process.Start((string)item.Tag); }
            catch { }
        }

        void ClearDetail()
        {
            _headerBlock.Text = "";
            _bodyBox.Text = "";
            _attachList.Items.Clear();
        }
    }
}
