using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace WatchBox
{
    public class SearchWindow : Window
    {
        ComboBox _cmbProfile;
        TreeView _treeView;
        ListView _itemList;
        TextBox _txtSearch;
        UIElement _placeholder;
        TextBlock _statusText;

        // Mail detail
        TextBlock _headerBlock;
        TextBox _bodyBox;
        ListBox _attachList;

        // Folder detail
        TextBlock _fileInfoBlock;

        // Right panel containers
        DockPanel _mailDetailPanel;
        DockPanel _folderDetailPanel;
        Grid _rightGrid;

        Dictionary<int, List<string[]>> _profileRecords = new Dictionary<int, List<string[]>>();
        Dictionary<int, string> _profileTypes = new Dictionary<int, string>();
        List<string[]> _activeRecords = new List<string[]>();
        List<string[]> _currentRecords = new List<string[]>();
        string _selectedFolder = "";
        string _currentType = "mail";
        Grid _mainGrid;
        ColumnDefinition _leftCol, _centerCol, _rightCol, _splitter2Col;
        GridSplitter _vSplitter2;

        // fin-studio light theme
        static readonly Brush TreeBg = new SolidColorBrush(Color.FromRgb(250, 250, 250));
        static readonly Brush TreeFg = new SolidColorBrush(Color.FromRgb(80, 80, 80));
        static readonly Brush TreeSelBg = new SolidColorBrush(Color.FromRgb(208, 232, 247));
        static readonly Brush TreeSelFg = new SolidColorBrush(Color.FromRgb(26, 26, 26));
        static readonly Brush TreeHoverBg = new SolidColorBrush(Color.FromRgb(230, 230, 230));
        static readonly Brush AccentBrush = new SolidColorBrush(Color.FromRgb(37, 99, 235));
        static readonly Brush PanelBg = Brushes.White;
        static readonly Brush BorderColor = new SolidColorBrush(Color.FromRgb(230, 230, 230));
        static readonly Brush MutedFg = new SolidColorBrush(Color.FromRgb(144, 144, 144));
        static readonly Brush SubtleFg = new SolidColorBrush(Color.FromRgb(80, 80, 80));
        static readonly Brush BarBg = new SolidColorBrush(Color.FromRgb(250, 250, 250));

        public SearchWindow()
        {
            Title = "Viewer";
            ResizeMode = ResizeMode.CanResizeWithGrip;
            Width = 960; Height = 620;
            MinWidth = 700; MinHeight = 400;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = PanelBg;
            FontFamily = new FontFamily("Segoe UI");
            FontSize = 13;

            var root = new DockPanel();

            // -- Top: search bar --
            var searchBar = new Border
            {
                BorderThickness = new Thickness(0, 0, 0, 1),
                BorderBrush = BorderColor, Background = BarBg,
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
                Text = "Search...",
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
                BorderBrush = BorderColor, Background = BarBg,
                Padding = new Thickness(12, 6, 12, 6)
            };
            _statusText = new TextBlock { Text = "Ready", Foreground = MutedFg, FontSize = 11 };
            statusBar.Child = _statusText;
            DockPanel.SetDock(statusBar, Dock.Bottom);
            root.Children.Add(statusBar);

            // -- Main grid (5 columns: left | splitter | center | splitter | right) --
            _mainGrid = new Grid();
            _leftCol = new ColumnDefinition { Width = new GridLength(240), MinWidth = 140 };
            _centerCol = new ColumnDefinition { Width = new GridLength(300), MinWidth = 180 };
            _splitter2Col = new ColumnDefinition { Width = GridLength.Auto };
            _rightCol = new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star), MinWidth = 200 };
            _mainGrid.ColumnDefinitions.Add(_leftCol);
            _mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            _mainGrid.ColumnDefinitions.Add(_centerCol);
            _mainGrid.ColumnDefinitions.Add(_splitter2Col);
            _mainGrid.ColumnDefinitions.Add(_rightCol);

            // Left: profile combo + VS Code-style tree
            var leftPanel = new DockPanel { Background = TreeBg };
            _cmbProfile = new ComboBox { Margin = new Thickness(4), FontSize = 12 };
            _cmbProfile.SelectionChanged += OnProfileChanged;
            DockPanel.SetDock(_cmbProfile, Dock.Top);
            leftPanel.Children.Add(_cmbProfile);

            _treeView = new TreeView
            {
                BorderThickness = new Thickness(0),
                Background = TreeBg, Foreground = TreeFg,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12, Padding = new Thickness(0, 4, 0, 4)
            };
            _treeView.Resources.Add(SystemColors.HighlightBrushKey, TreeSelBg);
            _treeView.Resources.Add(SystemColors.HighlightTextBrushKey, TreeSelFg);
            _treeView.Resources.Add(SystemColors.InactiveSelectionHighlightBrushKey, TreeSelBg);
            _treeView.Resources.Add(SystemColors.InactiveSelectionHighlightTextBrushKey, TreeSelFg);
            ScrollViewer.SetHorizontalScrollBarVisibility(_treeView, ScrollBarVisibility.Auto);
            _treeView.SelectedItemChanged += OnFolderSelected;
            leftPanel.Children.Add(_treeView);

            Grid.SetColumn(leftPanel, 0);
            _mainGrid.Children.Add(leftPanel);
            _mainGrid.Children.Add(MkVSplitter(1));
            _vSplitter2 = MkVSplitter(3);
            _mainGrid.Children.Add(_vSplitter2);

            // Right: content area (switches between mail and folder layouts)
            _rightGrid = new Grid();
            Grid.SetColumn(_rightGrid, 4);
            _mainGrid.Children.Add(_rightGrid);

            BuildMailLayout();
            BuildFolderLayout();

            root.Children.Add(_mainGrid);
            Content = root;

            Loaded += (s, e) => LoadData();
        }

        void BuildMailLayout()
        {
            // Item list (goes in center column for mail)
            _itemList = new ListView { BorderThickness = new Thickness(0) };
            var gridView = new GridView();
            gridView.Columns.Add(new GridViewColumn
            {
                Header = "Date", Width = 90,
                DisplayMemberBinding = new Binding("[0]"),
                HeaderContainerStyle = LeftAlignStyle()
            });
            gridView.Columns.Add(new GridViewColumn
            {
                Header = "Name",
                DisplayMemberBinding = new Binding("[1]"),
                HeaderContainerStyle = LeftAlignStyle()
            });
            var extraCol = new GridViewColumn
            {
                Width = 50,
                DisplayMemberBinding = new Binding("[2]"),
                HeaderContainerStyle = LeftAlignStyle()
            };
            extraCol.Header = new TextBlock
            {
                Text = "\uE723",
                FontFamily = new FontFamily("Segoe MDL2 Assets"),
                FontSize = 12
            };
            gridView.Columns.Add(extraCol);
            _itemList.View = gridView;
            _itemList.SelectionChanged += OnItemSelected;
            _itemList.SizeChanged += (s, e) =>
            {
                if (gridView.Columns.Count > 2)
                    gridView.Columns[1].Width = _itemList.ActualWidth - 152;
            };

            // Mail detail panel (header + body + attachments, goes in right column)
            _mailDetailPanel = new DockPanel();
            var detailGrid = new Grid();
            detailGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star), MinHeight = 100 });
            detailGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            detailGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(100), MinHeight = 60 });

            var bodyPanel = new DockPanel();
            _headerBlock = new TextBlock
            {
                Padding = new Thickness(12, 10, 12, 10),
                Background = new SolidColorBrush(Color.FromRgb(250, 250, 250)),
                TextWrapping = TextWrapping.Wrap, FontSize = 12, Foreground = SubtleFg
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
            detailGrid.Children.Add(bodyPanel);
            detailGrid.Children.Add(MkHSplitter(1));

            var attachPanel = new DockPanel();
            var attachHdr = new TextBlock
            {
                Text = " Attachments", FontSize = 11,
                Padding = new Thickness(12, 4, 0, 4),
                Foreground = SubtleFg, Background = BarBg
            };
            DockPanel.SetDock(attachHdr, Dock.Top);
            attachPanel.Children.Add(attachHdr);
            _attachList = new ListBox { BorderThickness = new Thickness(0), Background = Brushes.White };
            _attachList.MouseDoubleClick += OnAttachDoubleClick;
            attachPanel.Children.Add(_attachList);
            Grid.SetRow(attachPanel, 2);
            detailGrid.Children.Add(attachPanel);

            _mailDetailPanel.Children.Add(detailGrid);
        }

        void BuildFolderLayout()
        {
            _folderDetailPanel = new DockPanel();

            // Split: file list (top, large) + file info (bottom, small)
            var splitGrid = new Grid();
            splitGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star), MinHeight = 150 });
            splitGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            splitGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(120), MinHeight = 60 });

            // Reuse _itemList (it gets moved between panels)
            // File info panel
            var infoPanel = new DockPanel();
            var infoHdr = new TextBlock
            {
                Text = " File Info", FontSize = 11,
                Padding = new Thickness(12, 4, 0, 4),
                Foreground = SubtleFg, Background = BarBg
            };
            DockPanel.SetDock(infoHdr, Dock.Top);
            infoPanel.Children.Add(infoHdr);
            _fileInfoBlock = new TextBlock
            {
                Padding = new Thickness(12, 8, 12, 8),
                FontSize = 12, Foreground = SubtleFg,
                TextWrapping = TextWrapping.Wrap
            };
            infoPanel.Children.Add(_fileInfoBlock);
            Grid.SetRow(infoPanel, 2);
            splitGrid.Children.Add(MkHSplitter(1));
            splitGrid.Children.Add(infoPanel);

            _folderDetailPanel.Children.Add(splitGrid);
        }

        void SwitchLayout(string type)
        {
            _currentType = type;
            _rightGrid.Children.Clear();

            // Remove item list from wherever it is
            var parent = _itemList.Parent as Panel;
            if (parent != null) parent.Children.Remove(_itemList);
            if (_mainGrid.Children.Contains(_itemList)) _mainGrid.Children.Remove(_itemList);

            if (type == "folder")
            {
                // 2-pane: left tree (wide) + right (list + info)
                _leftCol.Width = new GridLength(300);
                _centerCol.Width = new GridLength(0);
                _centerCol.MinWidth = 0;
                _splitter2Col.Width = new GridLength(0);
                _vSplitter2.Visibility = Visibility.Collapsed;

                var grid = (Grid)_folderDetailPanel.Children[0];
                if (!grid.Children.Contains(_itemList))
                {
                    Grid.SetRow(_itemList, 0);
                    grid.Children.Insert(0, _itemList);
                }
                _rightGrid.Children.Add(_folderDetailPanel);
            }
            else
            {
                // 3-pane: left tree + center list + right detail
                _leftCol.Width = new GridLength(220);
                _centerCol.Width = new GridLength(300);
                _centerCol.MinWidth = 180;
                _splitter2Col.Width = GridLength.Auto;
                _vSplitter2.Visibility = Visibility.Visible;

                Grid.SetColumn(_itemList, 2);
                _mainGrid.Children.Add(_itemList);
                _rightGrid.Children.Add(_mailDetailPanel);
            }
        }

        // --- Tree (relative to output_root, all expanded) ---

        void RebuildTree()
        {
            _treeView.Items.Clear();

            int sel = _cmbProfile.SelectedIndex;
            string outputRoot = sel >= 0 ? Config.PGet(sel, "output_root") : "";
            if (!string.IsNullOrEmpty(outputRoot))
                outputRoot = outputRoot.TrimEnd('\\');

            // Collect unique folder paths and file counts
            // For mail: use folder_path as-is (Outlook paths)
            // For folder: make relative to output_root
            var folderCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in _activeRecords)
            {
                string fp = RecordFolderPath(r);
                if (string.IsNullOrEmpty(fp)) continue;
                string rel = fp.TrimStart('\\');
                if (_currentType != "mail" && !string.IsNullOrEmpty(outputRoot) &&
                    fp.StartsWith(outputRoot, StringComparison.OrdinalIgnoreCase))
                    rel = fp.Substring(outputRoot.Length).TrimStart('\\');
                if (rel.Length == 0) rel = ".";
                if (!folderCounts.ContainsKey(rel)) folderCounts[rel] = 0;
                folderCounts[rel]++;
            }

            // Build tree nodes
            var nodes = new Dictionary<string, TreeViewItem>(StringComparer.OrdinalIgnoreCase);
            var sorted = new List<string>(folderCounts.Keys);
            sorted.Sort(StringComparer.OrdinalIgnoreCase);

            foreach (var rel in sorted)
            {
                if (rel == ".")
                {
                    var rootNode = MkTreeNode(Path.GetFileName(outputRoot), ".", folderCounts[rel]);
                    nodes["."] = rootNode;
                    _treeView.Items.Add(rootNode);
                    continue;
                }
                var parts = rel.Split('\\');
                string cum = "";
                TreeViewItem parent = null;
                for (int i = 0; i < parts.Length; i++)
                {
                    cum = cum.Length == 0 ? parts[i] : cum + "\\" + parts[i];
                    TreeViewItem existing;
                    if (nodes.TryGetValue(cum, out existing)) { parent = existing; continue; }

                    int count = 0;
                    if (i == parts.Length - 1) folderCounts.TryGetValue(rel, out count);
                    var node = MkTreeNode(parts[i], cum, count);
                    nodes[cum] = node;
                    if (parent != null) parent.Items.Add(node);
                    else _treeView.Items.Add(node);
                    parent = node;
                }
            }
        }

        TreeViewItem MkTreeNode(string name, string tag, int count)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal };
            sp.Children.Add(new TextBlock
            {
                Text = "\uE8B7",
                FontFamily = new FontFamily("Segoe MDL2 Assets"),
                FontSize = 10, Foreground = MutedFg,
                Margin = new Thickness(0, 0, 5, 0),
                VerticalAlignment = VerticalAlignment.Center
            });
            var label = name;
            if (count > 0) label += "  " + count;
            sp.Children.Add(new TextBlock
            {
                Text = label, VerticalAlignment = VerticalAlignment.Center, Foreground = TreeFg
            });
            return new TreeViewItem
            {
                Header = sp, Tag = tag, IsExpanded = true,
                Padding = new Thickness(2, 1, 2, 1)
            };
        }

        void OnFolderSelected(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var item = _treeView.SelectedItem as TreeViewItem;
            if (item == null || item.Tag == null) return;
            _selectedFolder = (string)item.Tag;
            ApplyFilter();
        }

        // --- Helpers ---

        GridSplitter MkVSplitter(int col)
        {
            var sp = new GridSplitter
            {
                Width = 4, HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Stretch,
                Background = BorderColor
            };
            Grid.SetColumn(sp, col);
            return sp;
        }

        GridSplitter MkHSplitter(int row)
        {
            var sp = new GridSplitter
            {
                Height = 4, HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Center,
                Background = BorderColor
            };
            Grid.SetRow(sp, row);
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

        static bool IsMailRecord(string[] cols) { return cols.Length >= 11; }

        static string RecordDate(string[] r)
        {
            return IsMailRecord(r) ? (r.Length > 4 ? r[4] : "") : (r.Length > 6 ? r[6] : "");
        }

        static string RecordName(string[] r)
        {
            return IsMailRecord(r) ? (r.Length > 3 ? r[3] : "") : (r.Length > 1 ? r[1] : "");
        }

        static string RecordFolderPath(string[] r)
        {
            return IsMailRecord(r) ? (r.Length > 5 ? r[5] : "") : (r.Length > 3 ? r[3] : "");
        }

        static string RecordExtra(string[] r)
        {
            if (IsMailRecord(r))
            {
                string att = r.Length > 8 ? r[8] : "";
                if (string.IsNullOrEmpty(att)) return "";
                return att.Split('|').Length.ToString();
            }
            return r.Length > 5 ? FormatSize(r[5]) : "";
        }

        static string FormatSize(string sizeStr)
        {
            long size;
            if (!long.TryParse(sizeStr, out size)) return sizeStr;
            if (size < 1024) return size + " B";
            if (size < 1048576) return (size / 1024) + " KB";
            return string.Format("{0:0.0} MB", size / 1048576.0);
        }

        static bool RecordMatchesQuery(string[] r, string q)
        {
            if (IsMailRecord(r))
            {
                string subject = r.Length > 3 ? r[3].ToLower() : "";
                string email = r.Length > 1 ? r[1].ToLower() : "";
                string name = r.Length > 2 ? r[2].ToLower() : "";
                string body = r.Length > 10 ? r[10].ToLower() : "";
                return subject.Contains(q) || email.Contains(q) || name.Contains(q) || body.Contains(q);
            }
            string fileName = r.Length > 1 ? r[1].ToLower() : "";
            string relPath = r.Length > 4 ? r[4].ToLower() : "";
            return fileName.Contains(q) || relPath.Contains(q);
        }

        // --- Load data ---

        void LoadData()
        {
            _profileRecords.Clear();
            _profileTypes.Clear();
            _cmbProfile.Items.Clear();

            for (int p = 0; p < Config.ProfileCount; p++)
            {
                _cmbProfile.Items.Add(Config.PGet(p, "name", "Profile " + (p + 1)));
                string type = Config.PGet(p, "type", "mail");
                _profileTypes[p] = type;
                var records = new List<string[]>();
                string root = Config.PGet(p, "output_root");
                if (!string.IsNullOrEmpty(root))
                {
                    string csvPath = ManifestIO.ResolvePath(root);
                    if (File.Exists(csvPath))
                        foreach (var line in File.ReadAllLines(csvPath, System.Text.Encoding.UTF8))
                        {
                            if (string.IsNullOrEmpty(line)) continue;
                            var cols = line.Split(',');
                            if (cols.Length < 2) continue;
                            if (cols[0] == "entry_id" || cols[0] == "item_id") continue;
                            records.Add(cols);
                        }
                }
                _profileRecords[p] = records;
            }

            if (_cmbProfile.Items.Count > 0)
                _cmbProfile.SelectedIndex = 0;
        }

        // --- Profile selection ---

        void OnProfileChanged(object sender, SelectionChangedEventArgs e)
        {
            int sel = _cmbProfile.SelectedIndex;
            if (sel < 0) return;
            _selectedFolder = "";

            string type = "mail";
            if (_profileTypes.ContainsKey(sel)) type = _profileTypes[sel];
            SwitchLayout(type);

            RebuildActive();
            RebuildTree();
            ApplyFilter();
        }

        void RebuildActive()
        {
            _activeRecords.Clear();
            int sel = _cmbProfile.SelectedIndex;
            if (sel < 0) return;
            List<string[]> recs;
            if (_profileRecords.TryGetValue(sel, out recs))
                _activeRecords.AddRange(recs);
        }

        // --- Filter ---

        void ApplyFilter()
        {
            _itemList.Items.Clear();
            _currentRecords.Clear();
            ClearDetail();

            string q = (_txtSearch.Text ?? "").Trim().ToLower();

            foreach (var r in _activeRecords)
            {
                if (_selectedFolder.Length > 0 && _selectedFolder != ".")
                {
                    string fp = RecordFolderPath(r).TrimStart('\\');
                    if (_currentType != "mail")
                    {
                        int pidx = _cmbProfile.SelectedIndex;
                        string oRoot = pidx >= 0 ? Config.PGet(pidx, "output_root").TrimEnd('\\') : "";
                        if (!string.IsNullOrEmpty(oRoot) &&
                            fp.StartsWith(oRoot, StringComparison.OrdinalIgnoreCase))
                            fp = fp.Substring(oRoot.Length).TrimStart('\\');
                    }
                    if (!fp.StartsWith(_selectedFolder, StringComparison.OrdinalIgnoreCase))
                        continue;
                }
                if (q.Length > 0 && !RecordMatchesQuery(r, q))
                    continue;

                _currentRecords.Add(r);
                string date = RecordDate(r);
                if (date.Length >= 10) date = date.Substring(0, 10);
                _itemList.Items.Add(new[] { date, RecordName(r), RecordExtra(r) });
            }

            _statusText.Text = string.Format("{0} item(s)", _currentRecords.Count);
        }

        // --- Item selected ---

        void OnItemSelected(object sender, SelectionChangedEventArgs e)
        {
            int idx = _itemList.SelectedIndex;
            if (idx < 0 || idx >= _currentRecords.Count) return;
            var r = _currentRecords[idx];

            if (_currentType == "folder")
                ShowFolderFileInfo(r);
            else
                ShowMailDetail(r);
        }

        void ShowMailDetail(string[] r)
        {
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
            if (!string.IsNullOrEmpty(mailDir) && Directory.Exists(mailDir))
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

        void ShowFolderFileInfo(string[] r)
        {
            string fileName = r.Length > 1 ? r[1] : "";
            string filePath = r.Length > 2 ? r[2] : "";
            string relativePath = r.Length > 4 ? r[4] : "";
            string fileSize = r.Length > 5 ? FormatSize(r[5]) : "";
            string modifiedAt = r.Length > 6 ? r[6] : "";

            string created = "";
            try
            {
                if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
                    created = new FileInfo(filePath).CreationTime.ToString("yyyy-MM-dd HH:mm:ss");
            }
            catch { }

            _fileInfoBlock.Text = string.Format(
                "Name:      {0}\nPath:      {1}\nSize:      {2}\nModified:  {3}\nCreated:   {4}",
                fileName, relativePath, fileSize, modifiedAt, created);
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
            if (_headerBlock != null) _headerBlock.Text = "";
            if (_bodyBox != null) _bodyBox.Text = "";
            if (_attachList != null) _attachList.Items.Clear();
            if (_fileInfoBlock != null) _fileInfoBlock.Text = "";
        }
    }
}
