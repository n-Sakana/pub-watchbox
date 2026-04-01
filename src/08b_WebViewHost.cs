using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;

namespace WatchBox
{
    // Reusable base class for WebView2-hosted HTML windows.
    // Subclasses override HandleMessage to process JS -> C# actions.
    public class WebViewHost : Window
    {
        protected WebView2 _webView;
        protected bool _isReady;
        string _htmlFile;

        // Loading overlay (shown until page renders)
        Grid _rootGrid;
        Border _loadingOverlay;

        // Shared environment (pre-initialized at app startup for speed)
        static CoreWebView2Environment _sharedEnv;

        // Pre-create the environment so browser process is warm
        public static async System.Threading.Tasks.Task WarmUpAsync()
        {
            if (_sharedEnv != null) return;
            try
            {
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.LocalApplicationData),
                    "watchbox", "webview2");
                Directory.CreateDirectory(userDataFolder);
                _sharedEnv = await CoreWebView2Environment.CreateAsync(
                    null, userDataFolder);
            }
            catch { }
        }

        public WebViewHost(string title, string htmlFile,
            double width, double height)
        {
            Title = title;
            Width = width; Height = height;
            Background = new SolidColorBrush(Color.FromRgb(250, 250, 250));
            FontFamily = new FontFamily("Segoe UI");
            _htmlFile = htmlFile;

            // Root grid: WebView2 behind, loading overlay in front
            _rootGrid = new Grid();

            _webView = new WebView2();
            _webView.Visibility = Visibility.Hidden;
            _rootGrid.Children.Add(_webView);

            // Simple loading indicator
            _loadingOverlay = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(250, 250, 250))
            };
            var loadingText = new TextBlock
            {
                Text = "Loading...",
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
                FontSize = 13
            };
            _loadingOverlay.Child = loadingText;
            _rootGrid.Children.Add(_loadingOverlay);

            Content = _rootGrid;
            Loaded += OnLoaded;
        }

        async void OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_sharedEnv == null) await WarmUpAsync();
                await _webView.EnsureCoreWebView2Async(_sharedEnv);

                _webView.DefaultBackgroundColor
                    = System.Drawing.Color.FromArgb(255, 250, 250, 250);

                // Map web/ folder to https://app.local/
                string appDir = Config.Get("app_dir");
                string webFolder = Path.Combine(appDir, "web");
                _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "app.local", webFolder,
                    CoreWebView2HostResourceAccessKind.Allow);

                _webView.CoreWebView2.WebMessageReceived += OnMessageReceived;

                _webView.CoreWebView2.NewWindowRequested += (s2, e2) =>
                    e2.Handled = true;

                // Wait for page to render before showing
                _webView.CoreWebView2.NavigationCompleted += OnFirstNavComplete;

                _isReady = true;
                _webView.CoreWebView2.Navigate(
                    "https://app.local/" + _htmlFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show("WebView2 failed to initialize.\n\n" +
                    ex.Message, "watchbox", MessageBoxButton.OK,
                    MessageBoxImage.Error);
                Close();
            }
        }

        void OnFirstNavComplete(object sender,
            CoreWebView2NavigationCompletedEventArgs e)
        {
            _webView.CoreWebView2.NavigationCompleted -= OnFirstNavComplete;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                // Swap: hide loading, show WebView
                _webView.Visibility = Visibility.Visible;
                _rootGrid.Children.Remove(_loadingOverlay);
                _loadingOverlay = null;
            }));
        }

        void OnMessageReceived(object sender,
            CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string raw = e.WebMessageAsJson;
                string action = ExtractJsonString(raw, "action");
                if (!string.IsNullOrEmpty(action))
                    HandleMessage(action, raw);
            }
            catch { }
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (_webView != null)
            {
                _webView.Dispose();
                _webView = null;
            }
            base.OnClosing(e);
        }

        // Override in subclasses to handle specific actions
        protected virtual void HandleMessage(string action, string json) { }

        // Send a message to JavaScript
        protected void SendMessage(string action, string data)
        {
            if (!_isReady || _webView == null || _webView.CoreWebView2 == null) return;
            string msg = string.Format(
                "{{\"action\":\"{0}\",\"data\":{1}}}",
                JsonEsc(action), data ?? "null");
            try { _webView.CoreWebView2.PostWebMessageAsJson(msg); }
            catch { }
        }

        // --- Minimal JSON helpers (no external library) ---

        protected static string ExtractJsonString(string json, string key)
        {
            var m = Regex.Match(json,
                "\"" + Regex.Escape(key) + "\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
            return m.Success
                ? m.Groups[1].Value.Replace("\\\"", "\"").Replace("\\\\", "\\")
                : "";
        }

        protected static string ExtractJsonObject(string json, string key)
        {
            int idx = json.IndexOf("\"" + key + "\"");
            if (idx < 0) return "{}";
            idx = json.IndexOf(':', idx);
            if (idx < 0) return "{}";
            idx++;
            while (idx < json.Length && json[idx] == ' ') idx++;
            if (idx >= json.Length) return "{}";
            char open = json[idx];
            char close = open == '{' ? '}' : open == '[' ? ']' : '\0';
            if (close == '\0') return "{}";
            int depth = 1; int start = idx; idx++;
            bool inStr = false;
            while (idx < json.Length && depth > 0)
            {
                char c = json[idx];
                if (c == '\\' && inStr) { idx += 2; continue; }
                if (c == '"') inStr = !inStr;
                else if (!inStr && c == open) depth++;
                else if (!inStr && c == close) depth--;
                idx++;
            }
            return json.Substring(start, idx - start);
        }

        protected static string JsonEsc(string value)
        {
            return (value ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"")
                .Replace("\r\n", "\\n").Replace("\r", "\\n").Replace("\n", "\\n")
                .Replace("\t", "\\t");
        }

        protected static string JsonPair(string key, string value)
        {
            return string.Format("\"{0}\":\"{1}\"", key, JsonEsc(value ?? ""));
        }

        // Check if WebView2 runtime is available
        public static bool IsAvailable()
        {
            try
            {
                string ver = CoreWebView2Environment
                    .GetAvailableBrowserVersionString();
                return !string.IsNullOrEmpty(ver);
            }
            catch { return false; }
        }
    }
}
