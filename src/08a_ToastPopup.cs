using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace WatchBox
{
    // Lightweight toast-style popup in bottom-right corner.
    // Shows for a few seconds then fades out.
    public class ToastPopup : Window
    {
        DispatcherTimer _timer;

        public ToastPopup(string title, string message)
        {
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            ShowInTaskbar = false;
            Topmost = true;
            ResizeMode = ResizeMode.NoResize;
            Width = 320; Height = 90;
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = SystemParameters.WorkArea.Right - 330;
            Top = SystemParameters.WorkArea.Bottom - 100;

            var border = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(55, 120, 200)),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(16, 12, 16, 12),
                Margin = new Thickness(4),
                Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    BlurRadius = 12, ShadowDepth = 2, Opacity = 0.3,
                    Color = Colors.Black
                }
            };

            var panel = new StackPanel();

            var titleBlock = new TextBlock
            {
                Text = title,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            panel.Children.Add(titleBlock);

            var msgBlock = new TextBlock
            {
                Text = message,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(200, 200, 200)),
                TextTrimming = TextTrimming.CharacterEllipsis,
                Margin = new Thickness(0, 4, 0, 0)
            };
            panel.Children.Add(msgBlock);

            border.Child = panel;
            Content = border;

            MouseLeftButtonDown += (s, e) => Close();

            Loaded += OnLoaded;
        }

        void OnLoaded(object sender, RoutedEventArgs e)
        {
            // Slide in from right
            var slideIn = new DoubleAnimation
            {
                From = Left + 50, To = Left,
                Duration = TimeSpan.FromMilliseconds(200)
            };
            BeginAnimation(LeftProperty, slideIn);

            // Auto-close after 4 seconds
            _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(4) };
            _timer.Tick += (s2, e2) =>
            {
                _timer.Stop();
                FadeAndClose();
            };
            _timer.Start();
        }

        void FadeAndClose()
        {
            var fadeOut = new DoubleAnimation
            {
                From = 1, To = 0,
                Duration = TimeSpan.FromMilliseconds(300)
            };
            fadeOut.Completed += (s, e) => Close();
            BeginAnimation(OpacityProperty, fadeOut);
        }

        // Static helper: show toast on UI thread
        public static void Show(string title, string message)
        {
            try
            {
                var popup = new ToastPopup(title, message);
                popup.Show();
            }
            catch { }
        }
    }
}
