using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using Vanara.PInvoke;
using static Vanara.PInvoke.User32;

namespace Switcheroo
{
    public partial class MainWindow : Window
    {
        private const int MAX_WIDTH = 1200;
        private const int MAX_WINDOW_HEIGHT = 150;
        private const int REFRESH_INTERVAL = 10;

        // Global hotkey
        private const int HOTKEY_ID = 0xB001;
        private const int WM_HOTKEY = 0x0312;
        private const int MOD_NOREPEAT = 0x4000; // avoids key repeat storms
        private const int VK_OEM_PERIOD = 0xBE;  // '.' key

        private readonly List<HWND> _selectedWindows = new();
        private readonly Dictionary<Border, HTHUMBNAIL> _thumbnailHandles = new();
        private int _selectedWindowIndex = -1;
        private Border? _lastCycledContainer = null;
        private bool _selectionUIVisible = true;
        private HwndSource? _hwndSrc;
        private DispatcherTimer _refreshTimer;

        // --- Begin interop for blur ---
        [StructLayout(LayoutKind.Sequential)]
        private struct AccentPolicy
        {
            public int AccentState;
            public int AccentFlags;
            public int GradientColor;
            public int AnimationId;
        }

        private enum AccentState
        {
            ACCENT_DISABLED = 0,
            ACCENT_ENABLE_GRADIENT = 1,
            ACCENT_ENABLE_TRANSPARENTGRADIENT = 2,
            ACCENT_ENABLE_BLURBEHIND = 3,
            ACCENT_ENABLE_ACRYLICBLURBEHIND = 4,
            ACCENT_INVALID_STATE = 5
        }

        private enum WindowCompositionAttribute
        {

            WCA_ACCENT_POLICY = 19
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct WindowCompositionAttributeData
        {
            public WindowCompositionAttribute Attribute;
            public IntPtr Data;
            public int SizeOfData;
        }

        public MainWindow()
        {
            InitializeComponent();

            Topmost = true;
            StateChanged += MainWindow_StateChanged;
            Deactivated += MainWindow_Deactivated;

            InitializeRefreshTimer();

            LoadWindows();
            UpdateSelectedCount();
        }

        // Hook WndProc and register hotkey once HWND exists
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            _hwndSrc = (HwndSource)PresentationSource.FromVisual(this)!;
            _hwndSrc.AddHook(WndProc);

            // Register Ctrl + . as a global hotkey
            RegisterHotKey(
                _hwndSrc.Handle,
                HOTKEY_ID,
                HotKeyModifiers.MOD_CONTROL | HotKeyModifiers.MOD_NOREPEAT,
                (uint)VK_OEM_PERIOD
            );


        }

        protected override void OnClosed(EventArgs e)
        {
            _refreshTimer?.Stop();
            try
            {
                if (_hwndSrc != null)
                {
                    UnregisterHotKey(_hwndSrc.Handle, HOTKEY_ID);
                    _hwndSrc.RemoveHook(WndProc);
                }
            }
            catch { /* ignore */ }

            base.OnClosed(e);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_HOTKEY && wParam.ToInt32() == HOTKEY_ID)
            {
                // If selector is visible, hide & start cycling; if hidden, cycle next.
                if (_selectionUIVisible)
                {
                    BeginCyclingAndHide();
                }
                else
                {
                    if (_selectedWindows.Count == 0)
                    {
                        // Nothing selected -> show selector so user can pick
                        ShowSelector();
                    }
                    else
                    {
                        CycleSelectedWindows();
                    }
                }

                handled = true;
            }
            return IntPtr.Zero;
        }

        private void MainWindow_StateChanged(object? sender, EventArgs e)
        {
            if (WindowState == WindowState.Normal)
            {
                Topmost = true;
                _refreshTimer.Start(); // Start timer when visible
            }

            else if (WindowState == WindowState.Minimized)
            {
                _refreshTimer.Stop(); // Stop timer when minimized
            }
        }

        private void MainWindow_Deactivated(object? sender, EventArgs e)
        {
            // When we’re in “cycling” mode (selector hidden), keep UI hidden if we lose focus
            if (!_selectionUIVisible)
                Hide();
        }

        // UI actions
        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            BeginCyclingAndHide();
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            LoadWindows();
        }

        private void CleanupThumbnails()
        {
            // Unregister all active DWM thumbnails
            foreach (var thumbHandle in _thumbnailHandles.Values)
            {
                try
                {
                    DwmApi.DwmUnregisterThumbnail(thumbHandle);
                }
                catch
                {
                    // Ignore errors, window might be closed
                }
            }

            // Clear our tracking and the WPF panel
            _thumbnailHandles.Clear();
            WindowPanel.Children.Clear();
        }

        private void BeginCyclingAndHide()
        {
            //// If no selection, keep selector open
            //if (_selectedWindows.Count == 0)
            //{
            //    MessageBox.Show("Select one or more windows first.", "Switcheroo", MessageBoxButton.OK, MessageBoxImage.Information);
            //    return;
            //}

            _selectionUIVisible = false;
            // Minimize + hide so WM_HOTKEY still arrives
            WindowState = WindowState.Minimized;
            _refreshTimer.Stop();
            Hide();

            // Immediately activate first (or next) window on first invocation
            CycleSelectedWindows();
        }

        private void InitializeRefreshTimer()
        {
            _refreshTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(REFRESH_INTERVAL)
            };

            // When the timer ticks, just call your existing LoadWindows method
            _refreshTimer.Tick += (s, e) => LoadWindows();

            _refreshTimer.Start(); // Start the timer immediately
        }

        private void ShowSelector()
        {
            _selectionUIVisible = true;
            Show();
            WindowState = WindowState.Normal;
            Activate();
            Topmost = true;
        }

        private void LoadWindows()
        {
            CleanupThumbnails();

            WindowPanel.Children.Clear();
            //_selectedWindows.Clear();
            //_selectedWindowIndex = -1;
            _lastCycledContainer = null;

            int total = CountWindows();
            if (total == 0)
                return;

            EnumWindows((hwnd, _) =>
            {
                if (IsCandidateWindow(hwnd))
                {
                    string title = GetWindowTitle(hwnd);
                    if (!string.IsNullOrWhiteSpace(title))
                    {
                        AddThumbnail(hwnd, title, total);
                    }
                }
                return true;
            }, IntPtr.Zero);

            _selectionUIVisible = true;
            ShowSelector();
            UpdateSelectedCount();
        }

        private int CountWindows()
        {
            int count = 0;
            EnumWindows((hwnd, _) =>
            {
                if (IsCandidateWindow(hwnd))
                    count++;
                return true;
            }, IntPtr.Zero);
            return count;
        }

        private static string GetWindowTitle(HWND hwnd)
        {
            int len = GetWindowTextLength(hwnd);
            if (len <= 0) return string.Empty;

            var sb = new StringBuilder(len + 1);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private bool IsCandidateWindow(HWND hwnd)
        {
            if (_hwndSrc != null && hwnd == _hwndSrc.Handle) return false;
            if (!IsWindowVisible(hwnd)) return false;
            if (GetWindowTextLength(hwnd) <= 0) return false;
            if (IsToolWindow(hwnd)) return false;
            if (IsCloaked(hwnd)) return false;
            return true;
        }

        private static bool IsToolWindow(HWND hwnd)
        {
            long ex = GetWindowLongPtr(hwnd, WindowLongFlags.GWL_EXSTYLE).ToInt64();
            return (ex & (long)WindowStylesEx.WS_EX_TOOLWINDOW) != 0;
        }

        private static bool IsCloaked(HWND hwnd)
        {
            // DWM cloaking check
            uint cloaked = 0;
            IntPtr ptr = Marshal.AllocHGlobal(sizeof(uint));
            try
            {
                Marshal.WriteInt32(ptr, 0);
                DwmApi.DwmGetWindowAttribute(hwnd,
                    DwmApi.DWMWINDOWATTRIBUTE.DWMWA_CLOAKED,
                    ptr,
                    sizeof(uint));
                cloaked = (uint)Marshal.ReadInt32(ptr);
                return cloaked != 0;
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }
        }

        private void AddThumbnail(HWND hwnd, string title, int total)
        {
            int maxWindowWidth = (MAX_WIDTH / total) - 40;
            GetWindowRect(hwnd, out RECT rect);

            int width = rect.right - rect.left;
            int height = rect.bottom - rect.top;

            float aspectRatio = (float)width / height;

            // calculate width and height based on aspect ratio
            if (aspectRatio >= 1.0f)
            {
                // Wider than tall
                width = Math.Min(maxWindowWidth, width);
                height = Math.Min(MAX_WINDOW_HEIGHT, (int)(width / aspectRatio));
            }
            else
            {
                // Taller than wide
                height = 150; // fixed height
                width = (int)(height * aspectRatio);
                if (width > maxWindowWidth)
                {
                    width = maxWindowWidth;
                    height = (int)(width / aspectRatio);
                }
            }

            SolidColorBrush borderBrush = _selectedWindows.Contains(hwnd) ? Brushes.LimeGreen : Brushes.Gray;
            //int maxWindowWidth = Math.Max(180, (MAX_WIDTH / Math.Max(1, total)) - 40);

            

            var container = new Border
            {
                Width = width,
                Height = height,
                Margin = new Thickness(8),
                BorderBrush = borderBrush,
                BorderThickness = new Thickness(4),
                Background = Brushes.Black,
                Cursor = Cursors.Hand,
                Tag = hwnd
            };
            container.MouseLeftButtonUp += Thumbnail_MouseLeftButtonUp;

            var grid = new Grid();
            container.Child = grid;
            WindowPanel.Children.Add(container);

            // Label only (DWM thumbnail will appear on top — if you need text on top of preview,
            // switch to a captured snapshot approach instead of DWM thumbnails)
            //var overlayText = new TextBlock
            //{
            //    Text = title,
            //    Foreground = Brushes.White,
            //    TextWrapping = TextWrapping.Wrap,
            //    HorizontalAlignment = HorizontalAlignment.Center,
            //    VerticalAlignment = VerticalAlignment.Center,
            //    Margin = new Thickness(6),
            //    Opacity = 0.85
            //};
            //grid.Children.Add(overlayText);

            container.Loaded += (_, __) =>
            {
                var hwndSrc = (HwndSource)PresentationSource.FromVisual(this)!;
                if (hwndSrc == null) return;



                if (DwmApi.DwmRegisterThumbnail(hwndSrc.Handle, hwnd, out var thumb) == 0)
                {
                    _thumbnailHandles[container] = thumb;
                    double padding = 6;
                    double width = container.ActualWidth - 2 * padding;
                    double height = container.ActualHeight - 2 * padding;

                    var relative = container.TransformToAncestor(this).Transform(new Point(0, 0));
                    double left = relative.X + (container.ActualWidth - width) / 2;
                    double top = relative.Y + (container.ActualHeight - height) / 2;

                    var props = new DwmApi.DWM_THUMBNAIL_PROPERTIES
                    {
                        dwFlags = DwmApi.DWM_TNP.DWM_TNP_VISIBLE |
                                  DwmApi.DWM_TNP.DWM_TNP_RECTDESTINATION |
                                  DwmApi.DWM_TNP.DWM_TNP_SOURCECLIENTAREAONLY,
                        fVisible = true,
                        fSourceClientAreaOnly = false,
                        rcDestination = new RECT((int)left, (int)top, (int)(left + width), (int)(top + height))
                    };

                    DwmApi.DwmUpdateThumbnailProperties(thumb, props);
                }
            };
        }

        private void Thumbnail_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is not Border container) return;
            var hwnd = (HWND)container.Tag;

            if (_selectedWindows.Contains(hwnd))
            {
                _selectedWindows.Remove(hwnd);
                container.BorderBrush = Brushes.Gray;
            }
            else
            {
                _selectedWindows.Add(hwnd);
                container.BorderBrush = Brushes.LimeGreen;
            }

            // Reset cycling index if the selection set changes
            _selectedWindowIndex = -1;
            UpdateSelectedCount();
        }

        private void UpdateSelectedCount()
        {
            if (SelectedCount != null)
                SelectedCount.Text = _selectedWindows.Count == 0
                    ? "No windows selected."
                    : $"{_selectedWindows.Count} selected.";
        }

        // ===== Cycling logic =====

        private void CycleSelectedWindows()
        {
            if (_selectedWindows.Count == 0)
            {
                ShowSelector();
                return;
            }

            // Restore previous cycled container’s brush
            if (_lastCycledContainer != null)
            {
                if (_lastCycledContainer.Tag is HWND prevHwnd)
                {
                    _lastCycledContainer.BorderBrush = _selectedWindows.Contains(prevHwnd) ? Brushes.LimeGreen : Brushes.Gray;
                }
                else
                {
                    _lastCycledContainer.BorderBrush = Brushes.Gray;
                }
                _lastCycledContainer = null;
            }

            // Advance index
            _selectedWindowIndex++;
            if (_selectedWindowIndex >= _selectedWindows.Count)
                _selectedWindowIndex = 0;

            // Clean up handles that may have closed
            int safety = 0;
            while (_selectedWindowIndex < _selectedWindows.Count && (!IsWindow(_selectedWindows[_selectedWindowIndex]) || !IsWindowVisible(_selectedWindows[_selectedWindowIndex])))
            {
                _selectedWindows.RemoveAt(_selectedWindowIndex);
                if (_selectedWindows.Count == 0)
                {
                    ShowSelector();
                    return;
                }
                if (++safety > 100) break;
            }

            if (_selectedWindows.Count == 0)
            {
                ShowSelector();
                return;
            }

            var hwndToActivate = _selectedWindows[_selectedWindowIndex];

            if (IsIconic(hwndToActivate))
                ShowWindow(hwndToActivate, ShowWindowCommand.SW_RESTORE);

            // Temporarily drop Topmost so target window can come to front reliably
            Topmost = false;
            SetForegroundWindow(hwndToActivate);
            // Bring to top as a fallback
            BringWindowToTop(hwndToActivate);
            // Immediately hide selector again (Alt-Tab feel)
            Hide();
            Topmost = true;

            // Update highlight (useful when selector is visible again)
            foreach (var child in WindowPanel.Children.OfType<Border>())
            {
                if (child.Tag is HWND h && h == hwndToActivate)
                {
                    child.BorderBrush = Brushes.DodgerBlue;
                    _lastCycledContainer = child;
                    break;
                }
            }
        }
    }
}
