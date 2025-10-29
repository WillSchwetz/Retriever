using System; // MODIFIED: Added for EventArgs
using System.Collections.Generic; // MODIFIED: Added for List<>
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Vanara.PInvoke;

namespace Switcheroo;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    // MODIFIED: List to hold selected windows and index for cycling
    private List<HWND> _selectedWindows = new List<HWND>();
    private int _selectedWindowIndex = -1;

    public MainWindow()
    {
        InitializeComponent();

        // MODIFIED: Make the window always on top
        this.Topmost = true;

        // MODIFIED: Add handlers for key down and state changed
        this.KeyDown += MainWindow_KeyDown;
        this.StateChanged += MainWindow_StateChanged;

        LoadWindows();
    }

    // MODIFIED: Handle window restore
    private void MainWindow_StateChanged(object sender, EventArgs e)
    {
        // When the window is restored, make sure it's topmost
        if (this.WindowState == WindowState.Normal)
        {
            this.Topmost = true;
        }
    }

    // MODIFIED: Handle the keydown event for hotkeys
    private void MainWindow_KeyDown(object sender, KeyEventArgs e)
    {
        // Check for Control + . (OemPeriod is the '.' key)
        if (e.Key == Key.OemPeriod && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
        {
            CycleSelectedWindows();
        }
    }

    // MODIFIED: Logic to cycle through the selected windows
    private void CycleSelectedWindows()
    {
        if (_selectedWindows.Count == 0)
        {
            // Nothing to cycle
            return;
        }

        // Increment and wrap the index
        _selectedWindowIndex++;
        if (_selectedWindowIndex >= _selectedWindows.Count)
        {
            _selectedWindowIndex = 0;
        }

        // Get the window handle to activate
        HWND hwndToActivate = _selectedWindows[_selectedWindowIndex];

        // Activate the target window
        User32.SetForegroundWindow(hwndToActivate);

        this.Activate();

        //// Minimize this app so you can see the window
        //this.WindowState = WindowState.Minimized;
    }

    private void LoadWindows()
    {
        // MODIFIED: Clear old windows and selection before loading new ones
        WindowPanel.Children.Clear();
        _selectedWindows.Clear();
        _selectedWindowIndex = -1;

        User32.EnumWindows((hwnd, _) =>
        {
            if (User32.IsWindowVisible(hwnd) &&
                User32.GetWindowTextLength(hwnd) > 0)
            {
                int length = User32.GetWindowTextLength(hwnd);
                var sb = new System.Text.StringBuilder(length + 1);
                User32.GetWindowText(hwnd, sb, sb.Capacity);
                string title = sb.ToString();

                AddThumbnail(hwnd, title);
            }
            return true;
        }, IntPtr.Zero);
    }

    // private Border _selectedThumbnail = null; // MODIFIED: Removed unused variable

    private void AddThumbnail(Vanara.PInvoke.HWND hwnd, string title)
    {
        // Skip minimized windows
        if (User32.IsIconic(hwnd))
            return;

        // Allocate space for a DWORD
        uint cloaked = 0;
        IntPtr ptr = Marshal.AllocHGlobal(Marshal.SizeOf(cloaked));
        Marshal.WriteInt32(ptr, 0);

        // Call DwmGetWindowAttribute
        var hr = DwmApi.DwmGetWindowAttribute(hwnd,
            DwmApi.DWMWINDOWATTRIBUTE.DWMWA_CLOAKED,
            ptr,
            Marshal.SizeOf(typeof(uint)));

        if (hr == 0)
        {
            cloaked = (uint)Marshal.ReadInt32(ptr);
            if (cloaked != 0)
            {
                Marshal.FreeHGlobal(ptr);
                return; // Skip cloaked window
            }
        }

        Marshal.FreeHGlobal(ptr);

        // Create container for thumbnail
        var container = new Border
        {
            Width = 300,
            Height = 200,
            Margin = new Thickness(10),
            BorderBrush = System.Windows.Media.Brushes.Gray, // Default border
            BorderThickness = new Thickness(10),
            Background = System.Windows.Media.Brushes.Black,
            Cursor = System.Windows.Input.Cursors.Hand
        };

        // MODIFIED: Store the HWND in the Tag property for easy access
        container.Tag = hwnd;

        // MODIFIED: Attach the click handler
        container.MouseLeftButtonUp += Thumbnail_MouseLeftButtonUp;

        var grid = new Grid();
        container.Child = grid;
        WindowPanel.Children.Add(container);

        var overlayText = new TextBlock
        {
            Text = title,
            Foreground = System.Windows.Media.Brushes.White,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            FontWeight = FontWeights.Bold,
            Margin = new Thickness(5)
        };

        grid.Children.Add(overlayText);

        var subtitleText = new TextBlock
        {
            Text = "Subtitle goes here",
            Foreground = Brushes.LightGray,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(5),
            FontStyle = FontStyles.Italic
        };

        grid.Children.Add(subtitleText);


        container.Loaded += (_, __) =>
        {
            var hwndSrc = (HwndSource)PresentationSource.FromVisual(this);
            if (hwndSrc == null) return;

            if (DwmApi.DwmRegisterThumbnail(hwndSrc.Handle, hwnd, out var thumb) == 0)
            {
                double padding = 10;

                double width = container.ActualWidth - 2 * padding;
                double height = container.ActualHeight - 2 * padding;

                var relativePoint = container.TransformToAncestor(this).Transform(new Point(0, 0));

                // Center the thumbnail
                double left = relativePoint.X + (container.ActualWidth - width) / 2;
                double top = relativePoint.Y + (container.ActualHeight - height) / 2;

                var props = new DwmApi.DWM_THUMBNAIL_PROPERTIES
                {
                    dwFlags = DwmApi.DWM_TNP.DWM_TNP_VISIBLE |
                                DwmApi.DWM_TNP.DWM_TNP_RECTDESTINATION |
                                DwmApi.DWM_TNP.DWM_TNP_SOURCECLIENTAREAONLY,
                    fVisible = true,
                    fSourceClientAreaOnly = false,
                    rcDestination = new RECT(
                        (int)left,
                        (int)top,
                        (int)(left + width),
                        (int)(top + height)
                    )
                };

                DwmApi.DwmUpdateThumbnailProperties(thumb, props);

                // MODIFIED: Removed the original click handler from here
                // as we added a more advanced one above.
            }
        };
    }

    // MODIFIED: New method to handle thumbnail clicks
    private void Thumbnail_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        var container = sender as Border;
        if (container == null) return;

        // Get the HWND we stored in the Tag
        var hwnd = (HWND)container.Tag;

        if (_selectedWindows.Contains(hwnd))
        {
            // Already selected, so DE-SELECT it
            _selectedWindows.Remove(hwnd);
            container.BorderBrush = Brushes.Gray; // Set border back to default
            _selectedWindowIndex = -1; // Reset cycle index
        }
        else
        {
            // Not selected, so SELECT it
            _selectedWindows.Add(hwnd);
            container.BorderBrush = Brushes.LimeGreen; // Set border to bright green
        }
    }
}