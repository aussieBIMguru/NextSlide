using System.Windows;
using NextSlide.Models;
using NextSlide.ViewModels;

namespace NextSlide.Views;

public partial class MainWindow : Window
{
    public MainWindow(MainViewModel viewModel, AppSettings settings)
    {
        InitializeComponent();
        DataContext = viewModel;
        RestoreWindowState(settings);
    }

    private void RestoreWindowState(AppSettings settings)
    {
        if (!double.IsNaN(settings.WindowLeft) && !double.IsNaN(settings.WindowTop))
        {
            var left = settings.WindowLeft;
            var top = settings.WindowTop;

            // Guard against restoring a position from a monitor that's no
            // longer connected — fall back to WPF's default startup
            // location (see WindowStartupLocation in MainWindow.xaml)
            // instead of placing the window off-screen.
            var withinVirtualScreen =
                left >= SystemParameters.VirtualScreenLeft &&
                left < SystemParameters.VirtualScreenLeft + SystemParameters.VirtualScreenWidth &&
                top >= SystemParameters.VirtualScreenTop &&
                top < SystemParameters.VirtualScreenTop + SystemParameters.VirtualScreenHeight;

            if (withinVirtualScreen)
            {
                Left = left;
                Top = top;
            }
        }

        if (settings.WindowWidth > 0)
            Width = settings.WindowWidth;
        if (settings.WindowHeight > 0)
            Height = settings.WindowHeight;

        if (settings.WindowMaximized)
            WindowState = WindowState.Maximized;
    }

    /// <summary>
    /// Copies the current window bounds (and, via the view model,
    /// non-window state) back into <paramref name="settings"/> so
    /// App.xaml.cs can persist them on exit. The template's default save
    /// policy is "explicit save on clean exit" (see SettingsService's
    /// ScheduleAutosave for the opt-in debounced alternative), so this is
    /// called once from App.xaml.cs rather than on every LocationChanged /
    /// SizeChanged event.
    /// </summary>
    public void PersistWindowState(AppSettings settings)
    {
        // Capture the restore bounds, not the maximized bounds, so
        // un-maximizing on the next launch doesn't leave the window
        // full-screen with nowhere sensible to shrink back to.
        if (WindowState == WindowState.Normal)
        {
            settings.WindowLeft = Left;
            settings.WindowTop = Top;
            settings.WindowWidth = Width;
            settings.WindowHeight = Height;
        }

        settings.WindowMaximized = WindowState == WindowState.Maximized;

        if (DataContext is MainViewModel viewModel)
            viewModel.PersistToSettings();
    }
}
