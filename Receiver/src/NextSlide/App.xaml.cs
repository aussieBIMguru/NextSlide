using System;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using NextSlide.Models;
using NextSlide.Services;
using NextSlide.ViewModels;
using NextSlide.Views;

namespace NextSlide;

/// <summary>
/// Startup/shutdown orchestration for all three RunModes. ShutdownMode is
/// set to OnExplicitShutdown in App.xaml because Silent mode has no window
/// to trigger the default "last window closed" shutdown — every exit path
/// here ends by calling Shutdown() itself.
/// </summary>
public partial class App : Application
{
    private SettingsService? _settingsService;
    private AppSettings? _settings;
    private TrayIconService? _trayIconService;
    private MainViewModel? _mainViewModel;
    private MainWindow? _mainWindow;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        _settingsService = new SettingsService();
        _settings = _settingsService.Load();

        // Command-line override: `--silent` forces silent/tray-only mode
        // for this launch regardless of the saved RunMode — useful for a
        // Task Scheduler entry or a Startup-folder shortcut that shouldn't
        // depend on whatever the user last left RunMode set to.
        var forceSilent = e.Args.Any(arg => string.Equals(arg, "--silent", StringComparison.OrdinalIgnoreCase));
        var effectiveRunMode = forceSilent ? RunMode.Silent : _settings.RunMode;

        _mainViewModel = new MainViewModel(_settings, _settingsService);

        // The tray icon is always created regardless of RunMode — only
        // whether the window shows on launch, and what closing it does,
        // changes between modes. MainViewModel's poll loop (SlidePollingService)
        // is independent of window visibility too: hiding to tray via
        // WindowedTrayOnClose (NextSlide's default — see AppSettings) never
        // stops it, which is the whole point of running in the tray.
        _trayIconService = new TrayIconService(monogram: AppInfo.Monogram, tooltipText: AppInfo.Name);
        _trayIconService.ShowRequested += (_, _) => ShowMainWindow();
        _trayIconService.ExitRequested += (_, _) => ExitApplication();

        if (effectiveRunMode != RunMode.Silent)
            ShowMainWindow();
    }

    private void ShowMainWindow()
    {
        if (_mainWindow is null)
        {
            _mainWindow = new MainWindow(_mainViewModel!, _settings!);
            _mainWindow.Closing += MainWindow_Closing;
        }

        _mainWindow.Show();

        if (_mainWindow.WindowState == WindowState.Minimized)
            _mainWindow.WindowState = WindowState.Normal;

        _mainWindow.Activate();
    }

    private void MainWindow_Closing(object? sender, CancelEventArgs e)
    {
        if (_settings!.RunMode == RunMode.WindowedTrayOnClose)
        {
            // Hide to tray instead of closing. The window instance stays
            // alive and is reshown by TrayIconService's Show/double-click.
            e.Cancel = true;
            _mainWindow!.Hide();
            return;
        }

        // WindowedExitOnClose — a real close means exit the whole app.
        ExitApplication();
    }

    private void ExitApplication()
    {
        if (_mainWindow is not null)
        {
            _mainWindow.Closing -= MainWindow_Closing;
            _mainWindow.PersistWindowState(_settings!);
        }

        _settingsService?.Save(_settings!);
        _mainViewModel?.Dispose();
        _trayIconService?.Dispose();
        Shutdown();
    }
}
