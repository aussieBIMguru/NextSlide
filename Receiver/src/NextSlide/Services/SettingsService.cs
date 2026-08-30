using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using NextSlide.Models;

namespace NextSlide.Services;

/// <summary>
/// Loads and saves AppSettings as JSON under the user's AppData folder.
/// Defaults to explicit save (call Save/SaveAsync yourself — App.xaml.cs
/// does this on clean exit); ScheduleAutosave below is the opt-in
/// alternative if a derived app wants debounce-on-change instead.
/// </summary>
public sealed class SettingsService
{
    /// <summary>
    /// false (default) = %LOCALAPPDATA%\{Publisher}\{Name}\settings.json
    ///   (machine-specific — the template default, since most system-app
    ///   settings like window position are meaningless on another machine).
    /// true  = %APPDATA%\{Publisher}\{Name}\settings.json (roams with the
    ///   user profile on a domain/roaming-profile setup).
    /// </summary>
    private const bool UseRoamingAppData = false;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },

        // AppSettings.WindowLeft/WindowTop deliberately default to
        // double.NaN as a "no saved position yet" sentinel (see their doc
        // comments) — plain JSON has no representation for NaN, so without
        // this, any Save() call made before a window position is ever
        // captured (e.g. on first run, before the window has been closed
        // once) throws ArgumentException instead of writing the file. This
        // makes System.Text.Json write/read it as the literal NaN instead.
        NumberHandling = JsonNumberHandling.AllowNamedFloatingPointLiterals
    };

    private readonly string _settingsFilePath;
    private readonly object _autosaveLock = new();
    private Timer? _autosaveTimer;
    private AppSettings? _autosaveTarget;

    public SettingsService()
    {
        var root = Environment.GetFolderPath(UseRoamingAppData
            ? Environment.SpecialFolder.ApplicationData
            : Environment.SpecialFolder.LocalApplicationData);

        var folder = Path.Combine(root, AppInfo.Publisher, AppInfo.Name);
        Directory.CreateDirectory(folder);
        _settingsFilePath = Path.Combine(folder, "settings.json");
    }

    /// <summary>Full path to settings.json — handy for a "Reveal in Explorer" menu item.</summary>
    public string SettingsFilePath => _settingsFilePath;

    public AppSettings Load()
    {
        try
        {
            if (!File.Exists(_settingsFilePath))
                return new AppSettings();

            var json = File.ReadAllText(_settingsFilePath);
            var settings = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions);

            // TEMPLATE: if settings?.SchemaVersion is older than the current
            // AppSettings.SchemaVersion, migrate old field shapes here
            // before returning. Nothing to migrate yet at schema version 1.

            return settings ?? new AppSettings();
        }
        catch
        {
            // Corrupt or unreadable settings file — start fresh rather than
            // crashing the app on launch.
            return new AppSettings();
        }
    }

    public void Save(AppSettings settings)
    {
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        File.WriteAllText(_settingsFilePath, json);
    }

    public async Task SaveAsync(AppSettings settings, CancellationToken cancellationToken = default)
    {
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        await File.WriteAllTextAsync(_settingsFilePath, json, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// TEMPLATE EXAMPLE — opt-in autosave. Call this (e.g. from a
    /// property-changed handler) to debounce-save <paramref name="settings"/>
    /// a short time after the last change, instead of relying on an
    /// explicit Save. Not wired up by default — see README.md "Settings".
    /// </summary>
    public void ScheduleAutosave(AppSettings settings, TimeSpan? debounce = null)
    {
        lock (_autosaveLock)
        {
            _autosaveTarget = settings;
            var delay = debounce ?? TimeSpan.FromSeconds(1.5);

            _autosaveTimer ??= new Timer(_ =>
            {
                AppSettings? toSave;
                lock (_autosaveLock)
                {
                    toSave = _autosaveTarget;
                }

                if (toSave is not null)
                    Save(toSave);
            });

            _autosaveTimer.Change(delay, Timeout.InfiniteTimeSpan);
        }
    }
}
