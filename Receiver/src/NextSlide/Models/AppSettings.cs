namespace NextSlide.Models;

/// <summary>
/// Persisted application settings, serialized as JSON by SettingsService.
/// Bump <see cref="SchemaVersion"/> whenever a field is added, removed, or
/// its meaning changes, and add a migration note (or code, in
/// SettingsService.Load) if older settings files on disk need upgrading.
/// </summary>
public sealed class AppSettings
{
    /// <summary>
    /// v2 (2026-08-30): dropped the base template's DummyItemCount example
    /// field and added LastSessionName/LastSheetUrl. No migration code
    /// needed — System.Text.Json just leaves the two new fields at their
    /// defaults when loading a v1 file, and silently drops the removed
    /// one, which is harmless here (nothing derives from it).
    /// </summary>
    public int SchemaVersion { get; set; } = 2;

    // The template default is WindowedExitOnClose; NextSlide is meant to
    // keep polling in the background once set up, so it defaults to
    // close-to-tray instead.
    public RunMode RunMode { get; set; } = RunMode.WindowedTrayOnClose;

    // Last known main window bounds, used to restore the window on the next
    // launch. Left/Top of double.NaN means "no saved position yet" (first
    // run) — MainWindow falls back to WPF's centered startup location.
    public double WindowLeft { get; set; } = double.NaN;
    public double WindowTop { get; set; } = double.NaN;
    public double WindowWidth { get; set; } = 1000;
    public double WindowHeight { get; set; } = 650;
    public bool WindowMaximized { get; set; }

    // Pre-fills the Session/Sheet URL fields on next launch so the user
    // doesn't have to retype them — but never auto-locks on startup, so a
    // stale/wrong value from a previous run can't silently start polling
    // and driving PowerPoint before the user has looked at it.
    public string? LastSessionName { get; set; }
    public string? LastSheetUrl { get; set; }
}
