namespace NextSlide;

/// <summary>
/// Single source of truth for the app's display name, tray monogram, and
/// AppData publisher folder. The window title, the tray icon tooltip and
/// glyph, and the settings file path (%LOCALAPPDATA%\Gavin\NextSlide\) all
/// read from here.
/// </summary>
public static class AppInfo
{
    public const string Name = "NextSlide";
    public const string Monogram = "NS";
    public const string Publisher = "Gavin";
}
