using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;

namespace NextSlide.Services;

/// <summary>
/// Wraps a WinForms NotifyIcon so the tray icon works the same way
/// regardless of RunMode — it is always created (see App.xaml.cs), only
/// what happens on Show/Exit and window-close differs between modes.
///
/// This file is the one deliberately WinForms/System.Drawing file in the
/// project (see NextSlide.csproj's ImplicitUsings comment) — it has no
/// `using System.Windows;`, so there is no ambiguity with WPF's own
/// Application/Color/Rectangle-named types.
/// </summary>
public sealed class TrayIconService : IDisposable
{
    private readonly NotifyIcon _notifyIcon;

    public event EventHandler? ShowRequested;
    public event EventHandler? ExitRequested;

    public TrayIconService(string monogram, string tooltipText)
    {
        var menu = new ContextMenuStrip();
        menu.Items.Add("Show", null, (_, _) => ShowRequested?.Invoke(this, EventArgs.Empty));
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Exit", null, (_, _) => ExitRequested?.Invoke(this, EventArgs.Empty));

        _notifyIcon = new NotifyIcon
        {
            Icon = CreateMonogramIcon(monogram),
            Text = Truncate(tooltipText, 63), // NotifyIcon.Text is capped at 63 chars.
            Visible = true,
            ContextMenuStrip = menu
        };

        _notifyIcon.DoubleClick += (_, _) => ShowRequested?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Renders a small rounded square with a two-letter monogram to an
    /// in-memory bitmap and wraps it as an Icon, so the template needs no
    /// .ico asset. TEMPLATE NOTE: baking a real .ico at design time
    /// (Resources\app.ico, wired via &lt;ApplicationIcon&gt; in the csproj)
    /// is the other valid approach, and is simpler once the app has real
    /// branding — swap to that when this stops being a placeholder.
    /// </summary>
    public static Icon CreateMonogramIcon(string monogram, int size = 32)
    {
        using var bitmap = new Bitmap(size, size);
        using var g = Graphics.FromImage(bitmap);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.AntiAlias;

        // Kept in sync with Theme.xaml's Accent brush (#FF7C5CFC) by hand —
        // there's only one place that matters visually, so it's not worth
        // sharing a single color constant across a WPF brush and a GDI+ one.
        using var backgroundBrush = new SolidBrush(ColorTranslator.FromHtml("#FF7C5CFC"));
        using var backgroundPath = CreateRoundedRectanglePath(new Rectangle(0, 0, size, size), size / 4);
        g.FillPath(backgroundBrush, backgroundPath);

        using var font = new Font("Segoe UI", size * 0.4f, FontStyle.Bold, GraphicsUnit.Pixel);
        using var textBrush = new SolidBrush(Color.White);
        var format = new StringFormat
        {
            Alignment = StringAlignment.Center,
            LineAlignment = StringAlignment.Center
        };
        g.DrawString(monogram.ToUpperInvariant(), font, textBrush, new RectangleF(0, 0, size, size), format);

        // The HICON handle returned by GetHicon() is owned by the caller
        // (Icon.FromHandle does not take ownership or destroy it). This
        // template creates exactly one tray icon for the process lifetime,
        // so the single leaked handle is released by the OS on exit; a
        // service that creates many icons at runtime should track the
        // handle and call DestroyIcon on it once done.
        var handle = bitmap.GetHicon();
        return Icon.FromHandle(handle);
    }

    private static GraphicsPath CreateRoundedRectanglePath(Rectangle bounds, int radius)
    {
        var path = new GraphicsPath();
        var diameter = radius * 2;

        path.AddArc(bounds.X, bounds.Y, diameter, diameter, 180, 90);
        path.AddArc(bounds.Right - diameter, bounds.Y, diameter, diameter, 270, 90);
        path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(bounds.X, bounds.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();

        return path;
    }

    private static string Truncate(string value, int maxLength)
        => value.Length <= maxLength ? value : value[..maxLength];

    public void Dispose()
    {
        _notifyIcon.Visible = false;
        _notifyIcon.Icon?.Dispose();
        _notifyIcon.Dispose();
    }
}
