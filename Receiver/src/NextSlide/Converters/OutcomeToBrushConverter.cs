using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using NextSlide.Models;

namespace NextSlide.Converters;

/// <summary>
/// Maps a CommandOutcome to the matching status brush defined in
/// Resources/Theme.xaml — the same fallback-safe pattern the base
/// template used for WorkItemStatus. Falls back to Text.Secondary /
/// Brushes.Gray so a theme edit that renames a key never crashes the UI.
/// </summary>
public sealed class OutcomeToBrushConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var resourceKey = value switch
        {
            CommandOutcome.Fired => "Status.Success",
            CommandOutcome.Failed => "Status.Error",
            _ => "Text.Secondary"
        };

        if (Application.Current?.TryFindResource(resourceKey) is Brush brush)
            return brush;

        return Brushes.Gray;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException($"{nameof(OutcomeToBrushConverter)} does not support ConvertBack.");
}
