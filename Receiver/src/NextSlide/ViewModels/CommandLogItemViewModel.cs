using System;
using NextSlide.Models;

namespace NextSlide.ViewModels;

/// <summary>
/// One row in the command-log DataGrid. Immutable once created — a
/// processed sheet row's outcome never changes after the fact — so unlike
/// WorkItemViewModel in the base template this doesn't need to be an
/// ObservableObject; the grid picks up new rows through the
/// ObservableCollection they're inserted into (MainViewModel.CommandLog),
/// newest at index 0. Only reports rows processed since this app launched
/// — nothing here is loaded from the dedupe store's disk dump, which
/// exists purely to prevent re-firing, not to reconstruct history.
/// </summary>
public sealed class CommandLogItemViewModel
{
    public CommandLogItemViewModel(SheetCommandRow row, CommandOutcome outcome, string detail)
    {
        ProcessedAt = DateTime.Now;
        SheetTimestamp = row.Timestamp;
        Command = row.Command;
        SlideNumber = row.SlideNumber;
        Session = row.Session;
        Outcome = outcome;
        Detail = detail;
    }

    /// <summary>When this app processed the row (local wall clock) — kept separate from SheetTimestamp, which is in the spreadsheet's own configured timezone.</summary>
    public DateTime ProcessedAt { get; }

    public DateTime SheetTimestamp { get; }
    public string Command { get; }
    public int? SlideNumber { get; }
    public string Session { get; }
    public CommandOutcome Outcome { get; }
    public string Detail { get; }
}
