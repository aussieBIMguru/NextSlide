using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Windows.Threading;
using NextSlide.Models;

namespace NextSlide.Services;

/// <summary>One sheet row that's been processed, for MainViewModel to turn into a command-log entry.</summary>
public sealed class SlideCommandProcessedEventArgs : EventArgs
{
    public SlideCommandProcessedEventArgs(SheetCommandRow row, CommandOutcome outcome, string detail)
    {
        Row = row;
        Outcome = outcome;
        Detail = detail;
    }

    public SheetCommandRow Row { get; }
    public CommandOutcome Outcome { get; }
    public string Detail { get; }
}

/// <summary>
/// Owns the ~0.5s poll loop: fetch the sheet, filter to our Session, dedupe,
/// and drive PowerPoint for whatever's new — see the project's handover
/// doc §5 for the design this implements end to end.
///
/// Uses a <see cref="DispatcherTimer"/> rather than a background
/// Task+Task.Delay loop specifically so its Tick handler — and therefore
/// every <see cref="PowerPointController"/> call it makes — runs on the
/// UI/STA thread that constructed this service. COM automation requires
/// that; construct this from MainViewModel (on the WPF UI thread), never
/// from a background thread. The network fetch inside OnTick is awaited
/// without ConfigureAwait(false) for the same reason WorkSimulationService
/// in the base template doesn't: the continuation needs to come back to
/// this thread to safely touch PowerPoint's COM objects.
///
/// <see cref="OnTick"/> is <c>async void</c> (the only option for a
/// DispatcherTimer.Tick handler), which means any exception that escapes it
/// doesn't just fail this poll — it crashes the entire application, since
/// there's no caller left to observe it. <see cref="PowerPointController"/>
/// is written to never let a "PowerPoint closed" failure escape as an
/// exception, but the try/catch around its call below is a deliberate
/// second line of defense against that specific failure mode, not just
/// tidiness.
/// </summary>
public sealed class SlidePollingService : IDisposable
{
    private readonly HttpClient _httpClient = new() { Timeout = TimeSpan.FromSeconds(10) };
    private readonly CommandDedupeStore _dedupeStore;
    private readonly PowerPointController _powerPoint = new();
    private readonly DispatcherTimer _timer;
    private bool _tickInProgress;

    public event EventHandler<SlideCommandProcessedEventArgs>? CommandProcessed;
    public event EventHandler<string>? PollError;

    /// <summary>
    /// Raised the moment a command attempt discovers the target
    /// presentation is gone — PowerPoint closed, or that specific
    /// presentation was closed — rather than just "not presenting right
    /// now" or "rejected". MainViewModel uses this to stop polling and drop
    /// back to the presentation picker instead of silently re-attempting
    /// (and re-logging a Failed row for) the same dead target on every
    /// subsequent tick.
    /// </summary>
    public event EventHandler? PresentationUnavailable;

    /// <summary>Session name to filter sheet rows to. Set before <see cref="Start"/>.</summary>
    public string? Session { get; set; }

    /// <summary>The gviz query URL built from the pasted Sheet URL (see GoogleSheetReader.BuildGvizUrl).</summary>
    public string? GvizUrl { get; set; }

    /// <summary>The PowerPoint presentation name (PresentationOption.Name) to drive.</summary>
    public string? TargetPresentationName { get; set; }

    public SlidePollingService(CommandDedupeStore dedupeStore, TimeSpan? interval = null)
    {
        _dedupeStore = dedupeStore;
        _timer = new DispatcherTimer { Interval = interval ?? TimeSpan.FromSeconds(0.5) };
        _timer.Tick += OnTick;
    }

    public bool IsRunning => _timer.IsEnabled;

    public void Start() => _timer.Start();

    public void Stop() => _timer.Stop();

    private async void OnTick(object? sender, EventArgs e)
    {
        // Reentrancy guard: if a fetch is unusually slow (slow network),
        // don't stack up overlapping polls — just wait for the next tick.
        if (_tickInProgress)
            return;

        // Captured into locals (rather than re-reading the properties
        // below) so the null checks actually narrow what gets used —
        // Session/GvizUrl/TargetPresentationName are mutable properties
        // the view model can reassign between statements.
        var gvizUrl = GvizUrl;
        var session = Session;
        var targetPresentationName = TargetPresentationName;

        if (string.IsNullOrWhiteSpace(gvizUrl) ||
            string.IsNullOrWhiteSpace(session) ||
            string.IsNullOrWhiteSpace(targetPresentationName))
        {
            return;
        }

        _tickInProgress = true;
        try
        {
            IReadOnlyList<SheetCommandRow> rows;
            try
            {
                rows = await GoogleSheetReader.FetchRowsAsync(_httpClient, gvizUrl, CancellationToken.None);
            }
            catch (Exception ex)
            {
                PollError?.Invoke(this, $"Couldn't reach the sheet: {ex.Message}");
                return;
            }

            var sessionTrimmed = session.Trim();
            var mine = rows
                .Where(r => string.Equals(r.Session?.Trim(), sessionTrimmed, StringComparison.OrdinalIgnoreCase))
                .OrderBy(r => r.Timestamp)
                .ToList();

            if (mine.Count == 0)
                return;

            // "Now" for staleness purposes is the newest row timestamp
            // actually present, not this PC's clock — see
            // CommandDedupeStore.Prune's doc comment for why.
            var referenceNow = mine.Max(r => r.Timestamp);
            _dedupeStore.Prune(referenceNow);

            foreach (var row in mine)
            {
                if (!_dedupeStore.TryClaim(row))
                    continue; // already processed on a previous tick (or before a restart)

                if (_dedupeStore.IsStale(row.Timestamp, referenceNow))
                {
                    // Claimed so it never fires later, but deliberately
                    // not raised as an event at all — a backlog picked up
                    // right when hooking up (or right after a restart)
                    // would otherwise flood the command log with rows
                    // nothing was actually done about. See
                    // CommandDedupeStore's doc comment for why a restart
                    // doesn't need anything more than this to stay safe.
                    continue;
                }

                if (!RemoteCommandParser.TryParse(row.Command, out var command))
                {
                    CommandProcessed?.Invoke(this, new SlideCommandProcessedEventArgs(
                        row, CommandOutcome.Failed, $"Unrecognized command value '{row.Command}'."));
                    continue;
                }

                bool fired;
                string detail;
                bool presentationUnavailable;
                try
                {
                    fired = _powerPoint.TryExecuteCommand(
                        targetPresentationName, command, row.SlideNumber, out detail, out presentationUnavailable);
                }
                catch (Exception ex)
                {
                    // Should be unreachable — PowerPointController is meant
                    // to convert every COM failure into a false/detail
                    // return — but this is an async void tick, so treat any
                    // surprise the same way rather than letting it crash
                    // the app (see this class's doc comment).
                    fired = false;
                    detail = $"Unexpected error talking to PowerPoint: {ex.Message}";
                    presentationUnavailable = true;
                }

                CommandProcessed?.Invoke(this, new SlideCommandProcessedEventArgs(
                    row, fired ? CommandOutcome.Fired : CommandOutcome.Failed, detail));

                if (presentationUnavailable)
                {
                    // PowerPoint (or this specific presentation) is gone —
                    // no point trying the rest of this batch against the
                    // same dead target. MainViewModel drops back to the
                    // picker in response; polling itself stops as a side
                    // effect of that (see MainViewModel.OnPresentationUnavailable).
                    PresentationUnavailable?.Invoke(this, EventArgs.Empty);
                    break;
                }
            }
        }
        finally
        {
            _tickInProgress = false;
        }
    }

    public void Dispose()
    {
        _timer.Stop();
        _timer.Tick -= OnTick;
        _httpClient.Dispose();
    }
}
