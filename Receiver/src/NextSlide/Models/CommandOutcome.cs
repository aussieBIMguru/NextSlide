namespace NextSlide.Models;

/// <summary>
/// How a processed sheet row was handled, shown per-row in the command log
/// and driving its row color (see Converters/OutcomeToBrushConverter.cs).
///
/// Only rows an action was actually attempted on get logged at all — a
/// row that's too old to act on (see CommandDedupeStore.IsStale) is
/// claimed silently and never reaches the grid, so hooking up mid-session
/// and picking up a backlog of old sheet rows doesn't flood the log.
/// </summary>
public enum CommandOutcome
{
    /// <summary>Matched a hooked, live Slide Show and the command was sent successfully.</summary>
    Fired,

    /// <summary>Matched our Session but couldn't be sent (not in Slide Show mode, unrecognized Command, PowerPoint closed, etc.) — see the log row's Detail text.</summary>
    Failed
}
