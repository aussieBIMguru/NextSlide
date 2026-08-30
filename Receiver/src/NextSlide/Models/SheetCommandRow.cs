using System;

namespace NextSlide.Models;

/// <summary>
/// One parsed row from the Google Sheet's "Form Responses" tab — the four
/// columns the Form actually writes (see the project's handover doc §2):
/// Timestamp, Command, Slide #, Session. <see cref="Timestamp"/> is the
/// sheet's own value (parsed from gviz's "Date(y,m,d,h,mi,s)" cell format),
/// never wall-clock receipt time — that's what makes it safe to track
/// "already processed" across restarts and purges (see
/// <see cref="Services.CommandDedupeStore"/>).
/// </summary>
public sealed record SheetCommandRow(DateTime Timestamp, string Command, int? SlideNumber, string Session)
{
    /// <summary>
    /// Stable string built from every field, used as the input to the
    /// dedupe store's hash. Deliberately includes the raw Command text
    /// (not the parsed enum) so two rows that differ only by an
    /// unrecognized Command value still hash differently.
    /// </summary>
    public string DedupeKey => $"{Timestamp:O}|{Command}|{SlideNumber?.ToString() ?? ""}|{Session}";
}
