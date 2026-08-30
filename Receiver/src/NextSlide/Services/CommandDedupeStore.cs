using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using NextSlide.Models;

namespace NextSlide.Services;

/// <summary>
/// Tracks which sheet rows have already been acted on for the lifetime of
/// this run, so a row is never fired twice — the same requirement the
/// sheet's own 15-minute purge script depends on us honoring (see the
/// project's handover doc §5.2 and §6): track by the row's own Timestamp,
/// never row position, since a purge run shifts every row above the one
/// it deletes.
///
/// Purely in-memory — no disk dump. An earlier version persisted this to
/// a local JSON file (specifically so a restart wouldn't replay a command
/// it had just fired moments earlier), but in practice that let entries
/// pile up indefinitely across sessions and made the "already seen" table
/// grow far beyond what the live sheet actually contained. The simpler
/// fix: a restart can only ever re-encounter rows that are already older
/// than <see cref="_maxAge"/> by the time polling resumes — nothing a
/// fresh process just claimed a moment ago can still be sitting
/// unprocessed — so those rows come back in as stale, not fireable
/// (see <see cref="IsStale"/>), and stale rows are never logged to the
/// command grid (see SlidePollingService) or sent to PowerPoint. A
/// restart is simply invisible rather than something this needs to
/// survive.
/// </summary>
public sealed class CommandDedupeStore
{
    private readonly TimeSpan _maxAge;
    private readonly Dictionary<string, DateTime> _seen = new();

    /// <param name="maxAge">
    /// How old a row can be (relative to the newest row timestamp seen in
    /// a poll — see <see cref="Prune"/>) before it's pruned from this
    /// store and treated as too stale to act on. Defaults to 15 seconds.
    /// </param>
    public CommandDedupeStore(TimeSpan? maxAge = null)
    {
        _maxAge = maxAge ?? TimeSpan.FromSeconds(15);
    }

    /// <summary>
    /// True the first time this exact row is seen (and records it);
    /// false on every subsequent poll that still contains the same row.
    /// </summary>
    public bool TryClaim(SheetCommandRow row)
    {
        var hash = Hash(row.DedupeKey);
        if (_seen.ContainsKey(hash))
            return false;

        _seen[hash] = row.Timestamp;
        return true;
    }

    /// <summary>
    /// True if <paramref name="rowTimestamp"/> is already older than
    /// this store's max age relative to <paramref name="referenceNow"/>.
    /// Used to decide whether a newly-claimed row should actually be
    /// acted on and logged, or just silently claimed and dropped.
    /// </summary>
    public bool IsStale(DateTime rowTimestamp, DateTime referenceNow) =>
        rowTimestamp < referenceNow - _maxAge;

    /// <summary>
    /// Drops entries older than this store's max age relative to
    /// <paramref name="referenceNow"/> — keeps the in-memory table from
    /// growing for the lifetime of a long-running session. Callers pass
    /// the newest row Timestamp seen in the latest poll batch as
    /// <paramref name="referenceNow"/> — deliberately NOT this PC's
    /// DateTime.Now. The sheet's Timestamp column is written in the
    /// spreadsheet's own configured timezone, which generally won't match
    /// the receiver PC's local timezone; a wall-clock comparison would
    /// mis-age every row by that constant offset. Anchoring to the newest
    /// timestamp actually present cancels that offset out — only the
    /// relative age between rows matters.
    /// </summary>
    public void Prune(DateTime referenceNow)
    {
        var cutoff = referenceNow - _maxAge;
        var stale = _seen.Where(kv => kv.Value < cutoff).Select(kv => kv.Key).ToList();
        foreach (var key in stale)
            _seen.Remove(key);
    }

    private static string Hash(string input) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(input)));
}
