using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using NextSlide.Models;

namespace NextSlide.Services;

/// <summary>
/// Reads the Google Sheet the sender's Form writes to, via the
/// unauthenticated "gviz" query endpoint — no OAuth/service-account
/// credentials, as long as the sheet is shared "Anyone with the link can
/// view" (see the project's handover doc §5.1). Read-only: this app never
/// writes back to the sheet.
/// </summary>
public static class GoogleSheetReader
{
    private static readonly Regex SheetIdPattern = new(@"/spreadsheets/d/([a-zA-Z0-9-_]+)", RegexOptions.Compiled);
    private static readonly Regex GidPattern = new(@"[?&#]gid=(\d+)", RegexOptions.Compiled);
    private static readonly Regex GvizDatePattern =
        new(@"^Date\((\d+),(\d+),(\d+)(?:,(\d+),(\d+),(\d+))?\)$", RegexOptions.Compiled);

    /// <summary>
    /// Pulls the spreadsheet ID (and, if present, the tab's gid) out of any
    /// normal Google Sheets share/edit URL the user might paste — e.g.
    /// "https://docs.google.com/spreadsheets/d/{ID}/edit#gid=123". Returns
    /// false for anything that doesn't look like a Sheets URL at all.
    /// </summary>
    public static bool TryParseSheetUrl(string? url, out string sheetId, out string? gid)
    {
        sheetId = "";
        gid = null;

        if (string.IsNullOrWhiteSpace(url))
            return false;

        var idMatch = SheetIdPattern.Match(url);
        if (!idMatch.Success)
            return false;

        sheetId = idMatch.Groups[1].Value;

        var gidMatch = GidPattern.Match(url);
        if (gidMatch.Success)
            gid = gidMatch.Groups[1].Value;

        return true;
    }

    /// <summary>
    /// Builds the gviz JSON query URL for a sheet. Omitting gid defaults to
    /// the first (leftmost) tab, which is "Form Responses 1" in the normal
    /// Form-linked-Sheet setup this app is built against.
    /// </summary>
    public static string BuildGvizUrl(string sheetId, string? gid) =>
        string.IsNullOrEmpty(gid)
            ? $"https://docs.google.com/spreadsheets/d/{sheetId}/gviz/tq?tqx=out:json"
            : $"https://docs.google.com/spreadsheets/d/{sheetId}/gviz/tq?tqx=out:json&gid={gid}";

    public static async Task<IReadOnlyList<SheetCommandRow>> FetchRowsAsync(
        HttpClient httpClient, string gvizUrl, CancellationToken cancellationToken)
    {
        var text = await httpClient.GetStringAsync(gvizUrl, cancellationToken).ConfigureAwait(false);
        return ParseRows(text);
    }

    /// <summary>
    /// Parses a gviz response. The endpoint wraps its JSON in a JS call —
    /// <c>google.visualization.Query.setResponse({...});</c>, sometimes
    /// preceded by a <c>/*O_o*/</c> comment line — so this strips down to
    /// the outermost {...} object rather than assuming a fixed prefix.
    /// Columns are matched by header label (case-insensitive, ignoring the
    /// "#" in "Slide #") so a reordered or renamed-but-recognizable sheet
    /// tab still parses; falls back to the Form's default A–D column order
    /// (Timestamp, Command, Slide #, Session) if labels aren't found.
    /// </summary>
    internal static IReadOnlyList<SheetCommandRow> ParseRows(string gvizResponseText)
    {
        var start = gvizResponseText.IndexOf('{');
        var end = gvizResponseText.LastIndexOf('}');
        if (start < 0 || end <= start)
            return Array.Empty<SheetCommandRow>();

        using var doc = JsonDocument.Parse(gvizResponseText.Substring(start, end - start + 1));
        var root = doc.RootElement;

        if (root.TryGetProperty("status", out var status) &&
            status.ValueKind == JsonValueKind.String &&
            string.Equals(status.GetString(), "error", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "The sheet rejected the query — confirm the URL is a valid Sheets link and the sheet is shared 'Anyone with the link can view'.");
        }

        var table = root.GetProperty("table");
        var (idxTimestamp, idxCommand, idxSlide, idxSession) = ResolveColumnIndexes(table);

        var rows = new List<SheetCommandRow>();
        if (!table.TryGetProperty("rows", out var rowsElement))
            return rows;

        foreach (var row in rowsElement.EnumerateArray())
        {
            if (!row.TryGetProperty("c", out var cells))
                continue;

            var timestamp = ReadDate(cells, idxTimestamp);
            if (timestamp is null)
                continue; // Can't dedupe or age-check a row with no usable timestamp — skip it.

            var command = ReadString(cells, idxCommand) ?? "";
            var session = ReadString(cells, idxSession) ?? "";
            var slideNumber = ReadInt(cells, idxSlide);

            rows.Add(new SheetCommandRow(timestamp.Value, command, slideNumber, session));
        }

        return rows;
    }

    private static (int Timestamp, int Command, int Slide, int Session) ResolveColumnIndexes(JsonElement table)
    {
        // Form-default order, used whenever label matching below doesn't
        // find all four columns.
        int idxTimestamp = 0, idxCommand = 1, idxSlide = 2, idxSession = 3;

        if (!table.TryGetProperty("cols", out var cols))
            return (idxTimestamp, idxCommand, idxSlide, idxSession);

        var found = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var i = 0;
        foreach (var col in cols.EnumerateArray())
        {
            var label = col.TryGetProperty("label", out var l) ? l.GetString() : null;
            if (!string.IsNullOrWhiteSpace(label))
                found[label.Replace("#", "").Trim()] = i;
            i++;
        }

        if (found.TryGetValue("Timestamp", out var t)) idxTimestamp = t;
        if (found.TryGetValue("Command", out var c)) idxCommand = c;
        if (found.TryGetValue("Slide", out var s)) idxSlide = s;
        if (found.TryGetValue("Session", out var se)) idxSession = se;

        return (idxTimestamp, idxCommand, idxSlide, idxSession);
    }

    private static JsonElement? CellAt(JsonElement cells, int index)
    {
        if (cells.ValueKind != JsonValueKind.Array || index < 0 || index >= cells.GetArrayLength())
            return null;

        var cell = cells[index];
        return cell.ValueKind == JsonValueKind.Null ? null : cell;
    }

    private static string? ReadString(JsonElement cells, int index)
    {
        var cell = CellAt(cells, index);
        if (cell is not { } el || !el.TryGetProperty("v", out var v))
            return null;

        return v.ValueKind switch
        {
            JsonValueKind.String => v.GetString(),
            JsonValueKind.Number => v.GetRawText(),
            _ => null
        };
    }

    private static int? ReadInt(JsonElement cells, int index)
    {
        var cell = CellAt(cells, index);
        if (cell is not { } el || !el.TryGetProperty("v", out var v))
            return null;

        return v.ValueKind switch
        {
            JsonValueKind.Number => (int)v.GetDouble(),
            JsonValueKind.String when int.TryParse(v.GetString(), out var n) => n,
            _ => null
        };
    }

    /// <summary>
    /// gviz encodes datetime cells as the *string* "Date(year,monthIndex,
    /// day,hour,minute,second)" — not a JSON date type, and monthIndex is
    /// 0-based (JS convention), so January is 0.
    /// </summary>
    private static DateTime? ReadDate(JsonElement cells, int index)
    {
        var cell = CellAt(cells, index);
        if (cell is not { } el || !el.TryGetProperty("v", out var v) || v.ValueKind != JsonValueKind.String)
            return null;

        var match = GvizDatePattern.Match(v.GetString() ?? "");
        if (!match.Success)
            return null;

        try
        {
            var year = int.Parse(match.Groups[1].Value);
            var month = int.Parse(match.Groups[2].Value) + 1;
            var day = int.Parse(match.Groups[3].Value);
            var hour = match.Groups[4].Success ? int.Parse(match.Groups[4].Value) : 0;
            var minute = match.Groups[5].Success ? int.Parse(match.Groups[5].Value) : 0;
            var second = match.Groups[6].Success ? int.Parse(match.Groups[6].Value) : 0;
            return new DateTime(year, month, day, hour, minute, second);
        }
        catch (ArgumentOutOfRangeException)
        {
            return null;
        }
    }
}
