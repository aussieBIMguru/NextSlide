using System;

namespace NextSlide.Models;

/// <summary>
/// The three commands the sender page (remote.html / the Google Site embed)
/// ever writes into the sheet's Command column. These string values are a
/// contract with the sender — see the project's handover doc — not
/// something to invent independently.
/// </summary>
public enum RemoteCommand
{
    Next,
    Previous,
    GoToSlide
}

public static class RemoteCommandParser
{
    /// <summary>
    /// Matches a raw Command cell value to a <see cref="RemoteCommand"/>.
    /// Comparison is case-insensitive and trims whitespace, but otherwise
    /// expects exactly the sender's three literal strings ("Next",
    /// "Previous", "Go to Slide") — anything else (a typo, a blank row, a
    /// future command the sender doesn't send yet) is reported back as
    /// unrecognized rather than guessed at.
    /// </summary>
    public static bool TryParse(string? raw, out RemoteCommand command)
    {
        switch ((raw ?? string.Empty).Trim())
        {
            case var s when string.Equals(s, "Next", StringComparison.OrdinalIgnoreCase):
                command = RemoteCommand.Next;
                return true;

            case var s when string.Equals(s, "Previous", StringComparison.OrdinalIgnoreCase):
                command = RemoteCommand.Previous;
                return true;

            case var s when string.Equals(s, "Go to Slide", StringComparison.OrdinalIgnoreCase):
                command = RemoteCommand.GoToSlide;
                return true;

            default:
                command = default;
                return false;
        }
    }
}
