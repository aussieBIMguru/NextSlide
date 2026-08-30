using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using NextSlide.Models;

namespace NextSlide.Services;

/// <summary>
/// Drives a live PowerPoint Slide Show via COM automation, late-bound
/// through <c>dynamic</c> rather than a typed Primary Interop Assembly —
/// no COM/PIA reference, no NuGet package, works against whatever
/// PowerPoint version is installed (see README.md "PowerPoint automation"
/// for why this was chosen over simulated keystrokes).
///
/// Reaches PowerPoint via the Running Object Table, the same mechanism
/// VBA's <c>GetObject</c> uses — <see cref="TryGetRunningApplication"/>
/// P/Invokes oleaut32's GetActiveObject directly rather than calling
/// <c>System.Runtime.InteropServices.Marshal.GetActiveObject</c>, because
/// that convenience wrapper only ever existed in .NET Framework's mscorlib
/// and was never ported to modern .NET's Marshal class (CS0117 if you go
/// looking for it here). PowerPoint registers itself in the ROT once per
/// machine under the fixed ProgID "PowerPoint.Application" — in the
/// overwhelmingly common case of
/// one running PowerPoint instance (possibly with several presentations
/// open in it via File → Open, which share that one instance/process),
/// this reaches all of them. A second, fully separate POWERPNT.EXE
/// process — rare, and not how opening multiple files normally behaves on
/// Windows — would not be reachable this way; see README.md "Known
/// limitations".
///
/// All calls here must happen on an STA thread (WPF's UI thread is STA by
/// default) — see SlidePollingService's doc comment for why its timer
/// callback is safe to call straight into this class from.
/// </summary>
public sealed class PowerPointController
{
    /// <summary>
    /// Lists every open presentation in the running PowerPoint instance,
    /// flagging which one (if any) is actually in Slide Show mode right
    /// now. Returns an empty list (with <paramref name="diagnostic"/> set)
    /// if PowerPoint isn't running at all, rather than throwing — the
    /// combobox is expected to legitimately be empty until PowerPoint is
    /// launched.
    /// </summary>
    public IReadOnlyList<PresentationOption> ListOpenPresentations(out string? diagnostic)
    {
        diagnostic = null;
        var results = new List<PresentationOption>();

        dynamic app = TryGetRunningApplication();
        if (app is null)
        {
            diagnostic = "PowerPoint isn't running.";
            return results;
        }

        try
        {
            dynamic presentations = app.Presentations;
            var count = (int)presentations.Count;
            for (var i = 1; i <= count; i++) // PowerPoint collections are 1-indexed.
            {
                dynamic presentation = presentations[i];
                var name = (string)presentation.Name;
                var inSlideShow = IsInSlideShow(presentation);
                results.Add(new PresentationOption(name, inSlideShow));
            }
        }
        catch (COMException ex)
        {
            diagnostic = $"Couldn't read PowerPoint's open presentations: {ex.Message}";
        }

        return results;
    }

    /// <summary>
    /// Re-finds the named presentation (COM references from a previous
    /// poll tick are never cached — the user could have closed it) and, if
    /// it's currently in Slide Show mode, sends one command. Returns false
    /// with a human-readable <paramref name="detail"/> for every failure
    /// mode: PowerPoint closed, presentation closed, not presenting, or a
    /// rejected/out-of-range slide number.
    /// </summary>
    public bool TryExecuteCommand(string presentationName, RemoteCommand command, int? slideNumber, out string detail)
    {
        dynamic app = TryGetRunningApplication();
        if (app is null)
        {
            detail = "PowerPoint isn't running.";
            return false;
        }

        dynamic target = null;
        try
        {
            dynamic presentations = app.Presentations;
            var count = (int)presentations.Count;
            for (var i = 1; i <= count; i++)
            {
                dynamic presentation = presentations[i];
                if (string.Equals((string)presentation.Name, presentationName, StringComparison.OrdinalIgnoreCase))
                {
                    target = presentation;
                    break;
                }
            }
        }
        catch (COMException ex)
        {
            detail = $"Couldn't look up the presentation: {ex.Message}";
            return false;
        }

        if (target is null)
        {
            detail = $"'{presentationName}' is no longer open.";
            return false;
        }

        if (!IsInSlideShow(target))
        {
            detail = "Not currently in Slide Show mode.";
            return false;
        }

        try
        {
            dynamic view = target.SlideShowWindow.View;
            switch (command)
            {
                case RemoteCommand.Next:
                    view.Next();
                    break;

                case RemoteCommand.Previous:
                    view.Previous();
                    break;

                case RemoteCommand.GoToSlide:
                    if (slideNumber is not { } n || n < 1)
                    {
                        detail = "Missing or invalid slide number.";
                        return false;
                    }

                    view.GotoSlide(n);
                    break;

                default:
                    detail = $"Unrecognized command '{command}'.";
                    return false;
            }
        }
        catch (COMException ex)
        {
            // Covers e.g. GotoSlide with a number beyond the deck's slide
            // count — PowerPoint rejects it via HRESULT rather than a
            // silent no-op.
            detail = $"PowerPoint rejected the command: {ex.Message}";
            return false;
        }

        detail = "OK";
        return true;
    }

    private static bool IsInSlideShow(dynamic presentation)
    {
        try
        {
            return presentation.SlideShowWindow is not null;
        }
        catch (COMException)
        {
            return false;
        }
    }

    /// <summary>
    /// P/Invoke replacement for <c>Marshal.GetActiveObject</c> (removed
    /// from modern .NET — see this class's doc comment). Resolves the
    /// ProgID to a CLSID via ole32's CLSIDFromProgID, then asks oleaut32's
    /// GetActiveObject for whatever's registered under that CLSID in the
    /// Running Object Table — exactly what the old Marshal method did
    /// internally on .NET Framework.
    /// </summary>
    [DllImport("ole32.dll", PreserveSig = false)]
    private static extern void CLSIDFromProgID(
        [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid lpclsid);

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    private static dynamic TryGetRunningApplication()
    {
        try
        {
            CLSIDFromProgID("PowerPoint.Application", out var clsid);
            GetActiveObject(ref clsid, IntPtr.Zero, out var app);
            return app;
        }
        catch (COMException)
        {
            // MK_E_UNAVAILABLE from GetActiveObject — nothing registered in
            // the ROT under that CLSID, i.e. PowerPoint simply isn't
            // running. (CLSIDFromProgID failing — PowerPoint not
            // installed/registered — throws the same way and is handled
            // identically here.)
            return null;
        }
    }
}
