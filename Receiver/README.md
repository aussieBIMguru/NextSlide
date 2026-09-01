# NextSlide

Watches a Google Sheet for remote-clicker commands and drives PowerPoint.

NextSlide is the receiver half of the Slide Remote project: a phone/laptop
sends `Next` / `Previous` / `Go to Slide` commands to a Google Sheet (via a
hidden Google Form intake), and this app polls that Sheet, filters to the
Session it's watching, and drives a live PowerPoint Slide Show
accordingly. Forked from the **SystemApp** template in Gavin's Scaffold
collection — most of the tray/settings/theme infrastructure below is
unchanged template plumbing; the Slide Remote receiver logic is what's new.

This README is the developer-facing side of *this* half of the project:
how NextSlide specifically is built and how to work on it. For the
project as a whole (both halves, how they fit together), start at
**`../README.md`**. For running an actual presentation end-to-end (both
the sender page and this app together), see
**`userGuide_slideRemote.md`** (same folder); for how the sender/Sheet/
Form side was built, see `../Remote/webSetup_slideRemote.md` and
`../Remote/ai/handover_slideRemote.md`.

## Getting Started

1. Open `src/NextSlide.sln` in Visual Studio 2022+ (or any IDE with .NET
   10 SDK support).
2. Target framework is `net10.0-windows` — Windows only (WPF + COM
   automation).
3. **PowerPoint must be installed on this machine.** NextSlide drives it
   via COM automation (finds the running `PowerPoint.Application` through
   the Running Object Table) — there's no bundled/portable PowerPoint, and
   nothing to configure for this beyond having a normal desktop
   Office/PowerPoint install.
4. Build and run (F5). Zero NuGet packages are required.

## How it works

Two-step hookup, matching the panel across the top of the window:

1. **Session & Sheet** — type the same Session name the remote sender
   uses *and* paste the Sheet URL, then click **Lock**. Both fields are
   free text together and validated together — Lock only enables once
   the Session name is non-empty and the Sheet URL looks like a real
   Google Sheets link. Locking commits both at once (and pre-fills them
   from last run, though it never auto-locks). **Release** unlocks both
   fields again (and stops polling) without clearing either value.
2. **Hook to PowerPoint** — only enabled once locked. Lists every
   presentation currently open in PowerPoint (via COM), flagging which
   one is actually in Slide Show mode. Pick one; polling starts
   automatically, and stops the moment the session is released or the
   presentation stops being available (see "Recovering from PowerPoint
   closing" below).

Once running (~2 polls/second, i.e. every ~0.5s): fetch the sheet →
filter rows to the locked Session → skip anything already processed or
too old to act on → drive the picked presentation if (and only if) it's
actually in Slide Show mode. Only rows an action was actually attempted
on reach the log — green = fired, red = failed, with a Detail column
explaining any failure. Rows older than the staleness window (see below)
are claimed silently and never shown, so hooking up mid-session onto a
sheet that still has old, not-yet-purged rows in it doesn't flood the
grid.

**Recovering from PowerPoint closing mid-session:** every COM call into
PowerPoint is guarded (`PowerPointController`, see its class doc comment)
so a closed PowerPoint app, or just the specific presentation being
closed, can never throw its way out and crash NextSlide — it always comes
back as a normal failed-command result instead. The first time a poll
tick discovers the target is gone, `SlidePollingService` raises
`PresentationUnavailable`; `MainViewModel` responds by clearing the
picked presentation (which stops polling) and re-scanning, *without*
unlocking the Session/Sheet step — so reopening PowerPoint (or the same
deck) and clicking **Refresh** is all it takes to pick back up, no
re-typing the Session/Sheet URL.

**Dedup / "already seen" tracking** (`Services/CommandDedupeStore.cs`):
purely in-memory for the lifetime of the running app — rows are hashed
(SHA-256 over Timestamp+Command+Slide#+Session) into a table that's
never written to disk, so nothing accumulates across restarts or
sessions. Entries older than ~10 seconds are pruned automatically —
measured against the *newest row timestamp seen in that poll*, not this
PC's clock, since the Sheet's Timestamp column is in the spreadsheet's
own configured timezone and generally won't match the receiver's local
timezone. That same 10-second window is also what "too old to act on"
means above; an earlier build used 15s (and, before that, a disk-persisted
60-second window), narrowed further so a backlog picked up while the app
was busy opening/reconnecting can't still be fresh enough to fire in a
rapid-fire burst — a restart doesn't need to survive this table, since
anything it would have protected against re-firing is, by definition,
already older than the window by the time polling resumes.

**PowerPoint automation** (`Services/PowerPointController.cs`): late-bound
COM via `dynamic`, not a typed Primary Interop Assembly — no COM/PIA
reference and no NuGet package, and it works against whatever PowerPoint
version happens to be installed. Reaches PowerPoint through the Running
Object Table, the way VBA's `GetObject` does — via a direct P/Invoke to
oleaut32's `GetActiveObject` (resolving the "PowerPoint.Application"
ProgID to a CLSID first through ole32's `CLSIDFromProgID`), not
`System.Runtime.InteropServices.Marshal.GetActiveObject` — that
convenience wrapper only ever existed in .NET Framework and isn't present
in modern .NET's `Marshal` class. Every call site catches broadly
(`COMException` plus the handful of other exception types the CLR maps
well-known HRESULTs to — see `PowerPointController.IsComInteropFailure`)
rather than just `COMException`, specifically so an external process
disappearing mid-call can never crash the app.

## Project structure

```
README.md                       This file
userGuide_slideRemote.md        End-to-end usage guide (both apps)
ai/
  handover_slideRemote_receiver.md   Design decisions / build-run feedback log
  (your original brief to Claude)
src/
  NextSlide.sln
  NextSlide/
    App.xaml / App.xaml.cs        Startup/shutdown, RunMode wiring, tray icon
    AppInfo.cs                    Name / Monogram / Publisher — single source of truth
    Views/
      MainWindow.xaml / .cs       The hookup panel, command log, status bar
      MessageForm.xaml / .cs      Themed dialog (replaces MessageBox — see "Dialogs" below)
    ViewModels/
      MainViewModel.cs            Session/Sheet/PowerPoint state machine + polling wiring
      CommandLogItemViewModel.cs  One row in the command log
    Models/
      RemoteCommand.cs            Next/Previous/GoToSlide + string parser (the sender's contract)
      SheetCommandRow.cs          One parsed sheet row
      CommandOutcome.cs           Fired / Failed
      PresentationOption.cs       One entry in the "hook to PowerPoint" combobox
      AppSettings.cs              Persisted settings (schema-versioned)
      RunMode.cs                  Silent / WindowedExitOnClose / WindowedTrayOnClose
    Services/
      GoogleSheetReader.cs        Sheet URL parsing + gviz fetch/parse
      CommandDedupeStore.cs       In-memory hash table + staleness pruning (no disk persistence)
      PowerPointController.cs     COM automation (list presentations, fire commands)
      SlidePollingService.cs      The ~0.5s DispatcherTimer loop tying the above together
      SettingsService.cs          Load/save AppSettings as JSON under %LOCALAPPDATA%
      TrayIconService.cs          NotifyIcon wrapper, all run modes go through this
    Mvvm/
      ObservableObject.cs / RelayCommand.cs   Hand-rolled INotifyPropertyChanged/ICommand
    Converters/
      OutcomeToBrushConverter.cs  Maps CommandOutcome to a Theme.xaml status brush
    Resources/
      Theme.xaml                  Shared dark/violet theme (matches the sender's own page)
      app.ico                     NS monogram, baked at build time — see "App icon" below
```

## App icon

`Resources/app.ico` is the same NS monogram/accent-purple design
`TrayIconService` renders on the fly for the tray icon, baked into a
7-size `.ico` (16–256px) so it shows up everywhere the tray icon alone
can't reach:

- **The `.exe` itself** (Explorer, shortcuts) — via `<ApplicationIcon>`
  in `NextSlide.csproj`.
- **The taskbar button and Alt+Tab** while the window is open — via
  `Icon="/Resources/app.ico"` on `MainWindow` (and `MessageForm`, for its
  dialogs).

The tray icon (`TrayIconService.CreateMonogramIcon`) is deliberately left
as-is — it renders the identical design at runtime with zero asset
dependency, which was the point of that approach and still works fine.
To change the glyph/color later, keep both in sync by hand: the GDI+
drawing code in `TrayIconService.cs`, and `Resources/app.ico` (regenerate
it — the accent color `#7C5CFC`, `size/4` corner radius, and centered
bold monogram are all documented there).

## Settings & data files

`%LOCALAPPDATA%\Gavin\NextSlide\settings.json` — window position/size,
RunMode, and the last-used Session name / Sheet URL (pre-filled on next
launch for convenience, but never auto-locked — polling only ever starts
from an explicit Lock + presentation pick). That's the only file NextSlide
writes; the dedup table (see above) is in-memory only.

## Run modes

Defaults to **close-to-tray** (`WindowedTrayOnClose`) — closing the
window hides it and keeps polling in the background; the tray icon's
**Show**/double-click brings the window back, **Exit** truly quits. Change
the default in `AppSettings.RunMode`, or force silent/tray-only for a
single launch with `--silent` (handy for a Task Scheduler entry or
Startup-folder shortcut).

## Known limitations

- **Single PowerPoint instance.** The Running Object Table lookup reaches
  whichever `PowerPoint.Application` is registered in the Running Object
  Table — in the overwhelmingly common case (one running PowerPoint,
  possibly with several decks open via File → Open, which share that one
  instance) this sees all of them. A second, fully separate `POWERPNT.EXE`
  process is rare and not how opening multiple files normally behaves on
  Windows, and isn't reachable this way.
- **Read-only against the Sheet.** By design (see
  `../Remote/ai/handover_slideRemote.md` §5.6) — no "Done" status is ever
  written back, so no Sheets API write credentials are needed.
- **Polling interval is fixed at 0.5 second** (`SlidePollingService`'s
  constructor takes an optional interval if a derived need ever wants it
  configurable; 0.5s was chosen as responsive without being so frequent
  it risks overlapping fetches on a slow connection — the reentrancy
  guard in `OnTick` handles that case regardless).
- **The presentation list doesn't auto-refresh** while hooked up — click
  Refresh if you open/close presentations after locking in. The Sheet
  poll loop and the presentation list intentionally don't run on two
  independent timers touching COM at once.

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| Presentation list is empty | PowerPoint isn't running — open a deck, then click Refresh. |
| Picked a presentation but nothing fires | It needs to actually be in Slide Show mode (F5 in PowerPoint), not just open — check the combobox's "(presenting)"/"(not presenting)" suffix. |
| Every row shows Failed, "PowerPoint rejected the command" | Usually a `Go to Slide` number beyond the deck's slide count. |
| Status bar shows a sheet/network error | Confirm the pasted URL is the sheet's real link and it's shared "Anyone with the link can view" — see `../Remote/ai/handover_slideRemote.md` §6 for the same gotchas on the sender side (unpublished Form, Workspace account restrictions, etc. don't apply here, but sheet sharing does). |
| A command from more than ~10 seconds ago never fires and never shows up in the log | Expected: hooking up (or restarting) picks up whatever's still in the sheet (not yet purged) but won't replay a backlog of old clicks — those rows are claimed and dropped silently rather than logged. |
| PowerPoint (or the deck) was closed while NextSlide was still hooked up | Expected, and no longer a crash — the log shows one red row explaining PowerPoint's gone, the picker clears back to unselected, and the Session/Sheet lock stays in place. Reopen PowerPoint (and/or the deck), click **Refresh**, and pick it again. |

## Extending the template

See the base SystemApp template's own conventions (hand-rolled MVVM, one
`Theme.xaml`, explicit-save settings with an opt-in autosave path,
`MessageForm` instead of `MessageBox`/`TaskDialog` for every dialog) —
this app follows them throughout; nothing here diverges from that base
without a comment explaining why.

## Changelog

- **2026-08-30** — Initial build. Forked from the SystemApp template;
  implemented the Session/Sheet/PowerPoint hookup flow, gviz polling,
  hash-based dedup with disk persistence, COM-driven PowerPoint control,
  and the command log.
- **2026-08-30** — Fixed CS0117 (`Marshal.GetActiveObject` doesn't exist
  in modern .NET) by P/Invoking `CLSIDFromProgID`/`GetActiveObject`
  directly. Fixed a startup crash serializing the `NaN` "no window
  position yet" sentinel by enabling
  `JsonNumberHandling.AllowNamedFloatingPointLiterals`. Reworked
  `CommandDedupeStore` to be purely in-memory (no more disk dump — it was
  accumulating entries across sessions), dropped the staleness window
  from 60s to 15s, and stopped logging stale/skipped rows to the command
  grid at all — only rows an action was actually attempted on show up.
- **2026-08-30** — Added `Resources/app.ico` (baked NS-monogram icon) and
  wired it into `<ApplicationIcon>` plus both windows' `Icon=`, so the
  `.exe`, taskbar, and Alt+Tab all show the same glyph the tray icon
  already did. Split off `userGuide_slideRemote.md` as the end-to-end,
  non-developer guide to running a presentation with both apps.
- **2026-08-30** — Moved this file to `Receiver/README.md` (receiver
  build/structure detail) now that `README.md` at the repo root is the
  whole-project overview covering both the sender and the receiver.
- **2026-08-30** — Repo reorganized into `Receiver/` (`src/` for the VS
  solution, `ai/` for handover docs) and `Remote/` (the sender, with its
  own `ai/` folder), each with this same README/userGuide split. Only
  doc paths changed.
- **2026-09-01** — Fixed a crash when PowerPoint (or the driven
  presentation) was closed while still hooked up: `PowerPointController`
  now catches broadly at every COM call site instead of just
  `COMException`, and `SlidePollingService`/`MainViewModel` recover
  gracefully (clear the picker, re-scan, stay locked) via a new
  `PresentationUnavailable` event rather than repeating the failure every
  tick. Merged the Session and Sheet URL steps into one combined
  lock — both are now validated and locked together instead of
  sequentially — so the hookup panel is a two-step flow, not three.
  Narrowed the dedupe staleness window from 15s to 10s (so a backlog
  picked up while reconnecting can't fire as a rapid-fire burst) and
  sped up polling from 1s to 0.5s.
