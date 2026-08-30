# Slide Remote — Receiver (NextSlide) Handoff

**Status:** Built, delivered as a Visual Studio solution, three rounds of real-run feedback already incorporated, and now polished with an app icon and full documentation (2026-08-30). Companion to `../../Remote/ai/handover_slideRemote.md` (the sender side, which *is* fully tested end-to-end) and `../../Remote/webSetup_slideRemote.md`. For the developer-facing build/structure doc see `../README.md`; for the end-user, both-apps-together operating guide see `../userGuide_slideRemote.md`. Your own original brief for this project should live alongside this file in `Receiver/ai/`.

---

## 1. What this is

NextSlide is the PC-side receiver described as "not yet built" in §5 of `../../Remote/ai/handover_slideRemote.md`. It polls the Google Sheet the sender writes to, filters to a locked Session name, and drives a live PowerPoint Slide Show via COM automation. Built as a fork of Gavin's **SystemApp** WPF template (Scaffold collection) — the tray/settings/theme/MVVM plumbing is template infrastructure; everything Slide-Remote-specific is new.

Lives in this repository as `Receiver/src/` (`NextSlide.sln` + `NextSlide/` project), zero NuGet packages.

## 2. Key architecture decisions

- **PowerPoint automation: COM via late-bound `dynamic`, not a typed PIA.**
  Reaches the running instance through the Running Object Table (same
  mechanism VBA's `GetObject` uses), via a direct P/Invoke to oleaut32's
  `GetActiveObject` (resolving the "PowerPoint.Application" ProgID to a
  CLSID first through ole32's `CLSIDFromProgID`) — **not**
  `System.Runtime.InteropServices.Marshal.GetActiveObject`, which only
  ever existed in .NET Framework and doesn't exist in modern .NET (see
  §4, round 1). Chosen over simulated keystrokes specifically so the app
  can positively confirm "is this presentation actually in Slide Show
  mode" via `presentation.SlideShowWindow`, rather than inferring it from
  window titles — and chosen over a typed Interop/PIA reference (the
  `<COMReference>` csproj approach) because that requires knowing the
  exact installed Office type-library version at build time, which
  couldn't be verified from the (non-Windows) environment this was built
  in. Net effect: zero NuGet packages, works against whatever PowerPoint
  version is installed, no COM reference wizard step needed in Visual
  Studio.
  **Known limitation:** the ROT only exposes one `PowerPoint.Application`
  moniker, so a second fully separate `POWERPNT.EXE` process (rare — not
  how opening multiple decks normally behaves on Windows) wouldn't be
  reachable. The combobox lists every open presentation within whichever
  single instance *is* found.

- **Sheet reads: unauthenticated `gviz` endpoint**, exactly as scoped in
  `../../Remote/ai/handover_slideRemote.md` §5.1 — `https://docs.google.com/spreadsheets/d/{ID}/gviz/tq?tqx=out:json[&gid={GID}]`.
  The Sheet URL field accepts any normal share/edit link; the app extracts
  the ID and optional `gid` itself via regex rather than requiring the raw
  gviz URL.

- **Dedup by hashed row metadata, not row position, purely in-memory** —
  per the hard requirement in `../../Remote/ai/handover_slideRemote.md` §5.2/§6 (track by
  Timestamp, never row position). SHA-256 over
  `Timestamp|Command|Slide#|Session`, kept only in an in-process
  Dictionary — **no disk persistence** (see §4 round 3: an earlier
  version dumped this to `seen-commands.json`, which turned out to let
  entries accumulate across sessions and flood the log with stale
  backlog rows far exceeding the sheet's actual ~30 live rows). A restart
  starting with an empty table is safe: anything it would have protected
  against re-firing is, by the time polling resumes, already past the
  staleness window below — so it comes back in as stale rather than
  refired, and stale rows are dropped silently, never logged.

- **Staleness window: ~15 seconds** (was 60s in round 1 — narrowed per
  Gavin's feedback), measured against the newest row timestamp seen in a
  poll, never against the receiver PC's own clock. The Sheet's Timestamp
  column is written in the *spreadsheet's* configured timezone, which
  generally won't match the receiver's local timezone — comparing
  against `DateTime.Now` directly would silently mis-age every row by
  that constant offset. Anchoring to the newest timestamp actually
  present cancels the offset out; only relative deltas between rows
  matter. `CommandDedupeStore`'s constructor takes the window as an
  optional `TimeSpan?` if a future need wants it configurable.

- **The command log only ever shows rows an action was actually
  attempted on** (Fired or Failed) — stale/backlog rows are claimed (so
  they never fire later) but never raised as an event at all, so they
  never reach the grid. `CommandOutcome.Skipped` was removed entirely
  rather than kept-but-hidden, since nothing constructs it anymore.

- **UI state machine**: Session (free text) → Lock → Sheet URL (unlocked)
  → valid URL → PowerPoint picker (unlocked) → pick a presentation →
  polling starts. Release stops polling and clears the presentation list
  but deliberately keeps the Sheet URL value, per the spec ("If (1) is
  released, (2) still holds its value"). Last-used Session/Sheet URL are
  persisted and pre-filled on next launch, but never auto-locked — polling
  only ever starts from an explicit user action. Locking also does an
  eager settings save (so the Session name survives even if the app is
  killed before a clean exit) — see §4 round 2 for a bug this surfaced.

- **Read-only against the Sheet** — no "Done" status written back,
  matching the deliberate scope boundary in `../../Remote/ai/handover_slideRemote.md` §5.6
  (would require real Sheets API write credentials this design avoids).

- **App icon: baked `.ico` alongside the still-runtime-rendered tray
  icon**, deliberately two separate assets. `TrayIconService` draws the
  NS monogram (accent purple `#7C5CFC`, `size/4` corner radius, centered
  bold white monogram) live via GDI+ every time it builds the tray
  `NotifyIcon` — that code was left completely untouched (Gavin: "the
  glyph is fine"). A `Resources/app.ico`, generated outside .NET
  (Python/Pillow, since this was built without a Windows/.NET toolchain)
  to replicate that exact same drawing logic as a static 7-size (16–256px)
  icon, covers the three places the runtime-drawn tray icon can't reach:
  the `.exe` file itself (`<ApplicationIcon>` in the csproj), and the
  taskbar button/Alt+Tab for both windows (`Icon="/Resources/app.ico"` on
  `MainWindow` and `MessageForm`). The two stay in sync only by hand if
  the design ever changes — documented in `../README.md`'s "App icon"
  section.

## 3. Run behavior

Defaults to `WindowedTrayOnClose` (template default is
`WindowedExitOnClose`; changed for NextSlide since the whole point is to
keep polling after the window is closed). Polling runs on a
`DispatcherTimer` (~1s) tied to the WPF UI thread — required because COM
calls into PowerPoint must happen on that STA thread — and is unaffected
by the window being hidden to tray.

## 4. Build/run feedback log

This was built in a Linux sandbox with no Windows, no .NET SDK, no
PowerPoint, and no live Google Sheet to test against, so real Visual
Studio builds/runs on Gavin's machine are the actual test.

**Round 1 — CS0117 `'Marshal' does not contain a definition for 'GetActiveObject'`**
(`Services/PowerPointController.cs`, build-time). Root cause: that
convenience wrapper only ever existed in .NET Framework's mscorlib and
was never ported into modern .NET's
`System.Runtime.InteropServices.Marshal`. Fixed by replacing it with a
direct P/Invoke — `CLSIDFromProgID` (ole32.dll) to resolve the ProgID,
then `GetActiveObject` (oleaut32.dll) against that CLSID — which is
exactly what the old Marshal method did internally.

**Round 2 — runtime `System.ArgumentException` on clicking Lock**:
"'.NET number values such as positive and negative infinity cannot be
written as valid JSON...'" (`Services/SettingsService.cs`, via
`MainViewModel.LockSession`'s eager `_settingsService.Save`). Root cause:
`AppSettings.WindowLeft`/`WindowTop` default to `double.NaN` as a
"no saved position yet" sentinel (base template design, documented in
`AppSettings.cs`) — that sentinel is only ever overwritten with a real
number when the window closes (`MainWindow.PersistWindowState`, called
from `App.ExitApplication`). On a fresh install, clicking Lock saves
*before* the window has ever closed once, so `JsonSerializer.Serialize`
hit a live NaN and threw — plain JSON has no token for NaN. Fixed by
adding `NumberHandling = JsonNumberHandling.AllowNamedFloatingPointLiterals`
to `SettingsService`'s `JsonSerializerOptions`, rather than redesigning
the sentinel.

**Round 3 — worked, but the command log/dedup table was "compounding
over time" and skipping far more rows than the sheet's actual ~30**
(functional feedback, not a crash). Root cause: `seen-commands.json`
persisted claimed-row hashes across every session/restart with no upper
bound on the file, and stale entries only got pruned when *currently
matching* rows happened to trigger a `Prune()` call — anything left over
from an earlier test session with no current rows to key off of just sat
there indefinitely. Combined with the (then) 60-second staleness window
and every stale row still being logged as a **Skipped** grid entry, this
made the log balloon over a long testing session. Fixed by: dropping
disk persistence entirely (`CommandDedupeStore` is now pure in-memory,
reset every launch — see §2 above for why that's safe), narrowing the
staleness window to 15s, and no longer logging stale rows to the grid at
all (`CommandOutcome.Skipped` removed).

Confirmed by Gavin after round 3: "great it is all working as intended" —
the polling/dedup/COM-drive loop is now considered done, no outstanding
functional bugs.

**Round 4 — polish pass, not a bug fix.** Gavin asked for three finishing
touches once the app was confirmed working: (1) a real icon for the
`.exe`/taskbar/Alt+Tab (explicitly keeping the existing tray glyph
design, just extending it to the places the tray icon alone doesn't
reach — see the new bullet in §2 above); (2) a developer-facing README
for NextSlide mirroring how the sender side has one (`../README.md`,
mirrors `../../Remote/webSetup_slideRemote.md`/
`../../Remote/ai/handover_slideRemote.md`'s split — build/structure/
troubleshooting for whoever works on the code); (3) a combined,
end-user-facing guide covering both apps together — session setup and
running an actual presentation from a non-developer's perspective —
written as `../userGuide_slideRemote.md` (new; there was no existing
equivalent to mirror, since the sender's own docs are dev-facing too).
All three delivered together; no code paths changed besides the
icon/csproj/XAML wiring described in §2.

**Round 5 — repo reorganized.** Split from a flat layout into
`Receiver/` (this app: `src/` for the VS solution, this `ai/` folder for
handover docs, plus `README.md` and `userGuide_slideRemote.md`) and
`Remote/` (the sender: `remote.html` and `purge.gs` as live source
files, `webSetup_slideRemote.md`, and its own `ai/` folder), under a
root `README.md` covering both. No functional changes — doc
cross-references throughout were updated for the new paths.

Still not independently verified: the exact PowerPoint object-model
member names/casing beyond what's documented (`Presentations`, `Name`,
`SlideShowWindow`, `SlideShowWindow.View.Next/Previous/GotoSlide`) —
believed correct and, as of round 3, confirmed working in Gavin's real
run.

## 5. Possible follow-ups (not built)

- Auto-refresh of the presentation combobox while hooked up (currently
  manual Refresh + once automatically on Lock, to avoid two independent
  loops touching COM concurrently).
- `First`/`Last` commands — mentioned as optional in
  `../../Remote/ai/handover_slideRemote.md` §5.4 but the sender never
  actually sends them; not implemented.
- Configurable poll interval (currently a hardcoded ~1s, though
  `SlidePollingService`'s constructor already accepts an override).
- Configurable staleness window (currently a hardcoded 15s default,
  though `CommandDedupeStore`'s constructor already accepts an override)
  — could be exposed in the UI/settings if Gavin wants to tune it later.
