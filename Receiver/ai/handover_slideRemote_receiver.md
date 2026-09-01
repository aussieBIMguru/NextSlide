# Slide Remote — Receiver (NextSlide) Handoff

**Status:** Built, delivered as a Visual Studio solution, six rounds of real-run feedback already incorporated — including a round 6 crash fix (PowerPoint closing mid-session), a merged Session+Sheet lock step, a faster poll, and a narrower staleness window (2026-09-01). Companion to `../../Remote/ai/handover_slideRemote.md` (the sender side, which *is* fully tested end-to-end) and `../../Remote/webSetup_slideRemote.md`. For the developer-facing build/structure doc see `../README.md`; for the end-user, both-apps-together operating guide see `../userGuide_slideRemote.md`. Your own original brief for this project should live alongside this file in `Receiver/ai/`.

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

- **Staleness window: ~10 seconds** (60s in round 1, narrowed to 15s in
  round 3, narrowed again to 10s in round 6 — see below), measured
  against the newest row timestamp seen in a poll, never against the
  receiver PC's own clock. The Sheet's Timestamp column is written in the
  *spreadsheet's* configured timezone, which generally won't match the
  receiver's local timezone — comparing against `DateTime.Now` directly
  would silently mis-age every row by that constant offset. Anchoring to
  the newest timestamp actually present cancels the offset out; only
  relative deltas between rows matter. `CommandDedupeStore`'s constructor
  takes the window as an optional `TimeSpan?` if a future need wants it
  configurable.

- **The command log only ever shows rows an action was actually
  attempted on** (Fired or Failed) — stale/backlog rows are claimed (so
  they never fire later) but never raised as an event at all, so they
  never reach the grid. `CommandOutcome.Skipped` was removed entirely
  rather than kept-but-hidden, since nothing constructs it anymore.

- **UI state machine (as of round 6)**: Session name + Sheet URL (both
  free text together) → Lock (only enabled once both are valid) → both
  commit at once and become read-only → PowerPoint picker (unlocked) →
  pick a presentation → polling starts. Release stops polling, clears the
  presentation list, and unlocks both Session name and Sheet URL again —
  deliberately keeping both values in place so re-locking doesn't require
  retyping either. (Rounds 1–5 locked Session and Sheet URL sequentially,
  one field unlocking the next; round 6 merged them into one combined
  step per Gavin's feedback — see §4 round 6.) Last-used Session/Sheet URL
  are persisted and pre-filled on next launch, but never auto-locked —
  polling only ever starts from an explicit user action. Locking also
  does an eager settings save of both fields (so they survive even if the
  app is killed before a clean exit) — see §4 round 2 for a bug this
  surfaced. `SlidePollingService.PresentationUnavailable` (new in round 6)
  drops the flow back to "pick a presentation" — without unlocking
  Session/Sheet — whenever PowerPoint or the target presentation vanishes
  mid-session; see §4 round 6.

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
`DispatcherTimer` (~0.5s as of round 6, was ~1s in rounds 1–5) tied to the
WPF UI thread — required because COM calls into PowerPoint must happen on
that STA thread — and is unaffected by the window being hidden to tray.

- **COM failures never escape as exceptions (round 6).** Every
  `PowerPointController` call site that touches PowerPoint's COM object
  model catches broadly — `COMException` plus the handful of other
  exception types the CLR maps well-known HRESULTs to (`ArgumentException`,
  `NullReferenceException`, `InvalidComObjectException`, etc. — see
  `PowerPointController.IsComInteropFailure`'s doc comment) — rather than
  `catch (COMException)` alone. This matters specifically because
  `SlidePollingService.OnTick` is `async void` (the only option for a
  `DispatcherTimer.Tick` handler): any exception that escapes it has no
  caller left to observe it and crashes the whole app, not just that poll.
  `SlidePollingService` also wraps its `TryExecuteCommand` call site in its
  own try/catch as a second line of defense. See §4 round 6 for the crash
  this fixes.

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

**Round 6 — fixed a real crash, plus three tuning/UX requests, from
real-run feedback after round 5.** Gavin reported: closing PowerPoint
mid-session (without clicking Release first) threw and crashed the app;
the app "chews up" commands when opening/reconnecting because the
staleness window was too generous; and asked for the Session/Sheet URL
hookup to be simplified so the two lock in together as a validated pair,
with a faster poll. Changes:

1. **Crash fix.** `PowerPointController`'s COM call sites previously
   caught only `COMException`; a disconnected RCW or a well-known HRESULT
   the CLR maps to a different exception type (`NullReferenceException`,
   `ArgumentException`, etc.) could still escape. Since
   `SlidePollingService.OnTick` is `async void`, anything escaping it
   crashes the whole app, not just that tick — matching the reported
   symptom exactly. Broadened every catch to cover that wider set (see
   `IsComInteropFailure`), added a `presentationUnavailable` out-parameter
   to `TryExecuteCommand` so callers know *why* it failed (target gone vs.
   not presenting vs. rejected), and added a second, outer try/catch
   around the `TryExecuteCommand` call in `SlidePollingService` as a
   deliberate belt-and-suspenders layer. On top of just not crashing, a
   new `PresentationUnavailable` event lets `MainViewModel` recover
   gracefully — clears the picked presentation (stopping polling as a
   side effect) and re-scans, without unlocking Session/Sheet — instead of
   silently re-attempting (and re-logging Failed for) the same dead
   target every 0.5s forever.
2. **UI restructure.** Merged the Session name and Sheet URL fields into
   one combined lock step (`GroupBox` "1 · Session & Sheet") — Lock now
   validates both together (non-empty Session, parseable Sheet URL)
   rather than sequentially unlocking one after the other. `IsHookupEditable`
   replaces the old `IsSessionNameEditable`/`IsSheetUrlEditable` pair.
   Release still keeps both values in place, now explicitly for both
   fields rather than just the Sheet URL.
3. **Staleness window narrowed 15s → 10s** (`CommandDedupeStore`) — per
   Gavin's "chews up commands when opening" report: a backlog picked up
   right when hooking up (or reconnecting) could still be within the old
   15s window and fire in a rapid-fire burst instead of being silently
   dropped as stale. 10s trims that without cutting into a normal single
   command's round trip.
4. **Poll interval halved 1s → 0.5s** (`SlidePollingService`) — Gavin
   confirmed this isn't too frequent for the Sheet's gviz endpoint or the
   reentrancy-guarded tick loop.

All doc cross-references (this file, `../README.md`, `../userGuide_slideRemote.md`)
updated to match — new 2-step flow, 0.5s poll, 10s staleness, and the new
"PowerPoint closed mid-session" recovery path documented as expected
behavior rather than left implicit.

## 5. Possible follow-ups (not built)

- Auto-refresh of the presentation combobox while hooked up (currently
  manual Refresh + once automatically on Lock, to avoid two independent
  loops touching COM concurrently). Round 6's `PresentationUnavailable`
  handling covers the "target disappeared" case specifically, but the
  combobox still won't notice a *new* presentation being opened without a
  manual Refresh.
- `First`/`Last` commands — mentioned as optional in
  `../../Remote/ai/handover_slideRemote.md` §5.4 but the sender never
  actually sends them; not implemented.
- Configurable poll interval (currently a hardcoded ~0.5s as of round 6,
  though `SlidePollingService`'s constructor already accepts an override).
- Configurable staleness window (currently a hardcoded 10s default as of
  round 6, though `CommandDedupeStore`'s constructor already accepts an
  override) — could be exposed in the UI/settings if Gavin wants to tune
  it later.
