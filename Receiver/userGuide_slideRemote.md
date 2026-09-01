# Slide Remote — End-to-End Guide

How the sender page and the NextSlide receiver work together, and how to
run an actual presentation with them. This is the "just tell me how to
use the thing" doc — for the project overview see `../README.md`; for how
either half was *built*, see `../Remote/webSetup_slideRemote.md` /
`../Remote/ai/handover_slideRemote.md` (the sender/Sheet/Form side) and
`README.md` (NextSlide's own developer docs, same folder as this file).

---

## 1. The whole picture

```
Your phone (or any browser)              This PC
┌───────────────────────┐                ┌─────────────────────────┐
│  Sender page           │   writes a     │  NextSlide (tray/window) │
│  (Google Site)          │   row per   │  polls the Sheet ~2×/sec  │
│  ◀ Previous  Next ▶      ──tap──▶ Google │  for rows matching its   │
│  Go to Slide [__]       │   Sheet    │  Session, then drives      │
│  Session: [boardroom-a]│                │  PowerPoint accordingly  │
└───────────────────────┘                └────────────┬─────────────┘
                                                        ▼
                                              Live PowerPoint
                                              Slide Show on this PC
```

Two independent pieces, joined only by a **Session name** both sides
type in — there's no pairing, login, or network connection between the
phone and the PC directly. Whoever's clicking the remote and whichever
PC is watching just need to agree on the same Session text (anything you
like, e.g. `boardroom-a`) and both point at the same Google Sheet.

Everything both sides need is already built and working:

- **Sender** — a page you open on your phone/laptop browser (already
  published on a Google Site). Nothing to install.
- **Receiver** — NextSlide, a Windows app that runs on the PC actually
  running the presentation. Built from source (see `README.md`) —
  install/build it once, then just run it each time.

## 2. One-time setup (per side, done once)

### Sender side (already done)

The Google Form, Sheet, and Site are already built and published — see
`../Remote/webSetup_slideRemote.md` if this ever needs rebuilding or
moving to a new Google account. Day to day, nothing here needs touching;
skip to §3.

### Receiver side (once per PC)

1. Build NextSlide (see `README.md` → Getting Started) — needs Visual
   Studio and PowerPoint installed on whichever PC will actually run
   presentations. Zero NuGet packages, so this is a plain F5 once the
   solution's open.
2. That's it for setup — there's no install wizard, no config file to
   edit by hand. Everything else below happens from NextSlide's own
   window each time you use it.

Optional: if you want NextSlide always ready in the background, add a
shortcut to it in your Startup folder (Win+R → `shell:startup`) — it
opens in tray-friendly mode by default (closing the window just hides
it; the app keeps polling until you actually choose Exit from the tray
icon).

## 3. Running a presentation (every time)

Do this in order — each step unlocks the next.

### On the PC (NextSlide)

1. Open PowerPoint and your deck, if it isn't already open.
2. Launch NextSlide (or bring it up from the tray if it's already
   running).
3. Type a **Session** name — anything, as long as it matches what you'll
   type on the phone in a moment (e.g. `boardroom-a`) — and paste the
   **Google Sheet URL** (the sheet the Form writes to — ask whoever set
   up the sender side for the link, or grab it from the sheet's own
   address bar; it just needs to be shared "Anyone with the link can
   view"). Both fields fill in from last time automatically. Click
   **Lock** once both look right — it only enables once the Session name
   is filled in and the URL looks like a real Sheets link.
4. In **Hook to PowerPoint**, pick your deck from the dropdown. If it's
   not listed, click **Refresh** (it only lists decks that are open —
   it doesn't need to be in Slide Show mode yet to appear, just open).
5. Start the slide show in PowerPoint itself (**F5**, or **Slide Show →
   From Beginning**). NextSlide only actually drives a deck once
   PowerPoint says it's presenting — locking it in ahead of time is fine.

At this point the status bar at the bottom shows a green dot and
"Watching session '...' — hooked to '...'" — NextSlide is now live.

### On the phone (or any browser)

1. Open the sender page (the published Google Site link).
2. Type the **same Session name** you locked in on the PC — exactly the
   same text (case doesn't matter, spacing does).
3. Tap **Next** / **Previous**, or enter a number and tap **Go to
   Slide**.

Commands land within about a second. Every one that fires shows up at
the top of NextSlide's log, green for success; anything that couldn't be
sent shows red with a reason (see Troubleshooting in `README.md` for
what the common ones mean).

### Wrapping up

- Leaving the deck: just stop the Slide Show in PowerPoint (Esc) —
  NextSlide keeps watching the Sheet, it just won't have anything to
  drive until a show is running again.
- Done for the day: click **Release** in NextSlide (stops polling; the
  Sheet URL is remembered for next time), or just close the app — it
  drops to the tray by default, or right-click the tray icon → **Exit**
  to actually quit.
- Multiple rooms/presentations at once: run one NextSlide per PC, each
  locked to its own distinct Session name, all pointed at the *same*
  Sheet — each only ever reacts to its own Session's rows.

## 4. Quick troubleshooting

| Symptom | Check |
|---|---|
| Tapping buttons on the phone does nothing on the PC | Session names match exactly on both sides? NextSlide's status bar shows a green "watching" line? Slide Show actually started in PowerPoint (not just the deck open)? |
| PowerPoint isn't in the "Hook to" list | It needs to be open (not necessarily presenting yet) — click Refresh in NextSlide after opening it. |
| NextSlide's log shows red "Not currently in Slide Show mode" | Normal if you click the remote before pressing F5 in PowerPoint, or after leaving the show with Esc — nothing's broken, the click just arrived while nothing was presenting. Start (or resume) the Slide Show and try again. |
| A click on the phone shows "Sent" but nothing happens | That's expected and not an error — the sender can't confirm delivery (see `../Remote/ai/handover_slideRemote.md` §4), so "Sent" just means the tap registered locally. Check NextSlide's log for what actually happened. |
| Everything worked, then suddenly stopped after several minutes idle | The Sheet auto-purges rows older than 15 minutes — this doesn't affect *new* clicks, only very old unactioned rows, so this is very unlikely to be the cause; more likely check the PC's network connection. |
| Closed PowerPoint (or the deck) without clicking Release first | NextSlide notices and recovers on its own — no crash, and you don't need to re-lock. The status bar and log explain PowerPoint's gone, and the presentation dropdown clears. Reopen PowerPoint/the deck, click **Refresh**, and pick it again. |

For anything not covered here, `README.md`'s own Troubleshooting table
covers NextSlide-specific errors in more depth, and
`../Remote/webSetup_slideRemote.md`'s Troubleshooting table covers
sender-page/Form issues (mostly relevant only if the sender side itself
is being rebuilt).
