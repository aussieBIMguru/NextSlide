# Slide Remote

A phone/laptop remote control for a live PowerPoint presentation, built
from two independent halves that never talk to each other directly — a
public web page and a Windows app — joined only by a shared Google Sheet
and a typed Session name.

```
Your phone (or any browser)              This PC
┌───────────────────────┐                ┌─────────────────────────┐
│  Sender page            │   writes a   │  NextSlide (tray/window)  │
│  (Google Site)          │   row per    │  polls the Sheet ~1×/sec  │
│  ◀ Previous  Next ▶     ─tap─▶ Google   │  for rows matching its    │
│  Go to Slide [__]       │   Sheet      │  Session, then drives     │
│  Session: [boardroom-a] │              │  PowerPoint accordingly   │
└───────────────────────┘                └────────────┬─────────────┘
                                                        ▼
                                              Live PowerPoint
                                              Slide Show on this PC
```

- **Sender** — `Remote/`: a button page (Next / Previous / Go to Slide +
  a Session field) published on a Google Site, backed by a hidden Google
  Form that silently appends each tap as a row to a Google Sheet. No
  server, no API keys, no hosting of its own. Fully built and tested.
- **Receiver** — `Receiver/`: NextSlide, a Windows WPF app. Polls that
  same Sheet roughly once a second, filters to whichever Session it's
  locked to, and drives a live PowerPoint Slide Show via COM automation.
  Fully built and tested.

The two sides are deliberately decoupled: there's no pairing, login, or
direct network connection between the phone and the PC. Anyone with the
sender link and the same Session name can drive whichever PC is watching
that Session — which is also what lets one Sheet serve several
rooms/presentations at once, each receiver locked to its own Session.

## Where to go next

| I want to... | Read |
|---|---|
| Actually run a presentation with both apps | **[`Receiver/userGuide_slideRemote.md`](Receiver/userGuide_slideRemote.md)** |
| Build or rebuild the sender (Form/Sheet/Site) | **[`Remote/webSetup_slideRemote.md`](Remote/webSetup_slideRemote.md)** |
| Understand the sender's design decisions, gotchas, exact field IDs | **[`Remote/ai/handover_slideRemote.md`](Remote/ai/handover_slideRemote.md)** |
| Build or work on NextSlide (the receiver) | **[`Receiver/README.md`](Receiver/README.md)** |
| See the receiver's design decisions and build/run history | **[`Receiver/ai/handover_slideRemote_receiver.md`](Receiver/ai/handover_slideRemote_receiver.md)** |

## Repository layout

```
README.md                        This file — project overview

Receiver/                        The PC-side app (NextSlide)
  README.md                      Build/structure docs (developer-facing)
  userGuide_slideRemote.md       End-to-end usage guide (both apps, non-developer)
  src/
    NextSlide.sln
    NextSlide/                   The WPF project itself
  ai/
    handover_slideRemote_receiver.md   Design decisions, build/run feedback log
    (your original brief to Claude)

Remote/                          The sender (Form / Sheet / Site)
  remote.html                    The button page — live source, embedded on the Site
  purge.gs                       Apps Script — purges rows older than 15 min
  webSetup_slideRemote.md        How to build the Form/Sheet/Site from scratch
  ai/
    handover_slideRemote.md      Design decisions, field IDs, gotchas
```

`Remote/remote.html` and `Remote/purge.gs` are the actual, current source
for the two pieces of code the sender side runs — not just documentation.
`remote.html`'s contents get pasted whole into the Google Site's embed
block; `purge.gs` gets pasted into the Sheet's bound Apps Script project.
Neither one deploys from this repo automatically — see
`Remote/webSetup_slideRemote.md` for where each one goes.

## The contract between the two halves

Everything that connects the sender and receiver comes down to one shared
Google Sheet, written by the sender's Form and read by NextSlide:

- **Columns:** `Timestamp` (added automatically by the Form link),
  `Command` (`Next` / `Previous` / `Go to Slide`), `Slide #`, `Session`.
- **Session** is free text either side agrees on — it's the only "pairing"
  mechanism that exists.
- The Sheet **auto-purges rows older than 15 minutes** (`Remote/purge.gs`,
  see `Remote/webSetup_slideRemote.md` Part 5) — kept comfortably ahead of
  NextSlide's own ~15-second staleness window, and both sides track rows
  by **Timestamp**, never row position, since a purge run shifts
  everything above it.
- The sender **never learns whether a command actually fired** — its
  "Sent" flash is optimistic (see `Remote/ai/handover_slideRemote.md`
  §4); NextSlide's own command log, on the PC, is the real record of
  what happened.

## Status

Both halves are built and have been run for real: the sender end-to-end
from day one, and NextSlide through three rounds of build/run feedback
(see `Receiver/README.md`'s Changelog) plus a documentation and app-icon
polish pass. No open bugs as of 2026-08-30.
