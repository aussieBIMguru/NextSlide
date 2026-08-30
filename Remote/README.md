# Remote (sender)

The phone/laptop side of Slide Remote: a button page that silently writes
`Next` / `Previous` / `Go to Slide` commands to a Google Sheet for the
receiver (`../Receiver/`) to pick up. Built entirely out of free Google
tools — no server, no API keys, no hosting.

## Key files

| File | What it is |
|---|---|
| `remote.html` | The button page itself — live source, currently embedded on the Google Site. Edit here, then re-paste into the Site's embed block to publish a change. |
| `purge.gs` | Apps Script bound to the Google Sheet — deletes rows older than 15 minutes on a timer, so the Sheet doesn't grow forever. Edit here, then re-paste into the Sheet's Extensions → Apps Script editor. |
| `webSetup_slideRemote.md` | Full step-by-step build guide — start here if setting this up from scratch or on a new Google account. |
| `ai/handover_slideRemote.md` | Design decisions, exact Form field IDs, and gotchas already paid for — read before changing anything non-trivial. |

Neither `remote.html` nor `purge.gs` deploys automatically — both live
only in Google's own editors (the Site's embed block, the Sheet's Apps
Script project). The copies here are the source of truth; keep them in
sync by hand whenever you edit the live version in Google.

## Setup, in short

1. **Form** — a hidden 3-question Google Form (`Command`, `Slide #`,
   `Session`) whose only job is to append a row to a linked Sheet per
   submission. Must be **Published**, and built on a personal (not
   Workspace) Google account.
2. **Sheet** — auto-created from the Form's Responses link. Gets
   `purge.gs` bound to it on a 15-minute timer to keep it from growing
   forever.
3. **Site** — a Google Site with `remote.html` pasted into an Embed
   code block, published with sharing set to "Anyone with the link."
4. **Test** — open the published Site on a phone, tap a button, confirm
   a row lands in the Sheet.

Full instructions, including where to find the Form's field IDs and how
to wire up the purge trigger, are in `webSetup_slideRemote.md`.

## The moving part: Session

Whoever taps the remote and whichever PC is running the receiver just
need to type the same **Session** name — that's the whole pairing
mechanism, no login or direct connection involved. See
`../README.md` for how this fits together with the receiver side.