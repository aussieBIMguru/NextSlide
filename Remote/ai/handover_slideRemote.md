# Slide Remote — Sender Handoff

**Status:** Sender side is fully built and tested end‑to‑end. The receiver (the PC‑side app that reads commands and drives PowerPoint) is `Receiver/`, in this same repository — see `../../Receiver/README.md` for its build docs and `Receiver/ai/handover_slideRemote_receiver.md` for its decision log. This document captures the sender side as actually built. The two files it produced now live as the real, current source in this repo: **[`../remote.html`](../remote.html)** and **[`../purge.gs`](../purge.gs)** — the copies below are historical record, not the live version.

---

## 1. Architecture, as actually built

```
Phone / laptop (Google Site, public button page)
   │  silent background POST, no page reload
   ▼
Google Form  (hidden — nobody ever sees its real page; write‑only intake)
   │  auto-appends a row per submission
   ▼
Google Sheet ("the table")
   │  Receiver polls this, reads new rows
   ▼
Receiver (PC)
   │  COM automation
   ▼
Live PowerPoint Slide Show
```

Note: an earlier design pass explored Airtable instead of Google Forms/Sheets. The team pivoted to Forms specifically to get an instant, no-visible-submit "clicker" feel via a silent POST (`fetch(url, {mode:'no-cors'})`) — Airtable couldn't offer that without either exposing a write‑scoped API token client‑side or gambling on undocumented CORS behaviour. Any earlier Airtable‑based sketch/diagram is superseded by this document.

---

## 2. The table (Google Sheet)

- Auto‑created from the Form's **Responses → Sheets** link.
- Columns actually in use: **Timestamp** (added automatically by the Form link), **Command**, **Slide #**, **Session**.
- No `Sender` or `Status` column was built — those were discussed early on but dropped for simplicity. Add them later only if you want a visible audit trail; the current design deliberately avoids needing one (see §5).
- **Purge:** a time‑driven Apps Script trigger bound to the sheet, function `purgeOldCommands` (source: `../purge.gs`), runs every 15 minutes and deletes any row whose Timestamp is more than 15 minutes old. Runs under the sheet owner's own Google identity — no external auth needed.

⚠️ **This purge cutoff is a hard dependency for the receiver's design** — see §5.

---

## 3. The Form (hidden data intake)

- 3 questions, all **Short answer** type: `Command`, `Slide #`, `Session`.
- Settings confirmed: "Collect email addresses" **off**, "Limit to 1 response" **off**, "Requires sign in" **off**.
- Must be **Published** — an unpublished form silently rejects every external submission with a `401`, even with the above settings correct. This was the actual root cause the one time setup broke.
- Built under a **personal** Google account. A Workspace/company account can block anonymous external submissions at the organization level with no per‑form setting to override it — if this ever gets rebuilt on a work account and mysteriously 401s again with everything above looking right, that's almost certainly why.

**Submission endpoint** (swap `/viewform` for `/formResponse` on the form's normal link):
```
https://docs.google.com/forms/d/e/FORMID/formResponse
```

**Field → entry ID mapping** (from the "Get pre‑filled link" tool — reusable, doesn't change unless the form's questions are deleted and rebuilt):

| Field | Entry ID |
|---|---|
| Command | `entry.####` |
| Slide # | `entry.####` |
| Session | `entry.####` |

---

## 4. The sender page (embedded on the Google Site)

Live source: `../remote.html`. Embedded on the Site via Insert → Embed → Embed code (the whole file — styles, HTML, and script together, pasted as one block).

Behaviour notes:
- Session name is remembered per‑device via `localStorage` so a returning sender doesn't retype it — this is **not** shared between senders/devices.
- Because submissions use `mode: 'no-cors'`, the page can **never** read whether a POST actually succeeded — the on‑screen "Sent ✓" flash is optimistic, not a confirmation. The Sheet gaining a row (and eventually the slide changing) is the real confirmation.
- Layout uses a height:100%/flex reset to avoid the Sites embed box scrolling. Google Sites' own page margins/header can be minimized (Full width section, Header: None) but not fully removed through settings — a truly edge‑to‑edge feel would mean bookmarking this page's raw URL directly instead of going through Sites.

---

## 5. The receiver — how it actually turned out

The following was the spec for the not-yet-built receiver as of this section's original writing; `Receiver/` (see `../../Receiver/README.md`) now implements all of it:

1. **Read side** — polls the Sheet with no auth needed via the gviz endpoint, as long as the sheet is shared "anyone with the link can view":
   ```
   https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:json[&gid={GID}]
   ```
2. **Tracks "last processed" by Timestamp, never row position or count** — required both to survive restarts safely and to stay correct once the purge script deletes old rows out from under it. Implemented as an in-memory hash table, deliberately not persisted to disk (see the receiver's own decision log for why an earlier disk-persisted version was reworked).
3. **Filters to its own Session value** before acting on a row — this is what lets one Sheet serve multiple concurrent presentations/rooms.
4. **Maps Command (+ Slide #) to a PowerPoint action** — `Next`, `Previous`, `Go to Slide` (with the number), via COM automation. `First`/`Last` were left unimplemented since the sender never actually sends them.
5. **Polling interval:** 1 second.
6. **Not built, and not needed:** writing a "Done" status back to the Sheet. The local‑pointer approach in step 2 makes this unnecessary, and adding it later would require real Sheets API write credentials (OAuth or a service account) — something this design has deliberately avoided.

---

## 6. Gotchas already paid for — don't rediscover these

- An **unpublished** Form gives a `401` to everyone but its owner, with no other obvious symptom.
- A **Workspace/company** Google account can block anonymous form submissions at the org level, invisibly — use a personal account for this.
- `no-cors` POST responses are permanently unreadable by design. Don't build anything that depends on inspecting the fetch response.
- The purge cutoff (15 min) must always stay well ahead of the receiver's poll interval, and the receiver must track by **timestamp**, never row position.

---

## 7. Before sharing this code on git

Redact `FORM_URL` and the three `entry.NNNNNNN` IDs in `../remote.html` (and this document) before committing this repository publicly — they're not secret in the sense of granting access to anything sensitive (worst case is spam submissions to the Form), but there's no reason to publish them either.
