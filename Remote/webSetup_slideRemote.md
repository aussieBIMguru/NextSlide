# Slide Remote — Setup Guide

A step-by-step guide to building a phone/laptop remote control for a live PowerPoint presentation, using only free Google tools — a Google Form as a hidden write-only intake, a linked Google Sheet as the shared "table," and a Google Site as the public-facing remote. No servers, no API keys, no hosting.

**Before you start:** do this on a **personal Google account** (plain Gmail), not a company/Workspace account. Workspace accounts can silently block anonymous external form submissions at the organization level, with no per-form setting able to override it.

This guide walks through building the sender from scratch (useful if it ever needs rebuilding, or moving to a new Google account). The two files it produces — the button page and the purge script — are committed in this same folder as the current, live versions: **[`remote.html`](remote.html)** and **[`purge.gs`](purge.gs)**. Day to day, paste those straight in rather than retyping the code blocks below.

---

## Part 1 — Create the Form (the hidden data intake)

The Form itself is never seen by anyone using the remote — it's just the mechanism that writes rows into the Sheet.

1. Go to **forms.google.com** and click **Blank form** (the `+` tile).
2. Name it something internal, e.g. "Slide Remote — Commands." Nobody will see this title.
3. Add three questions, each set to **Short answer**:
   - `Command`
   - `Slide #`
   - `Session`
4. Click the **Settings** gear at the top. Under **Responses**, confirm:
   - **Collect email addresses** — off
   - **Requires sign in** — off
   - **Limit to 1 response** — off
5. Go to the **Responses** tab (next to Questions) and click the green **Sheets** icon → **Create a new spreadsheet** → **Create**. This opens the Sheet that will fill up with one row per submission, plus a free `Timestamp` column.
6. **Publish the form.** Look for a **Publish** button/toggle near the top of the editor. This step is easy to miss and its absence is the single most common cause of a broken setup — an unpublished form rejects every submission from anyone but its owner (it looks like a `401 Unauthorized` error if you're testing programmatically), even with every setting above correct.

---

## Part 2 — Find the field IDs

Every question in a Google Form has a stable internal ID you'll need in order to submit to it programmatically.

1. Back in the Form editor, click the **⋮** (three-dot) menu, top-right → **Get pre-filled link**.
2. A preview of the form opens. Type an obvious dummy value into each field so you can tell them apart later — e.g. `Command` → `TESTCOMMAND`, `Slide #` → `999`, `Session` → `TESTSESSION`.
3. Scroll down, click **Get Link**, then **Copy Link**.
4. Paste the copied link somewhere you can read it. It will look like:
   ```
   https://docs.google.com/forms/d/e/1FAIpQLS.../viewform?usp=pp_url&entry.111111111=TESTCOMMAND&entry.222222222=999&entry.333333333=TESTSESSION
   ```
5. Match each `entry.NNNNNNNNN=` number to the dummy value that follows it. Record them:

   | Field | Entry ID |
   |---|---|
   | Command | `entry.___________` |
   | Slide # | `entry.___________` |
   | Session | `entry.___________` |

6. Build your submission URL by taking that same link's form ID and swapping `/viewform` for `/formResponse`:
   ```
   https://docs.google.com/forms/d/e/{YOUR_FORM_ID}/formResponse
   ```
   This is the URL the button page will silently POST to — it's already public the moment the form is published, with nothing further to expose or deploy.

---

## Part 3 — Build the button page

This is the actual interface people will use — three buttons and a couple of fields, no visible form anywhere. **The current, live version of this file is [`remote.html`](remote.html)** in this folder, already wired up with this project's real Form URL and entry IDs — the steps below are for rebuilding it from scratch on a new Form.

1. Create a plain text file named `remote.html` and paste this in, replacing `FORM_URL` and the three `entry.NNNNNNN` placeholders with your own values from Part 2:

   ```html
   <style>
     html, body { margin: 0; padding: 0; height: 100%; }
     *, *::before, *::after { box-sizing: border-box; }
     .remote {
       display: flex;
       flex-direction: column;
       justify-content: center;
       align-items: center;
       gap: 14px;
       height: 100%;
       padding: 24px;
       font-family: sans-serif;
       overflow: hidden;
     }
     .divider { width: 100%; max-width: 360px; border-top: 1px solid #ddd; margin: 6px 0; }
     .goto-row { display: flex; gap: 8px; width: 100%; max-width: 360px; }
   </style>

   <div class="remote">
     <input id="session" type="text" placeholder="Presentation name (e.g. boardroom-a)"
            style="width:100%; max-width:360px; font-size:1rem; padding:10px;">

     <button onclick="send('Previous')" style="width:100%; max-width:360px; font-size:1.4rem; padding:20px 40px;">◀ Previous</button>
     <button onclick="send('Next')" style="width:100%; max-width:360px; font-size:1.4rem; padding:20px 40px;">Next ▶</button>

     <div class="divider"></div>

     <div class="goto-row">
       <input id="slideNum" type="number" min="1" placeholder="Slide #"
              style="flex:1; font-size:1rem; padding:10px;">
       <button onclick="sendGoTo()" style="flex:1; font-size:1.1rem; padding:10px;">Go to Slide</button>
     </div>

     <div id="status" style="color:#888; font-size:0.9rem; height:20px;"></div>
   </div>

   <script>
   const FORM_URL = "https://docs.google.com/forms/d/e/YOUR_FORM_ID/formResponse";
   const sessionInput = document.getElementById('session');
   const slideNumInput = document.getElementById('slideNum');

   try { sessionInput.value = localStorage.getItem('slideRemoteSession') || ''; } catch (e) {}

   function send(command, slideNum = "") {
     const session = sessionInput.value.trim();
     if (!session) {
       sessionInput.style.border = '2px solid red';
       sessionInput.focus();
       return;
     }
     sessionInput.style.border = '';
     try { localStorage.setItem('slideRemoteSession', session); } catch (e) {}

     const data = new URLSearchParams();
     data.append("entry.YOUR_COMMAND_ID", command);
     data.append("entry.YOUR_SLIDENUM_ID", slideNum);
     data.append("entry.YOUR_SESSION_ID", session);

     fetch(FORM_URL, { method: "POST", mode: "no-cors", body: data });

     const status = document.getElementById("status");
     status.textContent = "Sent " + command + (slideNum ? " → slide " + slideNum : "");
     setTimeout(() => status.textContent = "", 1200);
   }

   function sendGoTo() {
     const slideNum = slideNumInput.value.trim();
     if (!slideNum) {
       slideNumInput.style.border = '2px solid red';
       slideNumInput.focus();
       return;
     }
     slideNumInput.style.border = '';
     send('Go to Slide', slideNum);
   }
   </script>
   ```

2. **Test it in isolation before touching the Site.** Double-click `remote.html` to open it in a browser, tap Next, and check:
   - Open DevTools (F12) → **Network** tab → the `formResponse` request should show status **200**
   - The linked Sheet should show a new row within a second or two

   If you get a `401` here, go back and confirm the form is actually **Published** (Part 1, step 6) — that's the fix in nearly every case.

---

## Part 4 — Create the Site and embed the remote

1. Go to **sites.google.com** and click the **+ (blank)** tile to start a new site.
2. Click the title area and name it, e.g. "Slide Remote."
3. On the page you want the remote to live on, open the **Insert** panel on the right and choose **Embed → Embed code** (not "By URL" — that's for content already hosted elsewhere).
4. Paste in the entire contents of `remote.html` (styles, HTML, and script together) and click **Insert**.
5. Drag the corners of the inserted block to size it — the CSS above fills whatever box you give it without producing a scrollbar.
6. Optional, to reduce the surrounding page chrome: hover the section and use its width toggle to set **Full width**, and in **Page settings** (gear icon) set the header type to **None**. Full edge-to-edge with zero Sites chrome isn't achievable purely through these settings — bookmark the raw page directly instead if that matters to you.
7. Click **Publish** (top-right). The first time, you'll be asked to choose a web address and sharing level — pick **Anyone with the link**.
8. Open the published URL on your phone, tap **Next**, and confirm a new row lands in the Sheet.

---

## Part 5 — Set up automatic purging

This keeps the Sheet from growing forever by clearing old rows on a schedule. It's optional at low volumes but easy to set up once. **The current, live version of this script is [`purge.gs`](purge.gs)** in this folder.

1. Open the linked Sheet, then **Extensions → Apps Script**.
2. Delete any placeholder code and paste in:

   ```javascript
   function purgeOldCommands() {
     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1'); // match your actual tab name
     const data = sheet.getDataRange().getValues();
     const cutoffMinutes = 15;
     const now = new Date();

     for (let i = data.length - 1; i >= 1; i--) { // bottom-up, skip header row
       const rowAge = (now - new Date(data[i][0])) / 60000; // column A = Timestamp
       if (rowAge > cutoffMinutes) sheet.deleteRow(i + 1);
     }
   }
   ```

3. Click the **save icon** (or Ctrl+S). The first save asks you to name the project — call it something like "Slide Remote Purge."
4. Confirm `purgeOldCommands` now appears in the function dropdown next to the **Run** (▶) button in the toolbar — if it doesn't, the save didn't take; save again.
5. Click the **clock icon** (Triggers) in the left sidebar → **Add Trigger**, and set:
   - Choose which function to run → `purgeOldCommands`
   - Choose which deployment should run → Head
   - Select event source → Time-driven
   - Select type of time based trigger → Minutes timer
   - Select minute interval → Every 15 minutes
6. Click **Save**. You'll get a Google authorization prompt — this is expected, since the script needs permission to edit the sheet on a schedule with nobody watching. You'll likely see a **"Google hasn't verified this app"** screen; this is normal for a personal script only you use. Click **Advanced** → **Go to [project name] (unsafe)** → review the permissions → **Allow**.

That's a one-time approval — after this it runs quietly every 15 minutes on its own.

> **Keep this cutoff comfortably ahead of however often the receiver polls the sheet** (see `../Receiver/README.md`), and make sure the receiver tracks "what's new" by the row's Timestamp, not by row position — a purge run shifts every row above it, which would silently break any tracking based on row count or index.

---

## Part 6 — End-to-end test

1. Open the published Site URL on a phone.
2. Type a Session name, e.g. `test`.
3. Tap **Next** — confirm a row appears in the Sheet with `Next`, a blank Slide #, and `test` in Session.
4. Enter a number in **Slide #** and tap **Go to Slide** — confirm a row appears with `Go to Slide` and that number.
5. Leave it alone for 20+ minutes and check that the oldest test rows have disappeared, confirming the purge trigger is running.

At this point the sender side is fully working. The remaining piece — the receiver app that polls this Sheet and actually drives PowerPoint — is `Receiver/`, in this same repository; see `../Receiver/README.md`.

---

## Troubleshooting quick reference

| Symptom | Likely cause |
|---|---|
| `401` on the `formResponse` request | Form isn't Published (most common) — or, on a Workspace account, org-level restriction with no per-form fix |
| No row appears, no error shown | Check the actual field values match your `entry.NNNNNNN` mapping exactly — a typo sends data to the wrong (or no) field |
| Embedded page shows a scrollbar on the Site | Missing the `html, body { height:100%; margin:0 }` reset, or the embed box is sized smaller than the content needs |
| `purgeOldCommands` doesn't appear in the Trigger dropdown | The script wasn't saved/named yet — save it first, then return to Triggers |
| Rows aren't being purged | Confirm the sheet/tab name in the script matches exactly, and that the trigger is enabled and hasn't errored (check the Apps Script **Executions** log) |
