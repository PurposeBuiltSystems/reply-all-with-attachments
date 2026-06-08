# Reply All with Attachments — Outlook add-in

A one-click ribbon button on the Outlook message-read surface. Outlook drops
the original attachments when you **Reply** or **Reply All** (unlike Forward).
This button reopens a **Reply All** with those attachments restored.

- **Manifest:** add-in only (XML) — works on Outlook on the web, new Outlook for
  Windows, classic Outlook desktop, and Outlook for Mac.
- **Trigger:** one ribbon button, no task pane.
- **APIs:** `item.getAttachmentContentAsync` (Mailbox 1.8) +
  `displayReplyAllFormAsync` (Mailbox 1.9).
- **No build step, no frameworks** — plain Office.js, served by a tiny Node
  HTTPS server. Only dev dependencies are Microsoft's add-in tooling.

## You do NOT need the M365 Developer Program sandbox

The free sandbox tenant is gated and not required. Sideload against any
Microsoft 365 / Outlook account you already have.

## Prerequisites

- Node.js 18+ (you have v24).
- An Outlook account (Microsoft 365 / Exchange Online).

## First-time setup

```bash
npm install          # installs Microsoft's add-in dev tooling
npm run icons        # generates placeholder ribbon icons in assets/
npm run dev-certs    # installs a trusted localhost HTTPS certificate
```

`npm run dev-certs` will prompt for your password the first time — it adds a
local CA so the browser/Outlook trusts `https://localhost:3000`.

## Run it

1. Start the dev server (leave it running):

   ```bash
   npm start
   ```

   Visit <https://localhost:3000/> once to confirm the cert is trusted.

2. Sideload the add-in:

   - **Desktop (Windows/Mac):** `npm run sideload` — registers the add-in and
     launches Outlook. Stop later with `npm run stop`.
   - **Outlook on the web:** open Outlook → **Settings** (gear) → **General →
     Manage add-ins** (or **Get Add-ins → My add-ins → Add a custom add-in →
     Add from file**) and upload `manifest.xml`.

3. Open any received message → ribbon → **Reply All with Attachments**. A Reply
   All compose window opens with the original attachments attached.

## How it works

`src/commands/commands.js`:

1. Reads `item.attachments`, skipping inline images (the quoted reply keeps
   those automatically).
2. Fetches each attachment's bytes via `getAttachmentContentAsync`.
3. Maps each to a `ReplyFormAttachment`:
   - **Base64** (normal files) → re-attached directly.
   - **Url** (cloud attachments) → re-attached by URL.
   - **Eml / ICalendar** (attached messages/events) → wrapped as a `.eml`/`.ics`.
4. Calls `displayReplyAllFormAsync({ attachments })`.

Attachments that can't be carried over are skipped and reported via an
in-message notification.

## Known limits

- Total reply form data is size-limited; very large attachments may be rejected
  by Outlook. The add-in reports how many were skipped.
- `htmlBody` is capped at 32 KB by the API (we send an empty body; Outlook adds
  the quoted original).
- Requires Mailbox requirement set **1.9**. Very old Outlook builds without 1.9
  won't show the button.

## Publishing (later)

1. Host the contents of this folder on a real HTTPS site (not localhost).
2. Replace every `https://localhost:3000` URL in `manifest.xml` with that host.
3. Replace the placeholder icons in `assets/` with real artwork.
4. Generate a fresh `<Id>` GUID for production.
5. Validate: `npm run validate`. Then deploy via the Microsoft 365 admin center
   (Integrated Apps) or submit to AppSource.

## Project layout

```
manifest.xml              Add-in manifest (ribbon button + metadata)
index.html                Dev-host landing page
server.js                 HTTPS static server for https://localhost:3000
src/commands/commands.html  Function-command runtime page
src/commands/commands.js    The Reply All with Attachments logic
tools/make-icons.js       Generates placeholder PNG icons
assets/                   Generated icons
```
