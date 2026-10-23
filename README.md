# Reply All with Attachments — Outlook add-in

A one-click ribbon button on the Outlook message-read surface. Outlook drops
the original attachments when you **Reply** or **Reply All** (unlike Forward).
This button reopens a **Reply All** with those attachments restored.

- **Manifest:** add-in only (XML) — works on Outlook on the web, new Outlook for
  Windows, classic Outlook desktop, and Outlook for Mac.
- **Trigger:** one ribbon button, no task pane.
- **APIs:** delegated Microsoft Graph creates and populates the draft;
  Office.js opens it after the attachments and signature are verified.
- **Authentication:** nested app authentication through MSAL, with the
  delegated `Mail.ReadWrite` scope.
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

1. Signs in and creates one Reply All draft through Microsoft Graph.
2. Places the saved signature inside the reply's HTML body, placing images beneath the signature text and embedding logos with unique inline content IDs.
3. Copies the original file attachments, following collection paging and
   using upload sessions for large files.
4. Waits for the signature images and copied files to be visible on the draft.
5. Uses Outlook's asynchronous open API where available and waits for its
   callback before completing the command. A definite open failure is retried
   on the same draft. Older clients use the legacy open API.

The add-in creates drafts only. You review and send them yourself. If Outlook
cannot open a saved draft, the notification points to Drafts rather than
requiring another click that would create a second reply.

## Signature images

Paste a signature into the Signature pane and Save. Images appear on their own
lines beneath the signature text. Image display sizes are retained before
compression, and the signature is isolated from the quoted thread's formatting.
Existing saved signatures use this layout automatically when creating a reply.

Review the logo in the saved draft before sending. A missing or broken image
should be checked by reopening the draft from Drafts.

## Checks

Run `npm test` for page wiring, undeclared references, Graph transport and upload
checks, signature HTML regressions, and simulated first-click/open retry tests.
These tests do not connect to your mailbox. A live Outlook smoke test is still
needed for client-specific rendering and sync behavior.

## Known limits

- Exchange mailbox attachment limits still apply to drafts and upload sessions.
- Outlook on the web and new Outlook restrict the native open API to bodies of
  at most 32K characters. If native opening fails, use the saved draft in Drafts.
- Asynchronous open confirmation requires Mailbox 1.9. Older supported desktop
  clients use the legacy API, which cannot report whether the window opened.
- Outlook mobile does not support the native message-open APIs. The reply is
  still saved to Drafts for review.

## Updating the published add-in

The production manifest points to the existing GitHub Pages deployment.
Preserve its `<Id>` when updating this product; a new ID creates a separate
add-in. Validate with `npm run validate`, compare the local manifest version
with the package in Partner Center, and increment the version for a package
update. Upload package changes to the existing offer.

Before release, exercise a saved reply with small and large attachments,
signature images, a delayed draft sync, and a failed native-open attempt in
the actual Outlook clients being claimed. Reopening the same saved draft must
not create another draft or copy attachments again. Mobile saves the result
to Drafts; it does not support the desktop message-open API.

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
