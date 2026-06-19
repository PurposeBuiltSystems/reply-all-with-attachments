# AppSource submission playbook — Reply All with Attachments

Everything you need to submit this add-in to Microsoft AppSource (via Partner
Center). Copy/paste the listing fields, follow the steps, and check off the
items only you can do.

> Timeline reality check: Microsoft manually reviews Outlook add-ins. Expect
> **~2–4 weeks** and possibly a round or two of feedback. This is normal.

---

## A. What's already done (by us)

- ✅ Working, validated add-in (`manifest.xml` passes Microsoft's validator).
- ✅ Builds the reply via **Microsoft Graph** (`createReplyAll` + copy
  attachments) using **delegated `Mail.ReadWrite`** (signed-in user's own mailbox
  only) with **nested app authentication (MSAL)** — no publisher server.
- ✅ Hosted on HTTPS (GitHub Pages).
- ✅ AppSource-compliant icons: **64×64** (`IconUrl`) and **128×128**
  (`HighResolutionIconUrl`) in the manifest.
- ✅ Hosted **Privacy Policy**, **Terms of Use**, and **Support** pages — privacy
  policy describes the Graph data flow and the no-data-collected design.
- ✅ Listing logo (300×300) generated in `assets/icon-300.png`.
- ✅ Multitenant Entra app registered (client ID
  `87764ff9-16e7-4e2f-8164-38eff9f3a895`).

## B. What only YOU can do

- ⬜ **Create a Partner Center account + enroll in the Microsoft 365 and Copilot
  program** (see section F1). Requires: a Microsoft **Entra ID work/school
  account** (NOT a personal @outlook/@gmail account — e.g. an account on the
  purposebuilt.systems tenant), **authority to sign on the company's behalf**,
  and your **legal business name, address, and primary contact** (Microsoft
  verifies these). Enrollment itself is free.
- ✅ Real support email applied to the hosted pages: `Matthew@purposebuilt.systems`.
- ⬜ Capture **1–5 screenshots** of the add-in in Outlook (e.g., the ribbon
  button, and the Reply All window with attachments). Recommended 1366×768.
  You produce these once you've sideloaded and tested it.
- ⬜ Reserve the app name in Partner Center (must match the manifest DisplayName).
- ⬜ Submit and respond to any reviewer feedback.

> **Gating question:** PurposeBuiltSystems must be a real business entity you can
> represent, with an Entra/Microsoft 365 work account. The
> `Matthew@purposebuilt.systems` address suggests you already have that tenant —
> use it. If PurposeBuiltSystems isn't a registered entity, account verification
> is where you'll get stuck.

---

## C. Hosted URLs to paste into the listing

| Field | URL |
| --- | --- |
| Privacy policy | https://purposebuiltsystems.github.io/reply-all-with-attachments/privacy-policy.html |
| Terms of use | https://purposebuiltsystems.github.io/reply-all-with-attachments/terms.html |
| Support / Help | https://purposebuiltsystems.github.io/reply-all-with-attachments/support.html |

## D. Listing copy (paste into Partner Center)

**Name** (≤50 chars, must match manifest DisplayName)
```
Reply All with Attachments
```

**Summary** (≤100 chars)
```
Reply All without losing the original attachments — restore them with one click.
```

**Description** (≤10,000 chars; ~300–500 words recommended)
```
Outlook removes the original attachments when you Reply or Reply All — only
Forward keeps them. That means re-sharing a document with the whole thread
usually takes extra steps: forward the message, re-add everyone, and rewrite
your reply.

Reply All with Attachments fixes this with a single button on the Outlook
message ribbon. Open any email, click Reply All with Attachments, and a normal
Reply All opens with the original attachments already attached. Edit your
message and send — that's it.

KEY BENEFITS
• One click. No forwarding, no re-typing recipients, no manually re-attaching files.
• Keeps the conversation intact. It's a true Reply All, so the thread and all
  recipients are preserved.
• Works everywhere. Because the reply is built with Microsoft Graph, it works on
  Outlook on the web, new and classic Outlook on Windows, and Outlook on Mac.
• Optional signature. Set a reply signature once (text and logo) and it's placed
  at the top of your replies. It's saved only in your own Outlook roaming
  settings — never sent to us.

PRIVACY BY DESIGN — WE COLLECT NO DATA
This add-in collects no data. None. There are no accounts to create, no
analytics, no trackers, no advertising, and no PurposeBuilt Systems server in the
loop. When you click the button, the reply is built directly between your own
mailbox and Microsoft Graph — Microsoft 365's own service — so your email and
attachments never leave the Microsoft 365 boundary and are never seen, stored,
proxied, sold, or shared by us. The add-in uses only the delegated
Mail.ReadWrite permission, which restricts it to your own mailbox and never
touches anyone else's mail. The optional signature is stored only in your own
mailbox's roaming settings. In short: it's a button that does exactly what it
says and keeps nothing.

WHO IT'S FOR
Anyone who regularly replies to threads where the attachment needs to stay with
the conversation — project teams, support and operations staff, legal and
records workflows, and anyone who misses the old "reply with attachments" macro.

HOW TO USE
1. Open a received email that has attachments.
2. Click Reply All with Attachments on the ribbon. The first time, sign in with
   your mailbox account when prompted (one-time consent; later clicks are silent).
3. A Reply All draft opens with the original attachments attached. Send as usual.
```

**Search keywords** (a few, comma-separated)
```
reply all, attachments, reply with attachments, forward, email, attach, productivity
```

**Category:** Productivity
**Industries:** (optional) — General / Productivity
**Products:** Outlook

## E. Notes for certification (REQUIRED — paste into "Notes for certification")

```
TEST STEPS
1. Install/sideload the add-in for any Microsoft 365 mailbox (Outlook on the web,
   new or classic Outlook on Windows, or Outlook on Mac).
2. Open a RECEIVED email that has one or more file attachments (e.g. a PDF).
3. On the message ribbon, click "Reply All with Attachments" (group "Reply All +").
4. FIRST RUN ONLY: a standard Microsoft sign-in appears. Sign in with the current
   mailbox's account and consent to the "Mail.ReadWrite" (delegated) permission.
   This is handled entirely by Microsoft Entra; a regular user can grant this
   consent for their own mailbox. Subsequent clicks are silent (no prompt).
5. EXPECTED: a Reply All draft opens, addressed to the sender and all recipients,
   with the original attachments already attached and the quoted thread intact.
   Edit and send as normal.

OPTIONAL FEATURE — SIGNATURE (not required to test the core feature)
- A second ribbon button, "Signature", opens a task pane where the user can paste
  a reply signature (text and logo). It is stored ONLY in the user's own Outlook
  roaming settings (Microsoft 365) and is added to the top of replies. It is
  never transmitted to or stored by the publisher.

ARCHITECTURE / DATA HANDLING — THE ADD-IN COLLECTS NO DATA
- No publisher/ISV server is involved at any point. The static program files
  (HTML/JS/icons) are served from GitHub Pages; the reply itself is built by
  calling Microsoft Graph (createReplyAll, then copying the original attachments
  onto the draft) DIRECTLY from the user's own client.
- Authentication uses Microsoft nested app authentication (MSAL/NAA). There is no
  separate account, no license key, and no third-party service.
- Permission used: Microsoft Graph "Mail.ReadWrite", DELEGATED only. The add-in
  acts solely as the signed-in user and can access only that user's own mailbox.
  No application-level, mailbox-wide, or other-users' access is requested.
- The add-in collects, logs, stores, transmits, sells, and shares NO user data.
  Email and attachments travel only between the user's mailbox and Microsoft
  Graph, inside the Microsoft 365 service boundary; nothing is retained by the
  publisher.
- Multitenant Microsoft Entra app (client) ID:
  87764ff9-16e7-4e2f-8164-38eff9f3a895
```

---

## F. Step-by-step in Partner Center

### F1. One-time account setup

1. Sign in with your **Entra ID work account** (e.g. on the purposebuilt.systems
   tenant) at the enrollment page:
   https://partner.microsoft.com/dashboard/account/v3/enrollment/introduction/office
2. Provide your **publisher profile**: company info, publisher info, contact info
   (legal business name + address + primary contact — Microsoft verifies these).
3. Accept the **Microsoft Office Agreement** (and Microsoft AI Cloud Partner
   Program agreement if new). You must be authorized to sign for the company.
4. Confirm the **Microsoft 365 and Copilot** program shows as registered under
   Settings (gear) → Account settings → Programs.
5. When you create your **publisher**, name it `PurposeBuiltSystems` so it matches
   the manifest `<ProviderName>` (Microsoft requires them to match closely).

### F2. Create and submit the offer

1. Partner Center → **Marketplace offers** tile → **Microsoft 365 and Copilot**
   tab → **+ New offer → Office Add-in**.
2. **Name:** enter `Reply All with Attachments`, click **Check availability**,
   associate it with the **PurposeBuiltSystems** publisher, **Create**.
   (Publisher can't be changed later.)
3. **Product setup:**
   - **Single sign-on (SSO):** answer **No**. (This question is about Office
     manifest SSO via `WebApplicationInfo`/`getAccessToken`, which we do NOT use.
     We authenticate with MSAL nested app auth instead — different mechanism. The
     sign-in/consent is covered in the certification notes so reviewers expect it.)
   - No to additional purchases, No to Apple Store, No to lead-management.
   - Save draft.
4. **Packages:** upload `manifest.xml`. Wait for **Status = Complete**.
5. **Properties:** pick category **Productivity** (1–3 allowed). For
   **Legal and support info**, the easiest path is to check **Standard Contract**
   (use Microsoft's built-in EULA) — OR paste the Terms URL from section C.
   Then paste the **Privacy policy** URL and **Support** URL from section C.
   (Note: a Terms of Use does NOT count as the privacy policy — you need both,
   and ours already names the app, which certification requires.)
6. **Marketplace listings:** select language English; paste Name, Summary,
   Description, Search keywords from section D; upload the 300×300 logo
   (`assets/icon-300.png`) and your screenshots.
7. **Availability:** schedule "as soon as approved" (the schedule can't be
   changed after first publish); choose markets (all is fine); free.
8. **Notes for certification:** paste section E. IMPORTANT: do NOT write
   "contact me for details" — reviewers cannot contact you, and missing test
   steps auto-fail. Our notes are fully self-contained (no login needed), so
   you're covered.
9. **Review and publish.** Expect a first response in ~3–4 business days; full
   approval typically 2–4 weeks, sometimes with a feedback round. When approved
   it appears in AppSource and you can install it from the store yourself.

## G. Reference

- Submission guide: https://learn.microsoft.com/partner-center/marketplace-offers/add-in-submission-guide
- Pre-submission checklist: https://learn.microsoft.com/partner-center/marketplace-offers/checklist
- Publish Office add-ins to the marketplace: https://learn.microsoft.com/office/dev/add-ins/publish/publish-office-add-ins-to-appsource
- Effective listings: https://learn.microsoft.com/partner-center/marketplace-offers/create-effective-office-store-listings
