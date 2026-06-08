# AppSource submission playbook — Reply All with Attachments

Everything you need to submit this add-in to Microsoft AppSource (via Partner
Center). Copy/paste the listing fields, follow the steps, and check off the
items only you can do.

> Timeline reality check: Microsoft manually reviews Outlook add-ins. Expect
> **~2–4 weeks** and possibly a round or two of feedback. This is normal.

---

## A. What's already done (by us)

- ✅ Working, validated add-in (`manifest.xml` passes Microsoft's validator).
- ✅ Minimal permission (`read item`) — easiest path through review.
- ✅ Hosted on HTTPS (GitHub Pages).
- ✅ AppSource-compliant icons: **64×64** (`IconUrl`) and **128×128**
  (`HighResolutionIconUrl`) in the manifest.
- ✅ Hosted **Privacy Policy**, **Terms of Use**, and **Support** pages.
- ✅ Listing logo (300×300) generated in `assets/icon-300.png`.

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
• Smart about inline images. Pictures embedded in the message body stay in the
  quoted reply automatically and aren't duplicated.
• Handles large threads. Files, cloud attachments, and attached messages are
  carried over; anything too large for Outlook's reply limit is skipped with a
  clear notice.

PRIVACY BY DESIGN
The add-in has no servers, no accounts, and no tracking. It collects no data.
Reading attachments and building the reply happen entirely inside Outlook, on
your own device — your email and attachments are never sent anywhere. It
requests only the minimum "read item" permission.

WHO IT'S FOR
Anyone who regularly replies to threads where the attachment needs to stay with
the conversation — project teams, support and operations staff, legal and
records workflows, and anyone who misses the old "reply with attachments" macro.

HOW TO USE
1. Open a received email that has attachments.
2. Click Reply All with Attachments on the ribbon.
3. The Reply All window opens with the attachments attached. Send as usual.
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
1. Open Outlook (web, new Windows, classic Windows, or Mac) with any mailbox.
2. Open a RECEIVED email message that has one or more file attachments.
   (Any test message with, e.g., a PDF or image attachment works.)
3. On the message ribbon, click the "Reply All with Attachments" button
   (group "Reply All +").
4. EXPECTED: a Reply All compose window opens addressed to the sender and all
   recipients, with the original attachments already attached.

NOTES
- No sign-in, account, license key, or external service is required.
- The add-in collects/transmits no data; all work happens locally via the
  Office.js APIs getAttachmentContentAsync and displayReplyAllFormAsync.
- Permission requested: read item (minimum).
- Inline images in the original body are intentionally not re-attached (they
  remain in the quoted reply). Attachments too large for Outlook's reply-form
  limit are skipped and the user is notified.
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
3. **Product setup:** answer No to the SSO/Entra question, No to additional
   purchases, No to Apple Store, No to lead-management. Save draft.
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
