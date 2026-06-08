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

- ⬜ Create a **Partner Center** account and enroll in the **Microsoft 365 and
  Copilot** (marketplace) program. https://partner.microsoft.com/
- ⬜ Provide a **real support email** (replaces the placeholder in the hosted
  pages — tell me the address and I'll patch it).
- ⬜ Capture **1–5 screenshots** of the add-in in Outlook (e.g., the ribbon
  button, and the Reply All window with attachments). Recommended 1366×768.
  You produce these once you've sideloaded and tested it.
- ⬜ Reserve the app name in Partner Center (must match the manifest DisplayName).
- ⬜ Submit and respond to any reviewer feedback.

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

1. Go to https://partner.microsoft.com/ and create/sign in to a Partner Center
   account. Complete the (free) **Microsoft 365 and Copilot** marketplace
   enrollment. Account verification can take a few days.
2. **Marketplace offers → New offer → Office add-in** (a.k.a. Microsoft 365 app).
3. **Reserve your offer name:** `Reply All with Attachments` (matches the
   manifest DisplayName).
4. **Packages:** upload `manifest.xml`. Partner Center validates it on upload.
5. **Marketplace listing:** paste the Name, Summary, Description, Search
   keywords from section D; upload the 300×300 logo and your screenshots; paste
   the Privacy / Terms / Support URLs from section C.
6. **Notes for certification:** paste section E.
7. **Availability/Plans:** set it free and choose markets (you can pick all).
8. **Review and publish.** Microsoft runs automated + manual validation.
   Respond to any feedback; resubmit if asked. When approved, it appears in
   AppSource and you can install it from the store yourself.

## G. Reference

- Submission guide: https://learn.microsoft.com/partner-center/marketplace-offers/add-in-submission-guide
- Pre-submission checklist: https://learn.microsoft.com/partner-center/marketplace-offers/checklist
- Publish Office add-ins to the marketplace: https://learn.microsoft.com/office/dev/add-ins/publish/publish-office-add-ins-to-appsource
- Effective listings: https://learn.microsoft.com/partner-center/marketplace-offers/create-effective-office-store-listings
