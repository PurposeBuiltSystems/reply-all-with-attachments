/*
 * Reply All with Attachments — Microsoft Graph implementation (v2).
 *
 * WHY Graph: Outlook's client-side reply APIs can't reliably re-attach files on
 * Outlook on the web / New Outlook (office-js #4599). Doing it server-side via
 * Microsoft Graph works on every platform: we create the reply-all DRAFT and
 * copy the original attachments onto it through Graph, then open the draft.
 *
 * AUTH: Nested App Authentication (NAA) via MSAL — no backend server. Requires
 * an Entra app registration (see GRAPH-SETUP below) and the Mail.ReadWrite
 * delegated permission. With admin consent granted, token acquisition is silent
 * so the button stays one-click.
 *
 * GRAPH-SETUP: paste your Entra app's Application (client) ID into CLIENT_ID.
 */

/* global Office, msal */

var CLIENT_ID = "87764ff9-16e7-4e2f-8164-38eff9f3a895"; // Entra app registration (PurposeBuiltSystems)
var GRAPH = "https://graph.microsoft.com/v1.0";
var SCOPES = ["Mail.ReadWrite"];

var pcaPromise = null;

Office.onReady(function () {});
if (Office.actions && Office.actions.associate) {
  Office.actions.associate("replyAllWithAttachments", replyAllWithAttachments);
}

/** Lazily create the MSAL nestable client (NAA). */
function getPca() {
  if (!pcaPromise) {
    pcaPromise = msal.createNestablePublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: "https://login.microsoftonline.com/common",
      },
    });
  }
  return pcaPromise;
}

/**
 * Sign-in must never hang. An un-timed await on the popup flow leaves the
 * ribbon button doing nothing visible, with no way for the user to tell
 * whether it is working or stuck.
 */
function withTimeout(promise, ms, message) {
  var timer;
  return Promise.race([
    promise.then(function (v) { clearTimeout(timer); return v; },
                 function (e) { clearTimeout(timer); throw e; }),
    new Promise(function (_, reject) {
      timer = setTimeout(function () { reject(new Error(message)); }, ms);
    }),
  ]);
}

/** Get a Graph token: silent first, interactive only if needed. */
async function getToken() {
  var pca = await withTimeout(getPca(), 20000,
    "Sign-in didn't start. Fully quit Outlook (Cmd+Q) and reopen, then try again.");
  try {
    return (await withTimeout(pca.acquireTokenSilent({ scopes: SCOPES }), 20000, "silent timeout")).accessToken;
  } catch (e) {
    var interactive = await withTimeout(
      pca.acquireTokenPopup({ scopes: SCOPES }), 120000,
      "Sign-in didn't finish \u2014 a Microsoft sign-in window may have opened behind Outlook. " +
      "Check for it, finish signing in, and click again. If none appeared, fully quit Outlook " +
      "(Cmd+Q), reopen, and retry.");
    return interactive.accessToken;
  }
}

/** Thin Graph fetch helper. */
function delay(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

async function graph(token, method, path, body) {
  var res = await fetch(GRAPH + path, {
    method: method,
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    var text = await res.text();
    throw new Error("Graph " + method + " " + path + " → " + res.status + " " + text);
  }
  return res.status === 204 ? null : res.json();
}

/** @param {Office.AddinCommands.Event} event */
async function replyAllWithAttachments(event) {
  try {
    var item = Office.context.mailbox.item;
    // Office gives an EWS id; Graph needs a REST id.
    var restId = Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );

    var token = await getToken();

    // 1. Create the Reply All draft (Graph omits original attachments, like Outlook).
    var draft = await graph(token, "POST", "/me/messages/" + restId + "/createReplyAll", {});

    // 1b. Tidy the quoted-thread header (Graph mashes the divider + From/Sent/
    //     To/Subject onto one line for plain-text originals) and, if the user
    //     saved a signature, insert it above the quoted thread — embedding its
    //     logo as an inline cid attachment so it renders reliably (some clients
    //     strip raw data: URIs in mail).
    var sig = readSignature(Office.context.roamingSettings);
    var d = await graph(token, "GET", "/me/messages/" + draft.id + "?$select=body");
    var body = (d && d.body && d.body.content) || "";
    var tidied = tidyQuotedHeader(body);
    var content = tidied;
    var images = [];
    if (sig) {
      var ex = extractInlineImages(sig);
      images = ex.images;
      content = ex.html + "<br><br>" + tidied;
    }
    if (sig || tidied !== body) {
      for (var k = 0; k < images.length; k++) {
        var im = images[k];
        await graph(token, "POST", "/me/messages/" + draft.id + "/attachments", {
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: im.name,
          contentType: im.contentType,
          contentBytes: im.contentBytes,
          isInline: true,
          contentId: im.cid,
        });
      }
      await graph(token, "PATCH", "/me/messages/" + draft.id, {
        body: { contentType: "HTML", content: content },
      });
    }

    // 2. Fetch the original message's file attachments (skip inline images).
    //    isInline alone is unreliable — many clients embed signature logos
    //    without setting it — so also treat any attachment whose contentId is
    //    referenced as cid: in the quoted body as inline.
    var attachments = await graph(token, "GET", "/me/messages/" + restId + "/attachments");
    var files = (attachments.value || []).filter(function (a) {
      if (a["@odata.type"] !== "#microsoft.graph.fileAttachment") { return false; }
      if (a.isInline) { return false; }
      if (a.contentId && content.indexOf("cid:" + a.contentId) !== -1) { return false; }
      return true;
    });

    // 2b. createReplyAll already carries the inline images the quoted thread
    //     needs, and some senders attach the same file twice — skip anything
    //     already on the draft (or already copied) so nothing lands twice.
    var existing = await graph(
      token,
      "GET",
      "/me/messages/" + draft.id + "/attachments?$select=name,size"
    );
    var seen = {};
    (existing.value || []).forEach(function (a) { seen[a.name + "|" + a.size] = true; });

    // 3. Copy each onto the draft. (Large attachments may need their own GET for
    //    contentBytes; handled in a later pass if needed.)
    var skipped = 0;
    var copied = {};
    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      var key = f.name + "|" + f.size;
      if (seen[key]) { continue; }
      seen[key] = true;
      var bytes = f.contentBytes;
      if (!bytes) {
        // Collection didn't include the bytes — fetch the single attachment.
        var full = await graph(token, "GET", "/me/messages/" + restId + "/attachments/" + f.id);
        bytes = full && full.contentBytes;
      }
      if (!bytes) { skipped++; continue; }
      await graph(token, "POST", "/me/messages/" + draft.id + "/attachments", {
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: f.name,
        contentType: f.contentType,
        contentBytes: bytes,
      });
      copied[key] = true;
    }

    // 3b. Confirm the server actually has them before opening anything.
    //
    //     Reported symptom: the first attempt sometimes opens a reply with no
    //     attachments, and repeating it works. Every call above is awaited and
    //     graph() throws on any non-2xx, so a silent copy failure isn't
    //     possible - which points at the draft being opened before Outlook has
    //     caught up with the copies. Opening early is the one thing here we
    //     control, so don't: re-read the draft until the server reports what we
    //     expect, and only then open it.
    var expected = 0;
    for (var q = 0; q < files.length; q++) {
      if (copied[files[q].name + "|" + files[q].size]) { expected++; }
    }
    var present = 0;
    if (expected > 0) {
      for (var attempt = 0; attempt < 5; attempt++) {
        var check = await graph(
          token, "GET", "/me/messages/" + draft.id + "/attachments?$select=name,size,isInline"
        );
        present = (check.value || []).filter(function (a) {
          return !a.isInline && copied[a.name + "|" + a.size];
        }).length;
        if (present >= expected) { break; }
        await delay(400 * (attempt + 1));   // 0.4s, 0.8s, 1.2s, 1.6s
      }
    }

    // 4. Open the populated draft for the user to review/send. Opening a
    //    message form isn't available on every client (Outlook mobile) —
    //    the draft already exists either way, so point at Drafts instead of
    //    reporting a failure.
    try {
      var ewsId = Office.context.mailbox.convertToEwsId(
        draft.id,
        Office.MailboxEnums.RestVersion.v2_0
      );
      Office.context.mailbox.displayMessageForm(ewsId);
    } catch (openErr) {
      notify("info", "Reply All draft created with attachments - open your Drafts folder to review and send.");
    }

    if (skipped > 0) {
      notify("info", skipped + " attachment(s) could not be copied and were skipped.");
    } else if (expected > 0 && present < expected) {
      // The copies were accepted but the server still isn't reporting them.
      // Say so plainly rather than let the user discover an empty reply.
      notify("info", "Attachments were copied but Outlook may not show them yet - " +
        "close the draft and reopen it from Drafts before sending.");
    }
    finish(event);
  } catch (e) {
    notify("error", "Reply All with Attachments failed: " + (e && e.message ? e.message : e));
    finish(event);
  }
}

// Graph's createReplyAll collapses the quoted-thread header onto a single line
// for plain-text originals (e.g. "____ From: … Sent: … To: … Subject: …").
// Drop the divider onto its own line and put each English field label on its
// own line. Best-effort and English-only — leaves other layouts/locales as-is.
// Only run on the quoted body, never on the user's signature.
function tidyQuotedHeader(html) {
  return html
    .replace(/(_{5,})/g, "$1<br>")
    .replace(/ (From|Sent|To|Cc|Bcc|Subject):/g, "<br>$1:");
}

// Pull base64 data-URI images out of the signature HTML and rewrite them to
// reference inline attachments by cid. Returns the rewritten HTML plus the list
// of images to POST as inline attachments on the draft.
function extractInlineImages(html) {
  var images = [];
  var n = 0;
  function take(ctype, b64) {
    n++;
    var cid = "raasig" + n;
    var ext = (/\/(\w+)/.exec(ctype) || [])[1] || "png";
    images.push({ cid: cid, contentType: ctype, contentBytes: b64, name: "image" + n + "." + ext });
    return cid;
  }
  // Handles src="data:..." and src='data:...'
  var out = html.replace(
    /src\s*=\s*"data:([^;]+);base64,([^"]+)"/gi,
    function (m, ctype, b64) { return 'src="cid:' + take(ctype, b64) + '"'; }
  );
  out = out.replace(
    /src\s*=\s*'data:([^;]+);base64,([^']+)'/gi,
    function (m, ctype, b64) { return "src='cid:" + take(ctype, b64) + "'"; }
  );
  return { html: out, images: images };
}

// Read the signature saved by the settings pane (chunked in roamingSettings).
function readSignature(rs) {
  try {
    var n = rs.get("raa.sig.meta");
    if (!n) { return ""; }
    var out = "";
    for (var i = 0; i < n; i++) { out += rs.get("raa.sig." + i) || ""; }
    return out;
  } catch (e) { return ""; }
}

function notify(kind, text) {
  try {
    var item = Office.context.mailbox.item;
    if (!item || !item.notificationMessages) { return; }
    item.notificationMessages.replaceAsync("raAttach", {
      type:
        kind === "error"
          ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
          : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: String(text).substring(0, 150),
      icon: "Icon.16",
      persistent: false,
    });
  } catch (e) { /* ignore */ }
}

function finish(event) {
  if (event && typeof event.completed === "function") { event.completed(); }
}
