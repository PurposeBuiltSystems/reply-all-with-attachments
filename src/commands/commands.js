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

/** Get a Graph token: silent first, interactive only if needed. */
async function getToken() {
  var pca = await getPca();
  try {
    var silent = await pca.acquireTokenSilent({ scopes: SCOPES });
    return silent.accessToken;
  } catch (e) {
    var interactive = await pca.acquireTokenPopup({ scopes: SCOPES });
    return interactive.accessToken;
  }
}

/** Thin Graph fetch helper. */
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
    var attachments = await graph(token, "GET", "/me/messages/" + restId + "/attachments");
    var files = (attachments.value || []).filter(function (a) {
      return a["@odata.type"] === "#microsoft.graph.fileAttachment" && !a.isInline;
    });

    // 3. Copy each onto the draft. (Large attachments may need their own GET for
    //    contentBytes; handled in a later pass if needed.)
    var skipped = 0;
    for (var i = 0; i < files.length; i++) {
      var f = files[i];
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
    }

    // 4. Open the populated draft for the user to review/send.
    var ewsId = Office.context.mailbox.convertToEwsId(
      draft.id,
      Office.MailboxEnums.RestVersion.v2_0
    );
    Office.context.mailbox.displayMessageForm(ewsId);

    if (skipped > 0) {
      notify("info", skipped + " attachment(s) could not be copied and were skipped.");
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
