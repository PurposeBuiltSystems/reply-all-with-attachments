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

    // 1b. If the user saved a signature, insert it above the quoted thread.
    var sig = readSignature(Office.context.roamingSettings);
    if (sig) {
      var d = await graph(token, "GET", "/me/messages/" + draft.id + "?$select=body");
      var body = (d && d.body && d.body.content) || "";
      await graph(token, "PATCH", "/me/messages/" + draft.id, {
        body: { contentType: "HTML", content: sig + "<br><br>" + body },
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
