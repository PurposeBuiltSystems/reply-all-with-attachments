/*
 * Reply All with Attachments — event-based implementation.
 *
 * WHY event-based: Outlook's displayReplyAllForm "attachments" parameter is
 * broken on Outlook on the web and New Outlook (Microsoft bug office-js #4599) —
 * it silently drops the files. It only works on Classic Outlook. To work on the
 * modern clients, we instead attach the files INSIDE the reply using the compose
 * API addFileAttachmentFromBase64Async, which does work there.
 *
 * FLOW:
 *   1. Ribbon button (read mode): read the open message's non-inline attachments,
 *      fetch each one's bytes (base64), stash them in roamingSettings (chunked,
 *      since each setting is capped at 32 KB and the add-in total at ~2 MB), then
 *      open a normal Reply All form.
 *   2. OnNewMessageCompose handler (fires automatically when that reply opens):
 *      read the stash, re-attach each file via addFileAttachmentFromBase64Async,
 *      clear the stash, and call event.completed().
 *
 * The stash carries a timestamp; the handler only acts on a stash less than
 * 60 s old, so it never attaches to an unrelated new message/forward.
 *
 * Requirement set: Mailbox 1.10 (event-based activation). Permission:
 * read/write item (compose attach needs write).
 */

/* global Office */

var META_KEY = "raa.meta";          // JSON: { ts, files: [{ name, chunks }] }
var CHUNK_PREFIX = "raa.";          // raa.<fileIndex>.<chunkIndex> = base64 piece
var FRESH_MS = 60000;               // only honor a stash younger than 60s
var CHUNK_SIZE = 30000;             // chars per roamingSettings entry (<32 KB)
var MAX_TOTAL = 1800000;            // ~1.8M base64 chars (~2 MB roaming cap)

// Event-based (JS-only) runtimes do NOT run Office.onReady, so the handler
// associations must be registered at top level.
Office.onReady(function () {});
if (Office.actions && Office.actions.associate) {
  Office.actions.associate("replyAllWithAttachments", replyAllWithAttachments);
  Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
}

/* ------------------------------------------------------------------ */
/* Ribbon button (read mode)                                          */
/* ------------------------------------------------------------------ */

/** @param {Office.AddinCommands.Event} event */
async function replyAllWithAttachments(event) {
  try {
    var item = Office.context.mailbox.item;
    var atts = (item.attachments || []).filter(function (a) {
      return !a.isInline; // inline images stay in the quoted reply automatically
    });

    if (atts.length === 0) {
      item.displayReplyAllFormAsync("", function () { finish(event); });
      return;
    }

    // Fetch every attachment's bytes.
    var fetched = await Promise.all(
      atts.map(function (a) {
        return getContent(item, a.id).then(function (c) { return { att: a, content: c }; });
      })
    );

    // Build base64 file list, respecting the roamingSettings size cap.
    var files = [];
    var total = 0;
    var skipped = 0;
    fetched.forEach(function (f) {
      var built = buildFile(f.att, f.content);
      if (!built) { skipped++; return; }
      if (total + built.b64.length > MAX_TOTAL) { skipped++; return; }
      total += built.b64.length;
      files.push(built);
    });

    // Stash for the compose handler to pick up.
    await stash(files);

    // Open Reply All (no attachments param — the handler attaches them).
    item.displayReplyAllFormAsync("", function (res) {
      if (res.status === Office.AsyncResultStatus.Failed) {
        notify("error", "Couldn't open the reply: " + res.error.message);
      } else if (skipped > 0) {
        notify("info", skipped + " attachment(s) were too large to carry over and were skipped.");
      }
      finish(event);
    });
  } catch (e) {
    notify("error", "Reply All with Attachments failed: " + msg(e));
    finish(event);
  }
}

/* ------------------------------------------------------------------ */
/* OnNewMessageCompose handler (compose mode, fires on the reply)     */
/* ------------------------------------------------------------------ */

/** @param {Office.AddinCommands.Event} event */
function onMessageComposeHandler(event) {
  try {
    var rs = Office.context.roamingSettings;
    var raw = rs.get(META_KEY);
    if (!raw) { return event.completed(); } // not our reply — do nothing

    var meta = JSON.parse(raw);
    var stale = !meta || !meta.files || !meta.files.length || (nowMs() - meta.ts) > FRESH_MS;
    if (stale) {
      clearStash(rs);
      rs.saveAsync(function () { event.completed(); });
      return;
    }

    // Reconstruct each file's base64 from its chunks.
    var item = Office.context.mailbox.item;
    var files = meta.files.map(function (f, i) {
      var b64 = "";
      for (var j = 0; j < f.chunks; j++) {
        b64 += rs.get(CHUNK_PREFIX + i + "." + j) || "";
      }
      return { name: f.name, b64: b64 };
    });

    attachSequentially(item, files, 0, function () {
      clearStash(rs);
      rs.saveAsync(function () { event.completed(); });
    });
  } catch (e) {
    event.completed();
  }
}

function attachSequentially(item, files, idx, done) {
  if (idx >= files.length) { return done(); }
  item.addFileAttachmentFromBase64Async(
    files[idx].b64,
    files[idx].name,
    { isInline: false },
    function () { attachSequentially(item, files, idx + 1, done); } // continue regardless of per-file result
  );
}

/* ------------------------------------------------------------------ */
/* roamingSettings stash helpers                                      */
/* ------------------------------------------------------------------ */

function stash(files) {
  return new Promise(function (resolve, reject) {
    var rs = Office.context.roamingSettings;
    clearStash(rs); // drop any previous payload's chunk keys

    var meta = { ts: nowMs(), files: [] };
    files.forEach(function (f, i) {
      var chunks = 0;
      for (var pos = 0; pos < f.b64.length; pos += CHUNK_SIZE) {
        rs.set(CHUNK_PREFIX + i + "." + chunks, f.b64.substr(pos, CHUNK_SIZE));
        chunks++;
      }
      meta.files.push({ name: f.name, chunks: chunks });
    });
    rs.set(META_KEY, JSON.stringify(meta));

    rs.saveAsync(function (res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) { resolve(); }
      else { reject(res.error); }
    });
  });
}

/** Remove the previous payload's keys from the in-memory settings (caller saves). */
function clearStash(rs) {
  try {
    var raw = rs.get(META_KEY);
    if (raw) {
      var meta = JSON.parse(raw);
      (meta.files || []).forEach(function (f, i) {
        for (var j = 0; j < f.chunks; j++) { rs.remove(CHUNK_PREFIX + i + "." + j); }
      });
    }
  } catch (e) { /* ignore */ }
  rs.remove(META_KEY);
}

/* ------------------------------------------------------------------ */
/* Attachment content -> { name, b64 }                                */
/* ------------------------------------------------------------------ */

function buildFile(att, content) {
  var F = Office.MailboxEnums.AttachmentContentFormat;
  switch (content.format) {
    case F.Base64:
      return { name: att.name, b64: content.content };
    case F.Eml:
      return safeText(ensureExt(att.name, ".eml"), content.content);
    case F.ICalendar:
      return safeText(ensureExt(att.name, ".ics"), content.content);
    case F.Url:
    default:
      return null; // cloud/url attachments aren't carried in this version
  }
}

function safeText(name, text) {
  try { return { name: name, b64: utf8ToBase64(text) }; }
  catch (e) { return null; } // btoa may be unavailable in some runtimes
}

function ensureExt(name, ext) {
  if (!name) { return "attachment" + ext; }
  return name.toLowerCase().endsWith(ext) ? name : name + ext;
}

function utf8ToBase64(str) {
  return btoa(unescape(encodeURIComponent(str)));
}

/* ------------------------------------------------------------------ */
/* Small utilities                                                    */
/* ------------------------------------------------------------------ */

function getContent(item, attachmentId) {
  return new Promise(function (resolve, reject) {
    item.getAttachmentContentAsync(attachmentId, function (res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) { resolve(res.value); }
      else { reject(res.error); }
    });
  });
}

function nowMs() {
  return new Date().getTime();
}

function msg(e) {
  return e && e.message ? e.message : String(e);
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
  } catch (e) { /* notifications unavailable in some contexts */ }
}

function finish(event) {
  if (event && typeof event.completed === "function") { event.completed(); }
}
