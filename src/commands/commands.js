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

/* global Office, msal, SignatureLayout */

var CLIENT_ID = "87764ff9-16e7-4e2f-8164-38eff9f3a895"; // Entra app registration (PurposeBuiltSystems)
var GRAPH = "https://graph.microsoft.com/v1.0";
var SCOPES = ["Mail.ReadWrite"];

var pcaPromise = null;
var commandAssociated = false;

function associateCommand() {
  if (!commandAssociated && Office.actions && Office.actions.associate) {
    Office.actions.associate("replyAllWithAttachments", replyAllWithAttachments);
    commandAssociated = true;
  }
}
// Some cold runtimes expose actions only after Office finishes initializing.
// Register immediately when possible and try again when the host is ready.
associateCommand();
Office.onReady(associateCommand);

// Older cached HTML may not include the new layout script yet. Load it on
// demand so a fresh command script still works on the very first click.
async function ensureSignatureLayout() {
  if (typeof SignatureLayout !== "undefined") { return; }
  await new Promise(function (resolve, reject) {
    var timer = setTimeout(function () { reject(new Error("Signature layout didn't load. Check your connection and try again.")); }, 20000);
    var script = document.createElement("script");
    script.src = new URL("../signature-layout.js?v=2", document.baseURI).href;
    script.onload = function () {
      clearTimeout(timer);
      if (typeof SignatureLayout !== "undefined") { resolve(); }
      else { reject(new Error("Signature layout couldn't load. Reopen the Signature pane and try again.")); }
    };
    script.onerror = function () { clearTimeout(timer); reject(new Error("Signature layout couldn't load. Check your connection and try again.")); };
    document.head.appendChild(script);
  });
}

/** Lazily create the MSAL nestable client (NAA). */
function getPca() {
  if (!pcaPromise) {
    pcaPromise = msal.createNestablePublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: "https://login.microsoftonline.com/common",
      },
    }).catch(function (err) {
      // A rejected broker initialization must not poison later attempts.
      pcaPromise = null;
      throw err;
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
  var pca;
  for (var attempt = 0; attempt < 2; attempt++) {
    try {
      pca = await withTimeout(getPca(), 20000,
        "Sign-in didn't start. Fully quit Outlook (Cmd+Q) and reopen, then try again.");
      break;
    } catch (initError) {
      // Retry a definite initialization failure once, within the first click.
      // A timeout can still be running, so do not start a second broker then.
      if (attempt > 0 || pcaPromise) { throw initError; }
      await delay(250);
    }
  }
  try {
    // The pane's Sign out sets this. Honour it here, or that button is
    // decorative: the command would keep working on a silent token.
    var out = false;
    try { out = Office.context.roamingSettings.get("raaSignedOut") === true; } catch (e2) { out = false; }
    if (out) { throw new Error("signed out of the add-in"); }
    return (await withTimeout(pca.acquireTokenSilent({ scopes: SCOPES }), 20000, "silent timeout")).accessToken;
  } catch (e) {
    var interactive = await withTimeout(
      pca.acquireTokenPopup({ scopes: SCOPES }), 120000,
      "Sign-in didn't finish \u2014 a Microsoft sign-in window may have opened behind Outlook. " +
      "Check for it, finish signing in, and click again. If none appeared, fully quit Outlook " +
      "(Cmd+Q), reopen, and retry.");
    try {
      Office.context.roamingSettings.set("raaSignedOut", false);
      Office.context.roamingSettings.saveAsync(function () {});
    } catch (e3) { /* non-fatal */ }
    return interactive.accessToken;
  }
}

/** Thin Graph fetch helper. */
function delay(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }


// Bound both the request and its response body. A stalled connection must
// release Outlook's command runtime and show a recoverable error.
async function fetchBounded(url, opts) {
  var controller = typeof AbortController !== "undefined" ? new AbortController() : null;
  var options = Object.assign({}, opts);
  if (controller) { options.signal = controller.signal; }
  try {
    return await withTimeout(fetch(url, options), 30000, "Microsoft request timed out. Check your connection.");
  } catch (err) {
    if (controller) { controller.abort(); }
    throw err;
  }
}

// Retry throttles once, respecting Retry-After. A 503 on a write may have
// committed already, so only replay reads for that ambiguous failure.
async function fetchRetry(url, opts) {
  var res = await fetchBounded(url, opts);
  if (res.status === 429 || (res.status === 503 && (!opts || !opts.method || opts.method === "GET"))) {
    var retryAfter = res.headers.get("Retry-After");
    var wait = retryAfter && /^\d+(\.\d+)?$/.test(retryAfter.trim())
      ? Number(retryAfter) * 1000 : Date.parse(retryAfter) - Date.now();
    if (!isFinite(wait)) { wait = 3000; }
    wait = Math.max(0, wait);
    // Surface long throttles rather than retrying before the server allows it.
    if (wait > 15000) { return res; }
    await new Promise(function (r) { setTimeout(r, wait); });
    res = await fetchBounded(url, opts);
  }
  return res;
}

async function graph(token, method, path, body) {
  var url = path.indexOf(GRAPH + "/") === 0 ? path : GRAPH + path;
  var res = await fetchRetry(url, {
    method: method,
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    var text = await withTimeout(res.text(), 30000, "Microsoft error response timed out.");
    var error = new Error("Graph " + method + " " + path + " → " + res.status + " " + text);
    error.status = res.status;
    throw error;
  }
  return res.status === 204 ? null : withTimeout(res.json(), 30000, "Microsoft response timed out. Check your connection.");
}

// A new draft can briefly return 404 immediately after createReplyAll. Retry
// its reads only; never replay draft creation or attachment writes.
async function readNewDraft(token, path) {
  for (var attempt = 0; attempt < 3; attempt++) {
    try { return await graph(token, "GET", path); }
    catch (err) {
      if (err.status !== 404 || attempt === 2) { throw err; }
      await delay(250 * (attempt + 1));
    }
  }
}

/** Read every page, including attachment collections with server-side paging. */
async function graphAll(token, path, newDraft) {
  var out = [];
  var next = path;
  var pages = 0;
  while (next && pages++ < 100) {
    var page = newDraft ? await readNewDraft(token, next) : await graph(token, "GET", next);
    out = out.concat(page.value || []);
    next = page["@odata.nextLink"] || null;
    if (next && next.indexOf(GRAPH + "/") !== 0) {
      throw new Error("Unexpected Graph paging URL.");
    }
  }
  if (next) { throw new Error("Attachment results exceeded the scan limit."); }
  return out;
}

/** Graph requires an upload session for files at or above 3 MB. */
async function copyFileAttachment(token, draftId, file, base64) {
  var path = "/me/messages/" + draftId + "/attachments";
  // Small uploads already need base64. Avoid decoding a second full copy just
  // to measure it, especially when copying several files at once.
  var padding = /==$/.test(base64) ? 2 : /=$/.test(base64) ? 1 : 0;
  var byteLength = base64.length * 3 / 4 - padding;
  if (byteLength < 3000000) {
    return graph(token, "POST", path, {
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: file.name,
      contentType: file.contentType,
      contentBytes: base64,
    });
  }
  var binary = atob(base64);
  var session = await graph(token, "POST", path + "/createUploadSession", {
    AttachmentItem: { attachmentType: "file", name: file.name, size: binary.length, contentType: file.contentType },
  });
  // Sequential ranges under 4 MB; the final response must confirm creation.
  var CHUNK = 3276800;
  for (var start = 0; start < binary.length; start += CHUNK) {
    var end = Math.min(start + CHUNK, binary.length);
    var bytes = new Uint8Array(end - start);
    for (var i = start; i < end; i++) { bytes[i - start] = binary.charCodeAt(i); }
    var res = await fetchRetry(session.uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Type": "application/octet-stream",
        "Content-Range": "bytes " + start + "-" + (end - 1) + "/" + binary.length,
      },
      body: bytes,
    });
    // The session URL is pre-authenticated: do not send a Graph bearer token.
    if (!res.ok || (end === binary.length && res.status !== 201)) {
      throw new Error("Attachment upload " + file.name + " -> " + res.status);
    }
  }
}

/** @param {Office.AddinCommands.Event} event */
async function replyAllWithAttachments(event) {
  var creationStarted = false;
  try {
    var item = Office.context.mailbox.item;
    // Office gives an EWS id; Graph needs a REST id.
    var restId = Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );

    notify("info", "Signing in to Outlook…");
    var sig = readSignature(Office.context.roamingSettings);
    // Load/format the saved signature while Outlook's authentication broker
    // is working, rather than adding that work after sign-in.
    var layoutP = sig ? ensureSignatureLayout().then(function () {
      return extractInlineImages(SignatureLayout.prepare(sig));
    }) : Promise.resolve({ html: "", images: [] });
    layoutP.catch(function () {});
    var token = await getToken();

    // Each Graph round trip costs a few hundred ms, and they used to run one
    // after another. Anything that doesn't depend on the draft starts now.
    // An explicitly empty Office collection needs no Graph attachment scan.
    // Missing metadata keeps the original Graph fallback.
    var noAttachments = Array.isArray(item.attachments) && item.attachments.length === 0;
    var originalsP = noAttachments ? Promise.resolve([]) : graphAll(token, "/me/messages/" + restId + "/attachments");
    originalsP.catch(function () {}); // awaited below; don't report it twice

    // 1. Create the Reply All draft (Graph omits original attachments, like Outlook).
    notify("info", "Creating your reply draft…");
    creationStarted = true;
    var draft = await graph(token, "POST", "/me/messages/" + restId + "/createReplyAll", {});

    if (!draft || !draft.id) { throw new Error("Microsoft did not return the draft ID."); }
    notify("info", "Reading the original attachments…");

    // Its body and the inline images createReplyAll already carried over.
    var reads = await Promise.all([
      // createReplyAll normally returns the body already. Fetch only when a
      // client/service response omits it, preserving the compatibility path.
      draft.body && typeof draft.body.content === "string" ? Promise.resolve(draft) :
        readNewDraft(token, "/me/messages/" + draft.id + "?$select=body"),
      noAttachments ? Promise.resolve([]) : graphAll(token, "/me/messages/" + draft.id + "/attachments?$select=name,size", true),
    ]);
    var body = (reads[0] && reads[0].body && reads[0].body.content) || "";
    var existing = reads[1];

    // 1b. Tidy the quoted-thread header (Graph mashes the divider + From/Sent/
    //     To/Subject onto one line for plain-text originals) and, if the user
    //     saved a signature, insert it above the quoted thread — embedding its
    //     logo as an inline cid attachment so it renders reliably (some clients
    //     strip raw data: URIs in mail).
    var preparedSignature = await layoutP;
    var tidied = tidyQuotedHeader(body);
    var content = tidied;
    var images = [];
    if (sig) {
      var ex = preparedSignature;
      images = ex.images;
      content = prependSignature(tidied, ex.html);
    }

    // 2. The original message's file attachments (skip inline images).
    //    isInline alone is unreliable — many clients embed signature logos
    //    without setting it — so also treat any attachment whose contentId is
    //    referenced as cid: in the quoted body as inline.
    var attachments = await originalsP;
    var files = attachments.filter(function (a) {
      if (a["@odata.type"] !== "#microsoft.graph.fileAttachment") { return false; }
      if (a.isInline) { return false; }
      if (a.contentId && content.indexOf("cid:" + a.contentId) !== -1) { return false; }
      return true;
    });

    // 2b. createReplyAll already carries the inline images the quoted thread
    //     needs, and some senders attach the same file twice — skip anything
    //     already on the draft (or already copied) so nothing lands twice.
    var seen = {};
    existing.forEach(function (a) { seen[a.name + "|" + a.size] = true; });
    var toCopy = files.filter(function (f) {
      var key = f.name + "|" + f.size;
      if (seen[key]) { return false; }
      seen[key] = true;
      return true;
    });

    // 3. Write the signature and copy the files. They touch different parts of
    //    the draft, so they run side by side. The logo must exist before the
    //    body that references it, so that one pair stays in order.
    notify("info", "Copying " + toCopy.length + " attachment(s) and saving your signature…");
    var writeSignature = (async function () {
      if (!sig && tidied === body) { return; }
      await Promise.all(images.map(function (im) {
        return graph(token, "POST", "/me/messages/" + draft.id + "/attachments", {
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: im.name,
          contentType: im.contentType,
          contentBytes: im.contentBytes,
          isInline: true,
          contentId: im.cid,
        });
      }));
      await graph(token, "PATCH", "/me/messages/" + draft.id, {
        body: { contentType: "HTML", content: content },
      });
    })();
    var copyFiles = mapLimit(toCopy, 3, async function (f) {
      var bytes = f.contentBytes;
      if (!bytes) {
        // Collection didn't include the bytes — fetch the single attachment.
        var full = await graph(token, "GET", "/me/messages/" + restId + "/attachments/" + f.id);
        bytes = full && full.contentBytes;
      }
      if (!bytes) { return null; }
      await copyFileAttachment(token, draft.id, f, bytes);
      return f.name;
    });
    var results = await Promise.all([writeSignature, copyFiles]);
    var copiedNames = results[1].filter(function (n) { return n !== null; });
    var skipped = toCopy.length - copiedNames.length;

    // 4. Outlook opens its locally synced copy of the draft, which can lag the
    //    writes above. Wait briefly for the server to show them; if it still
    //    doesn't, open anyway and say what to check. Refusing to open left
    //    people waiting on an error for a draft that was already in Drafts.
    notify("info", "Checking that your reply has synced…");
    var ready;
    var verificationFailed = false;
    try { ready = await waitForDraft(token, draft.id, images, copiedNames); }
    catch (verifyError) {
      // All writes succeeded. A failed verification read should not prevent
      // opening the saved draft and make the user create another on click two.
      verificationFailed = true;
    }

    notify("info", "Opening your reply draft…");
    var notes = [];
    if (verificationFailed) { notes.push("Draft saved; verification was interrupted. Check attachments and signature before sending."); }
    else if (!ready.files) { notes.push("Attachments are still syncing: if any are missing, close the draft and reopen it from Drafts before sending."); }
    else if (!ready.signature) { notes.push("Your signature logo is still syncing; check it before sending."); }
    if (skipped > 0) { notes.push(skipped + " attachment(s) couldn't be copied."); }
    try {
      await openDraft(draft.id);
      notify("info", notes.length ? notes.join(" ") : "Reply draft ready. Review it before sending.");
    } catch (openErr) {
      notify("info", "Reply draft saved, but Outlook couldn't open it. Open Drafts to review it; you don't need to click Reply All again.");
    }
    finish(event);
  } catch (e) {
    notify("error", (creationStarted ? "Reply interrupted. Check Drafts before trying again. " : "Reply could not start. ") + (e && e.message ? e.message : e));
    finish(event);
  }
}

/** Run fn over items, at most `limit` at a time, keeping result order. */
async function mapLimit(items, limit, fn) {
  var out = new Array(items.length);
  var next = 0;
  async function worker() {
    while (next < items.length) {
      var i = next++;
      out[i] = await fn(items[i]);
    }
  }
  var workers = [];
  for (var w = 0; w < Math.min(limit, items.length); w++) { workers.push(worker()); }
  await Promise.all(workers);
  return out;
}

// Insert into the existing document's body, not before its <html>/<head>.
// An empty line goes ABOVE the signature, as Outlook's own replies do: the
// draft opens with the cursor at the top, and without it the first thing
// typed lands inside the signature and takes on its formatting.
function prependSignature(html, signature) {
  var prefix = "<div><br></div>" + signature + "<br><br>";
  if (/<body\b[^>]*>/i.test(html)) {
    return html.replace(/<body\b[^>]*>/i, function (bodyTag) { return bodyTag + prefix; });
  }
  return prefix + html;
}

function referencesImage(html, cid) {
  var refs = String(html || "").match(/cid:[^"'\s<>]+/gi) || [];
  return refs.some(function (ref) {
    try { return decodeURIComponent(ref.slice(4)) === cid; }
    catch (e) { return ref.slice(4) === cid; }
  });
}

/**
 * Poll until the server shows the signature logo (referenced and attached)
 * and every copied file. Copies are matched by NAME: Exchange's attachment
 * size counts stored properties too, so a copy needn't report the original's
 * size, and matching on it can wait out every retry for a file already there.
 */
async function waitForDraft(token, draftId, images, copiedNames) {
  var state = { signature: !images.length, files: !copiedNames.length };
  var wanted = {};
  copiedNames.forEach(function (n) { wanted[n] = (wanted[n] || 0) + 1; });
  for (var attempt = 0; attempt < 5 && !(state.signature && state.files); attempt++) {
    if (attempt > 0) { await delay(300 * attempt); }
    var reads = await Promise.all([
      state.signature ? null : readNewDraft(token, "/me/messages/" + draftId + "?$select=body"),
      graphAll(token, "/me/messages/" + draftId + "/attachments?$select=name,isInline,contentId", true),
    ]);
    var html = (reads[0] && reads[0].body && reads[0].body.content) || "";
    var atts = reads[1];
    if (!state.signature) {
      state.signature = images.every(function (im) {
        return referencesImage(html, im.cid) && atts.some(function (a) {
          return a.isInline && a.contentId === im.cid;
        });
      });
    }
    if (!state.files) {
      var have = {};
      atts.forEach(function (a) { if (!a.isInline) { have[a.name] = (have[a.name] || 0) + 1; } });
      state.files = Object.keys(wanted).every(function (n) { return (have[n] || 0) >= wanted[n]; });
    }
  }
  return state;
}

// Keep the command runtime alive until Outlook confirms the open request.
// Transient 'item not found' failures retry the SAME draft, without copying again.
async function openDraft(draftId) {
  var mailbox = Office.context.mailbox;
  var ewsId = mailbox.convertToEwsId(draftId, Office.MailboxEnums.RestVersion.v2_0);
  var requirements = Office.context.requirements;
  var useAsync = typeof mailbox.displayMessageFormAsync === "function" &&
    (!requirements || requirements.isSetSupported("Mailbox", "1.9"));
  if (!useAsync) {
    mailbox.displayMessageForm(ewsId);
    // Legacy API has no acknowledgement; avoid tearing its runtime down immediately.
    await delay(500);
    return;
  }
  for (var attempt = 0; attempt < 3; attempt++) {
    // A missing callback has an unknown outcome, so a timeout exits without
    // retrying. Only a definite failed callback is retried below.
    var result = await withTimeout(new Promise(function (resolve) {
      mailbox.displayMessageFormAsync(ewsId, resolve);
    }), 10000, "Outlook didn't confirm opening the draft. Check Drafts before trying again.");
    if (result && result.status === Office.AsyncResultStatus.Succeeded) { return; }
    if (attempt === 2) { throw new Error((result && result.error && result.error.message) || "Outlook couldn't open the draft."); }
    await delay(600 * (attempt + 1));
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
  var instance = Date.now().toString(36) + Math.random().toString(36).slice(2);
  function take(ctype, b64) {
    n++;
    // A Content-ID must be an addr-spec (id@domain), per RFC 2045/2392. A bare
    // token like "raasig1" is not, and Exchange silently DROPS the
    // <img src="cid:raasig1"> reference while KEEPING the attachment - leaving
    // an inline image nothing points at, which Outlook then renders wherever
    // it likes. That is the reported bug: the logo appearing beside the last
    // line of the signature instead of where it was placed.
    var cid = "raasig" + instance + "-" + n + "@purposebuilt.systems";
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
