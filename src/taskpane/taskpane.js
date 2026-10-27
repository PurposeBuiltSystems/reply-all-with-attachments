/*
 * Signature settings pane.
 *
 * Stores the user's pasted signature (HTML) in roamingSettings so the
 * Reply All with Attachments command can add it to each Graph-built reply.
 *
 * STORAGE LIMIT: Outlook roamingSettings caps the WHOLE settings object at
 * 32 KB (not per key — chunking doesn't raise the ceiling). A full-res logo
 * blows past that, so every embedded image is downscaled + JPEG-compressed to
 * fit a byte budget before it's stored. The HTML is still chunked for safety.
 *
 * IMAGES: a pasted signature brings its logo in as a temporary blob:/cid:/remote
 * reference, NOT the actual bytes — so the logo would silently drop on save. We
 * embed images as compressed base64 data URIs instead, which persist. Two paths:
 *   1. After paste, convert any loaded <img> to a compressed data URI (canvas),
 *      falling back to the image bytes the clipboard carried alongside the HTML.
 *   2. An "Add image…" file picker that embeds the chosen file directly — the
 *      guaranteed path when the logo is cross-origin (canvas would be tainted).
 */

/* global Office, SignatureLayout */

var META = "raa.sig.meta";        // number of chunks
var CHUNK = "raa.sig.";           // raa.sig.0, raa.sig.1, ...
var CHUNK_SIZE = 30000;

// Per-image base64 budget (chars). Leaves room under the 32 KB roaming cap for
// the signature text, the chunk keys, and anything else Office stores.
var IMG_BUDGET = 22000;

// Only genuinely oversized logos get reined in, and only when the signature
// itself didn't specify a size. Forcing a width and display:block on every
// image moved logos onto their own line, splitting the user's text.
var MAX_LOGO_W = 320;

var sigEl, statusEl;

/**
 * Outlook desktop caches this pane's HTML while ?v= still fetches today's
 * JavaScript, so startup can run new code against an old page. An unguarded
 * dereference here throws inside Office.onReady and Outlook reports it as
 * "Add-in Error" - the whole pane, not one control. A sibling add-in failed
 * certification (1120.3.7.8) on exactly this.
 */
function el(id) { return document.getElementById(id); }
function onClick(id, fn) { var e = el(id); if (e) { e.onclick = fn; } }


/**
 * Sign out (certification 1100.5.7.1). This add-in authenticates from the
 * ribbon command, not from this pane, but the two share an origin - so the
 * MSAL cache cleared here is the same one the command uses.
 *
 * Under nested app authentication Outlook owns the session and no add-in can
 * end it. This clears the tokens and account this add-in cached, and says so;
 * claiming to sign the user out of Outlook would be false.
 */
var AUTH_CLIENT_ID = "87764ff9-16e7-4e2f-8164-38eff9f3a895";
var signOutPca = null;

/**
 * Add-in sign-out state.
 *
 * Certification rejected the naive version on a sibling add-in: "after
 * clicking sign-out there is no response or not signed out." Under nested app
 * authentication Outlook owns the session and getAllAccounts() reports the
 * HOST's account, not a cache this add-in controls - so clearing MSAL's cache
 * changes nothing visible and the next silent token still works.
 *
 * What this add-in can honestly deliver is refusing to act until the user
 * authenticates again. The ribbon command reads this same flag before using a
 * silent token, so signing out here genuinely gates Reply All.
 */
var SIGNED_OUT_KEY = "raaSignedOut";

function readSignedOut() {
  try { return Office.context.roamingSettings.get(SIGNED_OUT_KEY) === true; } catch (e) { return false; }
}
function writeSignedOut(v) {
  try {
    Office.context.roamingSettings.set(SIGNED_OUT_KEY, !!v);
    Office.context.roamingSettings.saveAsync(function () {});
  } catch (e) { /* in-memory only is still correct for this session */ }
}

function authPca() {
  if (!signOutPca && typeof msal !== "undefined") {
    signOutPca = msal.createNestablePublicClientApplication({
      auth: { clientId: AUTH_CLIENT_ID, authority: "https://login.microsoftonline.com/common" },
    });
  }
  return signOutPca;
}

async function renderAuthState() {
  var who = null;
  if (!readSignedOut()) {
    try {
      var p = authPca();
      if (p) {
        var pca = await p;
        var a = (pca.getAllAccounts && pca.getAllAccounts()) || [];
        who = a.length ? (a[0].username || a[0].name || "signed in") : null;
      }
    } catch (e) { who = null; }
  }
  var w = el("authWho"); if (w) { w.textContent = who ? ("Signed in as " + who) : "Not signed in"; }
  var b = el("signOut"); if (b) { b.hidden = !who; }
}

async function doSignOut() {
  // Flip the state and the UI first: awaiting a broker handshake before
  // showing anything is what certification saw as the button doing nothing.
  writeSignedOut(true);
  var b = el("signOut"); if (b) { b.disabled = true; b.hidden = true; }
  var w = el("authWho");
  if (w) {
    w.textContent = "Signed out. Reply All with Attachments will ask you to sign in again " +
      "before its next use. Your Outlook session is separate and is not affected.";
  }
  try {
    var p = signOutPca;
    signOutPca = null;
    if (p) {
      var pca = await p;
      var accts = (pca.getAllAccounts && pca.getAllAccounts()) || [];
      for (var i = 0; i < accts.length; i++) {
        try {
          if (pca.clearCache) { await pca.clearCache({ account: accts[i] }); }
          else if (pca.logoutPopup) { await pca.logoutPopup({ account: accts[i] }); }
        } catch (e) { /* best effort; the enforced state stands */ }
      }
    }
  } finally {
    if (b) { b.disabled = false; }
  }
}

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

Office.onReady(async function () {
  sigEl = el("sig");
  statusEl = el("status");
  try { await ensureSignatureLayout(); }
  catch (e) { status(e.message); return; }
  if (sigEl) { sigEl.innerHTML = readSignature(Office.context.roamingSettings) || ""; }

  onClick("signOut", doSignOut);
  renderAuthState();

  onClick("save", save);
  onClick("clear", function () {
    if (!sigEl) { return; }
    sigEl.innerHTML = "";
    sigEl.focus();
  });

  // Embed a logo file directly (guaranteed to persist).
  onClick("addimg", function () {
    var f = el("imgfile");
    if (f) { f.click(); }
  });
  var imgfile = el("imgfile");
  if (imgfile) {
    imgfile.onchange = function (e) {
      var f = e.target.files && e.target.files[0];
      if (f) {
        readFileAsDataUrl(f, function (url) {
          embedFromDataUrl(url, function (small) {
            appendImageAtEnd(small);
            status("Logo added. Click Save to keep it.");
          });
        });
      }
      e.target.value = "";
    };
  }

  // After a paste, turn temporary image references into compressed embedded bytes.
  if (sigEl) {
    sigEl.addEventListener("paste", function (e) {
      var clipFiles = imageFilesFromClipboard(e.clipboardData);
      var html = e.clipboardData && e.clipboardData.getData && e.clipboardData.getData("text/html");
      if (html) {
        // Only intercept a real HTML paste. A screenshot arrives as a file with
        // no text/html, and the default paste plus embedImages already handles
        // that correctly - don't break it.
        e.preventDefault();
        insertAtCursor(cleanPastedHtml(html));
      }
      setTimeout(function () { embedImages(clipFiles); }, 0);
    });
  }
});

function save() {
  if (!sigEl) { return; }   // stale cached page: nothing to save from
  normalizeImages();        // preserve alignment and display dimensions before compression
  compressExistingImages(); // shrink only the stored image bytes
  var html = sigEl.innerHTML.trim();
  var rs = Office.context.roamingSettings;
  var unembedded = countUnembeddedImages();
  writeSignature(rs, html);
  status("Saving…");
  rs.saveAsync(function (res) {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      var m = (res.error && res.error.message) || "unknown error";
      if (/size/i.test(m)) {
        status("Too big to save (32 KB limit). Try a smaller/simpler logo, or " +
               "remove the image and Save the text only.");
      } else {
        status("Couldn't save: " + m);
      }
      return;
    }
    var msg = "Saved. It'll appear on your next Reply All with Attachments.";
    if (unembedded > 0) {
      msg += " ⚠ " + unembedded + " image(s) aren't embedded and won't appear — " +
             "use “Add image…” to insert your logo file directly, then Save again.";
    }
    status(msg);
  });
}

function status(t) { statusEl.textContent = t; }

/* ---- image embedding + compression ---- */

// Pull any image File objects the clipboard carried alongside the pasted HTML.
function imageFilesFromClipboard(dt) {
  var out = [];
  var items = (dt && dt.items) || [];
  for (var i = 0; i < items.length; i++) {
    if (items[i].kind === "file" && items[i].type.indexOf("image") === 0) {
      var f = items[i].getAsFile();
      if (f) { out.push(f); }
    }
  }
  return out;
}

// Convert every <img> to a compressed base64 data URI that fits the budget.
// Live images: compress straight from the element (works same-origin / blob:).
// If the element tainted the canvas (cross-origin), use the clipboard bytes.
function embedImages(clipFiles) {
  readFilesAsDataUrls(clipFiles, function (clipUrls) {
    var clipIdx = 0;
    var imgs = Array.prototype.slice.call(sigEl.querySelectorAll("img"));
    imgs.forEach(function (img) {
      if (img.src.indexOf("data:") === 0) {
        // Already embedded, but maybe full-res — recompress to fit the budget.
        embedFromDataUrl(img.src, function (small) { img.src = small; });
        return;
      }
      var direct = null;
      try {
        direct = compressLoadedImg(img, img.src.indexOf("data:") === 0 ? img.src : null);
      } catch (e) { direct = null; } // tainted
      if (direct) { img.src = direct; return; }
      if (clipIdx < clipUrls.length) {
        embedFromDataUrl(clipUrls[clipIdx++], function (small) { img.src = small; });
      }
    });
  });
}

// Load a data URL into a fresh (untainted) image, then compress it.
function embedFromDataUrl(dataUrl, cb) {
  var im = new Image();
  im.onload = function () {
    var out;
    try { out = compressLoadedImg(im, dataUrl); } catch (e) { out = dataUrl; }
    cb(out || dataUrl);
  };
  im.onerror = function () { cb(dataUrl); };
  im.src = dataUrl;
}

/**
 * Fit an image into IMG_BUDGET while degrading it as little as possible.
 *
 * Order matters: an image that already fits is kept BYTE-FOR-BYTE — no
 * canvas, no re-encode, no resize — so a modest PNG keeps its transparency
 * and its original dimensions. Only when it genuinely doesn't fit do we
 * shrink, still trying PNG (transparency intact) at each size before
 * falling back to JPEG on white, which is the lossy last resort.
 *
 * Throws if the source image tainted the canvas (caller falls back).
 */
function compressLoadedImg(im, originalDataUrl) {
  // 1. already small enough? keep exactly what the user gave us
  if (originalDataUrl && originalDataUrl.length <= IMG_BUDGET) { return originalDataUrl; }

  var w0 = im.naturalWidth || im.width;
  var h0 = im.naturalHeight || im.height;
  if (!w0 || !h0) { return null; }

  // 2. try progressively smaller, PNG first so transparency survives
  var longest = Math.max(w0, h0);
  var dims = [longest, 640, 480, 360, 320, 240, 200, 160, 128, 96, 72]
    .filter(function (d, i) { return i === 0 || d < longest; });
  var qualities = [0.92, 0.85, 0.75, 0.6, 0.45];
  var smallest = null;

  for (var d = 0; d < dims.length; d++) {
    var scale = Math.min(1, dims[d] / longest);
    var w = Math.max(1, Math.round(w0 * scale));
    var h = Math.max(1, Math.round(h0 * scale));

    var pc = document.createElement("canvas");
    pc.width = w; pc.height = h;
    pc.getContext("2d").drawImage(im, 0, 0, w, h); // no white fill: keep alpha
    var png = pc.toDataURL("image/png");           // throws if tainted
    if (png.length <= IMG_BUDGET) { return png; }
    smallest = png;

    // 3. PNG too big at this size — JPEG on white, the lossy fallback
    var c = document.createElement("canvas");
    c.width = w; c.height = h;
    var ctx = c.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, w, h);
    ctx.drawImage(im, 0, 0, w, h);
    for (var q = 0; q < qualities.length; q++) {
      var url = c.toDataURL("image/jpeg", qualities[q]);
      if (url.length <= IMG_BUDGET) { return url; }
      smallest = url;
    }
  }
  return smallest; // best effort if nothing hit the budget
}

// Recompress any embedded (data:) image already sitting in the box that's over
// budget — covers logos placed before this code loaded. Synchronous: data:
// images are same-origin and already decoded, so the canvas won't taint.
function compressExistingImages() {
  var imgs = sigEl.querySelectorAll("img");
  for (var i = 0; i < imgs.length; i++) {
    var img = imgs[i];
    if (img.src.indexOf("data:") !== 0) { continue; }
    if (img.src.length <= IMG_BUDGET) { continue; }
    if (!img.complete || !img.naturalWidth) { continue; }
    try {
      var u = compressLoadedImg(img, null); // over budget by definition here
      if (u) { img.src = u; }
    } catch (e) { /* leave as-is; save will report if still too big */ }
  }
}

/** Keep logo dimensions stable and place images beneath the signature text. */
function normalizeImages() {
  SignatureLayout.stabilize(sigEl, MAX_LOGO_W);
  SignatureLayout.imagesBelowText(sigEl);
}

function countUnembeddedImages() {
  var imgs = sigEl.querySelectorAll("img");
  var n = 0;
  for (var i = 0; i < imgs.length; i++) {
    if (imgs[i].src.indexOf("data:") !== 0) { n++; }
  }
  return n;
}

function readFileAsDataUrl(file, cb) {
  var r = new FileReader();
  r.onload = function () { cb(r.result); };
  r.readAsDataURL(file);
}

function readFilesAsDataUrls(files, cb) {
  if (!files.length) { cb([]); return; }
  var urls = [], left = files.length;
  files.forEach(function (f, i) {
    readFileAsDataUrl(f, function (url) {
      urls[i] = url;
      if (--left === 0) { cb(urls); }
    });
  });
}

/**
 * Clean what Outlook and Word put on the clipboard.
 *
 * Copying a signature out of Outlook yields <p class="MsoNormal"> whose
 * margins are zeroed by a stylesheet that does NOT come with the fragment.
 * A contenteditable never applies that stylesheet, so every paragraph falls
 * back to the browser default of 1em top and bottom - which is the reported
 * bug: pasting a signature adds a blank line between every line, in the box
 * and in the mail that goes out.
 *
 * So the margins are made explicit, and Word's private markup is dropped.
 * Text formatting is retained while logos move beneath the text. Real
 * formatting the user chose - bold, colour, links, sizes - is left alone,
 * because a paste that quietly restyles someone's signature is a worse bug
 * than the spacing.
 */
function cleanPastedHtml(html) {
  var clipboardHtml = String(html);
  html = clipboardHtml
    .replace(/<!--[\s\S]*?-->/g, "")                                  // Word conditional comments
    .replace(/<\/?o:p[^>]*>/gi, "")                                    // Word's empty-paragraph markers
    .replace(/<(style|script)[^>]*>[\s\S]*?<\/\1>/gi, "")             // stylesheets that won't apply anyway
    .replace(/<\/?(html|head|body|meta|link)[^>]*>/gi, "");

  var box = document.createElement("div");
  box.innerHTML = html;
  SignatureLayout.clipboardAlignment(box, clipboardHtml);

  var all = box.querySelectorAll("*");
  for (var i = 0; i < all.length; i++) {
    var el = all[i];
    var cls = el.getAttribute("class");
    if (cls && /\bMso/i.test(cls)) { el.removeAttribute("class"); }
    var st = el.getAttribute("style");
    if (st) {
      var kept = st.split(";").filter(function (d) { return d.trim() && !/^\s*mso-/i.test(d); }).join(";");
      if (kept.trim()) { el.setAttribute("style", kept); } else { el.removeAttribute("style"); }
    }
  }
  // The stylesheet that zeroed these is gone, so say it on the element.
  var blocks = box.querySelectorAll("p");
  for (var j = 0; j < blocks.length; j++) {
    var b = blocks[j];
    var bs = b.getAttribute("style") || "";
    if (!/(^|;)\s*margin/i.test(bs)) {
      b.setAttribute("style", (bs && bs.trim() ? bs.replace(/;?\s*$/, ";") : "") + "margin:0");
    }
  }
  SignatureLayout.stabilize(box, MAX_LOGO_W);
  SignatureLayout.imagesBelowText(box);
  return box.innerHTML;
}

/** Insert at the caret, leaving it where the user put it (unlike insertHtml). */
function insertAtCursor(html) {
  if (!sigEl) { return; }
  sigEl.focus();
  document.execCommand("insertHTML", false, html);
}

function appendImageAtEnd(src) {
  if (!sigEl) { return; }
  var img = document.createElement("img");
  img.src = src;
  sigEl.appendChild(img);
  normalizeImages();
  sigEl.focus();
  var sel = window.getSelection && window.getSelection();
  if (sel) {
    var end = document.createRange();
    end.selectNodeContents(sigEl);
    end.collapse(false);
    sel.removeAllRanges();
    sel.addRange(end);
  }
}

function insertHtml(html) {
  if (!sigEl) { return; }
  sigEl.focus();
  var sel = window.getSelection && window.getSelection();
  if (sel) {
    var end = document.createRange();
    end.selectNodeContents(sigEl);
    end.collapse(false);          // false = collapse to the END of the contents
    sel.removeAllRanges();
    sel.addRange(end);
  }
  document.execCommand("insertHTML", false, html);
}

/* ---- chunked storage helpers (shared shape with commands.js) ---- */

function readSignature(rs) {
  var n = rs.get(META);
  if (!n) { return ""; }
  var out = "";
  for (var i = 0; i < n; i++) { out += rs.get(CHUNK + i) || ""; }
  return out;
}

function writeSignature(rs, html) {
  // clear any previous chunks
  var prev = rs.get(META) || 0;
  for (var j = 0; j < prev; j++) { rs.remove(CHUNK + j); }
  if (!html) { rs.remove(META); return; }
  var count = 0;
  for (var pos = 0; pos < html.length; pos += CHUNK_SIZE) {
    rs.set(CHUNK + count, html.substr(pos, CHUNK_SIZE));
    count++;
  }
  rs.set(META, count);
}
