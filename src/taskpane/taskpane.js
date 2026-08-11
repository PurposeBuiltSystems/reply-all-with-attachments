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

/* global Office */

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
  try {
    var p = authPca();
    if (p) {
      var pca = await p;
      var a = (pca.getAllAccounts && pca.getAllAccounts()) || [];
      who = a.length ? (a[0].username || a[0].name || "signed in") : null;
    }
  } catch (e) { who = null; }
  var w = el("authWho"); if (w) { w.textContent = who ? ("Signed in as " + who) : "Not signed in"; }
  var b = el("signOut"); if (b) { b.hidden = !who; }
}

async function doSignOut() {
  var b = el("signOut"); if (b) { b.disabled = true; }
  try {
    var p = authPca();
    if (p) {
      var pca = await p;
      var accts = (pca.getAllAccounts && pca.getAllAccounts()) || [];
      for (var i = 0; i < accts.length; i++) {
        try {
          if (pca.clearCache) { await pca.clearCache({ account: accts[i] }); }
          else if (pca.logoutPopup) { await pca.logoutPopup({ account: accts[i] }); }
        } catch (e) { /* keep clearing the rest */ }
      }
    }
    signOutPca = null;
    var w = el("authWho");
    if (w) {
      w.textContent = "Signed out \u2014 this add-in's saved tokens are cleared. Your Outlook " +
        "session is separate and is not affected; no add-in can end it.";
    }
  } finally {
    if (b) { b.disabled = false; }
    setTimeout(renderAuthState, 2500);
  }
}

Office.onReady(function () {
  sigEl = el("sig");
  statusEl = el("status");
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
            insertHtml('<img src="' + small + '" />');
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
      setTimeout(function () { embedImages(clipFiles); }, 0);
    });
  }
});

function save() {
  if (!sigEl) { return; }   // stale cached page: nothing to save from
  compressExistingImages(); // shrink any oversized logo already in the box
  normalizeImages();        // cap display width + stack text below the logo
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

/**
 * Keep a logo from overflowing the message, WITHOUT relaying it out. The
 * signature is the user's; where the logo sits relative to their text is
 * their decision. So: never set display, and only impose a width when the
 * signature gave none and the image is genuinely oversized.
 */
function normalizeImages() {
  var imgs = sigEl.querySelectorAll("img");
  for (var i = 0; i < imgs.length; i++) {
    var img = imgs[i];
    img.style.maxWidth = "100%";               // never wider than the message
    var sized = img.getAttribute("width") || (img.style && img.style.width);
    if (sized) { continue; }                   // the signature already said how big
    var natural = img.naturalWidth || 0;
    if (natural > MAX_LOGO_W) {
      img.setAttribute("width", String(MAX_LOGO_W));
      img.style.height = "auto";
      img.removeAttribute("height");
    }
  }
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
 * Where the logo lands.
 *
 * Choosing a file opens the OS picker, which takes focus off the editable box.
 * focus() restores focus but NOT the caret, so execCommand("insertHTML") used
 * to insert wherever the browser happened to leave it - usually part-way up
 * the signature. That is the reported bug: the image appearing above the last
 * line of text.
 *
 * Restoring the remembered caret was the obvious fix and is still wrong: the
 * caret sits wherever you last typed, which is frequently not the last line,
 * so the logo still lands mid-signature and it still looks arbitrary.
 *
 * This button therefore always appends at the end. It is the one position that
 * is predictable before you click, it is what a signature logo almost always
 * wants, and it is stated next to the button so nothing is a surprise. Placing
 * an image precisely is still possible - pasting drops it at the cursor, which
 * the browser handles natively and this code does not touch.
 */
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
