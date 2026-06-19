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

var sigEl, statusEl;

Office.onReady(function () {
  sigEl = document.getElementById("sig");
  statusEl = document.getElementById("status");
  sigEl.innerHTML = readSignature(Office.context.roamingSettings) || "";

  document.getElementById("save").onclick = save;
  document.getElementById("clear").onclick = function () {
    sigEl.innerHTML = "";
    sigEl.focus();
  };

  // Embed a logo file directly (guaranteed to persist).
  document.getElementById("addimg").onclick = function () {
    document.getElementById("imgfile").click();
  };
  document.getElementById("imgfile").onchange = function (e) {
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

  // After a paste, turn temporary image references into compressed embedded bytes.
  sigEl.addEventListener("paste", function (e) {
    var clipFiles = imageFilesFromClipboard(e.clipboardData);
    setTimeout(function () { embedImages(clipFiles); }, 0);
  });
});

function save() {
  compressExistingImages(); // shrink any oversized logo already in the box
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
      try { direct = compressLoadedImg(img); } catch (e) { direct = null; } // tainted
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
    try { out = compressLoadedImg(im); } catch (e) { out = dataUrl; }
    cb(out || dataUrl);
  };
  im.onerror = function () { cb(dataUrl); };
  im.src = dataUrl;
}

// Downscale + JPEG-compress a loaded image until the data URL fits IMG_BUDGET.
// White background so transparent logos stay readable on white email bodies.
// Throws if the source image tainted the canvas (caller falls back).
function compressLoadedImg(im) {
  var w0 = im.naturalWidth || im.width;
  var h0 = im.naturalHeight || im.height;
  if (!w0 || !h0) { return null; }
  var dims = [240, 200, 160, 128, 96, 72];
  var qualities = [0.8, 0.7, 0.6, 0.5, 0.4];
  var smallest = null;
  for (var d = 0; d < dims.length; d++) {
    var scale = Math.min(1, dims[d] / Math.max(w0, h0));
    var w = Math.max(1, Math.round(w0 * scale));
    var h = Math.max(1, Math.round(h0 * scale));
    var c = document.createElement("canvas");
    c.width = w; c.height = h;
    var ctx = c.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, w, h);
    ctx.drawImage(im, 0, 0, w, h);
    for (var q = 0; q < qualities.length; q++) {
      var url = c.toDataURL("image/jpeg", qualities[q]); // throws if tainted
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
      var u = compressLoadedImg(img);
      if (u) { img.src = u; }
    } catch (e) { /* leave as-is; save will report if still too big */ }
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

function insertHtml(html) {
  sigEl.focus();
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
