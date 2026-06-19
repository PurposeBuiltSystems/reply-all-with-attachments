/*
 * Signature settings pane.
 *
 * Stores the user's pasted signature (HTML) in roamingSettings so the
 * Reply All with Attachments command can add it to each Graph-built reply.
 * roamingSettings caps each entry at 32 KB, so the HTML is chunked.
 *
 * IMAGES: a pasted signature brings its logo in as a temporary blob:/cid:/remote
 * reference, NOT the actual bytes — so the logo would silently drop on save. We
 * embed images as base64 data URIs instead, which persist. Two paths:
 *   1. After paste, auto-convert any loaded <img> to a data URI (canvas), and
 *      fall back to the image bytes the clipboard carried alongside the HTML.
 *   2. An "Add image…" file picker that embeds the chosen file directly — the
 *      guaranteed path when the logo is cross-origin (canvas would be tainted).
 */

/* global Office */

var META = "raa.sig.meta";        // number of chunks
var CHUNK = "raa.sig.";           // raa.sig.0, raa.sig.1, ...
var CHUNK_SIZE = 30000;

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
        insertHtml('<img src="' + url + '" />');
        status("Logo added. Click Save to keep it.");
      });
    }
    e.target.value = "";
  };

  // After a paste, try to turn temporary image references into embedded bytes.
  sigEl.addEventListener("paste", function (e) {
    var clipFiles = imageFilesFromClipboard(e.clipboardData);
    setTimeout(function () { embedImages(clipFiles); }, 0);
  });
});

function save() {
  var html = sigEl.innerHTML.trim();
  var rs = Office.context.roamingSettings;
  writeSignature(rs, html);
  status("Saving…");
  var unembedded = countUnembeddedImages();
  rs.saveAsync(function (res) {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      status("Couldn't save: " + (res.error && res.error.message));
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

/* ---- image embedding ---- */

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

// Convert every non-embedded <img> to a base64 data URI.
// First try canvas (works for same-origin / blob:); fall back to the raw
// image bytes the clipboard provided, in order.
function embedImages(clipFiles) {
  readFilesAsDataUrls(clipFiles, function (clipUrls) {
    var clipIdx = 0;
    var imgs = sigEl.querySelectorAll("img");
    for (var i = 0; i < imgs.length; i++) {
      var img = imgs[i];
      if (img.src.indexOf("data:") === 0) { continue; }
      var viaCanvas = canvasDataUrl(img);
      if (viaCanvas) { img.src = viaCanvas; continue; }
      if (clipIdx < clipUrls.length) { img.src = clipUrls[clipIdx++]; }
    }
  });
}

function canvasDataUrl(img) {
  try {
    if (!img.naturalWidth) { return null; }
    var c = document.createElement("canvas");
    c.width = img.naturalWidth;
    c.height = img.naturalHeight;
    c.getContext("2d").drawImage(img, 0, 0);
    return c.toDataURL("image/png"); // throws if the image tainted the canvas
  } catch (e) {
    return null;
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
