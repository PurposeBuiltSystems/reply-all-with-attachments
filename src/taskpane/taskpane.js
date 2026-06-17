/*
 * Signature settings pane.
 *
 * Stores the user's pasted signature (HTML) in roamingSettings so the
 * Reply All with Attachments command can add it to each Graph-built reply.
 * roamingSettings caps each entry at 32 KB, so the HTML is chunked.
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
});

function save() {
  var html = sigEl.innerHTML.trim();
  var rs = Office.context.roamingSettings;
  writeSignature(rs, html);
  status("Saving…");
  rs.saveAsync(function (res) {
    status(
      res.status === Office.AsyncResultStatus.Succeeded
        ? "Saved. It'll appear on your next Reply All with Attachments."
        : "Couldn't save: " + (res.error && res.error.message)
    );
  });
}

function status(t) { statusEl.textContent = t; }

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
