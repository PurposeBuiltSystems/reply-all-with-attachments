/* Exercise first-click draft opening and signature replication against real command code. */
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { JSDOM } = require("jsdom");
function load(options) {
  const storedSig = '<p align="center"><img width="100" src="data:image/png;base64,YWJj"></p>';
  const calls = [], opens = [], callbacks = [], notices = [], created = [], inline = [];
  let completed = 0, bodyReads = 0, body = "", fileReads = 0;
  let notFoundLeft = options.draftNotFound || 0;
  const context = {
    document: new JSDOM("", { url: "https://purposebuiltsystems.github.io/reply-all-with-attachments/src/commands/commands.html" }).window.document,
    setTimeout(fn, ms) { if (ms < 10000) { fn(); } return 1; }, clearTimeout() {},
    atob: text => Buffer.from(text, "base64").toString("binary"),
    Office: { onReady() {}, actions: { associate() {} }, AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
      MailboxEnums: { RestVersion: { v2_0: "v2" }, ItemNotificationMessageType: { ErrorMessage: "error", InformationalMessage: "info" } },
      context: { roamingSettings: { get(key) { return key === "raa.sig.meta" ? 1 : key === "raa.sig.0" ? storedSig : null; } },
        mailbox: { item: { itemId: "source", attachments: options.emptyAttachments ? [] : undefined, notificationMessages: { replaceAsync(key, notice) { notices.push(notice.message); } } },
          convertToRestId: id => id, convertToEwsId: id => id,
          displayMessageFormAsync(id, callback) { opens.push(id); callbacks.push(callback); } } } },
    fetch: async (url, opts) => {
      const endpoint = url.replace("https://graph.microsoft.com/v1.0", "");
      const method = opts.method; calls.push(method + " " + endpoint);
      if (method === "GET" && endpoint.startsWith("/me/messages/draft") && notFoundLeft > 0) {
        notFoundLeft--;
        return { ok: false, status: 404, text: async () => "Draft not visible yet" };
      }
      if (options.verificationError && endpoint.endsWith("/attachments?$select=name,isInline,contentId")) {
        return { ok: false, status: 400, text: async () => "Read unavailable" };
      }
      let result;
      if (endpoint.endsWith("/createReplyAll")) { created.push("draft"); result = { id: "draft" }; if (options.returnBody) { result.body = { content: "<html><body>Quoted</body></html>" }; } }
      else if (endpoint === "/me/messages/draft?$select=body") {
        bodyReads++; result = { body: { content: bodyReads < 3 || options.neverReady ? "<html><body>Quoted</body></html>" : body } };
      } else if (endpoint === "/me/messages/draft" && method === "PATCH") {
        body = JSON.parse(opts.body).body.content; result = {};
      } else if (endpoint === "/me/messages/draft/attachments" && method === "POST") {
        inline.push(JSON.parse(opts.body)); result = { id: "logo" };
      } else if (endpoint === "/me/messages/draft/attachments?$select=name,isInline,contentId") {
        if (options.missingFiles) { fileReads++; }
        result = { value: inline };
      } else if (endpoint === "/me/messages/source/attachments") {
        result = { value: options.missingFiles ? [{ "@odata.type": "#microsoft.graph.fileAttachment", name: "report.txt", size: 3, contentBytes: "YWJj" }] : [] };
      } else if (endpoint === "/me/messages/draft/attachments?$select=name,size") result = { value: [] };
      else throw new Error("Unexpected request: " + method + " " + endpoint);
      return { ok: true, status: 200, json: async () => result };
    } };
  context.self = context;
  for (const file of ["signature-layout.js", "commands/commands.js"]) {
    vm.runInNewContext(fs.readFileSync(path.join(__dirname, "../src", file), "utf8"), context);
  }
  context.getToken = async () => "token";
  return { context, calls, opens, callbacks, notices, created, inline,
    completed: () => completed, fileReads: () => fileReads,
    run: () => context.replyAllWithAttachments({ completed() { completed++; } }) };
}
(async () => {
  const fresh = load({ draftNotFound: 2 });
  const freshRun = fresh.run();
  await new Promise(setImmediate);
  assert.equal(fresh.opens.length, 1, "transient new-draft 404 reads recover within the first click");
  fresh.callbacks.shift()({ status: "succeeded" });
  await freshRun;
  assert.equal(fresh.created.length, 1);
  assert.equal(fresh.inline.length, 1, "read retries do not replay signature uploads");
  assert.equal(fresh.completed(), 1);
  const interrupted = load({ verificationError: true });
  const interruptedRun = interrupted.run();
  await new Promise(setImmediate);
  assert.equal(interrupted.opens.length, 1, "verification failure still opens the fully written draft");
  interrupted.callbacks.shift()({ status: "succeeded" });
  await interruptedRun;
  assert.equal(interrupted.created.length, 1);
  assert.equal(interrupted.inline.length, 1);
  assert.ok(interrupted.notices.some(n => n.includes("verification was interrupted")));
  assert.equal(interrupted.completed(), 1);
  const fast = load({ returnBody: true, emptyAttachments: true });
  const fastRun = fast.run();
  await new Promise(setImmediate);
  assert.equal(fast.opens.length, 1);
  const patchIndex = fast.calls.indexOf("PATCH /me/messages/draft");
  assert.ok(patchIndex > 0);
  assert.equal(fast.calls.slice(0, patchIndex).some(c => c.includes("?$select=body")), false,
    "reuse createReplyAll body instead of a pre-write GET");
  assert.equal(fast.calls.some(c => c === "GET /me/messages/source/attachments"), false,
    "explicitly empty Office attachments skip the source scan");
  assert.equal(fast.calls.some(c => c.endsWith("/attachments?$select=name,size")), false,
    "no source attachments means no duplicate scan");
  assert.ok(fast.calls.some(c => c.endsWith("/attachments?$select=name,isInline,contentId")),
    "signature verification still runs on the faster path");
  fast.callbacks.shift()({ status: "succeeded" });
  await fastRun;
  assert.equal(fast.completed(), 1);
  const g = load({});
  const running = g.run();
  await new Promise(setImmediate);
  assert.equal(g.completed(), 0, "command remains active while Outlook is opening");
  assert.equal(g.opens.length, 1);
  g.callbacks.shift()({ status: "failed", error: { code: 9049, message: "Not synced yet" } });
  await new Promise(setImmediate);
  assert.equal(g.opens.length, 2, "retry a definite first-open failure automatically");
  assert.equal(g.completed(), 0);
  g.callbacks.shift()({ status: "succeeded" });
  await running;
  assert.equal(g.completed(), 1);
  assert.deepEqual(g.created, ["draft"], "only one reply draft is created");
  assert.deepEqual(g.opens, ["draft", "draft"], "retry the same populated draft");
  assert.equal(g.inline.length, 1, "logo uploaded once despite stale reads and open retries");
  assert.equal(g.calls.some(c => c.startsWith("DELETE")), false, "stale body cannot delete the logo");
  // A draft that never verifies still opens (it already exists in Drafts),
  // with a notice naming what to check, instead of an error and no draft.
  const never = load({ neverReady: true });
  const neverRun = never.run();
  await new Promise(setImmediate);
  assert.equal(never.opens.length, 1, "open the saved draft even when the logo hasn't synced");
  never.callbacks.shift()({ status: "succeeded" });
  await neverRun;
  assert.equal(never.inline.length, 1);
  assert.equal(never.calls.some(c => c.startsWith("DELETE")), false);
  assert.ok(never.notices.some(n => n.includes("logo is still syncing")));
  assert.equal(never.completed(), 1);
  const missing = load({ missingFiles: true });
  // Mock the file-copy helper only; readiness checks still execute the real code.
  missing.context.copyFileAttachment = async () => {};
  const missingRun = missing.run();
  await new Promise(setImmediate);
  assert.equal(missing.fileReads(), 5, "bounded wait for copied files");
  assert.equal(missing.opens.length, 1, "open anyway once the wait runs out");
  missing.callbacks.shift()({ status: "succeeded" });
  await missingRun;
  assert.ok(missing.notices.some(n => n.includes("reopen it from Drafts")));
  assert.equal(missing.completed(), 1);
  const cachedPage = load({});
  delete cachedPage.context.SignatureLayout;
  cachedPage.context.document.head.appendChild = script => {
    assert.ok(script.src.endsWith("/src/signature-layout.js?v=2"));
    vm.runInNewContext(fs.readFileSync(path.join(__dirname, "../src/signature-layout.js"), "utf8"), cachedPage.context);
    script.onload();
  };
  const cachedRun = cachedPage.run();
  await new Promise(setImmediate);
  assert.equal(cachedPage.opens.length, 1, "a cached HTML page loads the new module before its first draft");
  cachedPage.callbacks.shift()({ status: "succeeded" });
  await cachedRun;
  assert.equal(cachedPage.created.length, 1);
  assert.equal(cachedPage.completed(), 1);
  const failedOpen = load({});
  failedOpen.context.Office.context.mailbox.displayMessageFormAsync = (id, cb) => {
    failedOpen.opens.push(id); cb({ status: "failed", error: { message: "Unavailable" } });
  };
  await failedOpen.run();
  assert.equal(failedOpen.opens.length, 3, "definite open failures have bounded retries");
  assert.equal(failedOpen.created.length, 1);
  assert.equal(failedOpen.completed(), 1);
  assert.ok(failedOpen.notices.some(n => n.includes("Open Drafts")));
  const timeout = load({});
  let expire;
  timeout.context.setTimeout = (fn, ms) => { if (ms === 10000) { expire = fn; } else if (ms < 10000) { fn(); } return 1; };
  const timedRun = timeout.run();
  await new Promise(setImmediate);
  assert.equal(timeout.completed(), 0);
  assert.equal(typeof expire, "function");
  expire();
  await timedRun;
  assert.equal(timeout.opens.length, 1, "an unknown timeout outcome does not open a duplicate window");
  assert.equal(timeout.completed(), 1);
  const legacy = load({});
  delete legacy.context.Office.context.mailbox.displayMessageFormAsync;
  legacy.context.Office.context.mailbox.displayMessageForm = id => legacy.opens.push(id);
  await legacy.run();
  assert.equal(legacy.opens.length, 1, "older Outlook clients retain their existing open API");
  assert.equal(legacy.completed(), 1);
  console.log("All draft-readiness and first-click regression tests passed.");
})().catch(err => { console.error(err); process.exitCode = 1; });
