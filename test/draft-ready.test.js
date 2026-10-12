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
  const context = {
    document: new JSDOM("", { url: "https://purposebuiltsystems.github.io/reply-all-with-attachments/src/commands/commands.html" }).window.document,
    setTimeout(fn, ms) { if (ms < 10000) { fn(); } return 1; }, clearTimeout() {},
    atob: text => Buffer.from(text, "base64").toString("binary"),
    Office: { onReady() {}, actions: { associate() {} }, AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
      MailboxEnums: { RestVersion: { v2_0: "v2" }, ItemNotificationMessageType: { ErrorMessage: "error", InformationalMessage: "info" } },
      context: { roamingSettings: { get(key) { return key === "raa.sig.meta" ? 1 : key === "raa.sig.0" ? storedSig : null; } },
        mailbox: { item: { itemId: "source", notificationMessages: { replaceAsync(key, notice) { notices.push(notice.message); } } },
          convertToRestId: id => id, convertToEwsId: id => id,
          displayMessageFormAsync(id, callback) { opens.push(id); callbacks.push(callback); } } } },
    fetch: async (url, opts) => {
      const endpoint = url.replace("https://graph.microsoft.com/v1.0", "");
      const method = opts.method; calls.push(method + " " + endpoint);
      let result;
      if (endpoint.endsWith("/createReplyAll")) { created.push("draft"); result = { id: "draft" }; }
      else if (endpoint === "/me/messages/draft?$select=body") {
        bodyReads++; result = { body: { content: bodyReads < 3 || options.neverReady ? "<html><body>Quoted</body></html>" : body } };
      } else if (endpoint === "/me/messages/draft" && method === "PATCH") {
        body = JSON.parse(opts.body).body.content; result = {};
      } else if (endpoint === "/me/messages/draft/attachments" && method === "POST") {
        inline.push(JSON.parse(opts.body)); result = { id: "logo" };
      } else if (endpoint === "/me/messages/draft/attachments?$select=isInline,contentId") {
        result = { value: inline };
      } else if (endpoint === "/me/messages/source/attachments") {
        result = { value: options.missingFiles ? [{ "@odata.type": "#microsoft.graph.fileAttachment", name: "report.txt", size: 3, contentBytes: "YWJj" }] : [] };
      } else if (endpoint === "/me/messages/draft/attachments?$select=name,size") result = { value: [] };
      else if (endpoint === "/me/messages/draft/attachments?$select=name,size,isInline") { fileReads++; result = { value: [] }; }
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
  const never = load({ neverReady: true });
  await never.run();
  assert.equal(never.opens.length, 0, "do not open a draft with unverified signature images");
  assert.equal(never.inline.length, 1);
  assert.equal(never.calls.some(c => c.startsWith("DELETE")), false);
  assert.ok(never.notices.some(n => n.includes("still syncing")));
  assert.equal(never.completed(), 1);
  const missing = load({ missingFiles: true });
  // Mock the file-copy helper only; readiness checks still execute the real code.
  missing.context.copyFileAttachment = async () => {};
  await missing.run();
  assert.equal(missing.opens.length, 0, "never open before copied files are visible");
  assert.equal(missing.fileReads(), 5);
  assert.equal(missing.completed(), 1);
  const cachedPage = load({});
  delete cachedPage.context.SignatureLayout;
  cachedPage.context.document.head.appendChild = script => {
    assert.ok(script.src.endsWith("/src/signature-layout.js?v=1"));
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
  timeout.context.setTimeout = (fn, ms) => { if (ms === 10000) { expire = fn; } else { fn(); } return 1; };
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
