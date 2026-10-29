/* Regression: paged originals with duplicate files produce one verified draft. */
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const original = { "@odata.type": "#microsoft.graph.fileAttachment", name: "report.txt", size: 3, contentBytes: "YWJj" };
const copied = [], notices = [], opened = [];
let completed = 0, checks = 0;
const context = { setTimeout: (fn, ms) => { if (ms < 10000) { fn(); } return 1; }, clearTimeout() {},
  atob: text => Buffer.from(text, "base64").toString("binary"),
  Office: { onReady() {}, actions: { associate() {} },
    MailboxEnums: { RestVersion: { v2_0: "v2" }, ItemNotificationMessageType: { ErrorMessage: "error", InformationalMessage: "info" } },
    context: { roamingSettings: { get() { return null; } }, mailbox: {
      item: { itemId: "original", notificationMessages: { replaceAsync(key, notice) { notices.push(notice.message); } } },
      convertToRestId: id => id, convertToEwsId: id => id, displayMessageForm: id => opened.push(id) } } },
  fetch: async (url, opts) => {
    const endpoint = url.replace("https://graph.microsoft.com/v1.0", "");
    const method = opts.method;
    let body;
    if (endpoint === "/me/messages/original/createReplyAll") body = { id: "draft" };
    else if (endpoint === "/me/messages/draft?$select=body") body = { body: { content: "<p>Quoted</p>" } };
    else if (endpoint === "/me/messages/original/attachments") body = { value: [original],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages/original/attachments?next=2" };
    else if (endpoint === "/me/messages/original/attachments?next=2") body = { value: [original] };
    else if (endpoint === "/me/messages/draft/attachments?$select=name,size") body = { value: [] };
    else if (endpoint === "/me/messages/draft/attachments" && method === "POST") {
      const payload = JSON.parse(opts.body); copied.push(payload); body = { id: "copied" };
    } else if (endpoint === "/me/messages/draft/attachments?$select=id,name,isInline") {
      // Exchange reports a different size for the copy; matching is by name.
      checks++; body = { value: [{ name: "report.txt", size: 912, isInline: false }] };
    } else throw new Error("Unexpected request: " + method + " " + endpoint);
    return { ok: true, status: 200, json: async () => body };
  } };
vm.runInNewContext(fs.readFileSync(path.join(__dirname, "../src/commands/commands.js"), "utf8"), context);
context.getToken = async () => "token";
(async () => {
  await context.replyAllWithAttachments({ completed() { completed++; } });
  assert.equal(completed, 1);
  assert.equal(copied.length, 1, "duplicate source files copied once across pages");
  assert.equal(copied[0].contentBytes, "YWJj");
  assert.deepEqual(opened, ["draft"]);
  assert.equal(checks, 1, "verification counts distinct copied files");
  assert.equal(notices.some(n => /still syncing|missing|failed/i.test(n)), false, "duplicates do not trigger a false missing-attachments warning");
  console.log("All Reply All command integration tests passed.");
})().catch(err => { console.error(err); process.exitCode = 1; });
