/* Exchange partial writes and Classic Outlook cached-draft opening regressions. */
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const path = require("node:path");
const source = fs.readFileSync(path.join(__dirname, "../src/commands/commands.js"), "utf8");
function load(options = {}) {
  const writes = [], notices = [], waits = [], opens = [], attached = [];
  let completed = 0, activeWrites = 0, maxWrites = 0;
  const files = ["same.pdf", "same.pdf", "third.txt"].map((name, i) => ({
    id: "source-" + i, name, size: 3, contentBytes: Buffer.from("abc" + i).toString("base64"),
    "@odata.type": "#microsoft.graph.fileAttachment"
  }));
  const context = {
    setTimeout(fn, ms) { if (ms < 10000) { waits.push(ms); fn(); } return 1; }, clearTimeout() {},
    atob: text => Buffer.from(text, "base64").toString("binary"),
    Office: { onReady() {}, actions: { associate() {} }, AsyncResultStatus: { Succeeded: "succeeded" },
      MailboxEnums: { RestVersion: { v2_0: "v2" }, ItemNotificationMessageType: { ErrorMessage: "error", InformationalMessage: "info" } },
      context: { platform: "PC", roamingSettings: { get() { return null; } }, mailbox: {
        diagnostics: { hostName: "Outlook" }, item: { itemId: "source", notificationMessages: {
          replaceAsync(key, notice) { notices.push(notice); }
        } }, convertToRestId: id => id, convertToEwsId: id => id,
        displayMessageFormAsync(id, cb) {
          opens.push(id);
          cb(opens.length <= (options.openFailures || 0) ?
            { status: "failed", error: { code: options.openCode || 9049, message: "Not synced" } } : { status: "succeeded" });
        }
      } } },
    fetch: async (url, opts) => {
      const endpoint = url.replace("https://graph.microsoft.com/v1.0", "");
      let body;
      if (endpoint.endsWith("/createReplyAll")) body = { id: "draft", body: { content: "<p>Quoted</p>" } };
      else if (endpoint === "/me/messages/source/attachments") body = { value: files };
      else if (endpoint.endsWith("/attachments?$select=name,size")) body = { value: [] };
      else if (endpoint.endsWith("/attachments?$select=id,name,isInline")) body = { value: attached };
      else if (endpoint === "/me/messages/draft/attachments" && opts.method === "POST") {
        const file = JSON.parse(opts.body);
        writes.push(file); activeWrites++; maxWrites = Math.max(maxWrites, activeWrites);
        await new Promise(setImmediate); activeWrites--;
        if (options.failSecond && writes.length === 2) {
          return { ok: false, status: 400, text: async () => "Copy failed" };
        }
        attached.push({ name: file.name, isInline: false }); body = { id: "attached" };
      } else throw new Error("Unexpected request " + endpoint);
      return { ok: true, status: 200, json: async () => body };
    }
  };
  vm.runInNewContext(source, context);
  context.getToken = async () => "token";
  return { context, writes, attached, notices, waits, opens, maxWrites: () => maxWrites,
    completed: () => completed, run: () => context.replyAllWithAttachments({ completed() { completed++; } }) };
}
(async () => {
  const all = load({ openFailures: 4 });
  await all.run();
  assert.equal(all.writes.length, 3, "distinct source IDs survive identical names and sizes");
  assert.equal(all.attached.length, 3);
  assert.equal(all.maxWrites(), 1, "no simultaneous writes to the same Exchange draft");
  assert.deepEqual(all.opens, ["draft", "draft", "draft", "draft", "draft"], "Classic retries the same saved draft");
  assert.deepEqual(all.waits, [1500, 3000, 4500, 6000]);
  assert.equal(all.completed(), 1);
  assert.match(all.notices.at(-1).message, /Reply draft ready/);

  const partial = load({ failSecond: true });
  await partial.run();
  assert.equal(partial.writes.length, 3, "failed middle copy cannot prevent the last file");
  assert.equal(partial.attached.length, 2);
  assert.deepEqual(partial.opens, ["draft"], "open partial saved draft for review instead of abandoning it");
  assert.match(partial.notices.at(-1).message, /CHECK ATTACHMENTS: 1/);
  assert.equal(partial.notices.at(-1).persistent, true);
  assert.equal(partial.completed(), 1);

  const denied = load({ openFailures: 1, openCode: 9000 });
  await denied.run();
  assert.equal(denied.opens.length, 1, "unrelated Classic errors are not retried as sync delays");
  assert.match(denied.notices.at(-1).message, /Open Drafts/);
  assert.equal(denied.notices.at(-1).persistent, true);
  assert.equal(denied.completed(), 1);
  console.log("All Windows partial-copy and cached-draft recovery tests passed.");
})().catch(err => { console.error(err); process.exitCode = 1; });
