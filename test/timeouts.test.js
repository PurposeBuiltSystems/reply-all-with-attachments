/* Real command regression: stalled transport/body completes with draft recovery. */
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const path = require("node:path");
const source = fs.readFileSync(path.join(__dirname, "../src/commands/commands.js"), "utf8");
function load(fetcher) {
  const timers = new Map(), notices = [], calls = [];
  let next = 0, completed = 0;
  const context = { AbortController, Promise,
    setTimeout(fn, ms) { const id = ++next; timers.set(id, { fn, ms }); return id; },
    clearTimeout(id) { timers.delete(id); },
    fetch(url, opts) { calls.push({ url, opts }); return fetcher(url, opts, calls.length); },
    Office: { onReady() {}, actions: { associate() {} },
      MailboxEnums: { RestVersion: { v2_0: "v2" }, ItemNotificationMessageType: { ErrorMessage: "error", InformationalMessage: "info" } },
      context: { roamingSettings: { get() { return null; } }, mailbox: {
        item: { itemId: "original", attachments: [], notificationMessages: {
          replaceAsync(key, notice) { notices.push(notice.message); }
        } }, convertToRestId: id => id
      } } }
  };
  vm.runInNewContext(source, context);
  context.getToken = async () => "token";
  return { context, calls, notices, timers,
    run: () => context.replyAllWithAttachments({ completed() { completed++; } }),
    completed: () => completed,
    expire() { const entry = [...timers].find(([, t]) => t.ms === 30000); assert.ok(entry, "request has a deadline"); timers.delete(entry[0]); entry[1].fn(); }
  };
}
async function flush() { for (let i = 0; i < 30; i++) { await Promise.resolve(); } }
(async () => {
  for (const stalledBody of [false, true]) {
    const g = load(() => stalledBody ? Promise.resolve({ ok: true, status: 201, json: () => new Promise(() => {}) }) : new Promise(() => {}));
    const running = g.run();
    await flush();
    assert.ok(g.notices.includes("Creating your reply draft…"));
    g.expire();
    await running;
    assert.equal(g.completed(), 1, "Outlook command always completes");
    assert.equal(g.calls.length, 1, "uncertain draft creation is never replayed");
    assert.match(g.notices.at(-1), /^Reply interrupted\. Check Drafts before trying again\./);
    assert.match(g.notices.at(-1), /timed out/);
    assert.equal(g.timers.size, 0, "no deadline timers left behind");
    if (!stalledBody) { assert.equal(g.calls[0].opts.signal.aborted, true); }
  }
  const failedWrite = load(async () => ({ status: 503, ok: false, headers: { get: () => "0" } }));
  assert.equal((await failedWrite.context.fetchRetry("url", { method: "POST" })).status, 503);
  assert.equal(failedWrite.calls.length, 1, "ambiguous service errors never replay writes");
  console.log("All stalled-request and draft recovery tests passed.");
})().catch(err => { console.error(err); process.exitCode = 1; });
