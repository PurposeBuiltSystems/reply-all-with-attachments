/* Cold host registration and authentication initialization recovery. */
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const source = fs.readFileSync(path.join(__dirname, "../src/commands/commands.js"), "utf8");
function load() {
  let ready;
  const context = {
    Office: { onReady(cb) { ready = cb; }, context: { roamingSettings: {
      get() { return false; }, set() {}, saveAsync(cb) { cb(); }
    } } },
    setTimeout(fn, ms) { if (ms < 10000) { fn(); } return 1; }, clearTimeout() {}
  };
  vm.runInNewContext(source, context);
  return { context, ready: () => ready() };
}
(async () => {
  const cold = load();
  const associations = [];
  cold.context.Office.actions = { associate(name, fn) { associations.push({ name, fn }); } };
  cold.ready();
  assert.equal(associations.length, 1, "register when actions become available at host readiness");
  assert.equal(associations[0].name, "replyAllWithAttachments");
  assert.equal(associations[0].fn, cold.context.replyAllWithAttachments);
  cold.ready();
  assert.equal(associations.length, 1, "readiness must not register twice");

  let creates = 0, popups = 0;
  const client = {
    acquireTokenSilent: async () => ({ accessToken: "token" }),
    acquireTokenPopup: async () => { popups++; return { accessToken: "popup" }; }
  };
  cold.context.msal = { createNestablePublicClientApplication: async () => {
    creates++;
    if (creates === 1) { throw new Error("Broker starting"); }
    return client;
  } };
  assert.equal(await cold.context.getToken(), "token", "recover broker startup failure on the same click");
  assert.equal(creates, 2);
  assert.equal(popups, 0, "broker recovery does not open an unnecessary sign-in popup");
  assert.equal(await cold.context.getToken(), "token");
  assert.equal(creates, 2, "reuse successful broker initialization");

  const failed = load();
  let failures = 0;
  failed.context.msal = { createNestablePublicClientApplication: async () => {
    failures++;
    throw new Error("Broker unavailable");
  } };
  await assert.rejects(failed.context.getToken(), /Broker unavailable/);
  assert.equal(failures, 2, "initialization retries stay bounded");
  assert.equal(failed.context.pcaPromise, null, "rejected initialization isn't cached");
  failed.context.msal.createNestablePublicClientApplication = async () => client;
  assert.equal(await failed.context.getToken(), "token", "a later attempt can recover without restarting Outlook");
  console.log("All cold-start regression tests passed.");
})().catch(err => { console.error(err); process.exitCode = 1; });
