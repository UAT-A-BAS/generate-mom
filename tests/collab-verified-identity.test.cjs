const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");

function extractFunctionSource(source, name) {
  const start = source.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `${name} should exist`);
  const bodyStart = source.indexOf("{", start);
  let depth = 0;

  for (let index = bodyStart; index < source.length; index += 1) {
    if (source[index] === "{") depth += 1;
    if (source[index] === "}") depth -= 1;
    if (depth === 0) return source.slice(start, index + 1);
  }

  throw new Error(`${name} body should close`);
}

const identityPill = {
  hidden: true,
  textContent: "",
  title: "",
  attributes: {},
  setAttribute(name, value) {
    this.attributes[name] = String(value);
  },
  removeAttribute(name) {
    delete this.attributes[name];
  },
};
const context = {
  collabState: { editorName: "Legacy Editor", identity: null },
  elements: {
    collabEditorName: { value: "Legacy Editor" },
    collabIdentityText: identityPill,
  },
  getCollabEditorName() {
    return "Legacy Editor";
  },
};
vm.createContext(context);
vm.runInContext(extractFunctionSource(html, "normalizeCollabIdentity"), context);
vm.runInContext(extractFunctionSource(html, "applyCollabIdentity"), context);

const verifiedResult = vm.runInContext(
  `applyCollabIdentity({
    email: "u075060@bca.co.id",
    displayName: "Alex Marcello",
    verified: true
  })`,
  context
);
assert.equal(verifiedResult, true);
assert.equal(context.collabState.editorName, "Alex Marcello");
assert.equal(context.collabState.identity.email, "u075060@bca.co.id");
assert.equal(context.collabState.identity.verified, true);
assert.equal(identityPill.hidden, false);
assert.equal(identityPill.textContent, "Microsoft ✓ Alex Marcello");
assert.match(identityPill.attributes["aria-label"], /Microsoft terverifikasi/);
assert.equal(
  context.elements.collabEditorName.value,
  "Legacy Editor",
  "verified identity must not overwrite the legacy name stored for direct access"
);

const legacyResult = vm.runInContext(
  `applyCollabIdentity({
    email: "spoofed@example.test",
    displayName: "Spoofed Name",
    verified: false
  })`,
  context
);
assert.equal(legacyResult, false);
assert.equal(context.collabState.editorName, "Legacy Editor");
assert.equal(context.collabState.identity, null);
assert.equal(identityPill.hidden, true);
assert.equal(identityPill.textContent, "");

const initIdentityCalls = [];
const initContext = {
  collabState: {
    clientId: "browser-client",
    version: 0,
    lastSyncedAt: "",
  },
  applyCollabIdentity(identity) {
    initIdentityCalls.push(identity);
  },
  updateCollabStatus() {},
  applyRemotePatch() {},
  sendCollabFullPayload() {},
  window: { setTimeout() {} },
};
vm.createContext(initContext);
vm.runInContext(extractFunctionSource(html, "handleCollabSocketMessage"), initContext);
initContext.handleCollabSocketMessage({
  data: JSON.stringify({
    type: "init",
    serverClientId: "server-client-42",
    identity: {
      email: "u075060@bca.co.id",
      displayName: "Alex Marcello",
      verified: true,
    },
    version: 3,
    users: 1,
  }),
});
assert.equal(initContext.collabState.clientId, "server-client-42");
assert.equal(initContext.collabState.version, 3);
assert.equal(initIdentityCalls.length, 1);
assert.equal(initIdentityCalls[0].email, "u075060@bca.co.id");

console.log("collab verified identity tests passed");
