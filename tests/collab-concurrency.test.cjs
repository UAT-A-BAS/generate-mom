const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");

assert.match(
  html,
  /baseVersion:\s*collabState\.version/,
  "every outgoing collaboration message should carry its server base version"
);
assert.match(
  html,
  /function sendCollabFullPayload\(options = \{}\)[\s\S]*?replace:\s*Boolean\(options\.replace\)/,
  "full snapshots should distinguish explicit replacements from session seeding"
);

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

const sentMessages = [];
const clearedTimers = [];
const context = {
  WebSocket: { OPEN: 1 },
  window: {
    clearTimeout(timer) {
      clearedTimers.push(timer);
    },
  },
  collabState: {
    active: true,
    sessionId: "session-1",
    clientId: "editor-a",
    editorName: "Editor A",
    socket: {
      readyState: 1,
      send(raw) {
        sentMessages.push(JSON.parse(raw));
      },
    },
    version: 7,
    pendingTimers: new Map([
      ["table1ProjectsState/0/projectName", { timer: 101, value: "Project Alpha" }],
      ["table3State/0/activity", { timer: 102, value: "Deploy service" }],
    ]),
  },
  updateCollabStatus() {},
  sendCollabFullPayload() {
    sentMessages.push({ type: "full" });
    return true;
  },
};
vm.createContext(context);

for (const functionName of [
  "isCollabSocketOpen",
  "clearPendingCollabPatchTimers",
  "sendCollabPatch",
  "flushPendingCollabChanges",
]) {
  vm.runInContext(extractFunctionSource(html, functionName), context);
}

assert.equal(context.flushPendingCollabChanges(), true);
assert.deepEqual(
  sentMessages.map(({ type, path, value }) => ({ type, path, value })),
  [
    {
      type: "patch",
      path: "table1ProjectsState/0/projectName",
      value: "Project Alpha",
    },
    {
      type: "patch",
      path: "table3State/0/activity",
      value: "Deploy service",
    },
  ],
  "idle flush must preserve concurrent edits by sending one patch per changed field"
);
assert.deepEqual(clearedTimers, [101, 102]);
assert.equal(context.collabState.pendingTimers.size, 0);

console.log("collab concurrency tests passed");
