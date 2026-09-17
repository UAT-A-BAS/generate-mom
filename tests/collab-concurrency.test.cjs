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

// The client must never mint its own sequence number: the server owns ordering, so the
// only thing the client sends is the last sequence it has actually seen.
assert.doesNotMatch(
  html,
  /collabState\.version/,
  "the client must not track a locally invented version cursor"
);
assert.match(
  html,
  /baseVersion:\s*collabState\.seq/,
  "every outgoing collaboration message should carry the last server sequence it saw"
);
assert.doesNotMatch(
  html,
  /version:\s*collabState\.seq\s*\+\s*1/,
  "the client must not optimistically advance a sequence number"
);
assert.match(
  html,
  /function sendCollabFullPayload\(options = \{}\)[\s\S]*?replace:\s*Boolean\(options\.replace\)/,
  "full snapshots should distinguish explicit replacements from session seeding"
);

function buildContext({ socketOpen = true } = {}) {
  const sentMessages = [];
  const clearedTimers = [];
  const context = {
    WebSocket: { OPEN: 1 },
    window: {
      clearTimeout(timer) {
        clearedTimers.push(timer);
      },
      setTimeout() {
        return 0;
      },
    },
    collabState: {
      active: true,
      offline: false,
      connected: socketOpen,
      sessionId: "session-1",
      clientId: "editor-a",
      editorName: "Editor A",
      socket: socketOpen
        ? {
            readyState: 1,
            send(raw) {
              sentMessages.push(JSON.parse(raw));
            },
          }
        : { readyState: 3, send() {} },
      seq: 7,
      pendingTimers: new Map([
        ["table1ProjectsState/0/projectName", { timer: 101, value: "Project Alpha" }],
        ["table3State/0/activity", { timer: 102, value: "Deploy service" }],
      ]),
      pendingStructural: new Map(),
      retryingPaths: new Map(),
    },
    collabMaxOpsPerMessage: 120,
    collabMaxMessageBytes: 850000,
    createCollabMutationId: () => "mutation-1",
    updateCollabStatus() {},
    showFeedback() {},
    // Signature bookkeeping lives outside this unit under test; the outbox behaviour is
    // what this file pins down.
    refreshSignaturesForTouchedPaths() {},
  };
  vm.createContext(context);
  for (const functionName of [
    "isCollabSocketOpen",
    "clearPendingCollabPatchTimers",
    "sendCollabPatch",
    "sendCollabStructuralPatch",
    "flushPendingCollabChanges",
  ]) {
    vm.runInContext(extractFunctionSource(html, functionName), context);
  }
  return { context, sentMessages, clearedTimers };
}

// 1) A healthy flush sends one op per changed field, so two editors working in different
//    tables cannot overwrite each other.
const healthy = buildContext();
assert.equal(healthy.context.flushPendingCollabChanges(), true);
assert.deepEqual(
  healthy.sentMessages.map(({ type, ops }) => ({ type, ops: ops.map((op) => op.path) })),
  [
    { type: "patch", ops: ["table1ProjectsState/0/projectName"] },
    { type: "patch", ops: ["table3State/0/activity"] },
  ],
  "flush must send one patch per changed field instead of one full snapshot"
);
assert.equal(
  healthy.sentMessages.every((message) => message.type === "patch"),
  true,
  "flush must never replace the whole shared draft"
);
assert.deepEqual(healthy.clearedTimers, [101, 102]);
assert.equal(healthy.context.collabState.pendingTimers.size, 0);

// 2) A flush attempted while the socket is down must keep the queue so reconnect replays
//    it. Dropping it here is the silent data loss the previous protocol had.
const offline = buildContext({ socketOpen: false });
assert.equal(offline.context.flushPendingCollabChanges(), false);
assert.equal(offline.sentMessages.length, 0);
assert.equal(
  offline.context.collabState.pendingTimers.get("table3State/0/activity")?.value,
  "Deploy service",
  "queued edits must survive a flush attempted on a closed socket"
);
assert.equal(
  offline.context.collabState.pendingTimers.size,
  2,
  "a failed flush must not discard the outbox"
);

// 3) Structural edits travel as the enclosing array, not the whole draft.
const structural = buildContext();
structural.context.sendCollabStructuralPatch = structural.context.sendCollabStructuralPatch;
structural.context.collabState.pendingStructural.set("table3State", {
  path: "table3State",
  value: [{ activity: "row one" }, { activity: "row two" }],
});
assert.equal(structural.context.flushPendingCollabChanges(), true);
const structuralMessages = structural.sentMessages.filter((message) =>
  message.ops.some((op) => op.path === "table3State")
);
assert.equal(structuralMessages.length, 1, "the queued structural array should be flushed");
assert.deepEqual(
  structuralMessages[0].ops.map((op) => ({ path: op.path, value: op.value })),
  [{ path: "table3State", value: [{ activity: "row one" }, { activity: "row two" }] }]
);
assert.equal(
  structural.sentMessages.every((message) => message.ops.every((op) => op.path !== "draft")),
  true,
  "structural flushes must not fall back to a whole-draft replacement"
);
assert.equal(structural.context.collabState.pendingStructural.size, 0);

console.log("collab concurrency tests passed");
