const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.join(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");

assert.match(html, /const collabDebounceMs = 3000;/, "collab autosave debounce should be 3 seconds");
assert.match(html, /const collabIdleMs = 5 \* 60 \* 1000;/, "idle timer should be 5 minutes");
assert.match(html, /const collabHiddenGraceMs = 60 \* 1000;/, "hidden-tab grace period should be 60 seconds");
assert.match(
  html,
  /const collabActivityEvents = \["keydown", "input", "click", "mousemove", "scroll", "touchstart"\];/,
  "activity detector should cover required browser activity events"
);

for (const functionName of [
  "flushPendingCollabChanges",
  "pauseCollabSession",
  "resumeCollabSession",
  "handleCollabVisibilityChange",
  "cleanupCollabBeforeUnload",
  "shouldReconnectCollab",
  "scheduleCollabReconnect",
]) {
  assert.match(html, new RegExp(`function ${functionName}\\(`), `${functionName} should exist`);
}

assert.match(
  html,
  /function\s+flushPendingCollabChanges\(\)\s*{[\s\S]*?clearPendingCollabPatchTimers\(\);[\s\S]*?return sendCollabFullPayload\(\);[\s\S]*?}/,
  "pending debounced changes should flush as a full draft before disconnect"
);
assert.match(
  html,
  /function\s+pauseCollabSession\(reason\)\s*{[\s\S]*?flushPendingCollabChanges\(\);[\s\S]*?closeCollabSocket\(`Collab \$\{reason\}`,\s*collabFlushCloseDelayMs\);[\s\S]*?}/,
  "idle or hidden pause should flush and gracefully close the socket"
);
assert.match(
  html,
  /function\s+shouldReconnectCollab\(\)\s*{[\s\S]*?!collabState\.idle[\s\S]*?!document\.hidden[\s\S]*?}/,
  "reconnect should be disabled while idle or hidden"
);
assert.match(
  html,
  /socket\.addEventListener\("close",[\s\S]*?if \(!shouldReconnectCollab\(\)\) {[\s\S]*?return;[\s\S]*?}[\s\S]*?scheduleCollabReconnect\(\);[\s\S]*?}\);/,
  "socket close should only reconnect when idle/visibility policy allows it"
);
assert.match(
  html,
  /document\.addEventListener\("visibilitychange", handleCollabVisibilityChange\);/,
  "Page Visibility API should drive hidden-tab lifecycle"
);
assert.match(html, /window\.addEventListener\("beforeunload", cleanupCollabBeforeUnload\);/);
assert.match(html, /window\.addEventListener\("pagehide", cleanupCollabBeforeUnload\);/);
assert.match(
  html,
  /target\.addEventListener\(eventName, markCollabUserActive, { capture: true, passive: true }\);/,
  "activity events should reset idle state without blocking UI events"
);

console.log("collab idle lifecycle tests passed");
