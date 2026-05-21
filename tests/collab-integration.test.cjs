const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const root = path.join(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");

for (const id of [
  "collabEditorName",
  "startCollabBtn",
  "copyShareLinkBtn",
  "collabStatus",
  "collabUsersText",
  "collabLastSyncedText",
]) {
  assert.match(html, new RegExp(`id="${id}"`), `${id} should be present in index.html`);
}

for (const functionName of [
  "collectDraftPayload",
  "applyDraftPayload",
  "connectCollabSession",
  "disconnectCollab",
  "sendCollabPatch",
  "applyRemotePatch",
  "updateCollabStatus",
]) {
  assert.match(html, new RegExp(`function ${functionName}\\(`), `${functionName} should exist`);
}

assert.match(html, /URLSearchParams|searchParams\.set\("session"/, "session URL handling should preserve query params");
assert.match(html, /new WebSocket\(/, "frontend should connect via WebSocket");
assert.match(html, /api\/collab/, "frontend should use the collab API route");
assert.match(html, /queueRemotePatchUntilBlur/, "focused local fields should queue remote patches");

const pagesFunctionPath = path.join(root, "functions", "api", "collab", "[sessionId].js");
const workerPath = path.join(root, "worker", "index.mjs");
const rootWranglerPath = path.join(root, "wrangler.toml");
const workerWranglerPath = path.join(root, "worker", "wrangler.toml");
const readmePath = path.join(root, "README.md");

for (const file of [pagesFunctionPath, workerPath, rootWranglerPath, workerWranglerPath, readmePath]) {
  assert.equal(fs.existsSync(file), true, `${path.relative(root, file)} should exist`);
}

const pagesFunction = fs.readFileSync(pagesFunctionPath, "utf8");
assert.match(pagesFunction, /MOM_COLLAB_SESSIONS/, "Pages Function should use DO binding");
assert.match(pagesFunction, /idFromName\(sessionId\)/, "Pages Function should route by sessionId");

const rootWrangler = fs.readFileSync(rootWranglerPath, "utf8");
assert.match(rootWrangler, /pages_build_output_dir\s*=\s*"\."/);
assert.match(rootWrangler, /script_name\s*=\s*"generate-mom-collab-worker-dev-staging"/);

const workerWrangler = fs.readFileSync(workerWranglerPath, "utf8");
assert.match(workerWrangler, /main\s*=\s*"index\.mjs"/);
assert.match(workerWrangler, /new_sqlite_classes\s*=\s*\["MomCollabSession"\]/);

(async () => {
  const workerModule = await import(pathToFileURL(workerPath).href);
  assert.equal(typeof workerModule.default.fetch, "function");
  assert.equal(typeof workerModule.MomCollabSession, "function");
  assert.equal(typeof workerModule.parseCollabSessionId, "function");
  assert.equal(workerModule.parseCollabSessionId(new Request("https://example.test/api/collab/abc-123")), "abc-123");

  const draft = { table3State: [{ activity: "old" }] };
  workerModule.setDraftPath(draft, "table3State/0/activity", "new");
  assert.equal(draft.table3State[0].activity, "new");

  console.log("collab integration tests passed");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
