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
assert.match(html, /id="collabEditorName"\s+type="hidden"|id="collabEditorName"[^>]*type="hidden"/, "editor name should be hidden");
assert.match(
  html,
  /if\s*\(\s*applyDraftDataFromJson\(raw\)\s*\)\s*{\s*window\.setTimeout\(\(\)\s*=>\s*sendCollabFullPayload\(\{\s*replace:\s*true\s*}\),\s*0\);/s,
  "loaded draft should use an explicit version-guarded session replacement"
);
assert.match(
  html,
  /function\s+resetAll\(\)\s*{[\s\S]*?window\.setTimeout\(\(\)\s*=>\s*sendCollabFullPayload\(\{\s*replace:\s*true\s*}\),\s*0\);[\s\S]*?}/,
  "clear all should use an explicit version-guarded session replacement"
);
assert.match(html, /if \(message\.type === "ack"\)/, "server acknowledgements should advance the client version");
assert.match(
  html,
  /function\s+applyActiveDatePickerValue\(\)\s*{[\s\S]*?const path = getControlCollabPath\(displayInput\);[\s\S]*?queueCollabFieldPatch\(path,\s*getCollabFieldValue\(displayInput\),\s*true\);[\s\S]*?}/,
  "date picker changes should sync through collab patches"
);
assert.ok(
  html.indexOf('class="collab-panel"') < html.indexOf('id="loadDraftBtn"'),
  "collab controls should sit before Load Draft Data"
);

const pagesFunctionPath = path.join(root, "functions", "api", "collab", "[sessionId].js");
const workerPath = path.join(root, "worker", "index.mjs");
const workerWranglerPath = path.join(root, "worker", "wrangler.toml");
const readmePath = path.join(root, "README.md");

for (const file of [pagesFunctionPath, workerPath, workerWranglerPath, readmePath]) {
  assert.equal(fs.existsSync(file), true, `${path.relative(root, file)} should exist`);
}

const pagesFunction = fs.readFileSync(pagesFunctionPath, "utf8");
assert.match(pagesFunction, /MOM_COLLAB_WORKER_URL/, "Pages Function should allow worker URL override");
assert.match(pagesFunction, /generate-mom-collab-worker\.alex-marcello08\.workers\.dev/, "Pages Function should default to production Worker");
assert.match(pagesFunction, /fetch\(new Request/, "Pages Function should proxy the WebSocket request");

const workerWrangler = fs.readFileSync(workerWranglerPath, "utf8");
assert.match(workerWrangler, /main\s*=\s*"index\.mjs"/);
assert.match(workerWrangler, /new_sqlite_classes\s*=\s*\["MomCollabSession"\]/);

(async () => {
  const workerModule = await import(pathToFileURL(workerPath).href);
  assert.equal(typeof workerModule.default.fetch, "function");
  assert.equal(typeof workerModule.MomCollabSession, "function");
  assert.equal(typeof workerModule.parseCollabSessionId, "function");
  assert.equal(typeof workerModule.shouldAcceptFullMessage, "function");
  assert.equal(workerModule.parseCollabSessionId(new Request("https://example.test/api/collab/abc-123")), "abc-123");

  assert.equal(
    workerModule.shouldAcceptFullMessage(null, 8, { type: "full", baseVersion: 0 }),
    true,
    "the first editor should be allowed to seed an empty session"
  );
  assert.equal(
    workerModule.shouldAcceptFullMessage({ version: "draft" }, 8, { type: "full", baseVersion: 8 }),
    false,
    "an implicit full snapshot must not replace an existing session"
  );
  assert.equal(
    workerModule.shouldAcceptFullMessage(
      { version: "draft" },
      8,
      { type: "full", replace: true, baseVersion: 7 }
    ),
    false,
    "a stale explicit replacement must not overwrite newer editor patches"
  );
  assert.equal(
    workerModule.shouldAcceptFullMessage(
      { version: "draft" },
      8,
      { type: "full", replace: true, baseVersion: 8 }
    ),
    true,
    "an explicit replacement based on the latest server version should be accepted"
  );

  const draft = { table3State: [{ activity: "old" }] };
  workerModule.setDraftPath(draft, "table3State/0/activity", "new");
  assert.equal(draft.table3State[0].activity, "new");

  console.log("collab integration tests passed");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
