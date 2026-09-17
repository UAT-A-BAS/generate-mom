const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { pathToFileURL } = require("node:url");
const { execFileSync } = require("node:child_process");

(async () => {
  const root = path.join(__dirname, "..");
  const artifact = path.join(root, "generate-mom-offline.html");
  assert.equal(fs.existsSync(artifact), true, "offline artifact must exist");
  const html = fs.readFileSync(artifact, "utf8");
  const builder = path.join(root, "tools/build-offline-html.mjs");
  const { buildOfflineHtml } = await import(pathToFileURL(builder).href);
  assert.equal(html, buildOfflineHtml().html, "artifact must match a fresh build");
  assert.equal(buildOfflineHtml().html, buildOfflineHtml().html, "build must be deterministic");
  execFileSync(process.execPath, [builder, "--check"], { cwd: root });
  assert.match(html, /<meta name="mom-offline-artifact" content="1">/);
  assert.match(html, /index\.html SHA-256: [a-f0-9]{64}/);
  assert.match(html, /MOM Generator/);
  const scripts = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)];
  assert.equal(scripts[0][1], "window.__MOM_OFFLINE__ = true;", "offline flag must execute first");
  assert.ok(scripts.length >= 2, "main application script must remain present");
  assert.match(html, /offline:\s*Boolean\(window\.__MOM_OFFLINE__\)/, "source must include its offline guard");
  const macro = html.match(/href="data:application\/octet-stream;base64,([A-Za-z0-9+/=]+)"/);
  assert.ok(macro, "Outlook macro must be embedded");
  assert.deepEqual(Buffer.from(macro[1], "base64"), fs.readFileSync(path.join(root, "ExportMOMToDraft.bas")));
  for (const forbidden of ["workers.dev", "pages.dev", "generate-mom.pages.dev", "wss://", "ws://", "/api/collab"]) {
    assert.equal(html.includes(forbidden), false, `artifact must not contain ${forbidden}`);
  }
  assert.doesNotMatch(html, /<script\b[^>]*\bsrc\s*=/i, "scripts must be inline");
  assert.doesNotMatch(html, /<link\b[^>]*\bhref\s*=\s*["']https?:/i, "no remote stylesheets");
  assert.doesNotMatch(html, /href="\.\/ExportMOMToDraft\.bas"/, "macro must not depend on a sibling file");
  console.log("Offline artifact tests passed: fresh, deterministic, embedded macro, no network endpoints");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
