const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const source = fs.readFileSync(
  path.join(__dirname, "..", "functions", "api", "collab", "[sessionId].js"),
  "utf8"
);

(async () => {
  const moduleUrl = `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;
  const pagesFunction = await import(moduleUrl);

  const productionContext = {
    env: {},
    request: new Request("https://generate-mom.pages.dev/api/collab/abc"),
  };
  assert.equal(
    pagesFunction.resolveCollabWorkerUrl(productionContext),
    "https://generate-mom-collab-worker.alex-marcello08.workers.dev",
    "the ordinary Cloudflare URL must keep routing to the unchanged legacy Worker"
  );

  const accessTestContext = {
    env: {},
    request: new Request("https://generate-mom-entra-test.pages.dev/api/collab/abc"),
  };
  assert.equal(
    pagesFunction.resolveCollabWorkerUrl(accessTestContext),
    "https://generate-mom-collab-worker-entra-test.alex-marcello08.workers.dev",
    "the isolated Access hostname must never fall back to production collaboration storage"
  );

  const overrideContext = {
    env: { MOM_COLLAB_WORKER_URL: "https://override.example.test" },
    request: accessTestContext.request,
  };
  assert.equal(
    pagesFunction.resolveCollabWorkerUrl(overrideContext),
    "https://override.example.test",
    "an explicit environment binding should remain the highest-priority route"
  );

  console.log("Pages collaboration routing tests passed");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
