import assert from "node:assert/strict";
import { readdirSync } from "node:fs";
import { createRequire } from "node:module";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const require = createRequire(import.meta.url);
const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");

function loadPlaywright() {
  try { return require("playwright"); } catch (error) {
    if (error.code !== "MODULE_NOT_FOUND") throw error;
  }
  // Also support Playwright already installed by npx, without downloading anything.
  const cache = join(process.env.npm_config_cache || join(homedir(), ".npm"), "_npx");
  let entries = [];
  try { entries = readdirSync(cache).sort(); } catch (error) {
    if (error.code !== "ENOENT") throw error;
  }
  for (const entry of entries) {
    try { return require(join(cache, entry, "node_modules", "playwright")); } catch (error) {
      if (error.code !== "MODULE_NOT_FOUND") throw error;
    }
  }
  throw new Error("Playwright is unavailable. Install it with npm install --no-save playwright and npx playwright install chromium.");
}

const outbound = [];
const sockets = [];
const pageErrors = [];
let browser;
try {
  const { chromium } = loadPlaywright();
  browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ offline: true, serviceWorkers: "block" });
  context.on("request", (request) => {
    if (!/^(file:|data:|blob:|about:)/i.test(request.url())) {
      outbound.push(request.url());
      console.log(`OUTBOUND ATTEMPT: ${request.method()} ${request.url()}`);
    }
  });
  await context.route("**", (route) => {
    if (/^(file:|data:|blob:|about:)/i.test(route.request().url())) return route.continue();
    return route.abort("internetdisconnected");
  });
  await context.addInitScript(() => {
    window.__offlineWebSocketAttempts = [];
    const OriginalWebSocket = window.WebSocket;
    window.WebSocket = class extends OriginalWebSocket {
      constructor(...args) {
        window.__offlineWebSocketAttempts.push(String(args[0]));
        throw new Error("WebSocket construction blocked by offline verification");
      }
    };
  });
  for (const session of [false, true]) {
    const page = await context.newPage();
    page.on("pageerror", (error) => pageErrors.push(error.message));
    const url = pathToFileURL(join(root, "generate-mom-offline.html"));
    if (session) url.searchParams.set("session", "offline-runtime-probe");
    await page.goto(url.href, { waitUntil: "load" });
    await page.locator("#table1Projects textarea, #table1Projects input").first().waitFor({ state: "visible" });
    assert.equal(await page.locator("#generateBtn").isVisible(), true, "form actions must render");
    assert.equal(await page.evaluate(() => window.__MOM_OFFLINE__), true);
    assert.equal((await page.locator("#collabModeText").textContent()).trim(), "Personal Draft");
    assert.equal((await page.locator("#collabConnectionText").textContent()).trim(), "Offline");
    assert.doesNotMatch(await page.locator("#collabStatus").innerText(), /\bLive\b/);
    // Exercise the control even when disabled/hidden to check the handler's guard.
    await page.locator("#startCollabBtn").dispatchEvent("click");
    await page.waitForTimeout(300);
    assert.equal((await page.locator("#collabModeText").textContent()).trim(), "Personal Draft");
    sockets.push(...await page.evaluate(() => window.__offlineWebSocketAttempts));
    await page.close();
  }
  for (const socket of sockets) console.log(`WEBSOCKET ATTEMPT: ${socket}`);
  assert.deepEqual(outbound, [], "no outbound requests may be attempted");
  assert.deepEqual(sockets, [], "no WebSocket may be constructed");
  assert.deepEqual(pageErrors, [], "application must boot without runtime errors");
  console.log(`PASS offline runtime: ${outbound.length} outbound requests; ${sockets.length} WebSocket constructions; form rendered; Personal Draft / Offline; plain and session URLs verified.`);
} catch (error) {
  console.error(`FAIL offline runtime: ${outbound.length} outbound requests; ${sockets.length} WebSocket constructions; ${error.message}`);
  process.exitCode = 1;
} finally {
  await browser?.close();
}
