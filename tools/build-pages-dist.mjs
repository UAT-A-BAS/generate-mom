import { cpSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

export const projectRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
export const distPath = resolve(projectRoot, "dist-pages");

// Only these artifacts are published. Tests, tooling and source notes stay out of the
// public site so the deployed bundle is exactly what a visitor needs.
const PUBLIC_FILES = [
  "index.html",
  "generate-mom-offline.html",
  "ExportMOMToDraft.bas",
  "Panduan menggunakan MOM Generator updated.docx",
];

export function buildPagesDist({ quiet = false } = {}) {
  rmSync(distPath, { recursive: true, force: true });
  mkdirSync(distPath, { recursive: true });

  for (const file of PUBLIC_FILES) {
    cpSync(resolve(projectRoot, file), resolve(distPath, file));
  }

  // Cloudflare Pages picks up `functions/` from the deploy root.
  cpSync(resolve(projectRoot, "functions"), resolve(distPath, "functions"), { recursive: true });

  // The collaboration session URL is generated, so every HTML response must be fresh;
  // otherwise a returning editor can load stale markup that talks an older protocol.
  writeFileSync(
    resolve(distPath, "_headers"),
    [
      "/*",
      "  X-Content-Type-Options: nosniff",
      "  Referrer-Policy: strict-origin-when-cross-origin",
      "",
      "/index.html",
      "  Cache-Control: no-store, must-revalidate",
      "",
      "/",
      "  Cache-Control: no-store, must-revalidate",
      "",
    ].join("\n")
  );

  const sizes = PUBLIC_FILES.map((file) => {
    const bytes = readFileSync(resolve(distPath, file)).length;
    return `${file} (${bytes} bytes)`;
  });
  if (!quiet) {
    console.log(`Built ${distPath}`);
    sizes.forEach((line) => console.log(`  ${line}`));
  }
  return { distPath, files: PUBLIC_FILES, sizes };
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    buildPagesDist();
  } catch (error) {
    console.error(`FAIL: ${error.message}`);
    process.exitCode = 1;
  }
}
