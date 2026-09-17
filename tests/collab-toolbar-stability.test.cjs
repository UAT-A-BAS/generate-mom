const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");

// Regression guard for a reported bug: pressing Start Collab made the button jump away.
// Two things moved the control cluster: the share button was unhidden (adding width to the
// row) and the primary label grew from "Start Collab" to "Restart Collab". Measured on the
// live page, that produced a 45px horizontal and 21px vertical jump on desktop.

assert.doesNotMatch(
  html,
  /copyShareLinkBtn\.hidden/,
  "the share button must not be unhidden; that adds width and shifts the row"
);
assert.doesNotMatch(
  html,
  /startCollabBtn\.textContent\s*=/,
  "the primary label must not be swapped by rewriting text content, which resizes the button"
);

// Both labels live in the same grid cell, so the button is always as wide as the longest one.
assert.match(
  html,
  /\.collab-primary-action\s*{\s*display:\s*inline-grid;\s*place-items:\s*center;\s*}/,
  "the primary action must size itself from stacked labels"
);
assert.match(
  html,
  /\.collab-primary-action \.collab-action-label\s*{\s*grid-area:\s*1 \/ 1;\s*white-space:\s*nowrap;\s*}/,
  "both labels must occupy the same grid cell so neither can change the button width"
);
assert.match(
  html,
  /\.collab-primary-action \.collab-action-label-live\s*{\s*visibility:\s*hidden;\s*}/,
  "the idle state must reserve the width of the longer live label"
);
assert.match(
  html,
  /\.collab-primary-action\.is-live \.collab-action-label-idle\s*{\s*visibility:\s*hidden;\s*}/,
  "the live state must swap which label is visible without changing geometry"
);

// Both labels must be present in the markup, and the inactive one hidden with `visibility`
// rather than `display`, otherwise it would stop reserving width.
for (const label of ["Start Collab", "Restart Collab"]) {
  assert.ok(
    html.includes(`>${label}</span>`),
    `the primary action must contain the "${label}" label`
  );
}
assert.doesNotMatch(
  html,
  /\.collab-action-label-live\s*{\s*display:\s*none/,
  "hiding the live label with display:none would reintroduce the width shift"
);

// The share button ships disabled in the markup so it cannot flash as enabled before the
// first status update, and its state is driven by `disabled` rather than visibility.
assert.match(
  html,
  /<button id="copyShareLinkBtn" type="button" class="btn-ghost" disabled>Copy Share Link<\/button>/,
  "the share button must start disabled and always occupy its slot"
);
assert.match(
  html,
  /elements\.copyShareLinkBtn\.disabled = collabState\.offline \|\| !collabState\.active;/,
  "the share button must toggle disabled, not visibility"
);
assert.match(
  html,
  /elements\.startCollabBtn\.classList\.toggle\("is-live", collabState\.active\);/,
  "the primary button must swap labels through a state class"
);

// A disabled control must not look broken.
assert.match(
  html,
  /\.collab-panel \.btn-ghost:disabled\s*{[\s\S]*?opacity:\s*0\.45;[\s\S]*?}/,
  "the disabled share button needs a visible disabled style"
);

console.log("collab toolbar stability tests passed");
