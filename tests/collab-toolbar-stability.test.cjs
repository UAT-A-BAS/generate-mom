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

// Copy Share Link only appears once a session exists, but it keeps its slot reserved. That
// is the only way to satisfy both requirements at once: hidden until Start Collab, yet
// nothing in the toolbar moves when it shows up. `visibility` keeps the box in the layout;
// `display: none` would collapse it and shove the pill row sideways.
assert.match(
  html,
  /<button\s+id="copyShareLinkBtn"[\s\S]*?class="btn-ghost collab-share-action"[\s\S]*?aria-hidden="true"[\s\S]*?tabindex="-1"[\s\S]*?disabled/,
  "the share button must ship hidden, untabbable and disabled"
);
assert.match(
  html,
  /\.collab-share-action\s*{\s*visibility:\s*hidden;\s*pointer-events:\s*none;\s*}/,
  "the inactive share button must be invisible but keep its box"
);
assert.match(
  html,
  /\.collab-share-action\.is-visible\s*{\s*visibility:\s*visible;\s*pointer-events:\s*auto;\s*}/,
  "the active share button must become visible in the same slot"
);
assert.doesNotMatch(
  html,
  /copyShareLinkBtn\.hidden\s*=/,
  "display-based hiding would reintroduce the layout shift"
);
assert.match(
  html,
  /const canShare = !collabState\.offline && collabState\.active;\s*elements\.copyShareLinkBtn\.classList\.toggle\("is-visible", canShare\);/,
  "visibility must be driven by whether a session exists"
);
assert.match(
  html,
  /elements\.copyShareLinkBtn\.setAttribute\("aria-hidden", canShare \? "false" : "true"\);[\s\S]*?elements\.copyShareLinkBtn\.tabIndex = canShare \? 0 : -1;/,
  "the inactive share button must leave the tab order and the accessibility tree"
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

// The status pills change wording when a session starts ("Personal Draft" -> "Live",
// "Offline" -> "Connected", "Sync: -" -> "Sync: 12.34"). Because the row is right-aligned,
// those wording changes slid the whole cluster and could even change how many lines it
// wrapped onto. Each value reserves the width of its longest variant.
assert.match(
  html,
  /\.collab-stable-pill \.collab-pill-value\s*{\s*display:\s*inline-grid;\s*place-items:\s*center;\s*}/,
  "dynamic pills must stack their value and sizer in one grid cell"
);
assert.match(
  html,
  /\.collab-stable-pill \.collab-pill-value::after\s*{\s*content:\s*attr\(data-sizer\);\s*grid-area:\s*1 \/ 1;\s*visibility:\s*hidden;\s*white-space:\s*nowrap;\s*}/,
  "the hidden sizer must measure the longest wording without showing it"
);
assert.match(
  html,
  /\.collab-stable-pill \.collab-pill-text\s*{\s*grid-area:\s*1 \/ 1;\s*white-space:\s*nowrap;\s*}/,
  "the visible value must share the sizer's grid cell"
);

for (const [id, sizer] of [
  ["collabModeText", "Personal Draft"],
  ["collabConnectionText", "Disconnected"],
  ["collabUsersText", "Users: 00"],
  ["collabLastSyncedText", "Sync: 00.00"],
]) {
  assert.match(
    html,
    new RegExp(`id="${id}"[^>]*collab-stable-pill`),
    `${id} must be a stable pill`
  );
  assert.ok(
    html.includes(`data-sizer="${sizer}"`),
    `${id} must reserve the width of its longest wording ("${sizer}")`
  );
}

// Updating a pill must go through the value span. Writing to the pill directly would delete
// the sizer that lives in the same grid cell, and the shifting would come straight back.
assert.match(
  html,
  /function setCollabPillText\(pill, text\)\s*{[\s\S]*?querySelector\("\.collab-pill-text"\)/,
  "status updates must target the value span so the sizer survives"
);
for (const ref of [
  "elements.collabModeText.textContent",
  "elements.collabConnectionText.textContent",
  "elements.collabUsersText.textContent",
  "elements.collabLastSyncedText.textContent",
]) {
  assert.equal(
    html.includes(ref),
    false,
    `${ref} would overwrite the sizer and must not be used`
  );
}

// The reserved row has to fit beside the buttons on one line. The toolbar only leaves about
// 494px of room at 1440px, and the sync pill is the widest, so its label stays short and
// second-precision is dropped from the timestamp.
assert.equal(
  /second:\s*"2-digit"/.test(html),
  false,
  "the sync timestamp must not carry seconds, or the reserved row no longer fits on one line"
);

// `hidden` must beat the pill's own `display`, otherwise an empty conflict pill still
// renders its status dot on a second row.
assert.match(
  html,
  /\.collab-pill\[hidden\]\s*{\s*display:\s*none;\s*}/,
  "hidden pills must actually be removed from the layout"
);

console.log("collab toolbar stability tests passed");
