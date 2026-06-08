/*
 * Generate all required icon sizes from the master artwork (icon-source.png)
 * using macOS's built-in `sips` (no third-party deps).
 *
 *   node tools/make-icons.js
 *
 * Sizes: 16/32/80 = ribbon buttons, 64 = AppSource IconUrl,
 * 128 = AppSource HighResolutionIconUrl, 300 = Partner Center listing logo.
 *
 * Replace icon-source.png with new artwork (square PNG) and re-run to refresh
 * every size at once.
 */

const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const root = path.join(__dirname, "..");
const source = path.join(root, "icon-source.png");
const outDir = path.join(root, "assets");
const SIZES = [16, 32, 64, 80, 128, 300];

if (!fs.existsSync(source)) {
  console.error("Missing icon-source.png in the project root. Add a square PNG and re-run.");
  process.exit(1);
}

fs.mkdirSync(outDir, { recursive: true });
SIZES.forEach(function (size) {
  const out = path.join(outDir, "icon-" + size + ".png");
  execFileSync("sips", ["-s", "format", "png", "-z", String(size), String(size), source, "--out", out], {
    stdio: "ignore",
  });
  console.log("wrote " + out);
});
