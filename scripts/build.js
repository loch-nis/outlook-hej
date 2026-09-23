// Writes the hosted site to docs/ (GitHub Pages) and manifest.xml pointing at it.
// Usage: node scripts/build.js https://<user>.github.io/<repo>
const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const baseUrl = (process.argv[2] || process.env.BASE_URL || "").replace(/\/$/, "");
if (!/^https:\/\//.test(baseUrl)) {
  console.error("Usage: node scripts/build.js https://<host>/<path>  (must be https)");
  process.exit(1);
}
const { version } = require(path.join(root, "package.json"));
const fill = (text) => text.replaceAll("{{BASE_URL}}", baseUrl).replaceAll("{{VERSION}}", version);

const out = path.join(root, "docs");
fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(path.join(out, "assets"), { recursive: true });

fs.writeFileSync(path.join(out, "commands.html"), fill(fs.readFileSync(path.join(root, "src/commands.html"), "utf8")));
fs.copyFileSync(path.join(root, "src/launchevent.js"), path.join(out, "launchevent.js"));
for (const icon of fs.readdirSync(path.join(root, "assets"))) {
  fs.copyFileSync(path.join(root, "assets", icon), path.join(out, "assets", icon));
}
const manifest = fill(fs.readFileSync(path.join(__dirname, "manifest.template.xml"), "utf8"));
fs.writeFileSync(path.join(root, "manifest.xml"), manifest);
fs.writeFileSync(path.join(out, "manifest.xml"), manifest);
fs.writeFileSync(path.join(out, "index.html"), '<!DOCTYPE html><meta charset="UTF-8"><title>Hej-hilsen</title><p>Outlook add-in. <a href="manifest.xml">manifest.xml</a></p>\n');

console.log(`Built docs/ and manifest.xml for ${baseUrl} (v${version})`);
