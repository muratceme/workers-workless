// Paylaşım görselini üretir: scripts/og_gorsel.html → site/assets/og.png (1200×630)
// Playwright gerekir:  npx playwright@1.56.1 ... ya da NODE_PATH ile mevcut kurulum.
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { createRequire } from "node:module";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const { chromium } = createRequire(import.meta.url)("playwright");
const b = await chromium.launch({ channel: "msedge", headless: true });
const p = await b.newPage({ viewport: { width: 1200, height: 630 } });
await p.goto(pathToFileURL(path.join(ROOT, "scripts", "og_gorsel.html")).href, { waitUntil: "networkidle" });
await p.evaluate(() => document.fonts.ready);
await p.screenshot({ path: path.join(ROOT, "site", "assets", "og.png") });
await b.close();
console.log("✓ site/assets/og.png");
