// SEO sürümünü önizleme sunucusu: node scripts/sunucu-seo.mjs → http://localhost:4322
// Açılışta dist-seo/ yeniden üretilir. Klasör adreslerinde index.html sunulur (GitHub Pages ile aynı davranış).
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const OUT = path.join(ROOT, "dist-seo");
const PORT = Number(process.env.PORT) || 4322;
const TYPES = { ".html": "text/html; charset=utf-8", ".css": "text/css; charset=utf-8", ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8", ".xml": "application/xml; charset=utf-8", ".txt": "text/plain; charset=utf-8",
  ".md": "text/markdown; charset=utf-8", ".py": "text/plain; charset=utf-8", ".png": "image/png", ".svg": "image/svg+xml", ".zip": "application/zip" };

spawnSync(process.execPath, [path.join(ROOT, "scripts", "build-seo.mjs")], { cwd: ROOT, stdio: "inherit" });
http.createServer((req, res) => {
  let p = decodeURIComponent(req.url.split("?")[0]);
  if (p.endsWith("/")) p += "index.html";
  const f = path.normalize(path.join(OUT, p));
  if (!f.startsWith(OUT)) { res.writeHead(403); return res.end(); }
  fs.readFile(f, (err, data) => {
    if (err) {
      // "/gorev/x" (sondaki / olmadan) → "/gorev/x/"
      if (fs.existsSync(path.join(f, "index.html"))) { res.writeHead(301, { Location: p + "/" }); return res.end(); }
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" }); return res.end("404");
    }
    res.writeHead(200, { "Content-Type": TYPES[path.extname(f)] || "application/octet-stream", "Cache-Control": "no-cache" });
    res.end(data);
  });
}).listen(PORT, () => console.log(`Gerisi Bende (SEO sürümü) → http://localhost:${PORT}`));
