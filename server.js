// Yerel geliştirme sunucusu:  npm start  →  http://localhost:4321
// Açılışta siteyi derler, site/ catalog/ library/ değiştikçe yeniden derler.
const http = require("http");
const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");

const PORT = Number(process.env.PORT) || 4321;
const ROOT = __dirname;
const DIST = path.join(ROOT, "dist");
const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".md": "text/markdown; charset=utf-8",
  ".py": "text/plain; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".zip": "application/zip",
};

// Her derleme ayrı bir Node sürecinde çalışır; böylece katalog alt modülleri önbellekte kalmaz.
let building = Promise.resolve();
const rebuild = () => (building = building.then(() => new Promise((resolve) => {
  const p = spawn(process.execPath, [path.join(ROOT, "scripts", "build.mjs")], { cwd: ROOT, stdio: "inherit" });
  p.on("exit", resolve);
  p.on("error", (e) => { console.error("✗ Derleme başlatılamadı: " + e.message); resolve(); });
})));

let timer;
for (const dir of ["site", "catalog", "library", "scripts"]) {
  fs.watch(path.join(ROOT, dir), { recursive: true }, (_, file) => {
    if (file && /__pycache__|cikti|\.pyc$/.test(file)) return;
    clearTimeout(timer);
    timer = setTimeout(rebuild, 200);
  });
}

rebuild().then(() => {
  http.createServer(async (req, res) => {
    await building;
    const urlPath = decodeURIComponent(req.url.split("?")[0]);
    const file = path.normalize(path.join(DIST, urlPath === "/" ? "index.html" : urlPath));
    if (!file.startsWith(DIST)) { res.writeHead(403); return res.end(); }
    fs.readFile(file, (err, data) => {
      if (err) { res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" }); return res.end("404"); }
      res.writeHead(200, { "Content-Type": TYPES[path.extname(file)] || "application/octet-stream", "Cache-Control": "no-cache" });
      res.end(data);
    });
  }).listen(PORT, () => console.log(`Workers / Workless → http://localhost:${PORT}`));
});
