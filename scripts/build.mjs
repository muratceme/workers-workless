// Siteyi dist/ klasörüne derler:
//   - site/ dosyalarını kopyalar
//   - katalogdaki görev kimliklerinin benzersizliğini doğrular
//   - library/ altındaki her hazır görev için ZIP + SHA-256 üretir
//   - assets/catalog.js dosyasını (katalog + hazır görev bilgisi) yazar
// Kullanım: node scripts/build.mjs
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import { fileURLToPath, pathToFileURL } from "node:url";
import { makeZip } from "./zip.mjs";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DIST = path.join(ROOT, "dist");
const TYPES = ["code-blocks", "agents"];
const SKIP = /(^|[\\/])(__pycache__|\.pytest_cache|cikti|\.venv)([\\/]|$)|\.pyc$|(^|[\\/])\.env$/;

const TR = { ç: "c", ğ: "g", ı: "i", İ: "i", ö: "o", ş: "s", ü: "u", Ç: "c", Ğ: "g", Ö: "o", Ş: "s", Ü: "u" };
export const slug = (s) => s.replace(/[çğıİöşüÇĞÖŞÜ]/g, (c) => TR[c]).toLowerCase()
  .replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");

const walk = (dir, base = dir) => fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => {
  const full = path.join(dir, e.name);
  const rel = path.relative(base, full).split(path.sep).join("/");
  if (SKIP.test(rel)) return [];
  return e.isDirectory() ? walk(full, base) : [rel];
});

export async function build({ quiet = false } = {}) {
  const t0 = Date.now();
  // her seferinde taze içe aktar (geliştirme sunucusu yeniden derlerken)
  const cat = await import(pathToFileURL(path.join(ROOT, "catalog/catalog.mjs")).href + `?t=${t0}`);

  // 1) görev kayıt defteri
  //    Bir görev bir kez [ad, açıklama, saat] olarak tanımlanır; başka rollerde yalnızca adıyla
  //    ("Banka Mutabakatı") anılarak yeniden kullanılır. Aynı kimlik farklı içerikle iki kez
  //    tanımlanırsa derleme durur.
  const ids = new Map();      // id -> {name, desc, hours, where}
  for (const [deptId, d] of Object.entries(cat.DEPTS)) {
    for (const r of d.roles) for (const t of r.tasks) {
      if (typeof t === "string") continue;
      const [name, desc, hours] = t;
      const id = slug(name);
      const prev = ids.get(id);
      if (prev && (prev.desc !== desc || prev.hours !== hours)) {
        throw new Error(`"${name}" görevi iki farklı tanımla var (${prev.where} ve ${deptId}/${r.id}). ` +
          `Aynı görevse ikinci yerde yalnızca adını yazın; farklı görevse adını değiştirin.`);
      }
      if (!prev) ids.set(id, { name, desc, hours, where: `${deptId}/${r.id}` });
    }
  }
  const DEPTS = {};
  for (const [deptId, d] of Object.entries(cat.DEPTS)) {
    const roleIds = new Set();
    DEPTS[deptId] = {
      ...d,
      roles: d.roles.map((r) => {
        if (roleIds.has(r.id)) throw new Error(`${deptId}: "${r.id}" rolü iki kez tanımlanmış.`);
        roleIds.add(r.id);
        const seen = new Set();
        return {
          ...r,
          tasks: r.tasks.map((t) => {
            const id = slug(typeof t === "string" ? t : t[0]);
            const g = ids.get(id);
            if (!g) throw new Error(`${deptId}/${r.id}: "${t}" adında tanımlı bir görev yok.`);
            if (seen.has(id)) throw new Error(`${deptId}/${r.id}: "${g.name}" görevi aynı rolde iki kez var.`);
            seen.add(id);
            return [g.name, g.desc, g.hours];
          }),
        };
      }),
    };
  }
  for (const s of cat.SECTORS) for (const d of s.depts) {
    if (!DEPTS[d]) throw new Error(`${s.id} sektöründe tanımsız departman: "${d}"`);
  }

  // 2) dist/ temizle, site/ kopyala
  fs.rmSync(DIST, { recursive: true, force: true });
  fs.cpSync(path.join(ROOT, "site"), DIST, { recursive: true });

  // 3) hazır görevler
  const legal = ["LICENSE", "SORUMLULUK_REDDI.md"].map((f) => ({ name: f, data: fs.readFileSync(path.join(ROOT, f)) }));
  const READY = { "code-blocks": {}, agents: {} };
  for (const type of TYPES) {
    const libDir = path.join(ROOT, "library", type);
    if (!fs.existsSync(libDir)) continue;
    for (const id of fs.readdirSync(libDir).sort()) {
      const dir = path.join(libDir, id);
      const manifestPath = path.join(dir, "task.json");
      if (!fs.statSync(dir).isDirectory() || id.startsWith("_") || !fs.existsSync(manifestPath)) continue;
      const m = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
      if (m.id !== id) throw new Error(`${type}/${id}: task.json içindeki id "${m.id}" klasör adıyla aynı olmalı.`);
      if (!ids.has(id)) throw new Error(`${type}/${id}: katalogda bu kimlikte bir görev yok.`);

      const files = walk(dir);
      const zip = makeZip([
        ...files.map((f) => ({ name: `${id}/${f}`, data: fs.readFileSync(path.join(dir, f)) })),
        ...legal.map((l) => ({ name: `${id}/${l.name}`, data: l.data })),
      ]);
      const zipRel = `downloads/${type}/${id}.zip`;
      fs.mkdirSync(path.join(DIST, "downloads", type), { recursive: true });
      fs.writeFileSync(path.join(DIST, zipRel), zip);

      // sitede önizleme için ana dosya + README
      const fileDir = path.join(DIST, "files", type, id);
      fs.mkdirSync(fileDir, { recursive: true });
      for (const f of [m.ana_dosya, "README.md"]) fs.copyFileSync(path.join(dir, f), path.join(fileDir, f));

      const reqs = fs.existsSync(path.join(dir, "requirements.txt"))
        ? fs.readFileSync(path.join(dir, "requirements.txt"), "utf8").split(/\r?\n/).filter((l) => l && !l.startsWith("#"))
        : [];
      READY[type][id] = {
        version: m.surum,
        python: m.python,
        main: m.ana_dosya,
        run: m.calistir,
        inputs: m.girdiler,
        outputs: m.ciktilar,
        deps: reqs,
        zip: zipRel,
        size: zip.length,
        sha256: crypto.createHash("sha256").update(zip).digest("hex"),
        files: files.length + legal.length,
      };
    }
  }

  // 4) katalog
  const payload = { CONFIG: cat.CONFIG, DEPTS, SECTORS: cat.SECTORS, SOON: cat.SOON, READY, BUILT: new Date().toISOString() };
  fs.writeFileSync(path.join(DIST, "assets", "catalog.js"), `window.WW = ${JSON.stringify(payload)};\n`);

  const n = Object.keys(READY["code-blocks"]).length + Object.keys(READY.agents).length;
  if (!quiet) console.log(`✓ dist/ hazır · ${ids.size} görev, ${n} indirilebilir paket · ${Date.now() - t0} ms`);
  return { tasks: ids.size, ready: n };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  build().catch((e) => { console.error("✗ " + e.message); process.exit(1); });
}
