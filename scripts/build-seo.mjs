// SEO sürümü (deneme): normal derlemenin üstüne Google'ın okuyabileceği statik sayfalar ekler → dist-seo/
//   /gorev/<görev>/      her görev için ayrı sayfa (başlık, açıklama, girdi/çıktı, örnek çıktı tablosu, kaynaklar)
//   /meslek/<rol>/       her rolün görev listesi
//   /sektor/<sektör>/    sektörün departman ve rolleri
//   /gorevler/           bütün görevlerin dizini
//   sitemap.xml, robots.txt, yapılandırılmış veri (JSON-LD)
// Uygulama (hash yönlendirmeli tek sayfa) aynen kalır; statik sayfalar ona bağlantı verir.
// Kullanım: node scripts/build-seo.mjs   ·   önizleme: node scripts/sunucu-seo.mjs (port 4322)
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { build, slug } from "./build.mjs";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DIST = path.join(ROOT, "dist");
const OUT = path.join(ROOT, "dist-seo");
const SITE = "https://muratceme.github.io/workers-workless/";  // alan adı alınınca değiştirin

await build({ quiet: true });
fs.rmSync(OUT, { recursive: true, force: true });
fs.cpSync(DIST, OUT, { recursive: true });

const raw = fs.readFileSync(path.join(DIST, "assets", "catalog.js"), "utf8");
const W = JSON.parse(raw.slice(raw.indexOf("=") + 1).trim().replace(/;$/, ""));
const ONIZ = fs.existsSync(path.join(ROOT, "data", "onizlemeler.json"))
  ? JSON.parse(fs.readFileSync(path.join(ROOT, "data", "onizlemeler.json"), "utf8")) : {};
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
const kisalt = (s, n) => (s.length <= n ? s : s.slice(0, n - 1).replace(/\s+\S*$/, "") + "…");

// ---- veri: görev → kullanıldığı roller ----
const GOREV = new Map();  // id -> {id,name,desc,hours, yerler:[{s,dId,d,r}]}
const ROL = new Map();    // rolSlug -> {s list, dId, d, r}
for (const s of W.SECTORS) for (const dId of s.depts) {
  const d = W.DEPTS[dId];
  for (const r of d.roles) {
    const rs = `${dId}-${r.id}`;
    if (!ROL.has(rs)) ROL.set(rs, { slug: rs, dId, d, r, sektorler: [] });
    ROL.get(rs).sektorler.push(s);
    for (const [name, desc, hours] of r.tasks) {
      const id = slug(name);
      if (!GOREV.has(id)) GOREV.set(id, { id, name, desc, hours, yerler: [] });
      GOREV.get(id).yerler.push({ s, dId, d, r, rs });
    }
  }
}
const hazir = (t) => ({ code: W.READY["code-blocks"][t.id], agents: W.READY.agents[t.id] });
const uygHref = (m, y, t) => `#/${m}/${y.s.id}/${y.dId}/${y.r.id}/${t.id}`;

// ---- sayfa iskeleti ----
const sayfa = ({ kok, baslik, aciklama, yol, govde, jsonld = [], kirintilar = [] }) => {
  const url = SITE + yol;
  const krn = kirintilar.length ? `<nav class="kirinti" aria-label="Konum">${[["Ana sayfa", kok], ...kirintilar]
    .map(([a, h], i, arr) => (i < arr.length - 1 ? `<a href="${h}">${esc(a)}</a>` : `<span>${esc(a)}</span>`)).join(" › ")}</nav>` : "";
  const ldKirinti = kirintilar.length ? [{
    "@context": "https://schema.org", "@type": "BreadcrumbList",
    itemListElement: [["Ana sayfa", ""], ...kirintilar.map(([a, h]) => [a, h])].map(([a], i) => ({ "@type": "ListItem", position: i + 1, name: a })),
  }] : [];
  return `<!doctype html>
<html lang="tr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(baslik)}</title>
<meta name="description" content="${esc(aciklama)}">
<link rel="canonical" href="${url}">
<meta property="og:type" content="website"><meta property="og:locale" content="tr_TR"><meta property="og:site_name" content="Gerisi Bende">
<meta property="og:title" content="${esc(baslik)}"><meta property="og:description" content="${esc(aciklama)}"><meta property="og:url" content="${url}">
<meta property="og:image" content="${SITE}assets/og.png"><meta name="twitter:card" content="summary_large_image">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='8' fill='%23050505'/%3E%3Cpath d='M9 16.5l4.5 4.5L23 11' stroke='%233DDC97' stroke-width='3' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E">
<link rel="stylesheet" href="${kok}assets/seo.css">
${[...ldKirinti, ...jsonld].map((j) => `<script type="application/ld+json">${JSON.stringify(j)}</script>`).join("\n")}
</head>
<body>
<header class="ust"><a class="logo" href="${kok}"><span class="mark"></span>Gerisi <b>Bende</b></a>
<nav><a href="${kok}gorevler/">Tüm görevler</a><a href="${kok}#/code">Kod Blokları</a><a href="${kok}#/agents">AI Agents</a></nav></header>
<main>${krn}${govde}</main>
<footer class="alt"><p>Gerisi Bende · Açık kaynak (MIT) · Ücretsiz · Türkçe · <a href="${kok}#/sss">SSS</a> · <a href="${kok}#/gizlilik">Gizlilik</a> · <a href="${kok}#/istek">Görev talep et</a></p></footer>
</body></html>`;
};
const yaz = (rel, html) => { const f = path.join(OUT, rel, "index.html"); fs.mkdirSync(path.dirname(f), { recursive: true }); fs.writeFileSync(f, html); };

const onizHtml = (o) => o.sayfalar.slice(0, 2).map((s) => `
  <figure class="oniz"><figcaption>${esc(o.dosya)} · ${esc(s.ad)} sayfası${s.toplam_satir > s.satirlar.length ? ` (ilk ${s.satirlar.length} / ${s.toplam_satir} satır)` : ""}</figcaption>
  <div class="tablo"><table><thead><tr>${s.satirlar[0].map((h) => `<th>${esc(h)}</th>`).join("")}</tr></thead>
  <tbody>${s.satirlar.slice(1).map((r) => `<tr>${r.map((c) => `<td>${esc(c)}</td>`).join("")}</tr>`).join("")}</tbody></table></div></figure>`).join("");
const guncHtml = (g) => !g ? "" : `<div class="gunc"><b>Son kontrol: ${esc(g.tarih || "-")}</b>${g.parametre ? ` · ${esc(g.parametre)}` : ""}
  <p>Mevzuata ya da resmî oranlara dayanır; kullanmadan önce kaynakları güncel resmî metinle karşılaştırın.</p>
  <ul>${g.kaynaklar.slice(0, 8).map((k) => `<li>${k.url ? `<a href="${esc(k.url)}" rel="noopener nofollow" target="_blank">${esc(k.ad)}</a>` : esc(k.ad)}</li>`).join("")}</ul></div>`;

// ---- görev sayfaları ----
const urls = [""];
for (const t of GOREV.values()) {
  const kok = "../../", h = hazir(t), y0 = t.yerler[0];
  const turler = [h.code && "kod bloğu", h.agents && "AI agent"].filter(Boolean);
  const baslik = `${t.name}${h.code ? " · Ücretsiz Excel / Python otomasyonu" : h.agents ? " · Ücretsiz AI agent" : ""} | Gerisi Bende`;
  const roller = [...new Map(t.yerler.map((y) => [y.rs, y])).values()];
  const paket = (m, p, ad) => !p ? "" : `
    <section class="paket ${m}"><h2>${ad}</h2>
    <div class="ikili"><div><h3>Girdi</h3><ul>${p.inputs.map((x) => `<li>${esc(x)}</li>`).join("")}</ul></div>
    <div><h3>Çıktı</h3><ul>${p.outputs.map((x) => `<li>${esc(x)}</li>`).join("")}</ul></div></div>
    ${m === "code" && ONIZ[t.id] && !ONIZ[t.id].hata ? `<h3>Örnek çıktı</h3>${onizHtml(ONIZ[t.id])}` : ""}
    ${guncHtml(p.guncellik)}
    <p><a class="dugme" href="${kok}${uygHref(m, y0, t)}">İncele ve indir →</a>
    <span class="kucuk">Sürüm ${esc(p.version)} · ${m === "code" ? "internet ve API anahtarı gerekmez" : "kendi API anahtarınızla çalışır"} · çift tıkla başlatıcı dahil</span></p></section>`;
  const govde = `<article>
    <p class="ust-etiket">${turler.length ? `Hazır ${turler.join(" + ")}` : "Yapım aşamasında"} · haftada ~${t.hours} saat kazanım</p>
    <h1>${esc(t.name)}</h1>
    <p class="giris">${esc(t.desc)}</p>
    ${paket("code", h.code, "Kod bloğu: kural belliyse")}
    ${paket("agents", h.agents, "AI agent: muhakeme gerekiyorsa")}
    ${!turler.length ? `<p>Bu görev henüz yazılmadı. <a href="${kok}#/istek">Talep ederek</a> sırasını öne alabilirsiniz.</p>` : ""}
    <section><h2>Bu işi yapan meslekler</h2><ul class="liste">${roller.map((y) => `<li><a href="${kok}meslek/${y.rs}/">${esc(y.r.name)}</a> <span>${esc(y.d.name)}${y.d.shared ? "" : ` · ${esc(y.s.name)}`}</span></li>`).join("")}</ul></section>
  </article>`;
  const jsonld = turler.length ? [{
    "@context": "https://schema.org", "@type": "SoftwareApplication", name: t.name, description: t.desc,
    applicationCategory: "BusinessApplication", operatingSystem: "Windows, macOS, Linux", inLanguage: "tr",
    isAccessibleForFree: true, license: "https://opensource.org/licenses/MIT", url: SITE + `gorev/${t.id}/`,
    offers: { "@type": "Offer", price: "0", priceCurrency: "TRY" },
  }] : [];
  yaz(`gorev/${t.id}`, sayfa({ kok, baslik, aciklama: kisalt(`${t.desc} Ücretsiz, açık kaynak, Türkçe.`, 158), yol: `gorev/${t.id}/`, govde, jsonld,
    kirintilar: [[y0.d.shared ? "Tüm sektörler" : y0.s.name, `${kok}sektor/${y0.s.id}/`], [y0.r.name, `${kok}meslek/${y0.rs}/`], [t.name, ""]] }));
  urls.push(`gorev/${t.id}/`);
}

// ---- meslek sayfaları ----
for (const R of ROL.values()) {
  const kok = "../../", saat = R.r.tasks.reduce((a, x) => a + x[2], 0);
  const sek = R.d.shared ? "Tüm sektörler" : [...new Set(R.sektorler.map((s) => s.name))].join(", ");
  const govde = `<article><p class="ust-etiket">${esc(R.d.name)} · ${esc(sek)}</p>
    <h1>${esc(R.r.name)} için hazır otomasyonlar</h1>
    <p class="giris">${esc(R.r.name)} rolünün tekrar eden ${R.r.tasks.length} işi için ücretsiz kod blokları ve AI agent'lar. Haftada yaklaşık ${saat} saat geri kazanım.</p>
    <ul class="kartlar">${R.r.tasks.map(([name, desc, hours]) => { const id = slug(name); const h = hazir({ id });
      return `<li><a href="${kok}gorev/${id}/"><b>${esc(name)}</b><span>${esc(kisalt(desc, 140))}</span><em>${h.code || h.agents ? "Hazır" : "Yakında"} · ~${hours} sa/hafta</em></a></li>`; }).join("")}</ul></article>`;
  yaz(`meslek/${R.slug}`, sayfa({ kok, baslik: `${R.r.name}: rutin işler için hazır otomasyonlar | Gerisi Bende`,
    aciklama: kisalt(`${R.r.name} için ${R.r.tasks.length} hazır görev: ${R.r.tasks.slice(0, 4).map((x) => x[0]).join(", ")}. Ücretsiz, Türkçe.`, 158),
    yol: `meslek/${R.slug}/`, govde, kirintilar: [[R.d.shared ? "Tüm sektörler" : R.sektorler[0].name, `${kok}sektor/${R.sektorler[0].id}/`], [R.r.name, ""]] }));
  urls.push(`meslek/${R.slug}/`);
}

// ---- sektör sayfaları ----
for (const s of W.SECTORS) {
  const kok = "../../";
  const govde = `<article><p class="ust-etiket">Sektör</p><h1>${esc(s.name)}</h1><p class="giris">${esc(s.tagline || "")}</p>
    ${s.depts.map((dId) => { const d = W.DEPTS[dId]; return `<section><h2>${esc(d.name)}</h2><ul class="liste">${d.roles.map((r) =>
      `<li><a href="${kok}meslek/${dId}-${r.id}/">${esc(r.name)}</a> <span>${r.tasks.length} görev</span></li>`).join("")}</ul></section>`; }).join("")}</article>`;
  yaz(`sektor/${s.id}`, sayfa({ kok, baslik: `${s.name} sektörü için hazır otomasyonlar | Gerisi Bende`,
    aciklama: kisalt(`${s.name}: ${s.tagline || ""} Her rol için ücretsiz kod blokları ve AI agent'lar.`, 158), yol: `sektor/${s.id}/`, govde,
    kirintilar: [[s.name, ""]] }));
  urls.push(`sektor/${s.id}/`);
}

// ---- tüm görevler dizini ----
{
  const kok = "../";
  const govde = `<article><h1>Tüm görevler</h1><p class="giris">${GOREV.size} görev, ${W.SECTORS.length} sektör. Her biri ücretsiz ve açık kaynak.</p>
    ${W.SECTORS.map((s) => `<section><h2><a href="${kok}sektor/${s.id}/">${esc(s.name)}</a></h2><ul class="sutunlu">${[...new Map(
      s.depts.flatMap((dId) => W.DEPTS[dId].roles.flatMap((r) => r.tasks.map((x) => [slug(x[0]), x[0]])))).entries()]
      .map(([id, ad]) => `<li><a href="${kok}gorev/${id}/">${esc(ad)}</a></li>`).join("")}</ul></section>`).join("")}</article>`;
  yaz("gorevler", sayfa({ kok, baslik: "Tüm görevler: hazır kod blokları ve AI agent'lar | Gerisi Bende",
    aciklama: `${GOREV.size} rutin iş için ücretsiz, Türkçe, açık kaynak otomasyon. Sektör ve mesleğe göre dizin.`, yol: "gorevler/", govde, kirintilar: [["Tüm görevler", ""]] }));
  urls.push("gorevler/");
}

// ---- ana sayfaya taranabilir bağlantılar ----
let ana = fs.readFileSync(path.join(OUT, "index.html"), "utf8");
ana = ana.replace("</main>", `  <nav class="seo-dizin" aria-label="Sektör dizini"><p>Sektöre göre tüm görevler:</p>
    ${W.SECTORS.map((s) => `<a href="sektor/${s.id}/">${esc(s.name)}</a>`).join("")}<a href="gorevler/">Tüm görevler →</a></nav>
  </main>`);
ana = ana.replace("</head>", `<style>.seo-dizin{max-width:1200px;margin:40px auto 0;padding:0 22px;display:flex;flex-wrap:wrap;gap:8px 16px;font-size:13px;color:#8a948f}.seo-dizin p{width:100%;margin:0}.seo-dizin a{color:#8a948f}</style>\n</head>`);
fs.writeFileSync(path.join(OUT, "index.html"), ana);

// ---- sitemap, robots, stil ----
const bugun = new Date().toISOString().slice(0, 10);
fs.writeFileSync(path.join(OUT, "sitemap.xml"), `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
${urls.map((u) => `  <url><loc>${SITE}${u}</loc><lastmod>${bugun}</lastmod></url>`).join("\n")}
</urlset>
`);
fs.writeFileSync(path.join(OUT, "robots.txt"), `User-agent: *\nAllow: /\nSitemap: ${SITE}sitemap.xml\n`);
fs.copyFileSync(path.join(ROOT, "scripts", "seo.css"), path.join(OUT, "assets", "seo.css"));
console.log(`✓ dist-seo/ hazır · ${GOREV.size} görev + ${ROL.size} meslek + ${W.SECTORS.length} sektör sayfası · sitemap ${urls.length} adres`);
