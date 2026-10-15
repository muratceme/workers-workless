/* =========================================================================
   Tasarım + tema seçici (GEÇİCİ). Karar verilince bu dosya, assets/themes.css,
   assets/layouts.css ve index.html'deki "GEÇİCİ" satırları silinir.
   Tasarım (düzen) = sayfanın yapısı · Tema = renk ve yazı tipi. Birlikte seçilebilir.
   ========================================================================= */
(() => {
  const GF = "https://fonts.googleapis.com/css2?display=swap&family=";
  const LAYOUTS = [
    { id: "", ad: "Orijinal", not: "Şu anki yapı", ic: '<rect x="3" y="3" width="18" height="4" rx="1"/><rect x="6" y="9" width="12" height="5" rx="1"/><rect x="3" y="16" width="5" height="5" rx="1"/><rect x="9.5" y="16" width="5" height="5" rx="1"/><rect x="16" y="16" width="5" height="5" rx="1"/>' },
    { id: "yan", ad: "Yan Panel", not: "Solda sabit menü ve sektör dizini", ic: '<rect x="3" y="3" width="6" height="18" rx="1"/><rect x="11" y="3" width="10" height="8" rx="1"/><rect x="11" y="13" width="4.5" height="8" rx="1"/><rect x="16.5" y="13" width="4.5" height="8" rx="1"/>' },
    { id: "bolunmus", ad: "Bölünmüş", not: "Dev başlık, yapışkan sol sütunlar", ic: '<rect x="3" y="3" width="10" height="9" rx="1"/><rect x="15" y="3" width="6" height="4" rx="1"/><rect x="15" y="9" width="6" height="3" rx="1"/><rect x="3" y="14" width="6" height="7" rx="1"/><rect x="11" y="14" width="10" height="7" rx="1"/>' },
    { id: "bento", ad: "Bento", not: "Her şey bir karo ızgarasında", ic: '<rect x="3" y="3" width="11" height="10" rx="1.5"/><rect x="16" y="3" width="5" height="4.5" rx="1.5"/><rect x="16" y="9" width="5" height="4" rx="1.5"/><rect x="3" y="15" width="6" height="6" rx="1.5"/><rect x="11" y="15" width="10" height="6" rx="1.5"/>' },
    { id: "dizin", ad: "Dizin", not: "Kart yerine numaralı satırlar", ic: '<path d="M3 5h18M3 10h18M3 15h18M3 20h18"/><path d="M6 3v19" opacity=".5"/>' },
    { id: "dok", ad: "Arama + Dok", not: "Arama ile açılış, altta yüzen menü", ic: '<rect x="4" y="5" width="16" height="4" rx="2"/><path d="M7 13h10" opacity=".6"/><rect x="7" y="17" width="10" height="4" rx="2"/>' },
  ];
  const THEMES = [
    { id: "", ad: "Orijinal", not: "Şu anki renkler", sw: ["#fbfbfd", "#0071e3", "#7c3aed"] },
    { id: "editoryal", ad: "Editoryal", not: "Sıcak kâğıt, serif başlık", scheme: "light", sw: ["#F7F6F3", "#2F3437", "#1F6C9F"],
      font: "Instrument+Serif:ital@0;1&family=Geist:wght@400;500;600&family=Geist+Mono:wght@400;500" },
    { id: "isvicre", ad: "İsviçre Baskı", not: "Kalın tipografi, tek kırmızı", scheme: "light", sw: ["#F4F4F0", "#111111", "#E61919"],
      font: "Archivo+Black&family=Archivo:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600" },
    { id: "terminal", ad: "Terminal", not: "Koyu, monospace, tarama çizgisi", scheme: "dark", sw: ["#0A0A0A", "#EAEAEA", "#FF2A2A"],
      font: "Space+Mono:wght@400;700&family=JetBrains+Mono:wght@400;500;700" },
    { id: "cam", ad: "Cam", not: "Derin siyah, ışık küreleri", scheme: "dark", sw: ["#050505", "#3DDC97", "#A78BFA"],
      font: "Outfit:wght@300;400;500;600;700&family=Geist+Mono:wght@400;500" },
    { id: "yumusak", ad: "Yumuşak Yapı", not: "Gümüş zemin, yüzen kartlar", scheme: "light", sw: ["#F2F3F5", "#0E1116", "#FF5B1F"],
      font: "Bricolage+Grotesque:opsz,wght@12..96,500..800&family=Plus+Jakarta+Sans:wght@400;500;600;700" },
  ];
  const root = document.documentElement;
  const $ = (s, el = document) => el.querySelector(s);
  const store = {
    get: (k) => { try { return localStorage.getItem(k) || ""; } catch { return ""; } },
    set: (k, v) => { try { v ? localStorage.setItem(k, v) : localStorage.removeItem(k); } catch {} },
  };
  const loaded = new Set();
  const loadFont = (t) => {
    if (!t.font || loaded.has(t.id)) return;
    loaded.add(t.id);
    const l = document.createElement("link");
    l.rel = "stylesheet";
    l.href = GF + t.font;
    document.head.appendChild(l);
  };
  const themeBtn = $("#themeToggle");
  const userTheme = () => {
    try { const t = localStorage.getItem("ww-theme"); if (t) return t; } catch {}
    return matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  };

  /* ------------------------------ düzenlere özel öğeler ------------------------------ */
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const curMode = () => (/^#\/agents/.test(location.hash) ? "agents" : "code");

  // Yan Panel: sektör dizini
  const side = document.createElement("div");
  side.className = "lx-side";
  side.innerHTML = `<p>Sektörler</p>${(window.WW ? window.WW.SECTORS : []).map((s) =>
    `<a href="#/code/${esc(s.id)}" data-sec="${esc(s.id)}">${esc(s.name)}<span>${s.depts.length}</span></a>`).join("")}`;
  $(".nav-inner").appendChild(side);
  side.addEventListener("click", (e) => {
    const a = e.target.closest("a[data-sec]");
    if (!a) return;
    e.preventDefault();
    location.hash = `#/${curMode()}/${a.dataset.sec}`;
  });
  const markSide = () => {
    const sec = location.hash.replace(/^#\/?/, "").split("/")[1];
    side.querySelectorAll("a").forEach((a) => a.classList.toggle("on", a.dataset.sec === sec));
  };
  window.addEventListener("hashchange", markSide);
  markSide();

  // Arama + Dok: büyük arama kutusu ve popüler aramalar
  const openSearch = (q) => {
    $("#openPalette").click();
    if (!q) return;
    const inp = $("#paletteInput");
    inp.value = q;
    inp.dispatchEvent(new Event("input"));
  };
  const title = $(".hero-title");
  title.insertAdjacentHTML("afterend", `
    <button class="lx-search" type="button" aria-label="Görev ara">
      <svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true"><circle cx="11" cy="11" r="7" fill="none" stroke="currentColor" stroke-width="2"/><path d="M20 20l-3.5-3.5" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      <span class="lx-q">Hangi işi bitirmek istiyorsunuz? Görev, rol ya da departman…</span>
      <kbd>Ctrl K</kbd>
      <span class="lx-go">Ara</span>
    </button>
    <div class="lx-pop"><span>Sık arananlar:</span>${["mutabakat", "bordro", "CV", "hasar", "KVKK", "fatura", "stok"].map((w) =>
      `<button type="button" data-q="${w}">${w}</button>`).join("")}</div>`);
  $(".lx-search").addEventListener("click", () => openSearch(""));
  $(".lx-pop").addEventListener("click", (e) => { const b = e.target.closest("[data-q]"); if (b) openSearch(b.dataset.q); });

  /* ------------------------------ seçici ------------------------------ */
  const item = (kind, x, i, swatch) => `
    <button class="ds-item" type="button" role="menuitemradio" aria-checked="false" data-kind="${kind}" data-id="${x.id}">
      ${swatch}
      <span class="ds-txt"><b>${x.ad}</b><small>${x.not}</small></span>
      <span class="ds-n">${i === 0 ? "—" : i}</span>
    </button>`;
  const wrap = document.createElement("div");
  wrap.className = "ds";
  wrap.innerHTML = `
    <button class="icon-btn ds-btn" type="button" aria-haspopup="true" aria-expanded="false" aria-controls="dsMenu" title="Tasarım ve tema seç">
      <svg viewBox="0 0 24 24" width="17" height="17" aria-hidden="true"><rect x="3.5" y="3.5" width="7" height="7" rx="1.5" fill="none" stroke="currentColor" stroke-width="2"/><rect x="13.5" y="3.5" width="7" height="7" rx="1.5" fill="currentColor"/><rect x="3.5" y="13.5" width="7" height="7" rx="1.5" fill="currentColor"/><rect x="13.5" y="13.5" width="7" height="7" rx="1.5" fill="none" stroke="currentColor" stroke-width="2"/></svg>
      <span class="ds-label">Tasarım</span>
    </button>
    <div class="ds-menu" id="dsMenu" role="menu" aria-label="Tasarım ve tema" hidden>
      <div class="ds-cols">
        <div class="ds-col">
          <p class="ds-head">Tasarım · sayfa yapısı</p>
          ${LAYOUTS.map((l, i) => item("layout", l, i, `<span class="ds-ic" aria-hidden="true"><svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="1.6">${l.ic}</svg></span>`)).join("")}
        </div>
        <div class="ds-col">
          <p class="ds-head">Tema · renk ve yazı</p>
          ${THEMES.map((t, i) => item("theme", t, i, `<span class="ds-sw" aria-hidden="true">${t.sw.map((c) => `<i style="background:${c}"></i>`).join("")}</span>`)).join("")}
        </div>
      </div>
      <p class="ds-foot">Geçici seçici · ikisi birlikte kullanılabilir · bağlantı: <code>?duzen=…&amp;tasarim=…</code></p>
    </div>`;
  $(".nav-actions").insertBefore(wrap, themeBtn || null);
  const btn = $(".ds-btn", wrap);
  const menu = $(".ds-menu", wrap);

  function sync() {
    menu.querySelectorAll(".ds-item").forEach((b) => {
      const cur = b.dataset.kind === "layout" ? root.dataset.layout || "" : root.dataset.design || "";
      b.setAttribute("aria-checked", String(b.dataset.id === cur));
    });
  }
  function applyTheme(id) {
    const t = THEMES.find((x) => x.id === id) || THEMES[0];
    loadFont(t);
    if (t.id) root.dataset.design = t.id; else delete root.dataset.design;
    // Temaların kendi açık/koyu düzeni var; açık/koyu düğmesi yalnızca orijinal temada çalışır.
    root.dataset.theme = t.scheme || userTheme();
    if (themeBtn) themeBtn.hidden = Boolean(t.scheme);
    store.set("ww-design", t.id);
    sync();
  }
  function applyLayout(id) {
    const l = LAYOUTS.find((x) => x.id === id) || LAYOUTS[0];
    if (l.id) root.dataset.layout = l.id; else delete root.dataset.layout;
    store.set("ww-layout", l.id);
    // Yeni yapıda görünür olan bölümlerin giriş animasyonu ve sayaçları tetiklensin
    document.querySelectorAll(".reveal").forEach((el) => {
      const r = el.getBoundingClientRect();
      if (r.top < innerHeight && r.bottom > 0) el.classList.add("in");
    });
    sync();
  }

  const open = (v) => {
    menu.hidden = !v;
    btn.setAttribute("aria-expanded", String(v));
    if (v) (menu.querySelector('[aria-checked="true"]') || menu.querySelector(".ds-item")).focus();
  };
  btn.addEventListener("click", () => open(menu.hidden));
  menu.addEventListener("click", (e) => {
    const it = e.target.closest(".ds-item");
    if (!it) return;
    (it.dataset.kind === "layout" ? applyLayout : applyTheme)(it.dataset.id);
  });
  document.addEventListener("click", (e) => { if (!wrap.contains(e.target)) open(false); });
  document.addEventListener("keydown", (e) => {
    if (menu.hidden) return;
    const items = [...menu.querySelectorAll(".ds-item")];
    const i = items.indexOf(document.activeElement);
    if (e.key === "Escape") { open(false); btn.focus(); }
    else if (e.key === "ArrowDown") { e.preventDefault(); items[(i + 1) % items.length].focus(); }
    else if (e.key === "ArrowUp") { e.preventDefault(); items[(i - 1 + items.length) % items.length].focus(); }
  });

  // ?duzen=bento&tasarim=cam bağlantısı paylaşılabilir; yoksa son seçim
  const q = new URLSearchParams(location.search);
  const pick = (key, list, fallback) => (q.has(key) && list.some((x) => x.id === q.get(key)) ? q.get(key) : store.get(fallback));
  applyLayout(pick("duzen", LAYOUTS, "ww-layout"));
  applyTheme(pick("tasarim", THEMES, "ww-design"));
})();
