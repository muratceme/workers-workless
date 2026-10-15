/* =========================================================================
   Tasarım seçici (GEÇİCİ). Karar verilince bu dosya, assets/themes.css ve
   index.html'deki iki satır (link + script) ile head'deki küçük betik silinir.
   ========================================================================= */
(() => {
  const GF = "https://fonts.googleapis.com/css2?display=swap&family=";
  const DESIGNS = [
    { id: "", ad: "Orijinal", not: "Şu anki tasarım", sw: ["#fbfbfd", "#0071e3", "#7c3aed"] },
    { id: "editoryal", ad: "Editoryal", not: "Sıcak kâğıt, serif başlık, sade", scheme: "light", sw: ["#F7F6F3", "#2F3437", "#1F6C9F"],
      font: "Instrument+Serif:ital@0;1&family=Geist:wght@400;500;600&family=Geist+Mono:wght@400;500" },
    { id: "isvicre", ad: "İsviçre Baskı", not: "Izgara, kalın tipografi, tek kırmızı", scheme: "light", sw: ["#F4F4F0", "#111111", "#E61919"],
      font: "Archivo+Black&family=Archivo:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600" },
    { id: "terminal", ad: "Terminal", not: "Koyu, monospace, tarama çizgileri", scheme: "dark", sw: ["#0A0A0A", "#EAEAEA", "#FF2A2A"],
      font: "Space+Mono:wght@400;700&family=JetBrains+Mono:wght@400;500;700" },
    { id: "cam", ad: "Cam", not: "Derin siyah, ışık küreleri, çift çerçeve", scheme: "dark", sw: ["#050505", "#3DDC97", "#A78BFA"],
      font: "Outfit:wght@300;400;500;600;700&family=Geist+Mono:wght@400;500" },
    { id: "yumusak", ad: "Yumuşak Yapı", not: "Gümüş zemin, iri grotesk, yüzen kartlar", scheme: "light", sw: ["#F2F3F5", "#0E1116", "#FF5B1F"],
      font: "Bricolage+Grotesque:opsz,wght@12..96,500..800&family=Plus+Jakarta+Sans:wght@400;500;600;700" },
  ];
  const KEY = "ww-design";
  const root = document.documentElement;
  const store = {
    get: () => { try { return localStorage.getItem(KEY) || ""; } catch { return ""; } },
    set: (v) => { try { v ? localStorage.setItem(KEY, v) : localStorage.removeItem(KEY); } catch {} },
  };
  const loaded = new Set();
  const loadFont = (d) => {
    if (!d.font || loaded.has(d.id)) return;
    loaded.add(d.id);
    const l = document.createElement("link");
    l.rel = "stylesheet";
    l.href = GF + d.font;
    document.head.appendChild(l);
  };

  const themeBtn = document.getElementById("themeToggle");
  const userTheme = () => {
    try { const t = localStorage.getItem("ww-theme"); if (t) return t; } catch {}
    return matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  };

  function apply(id) {
    const d = DESIGNS.find((x) => x.id === id) || DESIGNS[0];
    loadFont(d);
    if (d.id) root.dataset.design = d.id; else delete root.dataset.design;
    // Alternatif tasarımların kendi açık/koyu düzeni var; tema düğmesi yalnızca orijinalde çalışır.
    root.dataset.theme = d.scheme || userTheme();
    if (themeBtn) themeBtn.hidden = Boolean(d.scheme);
    menu.querySelectorAll(".ds-item").forEach((b) => b.setAttribute("aria-checked", String(b.dataset.id === d.id)));
    store.set(d.id);
  }

  // Düğme + menü (sağ üst)
  const wrap = document.createElement("div");
  wrap.className = "ds";
  wrap.innerHTML = `
    <button class="icon-btn ds-btn" type="button" aria-haspopup="true" aria-expanded="false" aria-controls="dsMenu" title="Tasarım seç">
      <svg viewBox="0 0 24 24" width="17" height="17" aria-hidden="true"><rect x="3.5" y="3.5" width="7" height="7" rx="1.5" fill="none" stroke="currentColor" stroke-width="2"/><rect x="13.5" y="3.5" width="7" height="7" rx="1.5" fill="currentColor"/><rect x="3.5" y="13.5" width="7" height="7" rx="1.5" fill="currentColor"/><rect x="13.5" y="13.5" width="7" height="7" rx="1.5" fill="none" stroke="currentColor" stroke-width="2"/></svg>
      <span class="ds-label">Tasarım</span>
    </button>
    <div class="ds-menu" id="dsMenu" role="menu" aria-label="Tasarım alternatifleri" hidden>
      <p class="ds-head">Tasarım alternatifleri · geçici</p>
      ${DESIGNS.map((d, i) => `
        <button class="ds-item" type="button" role="menuitemradio" aria-checked="false" data-id="${d.id}">
          <span class="ds-sw" aria-hidden="true">${d.sw.map((c) => `<i style="background:${c}"></i>`).join("")}</span>
          <span class="ds-txt"><b>${d.ad}</b><small>${d.not}</small></span>
          <span class="ds-n">${i === 0 ? "—" : i}</span>
        </button>`).join("")}
    </div>`;
  const actions = document.querySelector(".nav-actions");
  actions.insertBefore(wrap, themeBtn || null);
  const btn = wrap.querySelector(".ds-btn");
  const menu = wrap.querySelector(".ds-menu");

  const open = (v) => {
    menu.hidden = !v;
    btn.setAttribute("aria-expanded", String(v));
    if (v) (menu.querySelector('[aria-checked="true"]') || menu.querySelector(".ds-item")).focus();
  };
  btn.addEventListener("click", () => open(menu.hidden));
  menu.addEventListener("click", (e) => {
    const it = e.target.closest(".ds-item");
    if (!it) return;
    apply(it.dataset.id);
    open(false);
    btn.focus();
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

  // ?tasarim=isvicre bağlantısı paylaşılabilir; yoksa son seçim
  const q = new URLSearchParams(location.search).get("tasarim");
  apply(q !== null && DESIGNS.some((d) => d.id === q) ? q : store.get());
})();
