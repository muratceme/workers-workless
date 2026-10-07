/* =========================================================================
   Workers / Workless — uygulama
   Rotalar:  #/                     ana sayfa
             #/code | #/agents      sektörler
             #/<mod>/<sektör>                       departmanlar
             #/<mod>/<sektör>/<departman>           roller
             #/<mod>/<sektör>/<departman>/<rol>     görevler
             #/<mod>/<sektör>/<departman>/<rol>/<görev>   görev detayı
   ========================================================================= */
(() => {
  "use strict";
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const { CONFIG, DEPTS, SECTORS, SOON, READY } = window.WW;

  /* ------------------------------ yardımcılar ------------------------------ */
  const TR = { ç: "c", ğ: "g", ı: "i", İ: "i", ö: "o", ş: "s", ü: "u", Ç: "c", Ğ: "g", Ö: "o", Ş: "s", Ü: "u" };
  const slug = (s) => s.replace(/[çğıİöşüÇĞÖŞÜ]/g, (c) => TR[c]).toLowerCase()
    .replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  const norm = (s) => s.toLocaleLowerCase("tr").replace(/[çğıöşü]/g, (c) => TR[c]);
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const initials = (s) => s.split(/\s+/).filter((w) => /^[A-ZÇĞİÖŞÜ]/.test(w)).slice(0, 2).map((w) => w[0]).join("");
  const fmt = (n) => n.toLocaleString("tr-TR");
  const MODES = {
    code:   { label: "Kod Blokları", folder: "code-blocks", unit: "kod bloğu", c1: "#0071e3", c2: "#00b8b0" },
    agents: { label: "AI Agents",    folder: "agents",      unit: "agent",     c1: "#7c3aed", c2: "#ff2d78" },
  };

  /* ------------------------------ veri ağacı ------------------------------ */
  const TREE = SECTORS.map((s) => ({
    ...s,
    depts: s.depts.map((id) => {
      const d = DEPTS[id];
      return {
        id, name: d.name, short: d.short, desc: d.desc, shared: !!d.shared,
        roles: d.roles.map((r) => ({
          ...r,
          tasks: r.tasks.map(([name, desc, hours]) => ({ id: slug(name), name, desc, hours })),
        })),
      };
    }),
  }));
  const sum = (arr, f) => arr.reduce((a, x) => a + f(x), 0);
  const roleHours = (r) => sum(r.tasks, (t) => t.hours);
  const deptTasks = (d) => sum(d.roles, (r) => r.tasks.length);
  const sectorTasks = (s) => sum(s.depts, deptTasks);

  // Her görev tek bir kez listelenir: ortak departmanlar (İK, Finans, BT) tüm sektörlerde
  // aynı görevi gösterir, kodu da tek bir klasörde durur.
  const INDEX = [];
  const seen = new Set();
  TREE.forEach((s) => s.depts.forEach((d) => d.roles.forEach((r) => r.tasks.forEach((t) => {
    if (seen.has(t.id)) return;
    seen.add(t.id);
    const where = d.shared ? "Tüm sektörler" : s.name;
    INDEX.push({ s, d, r, t, where, hay: norm(`${t.name} ${r.name} ${d.name} ${where} ${t.desc}`) });
  }))));

  const readyOf = (m, t) => (READY[MODES[m].folder] || {})[t.id];
  const isReady = (t) => !!(readyOf("code", t) || readyOf("agents", t));
  const readyCount = (tasks, m) => tasks.filter((t) => readyOf(m, t)).length;

  const STATS = {
    sectors: TREE.length,
    depts: new Set(TREE.flatMap((s) => s.depts.map((d) => d.id))).size,
    roles: new Set(INDEX.map((x) => x.r.id)).size,
    tasks: INDEX.length,
    hours: sum(INDEX, (x) => x.t.hours),
    ready: Object.keys(READY["code-blocks"]).length + Object.keys(READY.agents).length,
  };

  /* ------------------------------ ikonlar ------------------------------ */
  const P = {
    bank: '<path d="M3 10l9-6 9 6M5 10v8M9.5 10v8M14.5 10v8M19 10v8M3 20h18"/>',
    shield: '<path d="M12 3l8 3v6c0 4.5-3.4 8.2-8 9-4.6-.8-8-4.5-8-9V6z"/><path d="M8.5 12l2.5 2.5 4.5-5"/>',
    factory: '<path d="M3 20V10l5 3V10l5 3V6h3l1 7h4v7z"/><path d="M7 17h2M12 17h2M17 17h1"/>',
    thread: '<path d="M7 3h10M7 21h10M8 3v18M16 3v18"/><path d="M8 7l8 3M8 11l8 3M8 15l8 3"/>',
    bag: '<path d="M5 8h14l-1 12H6z"/><path d="M9 8V6a3 3 0 0 1 6 0v2"/>',
    pulse: '<path d="M3 12h4l2-5 4 10 2-5h6"/>',
    truck: '<path d="M3 6h11v10H3zM14 10h4l3 3v3h-7"/><circle cx="7" cy="18" r="1.6"/><circle cx="17" cy="18" r="1.6"/>',
    scale: '<path d="M12 4v16M6 20h12M5 8h14M5 8l-3 6a3 3 0 0 0 6 0zM19 8l-3 6a3 3 0 0 0 6 0z"/>',
    bolt: '<path d="M13 3L5 13h6l-1 8 8-10h-6z"/>',
    book: '<path d="M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2z"/><path d="M4 19V5"/>',
    code: '<path d="M8 7l-5 5 5 5M16 7l5 5-5 5"/>',
    spark: '<path d="M12 3l1.8 4.7L18.5 9.5l-4.7 1.8L12 16l-1.8-4.7L5.5 9.5l4.7-1.8z"/>',
  };
  const icon = (n, size = 24) =>
    `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${P[n] || P.code}</svg>`;

  /* ------------------------------ durum ------------------------------ */
  let mode = "code";
  const repoUrl = () => `https://github.com/${CONFIG.githubUser}/${CONFIG.repo}`;
  const taskPath = (m, s, d, r, t) => `${MODES[m].folder}/${s.id}/${d.id}/${r.id}/${t.id}`;
  const href = (...parts) => "#/" + parts.filter(Boolean).join("/");

  function setMode(m) {
    mode = m;
    document.body.dataset.mode = m;
    $$("[data-mode-btn]").forEach((b) => b.setAttribute("aria-selected", String(b.dataset.modeBtn === m)));
  }

  /* ------------------------------ tema ------------------------------ */
  const root = document.documentElement;
  const store = {
    get: (k) => { try { return localStorage.getItem(k); } catch { return null; } },
    set: (k, v) => { try { localStorage.setItem(k, v); } catch {} },
  };
  const sysDark = window.matchMedia("(prefers-color-scheme: dark)");
  root.dataset.theme = store.get("ww-theme") || (sysDark.matches ? "dark" : "light");
  $("#themeToggle").addEventListener("click", () => {
    root.dataset.theme = root.dataset.theme === "dark" ? "light" : "dark";
    store.set("ww-theme", root.dataset.theme);
  });

  /* ------------------------------ kart şablonları ------------------------------ */
  const sectorCard = (s, i, m = mode) => `
    <a class="card spot" href="${href(m, s.id)}" style="--i:${i}">
      <span class="card-icon">${icon(s.icon, 26)}</span>
      <span class="card-go" aria-hidden="true">→</span>
      <h3>${esc(s.name)}</h3>
      <p>${esc(s.tagline)}</p>
      <div class="card-meta"><span class="tag">${s.depts.length} departman</span><span class="tag accent">${sectorTasks(s)} görev</span></div>
    </a>`;

  const soonCard = (x, i) => `
    <div class="card soon" style="--i:${i}">
      <span class="card-icon">${icon(x.icon, 26)}</span>
      <h3>${esc(x.name)}</h3>
      <p>Yakında. Bu sektörü siz başlatın.</p>
      <div class="card-meta"><a class="tag" href="${repoUrl()}" target="_blank" rel="noopener">Katkı ver ↗</a></div>
    </div>`;

  const deptCard = (s, d, i) => `
    <a class="card spot" href="${href(mode, s.id, d.id)}" style="--i:${i}">
      <span class="mono-tile">${esc(d.short)}</span>
      <span class="card-go" aria-hidden="true">→</span>
      <h3>${esc(d.name)}</h3>
      <p>${esc(d.desc)}</p>
      <div class="card-meta">${d.shared ? '<span class="tag">Tüm sektörlerde ortak</span>' : ""}<span class="tag">${d.roles.length} rol</span><span class="tag accent">${deptTasks(d)} görev</span></div>
    </a>`;

  const roleCard = (s, d, r, i) => {
    const shown = r.tasks.slice(0, 3).map((t) => `<li>${esc(t.name)}${readyOf(mode, t) ? ' <span class="tag ok sm">Hazır</span>' : ""}</li>`).join("");
    const nReady = readyCount(r.tasks, mode);
    const more = r.tasks.length > 3 ? `<li class="more">+${r.tasks.length - 3} görev daha</li>` : "";
    return `
    <a class="card spot" href="${href(mode, s.id, d.id, r.id)}" style="--i:${i}">
      <span class="avatar" data-level="${esc(r.level)}">${esc(initials(r.name))}</span>
      <span class="tag level">${esc(r.level)}</span>
      <h3>${esc(r.name)}</h3>
      <ul class="role-tasks">${shown}${more}</ul>
      <div class="card-meta"><span class="tag accent">~${roleHours(r)} saat / hafta</span>${nReady ? `<span class="tag ok">${nReady} hazır</span>` : ""}</div>
    </a>`;
  };

  const taskCard = (s, d, r, t, i) => `
    <a class="card spot task-card" href="${href(mode, s.id, d.id, r.id, t.id)}" style="--i:${i}">
      <span class="task-num">${String(i + 1).padStart(2, "0")}</span>
      <h3>${esc(t.name)}</h3>
      <p>${esc(t.desc)}</p>
      <div class="card-meta">
        <span class="tag accent">~${t.hours} sa / hafta</span>
        <span class="tag">${mode === "code" ? "Python" : "Agent"}</span>
        ${readyOf(mode, t) ? '<span class="tag ok">Hazır · indir</span>' : '<span class="tag">Yakında</span>'}
      </div>
      <span class="task-cta">${!readyOf(mode, t) ? "Detaylar" : mode === "code" ? "Kodu incele ve indir" : "Agent'ı kur"} →</span>
    </a>`;

  /* ------------------------------ explorer ------------------------------ */
  function renderExplorer(s, d, r) {
    const level = r ? 3 : d ? 2 : s ? 1 : 0;
    const M = MODES[mode];

    // stepper
    $$("#stepper li").forEach((li, i) => {
      li.className = i < level ? "done" : i === level ? "current" : "";
    });

    // breadcrumb
    const crumbs = [[M.label, href(mode)]];
    if (s) crumbs.push([s.name, href(mode, s.id)]);
    if (d) crumbs.push([d.name, href(mode, s.id, d.id)]);
    if (r) crumbs.push([r.name, href(mode, s.id, d.id, r.id)]);
    $("#crumbs").innerHTML = crumbs.map(([n, h], i) =>
      i === crumbs.length - 1 ? `<span class="here">${esc(n)}</span>` : `<a href="${h}">${esc(n)}</a><span class="sep">/</span>`
    ).join("");

    // başlık + alt başlık
    const title = $("#exTitle"), sub = $("#exSub");
    let html = "";
    if (level === 0) {
      title.innerHTML = mode === "code" ? `Hangi sektördesiniz?` : `Hangi sektörde <span class="grad">agent</span> çalışsın?`;
      sub.textContent = mode === "code"
        ? "Kod blokları: indir, çalıştır. API anahtarı yok, veri bilgisayarınızdan çıkmaz."
        : "AI Agents: indir, .env dosyasına kendi API anahtarınızı ekleyin, kendi bilgisayarınızda çalıştırın.";
      html = TREE.map((x, i) => sectorCard(x, i)).join("") + SOON.map((x, i) => soonCard(x, TREE.length + i)).join("");
    } else if (level === 1) {
      title.textContent = s.name;
      sub.textContent = `${s.depts.length} departman, ${sectorTasks(s)} ${M.unit}. Bir departman seçin.`;
      html = s.depts.map((x, i) => deptCard(s, x, i)).join("");
    } else if (level === 2) {
      title.textContent = d.name;
      sub.textContent = `${s.name} sektöründe ${d.roles.length} rol. Hangi koltuktasınız?`;
      html = d.roles.map((x, i) => roleCard(s, d, x, i)).join("");
    } else {
      title.textContent = r.name;
      const nr = readyCount(r.tasks, mode);
      sub.textContent = `${r.tasks.length} görev${nr ? `, ${nr} tanesi indirilebilir` : ""} · haftada yaklaşık ${roleHours(r)} saat geri kazanım.`;
      html = r.tasks.map((x, i) => taskCard(s, d, r, x, i)).join("");
    }
    // animasyonu yeniden tetikle
    [title, sub].forEach((el) => { el.style.animation = "none"; void el.offsetWidth; el.style.animation = ""; });
    $("#exGrid").innerHTML = html;
  }

  /* ------------------------------ sözdizimi vurgulama ------------------------------ */
  function highlight(code) {
    const re = /(#[^\n]*)|("""[\s\S]*?"""|f?"[^"\n]*"|'[^'\n]*')|\b(def|import|from|return|if|else|for|in|as|None|True|False|with|class|async|await|not|and|or|cd|pip|python|npx|cp)\b|\b(\d+(?:\.\d+)?)\b|(\w+)(?=\()/g;
    let out = "", last = 0, m;
    while ((m = re.exec(code))) {
      out += esc(code.slice(last, m.index));
      const cls = m[1] ? "c" : m[2] ? "s" : m[3] ? "k" : m[4] ? "n" : "f";
      out += `<span class="tok-${cls}">${esc(m[0])}</span>`;
      last = re.lastIndex;
    }
    return out + esc(code.slice(last));
  }
  const codebox = (file, code, id) => `
    <div class="codebox">
      <div class="codebox-bar"><span>${esc(file)}</span><button class="copy-btn" data-copy="${id}">Kopyala</button></div>
      <pre id="${id}"><code>${highlight(code)}</code></pre>
    </div>`;

  /* ------------------------------ görev sayfası (sheet) ------------------------------ */
  const kb = (n) => (n < 1048576 ? `${Math.max(1, Math.round(n / 1024))} KB` : `${(n / 1048576).toFixed(1)} MB`);
  const ghTree = (m, t) => `${repoUrl()}/tree/${CONFIG.branch}/library/${MODES[m].folder}/${t.id}`;
  const issueUrl = (m, t) => `${repoUrl()}/issues/new?labels=gorev-talebi&title=${encodeURIComponent(`Görev talebi: ${t.name} (${MODES[m].label})`)}`;
  const files = new Map();
  const loadFile = (url) => {
    if (!files.has(url)) files.set(url, fetch(url).then((r) => (r.ok ? r.text() : Promise.reject(new Error(r.status)))));
    return files.get(url);
  };

  function readyBody(m, t, pkg) {
    const agent = m === "agents";
    const deps = pkg.deps.map((x) => x.split("==")[0]).join(" · ");
    const steps = [
      ["Paketi indirin ve klasöre çıkarın", "İndirmeden önce kullanım koşullarını onaylamanız istenir. Terminali çıkardığınız klasörde açın.", `cd ${t.id}`],
      ["Bağımlılıkları kurun", `Python ${esc(pkg.python)} gerekir.`, "pip install -r requirements.txt"],
      ...(agent ? [["API anahtarınızı ekleyin", "<code>.env.example</code> dosyasını <code>.env</code> olarak kopyalayıp kendi anahtarınızı yazın. Agent veri göndermeden önce onayınızı ister.",
        "copy .env.example .env      # macOS / Linux: cp .env.example .env"]] : []),
      ["Örnek veriyle deneyin", "Paket kurgusal örnek verilerle gelir.", `python ${pkg.main}`],
      ["Kendi verinizle çalıştırın", "Girdi standardı paketteki README.md dosyasında.", pkg.run],
    ];
    return `
      <div class="facts">
        <div class="fact"><b>~${t.hours} sa</b><span>haftalık kazanım</span></div>
        <div class="fact"><b>Python ${esc(pkg.python)}</b><span>${esc(deps)}</span></div>
        <div class="fact"><b>${agent ? "Kendi API anahtarınız" : "İnternet yok"}</b><span>${agent ? "Claude · OpenAI · Ollama (yerel)" : "veri bilgisayarınızdan çıkmaz"}</span></div>
      </div>
      <div class="actions">
        <button class="btn btn-primary" data-download="${m}">İndir · ZIP ${kb(pkg.size)}</button>
        <a class="btn btn-ghost" href="${ghTree(m, t)}" target="_blank" rel="noopener">GitHub'da incele ↗</a>
      </div>
      <p class="note">Sürüm ${esc(pkg.version)} · ${pkg.files} dosya · otomatik testlerden geçti</p>
      <div class="io">
        <div><h4>Girdi</h4><ul>${pkg.inputs.map((x) => `<li>${esc(x)}</li>`).join("")}</ul></div>
        <div><h4>Çıktı</h4><ul>${pkg.outputs.map((x) => `<li>${esc(x)}</li>`).join("")}</ul></div>
      </div>
      <h4>Kurulum</h4>
      <ol class="steps">${steps.map(([h, p, c], i) => `<li><h5>${h}</h5>${p ? `<p class="step-p">${p}</p>` : ""}${codebox("terminal", c, `st${i}`)}</li>`).join("")}</ol>
      <h4>Kod</h4>
      <div id="preview">${codebox(pkg.main, "Yükleniyor…", "pv")}</div>`;
  }

  function wipBody(m, t) {
    return `
      <div class="wip">
        <span class="wip-dot" aria-hidden="true"></span>
        <div>
          <b>Yapım aşamasında</b>
          <p>Bu görevin ${m === "code" ? "kod bloğu" : "AI agent"} sürümü henüz yazılmadı. Sıralamayı gelen talepler belirliyor.</p>
        </div>
      </div>
      <div class="facts">
        <div class="fact"><b>~${t.hours} sa</b><span>tahmini haftalık kazanım</span></div>
        <div class="fact"><b>${m === "code" ? "API yok" : "Kendi anahtarınız"}</b><span>${m === "code" ? "tamamen yerel çalışacak" : "Claude · OpenAI · Ollama"}</span></div>
        <div class="fact"><b>Ücretsiz</b><span>açık kaynak, MIT lisansı</span></div>
      </div>
      <div class="actions">
        <a class="btn btn-primary" href="${issueUrl(m, t)}" target="_blank" rel="noopener">Bu görevi talep et ↗</a>
        <a class="btn btn-ghost" href="${repoUrl()}" target="_blank" rel="noopener">Katkı ver ↗</a>
      </div>`;
  }

  let lastFocus = null;
  let current = null;
  function openSheet(s, d, r, t) {
    current = { s, d, r, t };
    const other = mode === "code" ? "agents" : "code";
    const O = MODES[other];
    const pkg = readyOf(mode, t);
    const otherReady = readyOf(other, t);
    const pathChips = [d.shared ? "Tüm sektörler" : s.name, d.name, r.name].map((x) => `<span class="tag">${esc(x)}</span>`).join("");

    $("#sheetBody").innerHTML = `
      <div class="sheet-path">${pathChips}<span class="tag accent">${MODES[mode].label}</span>${pkg ? '<span class="tag ok">Hazır</span>' : ""}</div>
      <h2 id="sheetTitle">${esc(t.name)}</h2>
      <p class="sheet-desc">${esc(t.desc)}</p>
      ${pkg ? readyBody(mode, t, pkg) : wipBody(mode, t)}
      <a class="crosslink" href="${href(other, s.id, d.id, r.id, t.id)}">
        <span class="x-ic" style="background:linear-gradient(135deg,${O.c1},${O.c2})">${icon(other === "code" ? "code" : "spark", 18)}</span>
        <span><b>${other === "code" ? "Kod bloğu" : "AI agent"} sürümü${otherReady ? " de hazır" : " (yakında)"}</b>
        <span>${other === "code" ? "API anahtarı gerektirmeyen, kural tabanlı sürüm." : "Dağınık belgelerde yorum yapabilen, kendi API anahtarınızla çalışan sürüm."}</span></span>
        <span class="arr">→</span>
      </a>`;

    if (pkg) {
      loadFile(`files/${MODES[mode].folder}/${t.id}/${pkg.main}`).then((code) => {
        if (!current || current.t !== t || !$("#preview")) return;
        const lines = code.split("\n");
        const shown = lines.slice(0, 90).join("\n") + (lines.length > 90 ? `\n\n# … ${lines.length - 90} satır daha (tamamı pakette)` : "");
        $("#preview").innerHTML = codebox(`${pkg.main} · ${lines.length} satır`, shown, "pv");
      }, () => { if ($("#preview")) $("#preview").innerHTML = '<p class="note">Önizleme yüklenemedi.</p>'; });
    }

    const wasOpen = !$("#sheet").hidden;
    $("#sheet").hidden = false;
    $("#sheetBackdrop").hidden = false;
    document.body.style.overflow = "hidden";
    $("#sheet").scrollTop = 0;
    if (!wasOpen) { lastFocus = document.activeElement; $("#sheetClose").focus(); }
  }

  /* ------------------------------ indirme onayı ------------------------------ */
  const CONSENT = {
    common: [
      "Yazılımın ücretsiz ve <b>“olduğu gibi”</b> sunulduğunu, doğruluğu veya amaca uygunluğu konusunda hiçbir garanti verilmediğini kabul ediyorum.",
      "Yazılımı <b>kendi bilgisayarımda, kendi verimle ve kendi riskimle</b> çalıştıracağımı; kullanımından doğabilecek veri kaybı, hatalı sonuç, mali kayıp veya başka herhangi bir zarardan yazarların sorumlu tutulamayacağını kabul ediyorum.",
      "Çıktıların karar desteği olduğunu ve bunlara dayanarak işlem yapmadan önce <b>kontrol edeceğimi</b> kabul ediyorum.",
    ],
    agents: [
      "Agent'ın <b>kendi API anahtarımla</b> üçüncü taraf bir yapay zekâ sağlayıcısına veri göndereceğini; <b>API ücretlerinin</b>, anahtarın güvenliğinin ve gönderilen verinin <b>KVKK/GDPR</b> ile şirket politikalarına uygunluğunun tamamen benim sorumluluğumda olduğunu kabul ediyorum.",
    ],
  };
  let pending = null;
  function openConsent(m) {
    const t = current.t, pkg = readyOf(m, t);
    pending = { m, t, pkg };
    const items = [...CONSENT.common, ...(m === "agents" ? CONSENT.agents : [])];
    $("#consentPkg").innerHTML = `<b>${esc(t.name)}</b> · ${m === "agents" ? "AI Agent" : "Kod Bloğu"} · v${esc(pkg.version)} · ${kb(pkg.size)}`;
    $("#consentList").innerHTML = items.map((x, i) =>
      `<label class="check"><input type="checkbox" data-c="${i}"><span>${x}</span></label>`).join("");
    $("#consentHash").textContent = pkg.sha256;
    syncConsent();
    $("#consent").hidden = false;
    $("#consentList input").focus();
  }
  function syncConsent() {
    const boxes = $$("#consentList input");
    const ok = boxes.length > 0 && boxes.every((b) => b.checked);
    $("#consentAll").checked = ok;
    $("#consentGo").disabled = !ok;
  }
  function closeConsent() { $("#consent").hidden = true; pending = null; }
  $("#consentList").addEventListener("change", syncConsent);
  $("#consentAll").addEventListener("change", (e) => { $$("#consentList input").forEach((b) => (b.checked = e.target.checked)); syncConsent(); });
  $("#consentCancel").addEventListener("click", closeConsent);
  $("#consent").addEventListener("click", (e) => { if (e.target.id === "consent") closeConsent(); });
  $("#consentGo").addEventListener("click", () => {
    if (!pending || $("#consentGo").disabled) return;
    const a = document.createElement("a");
    a.href = pending.pkg.zip;
    a.download = `${pending.t.id}-${pending.m === "code" ? "kod" : "agent"}-v${pending.pkg.version}.zip`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    closeConsent();
    toast("İndirme başladı. Önce README.md dosyasını okuyun.");
  });

  function hideSheet() {
    if ($("#sheet").hidden) return;
    $("#sheet").hidden = true;
    $("#sheetBackdrop").hidden = true;
    document.body.style.overflow = "";
    if (lastFocus && lastFocus.focus) lastFocus.focus();
  }
  function closeSheet() {
    const p = parts();
    if (p.length === 5) location.hash = href(...p.slice(0, 4));
    else hideSheet();
  }
  $("#sheetClose").addEventListener("click", closeSheet);
  $("#sheetBackdrop").addEventListener("click", closeSheet);

  document.addEventListener("click", (e) => {
    const c = e.target.closest("[data-copy]");
    if (c) {
      const txt = $("#" + c.dataset.copy).innerText;
      (navigator.clipboard ? navigator.clipboard.writeText(txt) : Promise.reject()).then(
        () => { c.textContent = "Kopyalandı ✓"; setTimeout(() => (c.textContent = "Kopyala"), 1500); },
        () => toast("Kopyalanamadı, elle seçin."),
      );
    }
    const dl = e.target.closest("[data-download]");
    if (dl) openConsent(dl.dataset.download);
  });

  /* ------------------------------ yönlendirme ------------------------------ */
  const parts = () => location.hash.replace(/^#\/?/, "").split("/").filter(Boolean);

  function route() {
    const p = parts();
    const navKey = p[0] || "";
    $$("[data-nav]").forEach((a) => a.classList.toggle("active", a.dataset.nav === navKey));

    if (!MODES[p[0]]) {
      hideSheet();
      $("#home").hidden = false;
      $("#explorer").hidden = true;
      if (p[0] && document.getElementById(p[0])) document.getElementById(p[0]).scrollIntoView({ behavior: "smooth" });
      else window.scrollTo({ top: 0 });
      return;
    }

    const prevMode = mode;
    setMode(p[0]);
    const s = TREE.find((x) => x.id === p[1]);
    const d = s && s.depts.find((x) => x.id === p[2]);
    const r = d && d.roles.find((x) => x.id === p[3]);
    const t = r && r.tasks.find((x) => x.id === p[4]);

    const wasHome = !$("#home").hidden;
    $("#home").hidden = true;
    $("#explorer").hidden = false;

    const key = [p[0], s && s.id, d && d.id, r && r.id].join("/");
    if (key !== route.lastKey || wasHome || prevMode !== mode) {
      renderExplorer(s, d, r);
      if (!t) window.scrollTo({ top: 0 });
      route.lastKey = key;
    }
    if (t) openSheet(s, d, r, t);
    else hideSheet();
  }
  window.addEventListener("hashchange", route);

  // mod düğmeleri: aynı konumda kal, sadece modu değiştir
  $$("[data-mode-btn]").forEach((b) => b.addEventListener("click", () => {
    const p = parts();
    p[0] = b.dataset.modeBtn;
    location.hash = href(...p.slice(0, 4));
  }));

  /* ------------------------------ ana sayfa ------------------------------ */
  $("#homeSectors").innerHTML = TREE.map((s, i) => sectorCard(s, i, "code")).join("");
  $("#homeReady").innerHTML = INDEX.filter((x) => isReady(x.t)).map((x, i) => {
    const c = readyOf("code", x.t), g = readyOf("agents", x.t);
    return `
    <a class="card spot ready-card" href="${href(c ? "code" : "agents", x.s.id, x.d.id, x.r.id, x.t.id)}" style="--i:${i}">
      <span class="task-num">${esc(x.where)} › ${esc(x.d.name)} › ${esc(x.r.name)}</span>
      <h3>${esc(x.t.name)}</h3>
      <p>${esc(x.t.desc)}</p>
      <div class="card-meta">
        ${c ? `<span class="tag ok">Kod bloğu v${esc(c.version)}</span>` : ""}
        ${g ? `<span class="tag ok agent">AI Agent v${esc(g.version)}</span>` : ""}
        <span class="tag accent">~${x.t.hours} sa / hafta</span>
      </div>
    </a>`;
  }).join("");
  $("#ghLink").href = repoUrl();

  // marquee
  const shuffled = INDEX.slice().sort(() => Math.random() - 0.5);
  const chip = (x) => `<a class="chip" href="${href("code", x.s.id, x.d.id, x.r.id, x.t.id)}">${esc(x.t.name)} <i>${esc(x.r.name)}</i></a>`;
  const half = Math.min(18, Math.floor(shuffled.length / 2));
  const row1 = shuffled.slice(0, half).map(chip).join("");
  const row2 = shuffled.slice(half, half * 2).map(chip).join("");
  $("#mq1").innerHTML = row1 + row1;
  $("#mq2").innerHTML = row2 + row2;

  // dönen rol cümlesi
  const roles = [];
  TREE.forEach((s) => s.depts.forEach((d) => d.roles.forEach((r) => {
    if (!roles.some((x) => x.name === r.name)) roles.push(r);
  })));
  let ri = 0;
  setInterval(() => {
    const a = $("#rotRole"), b = $("#rotCount");
    a.classList.add("out"); b.classList.add("out");
    setTimeout(() => {
      ri = (ri + 1 + Math.floor(Math.random() * 3)) % roles.length;
      a.textContent = roles[ri].name; b.textContent = roles[ri].tasks.length;
      a.classList.remove("out"); b.classList.remove("out");
    }, 350);
  }, 2600);

  // logo: Workers → Workless
  setInterval(() => $(".logo").classList.toggle("swapped"), 3200);

  /* ------------------------------ terminal demosu ------------------------------ */
  const DEMOS = [
    { file: "cv-raporlama / main.py", before: "8 saat", after: "0,9 sn", lines: [
      ["p", "$ python main.py --girdi ./basvurular --ilan ilan.txt"],
      ["dim", "-> 142 dosya bulundu (PDF 118 · DOCX 24)"],
      ["ok", "[OK] 142 CV işlendi (3 tanesinde uyarı var)"],
      ["ok", "[OK] 38 aday ilan uyumu %60 ve üzerinde"],
      ["hl", "[OK] Rapor: cikti/aday_raporu.xlsx"],
    ]},
    { file: "banka-mutabakati / main.py", before: "6 saat", after: "1,4 sn", lines: [
      ["p", "$ python main.py --banka ekim_ekstre.xlsx --defter muavin_102.xlsx"],
      ["ok", "[OK] 3.171 kayıt eşleşti · açık kalem: banka 14, defter 9"],
      ["dim", "   komisyon/BSMV 8 · zamanlama 6 · mükerrer 2 · tutar hatası 1"],
      ["hl", "[OK] Rapor: cikti/mutabakat.xlsx"],
    ]},
    { file: "banka-mutabakati / agent.py", before: "2 saat", after: "40 sn", agent: true, lines: [
      ["p", "$ python agent.py --banka ekim_ekstre.xlsx --defter muavin_102.xlsx"],
      ["ok", "[OK] Deterministik eşleştirme: 3.171 eşleşti · açık kalem 23"],
      ["dim", "[?] Yalnızca 23 açık kalem gönderilecek. Devam edilsin mi? [e/H] e"],
      ["ok", "[OK] 23/23 açık kalem yorumlandı · yevmiye taslakları hazır"],
      ["hl", "[OK] Rapor: cikti/mutabakat_ai.xlsx"],
    ]},
  ];
  const term = $("#term");
  let di = 0;
  const wait = (ms) => new Promise((res) => setTimeout(res, ms));
  async function runDemo() {
    const reduce = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    for (;;) {
      const demo = DEMOS[di++ % DEMOS.length];
      $("#termTitle").textContent = demo.file;
      $("#termBefore").textContent = demo.before;
      $("#termAfter").textContent = demo.after;
      term.innerHTML = "";
      for (const [cls, text] of demo.lines) {
        const line = document.createElement("div");
        line.className = cls;
        term.appendChild(line);
        if (cls === "p" && !reduce) {
          for (let i = 1; i <= text.length; i++) { line.innerHTML = esc(text.slice(0, i)) + '<span class="cursor"></span>'; await wait(18); }
          line.textContent = text;
          await wait(350);
        } else {
          line.textContent = text;
          await wait(reduce ? 0 : 520);
        }
      }
      term.insertAdjacentHTML("beforeend", '<div><span class="p">$ </span><span class="cursor"></span></div>');
      await wait(3600);
    }
  }
  runDemo();

  /* ------------------------------ sayaçlar & reveal ------------------------------ */
  function countUp(el) {
    const target = STATS[el.dataset.count];
    const t0 = performance.now(), dur = 1400;
    const step = (now) => {
      const k = Math.min(1, (now - t0) / dur);
      el.textContent = fmt(Math.round(target * (1 - Math.pow(1 - k, 4))));
      if (k < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }
  const io = new IntersectionObserver((entries) => entries.forEach((e) => {
    if (!e.isIntersecting) return;
    e.target.classList.add("in");
    const c = e.target.querySelector("[data-count]");
    if (c) countUp(c);
    io.unobserve(e.target);
  }), { threshold: 0.15 });
  $$(".reveal, .device").forEach((el) => io.observe(el));

  /* ------------------------------ imleç ışığı ------------------------------ */
  document.addEventListener("pointermove", (e) => {
    const c = e.target.closest && e.target.closest(".spot");
    if (!c) return;
    const b = c.getBoundingClientRect();
    c.style.setProperty("--mx", `${e.clientX - b.left}px`);
    c.style.setProperty("--my", `${e.clientY - b.top}px`);
  });

  /* ------------------------------ komut paleti ------------------------------ */
  const pal = $("#palette"), pin = $("#paletteInput"), plist = $("#paletteList");
  let results = [], sel = 0;
  function openPalette() { pal.hidden = false; pin.value = ""; filter(); pin.focus(); }
  function closePalette() { pal.hidden = true; }
  function mark(text, words) {
    let h = esc(text);
    words.forEach((w) => {
      if (!w) return;
      const n = norm(text), i = n.indexOf(w);
      if (i >= 0) h = esc(text.slice(0, i)) + "<mark>" + esc(text.slice(i, i + w.length)) + "</mark>" + esc(text.slice(i + w.length));
    });
    return h;
  }
  function filter() {
    const words = norm(pin.value.trim()).split(/\s+/).filter(Boolean);
    results = words.length
      ? INDEX.filter((x) => words.every((w) => x.hay.includes(w)))
      : [...INDEX.filter((x) => isReady(x.t)), ...INDEX.filter((x) => !isReady(x.t)).slice(0, 6)];
    results = results.slice(0, 40);
    sel = 0;
    plist.innerHTML = results.length
      ? results.map((x, i) => `
        <li role="option" data-i="${i}" aria-selected="${i === 0}">
          <span class="p-ic">${esc(x.d.short)}</span>
          <span class="p-main"><b>${mark(x.t.name, words.slice(0, 1))}</b><small>${esc(x.where)} › ${esc(x.d.name)} › ${esc(x.r.name)}</small></span>
          <span class="p-h">${isReady(x.t) ? '<span class="tag ok sm">Hazır</span> ' : ""}~${x.t.hours} sa/hf</span>
        </li>`).join("")
      : `<li class="palette-empty">Sonuç yok. Bu görevi GitHub'da önerebilirsiniz.</li>`;
  }
  function choose(i) {
    const x = results[i];
    if (!x) return;
    closePalette();
    const m = MODES[parts()[0]] ? parts()[0] : "code";
    location.hash = href(m, x.s.id, x.d.id, x.r.id, x.t.id);
  }
  function moveSel(dir) {
    if (!results.length) return;
    sel = (sel + dir + results.length) % results.length;
    $$("li[data-i]", plist).forEach((li) => li.setAttribute("aria-selected", String(+li.dataset.i === sel)));
    const cur = $(`li[data-i="${sel}"]`, plist);
    if (cur) cur.scrollIntoView({ block: "nearest" });
  }
  pin.addEventListener("input", filter);
  pin.addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown") { e.preventDefault(); moveSel(1); }
    else if (e.key === "ArrowUp") { e.preventDefault(); moveSel(-1); }
    else if (e.key === "Enter") { e.preventDefault(); choose(sel); }
  });
  plist.addEventListener("click", (e) => { const li = e.target.closest("li[data-i]"); if (li) choose(+li.dataset.i); });
  pal.addEventListener("click", (e) => { if (e.target === pal) closePalette(); });
  $("#openPalette").addEventListener("click", openPalette);
  if (/Mac|iPhone|iPad/.test(navigator.platform)) $(".kbd-hint").textContent = "⌘K";

  document.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") { e.preventDefault(); pal.hidden ? openPalette() : closePalette(); }
    else if (e.key === "Escape") { if (!$("#consent").hidden) closeConsent(); else if (!pal.hidden) closePalette(); else closeSheet(); }
    else if (e.key === "/" && pal.hidden && !/input|textarea/i.test(document.activeElement.tagName)) { e.preventDefault(); openPalette(); }
  });

  /* ------------------------------ toast ------------------------------ */
  let toastT;
  function toast(msg) {
    const el = $("#toast");
    el.textContent = msg;
    el.classList.add("show");
    clearTimeout(toastT);
    toastT = setTimeout(() => el.classList.remove("show"), 2400);
  }

  route();
})();
