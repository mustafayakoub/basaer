/* Basaer shared page kit: header, footer, back-to-top, report a note, contact,
 * site statistics, "my readings" and the install button.
 * Include once per page: <script src="assets/basaer.js" defer></script>
 * Pages mark ayat with data-sura and data-aya so "my readings" and reports know where the reader is.
 */
(function () {
  "use strict";

  // Fill these before publishing. Empty values hide the matching control.
  const CONFIG = {
    siteName: "بصائر",
    tagline: "القرآن الكريم ومصاحفه وعلومه آيةً آية",
    contactEmail: "",          // e.g. "contact@basaere.com"
    reportEndpoint: "",        // URL that accepts POST JSON {type, message, email, page, sura, aya, quote}
    social: {                  // full profile URLs
      x: "", youtube: "", telegram: "", facebook: "", instagram: "", whatsapp: "", tiktok: ""
    },
    nav: [
      ["./", "الرئيسية"], ["tajweed.html", "مصحف التجويد"], ["surah.html", "بطاقات السور"],
      ["sources.html", "المصادر والتراخيص"], ["install.html", "ثبّت الموقع"]
    ],
    sections: [
      ["tajweed.html", "مصحف التجويد الملوّن"], ["surah.html", "بطاقات السور"], ["sources.html", "المصادر والتراخيص"],
      ["install.html", "تثبيت الموقع"]
    ]
  };

  const ICON = {
    up: '<path d="M12 19V5M5 12l7-7 7 7"/>',
    note: '<path d="M4 4h16v12H8l-4 4z"/><path d="M8 9h8M8 12h5"/>',
    mail: '<rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 7 9 6 9-6"/>',
    chart: '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
    book: '<path d="M6 3h11a1 1 0 0 1 1 1v16l-6-3-6 3V4a1 1 0 0 1 1-1z"/>',
    phone: '<rect x="7" y="2" width="10" height="20" rx="2"/><path d="M11 18h2"/>',
    desktop: '<rect x="2" y="4" width="20" height="13" rx="2"/><path d="M8 21h8M12 17v4"/>',
    x: '<path d="M4 4l16 16M20 4 4 20"/>',
    youtube: '<rect x="2" y="5" width="20" height="14" rx="4"/><path d="m10 9 5 3-5 3z"/>',
    telegram: '<path d="m21 4-18 7 6 2 2 6 3-4 5 4z"/><path d="m9 13 9-6"/>',
    facebook: '<path d="M15 3h-3a4 4 0 0 0-4 4v3H6v4h2v7h4v-7h3l1-4h-4V7a1 1 0 0 1 1-1h2z"/>',
    instagram: '<rect x="3" y="3" width="18" height="18" rx="5"/><circle cx="12" cy="12" r="4"/><circle cx="17.5" cy="6.5" r="1"/>',
    whatsapp: '<path d="M4 20l1.5-4A8 8 0 1 1 8 18.5z"/><path d="M9 9c0 3 3 6 6 6l1-1.5-2-1-1 1c-1-.5-2-1.5-2.5-2.5l1-1-1-2z"/>',
    tiktok: '<path d="M14 3v11a4 4 0 1 1-4-4"/><path d="M14 3c0 3 2 5 5 5"/>'
  };
  const SOCIAL_NAMES = { x: "إكس", youtube: "يوتيوب", telegram: "تيليجرام", facebook: "فيسبوك", instagram: "إنستغرام", whatsapp: "قناة واتساب", tiktok: "تيك توك" };

  const svg = (name) => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON[name]}</svg>`;
  const el = (html) => { const t = document.createElement("template"); t.innerHTML = html.trim(); return t.content.firstElementChild; };
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const AR = "٠١٢٣٤٥٦٧٨٩";
  const arNum = (n) => Number(n).toLocaleString("ar-EG");
  const toArDigits = (s) => String(s).replace(/\d/g, (d) => AR[d]);

  const store = {
    get(key, fallback) { try { const v = localStorage.getItem("basaer:" + key); return v ? JSON.parse(v) : fallback; } catch (e) { return fallback; } },
    set(key, value) { try { localStorage.setItem("basaer:" + key, JSON.stringify(value)); } catch (e) { /* storage blocked: feature degrades */ } }
  };

  /* ---------- header and footer ---------- */
  function header() {
    const here = location.pathname.split("/").pop() || "index.html";
    const links = CONFIG.nav.map(([href, label]) => `<a href="${href}"${(href === "./" ? "index.html" : href) === here ? ' aria-current="page"' : ""}>${label}</a>`).join("");
    document.body.prepend(el(`<header class="b-header"><div class="wrap"><a class="b-brand" href="./">${CONFIG.siteName}</a><nav class="b-nav" aria-label="أقسام الموقع">${links}</nav></div></header>`));
  }

  function footer() {
    const social = Object.entries(CONFIG.social).filter(([, url]) => url)
      .map(([k, url]) => `<a href="${esc(url)}" target="_blank" rel="noopener" aria-label="${SOCIAL_NAMES[k]}" title="${SOCIAL_NAMES[k]}">${svg(k)}</a>`).join("");
    const sections = CONFIG.sections.map(([href, label]) => `<li><a href="${href}">${label}</a></li>`).join("");
    const f = el(`<footer class="b-footer"><div class="wrap">
      <div class="grid">
        <div><h2>${CONFIG.siteName}</h2><p class="muted">${CONFIG.tagline}</p>
          ${social ? `<div class="b-social">${social}</div>` : ""}
          <div data-install-slot></div></div>
        <div><h2>أقسام الموقع</h2><ul>${sections}</ul></div>
      </div>
      <div class="b-bottom"><span>${CONFIG.siteName} · ${toArDigits(new Date().getFullYear())}</span>
        <div class="actions">
          ${CONFIG.contactEmail ? `<button class="btn" type="button" data-open="contact">${svg("mail")}تواصل معنا</button>` : ""}
          <button class="btn" type="button" data-open="stats">${svg("chart")}إحصاءات الموقع</button>
        </div></div></div></footer>`);
    document.body.append(f);
  }

  /* ---------- back to top with reading progress ---------- */
  function backToTop() {
    const R = 22, C = 2 * Math.PI * R;
    const b = el(`<button class="btn b-top" type="button" aria-label="إلى الأعلى" title="إلى الأعلى" hidden>
      <svg class="ring" viewBox="0 0 50 50" aria-hidden="true"><circle class="track" cx="25" cy="25" r="${R}"/><circle class="prog" cx="25" cy="25" r="${R}" stroke-dasharray="${C}" stroke-dashoffset="${C}"/></svg>${svg("up")}</button>`);
    floatEnd.append(b);
    const prog = b.querySelector(".prog");
    let ticking = false;
    const update = () => {
      const max = document.documentElement.scrollHeight - innerHeight;
      const ratio = max > 0 ? Math.min(1, scrollY / max) : 0;
      b.hidden = scrollY < 500;
      prog.style.strokeDashoffset = String(C * (1 - ratio));
      ticking = false;
    };
    addEventListener("scroll", () => { if (!ticking) { ticking = true; requestAnimationFrame(update); } }, { passive: true });
    b.addEventListener("click", () => { scrollTo({ top: 0 }); document.querySelector("h1, main")?.focus?.({ preventScroll: true }); });
    update();
  }

  /* ---------- dialogs ---------- */
  function dialog(id, title, bodyHtml, footHtml) {
    const d = el(`<dialog class="b-dialog" id="b-${id}" aria-labelledby="b-${id}-t"><header><h2 id="b-${id}-t">${title}</h2><button class="b-x" type="button" aria-label="إغلاق">×</button></header><div class="body">${bodyHtml}</div>${footHtml ? `<div class="foot">${footHtml}</div>` : ""}</dialog>`);
    d.querySelector(".b-x").addEventListener("click", () => d.close());
    d.addEventListener("click", (e) => { if (e.target === d) d.close(); });
    document.body.append(d);
    return d;
  }

  // Where the reader is: the ayah at about 40% of the screen height, where the eye rests.
  function currentAyah() {
    const line = Math.max(80, innerHeight * 0.4);
    // Ayat share lines, so ask which ayah is under the point first; fall back to geometry.
    let best = document.elementFromPoint(innerWidth / 2, line)?.closest("[data-sura][data-aya]") || null;
    if (!best) for (const a of document.querySelectorAll("[data-sura][data-aya]")) {
      const r = a.getBoundingClientRect();
      if (r.bottom > line) { best = a; break; }
    }
    return best ? { sura: +best.dataset.sura, aya: +best.dataset.aya, name: best.dataset.suraName || "" } : null;
  }

  function reportDialog() {
    const d = dialog("report", "أبلغ عن ملاحظة", `
      <p class="b-msg">نقرأ كل ملاحظة. يُرفق معها مكان القراءة تلقائيًّا.</p>
      <label>نوع الملاحظة<select id="b-r-type">
        <option>خطأ في نص الآية أو ضبطها</option><option>خطأ في تلوين التجويد</option>
        <option>خطأ في معلومة أو إحالة</option><option>اقتراح</option><option>أخرى</option></select></label>
      <label>الموضع<input id="b-r-where" readonly></label>
      <label>النص المحدَّد<input id="b-r-quote" placeholder="حدّد كلمة في الصفحة قبل فتح هذه النافذة"></label>
      <label>الملاحظة<textarea id="b-r-msg" required maxlength="2000"></textarea></label>
      <label>بريدك إن أردت الرد (اختياري)<input id="b-r-email" type="email" autocomplete="email"></label>`,
      `<button class="btn primary" type="button" id="b-r-send">أرسل الملاحظة</button><span class="b-msg" id="b-r-status" role="status"></span>`);
    let ctx = null;
    d.prepare = () => {
      ctx = currentAyah();
      d.querySelector("#b-r-where").value = (ctx ? `سورة ${ctx.name || ctx.sura}، الآية ${ctx.aya} · ` : "") + document.title;
      d.querySelector("#b-r-quote").value = String(getSelection() || "").trim().slice(0, 300);
      d.querySelector("#b-r-status").textContent = "";
    };
    d.querySelector("#b-r-send").addEventListener("click", async () => {
      const status = d.querySelector("#b-r-status");
      const msg = d.querySelector("#b-r-msg").value.trim();
      if (!msg) { status.textContent = "اكتب الملاحظة أولًا."; return; }
      const payload = { type: d.querySelector("#b-r-type").value, message: msg, email: d.querySelector("#b-r-email").value.trim(),
        page: location.href, sura: ctx?.sura ?? null, aya: ctx?.aya ?? null, quote: d.querySelector("#b-r-quote").value };
      if (CONFIG.reportEndpoint) {
        status.textContent = "يُرسل…";
        try {
          const r = await fetch(CONFIG.reportEndpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
          if (!r.ok) throw new Error(r.status);
          status.textContent = "وصلت ملاحظتك، جزاك الله خيرًا.";
          d.querySelector("#b-r-msg").value = "";
          return;
        } catch (e) { status.textContent = "تعذّر الإرسال الآن. "; }
      }
      // No endpoint, or it failed: hand the text to the reader to send by email.
      const text = `${payload.type}\n${d.querySelector("#b-r-where").value}\n${payload.quote ? "النص: " + payload.quote + "\n" : ""}\n${payload.message}\n\n${payload.page}`;
      try { await navigator.clipboard.writeText(text); status.textContent += "نُسخت الملاحظة؛ أرسلها إلى " + (CONFIG.contactEmail || "بريد الموقع") + "."; }
      catch (e) { status.textContent += "انسخ الملاحظة وأرسلها إلى " + (CONFIG.contactEmail || "بريد الموقع") + "."; }
    });
    return d;
  }

  function contactDialog() {
    const links = Object.entries(CONFIG.social).filter(([, u]) => u)
      .map(([k, u]) => `<li><a href="${esc(u)}" target="_blank" rel="noopener">${SOCIAL_NAMES[k]}</a></li>`).join("");
    return dialog("contact", "تواصل معنا", `
      <p>للتصحيحات والاقتراحات والتعاون العلمي.</p>
      <ul class="b-list"><li><span>البريد</span><span dir="ltr" style="user-select:all">${esc(CONFIG.contactEmail)}</span></li></ul>
      ${links ? `<ul class="b-list">${links}</ul>` : ""}`,
      `<button class="btn" type="button" id="b-c-copy">${svg("mail")}انسخ البريد</button><span class="b-msg" id="b-c-status" role="status"></span>`);
  }

  function statsDialog() {
    const d = dialog("stats", "إحصاءات الموقع", `<div class="b-stats" id="b-s-grid"><div><b>…</b><span>يُحمَّل</span></div></div>
      <p class="b-msg">أرقام المحتوى محسوبة من ملفات البيانات نفسها عند بناء الموقع.</p>`);
    d.prepare = async () => {
      const grid = d.querySelector("#b-s-grid");
      try {
        const s = await (await fetch("data/stats.json")).json();
        const rows = [[s.surahs, "سورة"], [s.ayat, "آية"], [s.words, "كلمة"], [s.letters, "حرفًا"],
          [s.makki, "سورة مكية"], [s.madani, "سورة مدنية"], [`${s.riwayat_available} من ${s.riwayat_planned}`, "روايات متاحة"],
          [s.tajweed_rules, "حكمًا تجويديًّا"], [s.tajweed_marks, "موضعًا ملوّنًا"], [s.cited_works, "مصدرًا محالًا إليه"], [s.citations, "إحالة في بطاقات السور"]];
        grid.innerHTML = rows.map(([v, l]) => `<div><b>${typeof v === "number" ? arNum(v) : toArDigits(v)}</b><span>${l}</span></div>`).join("");
      } catch (e) { grid.innerHTML = '<p class="b-empty">تعذّر تحميل الإحصاءات.</p>'; }
    };
    return d;
  }

  /* ---------- my readings: last position, bookmarks, history ---------- */
  function readings() {
    const d = dialog("readings", "قراءاتي", `
      <section><h3 style="font-size:1rem;margin-bottom:8px">آخر موضع</h3><div id="b-m-last"></div></section>
      <section><h3 style="font-size:1rem;margin-bottom:8px">العلامات</h3><ul class="b-list" id="b-m-marks"></ul></section>
      <section><h3 style="font-size:1rem;margin-bottom:8px">قرأت مؤخرًا</h3><ul class="b-list" id="b-m-hist"></ul></section>
      <p class="b-msg">تُحفظ قراءاتك في هذا المتصفح وحده، ولا تُرسل إلى أحد.</p>`,
      `<button class="btn primary" type="button" id="b-m-add">${svg("book")}ضع علامة هنا</button><span class="b-msg" id="b-m-status" role="status"></span>`);
    const link = (r) => `${r.page}#a${r.sura}-${r.aya}`;
    const label = (r) => `${r.name ? "سورة " + r.name : "سورة " + r.sura}، الآية ${arNum(r.aya)}`;
    const render = () => {
      const last = store.get("last", null);
      d.querySelector("#b-m-last").innerHTML = last ? `<ul class="b-list"><li><a href="${esc(link(last))}">${esc(label(last))}</a><span class="muted">${esc(last.title || "")}</span></li></ul>` : '<p class="b-empty">ابدأ القراءة في أي صفحة، وسيُحفظ موضعك هنا.</p>';
      const marks = store.get("marks", []);
      const ul = d.querySelector("#b-m-marks");
      ul.innerHTML = marks.length ? marks.map((m, i) => `<li><a href="${esc(link(m))}">${esc(label(m))}</a><button class="b-x" type="button" data-del="${i}" aria-label="احذف العلامة">×</button></li>`).join("") : '<li class="b-empty">لا علامات بعد.</li>';
      const hist = store.get("history", []);
      d.querySelector("#b-m-hist").innerHTML = hist.length ? hist.map((h) => `<li><a href="${esc(link(h))}">${esc(label(h))}</a><span class="muted">${esc(h.title || "")}</span></li>`).join("") : '<li class="b-empty">لا شيء بعد.</li>';
    };
    d.querySelector("#b-m-marks").addEventListener("click", (e) => {
      const i = e.target.closest("[data-del]")?.dataset.del;
      if (i == null) return;
      const marks = store.get("marks", []); marks.splice(+i, 1); store.set("marks", marks); render();
    });
    d.querySelector("#b-m-add").addEventListener("click", () => {
      const c = here, status = d.querySelector("#b-m-status");
      if (!c) { status.textContent = "افتح صفحة فيها آيات لتضع علامة."; return; }
      const marks = store.get("marks", []);
      if (!marks.some((m) => m.sura === c.sura && m.aya === c.aya && m.page === pageName())) {
        marks.unshift({ ...c, page: pageName(), title: document.title, at: Date.now() });
        store.set("marks", marks.slice(0, 100));
      }
      status.textContent = "وُضعت العلامة على " + label(c) + ".";
      render();
    });
    let here = null;  // captured before the dialog covers the page
    d.prepare = () => { here = currentAyah(); d.querySelector("#b-m-status").textContent = ""; render(); };

    // Track position on pages that show ayat (they may load after this runs).
    {
      let t = null;
      addEventListener("scroll", () => {
        clearTimeout(t);
        t = setTimeout(() => {
          const c = currentAyah();
          if (!c) return;
          const rec = { ...c, page: pageName(), title: document.title, at: Date.now() };
          store.set("last", rec);
          const hist = store.get("history", []).filter((h) => !(h.page === rec.page && h.sura === rec.sura));
          hist.unshift(rec);
          store.set("history", hist.slice(0, 12));
        }, 1200);
      }, { passive: true });
    }
    return d;
  }
  const pageName = () => location.pathname.split("/").pop() || "index.html";

  /* ---------- install (PWA) ---------- */
  function install() {
    const ua = navigator.userAgent;
    const ios = /iPhone|iPad|iPod/.test(ua) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
    const mobile = ios || /Android|Mobile/.test(ua);
    const standalone = matchMedia("(display-mode: standalone)").matches || navigator.standalone === true;
    const slot = document.querySelector("[data-install-slot]");
    if (!slot || standalone) return;
    const label = mobile ? "ثبّت الموقع تطبيقًا على هاتفك" : "ثبّت الموقع تطبيقًا على حاسوبك";
    const b = el(`<div style="display:grid;gap:6px;justify-items:start"><button class="btn primary" type="button">${svg(mobile ? "phone" : "desktop")}${label}</button><a href="install.html" class="b-msg">طريقة التثبيت خطوةً خطوة</a></div>`);
    slot.append(b);
    let deferred = null;
    addEventListener("beforeinstallprompt", (e) => { e.preventDefault(); deferred = e; });
    addEventListener("appinstalled", () => { b.remove(); });
    b.querySelector("button").addEventListener("click", async () => {
      if (deferred) {
        deferred.prompt();
        await deferred.userChoice.catch(() => null);
        deferred = null;
      } else {
        location.href = "install.html#" + (ios ? "ios" : mobile ? "android" : "desktop");
      }
    });
  }

  function registerServiceWorker() {
    if ("serviceWorker" in navigator && location.protocol === "https:") {
      navigator.serviceWorker.register("sw.js").catch(() => { /* offline support is optional */ });
    }
  }

  /* ---------- boot ---------- */
  let floatEnd;
  function boot() {
    if (!document.body.dataset.noChrome) { header(); footer(); }
    const floatStart = el('<div class="b-float start"></div>');
    floatEnd = el('<div class="b-float end"></div>');
    document.body.append(floatStart, floatEnd);

    const dialogs = { report: reportDialog(), stats: statsDialog(), readings: readings() };
    if (CONFIG.contactEmail) {
      dialogs.contact = contactDialog();
      dialogs.contact.querySelector("#b-c-copy").addEventListener("click", async () => {
        const s = dialogs.contact.querySelector("#b-c-status");
        try { await navigator.clipboard.writeText(CONFIG.contactEmail); s.textContent = "نُسخ البريد."; }
        catch (e) { s.textContent = "حدّد البريد وانسخه."; }
      });
    }
    floatStart.append(el(`<button class="btn" type="button" data-open="report">${svg("note")}أبلغ عن ملاحظة</button>`));
    floatEnd.append(el(`<button class="btn" type="button" data-open="readings">${svg("book")}قراءاتي</button>`));
    backToTop();
    document.addEventListener("click", (e) => {
      const k = e.target.closest("[data-open]")?.dataset.open;
      if (!k || !dialogs[k]) return;
      dialogs[k].prepare?.();
      dialogs[k].showModal();
    });
    install();
    registerServiceWorker();
  }

  window.Basaer = { CONFIG, store, esc, arNum, toArDigits };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot); else boot();
})();
