// WZFS certificates: one A4 page per award, in the designer's v2 house style (Nov 2026 preset),
// filled from the portal's Certificates tab.
// The portal hands the list over in localStorage (wzfs_cert_handoff_v1). The librarian's corrections are kept in
// this browser only (wzfs_cert_edits_v1) so a refresh doesn't lose them; nothing is written back to SharePoint.
//
// Tiers (from wzfs-certificate-style-preset):
//   T1 Academic     — medallion (star=Achievement, trophy=Distinction, medal=Outstanding Student,
//                     book=Teacher Award, globe=Participation); dark-gold award title
//   T2 Recognition  — no medallion; navy title; placement as a metallic pill (1st gold, 2nd silver, 3rd bronze)
//   T3 Values       — T2 layout in green (Green Cards, Core Values, PDE)
//   T4 Graduation   — landscape; student name in Arizonia script; kicker above the title
(function () {
  const $ = (s) => document.querySelector(s);
  const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const load = (k, d) => { try { return JSON.parse(localStorage.getItem(k) || "null") || d; } catch (_) { return d; } };
  const save = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (_) {} };
  const data = load("wzfs_cert_handoff_v1", null);
  if (!data || !data.items || !data.items.length) {
    $("#certs").innerHTML = `<div class="empty"><h2>No certificates to print</h2><p style="margin-top:10px">Open the <b>Certificates</b> tab in the WZFS Reporting portal and click <b>Print certificates</b>.</p></div>`;
    return;
  }
  const edits = load("wzfs_cert_edits_v1", {});
  const opts = Object.assign({ semester: data.semester, year: data.year, conferred: new Date().toISOString().slice(0, 10) }, edits._opts || {});

  const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const ordinal = (n) => n + ((n % 100 >= 11 && n % 100 <= 13) ? "th" : ["th", "st", "nd", "rd"][n % 10] || "th");
  const T = data.templates || {};
  const tplFor = (it, base) => Object.assign({}, T.default || {}, T[it.category] || {}, T[base] || {});
  // "Outstanding Student Award (Bronze)" → title + level; "PTL Award – October" → title + month.
  // "(Midterm)" / "(Finals)" in the catalogue name stay on the title line (e.g. HONOURS (MIDTERM)).
  const splitAward = (name) => {
    let t = String(name || "").trim(), level = "", m;
    if (!/\((midterm|finals?)\)\s*$/i.test(t) && (m = t.match(/^(.*?)\s*\(([^)]+)\)\s*$/))) { t = m[1]; level = m[2]; }
    else if ((m = t.match(/^(.*?)\s+[–-]\s+(.+)$/))) { t = m[1]; level = m[2]; }
    return { title: t, level };
  };
  const nameHtml = (n) => { const m = String(n || "").match(/^(.*?)\s*(\(.+\))\s*$/); return m ? `${esc(m[1])} <span class="called">${esc(m[2])}</span>` : esc(n); };
  const nameSize = (n, big) => { const l = String(n || "").length; return big ? (l > 34 ? 21 : l > 26 ? 25 : 29) : (l > 34 ? 17 : 22); };
  const titleSize = (s) => { const l = String(s || "").length; return l > 40 ? 15 : l > 26 ? 18 : 21; };

  // Placement → metallic pill (preset: 1st/Champion=gold, 2nd=silver, 3rd=bronze, platinum=#D9DADE).
  const pillFor = (lvl) => {
    const s = String(lvl || "").toLowerCase();
    if (/platinum/.test(s)) return "platinum";
    if (/1st|champion|first|\bgold/.test(s)) return "gold";
    if (/2nd|second|silver/.test(s)) return "silver";
    if (/3rd|third|bronze/.test(s)) return "bronze";
    return null;
  };

  // Tier-1 medallion icons (inline SVG, fill inherits from CSS).
  const ICONS = {
    star: '<svg viewBox="0 0 24 24"><path d="M12 2l2.9 6.3 6.9.8-5.1 4.7 1.4 6.8-6.1-3.4-6.1 3.4 1.4-6.8L2.2 9.1l6.9-.8z"/></svg>',
    trophy: '<svg viewBox="0 0 24 24"><path d="M6 3h12v2h3v2c0 2.8-2 5-4.7 5.4A6 6 0 0 1 13 15.9V18h3v2H8v-2h3v-2.1a6 6 0 0 1-3.3-3.5C5 11.9 3 9.8 3 7V5h3V3zM5 7c0 1.4.8 2.7 2 3.3A9 9 0 0 1 7 7H5zm14 0h-2a9 9 0 0 1-.1 3.3c1.3-.6 2.1-1.9 2.1-3.3V7z"/></svg>',
    medal: '<svg viewBox="0 0 24 24"><path fill-rule="evenodd" d="M7.5 2h3.2v5.2a5.6 5.6 0 0 0-3.7 1.9L7.5 2zm9 0h-3.2v5.2a5.6 5.6 0 0 1 3.7 1.9L16.5 2zM12 8.5a5.5 5.5 0 1 1 0 11 5.5 5.5 0 0 1 0-11zm0 1.6a3.9 3.9 0 1 0 0 7.8 3.9 3.9 0 0 0 0-7.8z"/></svg>',
    book: '<svg viewBox="0 0 24 24"><path d="M12 5C10.5 3.8 8.3 3 6 3 4.8 3 3.6 3.2 2.5 3.5v13C3.6 16.2 4.8 16 6 16c2.3 0 4.5.8 6 2 1.5-1.2 3.7-2 6-2 1.2 0 2.4.2 3.5.5v-13C20.4 3.2 19.2 3 18 3c-2.3 0-4.5.8-6 2zm0 2.2v9.6c-1.3-1-3.2-1.8-5-1.8-1 0-2 .1-2.9.4V6.2c.9-.3 1.9-.4 2.9-.4 1.8 0 3.7.8 5 1.4z"/></svg>',
    globe: '<svg viewBox="0 0 24 24"><path d="M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm7.9 9h-2.9a15 15 0 0 0-1.2-5A8 8 0 0 1 19.9 11zM12 4.1c.9 1.2 1.7 3.1 1.9 6.9h-3.8c.2-3.8 1-5.7 1.9-6.9zM8.2 6a15 15 0 0 0-1.2 5H4.1A8 8 0 0 1 8.2 6zM4.1 13h2.9c.2 2.2.6 3.9 1.2 5a8 8 0 0 1-4.1-5zm5.9 0h3.8c-.2 3.8-1 5.7-1.9 6.9-.9-1.2-1.7-3.1-1.9-6.9zm5.8 5a15 15 0 0 0 1.2-5h2.9a8 8 0 0 1-4.1 5z"/></svg>',
  };

  const frame = `<div class="frame-o"></div><div class="frame-i"></div><div class="dia tl"></div><div class="dia tr"></div><div class="dia bl"></div><div class="dia br"></div>`;
  const head = `<div class="head"><img src="cert/wzfs-logo.png" alt=""><div class="school">Weihai Zhongshi Foreign School</div><div class="city">Weihai City, Shandong, China</div></div>`;
  const rule = (cls, color) => `<svg class="${cls}" viewBox="0 0 128 6" aria-hidden="true"><line x1="0" y1="3" x2="54" y2="3" stroke="${color}" stroke-width=".8"/><line x1="74" y1="3" x2="128" y2="3" stroke="${color}" stroke-width=".8"/><polygon points="64,0 67,3 64,6 61,3" fill="${color}"/><circle cx="57.5" cy="3" r=".9" fill="${color}"/><circle cx="70.5" cy="3" r=".9" fill="${color}"/></svg>`;
  const signs = (data.signatories || []).map((s) => `<div class="sign"><div class="sig">${s.sig ? `<img src="${esc(s.sig)}" alt="">` : ""}</div>
    <div class="sigline"></div><div class="who">${esc(s.name)}</div><div class="role">${esc(s.title)}</div></div>`);
  const signsRow = `<div class="signs">${signs[0] || ""}<div class="wasc"><div class="cap">Accredited by</div><img src="cert/wasc-logo.png" alt="WASC"><div class="cap">WASC</div></div>${signs[1] || ""}</div>`;
  // An editable text block; the librarian's correction (if any) replaces the generated text.
  const ed = (id, field, cls, html, style = "") => { const k = `${id}|${field}`; return `<div class="${cls}${edits[k] != null ? " edited" : ""}" contenteditable data-k="${esc(k)}" style="${style}">${edits[k] != null ? edits[k] : html}</div>`; };
  const levelHtml = (it, level) => {
    if (!level) return "";
    const p = pillFor(level);
    return p ? ed(it.id, "level", `pill ${p}`, esc(level)) : ed(it.id, "level", "level", esc(level));
  };

  const page = (it, conferredHtml) => {
    const base = String(it.award || "").trim(), a = splitAward(base), t = tplFor(it, base);
    let title = t.title || a.title; if (it.teacherAward) title += " (Teacher Award)";
    const level = a.level, typeWord = t.type || (it.type || "").replace(/^Certificate of\s*/i, "") || "Recognition";
    const after = ((t.afterByLevel || {})[level] || t.after || "").replace("{level}", level);
    const term = `${esc(opts.semester)} Semester, School Year <span style="white-space:nowrap">${esc(opts.year)}</span>.`;

    // ── Tier 4: landscape graduation awards ─────────────────────────────
    if (t.theme === "prestige") {
      const t4 = Object.assign({ intro: "This certificate hereby recognizes that" }, T[it.category] || {}, T[base] || {});
      const m = !t4.kicker && title.match(/^The Award of (.+)$/i);
      const kicker = t4.kicker || (m ? "The Award of" : ""), big = t4.title || (m ? m[1] : title);
      const lead = (level ? (t4.lead || "through sustained demonstration of exemplary “{level}” is the worthy recipient of") : "is the worthy recipient of").replace("{level}", esc(level));
      return `<section class="cert land t4">${frame}${head}
      <div class="main"><img class="wm" src="cert/crest-mark.png" alt="">
        ${ed(it.id, "intro", "intro", esc(t4.intro))}
        ${ed(it.id, "name", "script", esc(it.name))}
        ${rule("nrule", "var(--gold)")}
        ${ed(it.id, "lead", "lead", lead, "max-width:180mm")}
        ${kicker ? ed(it.id, "kicker", "kicker", esc(kicker)) : ""}
        ${ed(it.id, "award", "ptitle", esc(big), `font-size:${big.length > 34 ? 17 : big.length > 24 ? 20 : 22.5}pt`)}
        <div class="term">for the School Year <span style="white-space:nowrap">${esc(opts.year)}</span>.</div>
        <div class="conferred">${conferredHtml}</div>
      </div>
      ${signsRow}
    </section>`;
    }

    // ── Tiers 1–3 (portrait) ────────────────────────────────────────────
    // Tier: T3 = values (green); T1 = academic medallion; everything else = T2 recognition.
    const T3 = t.theme === "green" || ["Green Cards", "Core Values", "PDE"].includes(it.category);
    const medalFor = { Achievement: "star", Distinction: "trophy", "Teacher Award": "book", Participation: "globe" };
    let medal = T3 ? null : medalFor[typeWord] || null;
    if (!T3 && it.category === "Outstanding Student") medal = "medal";
    const tier = T3 ? "t3" : medal ? "t1" : "t2";
    const titleCls = medal === "book" ? "title navy" : "title";
    const accent = T3 ? "var(--greenDark)" : "var(--gold)";
    const medalHtml = medal ? `<div class="medal ${medal === "globe" ? "navy" : medal === "book" ? "book" : ""}">${ICONS[medal]}</div>` : "";

    return `<section class="cert ${tier}">${frame}${head}
      <div class="main"><img class="wm" src="cert/crest-mark.png" alt="">
        <div class="certword">CERTIFICATE</div>
        <div class="ofline">of ${esc(typeWord)}</div>
        ${medalHtml}
        ${ed(it.id, "intro", "intro", esc(t.intro || "is awarded to"))}
        ${ed(it.id, "name", "name", nameHtml(it.name), `font-size:${nameSize(it.name, true)}pt`)}
        ${rule("nrule", accent)}
        ${t.lead === "" ? "" : ed(it.id, "lead", "lead", esc(t.lead || "for"), "max-width:150mm")}
        ${ed(it.id, "award", titleCls, esc(title), `font-size:${titleSize(title)}pt`)}
        ${levelHtml(it, level)}
        ${after ? ed(it.id, "after", "after", esc(after), "max-width:150mm") : ""}
        <div class="gap"></div>
        <div class="term">${term}</div>
        <div class="conferred">${conferredHtml}</div>
      </div>
      ${signsRow}
    </section>`;
  };

  const render = () => {
    const d = new Date((opts.conferred || new Date().toISOString().slice(0, 10)) + "T12:00:00");
    const conferred = `Conferred this ${ordinal(d.getDate()).replace(/(\d+)(\D+)/, "$1<sup>$2</sup>")} day of ${MONTHS[d.getMonth()]} ${d.getFullYear()}.`;
    $("#certs").innerHTML = data.items.map((it) => page(it, conferred)).join("");
    document.querySelectorAll("[data-k]").forEach((el) => (el.oninput = () => { edits[el.dataset.k] = el.innerHTML; save("wzfs_cert_edits_v1", edits); }));
    document.title = `WZFS Certificates ${data.label || ""} ${opts.semester} ${opts.year}`.replace(/\s+/g, " ").trim();
  };
  const setOpt = (k) => (e) => { opts[k] = e.target.value; edits._opts = opts; save("wzfs_cert_edits_v1", edits); render(); };

  $("#count").textContent = `${data.items.length} certificate${data.items.length > 1 ? "s" : ""}`;
  $("#semester").value = opts.semester; $("#year").value = opts.year; $("#conferred").value = opts.conferred;
  $("#semester").onchange = setOpt("semester"); $("#year").onchange = setOpt("year"); $("#conferred").onchange = setOpt("conferred");
  $("#print").onclick = () => window.print();
  $("#reset").onclick = () => {
    const mine = data.items.map((it) => it.id + "|");
    if (!confirm("Undo all your text corrections on these certificates?")) return;
    for (const k of Object.keys(edits)) if (mine.some((p) => k.startsWith(p))) delete edits[k];
    save("wzfs_cert_edits_v1", edits); render();
  };
  render();
})();
