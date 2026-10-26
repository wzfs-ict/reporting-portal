// WZFS certificates: one A4 page per award, in the library's house style, filled from the portal's Certificates tab.
// The portal hands the list over in localStorage (wzfs_cert_handoff_v1). The librarian's corrections are kept in
// this browser only (wzfs_cert_edits_v1) so a refresh doesn't lose them; nothing is written back to SharePoint.
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
  const nameSize = (n, big) => { const l = String(n || "").length; return big ? (l > 34 ? 24 : l > 26 ? 28 : 32) : (l > 34 ? 17 : 22); };
  const awardSize = (s) => { const l = String(s || "").length; return l > 40 ? 20 : l > 26 ? 23 : 27; };
  const chev = `<svg class="chev %C" viewBox="0 0 60 100" aria-hidden="true"><polygon points="60,0 20,50 60,100 44,100 4,50 44,0" fill="var(--chev)"/><polygon points="60,14 32,50 60,86 60,72 44,50 60,28" fill="var(--chev2)"/></svg>`;
  const tri = `<svg class="tri" style="top:0;left:0;width:210mm;height:297mm" viewBox="0 0 210 297" aria-hidden="true">
    <polygon points="0,0 92,0 80,22 0,22" fill="#1F6B3A"/><polygon points="150,190 210,160 210,297 120,297" fill="var(--green)"/><polygon points="168,215 210,195 210,297 150,297" fill="#1F6B3A"/></svg>`;
  const waves = `<svg style="position:absolute;inset:0;width:297mm;height:210mm" viewBox="0 0 297 210" aria-hidden="true">
    <path d="M0,0 H297 V8 C230,20 120,0 0,14 Z" fill="var(--navy)"/><path d="M0,14 C120,0 230,20 297,8" fill="none" stroke="var(--gold)" stroke-width=".8"/>
    <path d="M0,18 C110,5 230,24 297,13" fill="none" stroke="#C9D6EE" stroke-width="1.6"/>
    <path d="M0,210 H297 V198 C220,190 110,208 0,194 Z" fill="var(--navy)"/><path d="M0,194 C110,208 220,190 297,198" fill="none" stroke="var(--gold)" stroke-width=".8"/>
    <path d="M0,190 C110,203 220,185 297,193" fill="none" stroke="#C9D6EE" stroke-width="1.6"/>
    <path d="M12,32 v-8 h8 M285,32 v-8 h-8 M12,176 v8 h8 M285,176 v8 h-8" fill="none" stroke="var(--gold)" stroke-width=".8"/></svg>`;
  const orn = `<svg class="orn" viewBox="0 0 128 4" aria-hidden="true"><line x1="0" y1="2" x2="56" y2="2" stroke="var(--gold)" stroke-width=".6"/><line x1="72" y1="2" x2="128" y2="2" stroke="var(--gold)" stroke-width=".6"/>
    <polygon points="64,0 67,2 64,4 61,2" fill="var(--gold)"/><circle cx="58.5" cy="2" r=".8" fill="var(--gold)"/><circle cx="69.5" cy="2" r=".8" fill="var(--gold)"/></svg>`;
  const signs = (data.signatories || []).map((s) => `<div class="sign"><div class="sig">${s.sig ? `<img src="${esc(s.sig)}" alt="">` : ""}</div>
    <div class="who">${esc(s.name)}</div><div class="role">${esc(s.title)}</div></div>`).join("");
  // An editable text block; the librarian's correction (if any) replaces the generated text.
  const ed = (id, field, cls, html, style = "") => { const k = `${id}|${field}`; return `<div class="${cls}${edits[k] != null ? " edited" : ""}" contenteditable data-k="${esc(k)}" style="${style}">${edits[k] != null ? edits[k] : html}</div>`; };

  const page = (it, conferredHtml) => {
    const base = String(it.award || "").trim(), a = splitAward(base), t = tplFor(it, base);
    let title = t.title || a.title; if (it.teacherAward) title += " (Teacher Award)";
    const level = a.level, typeWord = t.type || (it.type || "").replace(/^Certificate of\s*/i, "") || "Recognition";
    const after = ((t.afterByLevel || {})[level] || t.after || "").replace("{level}", level);
    const term = `${esc(opts.semester)} Semester, School Year <span style="white-space:nowrap">${esc(opts.year)}</span>.`;
    if (t.theme === "prestige") {
      // Its own defaults instead of the portrait "is awarded to … for".
      const t = Object.assign({ intro: "This certificate hereby recognizes that" }, T[it.category] || {}, T[base] || {});
      // Landscape graduation awards: "The Award of the Chairman of the Board (Servant Leadership)" → kicker + title + quality.
      const m = !t.kicker && title.match(/^The Award of (.+)$/i);
      const kicker = t.kicker || (m ? "The Award of" : ""), big = t.title || (m ? m[1] : title);
      const lead = (level ? (t.lead || "through sustained demonstration of exemplary “{level}” is the worthy recipient of") : "is the worthy recipient of").replace("{level}", esc(level));
      return `<section class="cert land">${waves}
      <div class="head" style="top:17mm"><img src="cert/wzfs-logo.png" alt="" style="width:22mm;height:22mm"><div><div class="school">Weihai Zhongshi Foreign School</div><div class="city">Weihai City, Shandong China</div></div></div>
      <div class="main" style="top:42mm;bottom:62mm;left:30mm;right:30mm;gap:2.4mm">
        ${ed(it.id, "intro", "intro", esc(t.intro || "This certificate hereby recognizes that"))}
        ${ed(it.id, "name", "script", esc(it.name))}
        ${ed(it.id, "lead", "lead", lead, "max-width:170mm")}
        ${kicker ? ed(it.id, "kicker", "kicker", esc(kicker)) : ""}
        ${ed(it.id, "award", "ptitle", esc(big), `font-size:${big.length > 34 ? 22 : big.length > 24 ? 27 : 31}pt`)}
        <div class="term" style="font-family:'Segoe UI',Arial,sans-serif;font-size:12pt">for the School Year <span style="white-space:nowrap">${esc(opts.year)}</span>.</div>
        <div class="conferred" style="font-family:'Segoe UI',Arial,sans-serif;font-size:12pt;margin-top:-2mm">${conferredHtml}</div>
      </div>
      <div class="signs" style="top:150mm;left:50mm;right:95mm">${signs}</div>
      <div class="accr" style="top:152mm;left:auto;right:24mm">Accredited by:<img src="cert/wasc-logo.png" alt="WASC"></div>
    </section>`;
    }
    if (t.theme === "green") return `<section class="cert green">${tri}
      <div class="head" style="justify-content:flex-start;left:8mm;top:4mm;color:#fff"><img src="cert/wzfs-logo.png" alt="" style="width:16mm;height:16mm;background:#fff;border-radius:50%">
        <div class="school" style="color:#fff;font-size:15pt;text-align:left">Weihai Zhongshi<br>Foreign School</div></div>
      <div style="position:absolute;top:8mm;right:12mm;font-family:Georgia,serif;font-size:11pt;color:var(--green-dark)">Weihai City, Shandong Province, China</div>
      <div class="main">
        <div class="certword">CERTIFICATE</div>
        <div class="banner">of ${esc(typeWord)}</div>
        ${level ? ed(it.id, "level", "level", `-${esc(level)}-`) : ""}
        <div class="gap"></div>
        ${ed(it.id, "intro", "intro", esc(t.intro || "is awarded to"))}
        ${ed(it.id, "name", "name", nameHtml(it.name), `font-size:${nameSize(it.name)}pt`)}
        ${ed(it.id, "after", "after", `${esc(after)} for the ${esc(opts.semester)} Semester of School Year <span style="white-space:nowrap">${esc(opts.year)}</span>.`, "max-width:140mm")}
        <div class="conferred">${conferredHtml.replace("Conferred", "Given")}</div>
      </div>
      <div class="signs">${signs}</div>
      <div class="accr" style="left:auto;right:10mm;top:262mm;color:#fff">Accredited<img src="cert/wasc-logo.png" alt="WASC"></div>
    </section>`;
    return `<section class="cert">${chev.replace("%C", "tr")}${chev.replace("%C", "bl")}
      <div class="head"><img src="cert/wzfs-logo.png" alt=""><div><div class="school">Weihai Zhongshi Foreign School</div><div class="city">Weihai City, Shandong China</div></div></div>
      <div class="main">
        <div class="this">This</div>
        ${orn}
        <div class="certword">CERTIFICATE</div>
        <div class="of">of</div>
        <div class="typeword">${esc(typeWord)}</div>
        ${orn}
        ${ed(it.id, "intro", "intro", esc(t.intro || "is awarded to"))}
        ${ed(it.id, "name", "name", nameHtml(it.name), `font-size:${nameSize(it.name, true)}pt`)}
        ${t.lead === "" ? "" : ed(it.id, "lead", "lead", esc(t.lead || "for"), "max-width:150mm")}
        ${ed(it.id, "award", `award${t.awardStyle === "blue" ? " blue" : ""}`, esc(title), `font-size:${awardSize(title)}pt`)}
        ${level ? ed(it.id, "level", `level${t.levelStyle === "green" ? " green" : ""}`, t.levelStyle === "green" ? esc(level) : `-${esc(level)}-`) : ""}
        ${after ? ed(it.id, "after", "after", esc(after), "max-width:150mm") : ""}
        <div class="gap"></div>
        <div class="term">${term}</div>
        <div class="conferred">${conferredHtml}</div>
      </div>
      <div class="signs">${signs}</div>
      <div class="accr">Accredited by:<img src="cert/wasc-logo.png" alt="WASC"></div>
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
