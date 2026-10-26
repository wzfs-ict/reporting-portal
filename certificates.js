// WZFS certificates: one A4 page per award, filled from the portal's Certificates tab.
// The portal hands the list over in localStorage (wzfs_cert_handoff_v1); nothing is saved back.
(function () {
  const $ = (s) => document.querySelector(s);
  const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  let data = null;
  try { data = JSON.parse(localStorage.getItem("wzfs_cert_handoff_v1") || "null"); } catch (_) {}
  if (!data || !data.items || !data.items.length) {
    $("#certs").innerHTML = `<div class="empty"><h2>No certificates to print</h2><p style="margin-top:10px">Open the <b>Certificates</b> tab in the WZFS Reporting portal and click <b>Print certificates</b>.</p></div>`;
    $("#count").textContent = "Certificates";
    return;
  }

  const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const ordinal = (n) => n + ((n % 100 >= 11 && n % 100 <= 13) ? "th" : ["th", "st", "nd", "rd"][n % 10] || "th");
  // "English Token Award (Gold)" → title + level line; "PTL Award – October" → title + month line. "(Midterm)" never prints.
  const splitAward = (name) => {
    let t = String(name || "").replace(/\s*\(midterm\)\s*$/i, "").trim(), level = "";
    let m = t.match(/^(.*?)\s*\(([^)]+)\)\s*$/);
    if (m) { t = m[1]; level = m[2]; }
    else if ((m = t.match(/^(.*?)\s+[–-]\s+(.+)$/))) { t = m[1]; level = m[2]; }
    return { title: t, level };
  };
  // "Sion Kim (Amelia)" → SION KIM (Amelia)
  const nameHtml = (n) => { const m = String(n || "").match(/^(.*?)\s*(\(.+\))\s*$/); return m ? `${esc(m[1])} <span class="called">${esc(m[2])}</span>` : esc(n); };
  const nameSize = (n) => { const l = String(n || "").length; return l > 34 ? 24 : l > 26 ? 28 : 34; };
  const citationFor = (it) => {
    const c = data.citations || {};
    const base = c[String(it.award || "").replace(/\s*\(midterm\)\s*$/i, "")] || c[it.category] || (data.typeCitations || {})[it.type] || data.defaultCitation || "in recognition of outstanding effort and achievement";
    return `${esc(base)}<br>during the ${esc(data.semester)} Semester of School Year ${esc(data.year)}`;
  };
  const corner = `<svg class="corner %C" viewBox="0 0 100 100" aria-hidden="true">
    <polygon points="38,0 100,0 100,62" fill="var(--navy)"/><polygon points="20,0 34,0 100,66 100,80" fill="var(--blue)"/><polygon points="6,0 14,0 100,86 100,94" fill="var(--sky)"/></svg>`;
  const signs = (data.signatories || []).map((s) => `<div class="sign"><div class="sig">${s.sig ? `<img src="${esc(s.sig)}" alt="">` : ""}</div>
    <div class="who" contenteditable>${esc(s.name)}</div><div class="role" contenteditable>${esc(s.title)}</div></div>`).join("");

  const render = () => {
    const d = new Date(($("#conferred").value || new Date().toISOString().slice(0, 10)) + "T12:00:00");
    const conferred = `Conferred this ${ordinal(d.getDate())} day of ${MONTHS[d.getMonth()]} ${d.getFullYear()}.`.replace(/(\d+)(st|nd|rd|th)/, "$1<sup>$2</sup>");
    $("#certs").innerHTML = data.items.map((it) => { const a = splitAward(it.award); return `<section class="cert">
      ${corner.replace("%C", "tr")}${corner.replace("%C", "bl")}
      <div class="head"><img src="cert/wzfs-logo.png" alt=""><div><div class="school">Weihai Zhongshi<br>Foreign School</div><div class="city">Weihai City, Shandong China</div></div></div>
      <div class="body">
        <div class="this">This</div>
        <div class="rule"></div>
        <div class="award" contenteditable>${esc(a.title)}</div>
        ${a.level ? `<div class="level" contenteditable>– ${esc(a.level)} –</div>` : ""}
        <div class="rule"></div>
        <div class="presented">is presented to</div>
        <div class="name" contenteditable style="font-size:${nameSize(it.name)}pt">${nameHtml(it.name)}</div>
        <div class="citation" contenteditable>${citationFor(it)}</div>
        <div class="conferred" contenteditable>${conferred}</div>
      </div>
      <div class="signs">${signs}</div>
      <div class="accr">Accredited by: <img src="cert/wasc-logo.png" alt="WASC"></div>
    </section>`; }).join("");
  };

  $("#count").textContent = `${data.items.length} certificate${data.items.length > 1 ? "s" : ""} · ${data.semester} ${data.year}`;
  $("#conferred").value = new Date().toISOString().slice(0, 10);
  $("#conferred").onchange = render;
  $("#print").onclick = () => window.print();
  render();
})();
