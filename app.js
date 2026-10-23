"use strict";
// ─────────────────────────────────────────────────────────────
//  WZFS Reporting Portal
//  Static front end. SharePoint lists are the database; Microsoft Graph is the API.
//  Privacy is enforced by SharePoint (item-level permissions), not by this page.
// ─────────────────────────────────────────────────────────────
const C = window.APP_CONFIG;
const L = C.lists;
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["User.Read", "Sites.ReadWrite.All"];
const SETUP_SCOPES = ["User.Read", "Sites.ReadWrite.All", "Sites.Manage.All"];
const SP_SCOPES = [`https://${C.siteHostname}/AllSites.FullControl`];
const redirectUri = location.origin + location.pathname.replace(/index\.html$/, "");

const S = {
  me: "", meName: "", siteId: "", siteUrl: "", lists: {},
  students: [], subjects: [], catalogue: [], settingsItem: null, settings: {},
  marks: {}, comments: [], awards: [],
  tab: "classes", cls: null, back: "classes",
  advGrade: "", ptcStudent: "", awGrade: "", awForm: {}, awSel: new Set(), awSearch: "", awAll: false,
  loadErr: null,
};

// ── Small helpers ────────────────────────────────────────────
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const lc = (s) => String(s || "").trim().toLowerCase();
const blank = (v) => v === null || v === undefined || v === "";
const P = () => ({ ...C, ...S.settings });               // config merged with saved Settings
const term = () => ({ Semester: P().semester, AcademicYear: P().academicYear });
const inTerm = (f) => f.Semester === P().semester && f.AcademicYear === P().academicYear;
// Marks and advisor comments also belong to a reporting period (Midterm / Final). Awards don't.
const termP = () => ({ ...term(), Period: P().period });
const inPeriod = (f) => inTerm(f) && (f.Period || "Final") === P().period;
const isMidterm = () => P().period === "Midterm";
const termLabel = () => `${P().semester} ${P().academicYear} · ${P().period}`;
const marksListName = (g) => L.marksPrefix + g;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
// SheetJS is large; fetch it only when a screen needs spreadsheets (retries once, for patchy connections).
let xlsxLoading = null;
function loadXlsx() {
  if (window.XLSX) return Promise.resolve();
  const once = () => new Promise((resolve, reject) => {
    const s = document.createElement("script"); s.src = "vendor/xlsx-0.18.5.full.min.js";
    s.onload = resolve; s.onerror = () => { s.remove(); reject(new Error("Couldn't load the spreadsheet tool. Check the connection and try again.")); };
    document.head.appendChild(s);
  });
  return (xlsxLoading = xlsxLoading || once().catch(() => once()).catch((e) => { xlsxLoading = null; throw e; }));
}
function notice(kind, html) { return `<div class="notice ${kind}">${html}</div>`; }
function flash(kind, msg) {
  const el = $("#flash"); if (!el) return;
  el.innerHTML = notice(kind, esc(msg));
  if (kind === "ok") setTimeout(() => { if (el.textContent.includes(msg)) el.innerHTML = ""; }, 4000);
}
function behLabel(v) { const b = P().behaviourScale.find((x) => x.value === Number(v)); return b ? `${b.value} – ${b.label}` : ""; }
function subjectDisplay(f) {
  const lvl = (f.ClassLevel || "").trim();
  if (!lvl) return f.Subject;
  return lc(lvl).includes(lc(f.Subject)) ? lvl : `${f.Subject} (${lvl})`;
}

// ── Auth + Graph ─────────────────────────────────────────────
let pca, account;
async function initAuth() {
  pca = new msal.PublicClientApplication({
    auth: { clientId: C.clientId, authority: `https://login.microsoftonline.com/${C.tenantId}`, redirectUri },
    cache: { cacheLocation: "localStorage" },
  });
  await pca.initialize();
  const res = await pca.handleRedirectPromise();
  account = (res && res.account) || pca.getAllAccounts()[0] || null;
  if (account) pca.setActiveAccount(account);
}
async function token(scopes = SCOPES) {
  try { return (await pca.acquireTokenSilent({ scopes, account })).accessToken; }
  catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      sessionStorage.setItem("rp_tab", S.tab);
      await pca.acquireTokenRedirect({ scopes, account });
      return new Promise(() => {});
    }
    throw e;
  }
}
async function graph(method, path, body, scopes) {
  if (S.viewAs && method !== "GET") throw new Error("Preview only: nothing is saved while you're viewing as someone else. Switch back to yourself to make changes.");
  const r = await fetch(path.startsWith("http") ? path : GRAPH + path, {
    method, headers: { Authorization: "Bearer " + await token(scopes), "Content-Type": "application/json" },
    body: body ? JSON.stringify(body) : undefined });
  if (!r.ok) { let m = r.statusText; try { m = (await r.json()).error.message || m; } catch (_) {} throw Object.assign(new Error(`${m} (HTTP ${r.status})`), { status: r.status }); }
  return r.status === 204 ? null : r.json();
}
async function graphAll(path) {
  let out = [], next = path;
  while (next) { const j = await graph("GET", next); out = out.concat(j.value || []); next = j["@odata.nextLink"]; }
  return out;
}
// Up to 20 requests per call; retries throttled ones. reqs: [{method, url, body}]
async function batch(reqs, scopes) {
  const out = new Array(reqs.length);
  let pending = reqs.map((r, i) => i);
  for (let attempt = 0; attempt < 4 && pending.length; attempt++) {
    if (attempt) await sleep(1500 * attempt);
    const retry = [];
    for (let i = 0; i < pending.length; i += 20) {
      const chunk = pending.slice(i, i + 20);
      const res = await graph("POST", "/$batch", { requests: chunk.map((k) => ({
        id: String(k), method: reqs[k].method, url: reqs[k].url,
        ...(reqs[k].body ? { headers: { "Content-Type": "application/json" }, body: reqs[k].body } : {}),
      })) }, scopes);
      for (const x of res.responses) {
        if ((x.status === 429 || x.status === 503) && attempt < 3) retry.push(+x.id); else out[+x.id] = x;
      }
    }
    pending = retry;
  }
  return out;
}
const batchFailures = (res) => res.filter((x) => !x || x.status >= 300);
const itemsUrl = (listName) => `/sites/${S.siteId}/lists/${S.lists[listName].id}/items`;

// SharePoint REST (only used by Setup → Permissions)
const sp = (method, path, body, headers) => spAt(S.siteUrl, method, path, body, headers);
async function spAt(siteUrl, method, path, body, headers = {}) {
  if (S.viewAs && method !== "GET") throw new Error("Preview only: nothing is saved while you're viewing as someone else.");
  const r = await fetch(`${siteUrl}/_api/${path}`, {
    method, headers: { Authorization: "Bearer " + await token(SP_SCOPES), Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=verbose", ...headers },
    body: body ? JSON.stringify(body) : undefined });
  if (!r.ok) { let m = r.statusText; try { const j = await r.json(); m = (j["odata.error"] || j.error).message.value || m; } catch (_) {} throw Object.assign(new Error(`${m} (HTTP ${r.status})`), { status: r.status }); }
  const t = await r.text(); return t ? JSON.parse(t) : null;
}

// ── Data ─────────────────────────────────────────────────────
const norm = (it) => ({ id: it.id, f: it.fields || {}, by: lc(it.createdBy?.user?.email), byName: it.createdBy?.user?.displayName || "", modified: it.lastModifiedDateTime });
async function readList(name) {
  const l = S.lists[name]; if (!l) return [];
  return (await graphAll(`/sites/${S.siteId}/lists/${l.id}/items?$expand=fields&$top=500`)).map(norm);
}
async function loadAll() {
  const site = await graph("GET", `/sites/${C.siteHostname}:${C.sitePath}?$select=id,webUrl`);
  S.siteId = site.id; S.siteUrl = site.webUrl;
  const lists = await graphAll(`/sites/${S.siteId}/lists?$select=id,displayName,webUrl`);
  S.lists = Object.fromEntries(lists.map((l) => [l.displayName, l]));
  const settings = await readList(L.settings);
  S.settingsItem = settings.find((x) => x.f.Title === "config") || null;
  try { S.settings = S.settingsItem ? JSON.parse(S.settingsItem.f.Json || "{}") : {}; } catch (_) { S.settings = {}; }
  S.approvalItem = settings.find((x) => x.f.Title === "approvals") || null;
  try { S.approvals = S.approvalItem ? JSON.parse(S.approvalItem.f.Json || "{}") : {}; } catch (_) { S.approvals = {}; }
  const [students, subjects, catalogue] = await Promise.all([readList(L.students), readList(L.subjects), readList(L.catalogue)]);
  S.students = students.sort((a, b) => (a.f.FullName || "").localeCompare(b.f.FullName || ""));
  S.subjects = subjects.sort((a, b) => (a.f.SortOrder ?? 999) - (b.f.SortOrder ?? 999) || (a.f.Title || "").localeCompare(b.f.Title || ""));
  S.catalogueAll = catalogue.sort((a, b) => (a.f.SortOrder ?? 999) - (b.f.SortOrder ?? 999));
  // Teachers see active awards for this semester (Season blank or "Both" = all year).
  S.catalogue = teacherCatalogue();
  await Promise.all([loadMarks(), loadComments(), loadAwards(), loadPhotos()]);
}
async function loadMarks(grade) {
  const grades = grade ? [grade] : P().grades;
  await Promise.all(grades.map(async (g) => { S.marks[g] = (await readList(marksListName(g))).filter((x) => inPeriod(x.f)).map((x) => ({ ...x, list: g })); }));
}
// A mark row lives in the bobae's own grade list (r.list), so it reaches their report card and advisor.
// ClassGrade is the grade of the class they sit in, e.g. a G11 bobae in a G10 class. Blank = same as the list.
const allMarks = () => Object.values(S.marks).flat();
const classGradeOf = (r) => r.f.ClassGrade || r.list;
// Student photos: one Graph call lists the library with ready-made thumbnail links (no file downloads).
async function loadPhotos() {
  S.photos = {};
  const l = S.lists[L.photos]; if (!l) return;
  try {
    const files = await graphAll(`/sites/${S.siteId}/lists/${l.id}/drive/root/children?$select=name,id&$expand=thumbnails($select=medium)&$top=999`);
    for (const f of files) { const t = f.thumbnails && f.thumbnails[0] && f.thumbnails[0].medium; if (t) S.photos[f.name.replace(/\.[^.]+$/, "")] = t.url; }
  } catch (_) { /* photos are optional; the app works without them */ }
}
// <img> for a bobae's photo, or an initial in a grey tile when there is none.
function photo(id, size = 36) {
  const st = studentById(id)?.f, url = S.photos && S.photos[id];
  const h = Math.round(size * 1.25);
  return url ? `<img class="ph" src="${esc(url)}" width="${size}" height="${h}" loading="lazy" alt="">`
    : `<span class="ph ph-none" style="width:${size}px;height:${h}px;font-size:${Math.round(size / 2.4)}px">${esc(((st && (st.CalledName || st.FullName)) || "?").slice(0, 1))}</span>`;
}
async function loadComments() { S.comments = (await readList(L.comments)).filter((x) => inPeriod(x.f)); }
async function loadAwards() { S.awards = (await readList(L.awards)).filter((x) => inTerm(x.f)); }
const rosterFor = (g) => S.students.filter((s) => s.f.Grade === g && s.f.Status !== "Withdrawn");
const studentById = (id) => S.students.find((s) => s.f.Title === id);

// ── Roles (layout only — real access comes from SharePoint) ──
const isAdmin = () => P().admins.map(lc).includes(S.me);
// Principal: reviews and corrects every grade, and gives the go-ahead for report cards (admins can too, as backup).
const isPrincipal = () => (P().principals || []).map(lc).includes(S.me);
const canApprove = () => isPrincipal() || isAdmin();
// The go-ahead for one grade: the Principal (or an admin) for any grade, a division head for their own grades.
const canApproveGrade = (g) => canApprove() || headGrades().includes(g);
const reviewGrades = () => P().grades.filter(canApproveGrade);
const isLibrarian = () => (P().librarians || []).map(lc).includes(S.me);
const isPrimary = (g) => ((P().primary || {}).grades || []).includes(g);
const isLeader = () => isAdmin() || isPrincipal() || P().leadership.map(lc).includes(S.me);
const emailsOf = (s) => String(s || "").split(/[\s,;]+/).map(lc).filter((e) => e.includes("@"));
const advisorEmails = (g) => emailsOf(P().advisors[g] && P().advisors[g].email);   // advisor + co-advisor
const advisorGrades = () => P().grades.filter((g) => advisorEmails(g).includes(S.me));
const canAdvise = (g) => isAdmin() || isPrincipal() || advisorGrades().includes(g);
// Admin team: the only people who generate report cards (Export).
const isReportTeam = () => isAdmin() || (P().reportTeam || []).map(lc).includes(S.me);
// Division heads (e.g. Head of Primary) see their own grades read-only.
const headGrades = () => (P().divisionHeads || {})[S.me] || [];
const viewGrades = () => isLeader() ? P().grades : P().grades.filter((g) => advisorGrades().includes(g) || headGrades().includes(g));
// "View as" (admins only): preview the app as someone else. Read-only; graph() refuses writes.
const realAdmin = () => P().admins.map(lc).includes(S.realMe);
// Does a Teacher text (from the Subjects list) mean me? Matches my email, or my name in any order,
// ignoring titles ("Mrs. Xochitl Tobias" = "Tobias Xochitl"). Several teachers can be listed: "A, B / C".
const nameWords = (s) => lc(s).replace(/\b(mr|mrs|ms|miss|dr)\b\.?/g, " ").split(/[^a-zÀ-￿]+/).filter((w) => w.length > 1);
function isMyName(text) {
  if (!text) return true;
  if (emailsOf(text).includes(S.me)) return true;
  const mine = new Set([...nameWords(S.meName), ...nameWords(nameOf(S.me))]);
  return String(text).split(/[,;\/&]|\band\b/).some((part) => { const w = nameWords(part); return w.length > 0 && w.every((x) => mine.has(x)); });
}
const nameOf = (e) => (P().names || {})[e] || (Object.values(P().advisors || {}).find((a) => emailsOf(a.email)[0] === e)?.name) || e;
const advisorName = (g) => (P().advisors[g] && P().advisors[g].name) || "";

// ── Shell ────────────────────────────────────────────────────
function tabsFor() {
  if (S.viewAs === "parent") return [["parent", "Parent view"]];
  if (S.viewAs === "student" || (S.studentMode && !S.viewAs)) return [["mygoals", "My goals"]];
  // Admin team members who don't teach (Ryan, Peter, Keating…) only need report cards.
  const teaches = isAdmin() || (P().staff || []).map(lc).includes(S.me) || advisorGrades().length > 0;
  if (isLibrarian() && !teaches && !isReportTeam()) return [["certs", "Certificates"], ["help", "Help"]];
  if (isReportTeam() && !teaches) return [["export", "Report cards"], ["certs", "Certificates"], ...(viewGrades().length ? [["advisor", "Advisor view"]] : []), ["help", "Help"]];
  const t = [["classes", "My classes"], ["awards", "Awards"]];
  if (reviewGrades().length) t.unshift(["review", canApprove() ? "Principal review" : `Review (${reviewGrades()[0]}–${reviewGrades().slice(-1)[0]})`]);
  if (isLibrarian() || isAdmin()) t.push(["certs", "Certificates"]);
  if (viewGrades().length) t.push(["advisor", headGrades().length && !advisorGrades().length && !isLeader() ? "Division view" : "Advisor view"]);
  if (isReportTeam()) t.push(["export", "Report cards"]);
  if (isAdmin()) t.push(["setup", "Setup"]);
  t.push(["pde", "PDE (in progress)"], ["help", "Help"]);
  return t;
}
function renderShell() {
  $("#schoolName").textContent = C.schoolName;
  $("#appName").textContent = C.appName;
  $("#footer").textContent = `Mr Glen Joshua (${P().semester} ${P().academicYear.slice(0, 4)}) · ${C.version} · Uses Zhongshi Microsoft SharePoint as database`;
  if (!account) { $("#who").textContent = ""; $("#authBtn").innerHTML = ""; $("#tabs").innerHTML = ""; return; }
  $("#who").textContent = account.name || S.me;
  $("#authBtn").innerHTML = `${realAdmin() && !S.loadErr ? viewAsMenu() : ""}<button class="btn small" id="signOut">Sign out</button>`;
  $("#signOut").onclick = () => pca.logoutRedirect({ account });
  const va = $("#viewAs");
  if (va) va.onchange = () => {
    S.viewAs = va.value; S.me = ["parent", "student"].includes(va.value) ? va.value + "@preview" : (va.value || S.realMe);
    S.meName = { parent: "Parent (preview)", student: "a bobae (preview)" }[va.value] || nameOf(S.me);
    S.tab = tabsFor()[0][0]; S.awSel.clear(); render(); window.scrollTo(0, 0);
  };
  if (S.loadErr) { $("#tabs").innerHTML = ""; return; }
  const active = { newclass: "classes", entry: S.back, ptc: "advisor" }[S.tab] || S.tab;
  $("#tabs").innerHTML = tabsFor().map(([k, n]) => `<button data-tab="${k}" ${k === active ? 'aria-current="page"' : ""}>${n}</button>`).join("");
  $$("#tabs button").forEach((b) => (b.onclick = () => go(b.dataset.tab)));
}
function go(tab) { S.tab = tab; render(); window.scrollTo(0, 0); }
function viewAsMenu() {
  const p = P(), opt = (e, label) => `<option value="${esc(e)}" ${S.viewAs === e ? "selected" : ""}>${esc(label || nameOf(e))}</option>`;
  const advisorOpts = p.grades.flatMap((g) => advisorEmails(g).map((e, i) => opt(e, `${g} ${i ? "co-advisor" : "advisor"}: ${nameOf(e)}`)));
  const seen = new Set(p.grades.flatMap(advisorEmails));
  const teachers = (p.staff || []).filter((e) => !seen.has(lc(e))).sort((a, b) => nameOf(a).localeCompare(nameOf(b)));
  return `<select id="viewAs" class="btn small" style="max-width:230px;margin-right:6px" title="Preview the app as someone else (read-only)">
    <option value="">View as: myself</option>
    <optgroup label="Admin team (report cards)">${(p.reportTeam || []).map((e) => opt(e)).join("")}</optgroup>
    <optgroup label="Division heads">${Object.entries(p.divisionHeads || {}).map(([e, gs]) => opt(e, `${nameOf(e)} (${gs[0]}–${gs[gs.length - 1]})`)).join("")}</optgroup>
    <optgroup label="Leadership">${p.leadership.map((e) => opt(e)).join("")}</optgroup>
    ${(p.librarians || []).length ? `<optgroup label="Librarian (certificates)">${p.librarians.map((e) => opt(e)).join("")}</optgroup>` : ""}
    <optgroup label="Class advisors">${advisorOpts.join("")}</optgroup>
    <optgroup label="Subject teachers">${teachers.map((e) => opt(e)).join("")}</optgroup>
    <optgroup label="Parents and bobaes">${opt("parent", "A parent (what they receive)")}${opt("student", "A bobae (My goals)")}</optgroup>
  </select>`;
}
function render() {
  renderShell();
  const main = $("#main");
  if (!account) return signInScreen(main);
  if (S.loadErr) return errorScreen(main);
  const views = { review: viewReview, certs: viewCerts, classes: viewClasses, newclass: viewNewClass, entry: viewEntry, advisor: viewAdvisor, ptc: viewPtc, awards: viewAwards, export: viewExport, setup: viewSetup, pde: viewPde, help: viewHelp, parent: viewParent, mygoals: viewMyGoals };
  if (!tabsFor().some(([k]) => k === ({ newclass: "classes", entry: S.back, ptc: "advisor" }[S.tab] || S.tab))) S.tab = tabsFor()[0][0];
  (views[S.tab] || viewClasses)(main);
  if (S.viewAs) main.insertAdjacentHTML("afterbegin", notice("warn", `<b>Preview:</b> you're seeing the app as <b>${esc(S.meName)}</b>. Read-only: nothing is saved. Their real access is set in SharePoint, so a few things may look slightly different for them. <a href="#" id="endPreview">Back to myself</a>`)
);
  const ep = $("#endPreview"); if (ep) ep.onclick = (e) => { e.preventDefault(); const v = $("#viewAs"); v.value = ""; v.onchange(); };
}

function signInScreen(main) {
  main.innerHTML = `
  <div class="center" style="max-width:560px;margin:0 auto">
    <img src="logo.png" alt="WZFS crest" width="96" height="96" style="border-radius:50%;background:#fff">
    <h1 style="margin-top:12px">${esc(C.appName)}</h1>
    <p class="lede">Enter marks, behaviour, remarks and awards for your own classes. Report cards are built from this.</p>
    <p><button class="btn primary" id="signIn">Sign in with your school Microsoft account</button></p>
    <div class="card" style="text-align:left">
      <h3>How it works</h3>
      <ol class="steps small">
        <li><b>My classes</b>: add the grade and subject you teach. Your class list appears from the school roster.</li>
        <li>Type each student's <b>score</b>, pick a <b>behaviour</b> rating and add a short <b>remark</b>. It saves as you go.</li>
        <li>Press <b>Submit to advisor</b> when the class is complete.</li>
        <li><b>Awards</b>: tap students' names, pick the award and save. No typing names.</li>
      </ol>
      <p class="small muted">Only you can see and change what you enter. Works on campus without a VPN.</p>
    </div>
  </div>`;
  $("#signIn").onclick = () => pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
}
function errorScreen(main) {
  const e = S.loadErr;
  const noAccess = [401, 403, 404].includes(e.status);
  main.innerHTML = `
    <h1>Can't open the reporting data</h1>
    ${notice("error", noAccess
      ? `Signed in as <b>${esc(S.me)}</b>, but this account doesn't have access to the WZFS Reporting site yet. Please ask the ICT coordinator to add you.`
      : esc(e.message))}
    <p><button class="btn" id="retry">Try again</button> <button class="btn" id="switch">Use a different account</button></p>`;
  $("#retry").onclick = () => start();
  $("#switch").onclick = () => pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
}

// ── My classes ───────────────────────────────────────────────
function groupClasses(rows) {
  const m = new Map();
  for (const r of rows) {
    const k = `${classGradeOf(r)}|${r.f.Subject}|${r.f.ClassLevel || ""}|${r.by}`;
    if (!m.has(k)) m.set(k, { grade: classGradeOf(r), subject: r.f.Subject, level: r.f.ClassLevel || "", by: r.by, byName: r.byName || r.f.TeacherName, rows: [] });
    m.get(k).rows.push(r);
  }
  return [...m.values()].map((c) => ({ ...c,
    total: c.rows.length,
    guests: c.rows.filter((r) => r.list !== c.grade).length,
    done: c.rows.filter((r) => !blank(r.f.Score) && !blank(r.f.Behaviour)).length,
    submitted: c.rows.length > 0 && c.rows.every((r) => r.f.Submitted === true),
  }));
}
function statusBadge(c) {
  if (c.submitted) return `<span class="badge ok">Submitted</span>`;
  if (c.done === 0) return `<span class="badge">Not started</span>`;
  return `<span class="badge warn">In progress · ${c.done} of ${c.total}</span>`;
}
function viewClasses(main) {
  const order = (g) => P().grades.indexOf(g);
  const mine = groupClasses(allMarks().filter((r) => r.by === S.me)).sort((a, b) => order(a.grade) - order(b.grade) || a.subject.localeCompare(b.subject));
  const hasLists = P().grades.some((g) => S.lists[marksListName(g)]);
  main.innerHTML = `
    <h1>My classes</h1>
    <p class="lede"><span class="badge info">${esc(termLabel())}</span> Each class is one grade and one subject you teach.${isMidterm() ? " Mid-term marks go to your grade advisors for the parent-teacher conferences." : ""}</p>
    <div id="flash"></div>
    ${!hasLists ? notice("warn", isAdmin() ? `The SharePoint lists haven't been created yet. Go to <a href="#" id="toSetup">Setup</a>.` : "The reporting lists aren't set up yet. Please check with the ICT coordinator.") : ""}
    <div class="class-grid">
      ${mine.map((c, i) => `
        <button class="class-card" data-i="${i}">
          <div class="t">${esc(c.grade)} · ${esc(subjectDisplay({ Subject: c.subject, ClassLevel: c.level }))}</div>
          <div class="bar"><i style="width:${c.total ? Math.round((100 * c.done) / c.total) : 0}%"></i></div>
          <div class="small muted">${c.total} students${c.guests ? ` (${c.guests} from other grades)` : ""} &nbsp; ${statusBadge(c)}</div>
        </button>`).join("")}
      ${hasLists ? `<button class="class-card" id="addClass" style="border-style:dashed;text-align:center"><div class="t">+ Add a class</div><div class="small muted">Pick a grade and subject</div></button>` : ""}
    </div>
    ${mine.length === 0 && hasLists ? `<p class="muted small" style="margin-top:14px">You haven't added any classes this semester yet.</p>` : ""}`;
  $$(".class-card[data-i]").forEach((b) => (b.onclick = () => { const c = mine[+b.dataset.i]; openClass(c.grade, c.subject, c.level, c.by, "classes"); }));
  const add = $("#addClass"); if (add) add.onclick = () => go("newclass");
  const ts = $("#toSetup"); if (ts) ts.onclick = (e) => { e.preventDefault(); go("setup"); };
}
function openClass(grade, subject, level, by, back) { S.cls = { grade, subject, level, by }; S.back = back; go("entry"); }

// Pick bobaes from other grades into a class (e.g. Ryan and Blair, G11, sitting in a G10 class).
// `chosen` is a Set of Student IDs kept across grade switches; `exclude` are IDs already in the class.
function otherGradePicker(box, classGrade, exclude, chosen, onChange) {
  const others = P().grades.filter((g) => g !== classGrade && S.lists[marksListName(g)]);
  const near = others.filter((g) => Math.abs(P().grades.indexOf(g) - P().grades.indexOf(classGrade)) <= 2);
  let pick = box.dataset.pick && others.includes(box.dataset.pick) ? box.dataset.pick : (near[0] || others[0] || "");
  const draw = () => {
    box.dataset.pick = pick;
    const list = pick ? rosterFor(pick).filter((s) => !exclude.has(s.f.Title)) : [];
    const picked = [...chosen].map(studentById).filter(Boolean);
    box.innerHTML = `
      <p class="small muted" style="margin:0 0 8px">For bobaes who belong to another grade but sit in this ${esc(classGrade)} class. Their marks go on <b>their own</b> grade's report card and to their own advisor.</p>
      <div class="row" style="margin-bottom:8px"><label style="max-width:160px"><span>From grade</span><select class="ogGrade">${others.map((g) => `<option ${g === pick ? "selected" : ""}>${esc(g)}</option>`).join("")}</select></label></div>
      ${list.length ? `<div class="grid">${list.map((s) => `<label style="color:var(--text)"><input type="checkbox" class="ogStu" value="${esc(s.f.Title)}" ${chosen.has(s.f.Title) ? "checked" : ""}> ${esc(s.f.CalledName || s.f.FullName)} <span class="muted small">${esc(s.f.FullName)}</span></label>`).join("")}</div>`
        : `<p class="small muted">Nobody to add from ${esc(pick || "other grades")}.</p>`}
      ${picked.length ? `<p class="small" style="margin-top:8px"><b>Added from other grades:</b> ${picked.map((s) => `${esc(s.f.CalledName || s.f.FullName)} <span class="badge">${esc(s.f.Grade)}</span>`).join(" ")}</p>` : ""}`;
    $(".ogGrade", box).onchange = (e) => { pick = e.target.value; draw(); };
    $$(".ogStu", box).forEach((c) => (c.onchange = () => { c.checked ? chosen.add(c.value) : chosen.delete(c.value); draw(); onChange(); }));
  };
  draw();
}
function viewNewClass(main) {
  const grades = P().grades.filter((g) => S.lists[marksListName(g)]);
  main.innerHTML = `
    <h1>Add a class</h1>
    <p class="lede">Choose the grade and subject. Then untick any students who don't take it.</p>
    <div id="flash"></div>
    <div class="card"><div class="grid">
      <label><span>Grade</span><select id="ncGrade"><option value="">Choose…</option>${grades.map((g) => `<option>${esc(g)}</option>`).join("")}</select></label>
      <label><span>Subject</span><select id="ncSubject" disabled><option value="">Choose a grade first</option></select></label>
      <label class="wide" id="ncLevelPick" hidden><span>Class level (this subject is taught in groups; pick yours)</span><select id="ncLevelSel"></select></label>
      <label class="wide" id="ncLevelWrap"><span>Class level or set (optional). Use only when the grade is split into groups, e.g. <i>Chinese Intermediate A</i>. Leave blank if you teach the whole grade.</span><input id="ncLevel" maxlength="60"></label>
    </div></div>
    <div id="ncRoster"></div>
    <div class="sticky-actions">
      <button class="btn primary" id="ncCreate" disabled>Create class</button>
      <button class="btn" id="ncCancel">Cancel</button>
      <span class="small muted" id="ncCount"></span>
    </div>`;
  const gSel = $("#ncGrade"), sSel = $("#ncSubject");
  const guests = new Set();
  let updCreate = () => {};
  const refresh = () => {
    const g = gSel.value;
    const roster = g ? rosterFor(g) : [];
    for (const id of [...guests]) if (studentById(id)?.f.Grade === g) guests.delete(id);   // now part of the class grade itself
    $("#ncRoster").innerHTML = g ? `
      <div class="card">
        <div class="row" style="justify-content:space-between;align-items:center;margin-bottom:8px">
          <h3>Students in ${esc(g)} (${roster.length})</h3>
          <span><button class="btn small" id="ncAll">Select all</button> <button class="btn small" id="ncNone">Select none</button></span>
        </div>
        ${roster.length ? `<div class="grid">${roster.map((s) => `
          <label style="color:var(--text)"><input type="checkbox" class="ncStu" value="${esc(s.f.Title)}" checked> ${esc(s.f.CalledName || s.f.FullName)} <span class="muted small">${esc(s.f.FullName)}</span></label>`).join("")}</div>`
          : notice("warn", "No students found for this grade in the roster.")}
      </div>
      <details class="card" ${guests.size ? "open" : ""}><summary>Add bobaes from other grades${guests.size ? ` (${guests.size} added)` : ""}</summary><div id="ncOther" style="margin-top:10px"></div></details>` : "";
    const upd = () => { const n = $$(".ncStu:checked").length + guests.size; $("#ncCount").textContent = g ? `${n} students selected${guests.size ? `, ${guests.size} from other grades` : ""}` : "";
      const levelOk = $("#ncLevelPick").hidden || ($("#ncLevelSel").value && ($("#ncLevelSel").value !== "__other" || $("#ncLevel").value.trim()));
      $("#ncCreate").disabled = !(g && sSel.value && n && levelOk); };
    updCreate = upd;
    if (g) otherGradePicker($("#ncOther"), g, new Set(roster.map((s) => s.f.Title)), guests, upd);
    $$(".ncStu").forEach((c) => (c.onchange = upd));
    const a = $("#ncAll"), z = $("#ncNone");
    if (a) a.onclick = () => { $$(".ncStu").forEach((c) => (c.checked = true)); upd(); };
    if (z) z.onclick = () => { $$(".ncStu").forEach((c) => (c.checked = false)); upd(); };
    upd();
  };
  gSel.onchange = () => {
    const subs = S.subjects.filter((s) => s.f.Grade === gSel.value && s.f.PrintsOnCard !== false);
    sSel.disabled = !gSel.value;
    sSel.innerHTML = `<option value="">${subs.length ? "Choose…" : "No subjects set up for this grade"}</option>` + subs.map((s) => `<option value="${esc(s.f.Title)}">${esc(s.f.Title)}</option>`).join("");
    showLevels();
    refresh();
  };
  // Subjects taught in groups (config classLevels): offer the official level names, plus "Other…".
  const lvSel = $("#ncLevelSel"), lvIn = $("#ncLevel");
  const showLevels = () => {
    const groups = (C.classLevels?.[sSel.value] || []).filter((x) => x.grades.includes(gSel.value));
    $("#ncLevelPick").hidden = !groups.length;
    if (groups.length) {
      lvSel.innerHTML = `<option value="">Choose…</option>` + groups.map((x) => `<optgroup label="${esc(x.label)}">${x.levels.map((l) => `<option>${esc(l)}</option>`).join("")}</optgroup>`).join("")
        + `<option value="__other">Other… (type it)</option><option value="__none">Whole grade, no level</option>`;
      lvIn.value = "";
    }
    $("#ncLevelWrap").hidden = groups.length > 0;
  };
  lvSel.onchange = () => {
    const v = lvSel.value;
    $("#ncLevelWrap").hidden = v !== "__other";
    lvIn.value = v.startsWith("__") ? "" : v;
    if (v === "__other") lvIn.focus();
    updCreate();
  };
  lvIn.oninput = () => updCreate();
  sSel.onchange = () => { showLevels(); refresh(); };
  $("#ncCancel").onclick = () => go("classes");
  $("#ncCreate").onclick = async () => {
    const g = gSel.value, subject = sSel.value, level = $("#ncLevel").value.trim();
    const ids = [...$$(".ncStu:checked").map((c) => c.value), ...guests];
    const existing = allMarks().filter((r) => r.by === S.me && classGradeOf(r) === g && r.f.Subject === subject && (r.f.ClassLevel || "") === level);
    if (existing.length) { flash("warn", "You already have this class. Opening it."); return openClass(g, subject, level, S.me, "classes"); }
    const subj = S.subjects.find((s) => s.f.Grade === g && s.f.Title === subject);
    // Guard against picking someone else's subject by mistake (a second class would block the grade's report cards).
    const others = [...new Set(allMarks().filter((r) => r.by !== S.me && classGradeOf(r) === g && r.f.Subject === subject
      && (!level || !r.f.ClassLevel || lc(r.f.ClassLevel) === lc(level))).map((r) => r.byName || r.f.TeacherName))].filter(Boolean);
    const listed = subj && subj.f.Teacher && !isMyName(subj.f.Teacher) ? subj.f.Teacher : "";
    if ((others.length || listed) && !confirm(`${others.length ? `${others.join(", ")} already ${others.length > 1 ? "have" : "has"} a ${g} ${subject}${level ? ` (${level})` : ""} class.` : `The subject list says ${g} ${subject} is taught by ${listed}.`}\n\n`
      + `Continue only if you really teach this subject${level ? "" : " (or a separate group: then pick its class level first)"}. Otherwise press Cancel and choose your own subject.`)) return;
    $("#ncCreate").disabled = true; $("#ncCreate").textContent = "Creating…";
    try {
      await createRows(g, subject, subj ? subj.f.SubjectType : "I", level, ids);
      openClass(g, subject, level, S.me, "classes");
    } catch (e) { flash("error", e.message); $("#ncCreate").disabled = false; $("#ncCreate").textContent = "Create class"; }
  };
}
// g = the class grade. Each bobae's row is written to their OWN grade's list, tagged with ClassGrade = g.
async function createRows(g, subject, type, level, ids) {
  const lists = new Set();
  const reqs = ids.map((id) => {
    const s = studentById(id).f;
    if (!S.lists[marksListName(s.Grade)]) throw new Error(`There is no "${marksListName(s.Grade)}" list for ${s.FullName}. Ask ICT to create it on the Setup tab.`);
    lists.add(s.Grade);
    return { method: "POST", url: itemsUrl(marksListName(s.Grade)), body: { fields: {
      Title: `${s.CalledName || s.FullName} – ${subject}`, StudentID: id, FullName: s.FullName, CalledName: s.CalledName || "",
      Subject: subject, SubjectType: type || "I", ClassLevel: level, ClassGrade: g, ...termP(), Submitted: false, Concern: false, TeacherName: S.meName } } };
  });
  const res = await batch(reqs);
  const bad = batchFailures(res);
  await Promise.all([...lists].map((x) => loadMarks(x)));
  if (bad.length) throw new Error(`${bad.length} of ${reqs.length} students could not be added (${bad[0]?.body?.error?.message || "unknown error"}). Open the class and use "Add students" to retry.`);
}

// ── Marks entry ──────────────────────────────────────────────
function viewEntry(main) {
  const { grade, subject, level, by } = S.cls;   // grade = the class grade
  const rows = allMarks().filter((r) => r.by === by && classGradeOf(r) === grade && r.f.Subject === subject && (r.f.ClassLevel || "") === level)
    .sort((a, b) => (a.list !== grade) - (b.list !== grade) || (a.f.FullName || "").localeCompare(b.f.FullName || ""));
  const mine = by === S.me;
  const submitted = rows.length > 0 && rows.every((r) => r.f.Submitted === true);
  const advisesSome = rows.some((r) => canAdvise(r.list)) || canAdvise(grade);
  const supervisor = advisesSome && !mine;
  // Per row: the teacher (until submitted), an admin, or the advisor of that bobae's own grade can edit.
  const rowEditable = (r) => (mine && !submitted) || isAdmin() || (!mine && canAdvise(r.list));
  const editable = rows.some(rowEditable) || (mine && !submitted);
  const teacher = rows[0] ? (rows[0].byName || rows[0].f.TeacherName) : "";
  const scale = P().behaviourScale;
  const inClass = new Set(rows.map((r) => r.f.StudentID));
  const missingIds = rosterFor(grade).filter((s) => !inClass.has(s.f.Title));
  const listUrl = (r) => itemsUrl(marksListName(r.list));
  main.innerHTML = `
    <p><a href="#" id="back">← Back</a></p>
    <h1>${esc(grade)} · ${esc(subjectDisplay({ Subject: subject, ClassLevel: level }))}</h1>
    <p class="lede"><span class="badge info">${esc(termLabel())}</span> ${mine ? "Your class" : "Teacher: " + esc(teacher)} · ${rows.length} students · ${submitted ? `<span class="badge ok">Submitted</span>` : `<span class="badge warn">Not submitted</span>`}</p>
    <div id="flash"></div>
    ${submitted && mine && !isAdmin() ? notice("ok", "Submitted to the advisor. To change something, ask your grade advisor to reopen it.") : ""}
    ${supervisor ? notice("warn", "You are editing another teacher's class as advisor. Changes are logged in SharePoint version history.") : ""}
    <div class="legend"><b>Behaviour:</b> ${scale.map((b) => `<span>${b.value} = ${esc(b.label)}</span>`).join("")}<span>· Score 0–100</span><span>· ⚑ = the advisor should raise this at the parent meeting</span></div>
    <div class="table-wrap"><table class="entry">
      <thead><tr><th>#</th><th>Student</th><th class="num">Score</th><th>Behaviour</th><th>Remark (optional, max ${P().remarkMaxLength})</th><th class="center" title="Flag a concern for the advisor">⚑ Flag</th><th></th>${editable && (mine || isAdmin()) ? "<th></th>" : ""}</tr></thead>
      <tbody>
      ${rows.map((r, i) => `
        <tr data-id="${r.id}">
          <td class="muted">${i + 1}</td>
          <td><div class="who-cell">${photo(r.f.StudentID, 32)}<div><b>${esc(r.f.CalledName || r.f.FullName)}</b>${r.list !== grade ? ` <span class="badge info" title="Belongs to ${esc(r.list)}; this mark goes on their ${esc(r.list)} report card">${esc(r.list)}</span>` : ""}<div class="small muted">${esc(r.f.FullName)} · ${esc(r.f.StudentID)}</div></div></div></td>
          <td class="num"><input class="score" inputmode="decimal" autocomplete="off" value="${blank(r.f.Score) ? "" : esc(r.f.Score)}" ${rowEditable(r) ? "" : "disabled"} aria-label="Score for ${esc(r.f.FullName)}"></td>
          <td><select class="beh" ${rowEditable(r) ? "" : "disabled"} aria-label="Behaviour for ${esc(r.f.FullName)}"><option value=""></option>${scale.map((b) => `<option value="${b.value}" ${Number(r.f.Behaviour) === b.value ? "selected" : ""}>${b.value} – ${esc(b.label)}</option>`).join("")}</select></td>
          <td><textarea class="remark" maxlength="${P().remarkMaxLength}" ${rowEditable(r) ? "" : "disabled"} aria-label="Remark for ${esc(r.f.FullName)}">${esc(r.f.Remark || "")}</textarea></td>
          <td class="center"><input type="checkbox" class="concern" ${r.f.Concern ? "checked" : ""} ${rowEditable(r) ? "" : "disabled"} aria-label="Flag a concern about ${esc(r.f.FullName)} for the advisor"></td>
          <td class="rowstate"></td>
          ${editable && (mine || isAdmin()) ? `<td>${rowEditable(r) ? `<button class="btn small danger rm" title="Remove this student from the class">✕</button>` : ""}</td>` : ""}
        </tr>`).join("")}
      </tbody></table></div>
    ${editable && (mine || isAdmin()) ? `
      <details class="card"><summary>Add students${missingIds.length ? ` (${missingIds.length} in ${esc(grade)} not in this class)` : ""}, or bobaes from other grades</summary>
        ${missingIds.length ? `<div class="grid" style="margin-top:10px">${missingIds.map((s) => `<label style="color:var(--text)"><input type="checkbox" class="addStu" value="${esc(s.f.Title)}"> ${esc(s.f.CalledName || s.f.FullName)} <span class="muted small">${esc(s.f.FullName)}</span></label>`).join("")}</div>` : ""}
        <h3 style="margin-top:14px">From other grades</h3><div id="addOther"></div>
        <p><button class="btn" id="addStudents">Add selected</button></p>
      </details>` : ""}
    <div class="sticky-actions">
      ${mine && !submitted ? `<button class="btn primary" id="submit">Submit to advisor</button>` : ""}
      ${submitted && advisesSome ? `<button class="btn" id="reopen">Reopen for the teacher</button>` : ""}
      ${rows.length && ((mine && !submitted) || isAdmin()) ? `<button class="btn danger" id="delClass" style="margin-left:auto">Delete this class</button>` : ""}
      <span class="small muted" id="progress"></span>
    </div>`;
  $("#back").onclick = (e) => { e.preventDefault(); go(S.back); };
  const progress = () => {
    const n = rows.filter((r) => !blank(r.f.Score) && !blank(r.f.Behaviour)).length;
    $("#progress").textContent = `${n} of ${rows.length} complete · saves automatically`;
  };
  progress();
  for (const tr of $$("tbody tr[data-id]", main)) {
    const r = rows.find((x) => x.id === tr.dataset.id);
    const state = $(".rowstate", tr);
    const save = async (fields) => {
      state.className = "rowstate"; state.textContent = "Saving…";
      try {
        await graph("PATCH", `${listUrl(r)}/${r.id}/fields`, fields);
        Object.assign(r.f, fields); state.textContent = "Saved ✓"; progress();
      } catch (e) { state.className = "rowstate err"; state.textContent = "Not saved"; state.title = e.message; flash("error", `Couldn't save ${r.f.FullName}: ${e.message}`); }
    };
    const score = $(".score", tr), beh = $(".beh", tr), rem = $(".remark", tr);
    score.onchange = () => {
      const t = score.value.trim();
      if (t === "") { score.classList.remove("invalid"); return save({ Score: null }); }
      const n = Number(t);
      if (!Number.isFinite(n) || n < 0 || n > 100) { score.classList.add("invalid"); state.className = "rowstate err"; state.textContent = "0–100"; return; }
      score.classList.remove("invalid"); save({ Score: Math.round(n * 10) / 10 });
    };
    score.onkeydown = (e) => {
      if (e.key === "Enter" || e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault();
        const sib = e.key === "ArrowUp" ? tr.previousElementSibling : tr.nextElementSibling;
        if (sib) $(".score", sib).focus();
      }
    };
    beh.onchange = () => save({ Behaviour: beh.value ? Number(beh.value) : null });
    rem.onchange = () => save({ Remark: rem.value.trim() });
    const con = $(".concern", tr);
    con.onchange = () => save({ Concern: con.checked });
    const rm = $(".rm", tr);
    if (rm) rm.onclick = async () => {
      if (!confirm(`Remove ${r.f.FullName} from this class? Their score and remark for this subject will be deleted.`)) return;
      try { await graph("DELETE", `${listUrl(r)}/${r.id}`); await loadMarks(r.list); render(); }
      catch (e) { flash("error", e.message); }
    };
  }
  const sub = $("#submit");
  if (sub) sub.onclick = async () => {
    const missing = rows.filter((r) => blank(r.f.Score) || blank(r.f.Behaviour));
    if (missing.length && !confirm(`${missing.length} student(s) still have no score or behaviour:\n${missing.slice(0, 8).map((r) => "• " + r.f.FullName).join("\n")}${missing.length > 8 ? "\n…" : ""}\n\nSubmit anyway?`)) return;
    if (!missing.length && !confirm("Submit this class to the advisor? You won't be able to edit it afterwards unless the advisor reopens it.")) return;
    await setSubmitted(rows, true);
  };
  const ro = $("#reopen");
  if (ro) ro.onclick = () => setSubmitted(rows.filter((r) => isAdmin() || canAdvise(r.list)), false);
  // Created by mistake: the teacher (before submitting) or an admin removes the whole class.
  // Deleted rows go to the SharePoint recycle bin, so ICT can restore them if needed.
  const dc = $("#delClass");
  if (dc) dc.onclick = async () => {
    const filled = rows.filter((r) => !blank(r.f.Score) || (r.f.Remark || "").trim()).length;
    if (!confirm(`Delete ${grade} ${subjectDisplay({ Subject: subject, ClassLevel: level })}${mine ? "" : ` (${teacher})`}?\n\nThis removes all ${rows.length} students from the class${filled ? `, including ${filled} entered mark${filled > 1 ? "s" : ""} and remarks` : ""}. ICT can restore it from the SharePoint recycle bin if needed.`)) return;
    dc.disabled = true; dc.textContent = "Deleting…";
    const res = await batch(rows.map((r) => ({ method: "DELETE", url: `${listUrl(r)}/${r.id}` })));
    const bad = batchFailures(res);
    await Promise.all([...new Set(rows.map((r) => r.list))].map((x) => loadMarks(x)));
    if (bad.length) { render(); return flash("error", `${bad.length} of ${rows.length} rows could not be deleted: ${bad[0]?.body?.error?.message || ""}`); }
    go(S.back || "classes"); flash("ok", `Deleted ${grade} ${subject}.`);
  };
  const guests = new Set();
  const ao = $("#addOther"); if (ao) otherGradePicker(ao, grade, inClass, guests, () => {});
  const add = $("#addStudents");
  if (add) add.onclick = async () => {
    const ids = [...$$(".addStu:checked").map((c) => c.value), ...guests];
    if (!ids.length) return;
    const t = rows[0] ? rows[0].f.SubjectType : (S.subjects.find((s) => s.f.Grade === grade && s.f.Title === subject)?.f.SubjectType || "I");
    try { await createRows(grade, subject, t, level, ids); render(); } catch (e) { flash("error", e.message); }
  };
}
async function setSubmitted(rows, value) {
  const res = await batch(rows.map((r) => ({ method: "PATCH", url: `${itemsUrl(marksListName(r.list))}/${r.id}/fields`, body: { Submitted: value } })));
  const bad = batchFailures(res);
  await Promise.all([...new Set(rows.map((r) => r.list))].map((g) => loadMarks(g)));
  render();
  if (bad.length) flash("error", `${bad.length} rows could not be updated: ${bad[0]?.body?.error?.message || ""}`);
  else flash("ok", value ? "Submitted to the advisor." : "Reopened. The teacher can edit again.");
}

// ── Advisor view ─────────────────────────────────────────────
function viewAdvisor(main) {
  const grades = viewGrades();
  if (!S.advGrade || !grades.includes(S.advGrade)) S.advGrade = grades[0] || "";
  const g = S.advGrade;
  const rows = S.marks[g] || [];
  const classes = groupClasses(rows).sort((a, b) => a.subject.localeCompare(b.subject));
  const expected = S.subjects.filter((s) => s.f.Grade === g && s.f.PrintsOnCard !== false);
  const notStarted = expected.filter((s) => !classes.some((c) => c.subject === s.f.Title));
  const roster = rosterFor(g);
  const subjCols = [...new Set([...expected.map((s) => s.f.Title), ...classes.map((c) => c.subject)])];
  const commentFor = (id) => S.comments.find((c) => c.f.StudentID === id && c.f.Grade === g);
  const canComment = canAdvise(g);
  const onlyMine = rows.length > 0 && rows.every((r) => r.by === S.me) && !isAdmin();
  main.innerHTML = `
    <h1>Advisor view</h1>
    <div class="row" style="margin-bottom:12px">
      <label style="max-width:180px"><span>Grade</span><select id="advGrade">${grades.map((x) => `<option ${x === g ? "selected" : ""}>${esc(x)}</option>`).join("")}</select></label>
      <div class="muted small" style="padding-bottom:8px">Advisor: ${esc(advisorName(g) || "not set")} · <span class="badge info">${esc(termLabel())}</span></div>
      <button class="btn primary" id="toPtc" style="margin-left:auto">${isMidterm() ? "PTC sheets for" : "Student summaries for"} ${esc(g)} →</button>
    </div>
    <div id="flash"></div>
    ${onlyMine ? notice("warn", "You can only see your own classes in this grade. Ask ICT to give you advisor access in Setup → Permissions.") : ""}
    <div class="stats">
      <div class="stat"><div class="k">Subjects submitted</div><div class="v">${classes.filter((c) => c.submitted).length} / ${Math.max(expected.length, classes.length)}</div></div>
      <div class="stat"><div class="k">Not started</div><div class="v">${notStarted.length}</div></div>
      <div class="stat"><div class="k">Students</div><div class="v">${roster.length}</div></div>
      <div class="stat"><div class="k">⚑ Concerns flagged</div><div class="v">${rows.filter((r) => r.f.Concern).length}</div></div>
      <div class="stat"><div class="k">${isMidterm() ? "PTC notes written" : "Advisor comments written"}</div><div class="v">${roster.filter((s) => (commentFor(s.f.Title)?.f.Comment || "").trim()).length} / ${roster.length}</div></div>
    </div>
    <h2>Subjects</h2>
    <div class="table-wrap"><table>
      <thead><tr><th>Subject</th><th>Teacher</th><th>Progress</th><th></th></tr></thead>
      <tbody>
        ${classes.map((c, i) => `<tr><td>${esc(subjectDisplay({ Subject: c.subject, ClassLevel: c.level }))}${c.grade !== g ? ` <span class="badge info" title="These ${esc(g)} bobaes sit in a ${esc(c.grade)} class">in ${esc(c.grade)} class</span>` : ""}</td><td>${esc(c.byName)}</td><td>${statusBadge(c)}</td><td><button class="btn small" data-open="${i}">Open</button></td></tr>`).join("")}
        ${notStarted.map((s) => `<tr><td>${esc(s.f.Title)}</td><td class="muted">${esc(s.f.Teacher || "—")}</td><td><span class="badge bad">No class created</span></td><td></td></tr>`).join("")}
        ${!classes.length && !notStarted.length ? `<tr><td colspan="4" class="muted">No subjects yet.</td></tr>` : ""}
      </tbody></table></div>
    <h2>Marks overview</h2>
    <p class="small muted">Score / behaviour. Highlighted cells are in a class but not yet filled in.</p>
    <div class="table-wrap"><table class="matrix">
      <thead><tr><th class="sticky">Student</th>${subjCols.map((s) => `<th class="subj">${esc(s)}</th>`).join("")}</tr></thead>
      <tbody>${roster.map((s) => `<tr><td class="sticky"><b>${esc(s.f.CalledName || s.f.FullName)}</b><div class="small muted">${esc(s.f.FullName)}</div></td>${subjCols.map((sub) => {
        const r = rows.find((x) => x.f.StudentID === s.f.Title && x.f.Subject === sub);
        if (!r) return `<td class="cell muted">·</td>`;
        const miss = blank(r.f.Score) || blank(r.f.Behaviour);
        return `<td class="cell ${miss ? "missing" : ""}" title="${esc(r.byName)}${r.f.Remark ? " — " + esc(r.f.Remark) : ""}">${blank(r.f.Score) ? "–" : esc(r.f.Score)} / ${blank(r.f.Behaviour) ? "–" : esc(r.f.Behaviour)}${r.f.Concern ? ` <span style="color:var(--bad)" title="Concern flagged">⚑</span>` : ""}</td>`;
      }).join("")}</tr>`).join("")}</tbody></table></div>
    <h2>${isMidterm() ? "PTC notes" : "Advisor comments"}</h2>
    <p class="small muted">${!canComment ? "Read only." : isMidterm()
      ? "Optional, one per student. It prints at the top of their PTC sheet, e.g. \"Promoted from ESL to regular classes this semester\". Saves when you leave the box."
      : "One comment per student. It prints in the teacher's comments box on page 1 of the report card. Saves when you leave the box."}</p>
    <div class="table-wrap"><table>
      <thead><tr><th style="width:220px">Student</th><th>Comment</th><th></th></tr></thead>
      <tbody>${roster.map((s) => { const c = commentFor(s.f.Title); return `
        <tr data-sid="${esc(s.f.Title)}"><td><div class="who-cell">${photo(s.f.Title, 40)}<div><b>${esc(s.f.CalledName || s.f.FullName)}</b><div class="small muted">${esc(s.f.FullName)}</div></div></div></td>
        <td><textarea class="cmt" rows="3" maxlength="1200" ${canComment ? "" : "disabled"}>${esc(c?.f.Comment || "")}</textarea></td>
        <td class="rowstate"></td></tr>`; }).join("")}</tbody></table></div>`;
  $("#advGrade").onchange = (e) => { S.advGrade = e.target.value; render(); };
  $("#toPtc").onclick = () => go("ptc");
  $$("[data-open]").forEach((b) => (b.onclick = () => { const c = classes[+b.dataset.open]; openClass(c.grade, c.subject, c.level, c.by, "advisor"); }));
  for (const tr of $$("tr[data-sid]", main)) {
    const ta = $(".cmt", tr), state = $(".rowstate", tr), sid = tr.dataset.sid;
    ta.onchange = () => saveComment(g, sid, ta.value, state);
  }
}
// Advisor comment for one bobae: update their row for this period, or create it. Shows progress in `state`.
async function saveComment(g, sid, value, state) {
  state.className = "rowstate"; state.textContent = "Saving…";
  try {
    const c = S.comments.find((x) => x.f.StudentID === sid && x.f.Grade === g), text = value.trim();
    if (c) { await graph("PATCH", `${itemsUrl(L.comments)}/${c.id}/fields`, { Comment: text }); c.f.Comment = text; c.modified = new Date().toISOString(); }
    else {
      const st = studentById(sid).f;
      const it = await graph("POST", itemsUrl(L.comments), { fields: { Title: `${st.FullName} (${g})`, StudentID: sid, Grade: g, ...termP(), Comment: text } });
      S.comments.push(norm(it));
    }
    state.textContent = "Saved ✓";
  } catch (e) { state.className = "rowstate err"; state.textContent = "Not saved"; flash("error", e.message); }
}

// ── Principal review: every grade at a glance, corrections, and the go-ahead for report cards ──
// Go-aheads live in the "approvals" item of Reporting Settings, keyed by year|semester|period, then grade.
const approvalKey = () => `${P().academicYear}|${P().semester}|${P().period}`;
const approvalFor = (g) => (S.approvals[approvalKey()] || {})[g] || null;
// Latest change to anything that prints on this grade's report cards (marks, comments, awards).
function lastChange(g) {
  const ids = new Set(rosterFor(g).map((s) => s.f.Title));
  const items = [...(S.marks[g] || []), ...S.comments.filter((c) => c.f.Grade === g), ...S.awards.filter((a) => ids.has(a.f.StudentID))];
  return items.reduce((m, x) => (x.modified && x.modified > m ? x.modified : m), "");
}
async function setApproval(g, on) {
  // Re-read first so two people approving different grades don't overwrite each other.
  const fresh = (await readList(L.settings)).find((x) => x.f.Title === "approvals");
  let all = {}; try { all = fresh ? JSON.parse(fresh.f.Json || "{}") : {}; } catch (_) { all = {}; }
  const k = approvalKey(); all[k] = all[k] || {};
  if (on) all[k][g] = { by: S.me, byName: S.meName, at: new Date().toISOString() }; else delete all[k][g];
  const Json = JSON.stringify(all);
  if (fresh) await graph("PATCH", `${itemsUrl(L.settings)}/${fresh.id}/fields`, { Json });
  else await graph("POST", itemsUrl(L.settings), { fields: { Title: "approvals", Json } });
  S.approvals = all;
}
function gradeSummary(g) {
  const c = gradeChecks(g), n = rosterFor(g).length, ids = new Set(rosterFor(g).map((s) => s.f.Title));
  const aw = S.awards.filter((a) => ids.has(a.f.StudentID));
  const ready = c.classes > 0 && !c.open.length && !c.missing && c.comments === n && !c.dupes.length;
  const ok = approvalFor(g);
  return { g, c, n, aw, ready, ok, changed: ok && lastChange(g) > ok.at, concerns: (S.marks[g] || []).filter((r) => r.f.Concern).length };
}
function approvalBadge(s) {
  if (!s.ok) return `<span class="badge">Not yet</span>`;
  return `<span class="badge ok">Go-ahead ✓</span><div class="small muted">${esc(s.ok.byName)}, ${esc(s.ok.at.slice(0, 10))}</div>${s.changed ? `<div><span class="badge warn">Changed since</span></div>` : ""}`;
}
function viewReview(main) {
  const grades = reviewGrades().filter((g) => rosterFor(g).length);
  const sums = grades.map(gradeSummary);
  const g = grades.includes(S.revGrade) ? S.revGrade : "";
  const scale = (v) => (blank(v) ? "–" : esc(v));
  main.innerHTML = `
    <h1>${canApprove() ? "Principal review" : `Review: ${esc(grades[0] || "")}–${esc(grades[grades.length - 1] || "")}`}</h1>
    <p class="lede"><span class="badge info">${esc(termLabel())}</span> ${canApprove() ? "Every grade’s" : "Your grades’"} marks, behaviour, comments and awards in one place. Check them, then give the go-ahead so the Admin team can generate the report cards.</p>
    <div id="flash"></div>
    ${isMidterm() ? notice("warn", "This is the <b>Midterm</b> period: progress only. Report cards are made at the <b>Final</b>, so the go-ahead opens then.") : ""}
    <div class="stats">
      <div class="stat"><div class="k">Grades complete</div><div class="v">${sums.filter((s) => s.ready).length} / ${sums.length}</div></div>
      <div class="stat"><div class="k">Go-ahead given</div><div class="v">${sums.filter((s) => s.ok).length} / ${sums.length}</div></div>
      <div class="stat"><div class="k">⚑ Concerns flagged</div><div class="v">${sums.reduce((t, s) => t + s.concerns, 0)}</div></div>
      <div class="stat"><div class="k">Awards given</div><div class="v">${S.awards.length}</div></div>
    </div>
    <div class="table-wrap"><table>
      <thead><tr><th>Grade</th><th class="num">Bobaes</th><th>Subjects</th><th class="num">Blank marks</th><th class="num">Comments</th><th class="num">Awards</th><th class="num">⚑</th><th>Go-ahead</th><th></th></tr></thead>
      <tbody>${sums.map((s) => `
        <tr ${s.g === g ? 'style="background:var(--panel-2,rgba(120,140,170,.12))"' : ""}><td><b>${esc(s.g)}</b><div class="small muted">${esc(advisorName(s.g))}</div></td><td class="num">${s.n}</td>
          <td class="small">${!s.c.classes ? `<span class="badge">None yet</span>` : s.c.open.length ? `<span class="badge warn">${s.c.classes - s.c.open.length} of ${s.c.classes} in</span>` : `<span class="badge ok">All ${s.c.classes} in</span>`}</td>
          <td class="num">${s.c.missing ? `<span class="badge warn">${s.c.missing}</span>` : "0"}</td>
          <td class="num">${s.c.comments === s.n ? `<span class="badge ok">${s.c.comments}/${s.n}</span>` : `${s.c.comments}/${s.n}`}</td>
          <td class="num">${s.aw.length}</td><td class="num">${s.concerns || ""}</td>
          <td>${approvalBadge(s)}</td>
          <td style="white-space:nowrap"><button class="btn small" data-rev="${esc(s.g)}">Review</button>
            ${canApproveGrade(s.g) && !isMidterm() ? (s.ok && !s.changed ? `<button class="btn small" data-undo="${esc(s.g)}">Withdraw</button>` : `<button class="btn small ${s.ready ? "primary" : ""}" data-ok="${esc(s.g)}">${s.ok ? "Give again" : "Give go-ahead"}</button>`) : ""}</td></tr>`).join("")}
      </tbody></table></div>
    <div id="revDetail"></div>`;
  if (g) {
    const rows = S.marks[g] || [];
    const order = (r) => { const x = S.subjects.find((s) => s.f.Grade === g && s.f.Title === r.f.Subject); return x ? (x.f.SortOrder ?? 999) : 999; };
    $("#revDetail").innerHTML = `
      <div class="row" style="align-items:center;margin-top:18px"><h2 style="margin:0">${esc(g)} · ${esc(advisorName(g))}</h2>
        <button class="btn small" id="revEdit" style="margin-left:auto">${canAdvise(g) ? "Correct marks in the Advisor view →" : "Open the Advisor view →"}</button></div>
      <p class="small muted">Score / behaviour for each subject (hover for the teacher’s remark). Comments save when you leave the box.</p>
      <div class="table-wrap"><table>
        <thead><tr><th style="width:200px">Bobae</th><th>Marks</th><th style="width:34%">Advisor comment</th><th>Awards</th><th></th></tr></thead>
        <tbody>${rosterFor(g).map((st) => { const id = st.f.Title;
          const mine = rows.filter((r) => r.f.StudentID === id).sort((a, b) => order(a) - order(b));
          const cm = S.comments.find((c) => c.f.StudentID === id && c.f.Grade === g);
          const aw = S.awards.filter((a) => a.f.StudentID === id);
          return `<tr data-sid="${esc(id)}"><td><div class="who-cell">${photo(id, 40)}<div><b>${esc(st.f.CalledName || st.f.FullName)}</b><div class="small muted">${esc(st.f.FullName)}</div></div></div></td>
            <td class="small">${mine.map((r) => `<span class="badge ${blank(r.f.Score) || blank(r.f.Behaviour) ? "warn" : ""}" title="${esc(r.byName)}${r.f.Remark ? " — " + esc(r.f.Remark) : ""}" style="margin:2px">${esc(subjectDisplay(r.f))} ${scale(r.f.Score)}/${scale(r.f.Behaviour)}${r.f.Concern ? " ⚑" : ""}</span>`).join("") || `<span class="muted">No marks yet</span>`}</td>
            <td><textarea class="cmt" rows="3" maxlength="1200" ${canAdvise(g) ? "" : "disabled"}>${esc(cm?.f.Comment || "")}</textarea></td>
            <td class="small">${aw.map((a) => esc(a.f.AwardName) + (a.f.Points ? ` (${esc(a.f.Points)})` : "")).join("<br>") || `<span class="muted">–</span>`}</td>
            <td class="rowstate"></td></tr>`; }).join("")}</tbody></table></div>`;
    $("#revEdit").onclick = () => { S.advGrade = g; go("advisor"); };
    for (const tr of $$("#revDetail tr[data-sid]")) { const ta = $(".cmt", tr); ta.onchange = () => saveComment(g, tr.dataset.sid, ta.value, $(".rowstate", tr)); }
  }
  $$("[data-rev]").forEach((b) => (b.onclick = () => { S.revGrade = S.revGrade === b.dataset.rev ? "" : b.dataset.rev; render(); if (S.revGrade) $("#revDetail").scrollIntoView({ behavior: "smooth" }); }));
  $$("[data-ok]").forEach((b) => (b.onclick = async () => {
    const s = sums.find((x) => x.g === b.dataset.ok);
    if (!s.ready && !confirm(`${s.g} isn’t complete yet (${s.c.open.length} subjects waiting, ${s.c.missing} blank marks, ${s.c.comments}/${s.n} comments). Give the go-ahead anyway?`)) return;
    b.disabled = true; b.textContent = "Saving…";
    try { await setApproval(s.g, true); render(); flash("ok", `Go-ahead given for ${s.g}. The Admin team can now generate its report cards.`); }
    catch (e) { flash("error", e.message); b.disabled = false; }
  }));
  $$("[data-undo]").forEach((b) => (b.onclick = async () => {
    if (!confirm(`Withdraw the go-ahead for ${b.dataset.undo}?`)) return;
    try { await setApproval(b.dataset.undo, false); render(); flash("ok", `Go-ahead withdrawn for ${b.dataset.undo}.`); } catch (e) { flash("error", e.message); }
  }));
}

// ── Certificates (Librarian): what to print, class by class, and what's done ──
// Every award this semester except service points. "Printed" is ticked here and kept in SharePoint.
function viewCerts(main) {
  const order = P().grades;
  const certs = S.awards.filter((a) => !a.f.Points && a.f.CertificateType !== "Service Points")
    .map((a) => { const st = studentById(a.f.StudentID)?.f || {}; return { a, grade: a.f.Grade || st.Grade || "", name: st.CertificateName || a.f.FullName || "", called: st.CalledName || a.f.CalledName || "" }; })
    .sort((x, y) => order.indexOf(x.grade) - order.indexOf(y.grade) || x.name.localeCompare(y.name) || (x.a.f.AwardName || "").localeCompare(y.a.f.AwardName || ""));
  const F = S.certF = S.certF || { grade: "", cat: "", status: "todo" };
  const cats = [...new Set(certs.map((c) => c.a.f.Category).filter(Boolean))].sort();
  const shown = certs.filter((c) => (!F.grade || c.grade === F.grade) && (!F.cat || c.a.f.Category === F.cat)
    && (F.status === "all" || (F.status === "done") === !!c.a.f.Printed));
  const byGrade = order.map((g) => [g, shown.filter((c) => c.grade === g)]).filter(([, l]) => l.length);
  const done = certs.filter((c) => c.a.f.Printed).length;
  const sel = (id, opts, v) => `<select id="${id}">${opts.map(([k, t]) => `<option value="${esc(k)}" ${k === v ? "selected" : ""}>${esc(t)}</option>`).join("")}</select>`;
  main.innerHTML = `
    <div class="no-print">
      <h1>Certificates</h1>
      <p class="lede"><span class="badge info">${esc(P().semester)} ${esc(P().academicYear)}</span> Every certificate given this semester, class by class. Print the list, print the certificates, then tick them off.</p>
      <div id="flash"></div>
      <div class="stats">
        <div class="stat"><div class="k">Certificates</div><div class="v">${certs.length}</div></div>
        <div class="stat"><div class="k">To print</div><div class="v">${certs.length - done}</div></div>
        <div class="stat"><div class="k">Printed</div><div class="v">${done}</div></div>
      </div>
      <div class="row" style="margin-bottom:12px;align-items:flex-end">
        <label style="max-width:160px"><span>Class</span>${sel("cfGrade", [["", "All classes"], ...order.filter((g) => certs.some((c) => c.grade === g)).map((g) => [g, `${g} (${certs.filter((c) => c.grade === g).length})`])], F.grade)}</label>
        <label style="max-width:240px"><span>Award type</span>${sel("cfCat", [["", "All types"], ...cats.map((c) => [c, c])], F.cat)}</label>
        <label style="max-width:180px"><span>Show</span>${sel("cfStatus", [["todo", "Still to print"], ["done", "Printed"], ["all", "All"]], F.status)}</label>
        <button class="btn primary" id="cfPrint">🖨 Print this list</button>
        <button class="btn" id="cfXlsx">Excel</button>
        ${F.status !== "done" && shown.length ? `<button class="btn" id="cfMark">Mark all shown as printed</button>` : ""}
      </div>
    </div>
    <div class="cert-list">
      <h2 style="margin-top:0">Certificates to print · ${esc(P().semester)} ${esc(P().academicYear)}${F.grade ? ` · ${esc(F.grade)}` : ""}${F.cat ? ` · ${esc(F.cat)}` : ""}</h2>
      ${byGrade.map(([g, l]) => `<div class="cert-grade">
        <h3>${esc(g)} · ${esc(advisorName(g))} <span class="small muted">${l.length} certificate${l.length > 1 ? "s" : ""}</span></h3>
        <div class="table-wrap"><table><thead><tr><th style="width:40px;text-align:center">✓</th><th>Name on certificate</th><th>Award</th><th>Certificate</th><th>Given by</th><th>Date</th></tr></thead><tbody>
          ${l.map((c) => `<tr data-aid="${esc(c.a.id)}"><td style="text-align:center"><input type="checkbox" class="cPrinted" ${c.a.f.Printed ? "checked" : ""}></td>
            <td><b>${esc(c.name)}</b>${c.called && !c.name.includes(c.called) ? ` <span class="small muted">(${esc(c.called)})</span>` : ""}</td>
            <td>${esc(c.a.f.AwardName)}</td><td class="small">${esc(c.a.f.CertificateType || "")}</td><td class="small">${esc(c.a.byName || c.a.f.AwardedBy || "")}</td><td class="small">${esc((c.a.f.AwardDate || "").slice(0, 10))}</td></tr>`).join("")}
        </tbody></table></div></div>`).join("") || `<p class="muted">${certs.length ? "Nothing here with these filters." : "No certificates have been given this semester yet."}</p>`}
    </div>`;
  const setF = (k) => (e) => { F[k] = e.target.value; render(); };
  $("#cfGrade").onchange = setF("grade"); $("#cfCat").onchange = setF("cat"); $("#cfStatus").onchange = setF("status");
  $("#cfPrint").onclick = () => window.print();
  $("#cfXlsx").onclick = async () => { try { await loadXlsx(); download(sheetBook(shown.map((c) => ({ Class: c.grade, Name_on_certificate: c.name, Called_name: c.called, Student_ID: c.a.f.StudentID,
    Award: c.a.f.AwardName, Certificate_type: c.a.f.CertificateType, Given_by: c.a.byName || c.a.f.AwardedBy, Date: (c.a.f.AwardDate || "").slice(0, 10), Printed: c.a.f.Printed ? "Yes" : "" })), "Certificates"),
    `Certificates_${P().semester}${P().academicYear.slice(0, 4)}${F.grade ? "_" + F.grade : ""}.xlsx`); } catch (e) { flash("error", e.message); } };
  const setPrinted = async (ids, on) => {
    const res = await batch(ids.map((id) => ({ method: "PATCH", url: `${itemsUrl(L.awards)}/${id}/fields`, body: { Printed: on } })));
    const bad = batchFailures(res);
    for (const id of ids) { const a = S.awards.find((x) => x.id === id); if (a) a.f.Printed = on; }
    return bad.length;
  };
  $$(".cPrinted").forEach((cb) => (cb.onchange = async () => {
    const id = cb.closest("tr").dataset.aid;
    try { const bad = await setPrinted([id], cb.checked); if (bad) throw new Error("Not saved"); render(); } catch (e) { cb.checked = !cb.checked; flash("error", e.message); }
  }));
  const mk = $("#cfMark"); if (mk) mk.onclick = async () => {
    const ids = shown.filter((c) => !c.a.f.Printed).map((c) => c.a.id);
    if (!ids.length || !confirm(`Mark ${ids.length} certificate${ids.length > 1 ? "s" : ""} as printed?`)) return;
    mk.disabled = true; mk.textContent = "Saving…";
    try { const bad = await setPrinted(ids, true); render(); flash(bad ? "error" : "ok", bad ? `${bad} could not be saved.` : `${ids.length} marked as printed.`); } catch (e) { flash("error", e.message); }
  };
}

// ── PTC sheets (advisor's parent-meeting helper, not an official report) ──
function ptcStudents(g) {
  const rows = S.marks[g] || [];
  const order = (r) => { const s = S.subjects.find((x) => x.f.Grade === g && x.f.Title === r.f.Subject); return s ? (s.f.SortOrder ?? 999) : 999; };
  // Primary PTC sheets show only the core subjects (Chinese, English, Maths, Science, and Accelerated English where taught).
  const core = isPrimary(g) ? ((P().primary || {}).ptcSubjects || []).map(lc) : null;
  return rosterFor(g).map((st) => {
    const id = st.f.Title;
    const subj = rows.filter((r) => r.f.StudentID === id && (!core || core.includes(lc(r.f.Subject))))
      .sort((a, b) => (a.f.SubjectType || "I").localeCompare(b.f.SubjectType || "I") || order(a) - order(b));
    const scored = subj.filter((r) => !blank(r.f.Score));
    const avg = scored.length ? scored.reduce((a, r) => a + Number(r.f.Score), 0) / scored.length : null;
    const byScore = [...scored].sort((a, b) => Number(b.f.Score) - Number(a.f.Score));
    const concerns = subj.filter((r) => r.f.Concern || Number(r.f.Behaviour) >= 3);
    return { st: st.f, subj, avg, concerns,
      top: scored.length >= 4 ? byScore.slice(0, 2) : [],
      low: scored.length >= 4 ? byScore.slice(-2).reverse() : [],
      missing: subj.filter((r) => blank(r.f.Score) || blank(r.f.Behaviour)),
      remarks: subj.filter((r) => (r.f.Remark || "").trim() && !concerns.includes(r)),
      note: (S.comments.find((c) => c.f.StudentID === id && c.f.Grade === g)?.f.Comment || "").trim() };
  });
}
function ptcPage(g, d, i, n) {
  const name = d.st.CalledName || d.st.FullName;
  const behClass = (b) => ["", "b1", "b2", "b3", "b4"][Number(b)] || "";
  const short = (b) => (P().behaviourScale.find((x) => x.value === Number(b))?.label || "");
  const subjName = (r) => esc(subjectDisplay(r.f));
  const overall = d.avg === null ? `No marks have been entered for ${esc(name)} yet.`
    : `${esc(name)} is averaging <b>${d.avg.toFixed(1)}%</b> across ${d.subj.filter((r) => !blank(r.f.Score)).length} subjects at ${esc(P().period.toLowerCase())}.`
      + (d.top.length ? ` Strongest: ${d.top.map((r) => `${subjName(r)} (${r.f.Score}%)`).join(", ")}.` : "")
      + (d.low.length ? ` Most room to grow: ${d.low.map((r) => `${subjName(r)} (${r.f.Score}%)`).join(", ")}.` : "")
      + (d.concerns.length ? ` ${d.concerns.length} subject${d.concerns.length > 1 ? "s" : ""} raised a behaviour concern.` : "");
  return `
  <section class="ptc-page">
    <header class="ptc-head">
      <span>WZFS · ${esc(g)} – ${esc(advisorName(g))} · Parent-Teacher Conference · ${esc(P().period)} ${esc(P().semester)} ${esc(P().academicYear.slice(0, 4))}</span>
      <span>${d.avg === null ? "" : `Avg ${d.avg.toFixed(1)}%`}</span>
    </header>
    <div class="ptc-name">${S.photos && S.photos[d.st.Title] ? `<img class="ptc-photo" src="${esc(S.photos[d.st.Title])}" alt="">` : ""}<span>${esc(name)} <small>${esc(d.st.FullName)} · ${esc(d.st.Title)}</small></span></div>
    ${d.note ? `<div class="ptc-note">${esc(d.note)}</div>` : ""}
    <div class="ptc-body">
      <div class="ptc-main">
        <h4>Overall</h4><p>${overall}</p>
        ${d.missing.length ? `<p class="ptc-miss">Not yet entered: ${d.missing.map(subjName).join(", ")}</p>` : ""}
        <h4>Subject scores</h4>
        <table class="ptc-scores">${d.subj.map((r) => `
          <tr><td>${subjName(r)}</td>
            <td class="bar"><i class="${behClass(r.f.Behaviour)}" style="width:${blank(r.f.Score) ? 0 : Math.max(2, Math.min(100, Number(r.f.Score)))}%"></i></td>
            <td class="num">${blank(r.f.Score) ? "–" : esc(r.f.Score) + "%"}</td><td class="beh">${esc(short(r.f.Behaviour))}${r.f.Concern ? " ⚑" : ""}</td></tr>`).join("")}
        </table>
        <div class="ptc-key">${P().behaviourScale.map((b) => `<span><i class="${behClass(b.value)}"></i>${esc(b.label)}</span>`).join("")}</div>
        ${d.remarks.length ? `<h4>What teachers are saying</h4>${d.remarks.map((r) => `<p class="ptc-q"><b>${subjName(r)}:</b> “${esc(r.f.Remark)}”</p>`).join("")}` : ""}
        ${d.concerns.length ? `<h4>Behaviour &amp; concerns</h4>${d.concerns.map((r) => `<p class="ptc-q concern"><b>${subjName(r)} [${[r.f.Concern ? "⚑ flagged by teacher" : "", Number(r.f.Behaviour) >= 3 ? esc(short(r.f.Behaviour)) : ""].filter(Boolean).join(" · ")}]</b>${r.f.Remark ? ` “${esc(r.f.Remark)}”` : ""}</p>`).join("")}` : ""}
      </div>
      <aside class="ptc-notes"><h4>Notes / action points</h4><div class="lines"></div></aside>
    </div>
    <footer class="ptc-foot"><span>CONFIDENTIAL · PTC use only · WZFS ${esc(P().academicYear)} · ${esc(name)}</span><span>${i} / ${n}</span></footer>
  </section>`;
}
function viewPtc(main) {
  const grades = viewGrades();
  if (!S.advGrade || !grades.includes(S.advGrade)) S.advGrade = grades[0] || "";
  const g = S.advGrade;
  const all = ptcStudents(g);
  const pick = S.ptcStudent && all.some((d) => d.st.Title === S.ptcStudent) ? S.ptcStudent : "";
  const list = pick ? all.filter((d) => d.st.Title === pick) : all;
  const flagged = all.filter((d) => d.concerns.length).length;
  main.innerHTML = `
    <div class="no-print">
      <p><a href="#" id="back">← Advisor view</a></p>
      <h1>${isMidterm() ? "PTC sheets" : "Student summaries"}</h1>
      <p class="lede">One page per bobae for parent meetings, built from what subject teachers entered. <b>Not an official report</b>: it's marked confidential and isn't sent home.</p>
      <div class="row" style="margin-bottom:12px">
        <label style="max-width:160px"><span>Grade</span><select id="ptcGrade">${grades.map((x) => `<option ${x === g ? "selected" : ""}>${esc(x)}</option>`).join("")}</select></label>
        <label style="max-width:280px"><span>Bobae</span><select id="ptcStu"><option value="">All (${all.length})</option>${all.map((d) => `<option value="${esc(d.st.Title)}" ${d.st.Title === pick ? "selected" : ""}>${esc(d.st.CalledName || d.st.FullName)}${d.concerns.length ? " ⚑" : ""}</option>`).join("")}</select></label>
        <button class="btn primary" id="ptcPrint">🖨 Print ${pick ? "this sheet" : `all ${list.length}`}</button>
        <span class="small muted" style="padding-bottom:8px">${flagged} with concerns · Print → "Save as PDF" for a file · A4 portrait</span>
      </div>
    </div>
    <div class="ptc-wrap">${list.map((d, i) => ptcPage(g, d, i + 1, list.length)).join("") || `<p class="muted">No bobaes in ${esc(g)}.</p>`}</div>`;
  $("#back").onclick = (e) => { e.preventDefault(); go("advisor"); };
  $("#ptcGrade").onchange = (e) => { S.advGrade = e.target.value; S.ptcStudent = ""; render(); };
  $("#ptcStu").onchange = (e) => { S.ptcStudent = e.target.value; render(); };
  $("#ptcPrint").onclick = () => window.print();
}

// ── Parent view (preview only) ───────────────────────────────
// Parents don't sign in to this app yet (see PDE question 8). This shows admins what a parent is given.
function viewParent(main) {
  const all = P().grades.flatMap((g) => rosterFor(g));
  if (!S.parentStudent || !studentById(S.parentStudent)) S.parentStudent = all[0]?.f.Title || "";
  const st = studentById(S.parentStudent), g = st?.f.Grade;
  const d = st ? ptcStudents(g).find((x) => x.st.Title === st.f.Title) : null;
  const comment = st ? (S.comments.find((c) => c.f.StudentID === st.f.Title && c.f.Grade === g)?.f.Comment || "") : "";
  const awards = st ? S.awards.filter((a) => a.f.StudentID === st.f.Title) : [];
  main.innerHTML = `
    <div class="no-print">
      <h1>Parent view</h1>
      <p class="lede">Parents don't sign in to this app. ${isMidterm()
        ? "At mid-term they're given their child's <b>PTC sheet</b> at the parent-teacher conference; that's what this shows."
        : "At the end of the semester they receive the <b>report card</b>, printed by the Admin team. This is a summary of what will be on it."}</p>
      <label style="max-width:320px;margin-bottom:14px"><span>Bobae</span><select id="pvStu">${P().grades.map((gr) => { const r = rosterFor(gr); return r.length ? `<optgroup label="${esc(gr)}">${r.map((s) => `<option value="${esc(s.f.Title)}" ${s.f.Title === S.parentStudent ? "selected" : ""}>${esc(s.f.CalledName || s.f.FullName)} · ${esc(s.f.FullName)}</option>`).join("")}</optgroup>` : ""; }).join("")}</select></label>
    </div>
    ${!st ? `<p class="muted">No bobaes in the roster yet.</p>` : isMidterm() ? `<div class="ptc-wrap">${ptcPage(g, d, 1, 1)}</div>` : `
      <div class="card">
        <h2 style="margin-top:0">${esc(st.f.CalledName || st.f.FullName)} <span class="muted small">${esc(st.f.FullName)} · ${esc(g)} · ${esc(P().semester)} ${esc(P().academicYear)}</span></h2>
        <div class="table-wrap"><table><thead><tr><th>Subject</th><th class="num">Score</th><th>Behaviour</th><th>Teacher's remark</th></tr></thead><tbody>
          ${d.subj.map((r) => `<tr><td>${esc(subjectDisplay(r.f))}</td><td class="num">${blank(r.f.Score) ? "–" : esc(r.f.Score)}</td><td>${esc(behLabel(r.f.Behaviour))}</td><td class="small">${esc(r.f.Remark || "")}</td></tr>`).join("") || `<tr><td colspan="4" class="muted">No marks yet.</td></tr>`}
        </tbody></table></div>
        <h3>Advisor's comment</h3><p class="small">${comment ? esc(comment) : `<span class="muted">Not written yet.</span>`}</p>
        <h3>Awards and service points</h3><p class="small">${awards.length ? awards.map((a) => esc(a.f.AwardName) + (a.f.Points ? ` (${esc(a.f.Points)})` : "")).join(" · ") : `<span class="muted">None yet.</span>`}</p>
      </div>`}`;
  $("#pvStu").onchange = (e) => { S.parentStudent = e.target.value; render(); };
}

// ── Awards ───────────────────────────────────────────────────
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
// Teachers see every active award all year; each is labelled with the semester it is usually given in.
const teacherCatalogue = () => (S.catalogueAll || []).filter((x) => x.f.Active !== false);
const seasonLabel = (s) => (s === "Fall" || s === "Spring" ? s : "All year");
// Advisor-only awards (criteria document): only the bobae's grade advisor / co-advisor, or an admin, may give them.
const advisorOnly = (c) => c && c.f.GivenBy === "Advisors";
const mayGive = (c, grade) => !advisorOnly(c) || isAdmin() || advisorEmails(grade).includes(S.me);
// Categories offered for a grade: Primary skips the ones that don't apply there (config primary.hiddenAwardCategories),
// and at Midterm only the midterm awards (Top Student, Outstanding Student) can be given.
function awardCategoriesFor(g) {
  const hide = isPrimary(g) ? ((P().primary || {}).hiddenAwardCategories || []).map(lc) : [];
  const mid = isMidterm() ? (P().midtermAwardCategories || []).map(lc) : null;
  return [...new Set(S.catalogue.map((c) => c.f.Category).filter(Boolean))]
    .filter((c) => !hide.includes(lc(c)) && (!mid || mid.includes(lc(c))));
}
function awardOptions(cat, selectedId) {
  if (!cat) return `<option value="">Choose a category first</option>`;
  const rank = (c) => (c.f.Season === P().semester ? 0 : !c.f.Season || c.f.Season === "Both" ? 1 : 2);   // this semester's first
  const list = S.catalogue.filter((c) => c.f.Category === cat).sort((a, b) => rank(a) - rank(b) || (a.f.Title || "").localeCompare(b.f.Title || ""));
  return `<option value="">Choose…</option>` + list.map((c) => `<option value="${esc(c.id)}" ${c.id === selectedId ? "selected" : ""}>${esc(c.f.Title)} · ${seasonLabel(c.f.Season)}${advisorOnly(c) ? " · advisors only" : ""}</option>`).join("");
}
function viewAwards(main) {
  const F = S.awForm;
  if (F.cat && !S.catalogue.some((c) => c.f.Category === F.cat)) Object.assign(F, { cat: "", id: "" });
  const grades = P().grades;
  if (!S.awGrade) S.awGrade = advisorGrades()[0] || grades[0];
  const g = S.awGrade;
  const q = lc(S.awSearch);
  const roster = rosterFor(g).filter((s) => !q || lc(`${s.f.FullName} ${s.f.CalledName} ${s.f.Title}`).includes(q));
  const seeAll = isLeader();
  const visible = S.awards.filter((a) => (S.awAll ? true : a.f.Grade === g));
  const countFor = (id) => S.awards.filter((a) => a.f.StudentID === id).length;
  const cats = awardCategoriesFor(g);
  if (F.cat && !cats.includes(F.cat)) Object.assign(F, { cat: "", id: "" });
  const today = new Date().toISOString().slice(0, 10);
  main.innerHTML = `
    <h1>Awards</h1>
    <p class="lede">Tap one or more students, choose the award, then save. Each award is stored with the student's ID, so it always lands on the right report card.</p>
    <div id="flash"></div>
    ${!S.lists[L.awards] ? notice("warn", "The Awards list isn't set up yet.") : ""}
    ${!S.catalogue.length ? notice("warn", "The award catalogue is empty. ICT needs to import it on the Setup tab.") : ""}
    <div class="two-col">
      <div>
        <div class="row" style="margin-bottom:12px">
          <label style="max-width:160px"><span>Grade</span><select id="awGrade">${grades.map((x) => `<option ${x === g ? "selected" : ""}>${esc(x)}</option>`).join("")}</select></label>
          <label style="flex:1;min-width:180px"><span>Find a student</span><input id="awSearch" placeholder="Name or ID" value="${esc(S.awSearch)}"></label>
          <button class="btn small" id="awAllBtn">Select all shown</button>
          <button class="btn small" id="awClear">Clear</button>
        </div>
        <div class="student-grid">
          ${roster.map((s) => { const n = countFor(s.f.Title); const on = S.awSel.has(s.f.Title); return `
            <button class="stu" data-sid="${esc(s.f.Title)}" aria-pressed="${on}">
              ${n ? `<span class="badge info count" title="Awards this semester">${n}</span>` : ""}
              <div class="stu-photo">${photo(s.f.Title, 56)}</div>
              <div class="called">${esc(s.f.CalledName || s.f.FullName)}</div>
              <div class="full">${esc(s.f.FullName)}</div>
              <div class="id">${esc(s.f.Title)}</div>
              ${on ? `<span class="tick">✓</span>` : ""}
            </button>`; }).join("")}
          ${!roster.length ? `<p class="muted">No students match.</p>` : ""}
        </div>
      </div>
      <div class="card award-panel">
        <h3>Give an award or service points</h3>
        <p class="small muted" id="awSelText"></p>
        <label><span>Category</span><select id="awCat"><option value="">Choose…</option>${cats.map((c) => `<option ${c === F.cat ? "selected" : ""}>${esc(c)}</option>`).join("")}</select></label>
        <label style="margin-top:8px"><span>Award <span class="muted">(labelled Fall, Spring or All year)</span></span><select id="awName" ${F.cat ? "" : "disabled"}>${awardOptions(F.cat, F.id)}</select></label>
        <div id="awCriteria" class="small" style="margin-top:6px"></div>
        <label style="margin-top:8px" id="awTitleWrap" hidden><span>Award title for your subject, in your own words (e.g. ICT: Coding Champion)</span><input id="awTitle" maxlength="120" value="${esc(F.title || "")}"></label>
        <label style="margin-top:8px" id="awMonthWrap" hidden><span>Month</span><select id="awMonth">${MONTHS.map((m) => `<option ${m === (F.month || MONTHS[new Date().getMonth()]) ? "selected" : ""}>${m}</option>`).join("")}</select></label>
        <label style="margin-top:8px"><span>Certificate type</span><select id="awType">${P().certificateTypes.map((t) => `<option ${t === F.type ? "selected" : ""}>${esc(t)}</option>`).join("")}</select></label>
        <label style="margin-top:8px"><span id="awPointsLabel">Service points (optional). A number like 16, or Gold / Silver / Bronze. Leave blank for a certificate.</span><input id="awPoints" maxlength="20" value="${esc(F.points || "")}"></label>
        <label style="margin-top:8px"><span>Date</span><input id="awDate" type="date" value="${esc(F.date || today)}"></label>
        <p style="margin-top:12px"><button class="btn primary" id="awSave" disabled>Save award</button></p>
      </div>
    </div>
    <h2>${seeAll ? "Awards this semester" : "Awards you have given this semester"}</h2>
    <label class="small"><input type="checkbox" id="awAllGrades" ${S.awAll ? "checked" : ""}> Show all grades</label>
    <div class="table-wrap" style="margin-top:8px"><table>
      <thead><tr><th>Student</th><th>ID</th><th>Grade</th><th>Award</th><th>Type</th><th>Points</th><th>Date</th>${seeAll ? "<th>Given by</th>" : ""}<th></th></tr></thead>
      <tbody>${visible.sort((a, b) => (b.f.AwardDate || "").localeCompare(a.f.AwardDate || "")).map((a) => `
        <tr><td>${esc(a.f.FullName)}</td><td class="small">${esc(a.f.StudentID)}</td><td>${esc(a.f.Grade)}</td><td>${esc(a.f.AwardName)}<div class="small muted">${esc(a.f.Category)}</div></td>
        <td class="small">${esc(a.f.CertificateType)}</td><td>${esc(a.f.Points || "")}</td><td class="small">${esc((a.f.AwardDate || "").slice(0, 10))}</td>
        ${seeAll ? `<td class="small">${esc(a.byName)}</td>` : ""}
        <td>${a.by === S.me || isAdmin() ? `<button class="btn small danger" data-del="${a.id}">Delete</button>` : ""}</td></tr>`).join("")}
        ${!visible.length ? `<tr><td colspan="9" class="muted">None yet.</td></tr>` : ""}
      </tbody></table></div>`;
  const current = () => S.catalogue.find((x) => x.id === $("#awName").value);
  const updSel = () => {
    const n = S.awSel.size, c = current();
    const names = [...S.awSel].map((id) => studentById(id)?.f.CalledName || id);
    $("#awSelText").textContent = n ? `${n} selected: ${names.slice(0, 6).join(", ")}${n > 6 ? "…" : ""}` : "No students selected yet.";
    $("#awTitleWrap").hidden = !(c && c.f.CustomTitle);
    $("#awMonthWrap").hidden = !(c && c.f.AskMonth);
    const isPoints = c && c.f.CertificateType === "Service Points";
    $("#awPointsLabel").textContent = isPoints ? "Service points (required): a number like 16, or Gold / Silver / Bronze." : "Service points (optional). A number like 16, or Gold / Silver / Bronze. Leave blank for a certificate.";
    const blocked = c ? [...new Set([...S.awSel].map((id) => studentById(id)?.f.Grade).filter((g) => !mayGive(c, g)))] : [];
    $("#awCriteria").innerHTML = (c && c.f.Criteria ? notice("", `<b>Criteria:</b> ${esc(c.f.Criteria)}`) : "")
      + (blocked.length ? notice("warn", `<b>Advisors only.</b> Only the ${blocked.map(esc).join(" / ")} advisor or co-advisor can give this award. Please pass the name to them.`)
        : advisorOnly(c) ? `<p class="small muted">Advisors only: you're the advisor for this grade.</p>` : "");
    const ready = n && c && !blocked.length && (!c.f.CustomTitle || $("#awTitle").value.trim()) && (!isPoints || $("#awPoints").value.trim());
    $("#awSave").disabled = !ready;
  };
  const remember = () => Object.assign(S.awForm, { cat: $("#awCat").value, id: $("#awName").value, type: $("#awType").value, points: $("#awPoints").value,
    date: $("#awDate").value, month: $("#awMonth").value, title: $("#awTitle").value });
  ["#awType", "#awPoints", "#awDate", "#awMonth", "#awTitle"].forEach((s) => ($(s).oninput = $(s).onchange = () => { remember(); updSel(); }));
  $("#awGrade").onchange = (e) => { S.awGrade = e.target.value; S.awSel.clear(); render(); };
  const search = $("#awSearch");
  search.oninput = () => { S.awSearch = search.value; const pos = search.selectionStart; render(); const s2 = $("#awSearch"); s2.focus(); s2.setSelectionRange(pos, pos); };
  $("#awAllBtn").onclick = () => { roster.forEach((s) => S.awSel.add(s.f.Title)); render(); };
  $("#awClear").onclick = () => { S.awSel.clear(); render(); };
  $("#awAllGrades").onchange = (e) => { S.awAll = e.target.checked; render(); };
  $$(".stu").forEach((b) => (b.onclick = () => { const id = b.dataset.sid; S.awSel.has(id) ? S.awSel.delete(id) : S.awSel.add(id); render(); }));
  const cat = $("#awCat"), nm = $("#awName"), ty = $("#awType");
  cat.onchange = () => {
    nm.disabled = !cat.value;
    nm.innerHTML = awardOptions(cat.value, "");
    remember(); updSel();
  };
  nm.onchange = () => {
    const c = current();
    if (c && P().certificateTypes.includes(c.f.CertificateType)) ty.value = c.f.CertificateType;
    remember(); updSel();
  };
  $("#awSave").onclick = async () => {
    const c = current(); if (!c) return;
    if ([...S.awSel].some((id) => !mayGive(c, studentById(id)?.f.Grade))) return flash("warn", "This award can only be given by the bobae's advisor or co-advisor.");
    // What prints on the report card: the teacher's own title, or "Award – Month", or the catalogue name.
    // Midterm awards (Top Student, Outstanding Student) are marked so they're told apart from the Final ones.
    const awardName = (c.f.CustomTitle ? $("#awTitle").value.replace(/\s+/g, " ").trim() : c.f.AskMonth ? `${c.f.Title} – ${$("#awMonth").value}` : c.f.Title)
      + (isMidterm() ? " (Midterm)" : "");
    const ids = [...S.awSel];
    const dupes = ids.filter((id) => S.awards.some((a) => a.f.StudentID === id && lc(a.f.AwardName) === lc(awardName)));
    if (dupes.length && !confirm(`${dupes.length} of these students already have "${awardName}" this semester. Add it again?`)) return;
    if (c.f.CustomTitle) {
      const mine = S.awards.filter((a) => a.by === S.me && a.f.Category === c.f.Category && !S.catalogue.some((x) => x.f.Title === a.f.AwardName && !x.f.CustomTitle));
      if (mine.length + ids.length > 1 && !confirm(`Each teacher gives only ONE Teacher Award per semester, for one subject. You have already given ${mine.length} this semester. Save anyway?`)) return;
    }
    $("#awSave").disabled = true; $("#awSave").textContent = "Saving…";
    const url = itemsUrl(L.awards), points = $("#awPoints").value.trim(), date = $("#awDate").value;
    const res = await batch(ids.map((id) => { const s = studentById(id).f; return { method: "POST", url, body: { fields: {
      Title: id, StudentID: id, FullName: s.FullName, CalledName: s.CalledName || "", Grade: s.Grade,
      Category: c.f.Category, AwardName: awardName, CertificateType: ty.value, Points: points,
      ...term(), AwardDate: date ? date + "T00:00:00Z" : null, AwardedBy: S.meName, Printed: false } } }; }));
    const bad = batchFailures(res);
    await loadAwards();
    S.awSel.clear(); S.awForm.title = ""; S.awForm.points = ""; render();
    if (bad.length) flash("error", `${bad.length} award(s) could not be saved: ${bad[0]?.body?.error?.message || ""}`);
    else flash("ok", `Saved "${awardName}" for ${ids.length} student${ids.length > 1 ? "s" : ""}.`);
  };
  $$("[data-del]").forEach((b) => (b.onclick = async () => {
    if (!confirm("Delete this award?")) return;
    try { await graph("DELETE", `${itemsUrl(L.awards)}/${b.dataset.del}`); await loadAwards(); render(); } catch (e) { flash("error", e.message); }
  }));
  updSel();
}

// ── Export (admin) ───────────────────────────────────────────
function marksExportRows(g) {
  const rows = S.marks[g] || [];
  const out = [];
  const subjOrder = (r) => { const s = S.subjects.find((x) => x.f.Grade === g && x.f.Title === r.f.Subject); return s ? (s.f.SortOrder ?? 999) : 999; };
  for (const st of rosterFor(g)) {
    const id = st.f.Title;
    const comment = (S.comments.find((c) => c.f.StudentID === id && c.f.Grade === g)?.f.Comment || "").trim();
    const mine = rows.filter((r) => r.f.StudentID === id)
      .sort((a, b) => (a.f.SubjectType || "I").localeCompare(b.f.SubjectType || "I") || subjOrder(a) - subjOrder(b));
    const base = { Student_ID: id, Student_Name: st.f.FullName, Called_Name: st.f.CalledName || "", DOB: st.f.DOB || "", Grade: g, Advisor: advisorName(g),
      Department: st.f.Department || "", Grading_Period: `${P().semester} Semester` };
    if (!mine.length) { if (comment) out.push({ ...base, Teacher_Comments: comment }); continue; }
    for (const r of mine) out.push({ ...base,
      Subject_Name: subjectDisplay(r.f), Subject_Type: r.f.SubjectType || "I",
      Score: blank(r.f.Score) ? "" : r.f.Score, Behaviour: blank(r.f.Behaviour) ? "" : r.f.Behaviour,
      Subject_Remark: r.f.Remark || "", Teacher_Comments: comment, Subject_Teacher: r.byName || r.f.TeacherName || "" });
  }
  return out;
}
function gradeChecks(g) {
  const rows = S.marks[g] || [];
  const classes = groupClasses(rows);
  const seen = new Map(), dupes = [];
  for (const r of rows) { const k = r.f.StudentID + "|" + r.f.Subject; if (seen.has(k) && seen.get(k) !== r.by) dupes.push(r); else seen.set(k, r.by); }
  const notInRoster = rows.filter((r) => !studentById(r.f.StudentID));
  return { rows: rows.length, missing: rows.filter((r) => blank(r.f.Score) || blank(r.f.Behaviour)).length,
    open: classes.filter((c) => !c.submitted), classes: classes.length, dupes, notInRoster,
    comments: rosterFor(g).filter((s) => S.comments.some((c) => c.f.StudentID === s.f.Title && c.f.Grade === g && (c.f.Comment || "").trim())).length };
}
function download(wb, name) { XLSX.writeFile(wb, name); }
function sheetBook(rows, sheetName) { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), sheetName); return wb; }
// Awards in the column layout the generator reads (Student ID first, so it matches exactly).
function awardsForGenerator(grade) {
  return S.awards.filter((a) => !grade || a.f.Grade === grade)
    .sort((a, b) => (a.f.Grade || "").localeCompare(b.f.Grade || "") || (a.f.FullName || "").localeCompare(b.f.FullName || ""))
    .map((a) => { const s = studentById(a.f.StudentID)?.f || a.f; return {
      Student_ID: a.f.StudentID, Student_Name: s.FullName, Grade: a.f.Grade, Award: a.f.AwardName,
      Recognition: a.f.CertificateType, Points: a.f.Points || "", Date: (a.f.AwardDate || "").slice(0, 10) }; });
}
const yearLong = () => { const [a, b] = P().academicYear.split("-"); return b ? `${a}–${a.slice(0, 2)}${b}` : P().academicYear; };
// Hand one grade to the report card generator (same site, /cards/) — no files downloaded or uploaded.
// The generator reads the hand-off once and deletes it straight away.
function openGenerator(g) {
  const roster = rosterFor(g);
  const payload = { createdAt: Date.now(), by: S.meName, marks: marksExportRows(g), awards: awardsForGenerator(g), attendance: S.attendance ? S.attendance.rows : null,
    meta: { grade: g, advisor: advisorName(g), department: roster[0]?.f.Department || "", gradingPeriod: `${P().semester} Semester`, schoolYear: yearLong() } };
  try { localStorage.setItem("wzfs_reporting_handoff_v1", JSON.stringify(payload)); }
  catch (e) { return flash("error", "Your browser blocked the hand-off (storage full or private mode). Use “Download files instead” below."); }
  const w = window.open("cards/index.html", "_blank");
  if (!w) flash("warn", "Your browser blocked the new tab. Allow pop-ups for this site, then click Generate again.");
}
function viewExport(main) {
  const tag = `${P().semester}${P().academicYear.slice(0, 4)}`;
  const grades = P().grades.filter((g) => rosterFor(g).length);
  const totals = { certs: 0, points: 0 };
  const rows = grades.map((g) => {
    const c = gradeChecks(g), n = rosterFor(g).length;
    const aw = S.awards.filter((a) => a.f.Grade === g), certs = aw.filter((a) => !a.f.Points).length, pts = aw.length - certs;
    totals.certs += certs; totals.points += pts;
    const ready = c.classes > 0 && !c.open.length && !c.missing && c.comments === n && !c.dupes.length;
    // At Final, report cards wait for the Principal's go-ahead (Principal review tab).
    const ok = approvalFor(g), approved = isMidterm() || !!ok;
    return { g, c, n, certs, pts, ready, ok, approved, changed: ok && lastChange(g) > ok.at };
  });
  const readyCount = rows.filter((r) => r.ready).length;
  const restricted = new Set(teacherCatalogue().filter(advisorOnly).map((c) => lc(c.f.Category + "|" + c.f.Title)));
  const odd = S.awards.filter((a) => restricted.has(lc(a.f.Category + "|" + a.f.AwardName)) && !advisorEmails(a.f.Grade).includes(a.by) && !P().admins.map(lc).includes(a.by));
  const byCat = Object.entries(S.awards.reduce((m, a) => ((m[a.f.Category || "Other"] = (m[a.f.Category || "Other"] || 0) + 1), m), {})).sort((a, b) => b[1] - a[1]);
  main.innerHTML = `
    <h1>Report cards</h1>
    <p class="lede"><span class="badge info">${esc(P().semester)} ${esc(P().academicYear)}</span> Check each grade is complete, then click <b>Generate</b>. The report cards open in a new tab, ready to check and print. No files to download or upload.</p>
    <div id="flash"></div>
    ${isMidterm() ? notice("warn", `The reporting period is <b>Midterm</b>, so these are mid-term marks. For end-of-semester report cards, ICT switches the period to <b>Final</b>.`) : ""}
    <div class="stats">
      <div class="stat"><div class="k">Grades ready</div><div class="v">${readyCount} / ${rows.length}</div></div>
      <div class="stat"><div class="k">Certificates</div><div class="v">${totals.certs}</div></div>
      <div class="stat"><div class="k">Service points entries</div><div class="v">${totals.points}</div></div>
      <div class="stat"><div class="k">Attendance</div><div class="v" style="font-size:15px">${S.attendance ? `✓ ${esc(S.attendance.name)}<div class="small muted">${S.attendance.rows.length} rows</div>` : `<span class="muted">Not added</span>`}</div></div>
    </div>
    <div class="card">
      <h3>Attendance sheet</h3>
      <p class="small muted" style="margin-top:0">One Excel file for all students (from the attendance system). It's used for every grade you generate in this session, and isn't saved anywhere.</p>
      <div class="row"><label style="max-width:360px"><span>Attendance file</span><input type="file" id="attFile" accept=".xlsx,.xls,.csv"></label>
        ${S.attendance ? `<button class="btn small" id="attClear">Remove</button>` : ""}<button class="btn small" id="reload">↻ Reload latest marks and awards</button></div>
    </div>
    <div class="table-wrap"><table>
      <thead><tr><th>Grade</th><th class="num">Bobaes</th><th>Marks</th><th class="num">Advisor comments</th><th>Awards</th><th>Status</th><th></th></tr></thead>
      <tbody>${rows.map(({ g, c, n, certs, pts, ready, ok, approved, changed }) => `
        <tr><td><b>${esc(g)}</b><div class="small muted">${esc(advisorName(g))}</div></td><td class="num">${n}</td>
          <td class="small">${!c.classes ? `<span class="badge">No classes yet</span>` : c.open.length ? `<span class="badge warn">${c.classes - c.open.length} of ${c.classes} subjects in</span><div class="muted">Waiting: ${c.open.map((x) => esc(x.subject)).join(", ")}</div>` : `<span class="badge ok">All ${c.classes} subjects in</span>`}
            ${c.missing ? `<div><span class="badge warn">${c.missing} blank marks</span></div>` : ""}${c.dupes.length ? `<div><span class="badge bad">${c.dupes.length} duplicate subject rows</span></div>` : ""}</td>
          <td class="num">${c.comments === n ? `<span class="badge ok">${c.comments} / ${n}</span>` : `${c.comments} / ${n}`}</td>
          <td class="small">${certs} certificates<div class="muted">${pts} service points</div></td>
          <td>${ready ? `<span class="badge ok">Ready</span>` : `<span class="badge warn">Not complete</span>`}
            ${isMidterm() ? "" : ok ? `<div><span class="badge ok">Go-ahead ✓</span></div>${changed ? `<div><span class="badge warn">Changed since go-ahead</span></div>` : ""}` : `<div><span class="badge">Waiting for go-ahead</span></div>`}</td>
          <td><button class="btn ${ready && approved ? "primary" : ""} small" data-gen="${esc(g)}" ${c.rows || c.comments ? "" : "disabled"}>Generate →</button></td></tr>`).join("")}
      </tbody></table></div>
    ${odd.length ? notice("warn", `${odd.length} advisor-only award${odd.length > 1 ? "s were" : " was"} saved by someone who isn't that grade's advisor: `
      + odd.slice(0, 8).map((a) => `${esc(a.f.CalledName || a.f.FullName)} (${esc(a.f.Grade)}): ${esc(a.f.AwardName)}, by ${esc(a.byName || a.f.AwardedBy)}`).join("; ") + (odd.length > 8 ? "…" : "") + ". Check with the advisor before printing.") : ""}
    <details class="card"><summary>Awards collected this semester: ${S.awards.length} in ${byCat.length} categories</summary>
      <div class="table-wrap" style="margin-top:10px;max-width:520px"><table><thead><tr><th>Category</th><th class="num">Given</th></tr></thead><tbody>
        ${byCat.map(([k, v]) => `<tr><td>${esc(k)}</td><td class="num">${v}</td></tr>`).join("") || `<tr><td colspan="2" class="muted">None yet.</td></tr>`}
      </tbody></table></div>
      <p><button class="btn small" id="printList">Certificate print list (Excel)</button></p>
    </details>
    <details class="card"><summary>Download files instead</summary>
      <p class="small muted">Only needed if the Generate button can't open the report cards (for example, if pop-ups are blocked). Upload these at the old generator.</p>
      <div class="row">${grades.map((g) => `<button class="btn small" data-g="${esc(g)}">Marks ${esc(g)}</button>`).join("")}</div>
      <p class="row" style="margin-top:8px"><button class="btn small" id="awardsFile">Awards file (all grades)</button><button class="btn small" id="allMarks">All grades in one marks file</button></p>
    </details>
    <p class="small muted">"Ready" means every subject is submitted, no marks are blank, every bobae has an advisor comment, and no subject is entered twice. At Final, print a grade's report cards once the Principal has given the go-ahead (Principal review tab). You can still generate a check copy before that.</p>`;
  $("#reload").onclick = async () => { $("#reload").textContent = "Loading…"; await loadAll(); render(); flash("ok", "Up to date."); };
  $("#attFile").onchange = async (e) => {
    const f = e.target.files[0]; if (!f) return;
    try { const wb = await readSheet(f); S.attendance = { name: f.name, rows: XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" }) }; render(); flash("ok", `Attendance added: ${S.attendance.rows.length} rows.`); }
    catch (err) { flash("error", "Couldn't read the attendance file: " + err.message); }
  };
  const ac = $("#attClear"); if (ac) ac.onclick = () => { S.attendance = null; render(); };
  $$("[data-gen]").forEach((b) => (b.onclick = () => {
    const r = rows.find((x) => x.g === b.dataset.gen);
    if (!r.ready && !confirm(`${r.g} isn't complete yet${r.c.open.length ? ` (waiting for ${r.c.open.length} subject${r.c.open.length > 1 ? "s" : ""})` : ""}. Generate a check copy anyway?`)) return;
    if (r.ready && !r.approved && !confirm(`${isPrimary(r.g) ? "The Vice Principal (Primary)" : "The Principal"} hasn't given the go-ahead for ${r.g} yet. Generate a check copy only (not for printing)?`)) return;
    if (!S.attendance && !confirm("No attendance sheet added. Generate without attendance?")) return;
    openGenerator(r.g);
  }));
  loadXlsx().catch(() => {});   // warm it up for the attendance file and downloads
  const withXlsx = (fn) => async () => { try { await loadXlsx(); fn(); } catch (e) { flash("error", e.message); } };
  $$("[data-g]").forEach((b) => (b.onclick = withXlsx(() => download(sheetBook(marksExportRows(b.dataset.g), "Marks"), `Marks_${b.dataset.g}_${tag}.xlsx`))));
  $("#allMarks").onclick = withXlsx(() => download(sheetBook(P().grades.flatMap(marksExportRows), "Marks"), `Marks_AllGrades_${tag}.xlsx`));
  const printList = () => S.awards.filter((a) => !a.f.Points).map((a) => ({
    Grade: a.f.Grade, Student_ID: a.f.StudentID, Certificate_Name: studentById(a.f.StudentID)?.f.CertificateName || a.f.FullName,
    Award: a.f.AwardName, Certificate_Type: a.f.CertificateType, Given_By: a.byName || a.f.AwardedBy, Printed: a.f.Printed ? "Yes" : "No" }));
  $("#awardsFile").onclick = withXlsx(() => {
    const wb = sheetBook(awardsForGenerator(), "Awards");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(printList()), "Certificate print list");
    download(wb, `Awards_${tag}.xlsx`);
  });
  $("#printList").onclick = withXlsx(() => download(sheetBook(printList(), "Certificate print list"), `Certificates_${tag}.xlsx`));
}

// ── Setup (admin) ────────────────────────────────────────────
const T = (name, multi) => ({ name, text: multi ? { allowMultipleLines: true } : {} });
const N = (name, minimum, maximum, decimalPlaces = "automatic") => ({ name, number: { decimalPlaces, ...(minimum !== undefined ? { minimum, maximum } : {}) } });
const CH = (name, choices) => ({ name, choice: { choices, displayAs: "dropDownMenu" } });
const B = (name) => ({ name, boolean: {} });
const D = (name) => ({ name, dateTime: { format: "dateOnly" } });
// kind: "read" = staff read only; "entry" = staff contribute, item-level privacy; "marks" = entry + advisor Design
function listSpecs() {
  const specs = [
    { name: L.students, kind: "read", cols: [T("FullName"), T("CalledName"), T("CertificateName"), T("Grade"), T("Email"), T("Department"), T("Joined"), CH("Status", ["Active", "Withdrawn"]), T("DOB")] },
    { name: L.subjects, kind: "read", cols: [T("Grade"), CH("SubjectType", ["I", "II"]), T("Teacher"), B("PrintsOnCard"), N("SortOrder")] },
    { name: L.catalogue, kind: "read", cols: [T("Category"), T("CertificateType"), CH("Season", ["Both", "Fall", "Spring"]), T("Criteria", true), B("AskMonth"), B("CustomTitle"), T("GivenBy"), B("Active"), N("SortOrder")] },
    { name: L.settings, kind: "read", cols: [T("Json", true)] },
    // Document library of small ID photos, <StudentID>.jpg. Staff read; bobaes never have access.
    { name: L.photos, kind: "read", library: true, cols: [] },
    { name: L.awards, kind: "entry", cols: [T("StudentID"), T("FullName"), T("CalledName"), T("Grade"), T("Category"), T("AwardName"), T("CertificateType"), T("Points"), T("Semester"), T("AcademicYear"), D("AwardDate"), T("AwardedBy"), B("Printed")] },
    { name: L.comments, kind: "entry", cols: [T("StudentID"), T("Grade"), T("Semester"), T("AcademicYear"), T("Period"), T("Comment", true)] },
  ];
  for (const g of P().grades) specs.push({ name: marksListName(g), grade: g, kind: "marks", cols: [
    T("StudentID"), T("FullName"), T("CalledName"), T("Subject"), CH("SubjectType", ["I", "II"]), T("ClassLevel"), T("ClassGrade"), T("Semester"), T("AcademicYear"), T("Period"),
    N("Score", 0, 100, "one"), N("Behaviour", 1, 4, "none"), T("Remark", true), B("Concern"), B("Submitted"), T("TeacherName")] });
  return specs;
}
async function readSheet(file) {
  await loadXlsx();
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => { try { resolve(XLSX.read(new Uint8Array(fr.result), { type: "array" })); } catch (e) { reject(e); } };
    fr.onerror = () => reject(fr.error); fr.readAsArrayBuffer(file);
  });
}
const rowsOf = (wb, nameRx) => { const n = wb.SheetNames.find((s) => nameRx && nameRx.test(s)) || wb.SheetNames[0]; return XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: "" }); };
const col = (row, ...names) => { for (const n of names) { const k = Object.keys(row).find((k) => lc(k).replace(/[\s_]/g, "") === lc(n).replace(/[\s_]/g, "")); if (k && String(row[k]).trim() !== "") return String(row[k]).trim(); } return ""; };

// Upsert: match existing list items by key, PATCH changed ones, POST new ones.
async function upsert(listName, records, keyOf, existing) {
  const url = itemsUrl(listName);
  const byKey = new Map(existing.map((x) => [keyOf(x.f), x]));
  const reqs = [];
  for (const rec of records) {
    const cur = byKey.get(keyOf(rec));
    if (!cur) reqs.push({ method: "POST", url, body: { fields: rec } });
    else if (Object.keys(rec).some((k) => String(cur.f[k] ?? "") !== String(rec[k] ?? ""))) reqs.push({ method: "PATCH", url: `${url}/${cur.id}/fields`, body: rec });
  }
  const res = await batch(reqs);
  return { sent: reqs.length, failed: batchFailures(res) };
}

function viewSetup(main) {
  const specs = listSpecs();
  const missing = specs.filter((s) => !S.lists[s.name]);
  const p = P();
  main.innerHTML = `
    <h1>Setup</h1>
    <p class="lede">For ICT only. Work through the steps in order. Running a step again is safe.</p>
    <div id="flash"></div>

    <div class="card"><h3>1 · SharePoint lists</h3>
      <p class="small">${specs.length - missing.length} of ${specs.length} lists exist on <a href="${esc(S.siteUrl)}" target="_blank" rel="noopener">${esc(S.siteUrl)}</a>.</p>
      ${missing.length ? `<p class="small muted">Missing: ${missing.map((m) => esc(m.name)).join(", ")}</p>` : `<p><span class="badge ok">All lists exist</span></p>`}
      <button class="btn ${missing.length ? "primary" : ""}" id="createLists">${missing.length ? "Create the missing lists" : "Check lists for new columns"}</button>
      <span class="small muted">Also adds any columns a newer version of the app needs. Existing data is never touched.</span>
      <pre id="createLog" class="small muted" style="white-space:pre-wrap"></pre>
    </div>

    <div class="card"><h3>2 · Import data</h3>
      <p class="small muted">Use the prepared Fall 2026 workbooks. Existing rows are updated by key, new rows added; nothing is deleted.</p>
      <div class="grid" style="grid-template-columns:repeat(auto-fill,minmax(280px,1fr))">
        <label><span>Roster: <code>List_Students_Fall2026.xlsx</code> (${S.students.length} now)</span><input type="file" id="fStudents" accept=".xlsx,.xls,.csv"></label>
        <label><span>Subjects: <code>WZFS_Fall2026_UpperSchool_Subjects.xlsx</code> (${S.subjects.length} now)</span><input type="file" id="fSubjects" accept=".xlsx,.xls,.csv"></label>
        <label><span>Award catalogue: <code>WZFS_Awards_Catalogue_AllYear.xlsx</code> (${S.catalogue.length} now). Not the older Fall-only file: anything missing from the file you pick is hidden.</span><input type="file" id="fCatalogue" accept=".xlsx,.xls,.csv"></label>
        <label><span>Student photos: select all the files in <code>StudentPhotos_Fall2026\\ByStudentID</code> (${Object.keys(S.photos || {}).length} now). Each file is named by Student ID; a new upload replaces the old photo.</span><input type="file" id="fPhotos" accept=".jpg,.jpeg,.png" multiple ${S.lists[L.photos] ? "" : "disabled"}></label>
      </div>
      <pre id="importLog" class="small muted" style="white-space:pre-wrap"></pre>
    </div>

    <div class="card"><h3>3 · Term and people</h3>
      <div class="grid">
        <label><span>Academic year</span><input id="sYear" value="${esc(p.academicYear)}"></label>
        <label><span>Semester</span><select id="sSem">${p.semesters.map((s) => `<option ${s === p.semester ? "selected" : ""}>${esc(s)}</option>`).join("")}</select></label>
        <label><span>Reporting period: what teachers are entering now</span><select id="sPeriod">${p.periods.map((s) => `<option ${s === p.period ? "selected" : ""}>${esc(s)}</option>`).join("")}</select></label>
        <label class="wide"><span>Admins (full access in this app), one email per line</span><textarea id="sAdmins" rows="2">${esc(p.admins.join("\n"))}</textarea></label>
        <label class="wide"><span>Leadership (read-only across grades), one email per line. Also add them as Owners of the SharePoint site.</span><textarea id="sLead" rows="3">${esc(p.leadership.join("\n"))}</textarea></label>
        <label class="wide"><span>Admin team: the only people who can generate report cards (the Report cards tab). One email per line; admins are included automatically.</span><textarea id="sReport" rows="3">${esc((p.reportTeam || []).join("\n"))}</textarea></label>
        <label class="wide"><span>Division heads: read-only view of their own grades. One per line, as <code>email = G1 G2 G3</code></span><textarea id="sHeads" rows="2">${esc(Object.entries(p.divisionHeads || {}).map(([e, gs]) => `${e} = ${gs.join(" ")}`).join("\n"))}</textarea></label>
      </div>
      <h3 style="margin-top:14px">Class advisors</h3>
      <div class="table-wrap"><table><thead><tr><th>Grade</th><th>Advisor (prints on the card)</th><th>Co-advisor</th><th>School emails of advisor and co-advisor, comma-separated</th></tr></thead><tbody>
        ${p.grades.map((g) => `<tr><td>${esc(g)}</td><td><input class="advName" data-g="${esc(g)}" value="${esc(p.advisors[g]?.name || "")}"></td><td><input class="advCo" data-g="${esc(g)}" value="${esc(p.advisors[g]?.co || "")}"></td><td><input class="advEmail" data-g="${esc(g)}" value="${esc(p.advisors[g]?.email || "")}" placeholder="advisor@zhongshischool.org, coadvisor@zhongshischool.org"></td></tr>`).join("")}
      </tbody></table></div>
      <button class="btn primary" id="saveSettings">Save term and people</button>
    </div>

    <div class="card"><h3>4 · Permissions</h3>
      <p class="small">Puts everyone who enters data into the SharePoint group <b>${esc(C.staffGroup)}</b>, then for every list: stops inheriting, removes site Members/Visitors, keeps Owners on Full Control, turns on version history and:</p>
      <ul class="small">
        <li><b>Students, Subjects, Award Catalogue, Settings</b>: staff group can <b>Read</b>.</li>
        <li><b>Awards, Advisor Comments, Marks G…</b>: staff group can <b>Contribute</b>, with <b>item-level privacy</b> (each person reads and edits only what they created).</li>
        <li><b>Marks G…</b>: that grade's advisor also gets <b>Design</b>, so they can see and correct the whole grade. Division heads get the same on their grades.</li>
        <li><b>Admin team</b> (step 3) gets <b>Design</b> on Marks, Advisor Comments and Awards, so they can read every grade to generate report cards.</li>
      </ul>
      <label><span>Staff who enter marks or awards, one school email per line (advisors are added automatically)</span><textarea id="staffEmails" rows="6" placeholder="name@zhongshischool.org">${esc((p.staff || []).join("\n"))}</textarea></label>
      <p><button class="btn primary" id="applyPerms">Apply permissions</button></p>
      <p class="small muted">Needs the SharePoint <code>AllSites.FullControl</code> delegated permission (admin consent) on the app registration. If it fails, do it by hand with the links below.</p>
      <pre id="permLog" class="small" style="white-space:pre-wrap"></pre>
      <details><summary class="small">Manual links</summary>
        <div class="table-wrap" style="margin-top:8px"><table><thead><tr><th>List</th><th>Should be</th><th></th></tr></thead><tbody>
        ${specs.filter((s) => S.lists[s.name]).map((s) => { const id = S.lists[s.name].id; return `<tr><td>${esc(s.name)}</td>
          <td class="small">${s.kind === "read" ? "Staff: Read" : `Staff: Contribute + item-level "Read/edit items created by the user"${s.kind === "marks" ? `; advisors (${esc(advisorEmails(s.grade).join(", ") || "?")}): Design` : ""}`}</td>
          <td class="small"><a target="_blank" rel="noopener" href="${esc(S.siteUrl)}/_layouts/15/user.aspx?obj=%7B${id}%7D,list&List=%7B${id}%7D">Permissions</a>
          ${s.kind !== "read" ? ` · <a target="_blank" rel="noopener" href="${esc(S.siteUrl)}/_layouts/15/advsetng.aspx?List=%7B${id}%7D">Advanced settings</a>` : ""}</td></tr>`; }).join("")}
        </tbody></table></div>
      </details>
    </div>

    <div class="card"><h3>5 · Subjects and awards during the year</h3>
      <p class="small muted">Electives and awards change. Add or hide them here; no re-import needed. Hiding never deletes marks or awards already given.
        To rename something, add the new name and hide the old one.</p>
      <h3 style="margin-top:12px">Subjects</h3>
      <div class="row" style="margin-bottom:8px">
        <label style="max-width:150px"><span>Grade</span><select id="mgGrade">${p.grades.map((g) => `<option ${g === S.mgGrade ? "selected" : ""}>${esc(g)}</option>`).join("")}</select></label>
      </div>
      <div id="mgSubjects"></div>
      <h3 style="margin-top:16px">Award catalogue</h3>
      <div class="row" style="margin-bottom:8px">
        <label style="max-width:260px"><span>Category</span><select id="mcCat"></select></label>
        <label style="flex:1;min-width:180px"><span>Find</span><input id="mcFind" placeholder="Award name"></label>
      </div>
      <div id="mgAwards"></div>
    </div>

    ${pdeSetupCard()}

    <div class="card"><h3>6 · Deployment check</h3>
      <p class="small">Redirect URI for this address (add it in Entra → App registration → Authentication → Single-page application): <code>${esc(redirectUri)}</code></p>
      <p class="small">Signed in as <code>${esc(S.me)}</code> · site <code>${esc(S.siteId)}</code></p>
    </div>`;

  $("#createLists").onclick = async () => {
    $("#createLists").disabled = true; const log = $("#createLog"); log.textContent = "";
    for (const s of missing) {
      log.textContent += `Creating ${s.name}… `;
      try {
        await graph("POST", `/sites/${S.siteId}/lists`, { displayName: s.name, list: { template: s.library ? "documentLibrary" : "genericList" }, ...(s.cols.length ? { columns: s.cols } : {}) }, SETUP_SCOPES);
        log.textContent += "done\n";
      } catch (e) { log.textContent += "FAILED: " + e.message + "\n"; }
    }
    for (const s of specs.filter((x) => S.lists[x.name] && !x.library)) {
      try {
        const cols = await graph("GET", `/sites/${S.siteId}/lists/${S.lists[s.name].id}/columns?$select=name`, null, SETUP_SCOPES);
        const have = new Set(cols.value.map((c) => c.name));
        const add = s.cols.filter((c) => !have.has(c.name));
        for (const c of add) await graph("POST", `/sites/${S.siteId}/lists/${S.lists[s.name].id}/columns`, c, SETUP_SCOPES);
        if (add.length) log.textContent += `${s.name}: added ${add.map((c) => c.name).join(", ")}\n`;
      } catch (e) { log.textContent += `${s.name}: column check FAILED: ${e.message}\n`; }
    }
    await loadAll(); log.textContent += "Finished. All lists are up to date.\n";
    setTimeout(render, 1500);
  };

  const imp = (id, fn) => { $(id).onchange = async (e) => {
    const f = e.target.files[0]; if (!f) return; const log = $("#importLog");
    try { log.textContent = `Reading ${f.name}…\n`; const wb = await readSheet(f); log.textContent += await fn(wb) + "\n"; await loadAll(); setTimeout(render, 1500); }
    catch (err) { log.textContent += "FAILED: " + err.message + "\n"; }
  }; };
  $("#fPhotos").onchange = async (e) => {
    const files = [...e.target.files]; if (!files.length) return; const log = $("#importLog");
    const known = new Set(S.students.map((s) => s.f.Title));
    const ok = files.filter((f) => known.has(f.name.replace(/\.[^.]+$/, ""))), skipped = files.filter((f) => !ok.includes(f));
    log.textContent = `Uploading ${ok.length} photos…${skipped.length ? ` (skipping ${skipped.length} not named after a Student ID: ${skipped.slice(0, 5).map((f) => f.name).join(", ")}${skipped.length > 5 ? "…" : ""})` : ""}\n`;
    try {
      const drive = await graph("GET", `/sites/${S.siteId}/lists/${S.lists[L.photos].id}/drive?$select=id`);
      let done = 0, failed = 0;
      const one = async (f) => {
        const r = await fetch(`${GRAPH}/drives/${drive.id}/root:/${encodeURIComponent(f.name.replace(/\.jpeg$/i, ".jpg"))}:/content`, { method: "PUT", headers: { Authorization: "Bearer " + await token(), "Content-Type": f.type || "image/jpeg" }, body: f });
        if (r.ok) done++; else failed++;
        if ((done + failed) % 20 === 0) log.textContent += `  …${done + failed} of ${ok.length}\n`;
      };
      for (let i = 0; i < ok.length; i += 4) await Promise.all(ok.slice(i, i + 4).map(one));   // four at a time
      await loadPhotos();
      log.textContent += `Photos: ${done} uploaded, ${failed} failed. ${Object.keys(S.photos).length} bobaes now have a photo.\n`;
    } catch (err) { log.textContent += "FAILED: " + err.message + "\n"; }
  };
  imp("#fStudents", async (wb) => {
    const recs = rowsOf(wb, /student/i).map((r) => ({
      Title: col(r, "Student_ID", "StudentID", "ID"), FullName: col(r, "Full_Name", "Name"), CalledName: col(r, "Called_Name", "Nickname"),
      CertificateName: col(r, "Certificate_Name"), Grade: col(r, "Grade"), Email: lc(col(r, "Email")), Department: col(r, "Department"),
      Joined: col(r, "Joined"), Status: col(r, "Status") === "Withdrawn" ? "Withdrawn" : "Active",
      ...((d) => d ? { DOB: d } : {})(col(r, "DOB", "Date_of_Birth", "Date of birth", "BOD")) })).filter((r) => r.Title && r.FullName);
    const unknownGrades = [...new Set(recs.map((r) => r.Grade).filter((g) => !P().grades.includes(g)))];
    const dupes = [...new Set(recs.map((r) => r.Title).filter((id, i, a) => a.indexOf(id) !== i))];
    if (dupes.length) throw new Error(`The same Student ID appears more than once: ${dupes.map((id) => `${id} (${recs.filter((r) => r.Title === id).map((r) => r.Grade).join(" and ")})`).join(", ")}. Fix the workbook so each bobae has one row, then import again. Nothing was imported.`);
    const { sent, failed } = await upsert(L.students, recs, (f) => f.Title, S.students);
    return `Roster: ${recs.length} students read, ${sent} added/updated, ${failed.length} failed.` + (unknownGrades.length ? ` Note: grades not in the app's grade list: ${unknownGrades.join(", ")}.` : "");
  });
  imp("#fSubjects", async (wb) => {
    // Accepts one sheet with a Grade column (Upper School file) or one sheet per grade named "G1", "Grade 1"… (Primary file).
    const gradeOf = (s) => { const t = String(s || "").replace(/\s+/g, "").replace(/^grade/i, "G").toUpperCase(); return P().grades.find((g) => g.toUpperCase() === t || g.toUpperCase() === "G" + t) || ""; };
    // Same list as the report card generator (Subject II), matched on whole words, plus choir.
    const extracurricular = /\b(pe|p\.e\.|physical education|drama|orchestra|choir|debate|well-?being|pshe|careers|electives?|ecc|sports?|board games|food tech(nology)?|homeroom)\b/i;
    const seen = new Set(), recs = [], skipped = new Set();
    let order = 0;
    for (const sheet of wb.SheetNames) {
      const sheetGrade = gradeOf(sheet);
      for (const r of XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: "" })) {
        const grade = gradeOf(col(r, "Grade", "Class")) || sheetGrade;
        const name = col(r, "Suggested canonical name", "Subject", "Subjects", "Subject Name", "Subject (as used in Spring 2026)").replace(/\s+/g, " ");
        if (!name) continue;
        if (!grade) { skipped.add(sheet); continue; }
        const key = grade + "|" + lc(name); if (seen.has(key)) continue; seen.add(key);
        const typeCell = col(r, "Subject I / II", "Subject_Type", "Type", "Subject Type");
        recs.push({ Title: name, Grade: grade, SubjectType: typeCell ? (/ii|2|extra/i.test(typeCell) ? "II" : "I") : (extracurricular.test(name) ? "II" : "I"),
          Teacher: col(r, "Teacher for Fall 2026", "Teacher"), PrintsOnCard: !/^no/i.test(col(r, "Prints on report card?", "PrintsOnCard")), SortOrder: ++order });
      }
    }
    if (!recs.length) throw new Error(`No subjects found. Each sheet needs a "Subject" column, and either a "Grade" column or a sheet name like G1.`);
    const { sent, failed } = await upsert(L.subjects, recs, (f) => f.Grade + "|" + lc(f.Title), S.subjects);
    const grades = [...new Set(recs.map((r) => r.Grade))];
    return `Subjects: ${recs.length} read for ${grades.join(", ")} (${recs.filter((r) => r.SubjectType === "II").length} marked extracurricular), ${sent} added/updated, ${failed.length} failed.`
      + (skipped.size ? ` Skipped sheets with no grade: ${[...skipped].join(", ")}.` : "") + " Check the I/II types on the Subjects screen below.";
  });
  imp("#fCatalogue", async (wb) => {
    const all = await readList(L.catalogue);
    // Tidy: collapse odd spaces (incl. full-width), drop case-only duplicates ("Most improved in ICT" twice).
    const clean = (s) => String(s || "").replace(/[\s　]+/g, " ").trim();
    const seen = new Set(), dupes = [];
    const recs = rowsOf(wb, /choice list/i).map((r, i) => ({
      Title: clean(col(r, "Award name", "Award name (cleaned)", "Award_Name")), Category: clean(col(r, "Category")),
      CertificateType: ((t) => P().certificateTypes.find((x) => lc(x) === lc(t)) || t)(clean(col(r, "Certificate type", "Certificate_Type"))),
      Season: ((s) => ["Fall", "Spring"].find((x) => lc(x) === lc(s)) || "Both")(clean(col(r, "Season", "Semester"))),
      Criteria: clean(col(r, "Criteria")), GivenBy: /advisor/i.test(col(r, "Given by")) ? "Advisors" : "", AskMonth: /^y/i.test(col(r, "Ask for month")), CustomTitle: /^y/i.test(col(r, "Custom title")), Active: true, SortOrder: i + 1 }))
      .filter((r) => { if (!r.Title || !r.Category) return false; const k = lc(r.Category + "|" + r.Title); if (seen.has(k)) { dupes.push(r.Title); return false; } seen.add(k); return true; });
    const oddTypes = [...new Set(recs.map((r) => r.CertificateType).filter((t) => !P().certificateTypes.includes(t)))];
    const keyOf = (f) => lc(clean(f.Category) + "|" + clean(f.Title));
    // Retire extra copies already in SharePoint, and awards no longer in the file (kept for history, hidden from teachers).
    const first = new Map(); for (const x of all) if (!first.has(keyOf(x.f))) first.set(keyOf(x.f), x.id);
    const retire = all.filter((x) => x.f.Active !== false && (!seen.has(keyOf(x.f)) || first.get(keyOf(x.f)) !== x.id));
    const { sent, failed } = await upsert(L.catalogue, recs, keyOf, all.filter((x) => first.get(keyOf(x.f)) === x.id));
    if (retire.length) await batch(retire.map((x) => ({ method: "PATCH", url: `${itemsUrl(L.catalogue)}/${x.id}/fields`, body: { Active: false } })));
    return `Award catalogue: ${recs.length} awards in ${new Set(recs.map((r) => r.Category)).size} categories, ${sent} added/updated, ${retire.length} retired, ${failed.length} failed.`
      + (dupes.length ? ` Skipped ${dupes.length} duplicate names (same award, different capitals).` : "")
      + (oddTypes.length ? ` Check certificate type: ${oddTypes.join(", ")}.` : "");
  });

  $("#saveSettings").onclick = async () => {
    const lines = (id) => $(id).value.split(/[\s,;]+/).map(lc).filter(Boolean);
    const advisors = {};
    for (const g of P().grades) advisors[g] = { name: $(`.advName[data-g="${g}"]`).value.trim(), co: $(`.advCo[data-g="${g}"]`).value.trim(), email: emailsOf($(`.advEmail[data-g="${g}"]`).value).join(", ") };
    const divisionHeads = {};
    for (const line of $("#sHeads").value.split("\n")) {
      const [e, gs] = line.split("="); const email = lc(e); if (!email.includes("@")) continue;
      divisionHeads[email] = String(gs || "").split(/[\s,;]+/).map((x) => x.trim().toUpperCase()).filter((x) => P().grades.includes(x));
    }
    const json = { academicYear: $("#sYear").value.trim(), semester: $("#sSem").value, period: $("#sPeriod").value, admins: lines("#sAdmins"), leadership: lines("#sLead"),
      reportTeam: lines("#sReport"), divisionHeads, advisors };
    if (!json.admins.includes(S.me) && !confirm("You are removing yourself from the admins. Continue?")) return;
    try {
      if (S.settingsItem) await graph("PATCH", `${itemsUrl(L.settings)}/${S.settingsItem.id}/fields`, { Json: JSON.stringify(json, null, 1) });
      else await graph("POST", itemsUrl(L.settings), { fields: { Title: "config", Json: JSON.stringify(json, null, 1) } });
      await loadAll();
      if (S.pde.loaded) { try { await savePdeSettings(); await loadPde(); } catch (_) {} }   // bobaes read the term from the PDE site
      render(); flash("ok", "Saved. Re-run Permissions if advisors changed.");
    } catch (e) { flash("error", e.message); }
  };

  $("#applyPerms").onclick = async () => {
    const log = $("#permLog"); const out = (s) => (log.textContent += s + "\n");
    log.textContent = ""; $("#applyPerms").disabled = true;
    try { await applyPermissions(specs.filter((s) => S.lists[s.name]), $("#staffEmails").value.split(/[\s,;]+/).map(lc).filter((e) => e.includes("@")), out); out("Done. Check one list's Permissions page to confirm."); }
    catch (e) { out("FAILED: " + e.message); out("Use the manual links below instead."); }
    $("#applyPerms").disabled = false;
  };
  bindYearTools();
  bindPdeSetup();
}

// Setup step 5: maintain subjects and the award catalogue in place.
function bindYearTools() {
  const p = P();
  S.mgGrade = $("#mgGrade").value;
  const drawSubjects = () => {
    const g = S.mgGrade;
    const list = S.subjects.filter((s) => s.f.Grade === g);
    $("#mgSubjects").innerHTML = `
      <div class="table-wrap"><table><thead><tr><th>Subject</th><th>Type</th><th class="center">Shown to teachers</th></tr></thead><tbody>
        ${list.map((s) => `<tr><td>${esc(s.f.Title)}</td><td><select data-stype="${s.id}" style="width:190px"><option value="I">I · academic</option><option value="II" ${s.f.SubjectType === "II" ? "selected" : ""}>II · extracurricular</option></select></td>
          <td class="center"><input type="checkbox" data-subj="${s.id}" ${s.f.PrintsOnCard !== false ? "checked" : ""}></td></tr>`).join("")
          || `<tr><td colspan="3" class="muted">No subjects for ${esc(g)} yet.</td></tr>`}
      </tbody></table></div>
      <div class="row">
        <label style="flex:1;min-width:200px"><span>New subject or elective for ${esc(g)}</span><input id="nsName" maxlength="80" placeholder="e.g. Elective (Robotics)"></label>
        <label style="max-width:190px"><span>Type</span><select id="nsType"><option value="I">I · academic</option><option value="II">II · extracurricular</option></select></label>
        <button class="btn" id="nsAdd">Add</button>
      </div>`;
    $$("[data-subj]").forEach((c) => (c.onchange = async () => {
      try { await graph("PATCH", `${itemsUrl(L.subjects)}/${c.dataset.subj}/fields`, { PrintsOnCard: c.checked }); S.subjects.find((s) => s.id === c.dataset.subj).f.PrintsOnCard = c.checked; flash("ok", "Saved."); }
      catch (e) { c.checked = !c.checked; flash("error", e.message); }
    }));
    $$("[data-stype]").forEach((sel) => (sel.onchange = async () => {
      try { await graph("PATCH", `${itemsUrl(L.subjects)}/${sel.dataset.stype}/fields`, { SubjectType: sel.value }); S.subjects.find((s) => s.id === sel.dataset.stype).f.SubjectType = sel.value; flash("ok", "Saved. Classes created from now on use the new type."); }
      catch (e) { sel.value = sel.value === "I" ? "II" : "I"; flash("error", e.message); }
    }));
    $("#nsAdd").onclick = async () => {
      const name = $("#nsName").value.replace(/\s+/g, " ").trim(); if (!name) return;
      if (list.some((s) => lc(s.f.Title) === lc(name))) return flash("warn", `${g} already has "${name}". Tick it to show it again.`);
      try {
        const it = await graph("POST", itemsUrl(L.subjects), { fields: { Title: name, Grade: g, SubjectType: $("#nsType").value, PrintsOnCard: true,
          SortOrder: Math.max(0, ...S.subjects.map((s) => s.f.SortOrder || 0)) + 1 } });
        S.subjects.push(norm(it)); drawSubjects(); flash("ok", `Added "${name}" to ${g}.`);
      } catch (e) { flash("error", e.message); }
    };
  };
  $("#mgGrade").onchange = (e) => { S.mgGrade = e.target.value; drawSubjects(); };
  drawSubjects();

  const all = S.catalogueAll || [];
  const cats = [...new Set(all.map((c) => c.f.Category).filter(Boolean))].sort();
  $("#mcCat").innerHTML = `<option value="">All categories (${all.length})</option>` + cats.map((c) => `<option ${c === S.mcCat ? "selected" : ""}>${esc(c)}</option>`).join("");
  const drawAwards = () => {
    const q = lc($("#mcFind").value);
    const list = all.filter((c) => (!S.mcCat || c.f.Category === S.mcCat) && (!q || lc(c.f.Title).includes(q)));
    $("#mgAwards").innerHTML = `
      <div class="table-wrap" style="max-height:420px;overflow:auto"><table><thead><tr><th>Award</th><th>Category</th><th>Certificate type</th><th>Season</th><th class="center">Advisors only</th><th class="center">Shown</th></tr></thead><tbody>
        ${list.map((c) => `<tr><td>${esc(c.f.Title)}</td><td class="small">${esc(c.f.Category)}</td><td class="small">${esc(c.f.CertificateType || "")}</td>
          <td><select data-season="${c.id}" style="width:100px">${["Both", "Fall", "Spring"].map((s) => `<option ${(c.f.Season || "Both") === s ? "selected" : ""}>${s}</option>`).join("")}</select></td>
          <td class="center"><input type="checkbox" data-advonly="${c.id}" ${advisorOnly(c) ? "checked" : ""}></td>
          <td class="center"><input type="checkbox" data-active="${c.id}" ${c.f.Active !== false ? "checked" : ""}></td></tr>`).join("")
          || `<tr><td colspan="6" class="muted">Nothing matches.</td></tr>`}
      </tbody></table></div>
      <div class="grid" style="align-items:end">
        <label><span>Category (pick or type a new one)</span><input id="naCat" list="naCats" value="${esc(S.mcCat || "")}"><datalist id="naCats">${cats.map((c) => `<option value="${esc(c)}">`).join("")}</datalist></label>
        <label><span>New award name</span><input id="naName" maxlength="160"></label>
        <label><span>Certificate type</span><select id="naType">${p.certificateTypes.map((t) => `<option>${esc(t)}</option>`).join("")}</select></label>
        <label><span>Season</span><select id="naSeason"><option>Both</option><option>Fall</option><option>Spring</option></select></label>
        <div><button class="btn" id="naAdd">Add award</button></div>
      </div>`;
    const patch = async (id, fields, undo) => {
      try { await graph("PATCH", `${itemsUrl(L.catalogue)}/${id}/fields`, fields); Object.assign(all.find((c) => c.id === id).f, fields); S.catalogue = teacherCatalogue(); flash("ok", "Saved."); }
      catch (e) { undo(); flash("error", e.message); }
    };
    $$("[data-active]").forEach((c) => (c.onchange = () => patch(c.dataset.active, { Active: c.checked }, () => (c.checked = !c.checked))));
    $$("[data-advonly]").forEach((c) => (c.onchange = () => patch(c.dataset.advonly, { GivenBy: c.checked ? "Advisors" : "" }, () => (c.checked = !c.checked))));
    $$("[data-season]").forEach((s) => { const was = s.value; s.onchange = () => patch(s.dataset.season, { Season: s.value }, () => (s.value = was)); });
    $("#naAdd").onclick = async () => {
      const cat = $("#naCat").value.replace(/\s+/g, " ").trim(), name = $("#naName").value.replace(/\s+/g, " ").trim();
      if (!cat || !name) return flash("warn", "Give the award a category and a name.");
      if (all.some((c) => lc(c.f.Category) === lc(cat) && lc(c.f.Title) === lc(name))) return flash("warn", `"${name}" is already in ${cat}. Tick "Shown" if it's hidden.`);
      try {
        const it = await graph("POST", itemsUrl(L.catalogue), { fields: { Title: name, Category: cat, CertificateType: $("#naType").value, Season: $("#naSeason").value, Active: true,
          SortOrder: Math.max(0, ...all.map((c) => c.f.SortOrder || 0)) + 1 } });
        all.push(norm(it)); S.catalogue = teacherCatalogue();
        S.mcCat = cat; bindYearTools(); flash("ok", `Added "${name}".`);
      } catch (e) { flash("error", e.message); }
    };
  };
  $("#mcCat").onchange = (e) => { S.mcCat = e.target.value; drawAwards(); };
  $("#mcFind").oninput = drawAwards;
  drawAwards();
}
const claims = (email) => `i:0#.f|membership|${email}`;
async function applyPermissions(specs, staffEmails, out) {
  const roles = (await sp("GET", "web/roledefinitions?$select=Id,RoleTypeKind")).value;
  const role = (kind) => roles.find((r) => r.RoleTypeKind === kind).Id; // 2 Reader, 3 Contributor, 4 Web Designer, 5 Full Control
  const owners = await sp("GET", "web/associatedownergroup?$select=Id,Title");
  const members = await sp("GET", "web/associatedmembergroup?$select=Id").catch(() => null);
  const visitors = await sp("GET", "web/associatedvisitorgroup?$select=Id").catch(() => null);
  let group = await sp("GET", `web/sitegroups/getbyname('${encodeURIComponent(C.staffGroup)}')?$select=Id`).catch(() => null);
  if (!group) { group = await sp("POST", "web/sitegroups", { __metadata: { type: "SP.Group" }, Title: C.staffGroup, Description: "Teachers who enter marks and awards in WZFS Reporting" }); out(`Created group ${C.staffGroup}.`); }
  // Admin team + Principal: Design on marks, comments and awards (read every grade; the Principal can also correct).
  const reportTeam = [...new Set([...(P().reportTeam || []), ...P().admins, ...(P().principals || [])].map(lc))];
  const heads = Object.entries(P().divisionHeads || {}).map(([e, gs]) => [lc(e), gs]);
  const librarians = (P().librarians || []).map(lc);
  const everyone = [...new Set([...staffEmails, ...P().grades.flatMap(advisorEmails), ...reportTeam, ...heads.map(([e]) => e), ...librarians])];
  const userIds = {};
  const userId = async (e) => (userIds[e] = userIds[e] || (await sp("POST", "web/ensureuser", { logonName: claims(e) })).Id);
  for (const e of everyone) {
    try { await sp("POST", `web/sitegroups(${group.Id})/users`, { __metadata: { type: "SP.User" }, LoginName: claims(e) }); }
    catch (err) { out(`  ! Could not add ${e}: ${err.message}`); }
  }
  out(`${everyone.length} people are in ${C.staffGroup}.`);
  for (const s of specs) {
    const id = S.lists[s.name].id, base = `web/lists(guid'${id}')`;
    await sp("POST", `${base}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`);
    for (const g of [members, visitors]) if (g) await sp("POST", `${base}/roleassignments/getbyprincipalid(${g.Id})`, null, { "X-HTTP-Method": "DELETE" }).catch(() => {});
    await sp("POST", `${base}/roleassignments/addroleassignment(principalid=${owners.Id},roledefid=${role(5)})`);
    await sp("POST", `${base}/roleassignments/addroleassignment(principalid=${group.Id},roledefid=${role(s.kind === "read" ? 2 : 3)})`);
    const props = { __metadata: { type: "SP.List" }, EnableVersioning: true };
    if (s.kind !== "read") Object.assign(props, { ReadSecurity: 2, WriteSecurity: 2 });
    await sp("POST", base, props, { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" });
    let extra = "";
    const design = async (emails) => { for (const e of emails) await sp("POST", `${base}/roleassignments/addroleassignment(principalid=${await userId(e)},roledefid=${role(4)})`); };
    if (s.kind === "marks") {
      const advs = advisorEmails(s.grade);
      await design(advs);
      extra = advs.length ? ` · advisors ${advs.join(", ")}: Design` : " · no advisor email set";
      const hs = heads.filter(([, gs]) => gs.includes(s.grade)).map(([e]) => e);
      if (hs.length) { await design(hs); extra += ` · division head ${hs.join(", ")}: Design`; }
    }
    if (s.name === L.settings) {   // the go-ahead (Principal, division heads) is saved here
      const approvers = [...new Set([...(P().principals || []), ...heads.map(([e]) => e)].map(lc))];
      for (const e of approvers) await sp("POST", `${base}/roleassignments/addroleassignment(principalid=${await userId(e)},roledefid=${role(3)})`);
      if (approvers.length) extra += ` · Principal and division heads (${approvers.length}): Contribute (go-ahead)`;
    }
    if (s.name === L.awards) {   // division heads review every award in their grades; the librarian ticks certificates as printed
      const more = [...new Set([...heads.map(([e]) => e), ...librarians])].filter((e) => !reportTeam.includes(e));
      await design(more); if (more.length) extra += ` · Division heads and librarian (${more.length}): Design`;
    }
    if (s.kind !== "read") {   // Admin team reads every grade's marks, comments and awards to make report cards
      await design(reportTeam); extra += ` · Admin team (${reportTeam.length}): Design`;
      if (s.name === L.comments && heads.length) { await design(heads.map(([e]) => e)); extra += " · division heads: Design"; }
    }
    out(`✓ ${s.name}: Owners Full Control · ${C.staffGroup} ${s.kind === "read" ? "Read" : "Contribute, item-level"}${extra}`);
  }
}

// ── PDE: Academic Purpose-Driven Education (work in progress) ─
// A design preview for the teacher committee. Nothing here reads or writes SharePoint;
// every number on these pages is an invented example. Build the real lists only once the
// committee has agreed the workflow (see "Questions for the committee").
const PDE_VALUES = [
  { v: "Responsibility", ind: "Attendance", by: "Subject teachers + advisor",
    r: ["Misses significant days without notice; submits many catch-up forms.", "Misses some days; communicates occasionally; submits several catch-up forms.", "Consistent attendance; communicates absences; submits few catch-up forms.", "Perfect attendance; proactive in notifying; submits no catch-up forms."] },
  { v: "Respect", ind: "Attitude and conduct", by: "Subject teachers + advisor",
    r: ["Disruptive behaviour; often disrespectful.", "Sometimes respectful; improving conduct.", "Respectful; participates constructively.", "Highly respectful; leads others in positive conduct."] },
  { v: "Self-expression", ind: "Arts and performance activities", by: "Advisor",
    r: ["Limited creativity; follows instructions only.", "Shows some creativity; needs guidance.", "Demonstrates originality and personal style.", "Highly creative; pushes boundaries of self-expression."] },
  { v: "Holistic Education", ind: "Participation in special events", by: "Advisor",
    r: ["Seldom joins in; lacks awareness of activities.", "Participates occasionally; shows emerging interest.", "Regular participation; values learning opportunities.", "Proactively engages; inspires others to join."] },
  { v: "Wisdom", ind: "Physical activities", by: "Advisor",
    r: ["Disengaged; lacks commitment to physical health.", "Participates but shows little initiative.", "Committed to personal fitness; encourages others.", "Advocates for health; leads initiatives in physical wellness."] },
];
const PDE_SECTIONS = [["goals", "Bobae goals (live)"], ["overview", "Overview"], ["cycle", "How the year works"], ["roles", "Who does what"], ["screens", "Screens (preview)"], ["values", "Core values rubric"], ["teacher", "Teacher goals"], ["questions", "Questions for the committee"]];
const PDE_QUESTIONS = [
  { q: "What is PDE mainly for?",
    why: "Every other answer follows from this one.",
    options: ["<b>Academic achievement:</b> goals are mostly target scores in subjects.",
      "<b>Personal character:</b> goals are mostly about values and personal development.",
      "<b>Deliberately both:</b> academic goals and character goals, side by side."] },
  { q: "What should students set goals for?",
    options: ["Only the four core subjects (Chinese, English, Maths and Science)?",
      "Or other subjects too, such as ICT, Art and PE?"] },
  { q: "Should students set goals for core values?",
    options: ["Should students have separate goals for values and personal development, such as responsibility, kindness, perseverance and service?",
      "If so, which values should the school use? (The current PDE template rates the five school values: Wisdom, Holistic Education, Self-expression, Responsibility and Respect.)"] },
  { q: "When should students write their goals?",
    options: ["At the beginning of each semester?", "At the beginning of each month?", "After receiving mid-term results?", "At another point in the school calendar?"] },
  { q: "What should a goal and action plan contain?",
    options: ["A target score, a personal goal and specific actions? Some of these, or something else?",
      "Should expectations differ between Primary, Middle and High School?"] },
  { q: "How often should progress be reviewed, and by whom?",
    options: ["Monthly, at mid-term, or at both points?", "Who reviews and comments on progress: the subject teacher, the advisor, or both?"] },
  { q: "How should achievement be measured?",
    options: ["Through scores, teacher feedback, demonstrated behaviour, completed actions, or a combination of these?"] },
  { q: "What happens when a student reaches a goal, or doesn't?",
    options: ["When a goal is reached: set a new target, extend the existing goal, or maintain the achievement?",
      "When a goal is not reached: carry it forward with a new plan, or start fresh next semester?"] },
  { q: "How should parents and school leadership be involved?",
    options: ["Should parents see goals and progress online, or during PTC meetings?",
      "What overview should advisors, division heads and the Director receive?"],
    body: `<details class="small"><summary>Ways to let parents see only their own child (for reference)</summary><ul>
      <li><b>Printed or PDF page</b> each month or semester, given by the advisor or at the PTC meeting. No online access needed.</li>
      <li><b>Through the bobae's own login</b> at home, with the parent looking alongside their child.</li>
      <li><b>A parent list</b> (Student ID → parent email) kept by the Admin Office, so a parent's sign-in shows only their own children.</li></ul></details>` },
];
// A starting point the committee can accept, change or reject. Not school policy.
const PDE_FIRST_STEP = [
  "<b>G1–G5:</b> advisors help students set age-appropriate goals.",
  "<b>G6–G12:</b> students write their own goals and action plans, starting with the four core subjects.",
  "<b>Other subjects and core values:</b> left open for the committee to decide.",
  "<b>Timing:</b> one school-wide cycle, e.g. set goals at the start of each semester, review them at mid-term and at the end of the semester.",
];
function pdeBand(rate) { return rate >= 100 ? ["Excellent", "ok"] : rate >= 90 ? ["Satisfactory", "warn"] : ["Needs improvement", "bad"]; }
function pdeGoalCard(subject, target, months, plan) {
  const done = months.filter((m) => m[1] !== null);
  const latest = done.length ? done[done.length - 1][1] : null;
  const rate = latest === null ? null : Math.round((latest / target) * 100);
  const [band, cls] = rate === null ? ["Not started", ""] : pdeBand(rate);
  return `<div class="card" style="margin:0">
    <div class="row" style="justify-content:space-between;align-items:center"><h3 style="margin:0">${esc(subject)}</h3><span class="badge ${cls}">${band}</span></div>
    <p class="small muted" style="margin:4px 0 8px">Target <b>${target}%</b> · latest ${latest === null ? "–" : `<b>${latest}%</b>`} · achievement rate ${rate === null ? "–" : `<b>${rate}%</b>`}</p>
    <div style="display:flex;gap:6px;align-items:flex-end;height:70px;border-bottom:1px solid var(--border);position:relative">
      <div title="Target" style="position:absolute;left:0;right:0;bottom:${target * 0.7}px;border-top:2px dashed var(--accent)"></div>
      ${months.map(([m, s]) => `<div style="flex:1;text-align:center"><div style="height:${(s || 0) * 0.7}px;background:${s === null ? "transparent" : "var(--accent)"};opacity:${s === null ? 1 : .75};border-radius:3px 3px 0 0"></div></div>`).join("")}
    </div>
    <div style="display:flex;gap:6px" class="small muted">${months.map(([m]) => `<div style="flex:1;text-align:center">${m}</div>`).join("")}</div>
    <p class="small" style="margin:8px 0 0"><b>My action plan:</b> ${esc(plan)}</p>
  </div>`;
}
function viewPde(main) {
  const sec = PDE_SECTIONS.some(([k]) => k === S.pdeSec) ? S.pdeSec : (S.pde && S.pde.loaded ? "goals" : "overview");
  const ex = `<span class="badge">Example data</span>`;
  const body = {
    overview: `
      <div class="card pde-hero">
        <div class="big">🌱</div>
        <h1>Academic Purpose-Driven Education</h1>
        <p><span class="badge warn">Work in progress</span> <span class="badge">Being designed with the teacher committee</span></p>
        <p class="lede" style="max-width:680px;margin:10px auto">Every bobae sets a goal for each subject on the first day of school and works on an action plan with their teacher.
          Teachers record progress through the year, so the bobae, their parents and the Director can all see whether they're on track.</p>
      </div>
      <div class="stats">
        <div class="stat"><div class="k">Day 1</div><div class="v" style="font-size:16px">Goal + action plan</div></div>
        <div class="stat"><div class="k">Every month</div><div class="v" style="font-size:16px">Score + feedback</div></div>
        <div class="stat"><div class="k">Each semester</div><div class="v" style="font-size:16px">Core values 1–4</div></div>
        <div class="stat"><div class="k">End of year</div><div class="v" style="font-size:16px">Goal achievement ratio</div></div>
      </div>
      <div class="card"><h3>What this section is for</h3>
        <p class="small">These pages show how PDE <i>could</i> work in this app, so the committee can react to something concrete. Nothing here saves data.
          Once the committee agrees the workflow, PDE will use the same sign-in, roster and marks as the rest of the app, so <b>no score is ever typed twice</b>.</p>
        <p class="small">Start with <a href="#" data-pde="cycle">How the year works</a>, then go through <a href="#" data-pde="questions">Questions for the committee</a>.</p>
      </div>`,
    cycle: `
      <h2>How the year works ${ex}</h2>
      <div class="table-wrap"><table>
        <thead><tr><th>When</th><th>Bobae</th><th>Subject teacher</th><th>Advisor</th><th>Parents · Director</th></tr></thead>
        <tbody>
          <tr><td><b>Day 1</b></td><td>Sets a <b>target score</b> and <b>action plan</b> for each core subject</td><td>Agrees the goal; adds their part of the action plan</td><td>Enters goals for G1–G5</td><td>—</td></tr>
          <tr><td><b>Every month</b><br><span class="small muted">Sep–Dec, Feb–Jun</span></td><td>Checks whether they're on track; writes a short reflection</td><td><b>Core subjects</b> (Chinese, English, Maths, Science) record the month's score and a line of feedback</td><td>Follows up with bobaes who are off track</td><td>Parents see progress and can reply</td></tr>
          <tr><td><b>Each semester</b></td><td>—</td><td>Rate <b>Responsibility</b> and <b>Respect</b> (1–4)</td><td>Rate <b>all five core values</b> (1–4)</td><td>Director sees the overview</td></tr>
          <tr><td><b>End of semester / year</b></td><td>Final reflection</td><td>Final score (from the report card marks)</td><td>Confirms the goal achievement ratio</td><td>Goals met out of goals set, by class and school</td></tr>
        </tbody></table></div>
      <div class="card"><h3>Teachers of ICT, Art, PE and similar subjects</h3>
        <p class="small">They teach every class, so instead of a goal per bobae they write <b>one teacher goal and action plan per division</b> (Primary, Middle School, IGCSE). See <a href="#" data-pde="teacher">Teacher goals</a>.</p></div>
      <div class="card"><h3>Achievement rate (proposal)</h3>
        <p class="small"><b>Achievement rate = latest score ÷ target score.</b> For example, a target of 85% and a score of 80% gives 94%.</p>
        <p class="small">Bands: <span class="badge ok">Excellent</span> 100% or more · <span class="badge warn">Satisfactory</span> 90–99% · <span class="badge bad">Needs improvement</span> below 90%. The committee sets the real thresholds.</p></div>`,
    roles: `
      <h2>Who does what</h2>
      <div class="table-wrap"><table>
        <thead><tr><th>Who</th><th>Enters</th><th>Sees</th></tr></thead>
        <tbody>
          <tr><td><b>Bobae (G6–G12)</b></td><td>Own goals, action plans and monthly reflections</td><td>Own goals and progress only</td></tr>
          <tr><td><b>Bobae (G1–G5)</b></td><td>Nothing: the advisor enters for them</td><td>—</td></tr>
          <tr><td><b>Core subject teacher</b><br><span class="small muted">Chinese, English (all types), Maths, Science</span></td><td>Monthly score and feedback for their own classes; Responsibility and Respect ratings</td><td>Their own bobaes, their own subject</td></tr>
          <tr><td><b>ICT, Art, PE … teacher</b></td><td>Teacher goal and action plan per division; Responsibility and Respect ratings</td><td>Their own classes</td></tr>
          <tr><td><b>Class advisor / co-advisor</b></td><td>All five core values; goals for G1–G5; follow-up notes</td><td>Everything for their class</td></tr>
          <tr><td><b>Parent</b></td><td>Feedback</td><td>Their own child's progress</td></tr>
          <tr><td><b>Director / leadership</b></td><td>Nothing</td><td>Whole-school overview: goals met, core values, bobaes off track</td></tr>
        </tbody></table></div>
      <p class="small muted">Privacy works like the marks lists: each person can only open what SharePoint lets them, and bobaes never get access to a list that holds other bobaes' marks.</p>`,
    screens: `
      <h2>What a bobae would see ${ex}</h2>
      <p class="small muted">Example: a Grade 8 bobae in October. Bars are monthly scores; the dashed line is their target.</p>
      <div class="class-grid" style="grid-template-columns:repeat(auto-fill,minmax(280px,1fr))">
        ${pdeGoalCard("Maths", 85, [["Sep", 78], ["Oct", 84], ["Nov", null], ["Dec", null]], "Do 10 practice questions every Tuesday and Thursday; ask for help after class when I'm stuck.")}
        ${pdeGoalCard("English", 80, [["Sep", 72], ["Oct", 70], ["Nov", null], ["Dec", null]], "Read 20 minutes a day and write one paragraph each week for feedback.")}
        ${pdeGoalCard("Science", 75, [["Sep", 79], ["Oct", 82], ["Nov", null], ["Dec", null]], "Make flashcards for each topic and review them before every quiz.")}
        ${pdeGoalCard("Chinese", 90, [["Sep", 88], ["Oct", 91], ["Nov", null], ["Dec", null]], "Learn 15 new characters a week with my study group.")}
      </div>
      <h2>What a core subject teacher would enter each month ${ex}</h2>
      <div class="table-wrap"><table class="entry">
        <thead><tr><th>Bobae</th><th class="num">Target</th><th class="num">Sep</th><th class="num">Oct</th><th class="num">Rate</th><th>Feedback (October)</th></tr></thead>
        <tbody>
          ${[["Breck", 85, 78, 84], ["Coleta", 70, 62, 66], ["Dennis", 90, 91, 93]].map(([n, t, s, o]) => { const r = Math.round((o / t) * 100), [b, c] = pdeBand(r);
            return `<tr><td><b>${n}</b></td><td class="num">${t}%</td><td class="num">${s}</td><td class="num"><input class="score" value="${o}" disabled></td><td class="num"><span class="badge ${c}">${r}%</span></td><td><input disabled placeholder="One line of feedback"></td></tr>`; }).join("")}
        </tbody></table></div>
      <h2>What the Director would see ${ex}</h2>
      <div class="stats">
        <div class="stat"><div class="k">Goals on track (school)</div><div class="v">71%</div></div>
        <div class="stat"><div class="k">Bobaes off track in 2+ subjects</div><div class="v">14</div></div>
        <div class="stat"><div class="k">Average core values (out of 20)</div><div class="v">15.2</div></div>
        <div class="stat"><div class="k">Monthly scores entered</div><div class="v">88%</div></div>
      </div>
      <div class="table-wrap"><table>
        <thead><tr><th>Grade</th><th class="num">Goals on track</th><th class="num">Off track</th><th class="num">Core values /20</th></tr></thead>
        <tbody>${[["G6", 74, 2, 15.8], ["G7", 68, 3, 14.9], ["G8", 70, 4, 15.1], ["G9-1", 66, 5, 14.6]].map(([g, a, b, c]) => `<tr><td>${g}</td><td class="num">${a}%</td><td class="num">${b}</td><td class="num">${c}</td></tr>`).join("")}</tbody></table></div>
      <div class="two-col" style="grid-template-columns:1fr 1fr">
        <div><h3>Subject ranking: goals met</h3><div class="table-wrap"><table><thead><tr><th>#</th><th>Subject</th><th class="num">Goals met</th><th class="num">Avg rate</th></tr></thead><tbody>
          ${[["Chinese", 78, 101], ["Science", 72, 98], ["Maths", 64, 95], ["English", 51, 91]].map(([s, m, r], i) => `<tr><td>${i + 1}</td><td>${s}</td><td class="num">${m}%</td><td class="num">${r}%</td></tr>`).join("")}</tbody></table></div></div>
        <div><h3>Help needed</h3><div class="table-wrap"><table><thead><tr><th>Bobae</th><th>Grade</th><th>Off track in</th></tr></thead><tbody>
          ${[["Coleta", "G8", "English, Science"], ["Ethan", "G9-1", "Maths, English"], ["Lily", "G9-1", "Maths (2 months)"]].map(([n, g, s]) => `<tr><td>${n}</td><td>${g}</td><td class="small">${s}</td></tr>`).join("")}</tbody></table></div></div>
      </div>
      <p class="small muted">Parents would see the same cards as their child, plus a box to reply.</p>`,
    values: `
      <h2>Core values rubric</h2>
      <p class="small muted">From the current PDE template. Each value is rated 1–4. Subject teachers rate the first two; the advisor rates all five.</p>
      <div class="table-wrap"><table>
        <thead><tr><th>Core value</th><th>Rated by</th><th>Beginning (1)</th><th>Developing (2)</th><th>Proficient (3)</th><th>Advanced (4)</th></tr></thead>
        <tbody>${PDE_VALUES.map((x) => `<tr><td><b>${esc(x.v)}</b><div class="small muted">${esc(x.ind)}</div></td><td class="small">${esc(x.by)}</td>${x.r.map((t) => `<td class="small">${esc(t)}</td>`).join("")}</tr>`).join("")}</tbody>
      </table></div>`,
    teacher: `
      <h2>Teacher goals ${ex}</h2>
      <p class="small muted">For teachers who teach every class (ICT, Art, PE …): one goal and action plan per division. Example from ICT, 2025–26.</p>
      ${[["Primary School (Grades 2–5)", "By the end of the year, 85% of Primary students will design, fix and complete their own simple programs and digital projects.",
          ["Model debugging techniques; students build interactive stories, thinking through problems step by step.", "Run real-world data experiments; students organise results into spreadsheets and charts.", "Guide students to curate a digital folder showing their growth in typing, design and organisation."]],
         ["Middle School (Grades 6–8)", "Every student completes a real-world capstone project that combines computing and data management.",
          ["Move students from block coding to text-based languages through digital art and games.", "Facilitate mobile-app prototypes with a focus on user experience.", "Lead seminars on AI and privacy; students research and present arguments."]],
         ["Grades 9–10 (IGCSE ICT 0417)", "100% pass rate, with 70% of students at A*–B, through practical training and exam mastery.",
          ["Bi-weekly diagnostic theory tests so students close their own gaps.", "Monthly timed practical sprints under exam conditions.", "Workshops on complex spreadsheet formulas and CSS; coaching on structured written answers."]]]
        .map(([d, g, plan]) => `<div class="card"><h3>${esc(d)}</h3><p class="small"><b>Goal:</b> ${esc(g)}</p><ul class="small">${plan.map((p) => `<li>${esc(p)}</li>`).join("")}</ul>
          <p class="small muted">End of year: the teacher records whether the goal was met, with evidence.</p></div>`).join("")}`,
    questions: `
      <h2>Questions for the committee</h2>
      <p class="small">Initial discussion points. The app is ready to support whatever the school decides; the <b>educational decisions belong to the committee</b>. Nothing here is school policy yet.</p>
      ${PDE_QUESTIONS.map((x, i) => `
        <div class="card">
          <h3>${i + 1}. ${x.q}</h3>
          ${x.why ? `<p class="small muted" style="margin-top:0">${x.why}</p>` : ""}
          ${x.options ? `<ul class="small">${x.options.map((o) => `<li>${o}</li>`).join("")}</ul>` : ""}
          ${x.body || ""}
          ${x.suggest ? `<p class="small"><span class="badge info">Suggestion</span> ${x.suggest}</p>` : ""}
        </div>`).join("")}
      <div class="card" style="border-color:var(--accent)">
        <h3>A possible first step <span class="badge">Proposal, for the committee to accept or change</span></h3>
        <ul class="small">${PDE_FIRST_STEP.map((x) => `<li>${x}</li>`).join("")}</ul>
        <p class="small muted">The goal-setting already built in the app (target, goal and action plan for the four core subjects) is a starting point and can be changed once the committee decides.</p>
      </div>
      <p class="small">Proposed timing: the committee answers these in <b>Fall 2026</b>; PDE opens for one grade as a pilot, then for everyone in <b>Spring 2027</b>.</p>
      ${isAdmin() ? `<div class="card" style="border-style:dashed"><h3>ICT notes <span class="badge">Only admins see this</span></h3>
        <p class="small"><b>Student sign-in:</b> every bobae has a Microsoft school account, so G6–G12 can sign in to PDE. Before the pilot: run a short sign-in training per class and standardise account names so they're easy to remember (the current pattern is <code>fullname.calledname + graduation year</code>, e.g. <code>eunseongchoi.eden2033@</code>). Check G11–G12 Teams access at the same time.</p></div>` : ""}`,
    goals: `<div id="pdeGoals"></div>`,
  }[sec];
  main.innerHTML = `
    ${sec === "goals"
      ? notice("ok", "<b>Bobae goals are live.</b> Bobaes enter a target, goal and action plan for each core subject; you can read them and add feedback. The rest of PDE (monthly scores, core values) is still being designed with the teacher committee.")
      : notice("warn", "<b>Work in progress.</b> Academic Purpose-Driven Education is being designed with a committee of teachers. These pages are a preview to discuss; nothing here is saved, and all numbers are examples.")}
    <div class="row" style="gap:6px;margin-bottom:14px">${PDE_SECTIONS.map(([k, n]) => `<button class="btn small ${k === sec ? "primary" : ""}" data-pde="${k}">${n}</button>`).join("")}</div>
    ${body}`;
  $$("[data-pde]", main).forEach((b) => (b.onclick = (e) => { e.preventDefault(); S.pdeSec = b.dataset.pde; render(); window.scrollTo(0, 0); }));
  if (sec === "goals") viewPdeGoals($("#pdeGoals"));
}

// ── Help ─────────────────────────────────────────────────────
function viewHelp(main) {
  main.innerHTML = `
    <h1>Help</h1>
    <div class="card"><h3>Entering marks</h3><ol class="steps small">
      <li><b>My classes → + Add a class</b>. Pick the grade and subject. Untick students who don't take it. For split groups (e.g. Chinese levels) type the group name.</li>
      <li>Type a score (0–100), press <b>Enter</b> to move to the next student. Pick behaviour 1–4 and add a short remark if you like.</li>
      <li>Every change saves on its own. "Saved ✓" appears at the end of the row.</li>
      <li>When the class is finished, press <b>Submit to advisor</b>.</li></ol></div>
    <div class="card"><h3>Who can see what</h3><ul class="small">
      <li>You see and edit only the classes, awards and comments <b>you</b> created. SharePoint enforces this, not just this page.</li>
      <li>Your grade advisor can see all subjects for their grade and reopen a submitted class.</li>
      <li>Leadership and the Admin Office can see everything. Every change is kept in version history.</li></ul></div>
    <div class="card"><h3>Behaviour scale</h3><ul class="small">${P().behaviourScale.map((b) => `<li><b>${b.value}</b>: ${esc(b.label)}</li>`).join("")}</ul></div>
    <div class="card"><h3>Problems?</h3><p class="small">If you can't sign in or see "doesn't have access", contact the ICT coordinator with the email shown at the top right.</p></div>`;
}

// ── Start ────────────────────────────────────────────────────
async function start() {
  $("#main").innerHTML = `<div class="center muted">Loading…</div>`;
  try {
    if (!pca) await initAuth();
    if (!account) return render();
    S.me = S.realMe = lc(account.username); S.meName = account.name || S.me; S.viewAs = "";
    S.loadErr = null;
    try { await loadAll(); }
    catch (e) {
      // No access to the marks site: a bobae signing in for their goals lands on the PDE site instead.
      if ([401, 403, 404].includes(e.status) && typeof tryStudentMode === "function" && await tryStudentMode()) return render();
      throw e;
    }
    if (typeof loadPde === "function") await loadPde();   // never fails; sets S.pde.err instead
    const resume = sessionStorage.getItem("rp_tab"); sessionStorage.removeItem("rp_tab");
    if (resume && tabsFor().some(([k]) => k === resume)) S.tab = resume;
    else if (isAdmin() && !Object.keys(S.lists).includes(L.students)) S.tab = "setup";
  } catch (e) {
    S.loadErr = e;
  }
  render();
}
if (C.clientId.startsWith("PASTE")) {
  document.getElementById("main").innerHTML = `<div class="center"><h1>Almost ready</h1><p class="lede">Paste the Entra application (client) ID into <code>config.js</code>.</p></div>`;
} else start();
