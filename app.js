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
  advGrade: "", awGrade: "", awSel: new Set(), awSearch: "", awAll: false,
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
const marksListName = (g) => L.marksPrefix + g;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
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
async function sp(method, path, body, headers = {}) {
  const r = await fetch(`${S.siteUrl}/_api/${path}`, {
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
  const [students, subjects, catalogue] = await Promise.all([readList(L.students), readList(L.subjects), readList(L.catalogue)]);
  S.students = students.sort((a, b) => (a.f.FullName || "").localeCompare(b.f.FullName || ""));
  S.subjects = subjects.sort((a, b) => (a.f.SortOrder ?? 999) - (b.f.SortOrder ?? 999) || (a.f.Title || "").localeCompare(b.f.Title || ""));
  S.catalogue = catalogue.filter((x) => x.f.Active !== false).sort((a, b) => (a.f.SortOrder ?? 999) - (b.f.SortOrder ?? 999));
  await Promise.all([loadMarks(), loadComments(), loadAwards()]);
}
async function loadMarks(grade) {
  const grades = grade ? [grade] : P().grades;
  await Promise.all(grades.map(async (g) => { S.marks[g] = (await readList(marksListName(g))).filter((x) => inTerm(x.f)); }));
}
async function loadComments() { S.comments = (await readList(L.comments)).filter((x) => inTerm(x.f)); }
async function loadAwards() { S.awards = (await readList(L.awards)).filter((x) => inTerm(x.f)); }
const rosterFor = (g) => S.students.filter((s) => s.f.Grade === g && s.f.Status !== "Withdrawn");
const studentById = (id) => S.students.find((s) => s.f.Title === id);

// ── Roles (layout only — real access comes from SharePoint) ──
const isAdmin = () => P().admins.map(lc).includes(S.me);
const isLeader = () => isAdmin() || P().leadership.map(lc).includes(S.me);
const advisorGrades = () => Object.entries(P().advisors || {}).filter(([, a]) => a && lc(a.email) === S.me).map(([g]) => g);
const canAdvise = (g) => isAdmin() || advisorGrades().includes(g);
const advisorName = (g) => (P().advisors[g] && P().advisors[g].name) || "";

// ── Shell ────────────────────────────────────────────────────
function tabsFor() {
  const t = [["classes", "My classes"], ["awards", "Awards"]];
  if (advisorGrades().length || isLeader()) t.push(["advisor", "Advisor view"]);
  if (isAdmin()) t.push(["export", "Export"], ["setup", "Setup"]);
  t.push(["pde", "PDE"], ["help", "Help"]);
  return t;
}
function renderShell() {
  $("#schoolName").textContent = C.schoolName;
  $("#appName").textContent = C.appName;
  $("#footer").textContent = `Mr Glen Joshua (${P().semester} ${P().academicYear.slice(0, 4)}) · ${C.version} · Uses Zhongshi Microsoft SharePoint as database`;
  if (!account) { $("#who").textContent = ""; $("#authBtn").innerHTML = ""; $("#tabs").innerHTML = ""; return; }
  $("#who").textContent = account.name || S.me;
  $("#authBtn").innerHTML = `<button class="btn small" id="signOut">Sign out</button>`;
  $("#signOut").onclick = () => pca.logoutRedirect({ account });
  if (S.loadErr) { $("#tabs").innerHTML = ""; return; }
  const active = { newclass: "classes", entry: S.back }[S.tab] || S.tab;
  $("#tabs").innerHTML = tabsFor().map(([k, n]) => `<button data-tab="${k}" ${k === active ? 'aria-current="page"' : ""}>${n}</button>`).join("");
  $$("#tabs button").forEach((b) => (b.onclick = () => go(b.dataset.tab)));
}
function go(tab) { S.tab = tab; render(); window.scrollTo(0, 0); }
function render() {
  renderShell();
  const main = $("#main");
  if (!account) return signInScreen(main);
  if (S.loadErr) return errorScreen(main);
  const views = { classes: viewClasses, newclass: viewNewClass, entry: viewEntry, advisor: viewAdvisor, awards: viewAwards, export: viewExport, setup: viewSetup, pde: viewPde, help: viewHelp };
  (views[S.tab] || viewClasses)(main);
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
function classKey(f) { return `${f.Subject}|${f.ClassLevel || ""}`; }
function groupClasses(rows) {
  const m = new Map();
  for (const r of rows) {
    const k = classKey(r.f) + "|" + r.by;
    if (!m.has(k)) m.set(k, { subject: r.f.Subject, level: r.f.ClassLevel || "", by: r.by, byName: r.byName || r.f.TeacherName, rows: [] });
    m.get(k).rows.push(r);
  }
  return [...m.values()].map((c) => ({ ...c,
    total: c.rows.length,
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
  const mine = [];
  for (const g of P().grades) for (const c of groupClasses((S.marks[g] || []).filter((r) => r.by === S.me))) mine.push({ grade: g, ...c });
  const hasLists = P().grades.some((g) => S.lists[marksListName(g)]);
  main.innerHTML = `
    <h1>My classes</h1>
    <p class="lede">${esc(P().semester)} semester ${esc(P().academicYear)}. Each class is one grade and one subject you teach.</p>
    <div id="flash"></div>
    ${!hasLists ? notice("warn", isAdmin() ? `The SharePoint lists haven't been created yet. Go to <a href="#" id="toSetup">Setup</a>.` : "The reporting lists aren't set up yet. Please check with the ICT coordinator.") : ""}
    <div class="class-grid">
      ${mine.map((c, i) => `
        <button class="class-card" data-i="${i}">
          <div class="t">${esc(c.grade)} · ${esc(subjectDisplay({ Subject: c.subject, ClassLevel: c.level }))}</div>
          <div class="bar"><i style="width:${c.total ? Math.round((100 * c.done) / c.total) : 0}%"></i></div>
          <div class="small muted">${c.total} students &nbsp; ${statusBadge(c)}</div>
        </button>`).join("")}
      ${hasLists ? `<button class="class-card" id="addClass" style="border-style:dashed;text-align:center"><div class="t">+ Add a class</div><div class="small muted">Pick a grade and subject</div></button>` : ""}
    </div>
    ${mine.length === 0 && hasLists ? `<p class="muted small" style="margin-top:14px">You haven't added any classes this semester yet.</p>` : ""}`;
  $$(".class-card[data-i]").forEach((b) => (b.onclick = () => { const c = mine[+b.dataset.i]; openClass(c.grade, c.subject, c.level, c.by, "classes"); }));
  const add = $("#addClass"); if (add) add.onclick = () => go("newclass");
  const ts = $("#toSetup"); if (ts) ts.onclick = (e) => { e.preventDefault(); go("setup"); };
}
function openClass(grade, subject, level, by, back) { S.cls = { grade, subject, level, by }; S.back = back; go("entry"); }

function viewNewClass(main) {
  const grades = P().grades.filter((g) => S.lists[marksListName(g)]);
  main.innerHTML = `
    <h1>Add a class</h1>
    <p class="lede">Choose the grade and subject. Then untick any students who don't take it.</p>
    <div id="flash"></div>
    <div class="card"><div class="grid">
      <label><span>Grade</span><select id="ncGrade"><option value="">Choose…</option>${grades.map((g) => `<option>${esc(g)}</option>`).join("")}</select></label>
      <label><span>Subject</span><select id="ncSubject" disabled><option value="">Choose a grade first</option></select></label>
      <label class="wide"><span>Class level or set (optional). Use only when the grade is split into groups, e.g. <i>Chinese Intermediate A</i>. Leave blank if you teach the whole grade.</span><input id="ncLevel" maxlength="60"></label>
    </div></div>
    <div id="ncRoster"></div>
    <div class="sticky-actions">
      <button class="btn primary" id="ncCreate" disabled>Create class</button>
      <button class="btn" id="ncCancel">Cancel</button>
      <span class="small muted" id="ncCount"></span>
    </div>`;
  const gSel = $("#ncGrade"), sSel = $("#ncSubject");
  const refresh = () => {
    const g = gSel.value;
    const roster = g ? rosterFor(g) : [];
    $("#ncRoster").innerHTML = g ? `
      <div class="card">
        <div class="row" style="justify-content:space-between;align-items:center;margin-bottom:8px">
          <h3>Students in ${esc(g)} (${roster.length})</h3>
          <span><button class="btn small" id="ncAll">Select all</button> <button class="btn small" id="ncNone">Select none</button></span>
        </div>
        ${roster.length ? `<div class="grid">${roster.map((s) => `
          <label style="color:var(--text)"><input type="checkbox" class="ncStu" value="${esc(s.f.Title)}" checked> ${esc(s.f.CalledName || s.f.FullName)} <span class="muted small">${esc(s.f.FullName)}</span></label>`).join("")}</div>`
          : notice("warn", "No students found for this grade in the roster.")}
      </div>` : "";
    const upd = () => { const n = $$(".ncStu:checked").length; $("#ncCount").textContent = g ? `${n} students selected` : ""; $("#ncCreate").disabled = !(g && sSel.value && n); };
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
    refresh();
  };
  sSel.onchange = refresh;
  $("#ncCancel").onclick = () => go("classes");
  $("#ncCreate").onclick = async () => {
    const g = gSel.value, subject = sSel.value, level = $("#ncLevel").value.trim();
    const ids = $$(".ncStu:checked").map((c) => c.value);
    const existing = (S.marks[g] || []).filter((r) => r.by === S.me && r.f.Subject === subject && (r.f.ClassLevel || "") === level);
    if (existing.length) { flash("warn", "You already have this class. Opening it."); return openClass(g, subject, level, S.me, "classes"); }
    const subj = S.subjects.find((s) => s.f.Grade === g && s.f.Title === subject);
    $("#ncCreate").disabled = true; $("#ncCreate").textContent = "Creating…";
    try {
      await createRows(g, subject, subj ? subj.f.SubjectType : "I", level, ids);
      openClass(g, subject, level, S.me, "classes");
    } catch (e) { flash("error", e.message); $("#ncCreate").disabled = false; $("#ncCreate").textContent = "Create class"; }
  };
}
async function createRows(g, subject, type, level, ids) {
  const url = itemsUrl(marksListName(g));
  const reqs = ids.map((id) => {
    const s = studentById(id).f;
    return { method: "POST", url, body: { fields: {
      Title: `${s.CalledName || s.FullName} – ${subject}`, StudentID: id, FullName: s.FullName, CalledName: s.CalledName || "",
      Subject: subject, SubjectType: type || "I", ClassLevel: level, ...term(), Submitted: false, TeacherName: S.meName } } };
  });
  const res = await batch(reqs);
  const bad = batchFailures(res);
  await loadMarks(g);
  if (bad.length) throw new Error(`${bad.length} of ${reqs.length} students could not be added (${bad[0]?.body?.error?.message || "unknown error"}). Open the class and use "Add students" to retry.`);
}

// ── Marks entry ──────────────────────────────────────────────
function viewEntry(main) {
  const { grade, subject, level, by } = S.cls;
  const rows = (S.marks[grade] || []).filter((r) => r.by === by && r.f.Subject === subject && (r.f.ClassLevel || "") === level)
    .sort((a, b) => (a.f.FullName || "").localeCompare(b.f.FullName || ""));
  const mine = by === S.me;
  const submitted = rows.length > 0 && rows.every((r) => r.f.Submitted === true);
  const supervisor = canAdvise(grade) && !mine;
  const editable = (mine && !submitted) || (supervisor && canAdvise(grade)) || (isAdmin());
  const teacher = rows[0] ? (rows[0].byName || rows[0].f.TeacherName) : "";
  const scale = P().behaviourScale;
  const missingIds = rosterFor(grade).filter((s) => !rows.some((r) => r.f.StudentID === s.f.Title));
  main.innerHTML = `
    <p><a href="#" id="back">← Back</a></p>
    <h1>${esc(grade)} · ${esc(subjectDisplay({ Subject: subject, ClassLevel: level }))}</h1>
    <p class="lede">${mine ? "Your class" : "Teacher: " + esc(teacher)} · ${rows.length} students · ${submitted ? `<span class="badge ok">Submitted</span>` : `<span class="badge warn">Not submitted</span>`}</p>
    <div id="flash"></div>
    ${submitted && mine && !isAdmin() ? notice("ok", "Submitted to the advisor. To change something, ask your grade advisor to reopen it.") : ""}
    ${supervisor ? notice("warn", "You are editing another teacher's class as advisor. Changes are logged in SharePoint version history.") : ""}
    <div class="legend"><b>Behaviour:</b> ${scale.map((b) => `<span>${b.value} = ${esc(b.label)}</span>`).join("")}<span>· Score 0–100</span></div>
    <div class="table-wrap"><table class="entry">
      <thead><tr><th>#</th><th>Student</th><th class="num">Score</th><th>Behaviour</th><th>Remark (optional, max ${P().remarkMaxLength})</th><th></th>${editable && (mine || isAdmin()) ? "<th></th>" : ""}</tr></thead>
      <tbody>
      ${rows.map((r, i) => `
        <tr data-id="${r.id}">
          <td class="muted">${i + 1}</td>
          <td><b>${esc(r.f.CalledName || r.f.FullName)}</b><div class="small muted">${esc(r.f.FullName)} · ${esc(r.f.StudentID)}</div></td>
          <td class="num"><input class="score" inputmode="decimal" autocomplete="off" value="${blank(r.f.Score) ? "" : esc(r.f.Score)}" ${editable ? "" : "disabled"} aria-label="Score for ${esc(r.f.FullName)}"></td>
          <td><select class="beh" ${editable ? "" : "disabled"} aria-label="Behaviour for ${esc(r.f.FullName)}"><option value=""></option>${scale.map((b) => `<option value="${b.value}" ${Number(r.f.Behaviour) === b.value ? "selected" : ""}>${b.value} – ${esc(b.label)}</option>`).join("")}</select></td>
          <td><textarea class="remark" maxlength="${P().remarkMaxLength}" ${editable ? "" : "disabled"} aria-label="Remark for ${esc(r.f.FullName)}">${esc(r.f.Remark || "")}</textarea></td>
          <td class="rowstate"></td>
          ${editable && (mine || isAdmin()) ? `<td><button class="btn small danger rm" title="Remove this student from the class">✕</button></td>` : ""}
        </tr>`).join("")}
      </tbody></table></div>
    ${editable && missingIds.length && (mine || isAdmin()) ? `
      <details class="card"><summary>Add students (${missingIds.length} in ${esc(grade)} not in this class)</summary>
        <div class="grid" style="margin-top:10px">${missingIds.map((s) => `<label style="color:var(--text)"><input type="checkbox" class="addStu" value="${esc(s.f.Title)}"> ${esc(s.f.CalledName || s.f.FullName)} <span class="muted small">${esc(s.f.FullName)}</span></label>`).join("")}</div>
        <p><button class="btn" id="addStudents">Add selected</button></p>
      </details>` : ""}
    <div class="sticky-actions">
      ${mine && !submitted ? `<button class="btn primary" id="submit">Submit to advisor</button>` : ""}
      ${submitted && (canAdvise(grade)) ? `<button class="btn" id="reopen">Reopen for the teacher</button>` : ""}
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
        await graph("PATCH", `${itemsUrl(marksListName(grade))}/${r.id}/fields`, fields);
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
    const rm = $(".rm", tr);
    if (rm) rm.onclick = async () => {
      if (!confirm(`Remove ${r.f.FullName} from this class? Their score and remark for this subject will be deleted.`)) return;
      try { await graph("DELETE", `${itemsUrl(marksListName(grade))}/${r.id}`); await loadMarks(grade); render(); }
      catch (e) { flash("error", e.message); }
    };
  }
  const sub = $("#submit");
  if (sub) sub.onclick = async () => {
    const missing = rows.filter((r) => blank(r.f.Score) || blank(r.f.Behaviour));
    if (missing.length && !confirm(`${missing.length} student(s) still have no score or behaviour:\n${missing.slice(0, 8).map((r) => "• " + r.f.FullName).join("\n")}${missing.length > 8 ? "\n…" : ""}\n\nSubmit anyway?`)) return;
    if (!missing.length && !confirm("Submit this class to the advisor? You won't be able to edit it afterwards unless the advisor reopens it.")) return;
    await setSubmitted(grade, rows, true);
  };
  const ro = $("#reopen");
  if (ro) ro.onclick = () => setSubmitted(grade, rows, false);
  const add = $("#addStudents");
  if (add) add.onclick = async () => {
    const ids = $$(".addStu:checked").map((c) => c.value);
    if (!ids.length) return;
    const t = rows[0] ? rows[0].f.SubjectType : (S.subjects.find((s) => s.f.Grade === grade && s.f.Title === subject)?.f.SubjectType || "I");
    try { await createRows(grade, subject, t, level, ids); render(); } catch (e) { flash("error", e.message); }
  };
}
async function setSubmitted(grade, rows, value) {
  const url = itemsUrl(marksListName(grade));
  const res = await batch(rows.map((r) => ({ method: "PATCH", url: `${url}/${r.id}/fields`, body: { Submitted: value } })));
  const bad = batchFailures(res);
  await loadMarks(grade);
  render();
  if (bad.length) flash("error", `${bad.length} rows could not be updated: ${bad[0]?.body?.error?.message || ""}`);
  else flash("ok", value ? "Submitted to the advisor." : "Reopened. The teacher can edit again.");
}

// ── Advisor view ─────────────────────────────────────────────
function viewAdvisor(main) {
  const grades = isLeader() ? P().grades : advisorGrades();
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
      <div class="muted small" style="padding-bottom:8px">Advisor: ${esc(advisorName(g) || "not set")}</div>
    </div>
    <div id="flash"></div>
    ${onlyMine ? notice("warn", "You can only see your own classes in this grade. Ask ICT to give you advisor access in Setup → Permissions.") : ""}
    <div class="stats">
      <div class="stat"><div class="k">Subjects submitted</div><div class="v">${classes.filter((c) => c.submitted).length} / ${Math.max(expected.length, classes.length)}</div></div>
      <div class="stat"><div class="k">Not started</div><div class="v">${notStarted.length}</div></div>
      <div class="stat"><div class="k">Students</div><div class="v">${roster.length}</div></div>
      <div class="stat"><div class="k">Advisor comments written</div><div class="v">${roster.filter((s) => (commentFor(s.f.Title)?.f.Comment || "").trim()).length} / ${roster.length}</div></div>
    </div>
    <h2>Subjects</h2>
    <div class="table-wrap"><table>
      <thead><tr><th>Subject</th><th>Teacher</th><th>Progress</th><th></th></tr></thead>
      <tbody>
        ${classes.map((c, i) => `<tr><td>${esc(subjectDisplay({ Subject: c.subject, ClassLevel: c.level }))}</td><td>${esc(c.byName)}</td><td>${statusBadge(c)}</td><td><button class="btn small" data-open="${i}">Open</button></td></tr>`).join("")}
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
        return `<td class="cell ${miss ? "missing" : ""}" title="${esc(r.byName)}${r.f.Remark ? " — " + esc(r.f.Remark) : ""}">${blank(r.f.Score) ? "–" : esc(r.f.Score)} / ${blank(r.f.Behaviour) ? "–" : esc(r.f.Behaviour)}</td>`;
      }).join("")}</tr>`).join("")}</tbody></table></div>
    <h2>Advisor comments</h2>
    <p class="small muted">${canComment ? "One comment per student. It prints in the teacher's comments box on page 1. Saves when you leave the box." : "Read only."}</p>
    <div class="table-wrap"><table>
      <thead><tr><th style="width:220px">Student</th><th>Comment</th><th></th></tr></thead>
      <tbody>${roster.map((s) => { const c = commentFor(s.f.Title); return `
        <tr data-sid="${esc(s.f.Title)}"><td><b>${esc(s.f.CalledName || s.f.FullName)}</b><div class="small muted">${esc(s.f.FullName)}</div></td>
        <td><textarea class="cmt" rows="3" maxlength="1200" ${canComment ? "" : "disabled"}>${esc(c?.f.Comment || "")}</textarea></td>
        <td class="rowstate"></td></tr>`; }).join("")}</tbody></table></div>`;
  $("#advGrade").onchange = (e) => { S.advGrade = e.target.value; render(); };
  $$("[data-open]").forEach((b) => (b.onclick = () => { const c = classes[+b.dataset.open]; openClass(g, c.subject, c.level, c.by, "advisor"); }));
  for (const tr of $$("tr[data-sid]", main)) {
    const ta = $(".cmt", tr), state = $(".rowstate", tr), sid = tr.dataset.sid;
    ta.onchange = async () => {
      state.className = "rowstate"; state.textContent = "Saving…";
      try {
        const c = commentFor(sid), text = ta.value.trim();
        if (c) { await graph("PATCH", `${itemsUrl(L.comments)}/${c.id}/fields`, { Comment: text }); c.f.Comment = text; }
        else {
          const st = studentById(sid).f;
          const it = await graph("POST", itemsUrl(L.comments), { fields: { Title: `${st.FullName} (${g})`, StudentID: sid, Grade: g, ...term(), Comment: text } });
          S.comments.push(norm(it));
        }
        state.textContent = "Saved ✓";
      } catch (e) { state.className = "rowstate err"; state.textContent = "Not saved"; flash("error", e.message); }
    };
  }
}

// ── Awards ───────────────────────────────────────────────────
function viewAwards(main) {
  const grades = P().grades;
  if (!S.awGrade) S.awGrade = advisorGrades()[0] || grades[0];
  const g = S.awGrade;
  const q = lc(S.awSearch);
  const roster = rosterFor(g).filter((s) => !q || lc(`${s.f.FullName} ${s.f.CalledName} ${s.f.Title}`).includes(q));
  const seeAll = isLeader();
  const visible = S.awards.filter((a) => (S.awAll ? true : a.f.Grade === g));
  const countFor = (id) => S.awards.filter((a) => a.f.StudentID === id).length;
  const cats = [...new Set(S.catalogue.map((c) => c.f.Category).filter(Boolean))];
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
              <div class="called">${esc(s.f.CalledName || s.f.FullName)}</div>
              <div class="full">${esc(s.f.FullName)}</div>
              <div class="id">${esc(s.f.Title)}</div>
              ${on ? `<span class="tick">✓</span>` : ""}
            </button>`; }).join("")}
          ${!roster.length ? `<p class="muted">No students match.</p>` : ""}
        </div>
      </div>
      <div class="card award-panel">
        <h3>Give an award</h3>
        <p class="small muted" id="awSelText"></p>
        <label><span>Category</span><select id="awCat"><option value="">Choose…</option>${cats.map((c) => `<option>${esc(c)}</option>`).join("")}</select></label>
        <label style="margin-top:8px"><span>Award</span><select id="awName" disabled><option value="">Choose a category first</option></select></label>
        <label style="margin-top:8px"><span>Certificate type</span><select id="awType">${P().certificateTypes.map((t) => `<option>${esc(t)}</option>`).join("")}</select></label>
        <label style="margin-top:8px"><span>Service points (optional). A number like 16, or Gold / Silver / Bronze. Leave blank for a certificate.</span><input id="awPoints" maxlength="20"></label>
        <label style="margin-top:8px"><span>Date</span><input id="awDate" type="date" value="${today}"></label>
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
  const updSel = () => {
    const n = S.awSel.size;
    const names = [...S.awSel].map((id) => studentById(id)?.f.CalledName || id);
    $("#awSelText").textContent = n ? `${n} selected: ${names.slice(0, 6).join(", ")}${n > 6 ? "…" : ""}` : "No students selected yet.";
    $("#awSave").disabled = !(n && $("#awName").value);
  };
  $("#awGrade").onchange = (e) => { S.awGrade = e.target.value; S.awSel.clear(); render(); };
  const search = $("#awSearch");
  search.oninput = () => { S.awSearch = search.value; const pos = search.selectionStart; render(); const s2 = $("#awSearch"); s2.focus(); s2.setSelectionRange(pos, pos); };
  $("#awAllBtn").onclick = () => { roster.forEach((s) => S.awSel.add(s.f.Title)); render(); };
  $("#awClear").onclick = () => { S.awSel.clear(); render(); };
  $("#awAllGrades").onchange = (e) => { S.awAll = e.target.checked; render(); };
  $$(".stu").forEach((b) => (b.onclick = () => { const id = b.dataset.sid; S.awSel.has(id) ? S.awSel.delete(id) : S.awSel.add(id); render(); }));
  const cat = $("#awCat"), nm = $("#awName"), ty = $("#awType");
  cat.onchange = () => {
    const list = S.catalogue.filter((c) => c.f.Category === cat.value);
    nm.disabled = !cat.value;
    nm.innerHTML = `<option value="">Choose…</option>` + list.map((c) => `<option value="${esc(c.id)}">${esc(c.f.Title)}</option>`).join("");
    updSel();
  };
  nm.onchange = () => {
    const c = S.catalogue.find((x) => x.id === nm.value);
    if (c && P().certificateTypes.includes(c.f.CertificateType)) ty.value = c.f.CertificateType;
    updSel();
  };
  $("#awSave").onclick = async () => {
    const c = S.catalogue.find((x) => x.id === nm.value); if (!c) return;
    const ids = [...S.awSel];
    const dupes = ids.filter((id) => S.awards.some((a) => a.f.StudentID === id && a.f.AwardName === c.f.Title));
    if (dupes.length && !confirm(`${dupes.length} of these students already have "${c.f.Title}" this semester. Add it again?`)) return;
    $("#awSave").disabled = true; $("#awSave").textContent = "Saving…";
    const url = itemsUrl(L.awards), points = $("#awPoints").value.trim(), date = $("#awDate").value;
    const res = await batch(ids.map((id) => { const s = studentById(id).f; return { method: "POST", url, body: { fields: {
      Title: id, StudentID: id, FullName: s.FullName, CalledName: s.CalledName || "", Grade: s.Grade,
      Category: c.f.Category, AwardName: c.f.Title, CertificateType: ty.value, Points: points,
      ...term(), AwardDate: date ? date + "T00:00:00Z" : null, AwardedBy: S.meName, Printed: false } } }; }));
    const bad = batchFailures(res);
    await loadAwards();
    S.awSel.clear(); render();
    if (bad.length) flash("error", `${bad.length} award(s) could not be saved: ${bad[0]?.body?.error?.message || ""}`);
    else flash("ok", `Saved "${c.f.Title}" for ${ids.length} student${ids.length > 1 ? "s" : ""}.`);
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
    const base = { Student_ID: id, Student_Name: st.f.FullName, Called_Name: st.f.CalledName || "", Grade: g, Advisor: advisorName(g), Grading_Period: `${P().semester} Semester` };
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
function viewExport(main) {
  const tag = `${P().semester}${P().academicYear.slice(0, 4)}`;
  main.innerHTML = `
    <h1>Export for the report card generator</h1>
    <p class="lede">Download each grade's marks file and the awards file, then upload them with the attendance sheet at
      <a href="https://wzfs-report-card.vercel.app/" target="_blank" rel="noopener">wzfs-report-card.vercel.app</a>.</p>
    <div id="flash"></div>
    <p><button class="btn" id="reload">↻ Reload latest data</button></p>
    <div class="table-wrap"><table>
      <thead><tr><th>Grade</th><th class="num">Students</th><th class="num">Classes</th><th>Not submitted</th><th class="num">Blank marks</th><th class="num">Comments</th><th>Warnings</th><th></th></tr></thead>
      <tbody>${P().grades.map((g) => { const c = gradeChecks(g); return `
        <tr><td><b>${esc(g)}</b></td><td class="num">${rosterFor(g).length}</td><td class="num">${c.classes}</td>
        <td class="small">${c.open.length ? c.open.map((x) => esc(x.subject)).join(", ") : `<span class="badge ok">All in</span>`}</td>
        <td class="num">${c.missing || ""}</td><td class="num">${c.comments} / ${rosterFor(g).length}</td>
        <td class="small">${c.dupes.length ? `<span class="badge bad">${c.dupes.length} duplicate subject rows</span> ` : ""}${c.notInRoster.length ? `<span class="badge bad">${c.notInRoster.length} rows not in roster</span>` : ""}</td>
        <td><button class="btn small" data-g="${esc(g)}" ${c.rows || c.comments ? "" : "disabled"}>Marks file</button></td></tr>`; }).join("")}
      </tbody></table></div>
    <div class="row">
      <button class="btn primary" id="awardsFile">Awards file (all grades)</button>
      <button class="btn" id="allMarks">All grades in one marks file</button>
    </div>
    <p class="small muted" style="margin-top:12px">Duplicate subject rows mean two teachers created the same subject for the same student. Open the Advisor view to see who, and remove one before printing.</p>`;
  $("#reload").onclick = async () => { $("#reload").textContent = "Loading…"; await loadAll(); render(); flash("ok", "Up to date."); };
  $$("[data-g]").forEach((b) => (b.onclick = () => download(sheetBook(marksExportRows(b.dataset.g), "Marks"), `Marks_${b.dataset.g}_${tag}.xlsx`)));
  $("#allMarks").onclick = () => download(sheetBook(P().grades.flatMap(marksExportRows), "Marks"), `Marks_AllGrades_${tag}.xlsx`);
  $("#awardsFile").onclick = () => {
    const rows = S.awards.slice().sort((a, b) => (a.f.Grade || "").localeCompare(b.f.Grade || "") || (a.f.FullName || "").localeCompare(b.f.FullName || ""));
    // Column names chosen to match wzfs-report-card's awards parser (Award before any other "award" column).
    const forGenerator = rows.map((a) => { const s = studentById(a.f.StudentID)?.f || a.f; return {
      Student_ID: a.f.StudentID, Student_Name: s.FullName, Grade: a.f.Grade, Award: a.f.AwardName,
      Recognition: a.f.CertificateType, Points: a.f.Points || "", Date: (a.f.AwardDate || "").slice(0, 10) }; });
    const printList = rows.filter((a) => !a.f.Points).map((a) => ({
      Grade: a.f.Grade, Student_ID: a.f.StudentID, Certificate_Name: studentById(a.f.StudentID)?.f.CertificateName || a.f.FullName,
      Award: a.f.AwardName, Certificate_Type: a.f.CertificateType, Given_By: a.byName || a.f.AwardedBy, Printed: a.f.Printed ? "Yes" : "No" }));
    const wb = sheetBook(forGenerator, "Awards");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(printList), "Certificate print list");
    download(wb, `Awards_${tag}.xlsx`);
  };
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
    { name: L.students, kind: "read", cols: [T("FullName"), T("CalledName"), T("CertificateName"), T("Grade"), T("Email"), T("Department"), T("Joined"), CH("Status", ["Active", "Withdrawn"])] },
    { name: L.subjects, kind: "read", cols: [T("Grade"), CH("SubjectType", ["I", "II"]), T("Teacher"), B("PrintsOnCard"), N("SortOrder")] },
    { name: L.catalogue, kind: "read", cols: [T("Category"), T("CertificateType"), B("Active"), N("SortOrder")] },
    { name: L.settings, kind: "read", cols: [T("Json", true)] },
    { name: L.awards, kind: "entry", cols: [T("StudentID"), T("FullName"), T("CalledName"), T("Grade"), T("Category"), T("AwardName"), T("CertificateType"), T("Points"), T("Semester"), T("AcademicYear"), D("AwardDate"), T("AwardedBy"), B("Printed")] },
    { name: L.comments, kind: "entry", cols: [T("StudentID"), T("Grade"), T("Semester"), T("AcademicYear"), T("Comment", true)] },
  ];
  for (const g of P().grades) specs.push({ name: marksListName(g), grade: g, kind: "marks", cols: [
    T("StudentID"), T("FullName"), T("CalledName"), T("Subject"), CH("SubjectType", ["I", "II"]), T("ClassLevel"), T("Semester"), T("AcademicYear"),
    N("Score", 0, 100, "one"), N("Behaviour", 1, 4, "none"), T("Remark", true), B("Submitted"), T("TeacherName")] });
  return specs;
}
function readSheet(file) {
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
      ${missing.length ? `<p class="small muted">Missing: ${missing.map((m) => esc(m.name)).join(", ")}</p><button class="btn primary" id="createLists">Create the missing lists</button>` : `<span class="badge ok">All lists exist</span>`}
      <pre id="createLog" class="small muted" style="white-space:pre-wrap"></pre>
    </div>

    <div class="card"><h3>2 · Import data</h3>
      <p class="small muted">Use the prepared Fall 2026 workbooks. Existing rows are updated by key, new rows added; nothing is deleted.</p>
      <div class="grid">
        <label><span>Roster: <code>List_Students_Fall2026.xlsx</code> (${S.students.length} now)</span><input type="file" id="fStudents" accept=".xlsx,.xls,.csv"></label>
        <label><span>Subjects: <code>WZFS_Fall2026_UpperSchool_Subjects.xlsx</code> (${S.subjects.length} now)</span><input type="file" id="fSubjects" accept=".xlsx,.xls,.csv"></label>
        <label><span>Award catalogue: <code>WZFS_Fall2026_Awards_Catalogue.xlsx</code> (${S.catalogue.length} now)</span><input type="file" id="fCatalogue" accept=".xlsx,.xls,.csv"></label>
      </div>
      <pre id="importLog" class="small muted" style="white-space:pre-wrap"></pre>
    </div>

    <div class="card"><h3>3 · Term and people</h3>
      <div class="grid">
        <label><span>Academic year</span><input id="sYear" value="${esc(p.academicYear)}"></label>
        <label><span>Semester</span><select id="sSem">${p.semesters.map((s) => `<option ${s === p.semester ? "selected" : ""}>${esc(s)}</option>`).join("")}</select></label>
        <label class="wide"><span>Admins (full access in this app), one email per line</span><textarea id="sAdmins" rows="2">${esc(p.admins.join("\n"))}</textarea></label>
        <label class="wide"><span>Leadership (read-only across grades), one email per line. Also add them as Owners of the SharePoint site.</span><textarea id="sLead" rows="3">${esc(p.leadership.join("\n"))}</textarea></label>
      </div>
      <h3 style="margin-top:14px">Class advisors</h3>
      <div class="table-wrap"><table><thead><tr><th>Grade</th><th>Name (prints on the card)</th><th>School email</th></tr></thead><tbody>
        ${p.grades.map((g) => `<tr><td>${esc(g)}</td><td><input class="advName" data-g="${esc(g)}" value="${esc(p.advisors[g]?.name || "")}"></td><td><input class="advEmail" data-g="${esc(g)}" value="${esc(p.advisors[g]?.email || "")}" placeholder="name@zhongshischool.org"></td></tr>`).join("")}
      </tbody></table></div>
      <button class="btn primary" id="saveSettings">Save term and people</button>
    </div>

    <div class="card"><h3>4 · Permissions</h3>
      <p class="small">Puts everyone who enters data into the SharePoint group <b>${esc(C.staffGroup)}</b>, then for every list: stops inheriting, removes site Members/Visitors, keeps Owners on Full Control, turns on version history and:</p>
      <ul class="small">
        <li><b>Students, Subjects, Award Catalogue, Settings</b>: staff group can <b>Read</b>.</li>
        <li><b>Awards, Advisor Comments, Marks G…</b>: staff group can <b>Contribute</b>, with <b>item-level privacy</b> (each person reads and edits only what they created).</li>
        <li><b>Marks G…</b>: that grade's advisor also gets <b>Design</b>, so they can see and correct the whole grade.</li>
      </ul>
      <label><span>Staff who enter marks or awards, one school email per line (advisors are added automatically)</span><textarea id="staffEmails" rows="6" placeholder="name@zhongshischool.org"></textarea></label>
      <p><button class="btn primary" id="applyPerms">Apply permissions</button></p>
      <p class="small muted">Needs the SharePoint <code>AllSites.FullControl</code> delegated permission (admin consent) on the app registration. If it fails, do it by hand with the links below.</p>
      <pre id="permLog" class="small" style="white-space:pre-wrap"></pre>
      <details><summary class="small">Manual links</summary>
        <div class="table-wrap" style="margin-top:8px"><table><thead><tr><th>List</th><th>Should be</th><th></th></tr></thead><tbody>
        ${specs.filter((s) => S.lists[s.name]).map((s) => { const id = S.lists[s.name].id; return `<tr><td>${esc(s.name)}</td>
          <td class="small">${s.kind === "read" ? "Staff: Read" : `Staff: Contribute + item-level "Read/edit items created by the user"${s.kind === "marks" ? `; advisor (${esc(p.advisors[s.grade]?.email || "?")}): Design` : ""}`}</td>
          <td class="small"><a target="_blank" rel="noopener" href="${esc(S.siteUrl)}/_layouts/15/user.aspx?obj=%7B${id}%7D,list&List=%7B${id}%7D">Permissions</a>
          ${s.kind !== "read" ? ` · <a target="_blank" rel="noopener" href="${esc(S.siteUrl)}/_layouts/15/advsetng.aspx?List=%7B${id}%7D">Advanced settings</a>` : ""}</td></tr>`; }).join("")}
        </tbody></table></div>
      </details>
    </div>

    <div class="card"><h3>5 · Deployment check</h3>
      <p class="small">Redirect URI for this address (add it in Entra → App registration → Authentication → Single-page application): <code>${esc(redirectUri)}</code></p>
      <p class="small">Signed in as <code>${esc(S.me)}</code> · site <code>${esc(S.siteId)}</code></p>
    </div>`;

  const cl = $("#createLists");
  if (cl) cl.onclick = async () => {
    cl.disabled = true; const log = $("#createLog");
    for (const s of missing) {
      log.textContent += `Creating ${s.name}… `;
      try {
        await graph("POST", `/sites/${S.siteId}/lists`, { displayName: s.name, list: { template: "genericList" }, columns: s.cols }, SETUP_SCOPES);
        log.textContent += "done\n";
      } catch (e) { log.textContent += "FAILED: " + e.message + "\n"; }
    }
    await loadAll(); log.textContent += "Finished. Next: import data, then permissions.\n";
    setTimeout(render, 1200);
  };

  const imp = (id, fn) => { $(id).onchange = async (e) => {
    const f = e.target.files[0]; if (!f) return; const log = $("#importLog");
    try { log.textContent = `Reading ${f.name}…\n`; const wb = await readSheet(f); log.textContent += await fn(wb) + "\n"; await loadAll(); setTimeout(render, 1500); }
    catch (err) { log.textContent += "FAILED: " + err.message + "\n"; }
  }; };
  imp("#fStudents", async (wb) => {
    const recs = rowsOf(wb, /student/i).map((r) => ({
      Title: col(r, "Student_ID", "StudentID", "ID"), FullName: col(r, "Full_Name", "Name"), CalledName: col(r, "Called_Name", "Nickname"),
      CertificateName: col(r, "Certificate_Name"), Grade: col(r, "Grade"), Email: lc(col(r, "Email")), Department: col(r, "Department"),
      Joined: col(r, "Joined"), Status: col(r, "Status") === "Withdrawn" ? "Withdrawn" : "Active" })).filter((r) => r.Title && r.FullName);
    const unknownGrades = [...new Set(recs.map((r) => r.Grade).filter((g) => !P().grades.includes(g)))];
    const { sent, failed } = await upsert(L.students, recs, (f) => f.Title, S.students);
    return `Roster: ${recs.length} students read, ${sent} added/updated, ${failed.length} failed.` + (unknownGrades.length ? ` Note: grades not in the app's grade list: ${unknownGrades.join(", ")}.` : "");
  });
  imp("#fSubjects", async (wb) => {
    const seen = new Set(), recs = [];
    rowsOf(wb, /subjects/i).forEach((r, i) => {
      const grade = col(r, "Grade"), name = col(r, "Suggested canonical name", "Subject", "Subject (as used in Spring 2026)");
      if (!grade || !name || seen.has(grade + "|" + name)) return; seen.add(grade + "|" + name);
      recs.push({ Title: name, Grade: grade, SubjectType: /ii/i.test(col(r, "Subject I / II", "Subject_Type", "Type")) ? "II" : "I",
        Teacher: col(r, "Teacher for Fall 2026", "Teacher"), PrintsOnCard: !/^no/i.test(col(r, "Prints on report card?", "PrintsOnCard")), SortOrder: i + 1 });
    });
    const { sent, failed } = await upsert(L.subjects, recs, (f) => f.Grade + "|" + f.Title, S.subjects);
    return `Subjects: ${recs.length} read, ${sent} added/updated, ${failed.length} failed.`;
  });
  imp("#fCatalogue", async (wb) => {
    const all = await Promise.resolve(readList(L.catalogue));
    const recs = rowsOf(wb, /choice list/i).map((r, i) => ({
      Title: col(r, "Award name", "Award name (cleaned)", "Award_Name"), Category: col(r, "Category"),
      CertificateType: col(r, "Certificate type", "Certificate_Type"), Active: true, SortOrder: i + 1 })).filter((r) => r.Title && r.Category);
    const { sent, failed } = await upsert(L.catalogue, recs, (f) => f.Category + "|" + f.Title, all);
    return `Award catalogue: ${recs.length} awards read, ${sent} added/updated, ${failed.length} failed.`;
  });

  $("#saveSettings").onclick = async () => {
    const lines = (id) => $(id).value.split(/[\s,;]+/).map(lc).filter(Boolean);
    const advisors = {};
    for (const g of P().grades) advisors[g] = { name: $(`.advName[data-g="${g}"]`).value.trim(), email: lc($(`.advEmail[data-g="${g}"]`).value) };
    const json = { academicYear: $("#sYear").value.trim(), semester: $("#sSem").value, admins: lines("#sAdmins"), leadership: lines("#sLead"), advisors };
    if (!json.admins.includes(S.me) && !confirm("You are removing yourself from the admins. Continue?")) return;
    try {
      if (S.settingsItem) await graph("PATCH", `${itemsUrl(L.settings)}/${S.settingsItem.id}/fields`, { Json: JSON.stringify(json, null, 1) });
      else await graph("POST", itemsUrl(L.settings), { fields: { Title: "config", Json: JSON.stringify(json, null, 1) } });
      await loadAll(); render(); flash("ok", "Saved. Re-run Permissions if advisors changed.");
    } catch (e) { flash("error", e.message); }
  };

  $("#applyPerms").onclick = async () => {
    const log = $("#permLog"); const out = (s) => (log.textContent += s + "\n");
    log.textContent = ""; $("#applyPerms").disabled = true;
    try { await applyPermissions(specs.filter((s) => S.lists[s.name]), $("#staffEmails").value.split(/[\s,;]+/).map(lc).filter((e) => e.includes("@")), out); out("Done. Check one list's Permissions page to confirm."); }
    catch (e) { out("FAILED: " + e.message); out("Use the manual links below instead."); }
    $("#applyPerms").disabled = false;
  };
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
  const advisorEmails = Object.values(P().advisors).map((a) => lc(a.email)).filter(Boolean);
  const everyone = [...new Set([...staffEmails, ...advisorEmails])];
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
    if (s.kind === "marks") {
      const adv = lc(P().advisors[s.grade]?.email);
      if (adv) {
        const u = await sp("POST", "web/ensureuser", { logonName: claims(adv) });
        await sp("POST", `${base}/roleassignments/addroleassignment(principalid=${u.Id},roledefid=${role(4)})`);
        extra = ` · advisor ${adv}: Design`;
      } else extra = " · no advisor set";
    }
    out(`✓ ${s.name}: Owners Full Control · ${C.staffGroup} ${s.kind === "read" ? "Read" : "Contribute, item-level"}${extra}`);
  }
}

// ── PDE (placeholder) ────────────────────────────────────────
function viewPde(main) {
  main.innerHTML = `
    <div class="card pde-hero">
      <div class="big">🌱</div>
      <h1>Purpose Driven Education</h1>
      <p><span class="badge warn">In development</span></p>
      <p class="lede" style="max-width:620px;margin:10px auto">PDE is being designed with our teachers. When it opens here, each bobae's goals, core-value ratings and reflections will sit alongside the marks you already enter, so no score is ever typed twice.</p>
    </div>
    <div class="card"><h3>What's planned</h3>
      <ul class="small">
        <li><b>Goals</b>: a target score and action plan per subject. Bobaes in G6–G12 set their own; advisors enter them for G1–G5.</li>
        <li><b>Core values</b>: Wisdom, Holistic Education, Self Expression, Responsibility and Respect, rated once per bobae per semester instead of in every subject.</li>
        <li><b>Subject teachers</b>: one short comment on each goal. The achieved score comes from the marks entered here.</li>
        <li><b>Leadership</b>: core values out of 20, and goals met out of goals set, by class and school.</li>
      </ul>
      <p class="small muted">Timeline: Fall 2026, the teacher working group agrees the framework. Spring 2027, PDE opens here and can print as page 3 of the report card.</p>
    </div>`;
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
    S.me = lc(account.username); S.meName = account.name || S.me;
    S.loadErr = null;
    await loadAll();
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
