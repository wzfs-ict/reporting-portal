"use strict";
// ─────────────────────────────────────────────────────────────
//  PDE: bobae goals and action plans (Academic Purpose-Driven Education)
//  Lives on its OWN SharePoint site so bobaes never hold permission on the marks site.
//    PDE Goals     one row per bobae per core subject per semester (item-level: own rows only)
//    PDE Students  roster copy; each row is shared only with that bobae (unique per-item permissions)
//    PDE Settings  core subjects per grade (everyone reads)
//  Loaded after app.js and uses its helpers (graph, batch, S, P, esc, lc, …).
// ─────────────────────────────────────────────────────────────
const PD = () => ({
  sitePath: "/sites/WZFSPDE",
  lists: { goals: "PDE Goals", students: "PDE Students", settings: "PDE Settings" },
  studentGroup: "PDE Students",
  advisorEntryGrades: ["G1", "G2", "G3", "G4", "G5"],       // advisors enter goals for these grades
  coreSubjects: "chinese|english|literature|grammar|reading|written|writing|ielts|math|science|biology|chemistry|physics",
  goalMax: 200, planMax: 600, feedbackMax: 300,
  ...(C.pde || {}), ...((S.settings && S.settings.pde) || {}),
});
const selfEntryGrades = () => P().grades.filter((g) => !PD().advisorEntryGrades.includes(g));
S.pde = { siteId: "", siteUrl: "", lists: {}, goals: [], students: [], subjects: {}, settingsItem: null, err: null, loaded: false };

const pdeItems = (list) => `/sites/${S.pde.siteId}/lists/${S.pde.lists[PD().lists[list]].id}/items`;
async function pdeRead(list) {
  const l = S.pde.lists[PD().lists[list]]; if (!l) return [];
  return (await graphAll(`/sites/${S.pde.siteId}/lists/${l.id}/items?$expand=fields&$top=500`)).map(norm);
}
async function loadPde() {
  S.pde = { ...S.pde, err: null };
  try {
    const site = await graph("GET", `/sites/${C.siteHostname}:${PD().sitePath}?$select=id,webUrl`);
    S.pde.siteId = site.id; S.pde.siteUrl = site.webUrl;
    S.pde.lists = Object.fromEntries((await graphAll(`/sites/${site.id}/lists?$select=id,displayName,webUrl`)).map((l) => [l.displayName, l]));
    const [goals, students, settings] = await Promise.all([pdeRead("goals"), pdeRead("students"), pdeRead("settings")]);
    S.pde.settingsItem = settings.find((x) => x.f.Title === "subjects") || null;
    let js = {}; try { js = S.pde.settingsItem ? JSON.parse(S.pde.settingsItem.f.Json || "{}") : {}; } catch (_) {}
    S.pde.subjects = js.grades || {}; S.pde.term = js.academicYear ? { academicYear: js.academicYear, semester: js.semester } : null;
    // Bobaes can't read the marks site's settings, so the current term travels with the PDE settings.
    if (S.studentMode !== false && S.pde.term && !S.settingsItem) S.settings = { ...S.settings, ...S.pde.term };
    S.pde.goals = goals.filter((x) => inTerm(x.f));
    S.pde.students = students;
    S.pde.loaded = true;
  } catch (e) { S.pde.err = e; S.pde.loaded = false; }
}
// Signed-in account can't open the marks site: is it a bobae on the PDE site?
async function tryStudentMode() {
  await loadPde();
  if (S.pde.err) return false;
  const me = S.pde.students.find((s) => lc(s.f.Email) === S.me);
  if (!me) return false;
  S.studentMode = true; S.student = me.f; S.loadErr = null; S.tab = "mygoals";
  return true;
}
const goalOf = (studentId, subject) => S.pde.goals.find((g) => g.f.StudentID === studentId && g.f.Subject === subject);
const pdeSubjectsFor = (grade) => S.pde.subjects[grade] || [];

// One goal editor card. who = "bobae" | "advisor" (G1–G5) | "view"
function goalCard(st, subject, who) {
  const g = goalOf(st.Title, subject), f = g ? g.f : {};
  const ro = who === "view";
  const id = `${st.Title}|${subject}`.replace(/[^\w|-]/g, "_");
  return `<div class="card goal-card" data-sub="${esc(subject)}" data-sid="${esc(st.Title)}" style="margin:0">
    <div class="row" style="justify-content:space-between;align-items:center"><h3 style="margin:0">${esc(subject)}</h3>
      ${g ? `<span class="badge ok">Saved</span>` : `<span class="badge">Not set yet</span>`}</div>
    <label style="margin-top:8px"><span>Target score (0–100)</span><input class="g-target" inputmode="decimal" style="max-width:120px" value="${blank(f.TargetScore) ? "" : esc(f.TargetScore)}" ${ro ? "disabled" : ""} id="t_${id}"></label>
    <label style="margin-top:8px"><span>My goal (one sentence)</span><input class="g-goal" maxlength="${PD().goalMax}" value="${esc(f.Goal || "")}" ${ro ? "disabled" : ""} placeholder="e.g. Improve my essay writing so I reach 85% by the final exam"></label>
    <label style="margin-top:8px"><span>My action plan: what I will do, and how my teacher will help</span><textarea class="g-plan" rows="3" maxlength="${PD().planMax}" ${ro ? "disabled" : ""} placeholder="e.g. Write one paragraph every week and ask my teacher for feedback">${esc(f.ActionPlan || "")}</textarea></label>
    ${f.TeacherFeedback ? `<div class="notice" style="margin:8px 0 0"><b>Teacher's feedback</b>${f.FeedbackBy ? ` <span class="muted small">(${esc(f.FeedbackBy)})</span>` : ""}<br>${esc(f.TeacherFeedback)}</div>` : ""}
    ${ro ? "" : `<p style="margin:10px 0 0"><button class="btn primary small g-save">Save</button> <span class="small muted g-state"></span></p>`}
  </div>`;
}
function bindGoalCards(root, enteredBy) {
  $$(".goal-card", root).forEach((card) => {
    const btn = $(".g-save", card); if (!btn) return;
    btn.onclick = async () => {
      const st = S.studentMode && !S.viewAs ? S.student : (studentById(card.dataset.sid)?.f || S.student);
      const subject = card.dataset.sub, state = $(".g-state", card);
      const t = $(".g-target", card).value.trim(), n = Number(t);
      if (t !== "" && (!Number.isFinite(n) || n < 0 || n > 100)) { state.textContent = "Target must be 0–100"; state.style.color = "var(--bad)"; return; }
      const fields = { TargetScore: t === "" ? null : Math.round(n * 10) / 10, Goal: $(".g-goal", card).value.trim(), ActionPlan: $(".g-plan", card).value.trim() };
      if (fields.TargetScore === null && !fields.Goal && !fields.ActionPlan) { state.textContent = "Write a target, goal or plan first"; return; }
      btn.disabled = true; state.style.color = ""; state.textContent = "Saving…";
      try {
        const g = goalOf(st.Title, subject);
        if (g) { await graph("PATCH", `${pdeItems("goals")}/${g.id}/fields`, fields); Object.assign(g.f, fields); }
        else {
          const it = await graph("POST", pdeItems("goals"), { fields: { Title: `${st.CalledName || st.FullName} – ${subject}`, StudentID: st.Title, Email: lc(st.Email), FullName: st.FullName,
            CalledName: st.CalledName || "", Grade: st.Grade, Subject: subject, ...term(), EnteredBy: enteredBy, ...fields } });
          S.pde.goals.push(norm(it));
        }
        state.textContent = "Saved ✓"; const b = $(".badge", card); b.className = "badge ok"; b.textContent = "Saved";
        const counter = $("#goalCount"); if (counter) counter.textContent = `${pdeSubjectsFor(st.Grade).filter((s) => goalOf(st.Title, s)).length} / ${pdeSubjectsFor(st.Grade).length}`;
      } catch (e) { state.style.color = "var(--bad)"; state.textContent = e.message; }
      btn.disabled = false;
    };
  });
}

// ── Bobae: My goals ─────────────────────────────────────────
function viewMyGoals(main) {
  let st = S.student;
  if (S.viewAs === "student") {   // admin preview: choose which bobae
    const pool = selfEntryGrades().flatMap(rosterFor);
    if (!S.previewStudent || !studentById(S.previewStudent)) S.previewStudent = pool[0]?.f.Title || "";
    st = studentById(S.previewStudent)?.f;
  }
  if (!st) { main.innerHTML = notice("warn", "No bobae to show."); return; }
  const subs = pdeSubjectsFor(st.Grade);
  const done = subs.filter((s) => goalOf(st.Title, s)).length;
  main.innerHTML = `
    ${S.viewAs === "student" ? `<label class="no-print" style="max-width:320px;margin-bottom:12px"><span>Preview as bobae</span><select id="pvBobae">${selfEntryGrades().map((g) => { const r = rosterFor(g); return r.length ? `<optgroup label="${esc(g)}">${r.map((s) => `<option value="${esc(s.f.Title)}" ${s.f.Title === st.Title ? "selected" : ""}>${esc(s.f.CalledName || s.f.FullName)}</option>`).join("")}</optgroup>` : ""; }).join("")}</select></label>` : ""}
    <h1>Hi ${esc(st.CalledName || st.FullName)} 👋</h1>
    <p class="lede"><span class="badge info">${esc(P().semester)} ${esc(P().academicYear)}</span> ${esc(st.Grade)} · Set a goal for each subject and plan how you'll reach it. You can change it any time this semester.</p>
    <div class="stats"><div class="stat"><div class="k">Goals set</div><div class="v" id="goalCount">${done} / ${subs.length}</div></div></div>
    ${!subs.length ? notice("warn", "Your subjects haven't been set up yet. Please tell your advisor.") : ""}
    <div class="class-grid" style="grid-template-columns:repeat(auto-fill,minmax(300px,1fr))">${subs.map((s) => goalCard(st, s, "bobae")).join("")}</div>
    <p class="small muted" style="margin-top:14px">Only you, your teachers and your advisor can see your goals.</p>`;
  bindGoalCards(main, "Bobae");
  const pv = $("#pvBobae"); if (pv) pv.onchange = (e) => { S.previewStudent = e.target.value; render(); };
}

// ── Staff: Bobae goals (inside the PDE tab) ──────────────────
function viewPdeGoals(box) {
  if (!S.pde.loaded) {
    box.innerHTML = notice("warn", isAdmin()
      ? `The PDE site isn't set up yet${S.pde.err ? ` (${esc(S.pde.err.message)})` : ""}. Create the SharePoint site <code>${esc(PD().sitePath)}</code>, then use <b>Setup → 7 · PDE site</b>.`
      : "Bobae goals aren't open yet.");
    return;
  }
  // 1) Subject teacher: my classes → goals for that subject + one feedback box each.
  const myClasses = groupClasses(allMarks().filter((r) => r.by === S.me));
  const teach = myClasses.map((c) => {
    const rows = c.rows.map((r) => ({ r, st: studentById(r.f.StudentID)?.f, g: goalOf(r.f.StudentID, c.subject) })).filter((x) => x.st);
    return { c, rows, set: rows.filter((x) => x.g).length };
  }).filter((x) => x.rows.length && x.rows.some((y) => pdeSubjectsFor(y.st.Grade).includes(x.c.subject)));
  // 2) Advisor / head / leader: grade overview; advisors of G1–G5 enter goals on behalf.
  const grades = viewGrades();
  if (!S.pdeGrade || !grades.includes(S.pdeGrade)) S.pdeGrade = grades[0] || "";
  const g = S.pdeGrade, roster = g ? rosterFor(g) : [], subs = pdeSubjectsFor(g);
  const onBehalf = PD().advisorEntryGrades.includes(g) && canAdvise(g);
  box.innerHTML = `
    <h2>Bobae goals <span class="badge ok">Live</span></h2>
    <p class="small muted">Goals bobaes have set this semester (${esc(P().semester)} ${esc(P().academicYear)}). G6–G12 and ZSA enter their own; advisors enter them for G1–G5.</p>
    ${teach.length ? `<h3>Your classes</h3>${teach.map(({ c, rows, set }) => `
      <details class="card" ${teach.length === 1 ? "open" : ""}><summary><b>${esc(c.grade)} · ${esc(subjectDisplay({ Subject: c.subject, ClassLevel: c.level }))}</b> <span class="badge ${set === rows.length ? "ok" : "warn"}">${set} of ${rows.length} goals set</span></summary>
        <div class="table-wrap" style="margin-top:10px"><table><thead><tr><th>Bobae</th><th class="num">Target</th><th>Goal and action plan</th><th style="width:34%">Your feedback (one line, saves when you leave the box)</th><th></th></tr></thead><tbody>
          ${rows.map(({ st, g: gl }) => `<tr><td><b>${esc(st.CalledName || st.FullName)}</b><div class="small muted">${esc(st.Grade)}</div></td>
            <td class="num">${gl && !blank(gl.f.TargetScore) ? esc(gl.f.TargetScore) + "%" : "–"}</td>
            <td class="small">${gl ? `${esc(gl.f.Goal || "")}${gl.f.ActionPlan ? `<div class="muted">${esc(gl.f.ActionPlan)}</div>` : ""}` : `<span class="muted">Not set yet</span>`}</td>
            <td>${gl ? `<textarea class="fb" data-gid="${gl.id}" rows="2" maxlength="${PD().feedbackMax}">${esc(gl.f.TeacherFeedback || "")}</textarea>` : ""}</td><td class="rowstate"></td></tr>`).join("")}
        </tbody></table></div></details>`).join("")}` : ""}
    ${grades.length ? `
      <h3 style="margin-top:18px">Grade overview</h3>
      <div class="row" style="margin-bottom:10px"><label style="max-width:160px"><span>Grade</span><select id="pdeGrade">${grades.map((x) => `<option ${x === g ? "selected" : ""}>${esc(x)}</option>`).join("")}</select></label>
        <span class="small muted" style="padding-bottom:8px">${subs.length ? `Core subjects: ${subs.map(esc).join(", ")}` : "No core subjects set for this grade yet (Setup → 7)."}</span></div>
      <div class="table-wrap"><table class="matrix"><thead><tr><th class="sticky">Bobae</th>${subs.map((s) => `<th class="center small">${esc(s)}</th>`).join("")}<th class="num">Set</th>${onBehalf ? "<th></th>" : ""}</tr></thead><tbody>
        ${roster.map((s) => { const n = subs.filter((x) => goalOf(s.f.Title, x)).length; return `<tr><td class="sticky"><b>${esc(s.f.CalledName || s.f.FullName)}</b></td>
          ${subs.map((x) => { const gl = goalOf(s.f.Title, x); return `<td class="cell ${gl ? "" : "missing"}" title="${gl ? esc((gl.f.Goal || "") + " — " + (gl.f.ActionPlan || "")) : "Not set"}">${gl ? (blank(gl.f.TargetScore) ? "✓" : esc(gl.f.TargetScore) + "%") : "·"}</td>`; }).join("")}
          <td class="num">${n}/${subs.length}</td>${onBehalf ? `<td><button class="btn small" data-enter="${esc(s.f.Title)}">Enter goals</button></td>` : ""}</tr>`; }).join("") || `<tr><td class="muted" colspan="${subs.length + 2}">No bobaes.</td></tr>`}
      </tbody></table></div>
      <div id="enterBox"></div>` : ""}
    ${!teach.length && !grades.length ? `<p class="muted">Goals appear here for the classes you teach.</p>` : ""}`;
  $$(".fb", box).forEach((ta) => (ta.onchange = async () => {
    const st = ta.closest("tr").querySelector(".rowstate"); st.className = "rowstate"; st.textContent = "Saving…";
    const fields = { TeacherFeedback: ta.value.trim(), FeedbackBy: S.meName, FeedbackAt: new Date().toISOString() };
    try { await graph("PATCH", `${pdeItems("goals")}/${ta.dataset.gid}/fields`, fields); Object.assign(S.pde.goals.find((x) => x.id === ta.dataset.gid).f, fields); st.textContent = "Saved ✓"; }
    catch (e) { st.className = "rowstate err"; st.textContent = "Not saved"; st.title = e.message; }
  }));
  const gs = $("#pdeGrade", box); if (gs) gs.onchange = (e) => { S.pdeGrade = e.target.value; render(); };
  $$("[data-enter]", box).forEach((b) => (b.onclick = () => {
    const st = studentById(b.dataset.enter).f;
    $("#enterBox", box).innerHTML = `<h3 style="margin-top:16px">Goals for ${esc(st.CalledName || st.FullName)} (entered by you as advisor)</h3>
      <div class="class-grid" style="grid-template-columns:repeat(auto-fill,minmax(300px,1fr))">${subs.map((x) => goalCard(st, x, "advisor")).join("")}</div>`;
    bindGoalCards($("#enterBox", box), "Advisor: " + S.meName);
    $("#enterBox", box).scrollIntoView({ behavior: "smooth" });
  }));
}

// ── Setup → 7 · PDE site ─────────────────────────────────────
function pdeListSpecs() {
  const l = PD().lists;
  return [
    { name: l.goals, kind: "goals", cols: [T("StudentID"), T("Email"), T("FullName"), T("CalledName"), T("Grade"), T("Subject"), N("TargetScore", 0, 100, "one"),
      T("Goal", true), T("ActionPlan", true), T("Semester"), T("AcademicYear"), T("TeacherFeedback", true), T("FeedbackBy"), { name: "FeedbackAt", dateTime: { format: "dateTime" } }, T("EnteredBy")] },
    { name: l.students, kind: "roster", cols: [T("Email"), T("FullName"), T("CalledName"), T("Grade"), T("PermsFor")] },
    { name: l.settings, kind: "read", cols: [T("Json", true)] },
  ];
}
function defaultCoreSubjects() {
  const rx = new RegExp(PD().coreSubjects, "i"), out = {};
  for (const g of P().grades) out[g] = [...new Set(S.subjects.filter((s) => s.f.Grade === g && s.f.PrintsOnCard !== false && s.f.SubjectType !== "II" && rx.test(s.f.Title)).map((s) => s.f.Title))];
  return out;
}
function pdeSetupCard() {
  const ok = S.pde.siteId, specs = pdeListSpecs(), missing = specs.filter((s) => !S.pde.lists[s.name]);
  const subj = Object.keys(S.pde.subjects).length ? S.pde.subjects : defaultCoreSubjects();
  const upper = selfEntryGrades().flatMap(rosterFor).filter((s) => s.f.Email);
  return `<div class="card"><h3>7 · PDE site (bobae goals)</h3>
    ${!ok ? notice("warn", `Create a <b>Private Team site</b> named <b>WZFS PDE</b> at <code>${esc(PD().sitePath)}</code> (Owners: you and leadership; no members). Then refresh this page.${S.pde.err ? `<br><span class="small">${esc(S.pde.err.message)}</span>` : ""}`) : `
      <p class="small">Site: <a href="${esc(S.pde.siteUrl)}" target="_blank" rel="noopener">${esc(S.pde.siteUrl)}</a> · ${specs.length - missing.length} of ${specs.length} lists · ${S.pde.students.length} bobaes synced · ${S.pde.goals.length} goals this semester</p>
      <p><button class="btn ${missing.length ? "primary" : ""}" id="pdeLists">${missing.length ? "a) Create the PDE lists" : "a) Check PDE lists for new columns"}</button></p>
      <h3 style="margin-top:12px">b) Core subjects per grade</h3>
      <p class="small muted">Bobaes set a goal for each subject listed for their grade. Suggested from the Subjects list (Chinese, English of every kind, Maths, Science); edit freely, one grade per line: <code>G7: English Language; Maths; Science; Chinese</code></p>
      <textarea id="pdeSubjects" rows="8" style="font-family:Consolas,monospace;font-size:12.5px">${esc(P().grades.map((g) => `${g}: ${(subj[g] || []).join("; ")}`).join("\n"))}</textarea>
      <p><button class="btn" id="pdeSaveSubjects">Save core subjects</button></p>
      <h3 style="margin-top:12px">c) Bobaes and permissions</h3>
      <p class="small">Copies the ${upper.length} G6–G12 / ZSA bobaes with a school email to <b>PDE Students</b>, shares each row with that bobae only, puts them in the <b>${esc(PD().studentGroup)}</b> group, and sets:
        <b>PDE Goals</b>: bobaes Contribute with item-level privacy (own goals only), staff Design (read all, write feedback) · <b>PDE Settings</b>: everyone Read.
        Needs the SharePoint AllSites.FullControl permission, like step 4. Safe to run again; only new bobaes are processed.</p>
      <p><button class="btn primary" id="pdeSync">Sync bobaes and apply permissions</button></p>`}
    <pre id="pdeLog" class="small" style="white-space:pre-wrap"></pre></div>`;
}
function bindPdeSetup() {
  const log = $("#pdeLog"); if (!log) return;
  const out = (s) => { log.textContent += s + "\n"; };
  const b1 = $("#pdeLists");
  if (b1) b1.onclick = async () => {
    b1.disabled = true; log.textContent = "";
    for (const s of pdeListSpecs()) {
      try {
        if (!S.pde.lists[s.name]) { await graph("POST", `/sites/${S.pde.siteId}/lists`, { displayName: s.name, list: { template: "genericList" }, columns: s.cols }, SETUP_SCOPES); out(`Created ${s.name}`); }
        else {
          const have = new Set((await graph("GET", `/sites/${S.pde.siteId}/lists/${S.pde.lists[s.name].id}/columns?$select=name`, null, SETUP_SCOPES)).value.map((c) => c.name));
          const add = s.cols.filter((c) => !have.has(c.name));
          for (const c of add) await graph("POST", `/sites/${S.pde.siteId}/lists/${S.pde.lists[s.name].id}/columns`, c, SETUP_SCOPES);
          out(`${s.name}: ${add.length ? "added " + add.map((c) => c.name).join(", ") : "up to date"}`);
        }
      } catch (e) { out(`${s.name}: FAILED ${e.message}`); }
    }
    await loadPde(); out("Done."); setTimeout(render, 1200);
  };
  const b2 = $("#pdeSaveSubjects");
  if (b2) b2.onclick = async () => {
    const subjects = {};
    for (const line of $("#pdeSubjects").value.split("\n")) {
      const i = line.indexOf(":"); if (i < 0) continue;
      const g = line.slice(0, i).trim().toUpperCase(); if (!P().grades.includes(g)) continue;
      subjects[g] = [...new Set(line.slice(i + 1).split(";").map((x) => x.replace(/\s+/g, " ").trim()).filter(Boolean))];
    }
    try { await savePdeSettings(subjects); await loadPde(); flash("ok", `Saved core subjects for ${Object.keys(subjects).length} grades.`); }
    catch (e) { flash("error", e.message); }
  };
  const b3 = $("#pdeSync");
  if (b3) b3.onclick = async () => {
    b3.disabled = true; log.textContent = "";
    try { await pdeSyncAndPermissions(out); out("Done. Test with a bobae's account, or Setup → View as → A bobae."); }
    catch (e) { out("FAILED: " + e.message); }
    b3.disabled = false; await loadPde();
  };
}
// Core subjects + the current term, which bobaes need because they can't read the marks site's settings.
async function savePdeSettings(subjects = S.pde.subjects) {
  if (!S.pde.lists[PD().lists.settings]) return;
  const body = { Json: JSON.stringify({ grades: subjects, academicYear: P().academicYear, semester: P().semester }) };
  if (S.pde.settingsItem) await graph("PATCH", `${pdeItems("settings")}/${S.pde.settingsItem.id}/fields`, body);
  else await graph("POST", pdeItems("settings"), { fields: { Title: "subjects", ...body } });
}
async function pdeSyncAndPermissions(out) {
  const site = S.pde.siteUrl, api = (m, p, b, h) => spAt(site, m, p, b, h);
  // 1. Roster copy (Graph): add/update one row per G6–G12 / ZSA bobae with an email.
  const upper = selfEntryGrades().flatMap(rosterFor).filter((s) => s.f.Email);
  const recs = upper.map((s) => ({ Title: s.f.Title, Email: lc(s.f.Email), FullName: s.f.FullName, CalledName: s.f.CalledName || "", Grade: s.f.Grade }));
  const existing = new Map(S.pde.students.map((x) => [x.f.Title, x]));
  const reqs = [];
  for (const r of recs) { const cur = existing.get(r.Title);
    if (!cur) reqs.push({ method: "POST", url: pdeItems("students"), body: { fields: r } });
    else if (["Email", "FullName", "CalledName", "Grade"].some((k) => String(cur.f[k] || "") !== String(r[k]))) reqs.push({ method: "PATCH", url: `${pdeItems("students")}/${cur.id}/fields`, body: { ...r, PermsFor: "" } }); }
  const res = await batch(reqs); out(`Roster: ${recs.length} bobaes, ${reqs.length} added/updated, ${batchFailures(res).length} failed.`);
  await loadPde();
  // 2. Groups (SharePoint REST, PDE site).
  const roles = (await api("GET", "web/roledefinitions?$select=Id,RoleTypeKind")).value, role = (k) => roles.find((r) => r.RoleTypeKind === k).Id;
  const owners = await api("GET", "web/associatedownergroup?$select=Id");
  const group = async (title, desc) => (await api("GET", `web/sitegroups/getbyname('${encodeURIComponent(title)}')?$select=Id`).catch(() => null))
    || await api("POST", "web/sitegroups", { __metadata: { type: "SP.Group" }, Title: title, Description: desc });
  const staffG = await group(C.staffGroup, "Teachers and advisors: read bobae goals and write feedback");
  const stuG = await group(PD().studentGroup, "Bobaes (G6–G12, ZSA): enter their own goals");
  const staff = [...new Set([...(P().staff || []), ...P().grades.flatMap(advisorEmails), ...(P().reportTeam || []), ...P().admins, ...Object.keys(P().divisionHeads || {})].map(lc))];
  const addTo = async (g, emails, label) => { let bad = 0; for (const e of emails) { try { await api("POST", `web/sitegroups(${g.Id})/users`, { __metadata: { type: "SP.User" }, LoginName: claims(e) }); } catch (_) { bad++; } } out(`${label}: ${emails.length - bad} added${bad ? `, ${bad} could not be added (check the account)` : ""}.`); };
  await addTo(staffG, staff, C.staffGroup);
  await addTo(stuG, recs.map((r) => r.Email), PD().studentGroup);
  // 3. List permissions.
  const lists = PD().lists;
  const setList = async (name, grants, itemLevel) => {
    const base = `web/lists(guid'${S.pde.lists[name].id}')`;
    await api("POST", `${base}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`);
    for (const [pid, rid] of [[owners.Id, role(5)], ...grants]) await api("POST", `${base}/roleassignments/addroleassignment(principalid=${pid},roledefid=${rid})`);
    await api("POST", base, { __metadata: { type: "SP.List" }, EnableVersioning: true, ...(itemLevel ? { ReadSecurity: 2, WriteSecurity: 2 } : {}) }, { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" });
    return base;
  };
  await setList(lists.goals, [[stuG.Id, role(3)], [staffG.Id, role(4)]], true); out(`✓ ${lists.goals}: bobaes Contribute (own goals only), staff Design`);
  await setList(lists.settings, [[stuG.Id, role(2)], [staffG.Id, role(2)]], false); out(`✓ ${lists.settings}: everyone Read`);
  const rbase = await setList(lists.students, [[staffG.Id, role(2)]], false); out(`✓ ${lists.students}: staff Read; each row shared with its own bobae…`);
  // 4. Each roster row: visible to that bobae only (plus owners and staff).
  const todo = S.pde.students.filter((x) => lc(x.f.PermsFor) !== lc(x.f.Email) && x.f.Email);
  let done = 0;
  for (const x of todo) {
    try {
      const u = await api("POST", "web/ensureuser", { logonName: claims(lc(x.f.Email)) });
      const ib = `${rbase}/items(${x.id})`;
      await api("POST", `${ib}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`);
      for (const [pid, rid] of [[owners.Id, role(5)], [staffG.Id, role(2)], [u.Id, role(2)]]) await api("POST", `${ib}/roleassignments/addroleassignment(principalid=${pid},roledefid=${rid})`);
      await graph("PATCH", `${pdeItems("students")}/${x.id}/fields`, { PermsFor: lc(x.f.Email) });
      done++; if (done % 10 === 0) out(`  …${done} of ${todo.length}`);
    } catch (e) { out(`  ! ${x.f.FullName}: ${e.message}`); }
  }
  out(`Roster rows shared: ${done} of ${todo.length} needed${todo.length ? "" : " (all already done)"}.`);
}
