# WZFS Reporting portal: notes for Claude

The owner is Glen Joshua, ICT Coordinator at Weihai Zhongshi Foreign School (WZFS). Students are called "bobaes".

## What this is
- A static site: `index.html`, `app.js`, `config.js`, `pde.js` and `certificates.html/js`. It is served from GitHub Pages (wzfs-ict.github.io/reporting-portal) and also from Vercel. There is no build step; pushing to `main` deploys it.
- The database is SharePoint lists on weihaizhongshi.sharepoint.com `/sites/WZFSReporting`. PDE data is on `/sites/WZFSPDE`. Sign-in uses MSAL and data goes through Microsoft Graph.
- `P()` returns `config.js` merged with the "Reporting Settings" list.
- Item-level security: each teacher can edit only their own rows.
- `cards/` is the built report card generator. Its source repo, `wzfs-report-card`, is kept locally only and **never pushed**.
- `cert/` holds the logos and signatures for certificates. Librarians' edits stay in browser localStorage and are never saved to SharePoint.

## Data model
- Marks lists are one per grade: "Marks G1" … "Marks G12", "Marks G9-1", "Marks G9-2", "Marks ZSA".
- A guest bobae in a mixed-grade class (Chinese levels, IGCSE, electives) keeps their row in **their own grade's list**. `ClassGrade` is the grade of the class they sit in.
- **Item ids repeat across lists**: G3 #41 and G4 #41 are different bobaes. Never match rows by id alone. Match by list + id, or by position in the rendered table.
  - This bug was fixed in 35b4137. It had saved guests' marks into another bobae's row.
- Duplicate key for marks: StudentID | Subject | ClassLevel. The same subject at two levels is allowed.
- Students list: StudentID is in `Title`. Other fields include CertificateName and Grade.
- Student ID format (switched 26 Oct 2026):
  - Regular IDs: `WZFS` + year of graduation (2 digits) + a 4-digit running serial. Serials are school-wide and never reused.
  - ZSA students use `9` + their ZSA number.
  - Ask Glen for the next free serial before adding a bobae.
- Versioning is on for every list (50 versions). The recycle bin keeps items for 93 days.

## Rules
- **No student data in git**: no rosters, names lists, marks or photos. That data lives in `.local/` (git-ignored) and in OneDrive on the school PC.
- Never read or copy the "Account Actions Archive" sheet.
- Never enter passwords. Glen sends his own emails and messages; draft them only.
- Keep PDE out of the role guides.
- Code style: match `app.js` (small helpers like `$`/`$$`, template-string HTML, short comments only where something is non-obvious).

## Working in the cloud vs on the school PC
A cloud session can edit, commit and push this repo. It **cannot**:
- open SharePoint as Glen (the Chrome extension is only on the school PC);
- open OneDrive files (rosters, classwise xlsx, signatures, staff guides);
- open `.local/`.

When a task needs any of those (data scans, restoring marks, roster or ID changes, new signatures), say so plainly. Glen will connect the VPN on the school PC and continue there.

Teachers can check data themselves: the **Health check** card on the Setup tab (orphan IDs, bad scores, duplicates, photos, recycle bin) and **Snapshot**, which downloads an Excel backup of everything.
