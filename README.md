# WZFS Reporting Portal

Teachers enter **marks, behaviour, subject remarks and awards** for their own classes. Advisors collate and write comments. The Admin Office exports files for the existing report card generator ([wzfs-report-card.vercel.app](https://wzfs-report-card.vercel.app/)).

- **Database:** SharePoint lists on a private site in the school tenant. No server, no secrets.
- **Sign-in:** school Microsoft account (MSAL, delegated).
- **Privacy:** enforced by SharePoint item-level permissions. Each teacher reads and edits only the rows they created.
- **China without a VPN:** the page is static and its libraries are in `/vendor`. It only talks to `login.microsoftonline.com` and `graph.microsoft.com`.

## Files

| File | What it is |
|---|---|
| `index.html` | Page shell and styles (WZFS navy) |
| `app.js` | All logic: sign-in, My classes, marks entry, Advisor view, Awards, Export, Setup, PDE placeholder |
| `config.js` | **The only file you normally edit.** Client ID, site path, term, people |
| `vendor/` | Pinned `msal-browser@3.27.0` and `xlsx@0.18.5` |

## SharePoint lists

| List | Key (Title) | Staff access |
|---|---|---|
| Students | Student ID | Read |
| Subjects | Subject name (+ Grade) | Read |
| Award Catalogue | Award name (+ Category) | Read |
| Reporting Settings | `config` (JSON) | Read |
| Awards | Student ID | Contribute, item-level |
| Advisor Comments | Student (grade) | Contribute, item-level |
| Marks G6 … Marks ZSA (one per grade) | Student – Subject | Contribute, item-level; that grade's advisor has Design |

Score is limited to 0–100 and behaviour to 1–4 **in SharePoint itself**, so even editing the list directly can't store a bad value. Version history is on for every list.

## PDE: bobae goals (second SharePoint site)

Bobaes (G6–G12, ZSA) sign in with their school account and see only **My goals**: a target score, a one-sentence goal and an action plan for each core subject. Advisors enter goals for G1–G5. Subject teachers read their classes' goals and write one line of feedback (**PDE → Bobae goals**).

The data lives on a separate private site, **WZFS PDE** (`/sites/WZFSPDE`), so bobaes never hold any permission on the marks site:

| List | Bobaes | Staff (Reporting Staff group) |
|---|---|---|
| PDE Goals | Contribute, item-level (own goals only) | Design (read all, write feedback) |
| PDE Students | Each row shared only with its own bobae | Read |
| PDE Settings | Read (core subjects per grade + current term) | Read |

Set up in **Setup → 7 · PDE site**: create the lists, check the core subjects, then **Sync bobaes and apply permissions** (re-run after roster changes; only new bobaes are processed). The app recognises a bobae because their account can't open the marks site but appears in PDE Students.

Limit: SharePoint can't lock single columns, so a bobae could edit the feedback text on their own goal directly in SharePoint. The app never allows it, and version history shows who changed what.

## First-time setup

1. **SharePoint site**: create a Team site named `WZFS Reporting` (path `/sites/WZFSReporting`), set to Private. Owners: ICT, leadership, Admin Office. Don't add teachers as members.
2. **Entra app registration**: `WZFS Reporting Portal` → Single-page application.
   - Redirect URIs: `http://localhost:5173/` and the Vercel URL (ending in `/`).
   - Microsoft Graph delegated: `User.Read`, `Sites.ReadWrite.All`, `Sites.Manage.All`.
   - Optional: SharePoint delegated `AllSites.FullControl` with admin consent. This enables one-click permissions.
   - Put the Application (client) ID in `config.js`.
3. Open the app as an admin and go to **Setup**:
   1. Create the missing lists
   2. Import the roster, subjects and award catalogue (the prepared Fall 2026 workbooks)
   3. Confirm the term and the class advisors' emails
   4. Paste the staff emails, then **Apply permissions**. If that fails, use the manual links.
4. **Test with a second, non-admin account** before telling staff. It should see only its own classes, and the SharePoint list itself should show only its own rows.

## End of term

**Export** tab → marks file per grade + awards file → upload to the report card generator with the attendance sheet.

The export uses the generator's long format: `Student_ID, Student_Name, Called_Name, Grade, Advisor, Grading_Period, Subject_Name, Subject_Type, Score, Behaviour, Subject_Remark, Teacher_Comments, Subject_Teacher`.
Awards: `Student_ID, Student_Name, Grade, Award, Recognition, Points, Date`, plus a certificate print list sheet.

## Local testing

Any static server on port 5173, for example `npx serve -l 5173 .` (needs Node.js).

## Fallback with zero code

Staff can use the SharePoint lists directly with the list's **+ New** form. Item-level privacy still applies.

## Roadmap

1. **Fall 2026**: this web app. Marks, remarks, advisor comments, awards, export.
2. **Generator**: match awards by `Student_ID`, and print per-subject remarks.
3. **Spring 2027**: the PDE module (tab is in place, marked *in development*).
4. **Power Apps** front end on the same lists, added to Teams, like Purchase Requests.
