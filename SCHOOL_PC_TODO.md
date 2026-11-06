# To do on the school PC (from the cloud session, 2 Nov 2026)

Run `git pull` first. Work through this list in the live portal, signed in as Glen, then tick each item off and commit.
Refresh the page with Ctrl+F5 first, so the browser loads the new code. No student data goes in this file.

## 1. Awards
- [ ] Awards tab, at Midterm: the only category offered is **Honours & High Honours**.
- [ ] Use **View as** a subject teacher (e.g. Gregg): **PTL Awards** is not in the category list.
- [ ] **Report cards** tab: if the "category kept to named people" warning appears, click **Delete these N**. That removes every PTL award given by anyone other than Arsenio, Roger Russell, Beverly or an admin, including Gregg's.
- [ ] Confirm "Russell" is Roger Russell (rrussell@). If not, fix `awardCategoryGivers` in `config.js`.

## 1b. PTC sheets and scores (commits 80f6d08 and the rounding commit)
- [ ] Open a PTC sheet (advisor view) for a grade with marks. Check:
  - the bar colour follows the score: green Excellent, light green Good, amber Satisfactory, red Needs support. A low score with Excellent behaviour must show **red**;
  - behaviour is its own column headed "Behaviour"; "Needs Improvement" and ⚑ show in red;
  - the key at the bottom shows only the labels, no numbers.
- [ ] Check a parent view (View as → parent) shows the same.
- [ ] Scores are whole numbers: type 84.36 in a class and it saves as 84. Existing decimals (e.g. 84.4) show rounded on the class page, the PTC sheet and the report card export.
- [ ] Generate one report card and check the scores print as whole numbers.

## 1c. Department heads (new read-only "Department view")
- [ ] Setup → **Permissions** → run it. This gives Jayesh Patole (Science), Bijily Reacher (Maths), Mae Parker (Humanities) and Randy Wang (Chinese) read access to the G6–G12 marks lists. Without it their Department view stays empty.
- [ ] **View as** each of the four (Department heads group) and check the Department view lists only their subjects: Science = science subjects (not Computer Science or the Science Investigations elective), Maths, Humanities = English, Social Studies, Geography, History, IELTS, Chinese. Open a class: all boxes are locked.
- [ ] Confirm Mae Parker (mparker@) is the "Mae" meant, and Randy Wang's account is randy@. Fix `departmentHeads` in `config.js` if not.
- [ ] Note: SharePoint gives them Design access to those marks lists, so the app shows them only their subject, but SharePoint itself would let them see all subjects in those grades if they browsed the lists directly.

## 2. Subjects: Setup → card 5 → "Planned subject changes"
- [ ] Take a **Snapshot** first (Health check card).
- [ ] Click **Apply**. In one go, this:
  - **Maths (Bijily and Beverly, 3 Nov).** Each grade lists only its own subjects:
    - **G11:** AS Mathematics (Pure Math 1 & Probability and Statistics 1), AS Mathematics (Probability and Statistics 1), Mathematics. A2 Maths is hidden.
    - **G12:** A2 Mathematics (Pure Math 3 & Probability and Statistics 1), Pure Mathematics 1, Mathematics. AS Maths is hidden.
    - Michael's classes move with their marks. G11 "AS Maths (Pure Math 1 and Mechanics)" becomes the AS subject. His two G12 paper classes merge into the A2 subject: the row with marks is kept and the empty duplicate is deleted (it stays in the recycle bin). G12 "Maths" becomes Mathematics.
    - If Apply says a hidden subject **still has mark rows**, a class exists under it. Use **Rename** on that subject to move the class.
  - hides A2 Business Studies in G11 and G12;
  - adds Elective (Science Investigations), type II, to G1–G5 and ZSA.
- [ ] Every row should then show **Done**. If a row says "Old name not found", the live list uses a different name: use the **Rename** button in the subjects table for that one.
- [ ] If Apply reports **clashes**, a G12 bobae has marks on *both* papers. Those rows stay on Michael's old "AS Maths (Probability and Statistics 1)" card. Michael enters the combined mark on the A2 card, then deletes the old card.
- [ ] Ask Michael to refresh and check. He should see G11 AS Mathematics (Pure Math 1 & Probability and Statistics 1) and G12 A2 Mathematics (Pure Math 3 & Probability and Statistics 1), with one mark per bobae.
- [ ] The two bobaes not taking AS/A2: their maths teacher adds a **Mathematics** class for them (My classes → Add a class).
- [ ] After that, delete the `pendingSubjectChanges` entries from `config.js` and push.

## 3. Still open from before
- [ ] Classwise workbook: switch to the new IDs and add a Previous ID column.
- [ ] Upload the pending bobae photo (file name in `.local/`).
- [ ] Re-run the mixed-class check now that the midterm deadline has passed.
- [ ] Optional: add the Technology Record (Roadmap) to the Staff Hub links. It's internal, so share it with leadership only.

## Settled (no action)
- YCT4 score: whatever the teacher enters stays. Don't change it.
- The PDE Excellence and Resilience certificate wording stays as it is.
- ZSA keeps all 4 English courses.
