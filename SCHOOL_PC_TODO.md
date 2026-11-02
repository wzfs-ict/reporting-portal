# To do on the school PC (from the cloud session, 2 Nov 2026)

Run `git pull` first. Work through this list in the live portal, signed in as Glen, then tick each item off and commit.
Refresh the page with Ctrl+F5 first, so the browser loads the new code. No student data goes in this file.

## 1. Awards
- [ ] Awards tab, at Midterm: the only category offered is **Honours & High Honours**.
- [ ] Use **View as** a subject teacher (e.g. Gregg): **PTL Awards** is not in the category list.
- [ ] **Report cards** tab: if the "category kept to named people" warning appears, click **Delete these N**. That removes every PTL award given by anyone other than Arsenio, Roger Russell, Beverly or an admin, including Gregg's.
- [ ] Confirm "Russell" is Roger Russell (rrussell@). If not, fix `awardCategoryGivers` in `config.js`.

## 2. Subjects: Setup → card 5 → "Planned subject changes"
- [ ] Take a **Snapshot** first (Health check card).
- [ ] Click **Apply**. In one go, this:
  - **A-level Maths, each paper as its own subject**, all four in G11 and G12: AS Mathematics (Pure Mathematics 1), AS Mathematics (Probability and Statistics 1), A2 Mathematics (Pure Mathematics 3), A2 Mathematics (Probability and Statistics 1). Michael's classes move with their marks:
    - G11 "AS Maths (Pure Math 1 and Mechanics)" becomes AS Mathematics (Pure Mathematics 1);
    - his two G12 "AS Maths" classes (Pure Mathematics 3, Probability and Statistics 1) become the two **A2** subjects;
    - the old AS Maths and A2 Maths names become the Pure Mathematics subjects;
  - hides A2 Business Studies in G11 and G12;
  - adds Elective (Science Investigations), type II, to G1–G5 and ZSA.
- [ ] Every row should then show **Done**. If a row says "Old name not found", the live list uses a different name: use the **Rename** button in the subjects table for that one.
- [ ] Ask Michael to check his classes after the change, then refresh. The cards should read G11 AS Mathematics (Pure Mathematics 1), G12 A2 Mathematics (Pure Mathematics 3) and G12 A2 Mathematics (Probability and Statistics 1).
- [ ] Plain G12 "Maths" is left alone. Hide it if it isn't taught.
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
