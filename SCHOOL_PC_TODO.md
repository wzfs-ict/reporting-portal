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
  - renames the G11 and G12 AS Maths and A2 Maths subjects to Beverly's combined names (Pure Math 1 or 3 & Probability and Statistics 1), and moves every existing class and mark to the new names, guests included;
  - hides A2 Business Studies in G11 and G12;
  - adds Elective (Science Investigations), type II, to G1–G5 and ZSA.
- [ ] Every row should then show **Done**. If a row says "Old name not found", the live list uses a different name: use the **Rename** button in the subjects table for that one.
- [ ] Michael's separate P&S 1 and Pure Maths 3 subjects are **not** added; the combined subjects cover them.
- [ ] Open the Review page for G11 and G12: the maths classes show the new names, with their marks.
- [ ] Ask Beverly whether the plain G12 "Maths" is still taught.
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
