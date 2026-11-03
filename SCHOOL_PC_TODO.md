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
  - **Maths (Bijily, 3 Nov):** three subjects in G11 and G12: **AS Mathematics (Pure Math 1 & Probability and Statistics 1)**, **A2 Mathematics (Pure Math 3 & Probability and Statistics 1)**, and **Mathematics** (for the bobaes not taking AS or A2), plus **Pure Mathematics 1** on its own (Beverly). Michael's classes move with their marks:
    - G11 "AS Maths (Pure Math 1 and Mechanics)" becomes AS Mathematics (Pure Math 1 & Probability and Statistics 1);
    - his two G12 paper classes (Pure Mathematics 3, Probability and Statistics 1) merge into **A2 Mathematics (Pure Math 3 & Probability and Statistics 1)**. Of each bobae's two rows, the one with marks is kept and the empty one is deleted (it stays in the recycle bin);
    - old A2 Maths becomes A2 Mathematics (Pure Math 3 & Probability and Statistics 1); G12 "Maths" becomes Mathematics;
  - hides A2 Business Studies in G11 and G12;
  - adds Elective (Science Investigations), type II, to G1–G5 and ZSA.
- [ ] Every row should then show **Done**. If a row says "Old name not found", the live list uses a different name: use the **Rename** button in the subjects table for that one.
- [ ] If Apply reports **clashes**, a G12 bobae has marks on *both* papers. Those rows stay on Michael's old "AS Maths (Probability and Statistics 1)" card. Michael enters the combined mark on the A2 card, then deletes the old card.
- [ ] Ask Michael to refresh and check. He should see G11 AS Mathematics (Pure Math 1 & Probability and Statistics 1) and G12 A2 Mathematics (Pure Math 3 & Probability and Statistics 1), and one mark per bobae per level.
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
