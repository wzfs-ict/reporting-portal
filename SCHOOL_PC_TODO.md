# School PC — what to do

## Done from the cloud (no work needed on this PC)

- PTC sheets print fix (18 Nov): long teacher remarks no longer spill off the
  printed page. The app now shrinks an over-long sheet automatically when you
  print, and if it is still too long it uses a second page instead of cutting
  off words. Nothing for you to do — just print as usual.

## Still waiting (needs this PC, D:\Github\reporting-portal)

1. Put the two patch files into D:\Github\reporting-portal (the folder itself,
   next to app.js):
   - carry-classes-midterm-to-final.patch
   - certificates-v2-assets.patch
2. Open PowerShell in that folder and run:
     git am carry-classes-midterm-to-final.patch
     git am certificates-v2-assets.patch
3. Open GitHub Desktop and Push origin.

Notes:
- Use GitHub Desktop for pull/push (its network works). PowerShell git could
  not reach github.com on 18 Nov.
- Do NOT commit the .vs files Visual Studio shows. They are now ignored.

## Midterm → Final switch (do this when midterm is fully over)
- [ ] Take a **Snapshot** first (Setup → Health check card).
- [ ] Setup → card 3 → set **Reporting period** to **Final** → Save term and people.
- [ ] As yourself (Glen, you teach classes): open **My classes**. The "Carry your classes over" card lists your midterm classes, all ticked. Click **Copy selected classes to Final**. Classes appear with blank marks, "Not started".
- [ ] Check your **Midterm** marks are untouched: switch the period back to Midterm (Setup → card 3), open a class, see the marks. Switch back to Final again.
- [ ] **View as** a teacher who has classes (e.g. Gregg): the carry-over card shows for them too. In preview a copy click is blocked and saves nothing.
- [ ] Tell staff: "Open the portal, open My classes, click Copy my classes once. Then enter final marks."
- [ ] Same flow in spring: period Midterm + semester Spring offers to copy from Fall Final. Fall Midterm (new school year) starts with no copy — classes are added fresh.
