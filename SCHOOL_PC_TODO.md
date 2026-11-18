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
