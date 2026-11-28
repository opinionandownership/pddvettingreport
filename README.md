# PDD Handover Letter

Website: https://opinionandownership.github.io/pddvettingreport/

Fills in the PDD handover letter on the M. Raja / Opinion & Ownership letterhead and downloads it as PDF or Word.

## Reference numbers

Format: `BANK INITIALS / USER NO / RUNNING NO / Mon / Year`, e.g. `TFL/32/1/Nov/2026`.

The running number comes from a shared counter in a Google Sheet (one count per bank), so every
download gets its own number and no two people can receive the same one. If the counter can't be
reached, the site refuses to download rather than risk a duplicate.

## Connecting the counter (one-time)

1. Open Apps Script (from your Google Sheet via Extensions → Apps Script, or at script.google.com).
2. Replace all the code with `apps-script/Code.gs`, then Save.
3. Pick the function **setup** in the toolbar and click **Run**. Allow the permissions. The log shows which sheet is used
   (if the script isn't attached to a sheet, it creates "PDD Register" in your Drive).
4. Deploy → Manage deployments → Edit (pencil) → Version: **New version**. Execute as: **Me**. Who has access: **Anyone**. Deploy.
5. The web app URL goes in `index.html` at `const COUNTER_URL=...`.

The site shows "Reference counter connected" in green when it works.
The sheet has three tabs: **Counters** (edit "Next number" to set where a bank starts), **Register** (every number issued) and
**Users** (each person's user number; if a number is already taken, the first two letters stay and the third is swapped for the next letter of the name until a free one is found).
