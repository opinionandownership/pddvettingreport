# PDD Handover Letter

Website: https://opinionandownership.github.io/pddvettingreport/

Fills in the PDD handover letter on the M. Raja / Opinion & Ownership letterhead and downloads it as PDF or Word.

## Reference numbers

Format: `BANK INITIALS / USER NO / RUNNING NO / Mon / Year`, e.g. `TFL/32/1/Nov/2026`.

The running number comes from a shared counter in a Google Sheet (one count per bank), so every
download gets its own number and no two people can receive the same one. If the counter can't be
reached, the site refuses to download rather than risk a duplicate.

## Connecting the counter (one-time)

1. Create a new Google Sheet (e.g. "PDD Register").
2. Extensions → Apps Script. Replace the code with `apps-script/Code.gs`, then Save.
3. Deploy → New deployment → type **Web app**. Execute as: **Me**. Who has access: **Anyone**. Deploy and allow access.
4. Copy the web app URL (ends in `/exec`) and paste it into `index.html` at `const COUNTER_URL='';`.

The sheet gets two tabs: **Counters** (edit "Next number" to set where a bank starts) and **Register** (every number issued).
Only you can edit the sheet and this repository.
