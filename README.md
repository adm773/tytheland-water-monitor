# Tytheland Water Monitor

Google Apps Script bound to the **Water Monitor — Tytheland Master Meter** Google Sheet.
Every morning it reads the R S Hill master meter's Efento pulse logger (282C0242B88D, channel 1),
grades the night for leaks (CLEAR / WATCH / LEAK / NO DATA / CHECK), updates the Dashboard tab,
and emails a summary.

## Files

| File | Purpose |
| --- | --- |
| `Code.gs` | The whole script |
| `appsscript.json` | Project manifest (UK time zone, V8 runtime) |
| `.clasp.json` | Links this folder to Apps Script project `1NUwUig258wkLP-S6LUXHVNDZiv4o8GxTXvPEaFqJubwNT0YxxL331G-F` |
| `.claspignore` | Only `Code.gs` and `appsscript.json` are pushed |

## Deploying with clasp

```bash
npm install -g @google/clasp
clasp login              # once per machine; never commit ~/.clasprc.json
clasp pull               # first time only: check nothing in the live project would be lost
clasp push
```

`clasp push` replaces every file in the Apps Script project with the files here.

After pushing, reload the Sheet and run **Water Monitor → Refresh dashboard**.
If a chart looks wrong, run **Water Monitor → Check dashboard charts**: the Diagnostics tab then
lists each chart, the range it reads and how many numbers that range holds.

## Setup in a new Sheet

Use the **Water Monitor** menu in order: Set up credentials → Find measurement point →
Diagnose API → Backfill nights → Install daily trigger.
