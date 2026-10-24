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
| `.github/workflows/deploy.yml` | Deploys to Apps Script on every push to `main` |

## Automatic deployment

Every push to `main` that changes `Code.gs`, `appsscript.json` or the clasp config runs the
**Deploy to Apps Script** workflow (`.github/workflows/deploy.yml`). It:

1. checks `Code.gs` for syntax errors,
2. pulls the live project and stops if it holds edits made in the Apps Script editor that are not
   in GitHub (so nothing is lost silently),
3. pushes the code with clasp, and
4. moves the `deployed` tag to the commit that is now live (before the first deploy, the
   repo's first commit stands for what is live).

You can also run it by hand from **Actions → Deploy to Apps Script → Run workflow**. Tick
*overwrite live edits* only if you mean to discard changes made directly in the editor.

### One-time setup

1. Turn on **Google Apps Script API** at <https://script.google.com/home/usersettings>.
2. On your own computer: `npm install -g @google/clasp` then `clasp login`, and sign in with the
   Google account that owns the Sheet.
3. Copy the whole contents of `~/.clasprc.json` into a repository secret named `CLASPRC_JSON`
   (**Settings → Secrets and variables → Actions → New repository secret**). Never commit that file.
4. Run the workflow once by hand to make the first deploy.

The secret lets the workflow act on your Google account, so keep this repository private. If
deploys start failing with an authentication error (for example after a password change), run
`clasp login` again and replace the secret.

### Deploying by hand

```bash
clasp pull    # check nothing in the live project would be lost
clasp push
```

After a deploy, reload the Sheet and run **Water Monitor → Refresh dashboard**.
If a chart looks wrong, run **Water Monitor → Check dashboard charts**: the Diagnostics tab then
lists each chart, the range it reads and how many numbers that range holds.

## Setup in a new Sheet

Use the **Water Monitor** menu in order: Set up credentials → Find measurement point →
Diagnose API → Backfill nights → Install daily trigger.
