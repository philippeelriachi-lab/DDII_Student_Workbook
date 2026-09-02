# Getting this code onto the DDII Student Workbook

This repo is the Apps Script project (`src/`) for the master copy of the DDII
Student Workbook Google Sheet. It has to be pushed into the Sheet's bound
Apps Script project with `clasp` — a Drive connector cannot reach a bound
script project, and the Sheet itself has no code-upload UI.

Do this once, on your own machine (not in a cloud/remote session — `clasp
login` needs a browser you control).

## 1. Install clasp

```bash
npm install -g @google/clasp
```

Check you have a reasonably recent Node (`node -v`) first — clasp needs a
current LTS release.

## 2. Enable the Apps Script API (once, for your Google account)

Visit **https://script.google.com/home/usersettings** and turn on
"Google Apps Script API". Without this, `clasp login` and `clasp push` fail
with a permissions error.

## 3. Log in

```bash
clasp login
```

This opens a browser window for Google OAuth. Log in as the account that
owns (or has edit access to) the DDII Student Workbook. You'll see an
"unverified app" warning — that's expected for a personal Apps Script
project; click **Advanced → Go to (project name) (unsafe)** to continue.
Don't try to get rid of this screen — it's normal for a script that isn't
published to the Workspace Marketplace, and it will reappear the first time
each new copy of the workbook runs its own script too. Just document it for
students so it doesn't alarm them.

## 4. Find the Script ID

Open the **DDII_Student_Workbook** Google Sheet → **Extensions → Apps
Script**. This creates the Sheet's bound script project if it doesn't exist
yet. Then in the Apps Script editor: **Project Settings** (the gear icon in
the left sidebar) → copy the **Script ID** near the top.

## 5. Point this repo at that script

This repo already has the code in `src/`. Don't run `clasp clone` — that
would try to pull the (empty) remote project down and could clobber what's
here. Instead, wire the existing files up to the real project:

```bash
cp .clasp.json.example .clasp.json
```

Edit `.clasp.json` and paste the real Script ID in place of
`PASTE_YOUR_SCRIPT_ID_HERE`. It should look like:

```json
{
  "scriptId": "1AbC...xyz",
  "rootDir": "src"
}
```

`.clasp.json` is gitignored on purpose — it's local wiring, not something to
commit (it's specific to your machine and, if it ever held credentials,
those shouldn't end up in the repo either).

## 6. Push

From the repo root:

```bash
clasp push
```

If it warns that the manifest (`appsscript.json`) is about to be
overwritten, confirm — the remote one is just the default empty manifest
Apps Script created in step 4.

## 7. Verify

1. Reload the DDII_Student_Workbook spreadsheet tab. A **Tools** menu should
   appear within a few seconds (it's created by the `onOpen()` trigger in
   `Code.gs`).
2. Run **Tools → Diagnose**. On a clean workbook it should report the tabs,
   24 lookup lists, and 0 rows without an id (acceptance criterion #1 in the
   build brief).
3. Open a couple of the numbered tools from the menu and confirm they load
   without errors, and that closing without editing anything leaves the
   sheet unchanged (acceptance criterion #2).

## 8. Before distribution day

Everything above updates the **master** template only. Once you distribute
via Google Classroom ("make a copy for each student"), each copy gets its
own independent bound script — a later `clasp push` here will **not** reach
copies already handed out. Get the script right in the master first, then
distribute.

## Notes

- `clasp` syntax and flags do shift between versions — if a command above
  doesn't match what your installed `clasp` does, check `clasp --help` or
  the current docs at https://github.com/google/clasp before assuming
  something is broken here.
- The loop after this is simple: edit files in `src/`, `git commit`,
  `clasp push`, reload the sheet tab, test.
