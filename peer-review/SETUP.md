# Tool D · Peer Review — setup

A standalone Apps Script project, separate from the student workbook on purpose.
Every other tool is copied once per student; this one is shared. Ten students
write into one spreadsheet and you compile what they wrote about each other, so
a per-student copy would isolate exactly the data that has to pool.

Students get the web app URL. They never open the spreadsheet.

## The spreadsheet

Already created in your Drive:

**DDII · D Peer Review Collection**
`https://docs.google.com/spreadsheets/d/1Oj9LcWAh0_RvuUmz4RIpq80CEAYuUfcROHsU_Yin_q8/edit`

It is empty until you run `setup()` in step 4, which builds all four tabs.

## 1 · Attach a script to it

Open the spreadsheet → **Extensions → Apps Script**. This creates a
container-bound script, which is what lets it read and write the sheet without
any credentials.

## 2 · Get the Script ID

In the script editor: **Project Settings** (the gear) → copy the **Script ID**.

## 3 · Point clasp at it

From the repository root:

```
cd peer-review
copy .clasp.json.example .clasp.json      # Windows
cp   .clasp.json.example .clasp.json      # macOS / Linux
```

Open `.clasp.json` and paste the Script ID in place of the placeholder. Then:

```
clasp push
```

`.clasp.json` is gitignored — the ID stays on your machine.

The main workbook still pushes from the repository root. Two projects, two
`.clasp.json` files, one clone. Run `clasp push` from whichever folder you mean.

## 4 · Build the sheets

Back in the script editor, choose `setup` from the function dropdown and press
**Run**. Accept the permission prompt.

It creates **How to use**, **Responses**, **Package** and **Participation**,
with every formula, the student dropdown, column widths and frozen headers.
It is safe to run again at any time: it repairs structure and formulas and
never touches submitted reviews.

## 5 · Deploy

**Deploy → New deployment → Web app.**

| Setting | Value |
| --- | --- |
| Execute as | **Me** |
| Who has access | **Anyone with the link** (or your Workspace domain) |

Execute-as-me is what lets students write into a spreadsheet they cannot open.

Copy the web app URL and give that to the class. Not the spreadsheet link.

> Re-deploying after a `clasp push`: **Deploy → Manage deployments →** pencil →
> Version **New version → Deploy**. Without a new version students keep getting
> the old code.

## Changing the class list

`STUDENTS` in `Code.gs` is the list. Column A of **Participation** mirrors it and
feeds the Package dropdown. Change both together — edit `Code.gs`, `clasp push`,
then run `setup()` again — and re-deploy.

## What must not change

| Thing | Why |
| --- | --- |
| `KEY_QUESTION` — `Direction, in my words` | Appears exactly once per round, so one row carrying it is one completed review. Duplicate detection, the progress line, Package's received count and every Participation column key on it. |
| The three round labels in `ROUNDS` | Participation matches them character for character with `COUNTIFS`. A changed label silently zeroes that column. |
| Responses column order | `Round, Date, Presenting, Reviewer, Question, Response, Index, Submitted`. Package and Participation address columns by letter. |
| Column G | A sheet-owned formula, filled to row 3000. The script writes A–F and H and never touches G — writing anything there, `""` included, replaces the formula and drops that review out of Package. |
