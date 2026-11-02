# source-analysis

A Claude skill that reads a folder of CSV files from any source database and produces a
five-tab Excel analysis workbook: what the tables are, how they relate, and where the
problems are.

**The client's data never leaves the consultant's computer.** A local Python script reads
the CSVs; Claude only ever reads the finished workbook, which holds table and column names,
counts, percentages and inferred types — and no data values.

Built for Salesforce migration consultants, but the discovery is target-agnostic: it works
on any CRM, ERP, SQL Server / MySQL / Postgres / Access export, or a Salesforce org export.

---

## What you get

One `.xlsx` with five tabs, read left to right:

| Tab | What's on it | When you read it |
| --- | --- | --- |
| **Findings** | Every data-quality and structural issue, ranked High / Medium / Low | First. High rows before the kickoff call. |
| **Engagement Intake** | The questions the structure can't answer — for the client | Before scoping |
| **Source Data Analysis** | One row per table: classification, row count, proposed target object, include / exclude | Your first draft of migration scope |
| **Key Map** | Every primary and foreign key found, with a confidence rating | Focus on Low-confidence and amber rows |
| **Schema** | Every column with inferred type, max length, populated %, distinct count | Reference during field mapping |

## What it works out on its own

- **Files** — delimiter and encoding per file, ragged rows, missing or duplicate headers
- **Primary keys** — candidates verified unique and non-blank across every row
- **Foreign keys** — by naming convention and, for GUIDs and Salesforce Ids, by value;
  every candidate scored against live rows, best resolver wins. Handles `007` vs `7`,
  GUID braces and case, 15- vs 18-character Salesforce Ids
- **Polymorphic keys** — `WhatId`-style columns, and `X_id` + `X_type` pairs
- **Table classification** — Entity, Child / Detail, Lookup / Picklist, Junction,
  Polymorphic Child, Audit / History, Staging / Backup, System / Config — with its reason
- **Columns** — inferred type, max length (>255 flagged as Long Text Area), line breaks,
  leading zeros, populated %, distinct values, picklist candidates

---

## Install

### Claude (desktop or web)

1. Download the ZIP from the [Releases page](../../releases) — it is already shaped for
   upload, with `source-analysis/` at its root.
2. In Claude, open **Settings → Skills → Upload skill** and pick the ZIP.
3. Switch the skill on.

Check it took: in a new chat, ask what the source-analysis skill does.

> Downloading **Code → Download ZIP** from the green button instead gives you a
> `source-analysis-main/` wrapper, which the uploader rejects. Rename the inner folder to
> `source-analysis` and re-zip, or just use a release ZIP.

### Claude Code

```
/plugin install source-analysis --marketplace elisesfdc/source-analysis
```

On Claude Code 2.1.275 or later that's one step. On older versions, add the marketplace
first, then install:

```
/plugin marketplace add elisesfdc/source-analysis
/plugin install source-analysis@source-analysis-marketplace
```

Or clone it straight into your skills folder and skip plugins altogether:

```bash
git clone https://github.com/elisesfdc/source-analysis.git ~/.claude/skills/source-analysis
```

Start a new session and run `/skills` to confirm it loaded. Use `.claude/skills/` inside a
repo instead if you want it shared with whoever clones that repo. Skills installed in
Claude Code are separate from skills uploaded to Claude — install in both if you use both.

### Requirements

- Python 3 on the machine holding the extract
- `python3 -m pip install openpyxl` (one time; nothing else — standard library otherwise,
  and the script makes no network calls)

---

## Use it

Put the extract in a folder, connect that folder to Claude, and ask:

> Run a source analysis on `01_Source_Extract` for Acme Wealth. The source is their old CRM,
> Junxure. Put the workbook in `02_Analysis_Workbooks`.

Claude copies the tool next to your output folder, runs it on your machine, reads the
workbook, and walks you through the tabs.

### Or run it yourself

```bash
python3 scripts/source_analysis.py "01_Source_Extract" \
    --client "Acme Wealth" \
    --source "Junxure" \
    -o "02_Analysis_Workbooks/Acme_Wealth_Source_Data_Analysis.xlsx"
```

| Option | Does |
| --- | --- |
| `--profile profiles/<system>.json` | Apply system knowledge for a known source |
| `--target` | Target system label (default: Salesforce) |
| `--delimiter` / `--encoding` | Override the per-file auto-detection |
| `--recursive` | Include sub-folders |
| `--write-profile-template PATH` | Start a profile for a new source system |
| `--no-auto-profile` | Turn off profile auto-detection |

---

## Profiles

A profile is a JSON file of knowledge about a *source system* — table descriptions, target
objects, key overrides, intake questions. Bundled:

- `profiles/redtail.json` — RedTail CRM
- `profiles/salesforce_data_export.json` — Salesforce Data Export (auto-detected)
- `profiles/_template.json` — start your own

Profiles describe a system, never a client. Never commit a client name, a client's field
conventions, or anything from a specific engagement.

Built one for a system that isn't here? Pull requests welcome.

---

## Limits

- Foreign keys with integer values and no naming hint (`ShipVia` → `Shippers`) aren't
  found — add them as `fk_overrides` in a profile
- Composite primary keys aren't detected; reported as "no single-column primary key"
- Classification is structural. Confirm every row. Nothing here replaces the client's
  sign-off on scope.

## License

MIT — see [LICENSE](LICENSE).
