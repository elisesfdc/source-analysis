---
name: source-analysis
description: "Step 1 of ANY CSV-extract to Salesforce migration, whatever the source database: scan a folder of CSV files (any CRM, ERP, SQL Server / MySQL / Postgres / Access export, or a Salesforce org export) into a Source Data Analysis workbook — discovered primary and foreign keys, table classification, data-quality findings, column types — no data values. Use when the user has an extract from an unknown system, or asks \"what's in this export\", \"analyze these CSVs\", \"map out this database\", \"source analysis\". RedTail extracts: run it with profiles/redtail.json.\n"
---

# Source Data Analysis — any database (Step 1)

The first step of a data migration: find out what is actually in the client's extract
before anyone scopes, maps or loads anything.

## Privacy rules — non-negotiable, they are the point of this skill

1. **Separate the code from the data.** A deterministic Python script reads the client's
   records locally. The assistant never reads raw client rows into context. It works only
   from the workbook the script produces: table and column names, counts, percentages,
   inferred types, max lengths, distinct-value counts. No data values.
2. **Run the script on the computer that holds the extract.** If the extract folder is on
   the user's computer (a connected or linked folder), run `source_analysis.py` in that
   computer's shell, not in a cloud workspace.
3. **Bring the tool to the data, never the data to the tool.** An installed skill's files
   live with the assistant, not on the user's computer. Copy `scripts/source_analysis.py`
   and the `profiles/` folder (keep them side by side as `scripts/` and `profiles/`; the
   script finds profiles at `../profiles`) into a `_source-analysis-tool/` folder inside
   the output folder (e.g. `02_Analysis_Workbooks/_source-analysis-tool/`), and run it
   from there.
4. **Never upload, stage, sync, attach or copy the CSV files into a cloud workspace**, and
   never open, read, preview or `head` them. The only file the assistant reads is the
   finished workbook.
5. **Do not accept client workbooks as chat attachments either.** Uploading a deliverable
   stages it on the model provider's servers, which defeats the whole pattern. Read it in
   place on the user's machine instead, and say so if they attach one.
6. **If the user pastes raw client rows**, do not echo them back. Ask for the schema or an
   aggregate instead and continue from that.
7. **Read-only on the source.** The script opens the extract read-only. Every output is a
   NEW file. Nothing in the raw dataset is renamed, edited, moved or deleted. Say so
   explicitly when a consultant asks — it is a standard and reasonable question.
8. **No shell on the user's computer?** Stop. Do not fall back to the cloud. Tell the user
   to link their computer, or give them the exact command to run in their terminal
   themselves, and wait for the workbook.
9. **Scripts stay self-contained and offline.** Standard library plus `openpyxl` only,
   single file, no network calls. Preserve this property when modifying the script.

## Procedure

1. Get the extract folder path (e.g. `01_Source_Extract`) and the client name.
2. Ask whether a source profile exists for this system (`profiles/*.json`). RedTail →
   `profiles/redtail.json`. Unknown system → run without one; the generic rules plus the
   empirical checks do the work. **Salesforce Data Export is auto-detected** (Account,
   Contact and User tables with Salesforce system columns, read from headers only) and
   `profiles/salesforce_data_export.json` is applied without `--profile`: it treats
   `000000000000000AAA` as blank in every lookup, and identifies person accounts by a
   populated `Account.PersonContactId`. Any profile with a `detect` block is auto-applied
   the same way; `--no-auto-profile` turns this off.
3. Run `scripts/source_analysis.py` **on the user's computer** (see the privacy rules):

   ```
   python3 source_analysis.py "01_Source_Extract" --client "Client Name" --source "Legacy CRM v7" \
       -o "02_Analysis_Workbooks/Client_Name_Source_Data_Analysis.xlsx" [--profile profiles/<system>.json]
   ```

   Options: `--target` (default Salesforce), `--delimiter tab|comma|semicolon|pipe` and
   `--encoding` (both auto-detected per file), `--recursive` (sub-folders),
   `--write-profile-template PATH`. One-time: `python3 -m pip install openpyxl`.
4. Read the workbook and walk the consultant through it in this order: **Findings** (High
   first), **Engagement Intake**, **Source Data Analysis**, **Key Map** (Low confidence and
   amber rows), **Schema**.
5. When the same source system will come up again, offer to start a profile from
   `--write-profile-template` and fill it with what the consultant confirmed (table
   targets, key overrides, intake questions). Profiles hold system knowledge only — never
   client data.

## What the script detects on its own

- **Files:** delimiter (comma, tab, semicolon, pipe) and encoding (UTF-8, UTF-8 BOM,
  UTF-16, Latin-1 fallback) per file; ragged rows; missing or duplicate headers; `dbo.`
  schema and `tbl` prefixes in file names.
- **Primary keys:** `id`, `<table>id`, `<Table>ID`, GUID/key columns or the first column —
  verified unique and never blank across all rows. Non-unique candidates are a High
  finding.
- **Foreign keys:** snake_case, camelCase and PascalCase names; columns named exactly like
  another table's key (`CustomerID`); self-references (`parent_id`, `ReportsTo`); user
  references (`owner_id`, `CreatedById`, `*_by`) to a users table when one exists. Every
  candidate is scored on live rows; the best-resolving target wins. Key values are
  normalised (`007` = `7`, GUID case and braces, 15- vs 18-character Salesforce Ids), and
  held in memory as hashes only.
- **Value-based discovery:** GUID and Salesforce-style Ids are matched by value even with
  no naming hint; ids spanning several tables (`WhatId`) become polymorphic keys. Integer
  ids are never assigned on values alone — overlaps are reported as Ambiguous.
- **Polymorphic pairs:** `X_id` + `X_type` / `X_table` / `X_class`, with type values
  (`Crm::Contact`, `App\Models\Contact`, `Customer`) mapped to tables.
- **Classification:** Entity, Child / Detail, Lookup / Picklist, Junction, Polymorphic
  Child, Audit / History, Staging / Backup, System / Config, Unclassified — each with its
  reason.
- **Columns:** inferred type (Integer, Decimal, Date, Date/time, Boolean, GUID, Email,
  Phone, URL, Text, Mixed), max length (> 255 → Long Text Area), line breaks, leading
  zeros, populated %, distinct-value count, picklist candidates.

## Profiles

A profile (JSON) adds system knowledge: table descriptions, target objects, default
inclusion, UI locations and reference links; key overrides and notes; type-name map; system
and entity tables; and an intake that replaces the generic one. Intake rows can re-scope
tables live: `gates_tables` / `gates_classes`, or a table's `gated_by`, make a No answer
flip those tables to Not Included. See `profiles/_template.json`.

Profiles describe a *system*, never a client. Never put a client name, a client's field
conventions, or anything from a specific engagement in one.

## Limits — tell the consultant

- Keys with no naming hint and integer values (e.g. `ShipVia` → Shippers) are not found;
  add them as `fk_overrides`.
- Composite primary keys are not detected (reported as "no single-column primary key").
- Classification is structural; confirm every row. Nothing here replaces the client's
  sign-off.

## What the workbook holds

Names of tables and columns, counts, percentages, inferred types, max lengths,
distinct-value counts, and the class-name values of polymorphic `*_type` columns — no data
values. Never ask the user to paste raw CSV contents into chat.
