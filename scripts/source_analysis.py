#!/usr/bin/env python3
"""
MigrationOS Source Data Analysis — generic edition.
Step 1 of any CSV-extract → Salesforce migration, for ANY source database.

Point it at a folder of CSV files exported from any system (a CRM, an ERP, a homegrown SQL
Server / MySQL / Postgres / Access database, a Salesforce org …). It discovers the structure
on its own — primary keys, foreign keys, polymorphic keys, user references, table types,
column data types — and writes a branded Excel workbook:

  Cover                 — engagement header, live roll-ups, classification legend, how to use
  Engagement Intake     — decisions to settle with the client before analysis (some answers
                          re-scope tables automatically through live formulas)
  Findings              — data-quality issues worth knowing about on day one, ranked by severity
  Source Data Analysis  — one row per table: classification, primary key, counts, both-direction
                          key tracing, defaults, review and sign-off columns
  Key Map               — one row per foreign key: target, how it was found, confidence, and the
                          EMPIRICAL resolution % on live rows
  Schema                — one row per column: role, inferred type, max length, populated %,
                          distinct-value count

SOURCE PROFILES (optional): system-specific knowledge — table descriptions and target objects,
known key mappings, intake questions — lives in a JSON profile passed with --profile. Without
one, the script runs on generic conventions plus empirical checks. The RedTail playbook ships as
profiles/redtail.json. Write a blank one with --write-profile-template.

PRIVACY: the workbook holds structure and aggregates only — table and column names, counts,
percentages, inferred types, maximum lengths, distinct-value COUNTS, and the class-name values
of polymorphic *_type columns. No data values, no names, no identifiers. Key values are held in
memory as hashes only, to test joins. The script runs locally and makes no network connections.

USAGE:
    python3 source_analysis.py "/path/to/extract folder" --client "Acme" --source "Legacy CRM"
    python3 source_analysis.py "/path/to/extract" --client "Acme Wealth" --profile profiles/redtail.json
    python3 source_analysis.py --write-profile-template my_system.json

REQUIREMENT (one-time): python3 -m pip install openpyxl
"""

import argparse
import csv
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import date

VERSION = "v0.4 generic"

# ===========================================================================
# GENERIC CONVENTIONS — every name below is compared in normalized form
# (lower case, letters and digits only: "Deleted_At" == "deletedat" == "DeletedAt").
# A profile can extend any of these.
# ===========================================================================

DELETION_TIMESTAMP = ["deletedat", "deletedon", "datedeleted", "deleteddate", "removedat",
                      "deletiondate", "deletedtime", "deletedtimestamp"]
DELETION_BOOLEAN = ["isdeleted", "deleted", "deletedflag", "delflag", "isremoved", "removed",
                    "recordisdeleted", "softdeleted"]
NULL_TOKENS = {"", "null", "none", "nil", "nan", "\\n", "#n/a", "n/a"}
# Present but carrying no meaningful value (excluded from Populated %; still typed).
EMPTYISH = {"0", "false", "f", "no", "n", "0.0", "0.00"}

STRONG_KEY_SUFFIXES = {"id", "key", "fk", "guid", "uuid"}
WEAK_KEY_SUFFIXES = {"no", "num", "nbr", "code", "ref", "number"}   # only count when a table name matches

USER_STEMS = {"owner", "assignee", "assignedto", "creator", "author", "modifier", "user",
              "createdby", "modifiedby", "updatedby", "lastmodifiedby", "addedby", "enteredby"}
USER_TABLE_NAMES = {"user", "users", "sysuser", "appuser", "employee", "staff", "login",
                    "member", "aspnetuser", "useraccount", "systemuser", "adminuser"}

NOT_FK = {"taxid", "ssnid", "federalid", "externalid", "legacyid", "globalid", "tempid",
          "npi", "ein", "nationalid", "passportid", "driverslicenseid", "sessionid",
          "transactionid", "requestid", "correlationid", "trackingid"}

GENERIC_STEMS = {"type", "status", "code", "category", "stage", "kind", "role", "detail"}

SELF_STEMS = {"parent", "manager", "reportsto", "master", "original", "merged", "mergedinto",
              "duplicateof", "previous", "prior", "next", "root", "replacedby", "supersededby"}

POLY_TYPE_SUFFIXES = ["type", "table", "class", "model", "object", "entity", "objecttype"]

AUDIT_COLUMNS = {"createdat", "updatedat", "createdon", "updatedon", "modifiedat", "modifiedon",
                 "createddate", "lastmodifieddate", "systemmodstamp", "datecreated", "datemodified",
                 "lastupdated", "lastupdate", "rowversion", "timestamp", "insertedat", "insertdate",
                 "updatedate", "createdatetime", "modifieddatetime", "tempid", "globalid",
                 "archived", "archivedat", "recycledelete", "lastviewdate", "lastreferenceddate",
                 "etag", "version", "rowguid", "lockversion"}

LABEL_TOKENS = {"name", "description", "desc", "label", "title", "value", "code", "text",
                "display", "displayname", "abbreviation", "abbr", "caption", "meaning"}

AUDIT_TABLE_TOKENS = {"log", "logs", "audit", "audits", "history", "histories", "hist",
                      "changelog", "journal", "trail", "changes", "versions"}
STAGING_TABLE_TOKENS = {"bak", "backup", "backups", "old", "copy", "temp", "tmp", "staging",
                        "scratch", "test", "deprecated", "unused", "archive", "orig"}

TABLE_PREFIXES = {"tbl", "tb", "t", "dim", "fact", "ref", "lkp", "lu", "lk", "lookup"}

DISTINCT_CAP = 5000          # distinct-value counting stops here ("5,000+")
TYPE_SAMPLE = 20000          # values per column used for type inference
PICKLIST_MAX = 30            # low-cardinality text → picklist candidate
NOT_PICKLIST_TOKENS = {"name", "first", "last", "middle", "city", "street", "address", "line", "email", "phone",
                       "note", "notes", "comment", "comments", "description", "subject", "body", "memo", "url"}
FILE_EXTENSIONS = (".csv", ".tsv", ".txt", ".psv")

INCLUDED_OPTIONS = ["To Be Reviewed", "Included", "Not Included", "Not Applicable"]
ANALYSIS_OPTIONS = ["To Be Reviewed", "In progress", "Completed"]
CLIENT_REVIEW_OPTIONS = ["Not started", "In review", "Approved", "Not Applicable"]
INTAKE_ANSWER_OPTIONS = ["Yes", "No", "Undecided", "N/A"]
INTAKE_STATUS_OPTIONS = ["Open", "Decided", "N/A"]

# Generic Engagement Intake (Salesforce target, any source). A profile's "intake" replaces it.
# gates_classes / gates_tables: answer No → Not Included, Yes → Included, else To Be Reviewed.
GENERIC_INTAKE = [
    dict(topic="Source system & extract scope",
         question="What system and version produced this extract? Is it the full database or a partial / filtered export?",
         why="Tells you whether a table the Key Map can't find is missing from the export or never existed. A partial export changes every count in this workbook.",
         default="Confirm with the client's IT contact or vendor", drives="all tables"),
    dict(topic="Missing tables",
         question="Do the keys flagged 'Unresolved — no matching table' on the Key Map point at tables that exist in the source but were left out of the export?",
         why="An unresolved key usually means a lookup or parent table was not exported. Request it now; it is cheap on day one and expensive after mapping.",
         default="Request a re-export of any missing table", drives="Key Map — unresolved rows"),
    dict(topic="Existing data in the target org",
         question="Does the target Salesforce org already hold records for these entities (from an integration, a prior migration or manual entry)?",
         why="If yes, source records must be MATCHED to existing records instead of inserted, or the migration creates duplicates.",
         default="Ask — common when an integration already loads records", drives="Entity tables"),
    dict(topic="Match key",
         question="Which key identifies an existing Salesforce record for a source record (External ID, Tax ID, email, account number)?",
         why="Drives the local match step. The match runs on the consultant's machine; Claude sees counts only.",
         default="An External ID or a stable natural key", drives="Entity tables"),
    dict(topic="Target data model",
         question="Which Salesforce data model is the target: standard Sales/Service objects, an industry cloud (FSC, Nonprofit, Health), custom objects, or a mix?",
         why="Every target object and field API name downstream depends on it. Detect it from the org where possible rather than asking.",
         default="Detect from the connected org", drives="all migrated tables",
         options=["Standard objects", "Industry cloud", "Custom objects", "Mixed", "Undecided"]),
    dict(topic="Record scope & history cutoff",
         question="Migrate all records, or only active records / records after a cutoff date?",
         why="A cutoff can remove most of the volume in activity, note and transaction tables. Decide before mapping, not at load time.",
         default="All active records; history after an agreed cutoff", drives="Entity and Child / Detail tables"),
    dict(topic="Deleted records",
         question="Exclude soft-deleted rows (the Deletion Column on the Source Data Analysis tab) everywhere?",
         why="Standard rule. The Non-Deleted counts in this workbook already apply it.",
         default="Yes — always", drives="all tables with a deletion column"),
    dict(topic="Audit / history tables",
         question="Migrate audit, log and history tables?",
         why="Salesforce field history cannot be loaded, and these tables are usually large. Answering No re-scopes every Audit / History table automatically.",
         default="No — archive outside Salesforce", drives="tables classified Audit / History",
         gates_classes=["Audit / History"]),
    dict(topic="Staging / backup copies",
         question="Are the tables that look like backups, copies or staging tables safe to ignore?",
         why="Copies such as *_bak, *_old or *_2019 duplicate real tables. Answering No (do not migrate) re-scopes them automatically.",
         default="No — do not migrate; confirm none is the live table", drives="tables classified Staging / Backup",
         gates_classes=["Staging / Backup"]),
    dict(topic="System / configuration tables",
         question="Should any application configuration or system table be migrated?",
         why="Permissions, settings and integration plumbing are rebuilt in Salesforce, not loaded.",
         default="No", drives="tables classified System / Config", gates_classes=["System / Config"]),
    dict(topic="Users & record ownership",
         question="Map source users to Salesforce users for owner, created-by and assigned-to fields? Is there a users list if the extract has none?",
         why="Every User Reference row on the Key Map needs a user mapping. Inactive users need a fallback owner.",
         default="Map active users; fallback owner for the rest", drives="all user-reference columns"),
    dict(topic="Lookup / code tables",
         question="Convert lookup tables to picklists, or keep them as related objects? Who approves value cleanup?",
         why="Most lookup tables become picklist values; a few with their own attributes need an object.",
         default="Picklists, with a value map reviewed by the client", drives="tables classified Lookup / Picklist"),
    dict(topic="Many-to-many relationships",
         question="How should junction tables be represented (standard relationship objects, a custom junction object, or dropped)?",
         why="Each junction becomes a relationship object such as AccountContactRelation, CampaignMember or a custom junction.",
         default="Map to the nearest standard relationship object", drives="tables classified Junction"),
    dict(topic="Files & attachments",
         question="Where do files live: inside the database (blobs, file paths) or in a separate store? Has that store been exported?",
         why="A table of file metadata is not the files. Without the file export, nothing can be attached.",
         default="Request the file export on day one", drives="document / attachment tables"),
    dict(topic="Notes, activities & free text",
         question="Where do notes, emails and activity history land: Tasks/Events, Notes, or a custom object?",
         why="Changes the load path entirely. Long text over 255 characters needs Long Text Area fields (see Schema → Max Length).",
         default="Tasks for activities, Notes for notes", drives="note and activity tables"),
    dict(topic="Custom fields",
         question="Which source-only columns deserve a custom field in Salesforce?",
         why="Schema → % Populated shows which columns carry real data; columns that are rarely populated rarely justify a field.",
         default="Create only where populated % is material", drives="all migrated tables"),
    dict(topic="Legacy IDs",
         question="Store every source primary key in an External ID field on its target object?",
         why="Legacy IDs make upserts, reconciliation and later loads possible. They are also how child records find their parents during load.",
         default="Yes — always", drives="all migrated tables"),
]

# ===========================================================================
# Brand (MigrationOS — see references/brand.md in the field-mapping-matrix skill)
# ===========================================================================
DEEP_TEAL, TEAL_MID, LIGHT_TEAL, TEAL_TINT, TEAL_INK = "0F6E56", "1D9E75", "9FE1CB", "E9F6F0", "0F4A3A"
AMBER, INK, MUTED, PAPER, ZEBRA, HAIRLINE, WHITE = "9A5511", "2C2C2A", "6B6A64", "FCFAF5", "FAF7F0", "E6E3DA", "FFFFFF"
AMBER_TINT, RED, RED_TINT, YELLOW = "FBF1E2", "A32D2D", "F9E3E1", "FFF9DB"
DISPLAY_FONT, BODY_FONT, LINK = "Fraunces", "Inter", "smcmigrations.com"

CLASS_COLORS = {
    "Entity": (DEEP_TEAL, WHITE),
    "Child / Detail": (TEAL_TINT, TEAL_INK),
    "Lookup / Picklist": ("DDEFE7", TEAL_INK),
    "Junction (many-to-many)": ("CFE9DD", TEAL_INK),
    "Polymorphic Child": ("B9E0D0", TEAL_INK),
    "Audit / History": ("EDE9E0", MUTED),
    "Staging / Backup": ("EDE9E0", MUTED),
    "System / Config": (HAIRLINE, MUTED),
    "Unclassified": (AMBER_TINT, AMBER),
}

CLASS_TARGET_DEFAULTS = {
    "Entity": "TBD — standard or custom object",
    "Child / Detail": "Fields on the parent, or a related child object",
    "Lookup / Picklist": "Picklist values (or a lookup object if it has its own attributes)",
    "Junction (many-to-many)": "Relationship / junction object (e.g. AccountContactRelation, CampaignMember)",
    "Polymorphic Child": "Split by parent type; then fields or child records on each parent",
    "Audit / History": "Usually not migrated (field history cannot be loaded)",
    "Staging / Backup": "Not migrated — confirm it is not the live table",
    "System / Config": "Not migrated — rebuilt as Salesforce configuration",
    "Unclassified": "TBD — inspect the Schema tab",
}


# ===========================================================================
# Name handling
# ===========================================================================

def norm(s):
    return re.sub(r"[^a-z0-9]", "", s.lower())


def tokenize(s):
    """'CustomerID' → [customer, id]; 'createdById' → [created, by, id]; 'contact_family_id' → [...]."""
    s = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", s)
    s = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", s)
    s = re.sub(r"([A-Za-z])([0-9])", r"\1 \2", s)
    return [t for t in re.split(r"[^A-Za-z0-9]+", s.lower()) if t]


def singular(name):
    if len(name) <= 3:
        return name
    if name.endswith("ies"):
        return name[:-3] + "y"
    for end in ("sses", "shes", "ches", "xes", "ses", "zes"):
        if name.endswith(end):
            return name[:-2]
    if name.endswith("ss") or name.endswith("us") or name.endswith("is"):
        return name
    if name.endswith("s"):
        return name[:-1]
    return name


PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "profiles")


def _read_header(path):
    """First line only — column names, never data."""
    for enc in ("utf-8-sig", "latin-1"):
        try:
            with open(path, encoding=enc, newline="") as f:
                return [norm(h) for h in next(csv.reader(f), [])]
        except (UnicodeDecodeError, StopIteration):
            continue
        except OSError:
            return []
    return []


def detect_profile(paths):
    """Pick a profile whose 'detect' block matches the extract (file names + header columns)."""
    by_name = {table_base_name(n).lower(): p for n, p in paths}
    try:
        files = sorted(f for f in os.listdir(PROFILE_DIR) if f.endswith(".json") and not f.startswith("_"))
    except OSError:
        return None, None
    for fn in files:
        fp = os.path.join(PROFILE_DIR, fn)
        try:
            with open(fp, encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError):
            continue
        det = data.get("detect")
        if not det:
            continue
        need = [t.lower() for t in det.get("tables", [])]
        if not need or any(t not in by_name for t in need):
            continue
        ok = True
        for tbl, cols in det.get("columns", {}).items():
            hdr = set(_read_header(by_name.get(tbl.lower(), ""))) if tbl.lower() in by_name else set()
            if not all(norm(c) in hdr for c in cols):
                ok = False
                break
        if ok:
            return data, fp
    return None, None


def table_base_name(filename):
    stem = os.path.splitext(filename)[0]
    # strip schema prefixes like dbo.Customers / public.orders
    if "." in stem:
        stem = stem.split(".")[-1]
    return stem


def table_aliases(name):
    toks = tokenize(name)
    n = "".join(toks)
    forms = {n, singular(n)}
    if len(toks) > 1 and toks[0] in TABLE_PREFIXES:
        rest = "".join(toks[1:])
        forms |= {rest, singular(rest)}
    return {f for f in forms if f}


# ===========================================================================
# Value handling (values are never written out — only typed, counted and hashed)
# ===========================================================================

RE_INT = re.compile(r"[-+]?\d+")
RE_DEC = re.compile(r"[-+]?\$?(\d{1,3}(,\d{3})+|\d+)?\.\d+|[-+]?\$?\d{1,3}(,\d{3})+")
RE_DATE = re.compile(r"(\d{4}-\d{1,2}-\d{1,2}|\d{1,2}/\d{1,2}/\d{2,4}|\d{1,2}\.\d{1,2}\.\d{4}|\d{4}/\d{1,2}/\d{1,2})")
RE_TIME = re.compile(r"\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?\s*([AaPp][Mm])?")
RE_GUID = re.compile(r"\{?[0-9a-fA-F]{8}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{12}\}?")
RE_EMAIL = re.compile(r"[^@\s]+@[^@\s]+\.[A-Za-z]{2,}")
RE_URL = re.compile(r"(https?://|www\.)\S+", re.I)
RE_PHONE = re.compile(r"\+?[\d\s().\-]{7,22}((x|ext\.?)\s*\d{1,6})?", re.I)
RE_TOKEN = re.compile(r"[A-Za-z0-9_\-]{6,40}")
RE_SFID18 = re.compile(r"[a-zA-Z0-9]{15}[A-Z0-5]{3}")
BOOL_WORDS = {"true", "false", "yes", "no", "y", "n", "t", "f"}


def value_kind(v):
    lv = v.lower()
    if lv in BOOL_WORDS:
        return "bool"
    if RE_INT.fullmatch(v):
        return "int"
    if RE_DEC.fullmatch(v):
        return "dec"
    m = RE_DATE.match(v)
    if m:
        rest = v[m.end():].strip(" T")
        if not rest:
            return "date"
        if RE_TIME.match(rest):
            return "datetime"
    if RE_TIME.fullmatch(v):
        return "time"
    if RE_GUID.fullmatch(v):
        return "guid"
    if RE_EMAIL.fullmatch(v):
        return "email"
    if RE_URL.match(v):
        return "url"
    if RE_PHONE.fullmatch(v) and 7 <= sum(ch.isdigit() for ch in v) <= 15 and not RE_INT.fullmatch(v):
        return "phone"
    if RE_TOKEN.fullmatch(v) and any(c.isdigit() for c in v) and any(c.isalpha() for c in v):
        return "token"
    return "text"


def is_token_id(v):
    """A non-integer identifier (GUID, Salesforce Id, alphanumeric key) — safe to match by value."""
    if RE_GUID.fullmatch(v):
        return True
    return bool(RE_TOKEN.fullmatch(v)) and any(c.isdigit() for c in v) and any(c.isalpha() for c in v) and len(v) >= 8


def norm_key(v):
    """Canonical form for join tests: '007' == '7' == '7.0'; GUIDs case- and brace-insensitive;
    18-character Salesforce Ids == their 15-character form."""
    m = re.fullmatch(r"0*(\d+)(?:\.0+)?", v)
    if m:
        return m.group(1) or "0"
    if RE_GUID.fullmatch(v):
        return v.strip("{}").replace("-", "").lower()
    if len(v) == 18 and RE_SFID18.fullmatch(v):
        return v[:15]
    return v


TYPE_LABELS = {"int": "Integer", "dec": "Decimal", "bool": "Boolean", "date": "Date", "datetime": "Date/time",
               "time": "Time", "guid": "GUID", "email": "Email", "url": "URL", "phone": "Phone",
               "token": "Id / code token", "text": "Text"}


# ===========================================================================
# Profile
# ===========================================================================

class Profile:
    def __init__(self, data=None, path=""):
        d = data or {}
        self.path = path
        self.name = d.get("name", "")
        self.source_system = d.get("source_system", "")
        self.target_system = d.get("target_system", "")
        dc = d.get("deletion_columns", {})
        self.del_ts = [norm(c) for c in dc.get("timestamp", [])] + DELETION_TIMESTAMP
        self.del_bool = [norm(c) for c in dc.get("boolean", [])] + DELETION_BOOLEAN
        self.not_fk = NOT_FK | {norm(c) for c in d.get("not_fk_columns", [])}
        self.user_columns = {norm(c) for c in d.get("user_columns", [])}
        self.not_user = {norm(c) for c in d.get("not_user_columns", [])}
        self.generic_stems = GENERIC_STEMS | {norm(c) for c in d.get("generic_stems", [])}
        self.audit_cols = AUDIT_COLUMNS | {norm(c) for c in d.get("audit_columns", [])}
        self.type_map = {k.lower(): v.lower() for k, v in d.get("type_table_map", {}).items()}
        self.fk_overrides = {k.lower(): [t.lower() for t in v] for k, v in d.get("fk_overrides", {}).items()}
        self.fk_notes = {k.lower(): v for k, v in d.get("fk_override_notes", {}).items()}
        # extra values that mean "blank" in this source (e.g. Salesforce Data Export's 000000000000000AAA)
        self.null_tokens = {str(x).strip().lower() for x in d.get("null_tokens", [])}
        # keys whose resolution is not an issue on rows meeting a condition, e.g.
        # {"account.personcontactid": {"exempt_when": {"column": "IsPersonAccount", "values": ["true", "1"]}, "note": "..."}}
        self.conditional_keys = {k.lower(): v for k, v in d.get("conditional_keys", {}).items()}
        self.system_tables = {t.lower() for t in d.get("system_tables", [])}
        self.entity_tables = {t.lower() for t in d.get("entity_tables", [])}
        self.users_table = (d.get("users_table") or "").lower()
        self.tables = {k.lower(): v for k, v in d.get("tables", {}).items()}
        self.intake = d.get("intake") or GENERIC_INTAKE
        self.ui_label = d.get("ui_location_label", "Source UI Location")
        self.has_ui = any(v.get("ui") for v in self.tables.values())
        self.has_article = any(v.get("article") for v in self.tables.values())


PROFILE_TEMPLATE = {
    "name": "My Source System",
    "source_system": "My Source System v1 (exported to CSV)",
    "target_system": "Salesforce",
    "deletion_columns": {"timestamp": [], "boolean": []},
    "not_fk_columns": ["columns ending in _id that are NOT keys, e.g. tax_id"],
    "user_columns": ["columns holding a user id that the generic rules miss"],
    "not_user_columns": ["columns that look like user references but are free text, e.g. referred_by"],
    "users_table": "",
    "type_table_map": {"App::Contact": "contacts"},
    "null_tokens": ["values that mean blank in this source, e.g. 000000000000000AAA"],
    "detect": {"tables": ["customers", "orders"], "columns": {"customers": ["id", "created_at"]}},
    "conditional_keys": {"account.personcontactid": {"exempt_when": {"column": "IsPersonAccount", "values": ["true", "1"]},
                                                     "note": "Why the key is not an issue on these rows"}},
    "fk_overrides": {"orders.cust_ref": ["customers"], "status_id": ["order_statuses", "statuses"]},
    "fk_override_notes": {"orders.cust_ref": "Why this override exists / open question"},
    "system_tables": ["settings", "permissions"],
    "entity_tables": ["customers", "orders"],
    "ui_location_label": "Source UI Location",
    "tables": {
        "customers": {"description": "Customer master", "target": "Account",
                      "included": "Included", "ui": "Customers screen", "article": ""},
        "invoices": {"description": "Invoices", "target": "Custom object Invoice__c",
                     "included": "To Be Reviewed", "gated_by": "Invoices in scope"}
    },
    "intake": [
        {"topic": "Invoices in scope", "question": "Migrate invoices?", "why": "Largest table; often stays in the ERP.",
         "default": "No", "drives": "invoices, invoice_lines", "gates_tables": ["invoices", "invoice_lines"]},
        {"topic": "Target data model", "question": "Which data model?", "why": "Drives every mapping.",
         "default": "Detect from org", "drives": "all", "options": ["Standard", "Custom", "Undecided"]}
    ],
    "_notes": "Every key is optional. 'intake' REPLACES the generic intake when present. "
              "Table and column names are matched case-insensitively. 'included' is one of: "
              "To Be Reviewed, Included, Not Included, Not Applicable. 'gated_by' names an intake "
              "topic whose Yes/No answer flips the table's inclusion. Intake rows can also use "
              "'gates_classes' (e.g. [\"Audit / History\"]) to gate every table of a classification."
}


# ===========================================================================
# CSV scanning
# ===========================================================================

def raise_field_limit():
    limit = sys.maxsize
    while True:
        try:
            csv.field_size_limit(limit)
            return
        except OverflowError:
            limit = int(limit / 10)


def human_size(n):
    if n < 1000:
        return f"{n} bytes"
    if n < 1_000_000:
        return f"{int(round(n / 1000))} KB"
    if n < 1_000_000_000:
        return f"{n / 1_000_000:.1f} MB"
    return f"{n / 1_000_000_000:.2f} GB"


def detect_encoding(path):
    with open(path, "rb") as f:
        head = f.read(4)
    if head.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig", "UTF-8 (BOM)"
    if head.startswith(b"\xff\xfe") or head.startswith(b"\xfe\xff"):
        return "utf-16", "UTF-16"
    return "utf-8-sig", "UTF-8"


def detect_delimiter(path, encoding, ext):
    with open(path, "r", encoding=encoding, errors="replace", newline="") as f:
        first = f.readline()
    counts = {}
    for d in (",", "\t", ";", "|"):
        # count delimiters outside double-quoted segments
        counts[d] = len(re.sub(r'"[^"]*"', "", first).split(d)) - 1
    best = max(counts, key=counts.get)
    if counts[best] == 0:
        return "\t" if ext == ".tsv" else ("|" if ext == ".psv" else ",")
    return best


DELIM_NAMES = {",": "comma", "\t": "tab", ";": "semicolon", "|": "pipe"}


class Table:
    def __init__(self, filename, path, profile, forced_delim=None, forced_enc=None):
        self.filename = filename
        self.name = table_base_name(filename).lower()
        self.display = table_base_name(filename)
        self.path = path
        self.p = profile
        self.size = os.path.getsize(path)
        self.forced_delim, self.forced_enc = forced_delim, forced_enc
        self.aliases = table_aliases(self.display)
        self.tokens = tokenize(self.display)
        self.error = ""
        self.encoding_label = ""
        self.delim = ","
        self._reset()
        # derived later
        self.pk = None                       # column name
        self.pk_note = ""
        self.pk_kind = ""                    # "int" | "token" | ""
        self.classification = "Unclassified"
        self.class_reason = ""
        self.fks = []
        self.referenced_by = []
        self.key_cols = set()                # columns treated as keys (excluded from attributes)
        self.poly_by_value = set()           # token keys whose values span several tables

    def _reset(self):
        self.headers = []
        self.total = self.live = 0
        self.ragged = 0
        self.del_col = ""
        self.header_suspect = False
        self.dup_headers = []
        self.populated = Counter()
        self.distinct = {}                   # col -> set | None (overflow)
        self.types = {}                      # col -> Counter(kind)
        self.typed = Counter()
        self.only01 = {}                     # col -> bool
        self.maxlen = Counter()
        self.multiline = Counter()
        self.leading_zero = Counter()
        self.has_space = Counter()
        self.pk_sets = {}                    # candidate col -> set(hash) over ALL rows
        self.pk_dupes = Counter()
        self.pk_nulls = Counter()
        self.deleted_keys = {}               # candidate col -> set(hash) of deleted rows
        self.fk_values = {}                  # col -> Counter(hash(norm_key)) on live rows
        self.fk_kind = {}                    # col -> "int"/"token"/"mixed"
        self.token_cols = set()              # columns tracked only while every value is a token id
        self.poly_pairs = {}                 # id col -> type col
        self.exempt_values = {}              # conditional key col -> Counter(hash) on rows meeting the exemption
        self.poly_values = {}                # id col -> Counter((type, hash))
        self.type_values = {}                # type col -> Counter(type value)

    # ---- column-name tests -------------------------------------------------
    def is_user_col(self, col):
        n = norm(col)
        if n in self.p.not_user:
            return False
        if n in self.p.user_columns:
            return True
        toks = tokenize(col)
        core = toks[:-1] if toks and toks[-1] in STRONG_KEY_SUFFIXES else toks
        if not core:
            return False
        joined = "".join(core)
        return core[-1] == "by" or joined in USER_STEMS or (core[-1] == "user" and len(core) <= 3)

    def key_suffix(self, col):
        toks = tokenize(col)
        if not toks:
            return "", toks
        if toks[0] == "fk" and len(toks) > 1:
            return "strong", toks[1:]
        if toks[-1] in STRONG_KEY_SUFFIXES and len(toks) > 1:
            return "strong", toks[:-1]
        if toks[-1] in WEAK_KEY_SUFFIXES and len(toks) > 1:
            return "weak", toks[:-1]
        return "", toks

    def pk_name_candidates(self):
        names = ["id"] + [a + s for a in sorted(self.aliases) for s in ("id", "key", "guid", "uuid", "no", "number", "code")] \
                + ["uuid", "guid", "pk", "rowid", "recordid", "objectid", "code", "key"]
        out = []
        for want in names:
            for i, h in enumerate(self.headers):
                if norm(h) == want and i not in out:
                    out.append(i)
        return out

    # ---- scanning ----------------------------------------------------------
    def scan(self):
        if self.size == 0:
            return
        enc, label = (self.forced_enc, self.forced_enc) if self.forced_enc else detect_encoding(self.path)
        ext = os.path.splitext(self.filename)[1].lower()
        self.delim = self.forced_delim or detect_delimiter(self.path, enc, ext)
        try:
            self._scan(enc)
            self.encoding_label = label
        except UnicodeDecodeError:
            self._reset()
            self._scan("latin-1")
            self.encoding_label = "Latin-1 (fallback)"

    def _scan(self, encoding):
        with open(self.path, "r", encoding=encoding, newline="") as f:
            reader = csv.reader(f, delimiter=self.delim)
            try:
                raw = next(reader)
            except StopIteration:
                return
            self.headers = [h.strip().lstrip("﻿") for h in raw]
            n = len(self.headers)
            seen = Counter(h.lower() for h in self.headers)
            self.dup_headers = [h for h, c in seen.items() if c > 1]
            numeric = sum(1 for h in self.headers if not h or RE_INT.fullmatch(h) or RE_DEC.fullmatch(h) or RE_DATE.match(h))
            self.header_suspect = n > 0 and numeric / n >= 0.5
            norms = [norm(h) for h in self.headers]

            del_idx, del_style = None, None
            for cand in self.p.del_ts:
                if cand in norms:
                    del_idx, del_style = norms.index(cand), "ts"
                    break
            if del_idx is None:
                for cand in self.p.del_bool:
                    if cand in norms:
                        del_idx, del_style = norms.index(cand), "bool"
                        break
            if del_idx is not None:
                self.del_col = self.headers[del_idx]

            # polymorphic pairs: X_id + X_type (or X_table / X_class …)
            poly = {}
            for i, h in enumerate(self.headers):
                kind, core = self.key_suffix(h)
                if kind != "strong" or not core:
                    continue
                base = "".join(core)
                for suf in POLY_TYPE_SUFFIXES:
                    if base + suf in norms and norms.index(base + suf) != i:
                        poly[i] = norms.index(base + suf)
                        break
            self.poly_pairs = {self.headers[i]: self.headers[t] for i, t in poly.items()}

            pk_idx = self.pk_name_candidates()
            if 0 not in pk_idx and n:
                pk_idx.append(0)
            fk_idx = []
            for i, h in enumerate(self.headers):
                if i in poly or i == del_idx:
                    continue
                kind, _ = self.key_suffix(h)
                if norms[i] in self.p.not_fk:
                    continue
                if kind or self.is_user_col(h) or norms[i] in SELF_STEMS:
                    if norms[i] != "id":
                        fk_idx.append(i)
            token_idx = {i for i in range(n) if i not in fk_idx and i not in poly and i != del_idx
                         and norms[i] not in self.p.not_fk and norms[i] != "id"}

            for i in pk_idx:
                self.pk_sets[self.headers[i]] = set()
                self.deleted_keys[self.headers[i]] = set()
            for i in set(fk_idx) | token_idx:
                self.fk_values[self.headers[i]] = Counter()
            for h in self.headers:
                self.distinct[h] = set()
                self.types[h] = Counter()
                self.only01[h] = True
            for i in poly:
                self.poly_values[self.headers[i]] = Counter()
                self.type_values[self.headers[poly[i]]] = Counter()

            headers = self.headers
            cond = {}
            for i, h in enumerate(headers):
                rule = self.p.conditional_keys.get(f"{self.name}.{h.lower()}")
                if rule:
                    ew = rule.get("exempt_when", {})
                    cn = norm(ew.get("column", ""))
                    if cn in norms:
                        vals = None if ew.get("populated") else {str(x).strip().lower() for x in ew.get("values", [])}
                        cond[i] = (norms.index(cn), vals)
                        self.exempt_values[h] = Counter()
            for row in reader:
                if not row or all(c.strip() == "" for c in row):
                    continue
                self.total += 1
                if len(row) != n:
                    self.ragged += 1
                deleted = False
                if del_idx is not None and del_idx < len(row):
                    v = row[del_idx].strip()
                    deleted = (v.lower() not in NULL_TOKENS) if del_style == "ts" else (v.lower() not in (NULL_TOKENS | EMPTYISH))
                for i in pk_idx:
                    if i >= len(row):
                        continue
                    v = row[i].strip()
                    col = headers[i]
                    if v.lower() in NULL_TOKENS:
                        self.pk_nulls[col] += 1
                        continue
                    hv = hash(norm_key(v))
                    s = self.pk_sets[col]
                    if hv in s:
                        self.pk_dupes[col] += 1
                    else:
                        s.add(hv)
                    if deleted:
                        self.deleted_keys[col].add(hv)
                if deleted:
                    continue
                self.live += 1
                for i, cell in enumerate(row[:n]):
                    v = cell.strip()
                    lv = v.lower()
                    if lv in NULL_TOKENS:
                        continue
                    col = headers[i]
                    if lv not in EMPTYISH:
                        self.populated[col] += 1
                    L = len(v)
                    if L > self.maxlen[col]:
                        self.maxlen[col] = L
                    if "\n" in v:
                        self.multiline[col] += 1
                    if " " in v:
                        self.has_space[col] += 1
                    ds = self.distinct[col]
                    if ds is not None:
                        ds.add(hash(v))
                        if len(ds) > DISTINCT_CAP:
                            self.distinct[col] = None
                    if self.typed[col] < TYPE_SAMPLE:
                        k = value_kind(v)
                        self.types[col][k] += 1
                        self.typed[col] += 1
                        if k == "int" and len(v) > 1 and v[0] == "0":
                            self.leading_zero[col] += 1
                        if v not in ("0", "1"):
                            self.only01[col] = False
                    elif v not in ("0", "1"):
                        self.only01[col] = False
                    if i in token_idx:
                        if is_token_id(v):
                            self.fk_values[col][hash(norm_key(v))] += 1
                        else:
                            token_idx.discard(i)
                            del self.fk_values[col]
                    elif col in self.fk_values and lv not in ("0",):
                        self.fk_values[col][hash(norm_key(v))] += 1
                for i, (ci, vals) in cond.items():
                    if i < len(row) and ci < len(row) and (
                            (vals is None and row[ci].strip().lower() not in NULL_TOKENS)
                            or (vals is not None and row[ci].strip().lower() in vals)):
                        v = row[i].strip()
                        if v.lower() not in NULL_TOKENS:
                            self.exempt_values[headers[i]][hash(norm_key(v))] += 1
                for i, ti in poly.items():
                    if i < len(row) and ti < len(row):
                        v, t = row[i].strip(), row[ti].strip()
                        if v.lower() not in NULL_TOKENS and v != "0":
                            self.poly_values[headers[i]][(t, hash(norm_key(v)))] += 1
                        if t:
                            self.type_values[headers[ti]][t] += 1
            self.token_cols = {headers[i] for i in token_idx}
            # token-tracked columns that are entirely empty carry no evidence
            for col in list(self.token_cols):
                if not self.fk_values.get(col):
                    self.token_cols.discard(col)
                    self.fk_values.pop(col, None)

    # ---- derived column facts ----------------------------------------------
    def col_type(self, col):
        c = self.types.get(col)
        if not c:
            return ""
        total = sum(c.values())
        if self.only01.get(col) and set(c) <= {"int"}:
            return "Boolean (0/1)"
        merged = Counter()
        for k, v in c.items():
            merged[{"int": "num", "dec": "num", "date": "dt", "datetime": "dt"}.get(k, k)] += v
        top, cnt = merged.most_common(1)[0]
        if cnt / total >= 0.95:
            if top == "num":
                return "Decimal" if c.get("dec") else "Integer"
            if top == "dt":
                return "Date/time" if c.get("datetime") else "Date"
            return TYPE_LABELS[top]
        a = merged.most_common(2)
        name = lambda k: {"num": "Number", "dt": "Date"}.get(k, TYPE_LABELS.get(k, k))
        return f"Mixed: {name(a[0][0])} / {name(a[1][0])}"

    def distinct_count(self, col):
        ds = self.distinct.get(col)
        return None if ds is None else len(ds)

    def choose_pk(self):
        if not self.headers:
            return
        cands = [c for c in self.pk_sets]
        named = [c for c in cands if self.headers.index(c) in self.pk_name_candidates()]
        for col in named + [c for c in cands if c not in named]:
            if self.total and self.pk_dupes[col] == 0 and self.pk_nulls[col] == 0 and len(self.pk_sets[col]) == self.total:
                # first column only counts if it looks like an identifier, not a name or date
                if col not in named:
                    t = self.col_type(col)
                    if t not in ("Integer", "GUID", "Id / code token"):
                        continue
                self.pk = col
                t = self.col_type(col)
                self.pk_kind = "int" if t == "Integer" else ("token" if t in ("GUID", "Id / code token") else "other")
                break
        if self.pk is None and named:
            col = named[0]
            if self.pk_dupes[col] or self.pk_nulls[col]:
                self.pk_note = (f"{col} is not unique: {self.pk_dupes[col]:,} duplicate value(s), "
                                f"{self.pk_nulls[col]:,} blank")
        # free memory held for rejected candidates
        for col in list(self.pk_sets):
            if col != self.pk:
                del self.pk_sets[col]
                self.deleted_keys.pop(col, None)
        if self.pk:
            self.key_cols.add(self.pk)
            self.fk_values.pop(self.pk, None)
            self.token_cols.discard(self.pk)

    @property
    def pk_all(self):
        return self.pk_sets.get(self.pk, set()) if self.pk else set()

    @property
    def pk_deleted(self):
        return self.deleted_keys.get(self.pk, set()) if self.pk else set()


# ===========================================================================
# Key resolution
# ===========================================================================

def is_users_table(t, profile):
    if profile.users_table:
        return t.name == profile.users_table
    return bool(t.aliases & USER_TABLE_NAMES)


def stem_variants(t, col):
    kind, core = t.key_suffix(col)
    if not kind and (t.is_user_col(col) or norm(col) in SELF_STEMS):
        core = tokenize(col)
    if not core:
        return kind, [], []
    plain = ["".join(core[i:]) for i in range(len(core))]
    ts, first = singular("".join(t.tokens)), (t.tokens[0] if t.tokens else "")
    prefixed = []
    for s in plain:
        for pre in (ts, first):
            if pre and not s.startswith(pre):
                prefixed.append(pre + s)
    return kind, plain, prefixed


def candidate_tables(t, col, tables, profile):
    """Ordered (table, source) candidates for a key column, before the empirical check."""
    out, seen = [], set()

    def add(name, src, allow_self=False):
        if name == "__self__":
            name, allow_self = t.name, True
        if name == t.name and not allow_self:
            return
        if name in tables and name not in seen:
            out.append((name, src))
            seen.add(name)

    for key in (f"{t.name}.{col.lower()}", col.lower()):
        if key in profile.fk_overrides:
            for name in profile.fk_overrides[key]:
                add(name, "profile override", allow_self=True)
            break
    n = norm(col)
    for u in tables.values():
        if u is not t and u.pk and norm(u.pk) == n and n != "id":
            add(u.name, "matches that table's key name")
    kind, plain, prefixed = stem_variants(t, col)
    if plain and (plain[0] in SELF_STEMS or plain[-1] in SELF_STEMS or "".join(plain[:1]) in SELF_STEMS):
        add(t.name, "self-reference name", allow_self=True)
    by_alias = defaultdict(list)
    for u in tables.values():
        for a in u.aliases:
            by_alias[a].append(u.name)
    for s in prefixed:
        for form in (s, singular(s)):
            for name in by_alias.get(form, []):
                add(name, "naming convention")
    for s in plain:
        for form in (s, singular(s)):
            for name in by_alias.get(form, []):
                add(name, "naming convention")
    for s in plain[:1]:
        if s in profile.generic_stems or len(s) < 4:
            continue
        for u in tables.values():
            if any(a.endswith(s) or a.endswith(singular(s)) for a in u.aliases):
                add(u.name, "naming convention (suffix)")
    if t.is_user_col(col):
        for u in tables.values():
            if is_users_table(u, profile):
                add(u.name, "users table")
    return out


def type_to_table(type_val, tables, profile):
    key = type_val.strip().lower()
    if key in profile.type_map:
        return profile.type_map[key]
    last = re.split(r"::|\\|/|\.", type_val.strip())[-1]
    forms = {norm(last), singular(norm(last))}
    for u in tables.values():
        if u.aliases & forms or {singular(a) for a in u.aliases} & forms:
            return u.name
    if forms & USER_TABLE_NAMES:
        return "__users__"
    return ""


def score(counter, target):
    live_rows = sum(counter.values())
    hit = sum(n for h, n in counter.items() if h in target.pk_all)
    hit_deleted = sum(n for h, n in counter.items() if h in target.pk_deleted)
    return live_rows, hit, hit_deleted, (hit / live_rows if live_rows else None)


def resolve_keys(tables, profile):
    has_users_table = any(is_users_table(u, profile) for u in tables.values())
    token_targets = [u for u in tables.values() if u.pk and u.pk_kind == "token"]
    int_targets = [u for u in tables.values() if u.pk and u.pk_kind == "int"]

    for t in tables.values():
        # ---- ordinary keys (named like keys, or token-valued) -------------------
        for col, counter in list(t.fk_values.items()):
            if not counter:
                continue
            ctype = t.col_type(col)
            is_user = t.is_user_col(col)
            kind, _ = t.key_suffix(col)
            named_key = bool(kind) or is_user or norm(col) in SELF_STEMS
            # a column named like a key but holding sentences is an attribute
            if named_key and (t.has_space[col] / max(1, sum(t.types[col].values())) > 0.2 or ctype == "Text"):
                continue
            cands = candidate_tables(t, col, tables, profile)
            note = profile.fk_notes.get(f"{t.name}.{col.lower()}", "")
            results = []
            for name, src in cands:
                u = tables[name]
                if not u.pk:
                    results.append((name, src, sum(counter.values()), 0, 0, None))
                    continue
                lr, hit, hd, pct = score(counter, u)
                results.append((name, src, lr, hit, hd, pct))
            live_rows = sum(counter.values())
            scored = [r for r in results if r[5] is not None]
            best = max(scored, key=lambda r: (r[5], -results.index(r))) if scored else None

            value_pick, value_note = None, ""
            if (best is None or best[5] < 0.5) and (col in t.token_cols or named_key) and not (is_user and col not in t.token_cols):
                pool = token_targets if (col in t.token_cols or ctype in ("GUID", "Id / code token")) else (int_targets if named_key else [])
                hits = []
                for u in pool:
                    if u is t and not (kind and stem_variants(t, col)[1][:1] and stem_variants(t, col)[1][0] in SELF_STEMS):
                        continue
                    lr, hit, hd, pct = score(counter, u)
                    if pct is not None and pct >= 0.95:
                        hits.append((u.name, "value match", lr, hit, hd, pct))
                if not hits and pool is token_targets:
                    spread = []
                    for u in pool:
                        lr, hit, hd, pct = score(counter, u)
                        if pct and pct >= 0.01:
                            spread.append((u, lr, hit, hd, pct))
                    if len(spread) >= 2 and sum(x[4] for x in spread) >= 0.95:
                        for u, lr, hit, hd, pct in sorted(spread, key=lambda x: -x[4]):
                            sub_rows = hit
                            t.fks.append(dict(column=f"{col} [{u.display}]", role="Polymorphic Key", target=u.name, target_col=u.pk,
                                              rel="Polymorphic", how="inferred from values — confirm", conf="Medium",
                                              live_rows=sub_rows, resolved=hit, pct=1.0,
                                              note=f"{pct:.1%} of {col} values are {u.name} ids. No type column — the id itself identifies the parent."))
                            u.referenced_by.append((t.name, f"{col} [{u.display}]"))
                        t.key_cols.add(col)
                        t.poly_by_value.add(col)
                        continue
                distinct = t.distinct_count(col)
                if len(hits) == 1 and (pool is token_targets or (distinct is None or distinct >= 20)):
                    value_pick = hits[0]
                elif len(hits) > 1:
                    hits.sort(key=lambda r: -r[5])
                    value_note = ("Ambiguous — every value also exists in: " + ", ".join(h[0] for h in hits[:6])
                                  + (" …" if len(hits) > 6 else "") + ". Integer ids overlap across tables; confirm the target.")

            if best is None and value_pick is None and not value_note:
                if col in t.token_cols and not named_key:
                    continue                     # an ordinary token-valued attribute, not a key
                if kind == "weak" and not is_user:
                    continue                     # phone_number, zip_code … not keys
                if is_user and not has_users_table:
                    t.fks.append(dict(column=col, role="User Reference", target="Users (no users table in extract)", target_col="",
                                      rel="User reference", how="user column name", conf="—", live_rows=live_rows, resolved=0, pct=None,
                                      note="No users table in this extract — resolves to a users list captured from the source system (Engagement Intake)."))
                    t.key_cols.add(col)
                    continue
                t.fks.append(dict(column=col, role="Foreign Key", target="", target_col="", rel="Unresolved — no matching table",
                                  how="no candidate table", conf="—", live_rows=live_rows, resolved=0, pct=None,
                                  note=note or f"No table in this extract matches '{col}'. Confirm whether the source table is missing from the export."))
                t.key_cols.add(col)
                continue
            if best is None and value_pick is None:
                t.fks.append(dict(column=col, role="Foreign Key", target="", target_col="", rel="Ambiguous",
                                  how="value match", conf="Low", live_rows=live_rows, resolved=0, pct=None, note=value_note))
                t.key_cols.add(col)
                continue

            chosen, how_extra = best, ""
            if value_pick and (best is None or value_pick[5] > (best[5] or 0)):
                chosen = value_pick
                if best is not None:
                    how_extra = f" Name suggested {best[0]} ({best[5]*100:.1f}%), but values match {value_pick[0]}."
            name, src, lr, hit, hd, pct = chosen
            u = tables[name]
            cond_note = ""
            rule = profile.conditional_keys.get(f"{t.name}.{col.lower()}")
            ex = t.exempt_values.get(col)
            if rule and ex is None:
                cond_note = (f"Conditional rule ({rule.get('exempt_when', {}).get('column', '')}) not applied: "
                             f"that column is not in {t.filename}. Resolution below covers every row.")
            elif rule:
                ew = rule.get("exempt_when", {})
                label = ew.get("label", "")
                cond_txt = (f"{ew.get('column', '')} populated" if ew.get("populated")
                            else f"{ew.get('column', '')} = {'/'.join(ew.get('values', []))}")
                ex_rows = sum(ex.values()) if ex else 0
                ex_unres = sum(n for h, n in ex.items() if h not in u.pk_all) if ex else 0
                if ex_unres:
                    hit = min(lr, hit + ex_unres)
                    pct = hit / lr if lr else None
                if label:
                    other = ew.get("other_label", "other rows")
                    cond_note = (f"{ex_rows:,} {label} ({cond_txt}); {max(0, t.live - ex_rows):,} {other} (blank — nothing to resolve). "
                                 f"Not an issue on {label}"
                                 + (f": {ex_unres:,} without a matching {name} row are counted as resolved." if ex_unres else f" — all {ex_rows:,} resolve to {name}."))
                else:
                    cond_note = (f"Conditional rule ({cond_txt}): not an issue on those rows. {ex_rows:,} live rows meet it"
                                 + (f"; {ex_unres:,} without a matching {name} row are counted as resolved." if ex_unres else "; all resolve."))
                if rule.get("note"):
                    cond_note += f" {rule['note']}"
            if src == "value match":
                conf = "Medium" if u.pk_kind == "token" else "Low"
                how = "inferred from values — confirm"
            else:
                conf = "High" if pct >= 0.95 else ("Medium" if pct >= 0.5 else "Low")
                how = src if len(results) <= 1 else f"{src} — empirical choice among {len(results)}"
            if src == "profile override" and len(results) <= 1:
                how = "profile override"
            notes = [note] if note else []
            if cond_note:
                notes.append(cond_note)
            if how_extra:
                notes.append(how_extra.strip())
            ties = [r[0] for r in scored if r is not chosen and r[5] is not None and pct is not None and abs(r[5] - pct) < 1e-9]
            if ties:
                notes.append(f"TIE — {', '.join(ties)} resolve equally well; confirm the intended target.")
                conf = "Low" if conf != "Low" else conf
            alts = [f"{r[0]}: {r[5]*100:.1f}%" for r in results if r is not chosen and r[5] is not None]
            if alts:
                notes.append("Alternatives: " + "; ".join(alts) + ".")
            if lr and pct is not None and pct < 0.95:
                notes.append(f"{lr - hit:,} live rows point at ids not present in {name}.")
            if hd:
                notes.append(f"{hd:,} live rows point at DELETED {name} rows.")
            if is_user and is_users_table(u, profile):
                role, rel = "User Reference", "User reference"
            else:
                role, rel = "Foreign Key", ("Self-reference" if name == t.name else "")
            t.fks.append(dict(column=col, role=role, target=name, target_col=u.pk or "", rel=rel, how=how, conf=conf,
                              live_rows=lr, resolved=hit, pct=pct, note=" ".join(notes)))
            t.key_cols.add(col)
            tables[name].referenced_by.append((t.name, col))

        # ---- polymorphic -------------------------------------------------------
        for col, counter in t.poly_values.items():
            type_col = t.poly_pairs[col]
            t.key_cols |= {col, type_col}
            by_type = defaultdict(Counter)
            for (tv, h), n in counter.items():
                by_type[tv][h] += n
            for tv, sub in sorted(by_type.items()):
                lr = sum(sub.values())
                target = type_to_table(tv, tables, profile) if tv else ""
                if target == "__users__" or (target and is_users_table(tables[target], profile) and not tables[target].pk):
                    t.fks.append(dict(column=f"{col} [{tv}]", role="Polymorphic Key", target="Users", target_col="", rel="User reference",
                                      how="polymorphic", conf="—", live_rows=lr, resolved=0, pct=None, note="Resolves to the users list."))
                elif target and target in tables and tables[target].pk:
                    u = tables[target]
                    lr, hit, hd, pct = score(sub, u)
                    note = f"Where {type_col} = '{tv}'."
                    if lr and pct < 0.95:
                        note += f" {lr - hit:,} live rows unresolved."
                    if hd:
                        note += f" {hd:,} point at deleted {target} rows."
                    t.fks.append(dict(column=f"{col} [{tv}]", role="Polymorphic Key", target=target, target_col=u.pk, rel="Polymorphic",
                                      how="polymorphic type value", conf="High" if pct >= 0.95 else ("Medium" if pct >= 0.5 else "Low"),
                                      live_rows=lr, resolved=hit, pct=pct, note=note))
                    u.referenced_by.append((t.name, f"{col} [{tv}]"))
                else:
                    t.fks.append(dict(column=f"{col} [{tv or 'blank'}]", role="Polymorphic Key", target=tv or "(blank type)", target_col="",
                                      rel="Unresolved — unknown type", how="polymorphic type value", conf="—", live_rows=lr, resolved=0, pct=None,
                                      note=f"No table matched {type_col} value '{tv}'. Add it to type_table_map in a profile."))


def attributes(t, profile):
    out = []
    for c in t.headers:
        n = norm(c)
        if c in t.key_cols or c == t.del_col or n in profile.audit_cols or n == "id":
            continue
        out.append(c)
    return out


def has_label(t):
    for c in t.headers:
        toks = tokenize(c)
        if toks and (toks[-1] in LABEL_TOKENS or norm(c) in LABEL_TOKENS):
            return True
    return False


def classify(tables, profile):
    for t in tables.values():
        if t.error:
            t.classification, t.class_reason = "Unclassified", "File could not be parsed"
            continue
        if not t.headers:
            t.classification, t.class_reason = "Unclassified", "Empty file — no header row"
            continue
        toks = set(t.tokens)
        good = [f for f in t.fks if f["target"] and not f["rel"].startswith(("Unresolved", "Ambiguous"))]
        parents = [f for f in good if f["role"] == "Foreign Key" and f["target"] != t.name]
        audit_user = lambda c: tokenize(c)[:1] and tokenize(c)[0] in ("created", "modified", "updated", "lastmodified", "added", "entered", "deleted")
        users = [f for f in good if f["role"] == "User Reference" and not audit_user(f["column"])]
        user_cols = [f["column"] for f in t.fks if f["role"] == "User Reference" and not audit_user(f["column"])]
        attrs = attributes(t, profile)
        n_ref_by = len({r[0] for r in t.referenced_by if r[0] != t.name})
        poly = list(t.poly_pairs) + sorted(t.poly_by_value)
        if t.name in profile.system_tables:
            t.classification, t.class_reason = "System / Config", "Listed as a system table in the profile"
        elif t.name in profile.entity_tables:
            t.classification, t.class_reason = "Entity", f"Anchor object (profile); referenced by {n_ref_by} table(s)"
        elif toks & STAGING_TABLE_TOKENS or re.search(r"(_|\b)(19|20)\d{2}(\d{4})?$", t.display):
            t.classification, t.class_reason = "Staging / Backup", "Name looks like a copy, backup, staging or dated table"
        elif toks & AUDIT_TABLE_TOKENS:
            t.classification, t.class_reason = "Audit / History", "Name looks like an audit, log or history table"
        elif is_users_table(t, profile):
            t.classification, t.class_reason = "Entity", f"Users table; referenced by {n_ref_by} table(s)"
        elif poly:
            why = [f"{k}/{v} (a type column selects the parent)" for k, v in t.poly_pairs.items()] + \
                  [f"{k} (values span several parent tables)" for k in sorted(t.poly_by_value)]
            t.classification, t.class_reason = "Polymorphic Child", "; ".join(why)
        elif len(parents) + len(users) >= 2 and parents and len(attrs) <= 2:
            keys = [f["column"] for f in parents + users]
            t.classification, t.class_reason = "Junction (many-to-many)", f"{len(keys)} keys ({', '.join(keys)}) and {len(attrs)} attribute column(s)"
        elif not parents and not user_cols and has_label(t) and len(attrs) <= 3:
            t.classification, t.class_reason = "Lookup / Picklist", f"Key + label column, {len(attrs)} attribute(s), {t.live:,} live rows; referenced by {n_ref_by} table(s)"
        elif n_ref_by >= 1 and len(attrs) >= 3:
            t.classification, t.class_reason = "Entity", f"Referenced by {n_ref_by} table(s), {len(attrs)} attribute columns"
        elif not parents and len(attrs) > 3:
            t.classification, t.class_reason = "Entity", f"No parent keys, {len(attrs)} attribute columns"
        elif user_cols and n_ref_by >= 1:
            t.classification, t.class_reason = "Entity", f"Owned records ({', '.join(user_cols)}) referenced by {n_ref_by} table(s)"
        elif parents and attrs:
            t.classification, t.class_reason = "Child / Detail", f"{len(parents)} parent key(s) ({', '.join(f['column'] for f in parents)}) plus {len(attrs)} attribute column(s); nothing references it"
        elif not parents and has_label(t):
            t.classification, t.class_reason = "Lookup / Picklist", f"Key + label, {len(attrs)} attributes"
        else:
            t.classification, t.class_reason = "Unclassified", "Structure did not match a pattern — inspect the Schema tab"
    # relationship labels need the target's class
    for t in tables.values():
        for fk in t.fks:
            if fk["rel"] == "" and fk["target"] in tables:
                fk["rel"] = "Lookup" if tables[fk["target"]].classification == "Lookup / Picklist" else "Parent"


def column_role(t, col):
    n = norm(col)
    if col == t.pk:
        return "Primary Key"
    if t.del_col and col == t.del_col:
        return "Deletion Flag"
    if col in t.poly_pairs.values():
        return "Polymorphic Type"
    if col in t.poly_pairs or col in t.poly_by_value:
        return "Polymorphic Key"
    roles = {fk["role"] for fk in t.fks if fk["column"].split(" [")[0] == col}
    if "User Reference" in roles:
        return "User Reference"
    if "Foreign Key" in roles:
        return "Foreign Key"
    if n in t.p.not_fk:
        return "Attribute (identifier)"
    if n in t.p.audit_cols:
        return "Audit / System"
    return "Attribute"


# ===========================================================================
# Findings
# ===========================================================================

def build_findings(tables):
    F = []

    def add(sev, cat, table, col, text, rows, action):
        F.append((sev, cat, table, col, text, rows, action))

    for t in sorted(tables.values(), key=lambda x: x.name):
        if t.error:
            add("High", "File", t.filename, "", f"File could not be parsed: {t.error}", None, "Re-export the file; check delimiter and quoting.")
            continue
        if t.size == 0 or not t.headers:
            add("Low", "File", t.filename, "", "Empty file (no header row).", None, "Confirm the table is genuinely empty in the source.")
            continue
        if t.header_suspect:
            add("High", "File", t.filename, "", "First row looks like data, not column names — header row may be missing.", None, "Re-export with column headers.")
        if t.ragged:
            add("High", "File", t.filename, "", f"{t.ragged:,} row(s) have a different number of fields than the header — usually broken quoting or embedded delimiters.", t.ragged,
                "Re-export with proper quoting, or confirm the delimiter.")
        if t.dup_headers:
            add("Medium", "File", t.filename, ", ".join(t.dup_headers), "Duplicate column names in the header.", None, "Rename in the export; only the first copy is reliable.")
        if t.encoding_label.startswith("Latin-1"):
            add("Low", "File", t.filename, "", "Not valid UTF-8 — read as Latin-1. Accented characters may need checking.", None, "Ask for a UTF-8 export if names contain accents.")
        if t.pk_note:
            add("High", "Keys", t.filename, t.pk_note.split(" ")[0], f"Primary key candidate {t.pk_note}.", None, "Investigate duplicates before mapping; they break upserts and parent lookups.")
        elif not t.pk and t.classification not in ("Junction (many-to-many)",) and t.total:
            add("Medium", "Keys", t.filename, "", "No single-column primary key detected.", None, "Confirm the key (it may be composite). Child records cannot find this table without one.")
        if not t.total:
            add("Low", "Volume", t.filename, "", "Header only — no rows.", 0, "Out of scope unless the source has rows the export missed.")
        elif t.live == 0:
            add("Low", "Volume", t.filename, "", "Every row is soft-deleted (or the table has no rows).", t.total, "Usually out of scope.")
        elif t.total and t.del_col and (t.total - t.live) / t.total > 0.5:
            add("Low", "Volume", t.filename, t.del_col, f"{(t.total - t.live) / t.total:.0%} of rows are soft-deleted.", t.total - t.live, "Confirm the deletion rule with the client.")
        for fk in t.fks:
            if fk["rel"].startswith("Unresolved") and fk["role"] != "User Reference":
                add("Medium", "Keys", t.filename, fk["column"], f"Key does not resolve: {fk['note']}", fk["live_rows"], "Request the missing table, or confirm the column is not a key.")
            elif fk["rel"] == "Ambiguous":
                add("Medium", "Keys", t.filename, fk["column"], fk["note"], fk["live_rows"], "Name the target in a profile fk_override.")
            elif fk["pct"] is not None and fk["pct"] < 1 and fk["live_rows"]:
                add("Medium" if fk["pct"] < 0.95 else "Low", "Orphans", t.filename, fk["column"],
                    f"{fk['live_rows'] - fk['resolved']:,} of {fk['live_rows']:,} live rows ({1 - fk['pct']:.1%}) point at {fk['target']} rows that do not exist.",
                    fk["live_rows"] - fk["resolved"], "Decide: drop orphans, attach to a placeholder parent, or fix at source.")
            if "DELETED" in fk["note"] or "deleted " in fk["note"]:
                m = re.search(r"([\d,]+) (live rows point at DELETED|point at deleted)", fk["note"])
                if m:
                    add("Low", "Orphans", t.filename, fk["column"], f"{m.group(1)} live rows point at soft-deleted {fk['target']} rows.",
                        int(m.group(1).replace(",", "")), "Decide whether these follow their deleted parent out of scope.")
        long_cols = [c for c in t.headers if t.maxlen[c] > 255]
        for c in long_cols:
            add("Low", "Sizing", t.filename, c, f"Values up to {t.maxlen[c]:,} characters — longer than a 255-character Text field.", None,
                "Use a Long Text Area (max 131,072) or Rich Text field.")
        for c in t.headers:
            if t.leading_zero[c] and t.col_type(c) == "Integer":
                add("Low", "Formatting", t.filename, c, "Numbers with leading zeros — Excel and Data Loader prep can strip them.", t.leading_zero[c],
                    "Treat as Text through every step.")
        if t.live:
            empty = [c for c in t.headers if t.populated[c] == 0 and c != t.del_col]
            if empty and len(t.headers) > 1:
                add("Low", "Coverage", t.filename, ", ".join(empty[:8]) + (" …" if len(empty) > 8 else ""),
                    f"{len(empty)} of {len(t.headers)} columns are never populated on live rows.", None, "Exclude from the mapping.")
    order = {"High": 0, "Medium": 1, "Low": 2}
    F.sort(key=lambda f: (order[f[0]], f[1], f[2]))
    return F


# ===========================================================================
# Workbook
# ===========================================================================

def build_workbook(tables, folder, client, extract, source, target, profile, out_path):
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter as L
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.formatting.rule import CellIsRule

    thin = Side(style="thin", color=HAIRLINE)
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

    def font(size=9, bold=False, color=INK, name=BODY_FONT, italic=False):
        return Font(name=name, size=size, bold=bold, color=color, italic=italic)

    def fill(c):
        return PatternFill("solid", fgColor=c)

    def wrap(h="left", v="top"):
        return Alignment(horizontal=h, vertical=v, wrap_text=True)

    def masthead(ws, tag, ncols):
        for col in range(1, ncols + 1):
            ws.cell(row=1, column=col).fill = fill(DEEP_TEAL)
        ws.cell(row=1, column=1, value="MigrationOS™").font = font(12, bold=True, color=WHITE, name=DISPLAY_FONT)
        ws.cell(row=1, column=1).alignment = Alignment(vertical="center")
        c = ws.cell(row=1, column=ncols, value=tag)
        c.font = font(11, color=LIGHT_TEAL, name=DISPLAY_FONT)
        c.alignment = Alignment(horizontal="right", vertical="center")
        ws.row_dimensions[1].height = 24
        ws.sheet_view.showGridLines = False

    def note_row(ws, text, ncols, row=2):
        c = ws.cell(row=row, column=1, value=text)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        c.font = font(9, italic=True, color=MUTED)
        c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = 28

    def header_row(ws, row, columns):
        for i, (label, width) in enumerate(columns, start=1):
            c = ws.cell(row=row, column=i, value=label)
            c.font = font(9, bold=True, color=WHITE)
            c.fill = fill(DEEP_TEAL)
            c.alignment = wrap(v="center")
            c.border = BORDER
            ws.column_dimensions[L(i)].width = width
        ws.row_dimensions[row].height = 30

    def cell(ws, r, c, val, base, bold=False, color=INK, italic=False, h="left", fmt=None, blank=False):
        x = ws.cell(row=r, column=c, value=(None if blank else ("—" if val in (None, "") else val)))
        x.font = font(9, bold=bold, color=color, italic=italic)
        x.fill = fill(base)
        x.alignment = wrap(h=h)
        x.border = BORDER
        if fmt and (isinstance(val, (int, float)) or (isinstance(val, str) and val.startswith("="))):
            x.number_format = fmt
        return x

    def footer(ws, row, ncols, label):
        c = ws.cell(row=row, column=1, value=f"MigrationOS™  ·  {label}  ·  Source Data Analysis {VERSION}  ·  {LINK}")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        c.font = font(8, color=MUTED)

    def validation(ws, options, rng):
        dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(rng)

    wb = openpyxl.Workbook()
    ordered = sorted(tables.values(), key=lambda t: t.name)
    findings = build_findings(tables)

    # ------------------------------------------------------------ Engagement Intake
    ws_in = wb.create_sheet("Engagement Intake")
    IN_COLS = [("#", 5), ("Topic", 22), ("Question", 48), ("Why it matters", 52), ("Default / recommendation", 32),
               ("Answer", 14), ("Decision detail", 40), ("Decided by", 16), ("Date", 12), ("Status", 11), ("Drives", 36)]
    masthead(ws_in, f"{client}  ·  Engagement Intake", len(IN_COLS))
    note_row(ws_in, "Settle these with the client before mapping. Yellow cells are yours. Rows marked ⚑ re-scope tables live: "
                    "answer No and every table they drive flips to Not Included on the Source Data Analysis tab; Yes flips them to Included.", len(IN_COLS))
    header_row(ws_in, 3, IN_COLS)
    intake_first = 4
    topic_cell = {}
    gate_by_class, gate_by_table = {}, {}
    for idx, item in enumerate(profile.intake):
        r = intake_first + idx
        base = ZEBRA if idx % 2 else WHITE
        gates = item.get("gates_classes") or item.get("gates_tables")
        cell(ws_in, r, 1, idx + 1, base, h="center")
        cell(ws_in, r, 2, ("⚑ " if gates else "") + item["topic"], base, bold=True, color=TEAL_INK)
        cell(ws_in, r, 3, item.get("question", ""), base)
        cell(ws_in, r, 4, item.get("why", ""), base, color=MUTED)
        cell(ws_in, r, 5, item.get("default", ""), base, italic=True)
        for c in (6, 7, 8, 9):
            cell(ws_in, r, c, "", YELLOW, blank=True)
        cell(ws_in, r, 10, "Open", AMBER_TINT, bold=True, color=AMBER, h="center")
        cell(ws_in, r, 11, item.get("drives", ""), base, color=MUTED)
        validation(ws_in, item.get("options") or INTAKE_ANSWER_OPTIONS, f"F{r}")
        ws_in.row_dimensions[r].height = 58
        ref = f"'Engagement Intake'!$F${r}"
        topic_cell[item["topic"].lower()] = ref
        for cls in item.get("gates_classes", []):
            gate_by_class[cls] = ref
        for tn in item.get("gates_tables", []):
            gate_by_table[tn.lower()] = ref
    intake_last = intake_first + len(profile.intake) - 1
    validation(ws_in, INTAKE_STATUS_OPTIONS, f"J{intake_first}:J{intake_last}")
    ws_in.conditional_formatting.add(f"J{intake_first}:J{intake_last}",
        CellIsRule(operator="equal", formula=['"Decided"'], fill=fill(TEAL_TINT), font=Font(name=BODY_FONT, size=9, bold=True, color=TEAL_INK)))
    ws_in.conditional_formatting.add(f"J{intake_first}:J{intake_last}",
        CellIsRule(operator="equal", formula=['"N/A"'], fill=fill(HAIRLINE), font=Font(name=BODY_FONT, size=9, color=MUTED)))
    ws_in.freeze_panes = "C4"
    footer(ws_in, intake_last + 2, len(IN_COLS), "Engagement Intake")

    # ------------------------------------------------------------ Findings
    ws_f = wb.create_sheet("Findings")
    F_COLS = [("#", 5), ("Severity", 10), ("Category", 12), ("Table", 28), ("Column", 24), ("Finding", 60), ("Rows Affected", 11),
              ("Suggested Action", 44), ("Resolution / Owner", 30)]
    masthead(ws_f, f"{client}  ·  Findings", len(F_COLS))
    note_row(ws_f, "Data-quality issues found while scanning, most severe first. Counts only — no values. "
                   "High = fix or re-export before mapping; Medium = decide during mapping; Low = know about it.", len(F_COLS))
    header_row(ws_f, 3, F_COLS)
    r = 4
    sev_style = {"High": (RED_TINT, RED), "Medium": (AMBER_TINT, AMBER), "Low": (HAIRLINE, MUTED)}
    for i, (sev, cat, tbl, col, text, rows, action) in enumerate(findings):
        base = ZEBRA if i % 2 else WHITE
        cell(ws_f, r, 1, i + 1, base, h="center")
        sf, sc = sev_style[sev]
        cell(ws_f, r, 2, sev, sf, bold=True, color=sc, h="center")
        cell(ws_f, r, 3, cat, base, color=MUTED)
        cell(ws_f, r, 4, tbl, base, bold=True, color=TEAL_INK)
        cell(ws_f, r, 5, col, base)
        cell(ws_f, r, 6, text, base)
        cell(ws_f, r, 7, rows, base, h="right", fmt="#,##0")
        cell(ws_f, r, 8, action, base, color=MUTED)
        cell(ws_f, r, 9, "", YELLOW, blank=True)
        ws_f.row_dimensions[r].height = 30 if len(text) < 110 else 44
        r += 1
    f_last = max(r - 1, 4)
    if findings:
        ws_f.auto_filter.ref = f"A3:{L(len(F_COLS))}{f_last}"
    else:
        cell(ws_f, 4, 1, "No findings — the extract is structurally clean.", WHITE, italic=True, color=MUTED)
        ws_f.merge_cells(start_row=4, start_column=1, end_row=4, end_column=len(F_COLS))
    ws_f.freeze_panes = "D4"
    footer(ws_f, f_last + 2, len(F_COLS), "Findings")

    # ------------------------------------------------------------ Source Data Analysis
    ws_sa = wb.create_sheet("Source Data Analysis")
    SA = [("Source Table", "table", 28), ("Table Classification", "class", 18), ("Why This Classification", "reason", 34),
          ("Table Description", "desc", 40), ("Primary Key", "pk", 18), ("# of Records", "total", 11),
          ("# Non-Deleted Records", "live", 12), ("Deletion Column", "delcol", 12), ("File Size", "size", 10),
          ("File Format", "format", 14), ("References (this table looks up to →)", "refs", 34),
          ("Referenced By (← tables that look up to this one)", "refby", 34)]
    if profile.has_ui:
        SA.append((profile.ui_label, "ui", 30))
    if profile.has_article:
        SA.append(("Reference Link", "article", 20))
    SA += [("Analyst Notes / Questions", "notes", 34), ("Target Salesforce Object", "target", 32),
           ("Included in Migration", "included", 15), ("Analysis Status", "astatus", 14), ("Client Review Status", "cstatus", 14),
           ("Client Signoff Name", "signoff", 16), ("Link to Mapping Workbook", "link", 18), ("Client Notes", "cnotes", 28)]
    C = {key: i + 1 for i, (_, key, _) in enumerate(SA)}
    masthead(ws_sa, f"{client}  ·  Source Data Analysis", len(SA))
    note_row(ws_sa, f"Extract: {extract}  ·  {len(ordered)} tables scanned  ·  Counts exclude soft-deleted rows where a deletion column exists  ·  "
                    "Description, Target Object and Included are DEFAULTS (profile or inferred) — confirm or change every row.", len(SA))
    header_row(ws_sa, 3, [(a, w) for a, _, w in SA])
    sa_first = 4
    for idx, t in enumerate(ordered):
        r = sa_first + idx
        base = ZEBRA if idx % 2 else WHITE
        pt = profile.tables.get(t.name, {})
        desc = pt.get("description") or ""
        inferred_desc = not desc
        if t.error:
            desc = f"PARSE ERROR: {t.error}"
        elif t.size == 0 or not t.headers:
            desc = "Empty file (no header row)"
        elif inferred_desc:
            desc = f"Not in profile — {len(t.headers)} columns" + (f", {t.live:,} live rows" if t.live else "")
        target_obj = pt.get("target") or CLASS_TARGET_DEFAULTS.get(t.classification, "TBD")
        refs = sorted({f"{fk['target']} (via {fk['column']})" for fk in t.fks if fk["target"] and fk["role"] != "User Reference"
                       and not fk["rel"].startswith(("Unresolved", "Ambiguous"))})
        users = [fk["column"] for fk in t.fks if fk["role"] == "User Reference"]
        if users:
            refs.append(f"Users ({', '.join(users)})")
        unres = [fk["column"] for fk in t.fks if fk["rel"].startswith(("Unresolved", "Ambiguous"))]
        if unres:
            refs.append(f"⚠ unresolved: {', '.join(unres)}")
        ref_by = sorted({f"{tbl} ({col})" for tbl, col in t.referenced_by})
        cell(ws_sa, r, C["table"], t.filename, base, bold=True, color=TEAL_INK)
        cf, ct = CLASS_COLORS.get(t.classification, (AMBER_TINT, AMBER))
        cell(ws_sa, r, C["class"], t.classification, cf, bold=True, color=ct, h="center")
        cell(ws_sa, r, C["reason"], t.class_reason, base, color=MUTED)
        cell(ws_sa, r, C["desc"], desc, base, italic=inferred_desc, color=(MUTED if inferred_desc else INK))
        pk_text = t.pk or (f"⚠ {t.pk_note}" if t.pk_note else ("none detected" if t.headers else ""))
        cell(ws_sa, r, C["pk"], pk_text, base, bold=bool(t.pk), color=(INK if t.pk else AMBER))
        cell(ws_sa, r, C["total"], "ERROR" if t.error else t.total, base, h="right", fmt="#,##0")
        cell(ws_sa, r, C["live"], "ERROR" if t.error else t.live, base, h="right", fmt="#,##0")
        cell(ws_sa, r, C["delcol"], t.del_col, base, color=MUTED)
        cell(ws_sa, r, C["size"], human_size(t.size), base, color=MUTED, h="right")
        cell(ws_sa, r, C["format"], f"{DELIM_NAMES.get(t.delim, repr(t.delim))} · {t.encoding_label}" if t.headers else "", base, color=MUTED)
        cell(ws_sa, r, C["refs"], "\n".join(refs), base)
        cell(ws_sa, r, C["refby"], "\n".join(ref_by), base)
        if "ui" in C:
            ui = pt.get("ui", "")
            cell(ws_sa, r, C["ui"], ui or "Not yet documented", base, color=(INK if ui else MUTED), italic=not ui)
        if "article" in C:
            art = pt.get("article", "")
            ac = cell(ws_sa, r, C["article"], "Reference article" if art else "", base, color=TEAL_MID)
            if art:
                ac.hyperlink = art
        cell(ws_sa, r, C["notes"], "", YELLOW, blank=True)
        cell(ws_sa, r, C["target"], target_obj, base, italic=True)
        gate = None
        gated_by = (pt.get("gated_by") or "").lower()
        if gated_by and gated_by in topic_cell:
            gate = topic_cell[gated_by]
        elif t.name in gate_by_table:
            gate = gate_by_table[t.name]
        elif t.classification in gate_by_class:
            gate = gate_by_class[t.classification]
        if t.headers and t.live == 0 and not t.error:
            inc = "Not Applicable"
        elif gate:
            inc = f'=IF({gate}="No","Not Included",IF({gate}="Yes","Included","To Be Reviewed"))'
        else:
            inc = pt.get("included") or ("Not Applicable" if not t.headers else "To Be Reviewed")
        cell(ws_sa, r, C["included"], inc, base, h="center")
        cell(ws_sa, r, C["astatus"], "To Be Reviewed", base, h="center")
        cell(ws_sa, r, C["cstatus"], "Not started", base, h="center")
        for k in ("signoff", "link", "cnotes"):
            cell(ws_sa, r, C[k], "", YELLOW, blank=True)
        ws_sa.row_dimensions[r].height = max(30, 14 * max(len(refs), len(ref_by), 2))
    sa_last = sa_first + len(ordered) - 1
    colL = {k: L(v) for k, v in C.items()}
    validation(ws_sa, INCLUDED_OPTIONS, f"{colL['included']}{sa_first}:{colL['included']}{sa_last}")
    validation(ws_sa, ANALYSIS_OPTIONS, f"{colL['astatus']}{sa_first}:{colL['astatus']}{sa_last}")
    validation(ws_sa, CLIENT_REVIEW_OPTIONS, f"{colL['cstatus']}{sa_first}:{colL['cstatus']}{sa_last}")
    for key, good, warn in (("included", "Included", "Not Included"), ("astatus", "Completed", "In progress"), ("cstatus", "Approved", "In review")):
        rng = f"{colL[key]}{sa_first}:{colL[key]}{sa_last}"
        ws_sa.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{good}"'], fill=fill(TEAL_TINT),
                                         font=Font(name=BODY_FONT, size=9, bold=True, color=TEAL_INK)))
        ws_sa.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{warn}"'], font=Font(name=BODY_FONT, size=9, bold=True, color=AMBER)))
        ws_sa.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Not Applicable"'], fill=fill(HAIRLINE),
                                         font=Font(name=BODY_FONT, size=9, color=MUTED)))
    ws_sa.freeze_panes = "C4"
    ws_sa.auto_filter.ref = f"A3:{L(len(SA))}{sa_last}"
    footer(ws_sa, sa_last + 2, len(SA), "Source Data Analysis")

    # ------------------------------------------------------------ Key Map
    ws_km = wb.create_sheet("Key Map")
    KM_COLS = [("From Table", 26), ("From Column", 24), ("Key Role", 14), ("Resolves To Table", 26), ("To Column", 12),
               ("Relationship", 18), ("How Determined", 26), ("Confidence", 10), ("Live Rows With a Value", 11), ("Rows Resolved", 11),
               ("Resolution %", 10), ("Rows Unresolved", 11), ("Notes / Open Questions", 60)]
    masthead(ws_km, f"{client}  ·  Key Map", len(KM_COLS))
    note_row(ws_km, "One row per key. Resolution % = share of live rows whose value exists in the target's primary key (deleted target rows count as resolved; see Notes). "
                    "Below 95% is amber. Confidence: High = name + data agree; Medium = partial match or inferred from values; Low = ambiguous or weak.", len(KM_COLS))
    header_row(ws_km, 3, KM_COLS)
    r = 4
    km_rows = [(t, fk) for t in ordered for fk in sorted(t.fks, key=lambda f: (f["role"], f["column"]))]
    conf_style = {"High": (TEAL_TINT, TEAL_INK), "Medium": (AMBER_TINT, AMBER), "Low": (RED_TINT, RED), "—": (None, MUTED)}
    for idx, (t, fk) in enumerate(km_rows):
        base = ZEBRA if idx % 2 else WHITE
        bad = fk["rel"].startswith(("Unresolved", "Ambiguous"))
        cell(ws_km, r, 1, t.filename, base, bold=True, color=TEAL_INK)
        cell(ws_km, r, 2, fk["column"], base, bold=True)
        cell(ws_km, r, 3, fk["role"], base, color=MUTED)
        cell(ws_km, r, 4, fk["target"], base, bold=not bad, color=(AMBER if bad else INK))
        cell(ws_km, r, 5, fk["target_col"], base, color=MUTED, h="center")
        cell(ws_km, r, 6, fk["rel"], base, color=(AMBER if bad else INK))
        cell(ws_km, r, 7, fk["how"], base, color=MUTED)
        cf, cc = conf_style.get(fk["conf"], (None, MUTED))
        cell(ws_km, r, 8, fk["conf"], cf or base, bold=fk["conf"] != "—", color=cc, h="center")
        cell(ws_km, r, 9, fk["live_rows"], base, h="right", fmt="#,##0")
        if fk["pct"] is None:
            for c in (10, 11, 12):
                cell(ws_km, r, c, "", base, h="right")
        else:
            cell(ws_km, r, 10, fk["resolved"], base, h="right", fmt="#,##0")
            cell(ws_km, r, 11, f'=IF(I{r}=0,"",J{r}/I{r})', base, h="right", fmt="0.0%")
            cell(ws_km, r, 12, f"=I{r}-J{r}", base, h="right", fmt="#,##0")
        flag = bad or "confirm" in fk["note"].lower() or "OPEN" in fk["note"].upper() or "TIE" in fk["note"]
        cell(ws_km, r, 13, fk["note"], base, color=(AMBER if flag else MUTED))
        ws_km.row_dimensions[r].height = 30 if len(fk["note"]) < 120 else 48
        r += 1
    km_last = max(r - 1, 4)
    if km_rows:
        ws_km.conditional_formatting.add(f"K4:K{km_last}", CellIsRule(operator="lessThan", formula=["0.95"], fill=fill(AMBER_TINT),
                                         font=Font(name=BODY_FONT, size=9, bold=True, color=AMBER)))
        ws_km.auto_filter.ref = f"A3:{L(len(KM_COLS))}{km_last}"
    ws_km.freeze_panes = "C4"
    footer(ws_km, km_last + 2, len(KM_COLS), "Key Map")

    # ------------------------------------------------------------ Schema
    ws_sc = wb.create_sheet("Schema")
    SC_COLS = [("Table Name", 28), ("Column #", 8), ("Column Name", 28), ("Column Role", 16), ("Inferred Type", 16), ("Max Length", 9),
               ("Resolves To", 28), ("Resolution % (live rows)", 11), ("# Populated (live rows)", 11), ("% Populated (live rows)", 11),
               ("Distinct Values (live rows)", 11), ("Notes", 44)]
    masthead(ws_sc, f"{client}  ·  Schema", len(SC_COLS))
    note_row(ws_sc, f"Structure only — no data values. Populated % is measured on non-deleted rows (blank, null, 0, false and no count as not populated). "
                    f"Inferred Type uses up to {TYPE_SAMPLE:,} values per column; Distinct Values stops counting at {DISTINCT_CAP:,}.", len(SC_COLS))
    header_row(ws_sc, 3, SC_COLS)
    r, zi = 4, 0
    for t in ordered:
        fk_by_col = defaultdict(list)
        for fk in t.fks:
            fk_by_col[fk["column"].split(" [")[0]].append(fk)
        for pos, col in enumerate(t.headers, 1):
            base = ZEBRA if zi % 2 else WHITE
            zi += 1
            role = column_role(t, col)
            ctype = t.col_type(col)
            resolves, pct, notes = "", None, []
            if col in fk_by_col:
                es = fk_by_col[col]
                if role == "Polymorphic Key":
                    resolves = " / ".join(f"{e['target']} [{e['column'].split('[')[1].rstrip(']')}]" for e in es)
                    tot = sum(e["live_rows"] for e in es if e["pct"] is not None)
                    hit = sum(e["resolved"] for e in es if e["pct"] is not None)
                    pct = hit / tot if tot else None
                else:
                    e = es[0]
                    resolves = f"{e['target']}.{e['target_col']}" if e["target_col"] else (e["target"] or e["rel"])
                    pct = e["pct"]
                    if e["rel"].startswith(("Unresolved", "Ambiguous")):
                        notes.append(e["note"])
            elif role == "Polymorphic Type":
                tv = t.type_values.get(col, Counter())
                resolves = "class name selecting the parent table"
                if tv:
                    notes.append("Values: " + ", ".join(f"{k} ({v:,})" for k, v in sorted(tv.items())[:20]))
            elif role == "Deletion Flag":
                notes.append("Used to compute Non-Deleted counts")
            if role == "Primary Key":
                notes.append(f"Unique across all {t.total:,} rows")
            dc = t.distinct_count(col)
            if t.maxlen[col] > 255:
                notes.append(f"Over 255 characters → Long Text Area")
            if t.multiline[col]:
                notes.append(f"Line breaks in {t.multiline[col]:,} value(s)")
            if t.leading_zero[col] and ctype == "Integer":
                notes.append("Leading zeros — keep as Text")
            if (dc is not None and ctype == "Text" and 1 < dc <= PICKLIST_MAX and t.populated[col] >= 20 and role == "Attribute"
                    and dc / t.populated[col] <= 0.1 and t.maxlen[col] <= 60 and not t.multiline[col]
                    and not set(tokenize(col)) & NOT_PICKLIST_TOKENS):
                notes.append(f"Picklist candidate ({dc} distinct values)")
            if role == "Attribute" and t.populated[col] == 0 and t.live:
                notes.append("Never populated")
            cell(ws_sc, r, 1, t.filename, base, color=TEAL_INK, bold=(pos == 1))
            cell(ws_sc, r, 2, pos, base, h="center")
            cell(ws_sc, r, 3, col, base, bold=(role in ("Primary Key", "Foreign Key", "Polymorphic Key")))
            cell(ws_sc, r, 4, role, base, color=(AMBER if role == "User Reference" else (MUTED if role in ("Attribute", "Audit / System") else TEAL_INK)))
            cell(ws_sc, r, 5, ctype, base, color=(AMBER if ctype.startswith("Mixed") else INK))
            cell(ws_sc, r, 6, t.maxlen[col] or "", base, h="right", fmt="#,##0")
            cell(ws_sc, r, 7, resolves, base, color=(AMBER if notes and ("No table" in notes[0] or "Ambiguous" in notes[0]) else INK))
            cell(ws_sc, r, 8, pct if pct is not None else "", base, h="right", fmt="0.0%")
            cell(ws_sc, r, 9, t.populated.get(col, 0), base, h="right", fmt="#,##0")
            if t.live:
                cell(ws_sc, r, 10, f'=IF({t.live}=0,"",I{r}/{t.live})', base, h="right", fmt="0.0%")
            else:
                cell(ws_sc, r, 10, "", base, h="right")
            cell(ws_sc, r, 11, (f"{DISTINCT_CAP:,}+" if dc is None else dc), base, h="right", fmt="#,##0")
            cell(ws_sc, r, 12, "; ".join(notes), base, color=MUTED)
            r += 1
    sc_last = max(r - 1, 4)
    ws_sc.conditional_formatting.add(f"H4:H{sc_last}", CellIsRule(operator="lessThan", formula=["0.95"], fill=fill(AMBER_TINT),
                                     font=Font(name=BODY_FONT, size=9, bold=True, color=AMBER)))
    ws_sc.freeze_panes = "D4"
    ws_sc.auto_filter.ref = f"A3:{L(len(SC_COLS))}{sc_last}"
    footer(ws_sc, sc_last + 2, len(SC_COLS), "Schema")

    # ------------------------------------------------------------ Cover
    ws = wb.active
    ws.title = "Cover"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCD", (3, 42, 72, 3)):
        ws.column_dimensions[col].width = w
    for col in range(1, 5):
        ws.cell(row=1, column=col).fill = fill(DEEP_TEAL)
    ws.cell(row=1, column=2, value="MigrationOS™").font = font(14, bold=True, color=WHITE, name=DISPLAY_FONT)
    ws.cell(row=1, column=2).alignment = Alignment(vertical="center")
    c = ws.cell(row=1, column=3, value="Step 1 · Source Data Analysis")
    c.font = font(14, color=LIGHT_TEAL, name=DISPLAY_FONT)
    c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[1].height = 30
    ws.cell(row=3, column=2, value=f"{client} — Source Data Analysis").font = font(20, bold=True, name=DISPLAY_FONT)
    ws.merge_cells("B3:C3")
    ws.row_dimensions[3].height = 30
    meta = [("Client", client), ("Source system", source), ("Target", target), ("Extract", extract), ("Extract folder", folder),
            ("Source profile", (f"{profile.name} ({os.path.basename(profile.path)})"
                                + (" — auto-detected" if getattr(profile, "auto_detected", False) else "") if profile.path else "None — generic conventions + empirical checks")),
            ("Generated", date.today().strftime("%B %d, %Y")), ("Workbook version", f"Source Data Analysis {VERSION}")]
    r = 5
    for label, val in meta:
        ws.cell(row=r, column=2, value=label).font = font(9, bold=True, color=DEEP_TEAL)
        vc = ws.cell(row=r, column=3, value=val)
        vc.font = font(10)
        vc.alignment = wrap()
        r += 1
    r += 1
    ws.cell(row=r, column=2, value="AT A GLANCE").font = font(9, bold=True, color=MUTED)
    r += 1
    sa = "'Source Data Analysis'"
    inc = colL["included"]
    glance = [
        ("Tables in extract", f"=COUNTA({sa}!$A${sa_first}:$A${sa_last})"),
        ("Tables with live rows", f'=COUNTIF({sa}!${colL["live"]}${sa_first}:${colL["live"]}${sa_last},">0")'),
        ("Total live rows (all tables)", f'=SUM({sa}!${colL["live"]}${sa_first}:${colL["live"]}${sa_last})'),
        ("Tables currently Included", f'=COUNTIF({sa}!${inc}${sa_first}:${inc}${sa_last},"Included")'),
        ("Tables To Be Reviewed", f'=COUNTIF({sa}!${inc}${sa_first}:${inc}${sa_last},"To Be Reviewed")'),
        ("Tables Not Applicable / Not Included", f'=COUNTIF({sa}!${inc}${sa_first}:${inc}${sa_last},"Not Applicable")+COUNTIF({sa}!${inc}${sa_first}:${inc}${sa_last},"Not Included")'),
        ("Tables with no detected primary key", f'=COUNTIF({sa}!${colL["pk"]}${sa_first}:${colL["pk"]}${sa_last},"none detected")+COUNTIF({sa}!${colL["pk"]}${sa_first}:${colL["pk"]}${sa_last},"⚠*")'),
        ("Keys traced", f"=COUNTA('Key Map'!$B$4:$B${km_last})"),
        ("Keys resolving below 95%", f"=COUNTIF('Key Map'!$K$4:$K${km_last},\"<0.95\")"),
        ("Keys unresolved or ambiguous", f"=COUNTIF('Key Map'!$F$4:$F${km_last},\"Unresolved*\")+COUNTIF('Key Map'!$F$4:$F${km_last},\"Ambiguous*\")"),
        ("High-severity findings", f"=COUNTIF(Findings!$B$4:$B${f_last},\"High\")"),
        ("Engagement Intake items still Open", f"=COUNTIF('Engagement Intake'!$J${intake_first}:$J${intake_last},\"Open\")"),
    ]
    for i, (label, formula) in enumerate(glance):
        base = ZEBRA if i % 2 else WHITE
        lc = ws.cell(row=r, column=2, value=label); lc.font = font(10, bold=True, color=TEAL_INK); lc.fill = fill(base); lc.border = BORDER
        vc = ws.cell(row=r, column=3, value=formula); vc.font = font(10); vc.fill = fill(base); vc.border = BORDER
        vc.number_format = "#,##0"; vc.alignment = Alignment(horizontal="left")
        r += 1
    r += 1
    ws.cell(row=r, column=2, value="TABLE CLASSIFICATION LEGEND").font = font(9, bold=True, color=MUTED)
    r += 1
    legend = [
        ("Entity", "An anchor object with many attributes that other tables hang off. Migrates as a Salesforce object."),
        ("Child / Detail", "A parent key plus its own attributes. Usually folds into the parent's fields or a related object."),
        ("Lookup / Picklist", "Key + label rows referenced by other tables. Becomes picklist values, a Campaign, or a resolved text value."),
        ("Junction (many-to-many)", "Two or more keys and almost nothing else. Becomes a relationship object."),
        ("Polymorphic Child", "One table serving several parents via a *_type / *_id pair. Split by type before mapping."),
        ("Audit / History", "Logs, audit trails and change history. Usually archived, not migrated."),
        ("Staging / Backup", "Copies, backups, temp or dated tables. Confirm none is the live table, then exclude."),
        ("System / Config", "Application plumbing (permissions, settings, integrations). Rebuilt, never loaded."),
        ("Unclassified", "Structure matched no pattern — open the Schema tab and decide by hand."),
    ]
    for name, desc in legend:
        cf, ct = CLASS_COLORS[name]
        nc = ws.cell(row=r, column=2, value=name); nc.font = font(9, bold=True, color=ct); nc.fill = fill(cf); nc.border = BORDER
        nc.alignment = Alignment(horizontal="center", vertical="center")
        dc = ws.cell(row=r, column=3, value=desc); dc.font = font(9); dc.alignment = wrap(); dc.border = BORDER
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1
    ws.cell(row=r, column=2, value="HOW TO USE").font = font(9, bold=True, color=MUTED)
    r += 1
    how = ("1. Findings — start here. High items (broken files, duplicate keys) need a re-export or a client answer before anything else.\n"
           "2. Engagement Intake — settle every row with the client. ⚑ rows re-scope tables automatically.\n"
           "3. Source Data Analysis — one row per table. Check the Primary Key, References and Referenced By, then confirm Target Salesforce Object and Included "
           "in Migration and track Analysis and Client Review status through sign-off.\n"
           "4. Key Map — every key with how it was found, a confidence rating and its empirical resolution %. Low-confidence and amber rows need a decision.\n"
           "5. Schema — every column's role, inferred type, max length, populated % and distinct-value count. Use them to size fields and pick custom fields.\n"
           "Privacy: structure and counts only. No data values from the extract appear anywhere in this workbook.")
    hc = ws.cell(row=r, column=2, value=how)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    hc.font = font(9); hc.alignment = wrap(); hc.fill = fill(TEAL_TINT)
    ws.row_dimensions[r].height = 170
    r += 2
    fc = ws.cell(row=r, column=2, value=f"MigrationOS™ Template System  ·  Source Data Analysis {VERSION}  ·  {LINK}")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    fc.font = font(8, color=MUTED)

    wb.move_sheet("Cover", offset=-len(wb.sheetnames) + 1)
    wb.save(out_path)
    return findings


# ===========================================================================

def main():
    ap = argparse.ArgumentParser(description="MigrationOS Source Data Analysis — any CSV extract (Step 1).")
    ap.add_argument("folder", nargs="?", help="Folder containing the CSV export")
    ap.add_argument("-o", "--output", default=None, help="Output .xlsx (default: <Client>_Source_Data_Analysis.xlsx beside the folder)")
    ap.add_argument("--client", default="Client", help="Client name for the workbook header")
    ap.add_argument("--extract", default=None, help="Extract label (default: folder name)")
    ap.add_argument("--source", default=None, help="Source system name (default: from profile, else 'Source database (CSV extract)')")
    ap.add_argument("--target", default=None, help="Target system (default: from profile, else 'Salesforce')")
    ap.add_argument("--profile", default=None, help="Source profile JSON with system-specific knowledge")
    ap.add_argument("--no-auto-profile", action="store_true",
                    help="Do not auto-detect a profile (e.g. Salesforce Data Export) from the extract's tables")
    ap.add_argument("--delimiter", default=None, help="Force a delimiter (default: detected per file). Use 'tab' for tabs.")
    ap.add_argument("--encoding", default=None, help="Force an encoding (default: detected; UTF-8 with Latin-1 fallback)")
    ap.add_argument("--recursive", action="store_true", help="Also scan sub-folders")
    ap.add_argument("--write-profile-template", metavar="PATH", help="Write a blank profile JSON and exit")
    args = ap.parse_args()

    if args.write_profile_template:
        with open(args.write_profile_template, "w", encoding="utf-8") as f:
            json.dump(PROFILE_TEMPLATE, f, indent=2, ensure_ascii=False)
        print(f"Profile template written to: {args.write_profile_template}")
        return
    if not args.folder:
        ap.error("folder is required")
    folder = os.path.expanduser(args.folder)
    if not os.path.isdir(folder):
        sys.exit(f"ERROR: not a folder: {folder}")
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        sys.exit("ERROR: openpyxl is not installed.\nRun:  python3 -m pip install openpyxl")

    profile = Profile()
    if args.profile:
        with open(os.path.expanduser(args.profile), encoding="utf-8") as f:
            profile = Profile(json.load(f), path=args.profile)
    delim = {"tab": "\t", "\\t": "\t", "comma": ",", "semicolon": ";", "pipe": "|"}.get(args.delimiter, args.delimiter)

    raise_field_limit()
    paths = []
    if args.recursive:
        for root, _, names in os.walk(folder):
            paths += [(n, os.path.join(root, n)) for n in names if n.lower().endswith(FILE_EXTENSIONS)]
    else:
        paths = [(n, os.path.join(folder, n)) for n in os.listdir(folder) if n.lower().endswith(FILE_EXTENSIONS)]
    paths.sort(key=lambda p: p[0].lower())
    if not paths:
        sys.exit(f"ERROR: no {', '.join(FILE_EXTENSIONS)} files in {folder}")
    auto_detected = False
    if not args.profile and not args.no_auto_profile:
        data, fp = detect_profile(paths)
        if data:
            profile = Profile(data, path=fp)
            auto_detected = True
            print(f"Detected source: {profile.name} — applying {os.path.basename(fp)} (use --no-auto-profile to skip)")
    profile.auto_detected = auto_detected
    NULL_TOKENS.update(profile.null_tokens)
    print(f"Scanning {len(paths)} files in: {folder}" + (f"  (profile: {profile.name})" if profile.path else "") + "\n")

    tables = {}
    for i, (name, path) in enumerate(paths, 1):
        t = Table(name, path, profile, delim, args.encoding)
        if t.name in tables:                     # same table name in two sub-folders
            t.name = f"{t.name}~{i}"
        try:
            t.scan()
        except Exception as exc:                 # noqa: BLE001 — report and continue
            t.error = f"{type(exc).__name__}: {exc}"
        t.choose_pk()
        tables[t.name] = t
        print(f"  [{i:>3}/{len(paths)}] {name:<40} rows={t.total:<8} live={t.live:<8} cols={len(t.headers):<3} "
              f"pk={t.pk or '-':<14} {human_size(t.size)}" + (f"  ERROR {t.error}" if t.error else ""))

    resolve_keys(tables, profile)
    classify(tables, profile)

    extract = args.extract or os.path.basename(os.path.normpath(folder))
    source = args.source or profile.source_system or "Source database (CSV extract)"
    target = args.target or profile.target_system or "Salesforce"
    out = args.output or f"{args.client.replace(' ', '_')}_Source_Data_Analysis.xlsx"
    if not os.path.isabs(out):
        out = os.path.join(os.path.dirname(os.path.abspath(folder)), out)
    findings = build_workbook(tables, folder, args.client, extract, source, target, profile, out)

    fks = [fk for t in tables.values() for fk in t.fks]
    low = [fk for fk in fks if fk["pct"] is not None and fk["pct"] < 0.95]
    unres = [(t.name, fk["column"], fk["rel"]) for t in tables.values() for fk in t.fks if fk["rel"].startswith(("Unresolved", "Ambiguous"))]
    print("\nClassification: " + ", ".join(f"{k}={v}" for k, v in sorted(Counter(t.classification for t in tables.values()).items())))
    print(f"Keys traced: {len(fks)}   below 95%: {len(low)}   unresolved/ambiguous: {len(unres)}")
    for name, col, rel in unres:
        print(f"  {rel.split(' —')[0].upper():<11} {name}.{col}")
    sev = Counter(f[0] for f in findings)
    print(f"Findings: High={sev['High']}  Medium={sev['Medium']}  Low={sev['Low']}")
    print(f"\nWorkbook written to: {out}")


if __name__ == "__main__":
    main()
