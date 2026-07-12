"""
BPCL Uran - Document Control Index (DCI) Dashboard
Now backed by Supabase (PostgreSQL + Storage) for persistence.

Run with:  streamlit run dashboard.py

Required Streamlit secrets (see secrets.toml.example):
    SUPABASE_URL
    SUPABASE_SERVICE_ROLE_KEY
"""

import hashlib
import io
from datetime import datetime, date

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ----------------------------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------------------------
st.set_page_config(page_title="DCI Dashboard - BPCL Uran", layout="wide")

REV_BLOCK_START_COLS = [19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59]  # LATEST REV. col per block

APPROVED_CODES = {"1", "1 WITH COMMENTS"}
RETAINED_CODES = {"R"}
CODE2_CODES = {"2"}
CODE3_CODES = {"3"}
VOID_CODES = {"V"}

MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

AGING_BUCKET_ORDER = ["Not due yet", "0-7 days", "8-14 days", "15-21 days", ">21 days"]

CHART_PALETTE = ["#B4A7D6", "#F1948A", "#8EE4D0", "#f5c26b", "#8ab4f8", "#c9a0dc"]

STORAGE_BUCKET = "dci-workbooks"

# Original Excel-derived fields (unchanged shape expected by all the existing
# analytical tabs) plus the new dashboard-managed / editable fields, so a
# document loaded from Supabase looks exactly like a document parsed from Excel.
RECORD_COLUMNS = [
    "sr_no", "doc_no", "title", "discipline", "poc_submission",
    "latest_rev", "latest_uploaded_date", "eil_review_status", "latest_code",
    "latest_from_eil_date", "category", "l4_marked", "schedule_submission_date",
    "status", "poc", "eil_mandays", "eil_manhours", "stspl_mandays",
    "stspl_manhours", "revision_cycle", "expected_manhours",
    # -- new, dashboard-managed fields --
    "remarks", "assigned_engineer", "action_owner", "expected_completion_date",
    "forecast_month", "approval_date",
    # -- linkage fields (not shown in charts, used for Supabase round-trips) --
    "document_key", "document_id", "original_row_number",
]

HISTORY_COLUMNS = ["doc_no", "rev_no", "uploaded_date", "code", "from_eil_date"]

NUMERIC_COLUMNS = [
    "eil_mandays", "eil_manhours", "stspl_mandays",
    "stspl_manhours", "revision_cycle", "expected_manhours",
]

# Maps documents-table (Supabase) column names <-> dashboard RECORD_COLUMNS names.
DOC_TO_RECORD_FIELD = {
    "document_number": "doc_no",
    "document_title": "title",
    "discipline": "discipline",
    "vendor": "poc",
    "revision": "latest_rev",
    "dci_status": "status",
    "returned_code": "latest_code",
    "planned_submission_date": "schedule_submission_date",
    "actual_submission_date": "latest_uploaded_date",
    "review_date": "latest_from_eil_date",
    "approval_date": "approval_date",
    "forecast_month": "forecast_month",
    "remarks": "remarks",
    "assigned_engineer": "assigned_engineer",
    "action_owner": "action_owner",
    "expected_completion_date": "expected_completion_date",
}
STATUS_OPTIONS = [
    "FRESH SUBMISSION", "WITH EIL", "RETURNED BY EIL", "RESUBMISSION",
    "APPROVED", "APPROVED WITH COMMENTS", "ON HOLD", "CANCELLED",
]
RETURNED_CODE_OPTIONS = ["", "1", "1 WITH COMMENTS", "2", "3", "R", "V"]

ROOT_CAUSE_NOTES = [
    "Change in Tank Anchor Plate Design",
    "Dome Roof Rafter Section & Numbers and Compression plate thickness",
    "Suspended deck Design and Instrument Nozzle and Stillwell Size",
    "Intank Pump Selection and its document finalization",
    "Product Bottom Inlet Pipe Size reduction and Vapour Disengagement Design",
    "VMS Selection and its document finalization",
    "Tank Shell Insulation Arrangement change",
    "Tank Cool Down Sensor Arrangement change",
    "Change in Nos. of PSV & VSV and Change in their Platform arrangement",
    "Change in the Pump Roof Top maintenance platform arrangement",
    "Change in ROV arrangement at Liquid Platform and Multilevel Pipe Rack",
    "BOG line Orientation and Supporting on Shell and Trestle Design",
]


# ----------------------------------------------------------------------------
# SUPABASE CLIENT
# ----------------------------------------------------------------------------
@st.cache_resource
def get_supabase():
    from supabase import create_client

    url = st.secrets.get("SUPABASE_URL")
    key = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        return None
    return create_client(url, key)



# ----------------------------------------------------------------------------
# EXCEL PARSING (used only at import time, not for direct dashboard reads)
# ----------------------------------------------------------------------------
def parse_workbook_bytes(file_bytes):
    """Parse the uploaded .xlsx into (records, history_rows, project_name, contractor).
    Kept separate from Supabase concerns so it can be unit tested / reused.
    """
    import openpyxl

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if "DCI" not in wb.sheetnames:
        raise ValueError("This workbook does not contain a sheet named 'DCI'.")
    ws = wb["DCI"]

    project_name = ws.cell(row=2, column=2).value
    contractor = ws.cell(row=3, column=2).value

    records, history = [], []

    for r in range(9, ws.max_row + 1):
        doc_no = ws.cell(row=r, column=2).value
        if doc_no is None or str(doc_no).strip() == "":
            continue

        rec = {
            "sr_no": ws.cell(row=r, column=1).value,
            "doc_no": str(doc_no).strip(),
            "title": ws.cell(row=r, column=3).value,
            "discipline": ws.cell(row=r, column=4).value,
            "poc_submission": ws.cell(row=r, column=5).value,
            "latest_rev": ws.cell(row=r, column=10).value,
            "latest_uploaded_date": ws.cell(row=r, column=11).value,
            "eil_review_status": ws.cell(row=r, column=12).value,
            "latest_code": ws.cell(row=r, column=14).value,
            "latest_from_eil_date": ws.cell(row=r, column=15).value,
            "category": ws.cell(row=r, column=16).value,
            "l4_marked": ws.cell(row=r, column=17).value,
            "schedule_submission_date": ws.cell(row=r, column=18).value,
            "status": ws.cell(row=r, column=66).value,
            "poc": ws.cell(row=r, column=76).value,
            "eil_mandays": ws.cell(row=r, column=77).value,
            "eil_manhours": ws.cell(row=r, column=78).value,
            "stspl_mandays": ws.cell(row=r, column=79).value,
            "stspl_manhours": ws.cell(row=r, column=80).value,
            "revision_cycle": ws.cell(row=r, column=81).value,
            "expected_manhours": ws.cell(row=r, column=82).value,
            "remarks": "",
            "assigned_engineer": "",
            "action_owner": "",
            "expected_completion_date": None,
            "forecast_month": "",
            "approval_date": None,
            "original_row_number": r,
        }
        rec["document_key"] = f"{rec['doc_no']}::ROW-{r}"
        records.append(rec)

        row_history = []
        for start in REV_BLOCK_START_COLS:
            rev_no = ws.cell(row=r, column=start).value
            uploaded = ws.cell(row=r, column=start + 1).value
            code = ws.cell(row=r, column=start + 2).value
            from_eil = ws.cell(row=r, column=start + 3).value
            if rev_no is None and uploaded is None and code is None:
                continue
            row_history.append(
                {
                    "doc_no": rec["doc_no"],
                    "rev_no": rev_no,
                    "uploaded_date": _jsonable(uploaded),
                    "code": code,
                    "from_eil_date": _jsonable(from_eil),
                }
            )
        rec["_history_rows"] = row_history
        history.extend(row_history)

    return records, history, project_name, contractor


def _jsonable(value):
    """Make an openpyxl cell value safe to store in JSONB (dates -> isoformat)."""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return value


def validate_records(records):
    """Return (warnings, errors) lists for the import preview."""
    warnings, errors = [], []
    if not records:
        errors.append("No document rows were found under the 'DCI' sheet (starting row 9).")
        return warnings, errors

    doc_nos = [r["doc_no"] for r in records]
    seen, dupes = set(), set()
    for d in doc_nos:
        if d in seen:
            dupes.add(d)
        seen.add(d)
    if dupes:
        warnings.append(f"{len(dupes)} duplicate document number(s) found: {', '.join(list(dupes)[:10])}"
                         + (" ..." if len(dupes) > 10 else ""))

    missing_titles = sum(1 for r in records if not r.get("title"))
    if missing_titles:
        warnings.append(f"{missing_titles} row(s) are missing a document title.")

    return warnings, errors


def workbook_hash(file_bytes):
    return hashlib.sha256(file_bytes).hexdigest()


# ----------------------------------------------------------------------------
# IMPORT WORKFLOW (Supabase writes)
# ----------------------------------------------------------------------------
def find_import_by_hash(supabase, file_hash):
    res = (
        supabase.table("imports")
        .select("*")
        .eq("workbook_hash", file_hash)
        .eq("is_active", True)
        .order("uploaded_at", desc=True)
        .limit(1)
        .execute()
    )
    return res.data[0] if res.data else None


def list_imports(supabase):
    res = (
        supabase.table("imports")
        .select("*")
        .eq("is_active", True)
        .order("uploaded_at", desc=True)
        .execute()
    )
    return res.data or []


def upload_workbook_to_storage(supabase, import_id, file_bytes, filename):
    path = f"imports/{import_id}/original.xlsx"
    supabase.storage.from_(STORAGE_BUCKET).upload(
        path=path,
        file=file_bytes,
        file_options={
            "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "upsert": "true",
        },
    )
    return path


def run_import(supabase, file_bytes, filename, project_name, contractor,
                review_period_days, version_mode, parent_import_id=None):
    """Import a workbook without deactivating the current version prematurely.

    Supabase/PostgREST calls are not a single database transaction from Python, so
    this function uses a safe publish sequence: create pending import, upload and
    populate it, mark it complete, and only then archive the replaced import.
    """
    records, _history, wb_project_name, wb_contractor = parse_workbook_bytes(file_bytes)
    warnings, errors = validate_records(records)
    if errors:
        raise ValueError("; ".join(errors))

    file_hash = workbook_hash(file_bytes)
    version_number = 1
    if version_mode == "new_version" and parent_import_id:
        parent = (
            supabase.table("imports")
            .select("version_number")
            .eq("id", parent_import_id)
            .limit(1)
            .execute()
        )
        if parent.data:
            version_number = int(parent.data[0].get("version_number") or 1) + 1

    import_row = {
        "original_filename": filename,
        "display_name": filename,
        "project_name": project_name or wb_project_name,
        "contractor": contractor or wb_contractor,
        "version_number": version_number,
        "parent_import_id": parent_import_id if version_mode == "new_version" else None,
        "source_sheet_name": "DCI",
        "row_count": len(records),
        "workbook_hash": file_hash,
        "storage_bucket": STORAGE_BUCKET,
        "original_storage_path": "pending",
        "review_period_days": int(review_period_days),
        "validation_status": "warnings" if warnings else "clean",
        "import_status": "pending",
        "warning_count": len(warnings),
        "error_count": 0,
        "is_active": True,
    }
    created = supabase.table("imports").insert(import_row).execute()
    if not created.data:
        raise RuntimeError("Supabase did not return the newly created import row.")
    import_id = created.data[0]["id"]

    storage_path = None
    try:
        storage_path = upload_workbook_to_storage(supabase, import_id, file_bytes, filename)
        supabase.table("imports").update({"original_storage_path": storage_path}).eq("id", import_id).execute()
        _bulk_import_documents(supabase, import_id, records)
        supabase.table("imports").update({"import_status": "complete"}).eq("id", import_id).execute()

        if version_mode == "replace" and parent_import_id:
            supabase.table("imports").update({
                "is_active": False,
                "import_status": "archived",
            }).eq("id", parent_import_id).execute()

        st.cache_data.clear()
        return import_id, warnings
    except Exception:
        # Keep the previous active import untouched. Mark this import failed so it
        # cannot be selected as a valid workbook. Cascading cleanup can be done by
        # an admin without destroying evidence needed to diagnose the failure.
        supabase.table("imports").update({
            "is_active": False,
            "import_status": "failed",
        }).eq("id", import_id).execute()
        raise


def _bulk_import_documents(supabase, import_id, records, batch_size=200):
    for batch_start in range(0, len(records), batch_size):
        batch = records[batch_start:batch_start + batch_size]

        doc_rows = []
        for rec in batch:
            doc_rows.append({
                "import_id": import_id,
                "document_key": rec["document_key"],
                "original_row_number": rec["original_row_number"],
                "document_number": rec["doc_no"],
                "document_title": rec["title"],
                "discipline": rec["discipline"],
                "vendor": rec["poc"],
                "revision": None if rec["latest_rev"] is None else str(rec["latest_rev"]),
                "dci_status": rec["status"],
                "returned_code": None if rec["latest_code"] is None else str(rec["latest_code"]),
                "planned_submission_date": _jsonable(rec["schedule_submission_date"]),
                "actual_submission_date": _jsonable(rec["latest_uploaded_date"]),
                "review_date": _jsonable(rec["latest_from_eil_date"]),
                "approval_date": None,
                "forecast_month": "",
                "remarks": "",
                "assigned_engineer": "",
                "action_owner": "",
                "expected_completion_date": None,
                "last_changed_by": "import",
                "is_active": True,
            })

        inserted = supabase.table("documents").insert(doc_rows).execute()
        inserted_rows = inserted.data or []
        if len(inserted_rows) != len(doc_rows):
            raise RuntimeError(
                f"Inserted {len(inserted_rows)} of {len(doc_rows)} documents in a batch."
            )
        inserted_by_key = {row["document_key"]: row for row in inserted_rows}

        raw_rows, history_rows = [], []
        for rec in batch:
            doc = inserted_by_key.get(rec["document_key"])
            if doc is None:
                raise RuntimeError(f"Could not resolve inserted document: {rec['document_key']}")
            record_json = {k: _jsonable(v) for k, v in rec.items() if not k.startswith("_")}
            raw_rows.append({
                "document_id": doc["id"],
                "row_data": {"record": record_json, "history_rows": rec["_history_rows"]},
                "original_column_names": list(record_json.keys()),
                "original_row_number": rec["original_row_number"],
            })
            history_rows.append({
                "document_id": doc["id"],
                "import_id": import_id,
                "event_type": "imported",
                "field_name": None,
                "old_value": None,
                "new_value": None,
                "revision": doc.get("revision"),
                "returned_code": doc.get("returned_code"),
                "status": doc.get("dci_status"),
                "remarks": None,
                "changed_by": "import",
                "source": "import",
            })

        if raw_rows:
            supabase.table("document_raw_data").insert(raw_rows).execute()
        if history_rows:
            supabase.table("document_history").insert(history_rows).execute()


# ----------------------------------------------------------------------------
# REPOSITORY — load documents back out of Supabase into the shapes the
# existing analytical code (derive_fields / derive_history) already expects.
# ----------------------------------------------------------------------------
def _fetch_all(query_builder, page_size=1000):
    """Page through a Supabase query using .range() since PostgREST caps at 1000 rows."""
    rows, start = [], 0
    while True:
        res = query_builder.range(start, start + page_size - 1).execute()
        chunk = res.data or []
        rows.extend(chunk)
        if len(chunk) < page_size:
            break
        start += page_size
    return rows


@st.cache_data(ttl=60, show_spinner=False)
def load_documents(_supabase, import_id):
    """Returns (raw_df, raw_hist) shaped exactly like the original Excel parse,
    reconstructed from Supabase (raw JSON snapshot + current editable fields).
    The leading underscore on _supabase tells st.cache_data not to hash the client.
    """
    doc_rows = _fetch_all(_supabase.table("documents").select("*").eq("import_id", import_id).eq("is_active", True))
    if not doc_rows:
        return pd.DataFrame(columns=RECORD_COLUMNS), pd.DataFrame(columns=HISTORY_COLUMNS)

    doc_ids = [d["id"] for d in doc_rows]
    raw_rows = []
    for i in range(0, len(doc_ids), 200):
        chunk_ids = doc_ids[i:i + 200]
        res = _supabase.table("document_raw_data").select("*").in_("document_id", chunk_ids).execute()
        raw_rows.extend(res.data or [])
    raw_by_doc_id = {r["document_id"]: r for r in raw_rows}

    records, history = [], []
    for doc in doc_rows:
        raw = raw_by_doc_id.get(doc["id"], {})
        row_data = raw.get("row_data") or {}
        record = dict(row_data.get("record") or {})

        # Overlay the current (possibly edited) values from the documents table,
        # since those are the source of truth after any dashboard edit.
        for doc_field, record_field in DOC_TO_RECORD_FIELD.items():
            if doc_field in doc:
                record[record_field] = doc[doc_field]

        record["document_key"] = doc["document_key"]
        record["document_id"] = doc["id"]
        record["original_row_number"] = doc["original_row_number"]
        records.append(record)

        for h in (row_data.get("history_rows") or []):
            history.append(h)

    df = pd.DataFrame(records)
    for col in RECORD_COLUMNS:
        if col not in df.columns:
            df[col] = None
    df = df[RECORD_COLUMNS]

    hist = pd.DataFrame(history)
    for col in HISTORY_COLUMNS:
        if col not in hist.columns:
            hist[col] = None
    hist = hist[HISTORY_COLUMNS] if not hist.empty else pd.DataFrame(columns=HISTORY_COLUMNS)

    return df, hist


def get_import_meta(supabase, import_id):
    res = supabase.table("imports").select("*").eq("id", import_id).single().execute()
    return res.data


def get_document_history(supabase, document_id):
    res = (
        supabase.table("document_history")
        .select("*")
        .eq("document_id", document_id)
        .order("changed_at", desc=False)
        .execute()
    )
    return res.data or []


def get_history_for_documents(supabase, document_ids, batch_size=200):
    """Fetch history for many documents in a handful of batched requests
    (one .in_() query per batch_size ids) instead of one request per document.
    Use this instead of looping get_document_history() over a whole workbook —
    with hundreds of documents that loop is slow and prone to timing out.
    """
    document_ids = [d for d in document_ids if d is not None and not (isinstance(d, float) and pd.isna(d))]
    rows = []
    for i in range(0, len(document_ids), batch_size):
        chunk = document_ids[i:i + batch_size]
        res = (
            supabase.table("document_history")
            .select("*")
            .in_("document_id", chunk)
            .order("changed_at", desc=False)
            .execute()
        )
        rows.extend(res.data or [])
    return rows


DATE_RECORD_FIELDS = {
    "schedule_submission_date", "latest_uploaded_date", "latest_from_eil_date",
    "approval_date", "expected_completion_date",
}


def _normalise_record_value(field, value):
    """Normalize form/database scalars so previews and saves compare identically."""
    if field in DATE_RECORD_FIELDS:
        return _date_to_iso(value)
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def build_change_diff(form_values, current_record):
    diff = []
    for record_field, new_value in form_values.items():
        old_value = current_record.get(record_field)
        new_norm = _normalise_record_value(record_field, new_value)
        old_norm = _normalise_record_value(record_field, old_value)
        if new_norm != old_norm:
            diff.append({
                "field": record_field,
                "old": old_norm,
                "new": new_norm,
            })
    return diff


def save_document_changes(supabase, document_id, form_values, current_record, changed_by, event_type):
    """Persist changed fields atomically through the database RPC."""
    record_diff = build_change_diff(form_values, current_record)
    if not record_diff:
        return False, []

    changes = {}
    old_values = {}
    for item in record_diff:
        doc_field = next(
            (db_field for db_field, record_field in DOC_TO_RECORD_FIELD.items()
             if record_field == item["field"]),
            None,
        )
        if doc_field is None:
            continue
        changes[doc_field] = item["new"] or None
        old_values[doc_field] = item["old"] or None

    if not changes:
        return False, []

    supabase.rpc("fn_save_document_changes", {
        "p_document_id": document_id,
        "p_changes": changes,
        "p_old_values": old_values,
        "p_event_type": event_type,
        "p_changed_by": changed_by or "dashboard-user",
    }).execute()

    st.cache_data.clear()
    display_diff = [
        {
            "field": item["field"],
            "old": item["old"] or "(blank)",
            "new": item["new"] or "(blank)",
        }
        for item in record_diff
    ]
    return True, display_diff


def _date_to_iso(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, pd.Timestamp):
        if pd.isna(val):
            return None
        return val.date().isoformat()
    if isinstance(val, (date, datetime)):
        return val.isoformat()[:10]
    if isinstance(val, str):
        text = val.strip()
        if not text:
            return None
        parsed = pd.to_datetime(text, errors="coerce")
        return None if pd.isna(parsed) else parsed.date().isoformat()
    return None


# ----------------------------------------------------------------------------
# EXPORT — regenerate an .xlsx from the original workbook + current DB values
# ----------------------------------------------------------------------------
EXPORT_CELL_MAP = {
    # documents-table field -> (column index in the original DCI sheet)
    "document_title": 3,
    "revision": 10,
    "actual_submission_date": 11,
    "dci_status": 66,
    "returned_code": 14,
    "review_date": 15,
    "planned_submission_date": 18,
}
# Fields with no home in the original layout get appended as new columns.
EXTRA_EXPORT_COLUMNS = ["remarks", "assigned_engineer", "action_owner",
                         "expected_completion_date", "forecast_month", "approval_date"]
EXTRA_EXPORT_HEADERS = {
    "remarks": "Remarks (Dashboard)",
    "assigned_engineer": "Assigned Engineer (Dashboard)",
    "action_owner": "Action Owner (Dashboard)",
    "expected_completion_date": "Expected Completion Date (Dashboard)",
    "forecast_month": "Forecast Month (Dashboard)",
    "approval_date": "Approval Date (Dashboard)",
}


def generate_updated_workbook(supabase, import_id):
    import openpyxl

    meta = get_import_meta(supabase, import_id)
    original_bytes = supabase.storage.from_(STORAGE_BUCKET).download(meta["original_storage_path"])
    wb = openpyxl.load_workbook(io.BytesIO(original_bytes), data_only=False)
    ws = wb["DCI"]

    doc_rows = _fetch_all(supabase.table("documents").select("*").eq("import_id", import_id).eq("is_active", True))

    extra_start_col = ws.max_column + 1
    header_row = 8
    for i, field in enumerate(EXTRA_EXPORT_COLUMNS):
        ws.cell(row=header_row, column=extra_start_col + i, value=EXTRA_EXPORT_HEADERS[field])

    for doc in doc_rows:
        r = doc["original_row_number"]
        for field, col in EXPORT_CELL_MAP.items():
            ws.cell(row=r, column=col, value=doc.get(field))
        for i, field in enumerate(EXTRA_EXPORT_COLUMNS):
            ws.cell(row=r, column=extra_start_col + i, value=doc.get(field))

    # History sheet
    history_sheet_name = "Dashboard Change History"
    if history_sheet_name in wb.sheetnames:
        del wb[history_sheet_name]
    hs = wb.create_sheet(history_sheet_name)
    headers = ["Document Number", "Document Title", "Event Type", "Field Changed", "Old Value",
               "New Value", "Revision", "Returned Code", "Status", "Remarks", "Changed By", "Changed At"]
    hs.append(headers)

    doc_lookup = {d["id"]: d for d in doc_rows}
    all_history = []
    for i in range(0, len(doc_rows), 200):
        chunk_ids = [d["id"] for d in doc_rows[i:i + 200]]
        res = supabase.table("document_history").select("*").in_("document_id", chunk_ids) \
            .order("changed_at").execute()
        all_history.extend(res.data or [])

    for h in all_history:
        doc = doc_lookup.get(h["document_id"], {})
        hs.append([
            doc.get("document_number"), doc.get("document_title"), h.get("event_type"),
            h.get("field_name"), h.get("old_value"), h.get("new_value"), h.get("revision"),
            h.get("returned_code"), h.get("status"), h.get("remarks"), h.get("changed_by"),
            h.get("changed_at"),
        ])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"{(meta.get('project_name') or 'DCI').replace(' ', '_')}_DCI_Updated_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
    return buf.getvalue(), filename


def record_export(supabase, import_id, export_type, filename, row_count, exported_by="dashboard-user"):
    supabase.table("export_records").insert({
        "import_id": import_id,
        "export_type": export_type,
        "exported_by": exported_by,
        "filename": filename,
        "row_count": row_count,
        "success": True,
    }).execute()


# ----------------------------------------------------------------------------
# DERIVED-FIELD CALCULATIONS (unchanged from the original dashboard)
# ----------------------------------------------------------------------------
def clean_dates(df, cols):
    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NaT
        else:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


def normalize_code(code):
    if pd.isna(code):
        return ""
    if isinstance(code, (int, float)) and float(code).is_integer():
        return str(int(code))
    return " ".join(str(code).strip().upper().split())


def classify_code(code):
    value = normalize_code(code)
    if value in APPROVED_CODES:
        return "Approved (Code 1)"
    if value in RETAINED_CODES:
        return "Retained for Record (Code R)"
    if value in CODE2_CODES:
        return "Code 2 - Comments/Resubmit"
    if value in CODE3_CODES:
        return "Code 3 - Rejected/Resubmit"
    if value in VOID_CODES:
        return "Void"
    return "No code yet"


def normalize_status(value):
    if pd.isna(value):
        return ""
    status = " ".join(str(value).strip().upper().split())
    aliases = {
        "WITH EIL FOR REVIEW": "WITH EIL",
        "PENDING WITH EIL": "WITH EIL",
        "FRESH SUBMISSION PENDING": "FRESH SUBMISSION",
        "CODE 1": "APPROVED",
    }
    return aliases.get(status, status)


def status_fallback(row):
    status = normalize_status(row.get("status"))
    if status:
        return status
    if pd.isna(row.get("latest_uploaded_date")):
        return "FRESH SUBMISSION"
    if row["code_bucket"] in ("Approved (Code 1)", "Retained for Record (Code R)"):
        return "APPROVED"
    if pd.notna(row.get("latest_from_eil_date")):
        return "RETURNED BY EIL"
    return "WITH EIL"


def aging_bucket(age):
    if pd.isna(age):
        return None
    if age < 0:
        return "Not due yet"
    if age <= 7:
        return "0-7 days"
    if age <= 14:
        return "8-14 days"
    if age <= 21:
        return "15-21 days"
    return ">21 days"


def derive_fields(df, review_period_days, today):
    df = df.copy()
    df = clean_dates(df, ["latest_uploaded_date", "latest_from_eil_date", "schedule_submission_date"])

    for col in NUMERIC_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["code_bucket"] = df["latest_code"].apply(classify_code)
    df["is_void"] = df["code_bucket"] == "Void"

    df["status_clean"] = df.apply(status_fallback, axis=1)

    df["is_submitted"] = df["latest_uploaded_date"].notna() | (df["status_clean"] != "FRESH SUBMISSION")
    df["is_reviewed"] = df["latest_from_eil_date"].notna()
    df["is_approved_final"] = df["code_bucket"].isin(
        ["Approved (Code 1)", "Retained for Record (Code R)"]
    )

    df["submission_delay_days"] = np.where(
        df["latest_uploaded_date"].isna() & df["schedule_submission_date"].notna(),
        (today - df["schedule_submission_date"]).dt.days,
        np.nan,
    )
    df["review_age_days"] = np.where(
        df["latest_uploaded_date"].notna()
        & df["latest_from_eil_date"].isna()
        & ~df["is_approved_final"]
        & ~df["is_void"],
        (today - df["latest_uploaded_date"]).dt.days,
        np.nan,
    )
    df["is_submission_overdue"] = df["submission_delay_days"] > 0
    df["is_review_overdue"] = df["review_age_days"] > review_period_days

    def compute_age(row):
        if row["is_void"] or row["is_approved_final"]:
            return np.nan
        if row["status_clean"] == "FRESH SUBMISSION":
            base = row["schedule_submission_date"]
        else:
            base = row["latest_uploaded_date"]
        if pd.isna(base):
            return np.nan
        return (today - base).days

    df["age_days"] = df.apply(compute_age, axis=1)
    df["aging_bucket"] = df["age_days"].apply(aging_bucket)
    df["is_overdue"] = df["is_submission_overdue"].fillna(False) | df["is_review_overdue"].fillna(False)

    df["turnaround_days"] = (df["latest_from_eil_date"] - df["latest_uploaded_date"]).dt.days

    for c in ["latest_uploaded_date", "latest_from_eil_date", "schedule_submission_date"]:
        df[c + "_month"] = df[c].dt.to_period("M").astype(str)

    df["latest_rev_display"] = df["latest_rev"].apply(
        lambda v: "" if pd.isna(v) else str(v).strip()
    )

    return df


def derive_history(hist):
    hist = hist.copy()
    hist = clean_dates(hist, ["uploaded_date", "from_eil_date"])
    hist["code_bucket"] = hist["code"].apply(classify_code)
    hist["turnaround_days"] = (hist["from_eil_date"] - hist["uploaded_date"]).dt.days

    invalid_turnaround = hist["turnaround_days"] < 0
    invalid_count = int(invalid_turnaround.sum())
    if invalid_count:
        hist.loc[invalid_turnaround, "turnaround_days"] = np.nan
    st.session_state["_invalid_turnaround_count"] = invalid_count

    hist["uploaded_month"] = hist["uploaded_date"].dt.to_period("M").astype(str)
    hist["returned_month"] = hist["from_eil_date"].dt.to_period("M").astype(str)
    return hist


# ----------------------------------------------------------------------------
# SHARED CHART HELPER
# ----------------------------------------------------------------------------
def render_monthly_grouped_chart(data, date_col, title, count_label="Count"):
    d = data.dropna(subset=[date_col]).copy()
    if d.empty:
        st.info(f"No data available for '{title}' with the current filters.")
        return

    d["year"] = d[date_col].dt.year
    d["month_num"] = d[date_col].dt.month

    years = sorted(d["year"].unique(), reverse=True)
    grid_index = pd.MultiIndex.from_product([years, range(1, 13)], names=["year", "month_num"])
    grid = (
        d.groupby(["year", "month_num"]).size()
        .reindex(grid_index, fill_value=0)
        .reset_index(name=count_label)
    )
    grid["month_name"] = grid["month_num"].apply(lambda m: MONTH_ORDER[m - 1])
    grid["year"] = grid["year"].astype(str)

    year_labels = [str(y) for y in years]
    color_map = {y: CHART_PALETTE[i % len(CHART_PALETTE)] for i, y in enumerate(year_labels)}

    with st.container(border=True):
        fig = px.bar(
            grid,
            x="month_name",
            y=count_label,
            color="year",
            barmode="group",
            category_orders={"month_name": MONTH_ORDER, "year": year_labels},
            color_discrete_map=color_map,
        )
        fig.update_traces(marker_line_color="rgba(255,255,255,0.6)", marker_line_width=1)
        fig.update_layout(
            title=dict(
                text=title,
                x=0.5,
                xanchor="center",
                font=dict(family="Arial Black, Arial, sans-serif", size=22, color="#595959"),
            ),
            yaxis_title="No. of Deliverables",
            xaxis_title="",
            yaxis=dict(gridcolor="#e6e6e6", zeroline=False),
            xaxis=dict(showgrid=False),
            plot_bgcolor="white",
            paper_bgcolor="white",
            bargap=0.28,
            bargroupgap=0.04,
            legend=dict(
                orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5,
                title_text="", font=dict(size=13),
            ),
            margin=dict(t=70, b=10, l=10, r=10),
        )
        st.plotly_chart(fig, use_container_width=True)

        header_cells = "".join(
            f"<th style='padding:5px 10px;border:1px solid #d9d9d9;text-align:center;'>{m}</th>"
            for m in MONTH_ORDER
        )
        row_html = []
        for y in year_labels:
            color = color_map[y]
            series = grid.loc[grid["year"] == y].set_index("month_name").reindex(MONTH_ORDER)[count_label]
            value_cells = "".join(
                f"<td style='padding:5px 10px;border:1px solid #d9d9d9;text-align:center;'>{int(v)}</td>"
                for v in series
            )
            row_html.append(
                "<tr>"
                f"<td style='padding:5px 10px;border:1px solid #d9d9d9;white-space:nowrap;'>"
                f"<span style='display:inline-block;width:10px;height:10px;background:{color};"
                f"border:1px solid #999;margin-right:6px;'></span>{y}</td>"
                f"{value_cells}</tr>"
            )
        table_html = f"""
        <table style='border-collapse:collapse;width:100%;font-size:14px;color:#333;margin-top:4px;'>
            <tr><th style='padding:5px 10px;border:1px solid #d9d9d9;'></th>{header_cells}</tr>
            {''.join(row_html)}
        </table>
        """
        st.markdown(table_html, unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# SUPABASE CONNECTION GATE
# ----------------------------------------------------------------------------
supabase = get_supabase()
if supabase is None:
    st.title("BPCL Uran - Document Control Index Dashboard")
    st.error("Supabase is not configured yet.")
    st.markdown(
        "Add these to your Streamlit secrets (`.streamlit/secrets.toml` locally, "
        "or **Settings → Secrets** on Streamlit Community Cloud) and reload:\n\n"
        "```toml\n"
        'SUPABASE_URL = "https://your-project.supabase.co"\n'
        'SUPABASE_SERVICE_ROLE_KEY = "your-service-role-key"\n'
        "```\n\n"
        "Also run `supabase_schema.sql` once in your Supabase project's SQL editor "
        "(Table Editor → SQL Editor) to create the required tables, and create a "
        f"**private** Storage bucket named `{STORAGE_BUCKET}`."
    )
    st.stop()

# ----------------------------------------------------------------------------
# SIDEBAR
# ----------------------------------------------------------------------------
st.sidebar.title("DCI Dashboard")

with st.sidebar.expander("Workbook", expanded=True):
    imports = list_imports(supabase)
    import_options = {f"{i['display_name']} (v{i['version_number']}, {i['row_count']} docs)": i["id"] for i in imports}

    if imports:
        selected_label = st.selectbox("Active workbook", list(import_options.keys()))
        active_import_id = import_options[selected_label]
        active_import = next(i for i in imports if i["id"] == active_import_id)
        st.caption(
            f"Uploaded: {active_import['uploaded_at'][:10]}  \n"
            f"Status: {active_import['import_status']}  \n"
            f"Rows: {active_import['row_count']}"
        )
    else:
        active_import_id = None
        active_import = None
        st.info("No workbook imported yet. Upload one in the **Import / Export** tab.")

    new_upload = st.file_uploader("Upload a new DCI workbook (.xlsx)", type=["xlsx"], key="sidebar_uploader")
    if new_upload is not None:
        st.session_state["pending_upload_bytes"] = new_upload.getvalue()
        st.session_state["pending_upload_name"] = new_upload.name
        st.info("Go to the **Import / Export** tab to validate and confirm this upload.")

st.sidebar.markdown("---")
with st.sidebar.expander("Project Settings", expanded=True):
    default_review_days = active_import["review_period_days"] if active_import else 21
    review_period_days = st.number_input(
        "Allowed EIL review period (days)", min_value=1, max_value=90, value=default_review_days
    )

today = pd.Timestamp(datetime.today().date())

if active_import_id is not None:
    raw_df, raw_hist = load_documents(supabase, active_import_id)
else:
    raw_df, raw_hist = pd.DataFrame(columns=RECORD_COLUMNS), pd.DataFrame(columns=HISTORY_COLUMNS)

has_data = not raw_df.empty

if has_data:
    df = derive_fields(raw_df, review_period_days, today)
    hist = derive_history(raw_hist)
    ever_reviewed_docs = set(hist.loc[hist["from_eil_date"].notna(), "doc_no"])
    df["ever_reviewed"] = df["doc_no"].isin(ever_reviewed_docs)
else:
    df = pd.DataFrame(columns=RECORD_COLUMNS)
    hist = pd.DataFrame(columns=HISTORY_COLUMNS)

st.sidebar.markdown("---")
st.sidebar.subheader("Filters")

if has_data:
    disciplines = sorted([d for d in df["discipline"].dropna().unique()])
    vendors = sorted([v for v in df["poc"].dropna().unique()])
    statuses = sorted([s for s in df["status_clean"].dropna().unique()])
    code_buckets = sorted([c for c in df["code_bucket"].dropna().unique()])
    revisions = sorted([r for r in df["latest_rev_display"].unique() if r != ""], key=str.casefold)
else:
    disciplines = vendors = statuses = code_buckets = revisions = []
aging_buckets = AGING_BUCKET_ORDER

f_discipline = st.sidebar.multiselect("Discipline", disciplines)
f_vendor = st.sidebar.multiselect("Vendor / Consultant (POC)", vendors)
f_status = st.sidebar.multiselect("Status", statuses)
f_code = st.sidebar.multiselect("Approval code", code_buckets)
f_revision = st.sidebar.multiselect("Revision", revisions)
f_aging = st.sidebar.multiselect("Aging bucket (open docs only)", aging_buckets)
search = st.sidebar.text_input("Search document number / title")

f_date_range, include_undated = None, True
if has_data:
    date_cols_available = df["latest_uploaded_date"].dropna()
    if not date_cols_available.empty:
        min_d, max_d = date_cols_available.min().date(), date_cols_available.max().date()
        f_date_range = st.sidebar.date_input(
            "Uploaded date range", value=(min_d, max_d), min_value=min_d, max_value=max_d
        )
        include_undated = st.sidebar.checkbox("Include documents without an uploaded date", value=True)

filtered = df.copy()
if has_data:
    if f_discipline:
        filtered = filtered[filtered["discipline"].isin(f_discipline)]
    if f_vendor:
        filtered = filtered[filtered["poc"].isin(f_vendor)]
    if f_status:
        filtered = filtered[filtered["status_clean"].isin(f_status)]
    if f_code:
        filtered = filtered[filtered["code_bucket"].isin(f_code)]
    if f_revision:
        filtered = filtered[filtered["latest_rev_display"].isin(f_revision)]
    if f_aging:
        filtered = filtered[filtered["aging_bucket"].isin(f_aging)]
    if search:
        s = search.lower()
        filtered = filtered[
            filtered["doc_no"].str.lower().str.contains(s, na=False)
            | filtered["title"].fillna("").str.lower().str.contains(s, na=False)
        ]
    if f_date_range and isinstance(f_date_range, tuple) and len(f_date_range) == 2:
        start_d, end_d = pd.Timestamp(f_date_range[0]), pd.Timestamp(f_date_range[1])
        date_match = filtered["latest_uploaded_date"].between(start_d, end_d)
        if include_undated:
            date_match = date_match | filtered["latest_uploaded_date"].isna()
        filtered = filtered[date_match]

filtered_doc_numbers = set(filtered["doc_no"]) if has_data else set()
filtered_hist = hist[hist["doc_no"].isin(filtered_doc_numbers)].copy() if has_data else hist

st.sidebar.markdown("---")
with st.sidebar.expander("Quick Statistics", expanded=False):
    if has_data:
        st.metric("Filtered Documents", len(filtered))
        st.metric("Approved", int((filtered["code_bucket"] == "Approved (Code 1)").sum()))
        st.metric("Returned", int((filtered["status_clean"] == "RETURNED BY EIL").sum()))
        st.metric("Delayed", int(filtered["is_overdue"].sum()))
        st.metric("Under Review", int((filtered["status_clean"] == "WITH EIL").sum()))
    else:
        st.caption("No data loaded yet.")

# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
project_name = active_import["project_name"] if active_import else None
contractor = active_import["contractor"] if active_import else None

if has_data:
    as_of_candidates = pd.concat([df["latest_uploaded_date"], df["latest_from_eil_date"]]).dropna()
    as_of_date = as_of_candidates.max().strftime("%d %b %Y") if not as_of_candidates.empty else "N/A"
else:
    as_of_date = "N/A"

st.title(project_name or "Document Control Index Dashboard")
st.caption(f"Contractor: {contractor or 'N/A'}  |  Data as of: {as_of_date}  |  Total documents: {len(df)}")

tabs = st.tabs(
    [
        "Executive Summary",
        "Submission Progress",
        "Review Progress",
        "Approval Status",
        "Discipline Analysis",
        "Vendor Analysis",
        "Delays & Aging",
        "Manhours",
        "Document Register",
        "Update Document",
        "Document Timeline",
        "Import / Export",
    ]
)

def build_document_selector_options(dataframe):
    """Return stable labels keyed by document_id, even when doc numbers repeat."""
    options = []
    duplicate_counts = dataframe["doc_no"].value_counts(dropna=False)
    for _, row in dataframe.sort_values(["doc_no", "original_row_number"], na_position="last").iterrows():
        doc_no = str(row.get("doc_no") or "(no document number)")
        label = doc_no
        if duplicate_counts.get(row.get("doc_no"), 0) > 1:
            label = f"{doc_no} — Excel row {row.get('original_row_number')}"
        options.append((label, row.get("document_id")))
    return options


try:
    # ------------------------------------------------------------------
    # TAB 1: EXECUTIVE SUMMARY
    # ------------------------------------------------------------------
    with tabs[0]:
        if not has_data:
            st.info("No workbook imported yet. Go to **Import / Export** to upload one.")
        else:
            total = len(filtered)
            submitted = int(filtered["is_submitted"].sum())
            reviewed = int(filtered["is_reviewed"].sum())
            approved1 = int((filtered["code_bucket"] == "Approved (Code 1)").sum())
            retained = int((filtered["code_bucket"] == "Retained for Record (Code R)").sum())
            under_review = int((filtered["status_clean"] == "WITH EIL").sum())
            pending_first_sub = int((filtered["status_clean"] == "FRESH SUBMISSION").sum())
            overdue = int(filtered["is_overdue"].sum())

            r1 = st.columns(5)
            r1[0].metric("Total Documents", total)
            r1[1].metric("Submitted", submitted, f"{submitted/total*100:.0f}%" if total else None)
            r1[2].metric("Reviewed (latest rev)", reviewed, f"{reviewed/total*100:.0f}%" if total else None)
            r1[3].metric("Approved (Code 1)", approved1, f"{approved1/total*100:.0f}%" if total else None)
            r1[4].metric("Retained for Record", retained)

            r2 = st.columns(5)
            r2[0].metric("Under Review (with EIL)", under_review)
            r2[1].metric("Pending First Submission", pending_first_sub)
            r2[2].metric("Code 2 (comments)", int((filtered["code_bucket"] == "Code 2 - Comments/Resubmit").sum()))
            r2[3].metric("Void", int(filtered["is_void"].sum()))
            r2[4].metric("Overdue (submission or review)", overdue, delta_color="inverse")

            st.markdown("---")
            c1, c2 = st.columns([2, 1])
            with c1:
                breakdown = filtered["code_bucket"].value_counts().reset_index()
                breakdown.columns = ["Code", "Count"]
                fig = px.pie(breakdown, names="Code", values="Count", title="Approval Status Breakdown", hole=0.4)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                st.subheader("Root Cause - Manhour Overshoot")
                st.caption("Carried over from the previous Excel dashboard (view-only).")
                for i, note in enumerate(ROOT_CAUSE_NOTES, 1):
                    st.markdown(f"**{i}.** {note}")

    # ------------------------------------------------------------------
    # TAB 2: SUBMISSION PROGRESS
    # ------------------------------------------------------------------
    with tabs[1]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Monthly Submission Chart")
            render_monthly_grouped_chart(filtered_hist, "uploaded_date", "Monthly Submission Chart", "Submissions")

            st.markdown("---")
            st.subheader("Planned vs Actual First Submission")

            first_submission = (
                filtered_hist.dropna(subset=["uploaded_date"])
                .groupby("doc_no", as_index=False)["uploaded_date"].min()
                .rename(columns={"uploaded_date": "first_uploaded_date"})
            )

            planned = filtered.dropna(subset=["schedule_submission_date"]).copy()
            planned["planned_month"] = planned["schedule_submission_date"].dt.to_period("M").astype(str)
            planned_counts = planned.groupby("planned_month").size().reset_index(name="Planned")

            actual = filtered.merge(first_submission, on="doc_no", how="left").dropna(subset=["first_uploaded_date"])
            actual["actual_month"] = actual["first_uploaded_date"].dt.to_period("M").astype(str)
            actual_counts = actual.groupby("actual_month").size().reset_index(name="Actual")

            merged = pd.merge(
                planned_counts, actual_counts, left_on="planned_month", right_on="actual_month", how="outer"
            )
            merged["month"] = merged["planned_month"].combine_first(merged["actual_month"])
            merged = merged[["month", "Planned", "Actual"]].fillna(0).sort_values("month")
            fig2 = px.line(
                merged, x="month", y=["Planned", "Actual"], markers=True,
                title="Planned vs Actual First Submission by Month",
            )
            st.plotly_chart(fig2, use_container_width=True)

            cum = merged.copy()
            cum["Planned Cum"] = cum["Planned"].cumsum()
            cum["Actual Cum"] = cum["Actual"].cumsum()
            fig3 = px.line(cum, x="month", y=["Planned Cum", "Actual Cum"], markers=True, title="Cumulative Submission Progress")
            st.plotly_chart(fig3, use_container_width=True)

    # ------------------------------------------------------------------
    # TAB 3: REVIEW PROGRESS
    # ------------------------------------------------------------------
    with tabs[2]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Monthly Review Chart")
            render_monthly_grouped_chart(filtered_hist, "from_eil_date", "Monthly Review Chart", "Reviews")

            st.markdown("---")
            st.subheader("Submission-to-Review Turnaround")
            invalid_count = st.session_state.get("_invalid_turnaround_count", 0)
            if invalid_count:
                st.warning(
                    f"{invalid_count} revision record(s) had an EIL return date earlier than their "
                    "upload date and were excluded from turnaround calculations."
                )
            turn = filtered_hist.dropna(subset=["turnaround_days"])
            if not turn.empty:
                avg_turn = turn.groupby("returned_month")["turnaround_days"].mean().reset_index()
                fig2 = px.line(avg_turn, x="returned_month", y="turnaround_days", markers=True,
                                title="Average Turnaround Time (days) by Month")
                st.plotly_chart(fig2, use_container_width=True)
                st.metric("Average turnaround (all revisions, days)", f"{turn['turnaround_days'].mean():.1f}")
            else:
                st.info("No turnaround data available yet.")

    # ------------------------------------------------------------------
    # TAB 4: APPROVAL STATUS
    # ------------------------------------------------------------------
    with tabs[3]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Approval Code Breakdown (current / latest revision)")
            code_counts = filtered["code_bucket"].value_counts().reset_index()
            code_counts.columns = ["Code", "Count"]
            fig = px.bar(code_counts, x="Code", y="Count", title="Documents by Approval Code", text="Count")
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Approval Code by Revision (full history)")
            rc = filtered_hist.dropna(subset=["rev_no"]).groupby(["rev_no", "code_bucket"]).size().reset_index(name="Count")
            fig2 = px.bar(rc, x="rev_no", y="Count", color="code_bucket", title="Code outcome at each revision number", barmode="stack")
            st.plotly_chart(fig2, use_container_width=True)

            st.subheader("Revision Distribution")
            rev_dist = (
                filtered.loc[filtered["latest_rev_display"] != "", "latest_rev_display"]
                .value_counts()
                .reset_index()
            )
            rev_dist.columns = ["Latest Revision", "Count"]
            rev_dist = rev_dist.sort_values("Latest Revision", key=lambda s: s.str.casefold())
            fig3 = px.bar(rev_dist, x="Latest Revision", y="Count", title="Documents by Current Revision Number")
            st.plotly_chart(fig3, use_container_width=True)

    # ------------------------------------------------------------------
    # TAB 5: DISCIPLINE ANALYSIS
    # ------------------------------------------------------------------
    with tabs[4]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Documents by Discipline")
            disc_counts = filtered["discipline"].value_counts().reset_index()
            disc_counts.columns = ["Discipline", "Count"]
            fig = px.bar(disc_counts.sort_values("Count"), x="Count", y="Discipline", orientation="h",
                         title="Total Documents by Discipline", height=600)
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Pending Documents by Discipline")
            pending = filtered[~filtered["is_approved_final"] & ~filtered["is_void"]]
            pend_disc = pending["discipline"].value_counts().reset_index()
            pend_disc.columns = ["Discipline", "Pending"]
            fig2 = px.bar(pend_disc.sort_values("Pending"), x="Pending", y="Discipline", orientation="h",
                          title="Pending Documents by Discipline", height=600)
            st.plotly_chart(fig2, use_container_width=True)

            st.subheader("Overdue Documents by Discipline")
            overdue_df = filtered[filtered["is_overdue"] == True]
            od_disc = overdue_df["discipline"].value_counts().reset_index()
            od_disc.columns = ["Discipline", "Overdue"]
            if not od_disc.empty:
                fig3 = px.bar(od_disc.sort_values("Overdue"), x="Overdue", y="Discipline", orientation="h",
                              title="Overdue Documents by Discipline", height=500)
                st.plotly_chart(fig3, use_container_width=True)
            else:
                st.info("No overdue documents for the current filter.")

    # ------------------------------------------------------------------
    # TAB 6: VENDOR ANALYSIS
    # ------------------------------------------------------------------
    with tabs[5]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Documents by Vendor / Consultant")
            vend_counts = filtered["poc"].value_counts().reset_index()
            vend_counts.columns = ["Vendor / POC", "Count"]
            fig = px.bar(vend_counts.sort_values("Count").tail(20), x="Count", y="Vendor / POC", orientation="h",
                         title="Documents by Vendor (Top 20)", height=600)
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Top Vendors with Pending Documents")
            pending = filtered[~filtered["is_approved_final"] & ~filtered["is_void"]]
            pend_vend = pending["poc"].value_counts().reset_index().head(15)
            pend_vend.columns = ["Vendor / POC", "Pending"]
            fig2 = px.bar(pend_vend.sort_values("Pending"), x="Pending", y="Vendor / POC", orientation="h",
                          title="Top 15 Vendors by Pending Documents", height=500)
            st.plotly_chart(fig2, use_container_width=True)

    # ------------------------------------------------------------------
    # TAB 7: DELAYS & AGING
    # ------------------------------------------------------------------
    with tabs[6]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            open_docs = filtered[~filtered["is_approved_final"] & ~filtered["is_void"]]
            st.subheader("Aging Buckets (open documents)")
            bucket_counts = open_docs["aging_bucket"].value_counts().reindex(AGING_BUCKET_ORDER).fillna(0).reset_index()
            bucket_counts.columns = ["Aging Bucket", "Count"]
            fig = px.bar(bucket_counts, x="Aging Bucket", y="Count", text="Count", title="Open Documents by Aging Bucket")
            st.plotly_chart(fig, use_container_width=True)

            m = st.columns(3)
            m[0].metric("Total open (pending) documents", len(open_docs))
            m[1].metric("Late First Submissions", int(filtered["is_submission_overdue"].fillna(False).sum()))
            m[2].metric(f"EIL Reviews Overdue (> {review_period_days} days)", int(filtered["is_review_overdue"].fillna(False).sum()))
            st.caption(
                "\"Late First Submissions\" compares today against the scheduled submission date for "
                "documents not yet uploaded. \"EIL Reviews Overdue\" compares today against the upload "
                "date for documents still awaiting EIL return."
            )

            st.subheader("Overdue Documents")
            overdue_table = open_docs[open_docs["is_overdue"] == True][
                ["doc_no", "title", "discipline", "poc", "status_clean", "age_days", "schedule_submission_date", "latest_uploaded_date"]
            ].sort_values("age_days", ascending=False).copy()
            for date_col in ["schedule_submission_date", "latest_uploaded_date"]:
                overdue_table[date_col] = overdue_table[date_col].dt.strftime("%d-%b-%Y").fillna("")
            st.dataframe(overdue_table, use_container_width=True, hide_index=True)

    # ------------------------------------------------------------------
    # TAB 8: MANHOURS
    # ------------------------------------------------------------------
    with tabs[7]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            total_expected = filtered["expected_manhours"].fillna(0).sum()
            total_actual = (filtered["eil_manhours"].fillna(0) + filtered["stspl_manhours"].fillna(0)).sum()
            variance = total_actual - total_expected

            c = st.columns(3)
            c[0].metric("Planned Manhours (Expected)", f"{total_expected:,.0f}")
            c[1].metric("Actual Manhours (EIL + STSPL)", f"{total_actual:,.0f}")
            c[2].metric("Variance", f"{variance:,.0f}", delta_color="inverse")

            st.subheader("Manhours by Discipline")
            mh_disc = filtered.groupby("discipline")[["expected_manhours", "eil_manhours", "stspl_manhours"]].sum().reset_index()
            mh_disc["actual_manhours"] = mh_disc["eil_manhours"].fillna(0) + mh_disc["stspl_manhours"].fillna(0)
            fig = px.bar(mh_disc, x="discipline", y=["expected_manhours", "actual_manhours"], barmode="group",
                         title="Planned vs Actual Manhours by Discipline")
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Manhours by Vendor")
            mh_vend = filtered.groupby("poc")[["expected_manhours", "eil_manhours", "stspl_manhours"]].sum().reset_index()
            mh_vend["actual_manhours"] = mh_vend["eil_manhours"].fillna(0) + mh_vend["stspl_manhours"].fillna(0)
            mh_vend = mh_vend.sort_values("actual_manhours", ascending=False).head(15)
            fig2 = px.bar(mh_vend, x="poc", y=["expected_manhours", "actual_manhours"], barmode="group",
                          title="Planned vs Actual Manhours - Top 15 Vendors")
            st.plotly_chart(fig2, use_container_width=True)

            st.markdown("---")
            st.subheader("Root Cause Analysis for Manhour Overshoot")
            st.caption("Carried over from the original Excel dashboard (view-only reference list).")
            for i, note in enumerate(ROOT_CAUSE_NOTES, 1):
                st.markdown(f"**{i}.** {note}")

    # ------------------------------------------------------------------
    # TAB 9: DOCUMENT REGISTER
    # ------------------------------------------------------------------
    with tabs[8]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader(f"Document Register ({len(filtered)} documents matching filters)")

            register_cols = [
                "doc_no", "title", "discipline", "poc", "latest_rev_display", "status_clean",
                "code_bucket", "schedule_submission_date", "latest_uploaded_date", "latest_from_eil_date",
                "age_days",
            ]
            register_df = filtered[register_cols].rename(columns={
                "doc_no": "Document Number", "title": "Document Title", "discipline": "Discipline",
                "poc": "Vendor", "latest_rev_display": "Current Revision", "status_clean": "Current Status",
                "code_bucket": "Approval Code", "schedule_submission_date": "Submission Date",
                "latest_uploaded_date": "Last Uploaded", "latest_from_eil_date": "Review Date",
                "age_days": "Aging Days",
            }).copy()
            for c in ["Submission Date", "Last Uploaded", "Review Date"]:
                register_df[c] = pd.to_datetime(register_df[c]).dt.strftime("%d-%b-%Y").fillna("")
            register_df = register_df.sort_values("Document Number")

            st.dataframe(register_df, use_container_width=True, hide_index=True, height=500)

            st.markdown("---")
            st.caption("Select a document below to jump to Update Document or Document Timeline.")
            pick_options = ["-- select --"] + sorted(filtered["doc_no"].dropna().unique().tolist())
            picked = st.selectbox("Open document", pick_options, key="register_pick")
            if picked != "-- select --":
                st.session_state["selected_doc_no"] = picked
                st.success(f"'{picked}' selected. Open the **Update Document** or **Document Timeline** tab.")

    # ------------------------------------------------------------------
    # TAB 10: UPDATE DOCUMENT
    # ------------------------------------------------------------------
    with tabs[9]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Update Document")
            selector_options = build_document_selector_options(df)
            option_labels = [label for label, _ in selector_options]
            id_by_label = dict(selector_options)
            default_idx = 0
            preselected = st.session_state.get("selected_doc_no")
            if preselected:
                matching = [i for i, label in enumerate(option_labels) if label == preselected or label.startswith(f"{preselected} —")]
                if matching:
                    default_idx = matching[0]

            selected_label = st.selectbox(
                "Document", option_labels, index=default_idx, key="update_doc_select"
            )
            document_id = id_by_label.get(selected_label)

            if document_id is not None:
                doc_row = df[df["document_id"] == document_id].iloc[0]
                selected_doc_no = doc_row["doc_no"]

                st.markdown("#### Current Information (read-only)")
                ro1, ro2, ro3 = st.columns(3)
                ro1.text_input("Document Number", value=doc_row["doc_no"], disabled=True)
                ro1.text_input("Discipline", value=doc_row["discipline"] or "", disabled=True)
                ro2.text_input("Document Title", value=doc_row["title"] or "", disabled=True)
                ro2.text_input("Vendor / POC", value=doc_row["poc"] or "", disabled=True)
                ro3.text_input("Original Row Number", value=str(doc_row["original_row_number"]), disabled=True)
                ro3.text_input("Stable Document Key", value=str(doc_row["document_key"]), disabled=True)

                st.markdown("#### Editable Fields")
                with st.form("document_update_form"):
                    fc1, fc2 = st.columns(2)
                    with fc1:
                        cur_status = doc_row["status"] if doc_row["status"] in STATUS_OPTIONS else STATUS_OPTIONS[0]
                        new_status = st.selectbox("DCI Status", STATUS_OPTIONS, index=STATUS_OPTIONS.index(cur_status))
                        new_revision = st.text_input("Revision", value="" if pd.isna(doc_row["latest_rev"]) else str(doc_row["latest_rev"]))
                        cur_code = normalize_code(doc_row["latest_code"])
                        code_idx = RETURNED_CODE_OPTIONS.index(cur_code) if cur_code in RETURNED_CODE_OPTIONS else 0
                        new_code = st.selectbox("Returned Code", RETURNED_CODE_OPTIONS, index=code_idx)
                        new_forecast_month = st.text_input("Forecast Month", value=doc_row["forecast_month"] or "")
                        new_assigned_engineer = st.text_input("Assigned Engineer", value=doc_row["assigned_engineer"] or "")
                        new_action_owner = st.text_input("Action Owner", value=doc_row["action_owner"] or "")
                    with fc2:
                        def _safe_date(v):
                            ts = pd.to_datetime(v, errors="coerce")
                            return None if pd.isna(ts) else ts.date()

                        new_planned = st.date_input("Planned Submission Date", value=_safe_date(doc_row["schedule_submission_date"]))
                        new_actual = st.date_input("Actual Submission Date", value=_safe_date(doc_row["latest_uploaded_date"]))
                        new_review = st.date_input("Review Date", value=_safe_date(doc_row["latest_from_eil_date"]))
                        new_approval = st.date_input("Approval Date", value=_safe_date(doc_row["approval_date"]))
                        new_expected_completion = st.date_input("Expected Completion Date", value=_safe_date(doc_row["expected_completion_date"]))

                    new_remarks = st.text_area("Remarks", value=doc_row["remarks"] or "")

                    preview_clicked = st.form_submit_button("Save Changes")

                if preview_clicked:
                    form_values = {
                        "status": new_status, "latest_rev": new_revision, "latest_code": new_code,
                        "forecast_month": new_forecast_month, "assigned_engineer": new_assigned_engineer,
                        "action_owner": new_action_owner, "schedule_submission_date": new_planned,
                        "latest_uploaded_date": new_actual, "latest_from_eil_date": new_review,
                        "approval_date": new_approval, "expected_completion_date": new_expected_completion,
                        "remarks": new_remarks,
                    }
                    st.session_state["pending_doc_changes"] = (document_id, form_values, doc_row.to_dict())

                pending = st.session_state.get("pending_doc_changes")
                if pending and pending[0] == document_id:
                    _, form_values, current_record = pending
                    changes_preview = [
                        {
                            "Field": item["field"],
                            "Current Value": item["old"] or "(blank)",
                            "New Value": item["new"] or "(blank)",
                        }
                        for item in build_change_diff(form_values, current_record)
                    ]

                    if not changes_preview:
                        st.info("No changes detected.")
                        st.session_state.pop("pending_doc_changes", None)
                    else:
                        st.markdown("#### Change Preview")
                        st.dataframe(pd.DataFrame(changes_preview), use_container_width=True, hide_index=True)
                        cc1, cc2 = st.columns([1, 4])
                        if cc1.button("Confirm Save", type="primary"):
                            changed, diff = save_document_changes(
                                supabase, document_id, form_values, current_record,
                                changed_by="dashboard-user", event_type="field_updated",
                            )
                            st.session_state.pop("pending_doc_changes", None)
                            if changed:
                                st.success(f"Saved {len(diff)} change(s) to {selected_doc_no}.")
                                st.rerun()
                            else:
                                st.info("No changes detected.")
                        if cc2.button("Discard"):
                            st.session_state.pop("pending_doc_changes", None)
                            st.rerun()

                st.markdown("---")
                st.markdown("#### Recent History")
                recent = get_document_history(supabase, document_id)
                if recent:
                    recent_df = pd.DataFrame(recent[-10:][::-1])
                    show_cols = [c for c in ["changed_at", "event_type", "field_name", "old_value", "new_value", "changed_by"] if c in recent_df.columns]
                    st.dataframe(recent_df[show_cols], use_container_width=True, hide_index=True)
                else:
                    st.caption("No history yet for this document.")

    # ------------------------------------------------------------------
    # TAB 11: DOCUMENT TIMELINE
    # ------------------------------------------------------------------
    with tabs[10]:
        if not has_data:
            st.info("No workbook imported yet.")
        else:
            st.subheader("Document Timeline")
            selector_options = build_document_selector_options(df)
            option_labels = [label for label, _ in selector_options]
            id_by_label = dict(selector_options)
            default_idx = 0
            preselected = st.session_state.get("selected_doc_no")
            if preselected:
                matching = [i for i, label in enumerate(option_labels) if label == preselected or label.startswith(f"{preselected} —")]
                if matching:
                    default_idx = matching[0]
            timeline_label = st.selectbox(
                "Document", option_labels, index=default_idx, key="timeline_doc_select"
            )
            document_id = id_by_label.get(timeline_label)

            if document_id is not None:
                doc_row = df[df["document_id"] == document_id].iloc[0]
                timeline_doc_no = doc_row["doc_no"]
                st.caption(f"{doc_row['title'] or ''}  |  Discipline: {doc_row['discipline'] or 'N/A'}  |  Vendor: {doc_row['poc'] or 'N/A'}")

                events = get_document_history(supabase, document_id)
                if not events:
                    st.info("No history events recorded for this document yet.")
                else:
                    for i, ev in enumerate(events):
                        label = ev["event_type"].replace("_", " ").title()
                        if ev.get("field_name"):
                            label += f" — {ev['field_name'].replace('_', ' ').title()}"
                        ts = ev.get("changed_at", "")
                        try:
                            ts_display = pd.to_datetime(ts).strftime("%d %b %Y, %H:%M")
                        except Exception:
                            ts_display = ts

                        marker = "●" if i < len(events) - 1 else "◉"
                        bar = "│" if i < len(events) - 1 else " "
                        st.markdown(f"**{marker} {label}**")
                        details = []
                        if ev.get("old_value") or ev.get("new_value"):
                            details.append(f"{ev.get('old_value') or '(blank)'} → {ev.get('new_value') or '(blank)'}")
                        if ev.get("revision"):
                            details.append(f"Rev {ev['revision']}")
                        if ev.get("status"):
                            details.append(f"Status: {ev['status']}")
                        if ev.get("changed_by"):
                            details.append(f"By: {ev['changed_by']}")
                        st.caption(f"{ts_display}" + ("  |  " + "  |  ".join(details) if details else ""))
                        st.markdown(bar)

                    st.markdown("---")
                    st.markdown("#### Full History Table")
                    st.dataframe(pd.DataFrame(events), use_container_width=True, hide_index=True)

    # ------------------------------------------------------------------
    # TAB 12: IMPORT / EXPORT
    # ------------------------------------------------------------------
    with tabs[11]:
        st.subheader("Import")

        pending_bytes = st.session_state.get("pending_upload_bytes")
        pending_name = st.session_state.get("pending_upload_name")

        tab_upload = st.file_uploader("Upload a DCI workbook (.xlsx)", type=["xlsx"], key="import_tab_uploader")
        if tab_upload is not None:
            pending_bytes = tab_upload.getvalue()
            pending_name = tab_upload.name
            st.session_state["pending_upload_bytes"] = pending_bytes
            st.session_state["pending_upload_name"] = pending_name

        if pending_bytes is None:
            st.caption("No pending upload. Use the uploader above or the sidebar.")
        else:
            try:
                records, history, wb_project, wb_contractor = parse_workbook_bytes(pending_bytes)
                warnings, errors = validate_records(records)
                file_hash = workbook_hash(pending_bytes)
                dup = find_import_by_hash(supabase, file_hash)

                st.markdown("#### Preview")
                p1, p2, p3 = st.columns(3)
                p1.metric("Filename", pending_name)
                p2.metric("Rows Found", len(records))
                p3.metric("Warnings", len(warnings))

                if errors:
                    for e in errors:
                        st.error(e)
                if warnings:
                    for w in warnings:
                        st.warning(w)

                preview_df = pd.DataFrame(records[:10])
                cols_to_show = [c for c in ["doc_no", "title", "discipline", "status", "latest_rev", "latest_code"] if c in preview_df.columns]
                preview_display = preview_df[cols_to_show] if cols_to_show else preview_df
                # Raw Excel cells often mix plain numbers (1, 2, 3) with text ("R", "V",
                # "1 WITH COMMENTS") in the same column (e.g. latest_code, latest_rev).
                # Arrow can't serialize a mixed int/str object column, so stringify
                # everything for display purposes only -- the underlying import still
                # uses the original typed values.
                preview_display = preview_display.astype(object).where(preview_display.notna(), "").astype(str)
                st.dataframe(preview_display, use_container_width=True, hide_index=True)

                if dup:
                    st.info(
                        f"This exact workbook was already imported as **{dup['display_name']}** "
                        f"(v{dup['version_number']}, {dup['uploaded_at'][:10]})."
                    )

                if not errors:
                    st.markdown("#### Confirm Import")
                    mode_options = ["Import as new workbook"]
                    if imports:
                        mode_options = ["Create New Version", "Replace Current Import", "Import as Separate Workbook"]
                    version_choice = st.radio("Import mode", mode_options, horizontal=True)
                    ic1, ic2 = st.columns([1, 4])
                    if ic1.button("Confirm Import", type="primary"):
                        mode_map = {
                            "Create New Version": "new_version",
                            "Replace Current Import": "replace",
                            "Import as Separate Workbook": "separate",
                            "Import as new workbook": "separate",
                        }
                        mode = mode_map[version_choice]
                        parent_id = active_import_id if mode in ("new_version", "replace") else None
                        try:
                            with st.spinner("Importing..."):
                                new_import_id, import_warnings = run_import(
                                    supabase, pending_bytes, pending_name, wb_project, wb_contractor,
                                    review_period_days, mode, parent_id,
                                )
                            st.session_state.pop("pending_upload_bytes", None)
                            st.session_state.pop("pending_upload_name", None)
                            st.cache_data.clear()
                            st.success(f"Import complete: {len(records)} documents imported.")
                            st.rerun()
                        except Exception as ex:
                            st.error(f"Import failed: {ex}")
                    if ic2.button("Cancel Upload"):
                        st.session_state.pop("pending_upload_bytes", None)
                        st.session_state.pop("pending_upload_name", None)
                        st.rerun()
            except Exception as ex:
                st.error(f"Could not read this workbook: {ex}")

        st.markdown("---")
        st.subheader("Export")

        if not active_import_id:
            st.caption("No active workbook to export.")
        else:
            ec1, ec2 = st.columns(2)
            with ec1:
                if st.button("Generate Updated Workbook (.xlsx)"):
                    with st.spinner("Regenerating workbook from current data..."):
                        try:
                            wb_bytes, wb_filename = generate_updated_workbook(supabase, active_import_id)
                            record_export(supabase, active_import_id, "workbook", wb_filename, len(df))
                            st.download_button(
                                "Download Updated Workbook", data=wb_bytes, file_name=wb_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                        except Exception as ex:
                            st.error(f"Export failed: {ex}")
            with ec2:
                if has_data:
                    all_history_rows = get_history_for_documents(supabase, df["document_id"].dropna().unique().tolist())
                    if all_history_rows:
                        hist_csv = pd.DataFrame(all_history_rows).to_csv(index=False).encode("utf-8")
                        st.download_button(
                            "Download History Report (CSV)", data=hist_csv,
                            file_name=f"dci_history_report_{datetime.now().strftime('%Y-%m-%d')}.csv",
                            mime="text/csv",
                        )
                    else:
                        st.caption("No history events recorded yet.")

            if has_data:
                filtered_csv = filtered.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "Download Filtered Document Register (CSV)", data=filtered_csv,
                    file_name=f"dci_filtered_register_{datetime.now().strftime('%Y-%m-%d')}.csv",
                    mime="text/csv",
                )

            st.markdown("#### Previous Exports")
            exp_res = supabase.table("export_records").select("*").eq("import_id", active_import_id) \
                .order("exported_at", desc=True).limit(20).execute()
            if exp_res.data:
                st.dataframe(pd.DataFrame(exp_res.data), use_container_width=True, hide_index=True)
            else:
                st.caption("No exports recorded yet for this workbook.")

except Exception as e:
    st.error(
        "Something went wrong while rendering the dashboard tabs. "
        "The details below should help pin down the cause."
    )
    st.exception(e)
