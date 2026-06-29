import streamlit as st
import json
import re
import os
import urllib.request
from datetime import datetime
from io import BytesIO
import difflib
import html
import hashlib

try:
    from groq import Groq
    GROQ_OK = True
except ImportError:
    GROQ_OK = False

try:
    from docx import Document
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

try:
    import pdfplumber
    PDF_OK = True
except ImportError:
    PDF_OK = False

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    GOOGLE_OK = True
except ImportError:
    HttpError = Exception
    GOOGLE_OK = False

SERVICE_ACCOUNT_EMAIL = "content-qa-bot@bayut-competitor-gap-analysis.iam.gserviceaccount.com"

def get_service_account_email():
    """Return the service account email shown to users for Google Doc sharing."""
    try:
        info = dict(st.secrets.get("gcp_service_account", {}))
        return info.get("client_email") or SERVICE_ACCOUNT_EMAIL
    except Exception:
        return SERVICE_ACCOUNT_EMAIL

def friendly_google_api_error(e, doc_id=""):
    """Convert Google API errors into clear Streamlit-facing messages."""
    status = None
    try:
        status = getattr(getattr(e, "resp", None), "status", None)
    except Exception:
        status = None

    service_email = get_service_account_email()
    doc_hint = f" Document ID: {doc_id}." if doc_id else ""

    if status == 404:
        return (
            "Google Drive could not find or access this document."
            f"{doc_hint} Share the Google Doc with {service_email} as Editor, "
            "then try again. Also make sure the link is a Google Doc link copied from the browser, not a shortcut or a deleted/moved file."
        )
    if status == 403:
        return (
            "Google Drive found the document but permission is blocked."
            f"{doc_hint} Share it with {service_email} as Editor. "
            "If it is shared as 'Anyone with the link', change it to company-restricted sharing."
        )
    return f"Google API error: {e}"

st.set_page_config(page_title="Content QA System", page_icon="Q", layout="wide",
                   initial_sidebar_state="expanded")

PLATFORMS     = ["Bayut", "Dubizzle"]
CONTENT_TYPES = ["Landing page", "Blog post", "Property guide"]
LANGUAGES     = ["English", "Arabic"]

CAT_MAX = {
    "Content Quality":    25,
    "SEO & Structure":    20,
    "Language & Grammar": 20,
    "Brand Voice":        15,
    "Readability & Flow": 10,
    "Originality":        10,
}

GRADE_MAP = [
    (90, "A — Excellent"),
    (80, "B — Good"),
    (70, "C — Needs revision"),
    (60, "D — Major revision"),
    (0,  "F — Reject"),
]

# Weighted issue types used for comments and silent edits.
COMMENT_WEIGHTS = {
    "factual":            {"label": "Factual correction",        "deduction": 3.0, "color": "#fee2e2", "tc": "#991b1b"},
    "wrong_info_removed": {"label": "Wrong info removed",        "deduction": 2.0, "color": "#fee2e2", "tc": "#991b1b"},
    "source_alignment":   {"label": "Source alignment",          "deduction": 3.0, "color": "#fee2e2", "tc": "#991b1b"},
    "contradiction_fixed": {"label": "Contradiction fixed",      "deduction": 3.0, "color": "#fee2e2", "tc": "#991b1b"},
    "missing":            {"label": "Missing critical info",     "deduction": 1.5, "color": "#fef3c7", "tc": "#92400e"},
    "missing_info_added": {"label": "Missing info added",        "deduction": 1.2, "color": "#fef3c7", "tc": "#92400e"},
    "structural":         {"label": "Structural rewrite",        "deduction": 1.2, "color": "#fde8d8", "tc": "#9a3412"},
    "arabic_language":    {"label": "Arabic language correction", "deduction": 0.6, "color": "#e0f2fe", "tc": "#075985"},
    "grammar":            {"label": "Grammar / phrasing",        "deduction": 0.5, "color": "#f0f4ff", "tc": "#2D4A8A"},
    "rephrase":           {"label": "Rephrase only",             "deduction": 0.3, "color": "#f1f5f9", "tc": "#475569"},
    "formatting":         {"label": "Formatting only",           "deduction": 0.0, "color": "#f8fafc", "tc": "#64748b"},
}

LOW_IMPACT_EDIT_TYPES = {"rephrase", "grammar", "arabic_language", "formatting"}
HIGH_IMPACT_EDIT_TYPES = {"factual", "wrong_info_removed", "source_alignment", "contradiction_fixed", "missing", "missing_info_added"}
EVENT_ONLY_EDIT_TYPES = {"revision_event"}
REVISION_ROUND_PENALTY = 0.7  # per extra round

RECORDS_FILE = "qa_records.json"
RECORDS_SHEET_TAB = "qa_records"
RECORDS_SHEET_CHUNK_SIZE = 45000

# ── Shared persistence ─────────────────────────────────────────────────────
def _records_sheet_id():
    """Optional shared Google Sheet ID for dashboard records across all users."""
    try:
        return (st.secrets.get("QA_RECORDS_SHEET_ID", "") or os.environ.get("QA_RECORDS_SHEET_ID", "")).strip()
    except Exception:
        return os.environ.get("QA_RECORDS_SHEET_ID", "").strip()

def _records_sheet_tab():
    try:
        return (st.secrets.get("QA_RECORDS_SHEET_TAB", "") or RECORDS_SHEET_TAB).strip() or RECORDS_SHEET_TAB
    except Exception:
        return RECORDS_SHEET_TAB

def is_shared_dashboard_connected():
    """True when the dashboard is configured to use the shared Google Sheet backend."""
    return bool(_records_sheet_id())

def shared_dashboard_help_text():
    return (
        "Shared dashboard is not connected. Add QA_RECORDS_SHEET_ID to Streamlit secrets "
        "and share that Google Sheet with the service account as Editor. Until then, records "
        "are only saved inside this app instance and will not be shared reliably across laptops."
    )

def _get_sheets_service():
    if not GOOGLE_OK:
        raise Exception("google-api-python-client not installed")
    info = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(
        info,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.readonly",
        ],
    )
    return build("sheets", "v4", credentials=creds)

def _ensure_records_sheet(sheets_svc, sheet_id):
    """Create the records tab/header if needed."""
    tab = _records_sheet_tab()
    meta = sheets_svc.spreadsheets().get(spreadsheetId=sheet_id).execute()
    titles = [s.get("properties", {}).get("title") for s in meta.get("sheets", [])]
    if tab not in titles:
        sheets_svc.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": tab}}}]},
        ).execute()
    values = sheets_svc.spreadsheets().values().get(
        spreadsheetId=sheet_id, range=f"{tab}!A1:Z1"
    ).execute().get("values", [])
    
    if not values:
        # New formatted column headers
        header = ["Writer Name", "Title", "URL", "Platform (Bayut/Dubizzle)", "Score", "All Comments & Edits", "updated_at", "full_payload"]
        sheets_svc.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{tab}!A1:H1",
            valueInputOption="RAW",
            body={"values": [header]},
        ).execute()

def _load_records_from_sheet():
    sheet_id = _records_sheet_id()
    if not sheet_id:
        return None
    sheets = _get_sheets_service()
    _ensure_records_sheet(sheets, sheet_id)
    tab = _records_sheet_tab()
    
    # Read from row 2 up to column H where the payload is
    rows = sheets.spreadsheets().values().get(
        spreadsheetId=sheet_id, range=f"{tab}!A2:H"
    ).execute().get("values", [])
    
    records = []
    for row in rows:
        if len(row) < 8: 
            continue
        payload = row[7].strip() 
        if not payload:
            continue
        try:
            rec = json.loads(payload)
            if isinstance(rec, dict):
                records.append(rec)
        except Exception:
            continue
    return records

def _save_records_to_sheet(records):
    sheet_id = _records_sheet_id()
    if not sheet_id:
        return False
    sheets = _get_sheets_service()
    _ensure_records_sheet(sheets, sheet_id)
    tab = _records_sheet_tab()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    
    for rec in records:
        clean = _serialisable(rec)
        
        # Compile all comments and silent edits into a single text block
        comments_list = []
        for c in clean.get("classified", []):
            comments_list.append(f"• [تعليق - {c.get('label')}]: {c.get('text')}")
            
        for d in clean.get("diff_classified", []):
            original = d.get('original', '').strip()
            revised = d.get('revised', '').strip()
            if original or revised:
                comments_list.append(f"• [تعديل صامت - {d.get('label')}]: الكاتب قائل: '{original}' -> المحرر عدلها إلى: '{revised}'")
                
        all_comments_text = "\n".join(comments_list) if comments_list else "لا توجد ملاحظات أو تعديلات."

        # Extract values for specific columns
        writer_name = clean.get("writer", "—")
        title = clean.get("title", "—")
        url = clean.get("doc_url", clean.get("file_name", "—"))
        platform = clean.get("platform", "—")
        score = clean.get("qa_score", 0)
        
        # The full JSON payload to keep the dashboard working correctly
        payload_json = json.dumps(clean, ensure_ascii=False, separators=(",", ":"))
        
        row = [
            writer_name,       
            title,             
            url,               
            platform,          
            score,             
            all_comments_text, 
            now,               
            payload_json       
        ]
        rows.append(row)
        
    header = ["Writer Name", "Title", "URL", "Platform (Bayut/Dubizzle)", "Score", "All Comments & Edits", "updated_at", "full_payload"]
    
    # Clear out old rows and overwrite with new mapped ones
    sheets.spreadsheets().values().clear(
        spreadsheetId=sheet_id, range=f"{tab}!A:ZZ"
    ).execute()
    
    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{tab}!A1",
        valueInputOption="RAW",
        body={"values": [header] + rows},
    ).execute()
    return True

def _load_records_local():
    if not os.path.exists(RECORDS_FILE):
        return []
    try:
        with open(RECORDS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def _save_records_local(records):
    with open(RECORDS_FILE, "w", encoding="utf-8") as f:
        json.dump(_serialisable(records), f, ensure_ascii=False, indent=2)

def load_records():
    """Load dashboard records. Uses a shared Google Sheet when configured, otherwise local JSON."""
    try:
        shared = _load_records_from_sheet()
        if shared is not None:
            return shared
    except Exception as e:
        try:
            st.warning(f"Could not load shared dashboard records, using local cache only. Details: {e}")
        except Exception:
            pass
    return _load_records_local()

def save_records(records):
    """Save all dashboard records to the shared backend or local fallback."""
    if not _records_sheet_id():
        try:
            st.warning(shared_dashboard_help_text())
        except Exception:
            pass
        _save_records_local(records)
        return
    try:
        if _save_records_to_sheet(_serialisable(records)):
            return
    except Exception as e:
        try:
            st.warning(f"Could not save to the shared dashboard, saving locally only. Details: {e}")
        except Exception:
            pass
    _save_records_local(records)

def save_record(sub):
    records = load_records()
    key = _record_storage_key(sub)
    for r in records:
        if _record_storage_key(r) == key:
            return
    records.append(_serialisable(sub))
    save_records(records)

def update_record_decision(sub):
    records = load_records()
    key = _record_storage_key(sub)
    for r in records:
        if _record_storage_key(r) == key:
            r["editor_decision"] = sub.get("editor_decision", "")
            r["editor_notes"]    = sub.get("editor_notes", "")
            break
    save_records(records)

def _serialisable(obj):
    if isinstance(obj, dict):          return {k: _serialisable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)): return [_serialisable(i) for i in obj]
    if isinstance(obj, (str, int, float, bool)) or obj is None: return obj
    return str(obj)

def _safe_key_part(value):
    """Return a short safe string for Streamlit widget keys."""
    value = str(value or "").strip()
    value = re.sub(r"[^A-Za-z0-9_\-]+", "_", value)
    value = re.sub(r"_+", "_", value).strip("_")
    return value[:80] or "blank"

def _submission_identity(sub):
    fields = [
        sub.get("mode", ""),
        sub.get("title", ""),
        sub.get("writer", ""),
        sub.get("editor_name", ""),
        sub.get("date", ""),
        sub.get("platform", ""),
        sub.get("content_type", ""),
        sub.get("language", ""),
        sub.get("doc_url", ""),
        sub.get("file_name", ""),
    ]
    raw = "||".join(str(x or "") for x in fields)
    return hashlib.md5(raw.encode("utf-8")).hexdigest()[:12]

def _sub_widget_key(sub, prefix):
    render_uid = sub.get("_render_uid") or "single"
    return f"{prefix}_{_safe_key_part(sub.get('title','Untitled'))}_{_submission_identity(sub)}_{_safe_key_part(render_uid)}"

def _record_storage_key(sub):
    return (
        sub.get("mode", ""), sub.get("title", ""), sub.get("writer", ""),
        sub.get("editor_name", ""), sub.get("date", ""), sub.get("platform", ""),
        sub.get("content_type", ""), sub.get("language", ""), sub.get("doc_url", ""),
        sub.get("file_name", ""), round(float(sub.get("qa_score", 0) or 0), 2),
    )

# ── CSS ────────────────────────────────────────────────────────────────────
def inject_css():
    st.markdown("""
<style>
.stApp{background:#f5f6fb}
.block-container{max-width:1220px !important;padding-top:1.8rem !important;padding-left:2rem !important;padding-right:2rem !important;padding-bottom:3rem}
[data-testid="stVerticalBlockBorderWrapper"]{border:1px solid #e5e7eb !important;border-radius:20px !important;box-shadow:0 14px 35px rgba(17,24,39,.06) !important;background:#fff !important}
[data-testid="stForm"]{border:none !important;padding:0 !important;background:transparent !important}
[data-testid="stSidebar"]{background:#ffffff !important;border-right:1px solid #e5e7eb !important}
section[data-testid="stSidebar"]>div{padding:0 !important}
.sb-brand{display:flex;align-items:center;gap:10px;padding:22px 18px 18px 18px;border-bottom:1px solid #e5e7eb}
.sb-brand-icon{width:34px;height:34px;border-radius:50%;background:linear-gradient(135deg,#5b5ce2,#7c3aed);display:flex;align-items:center;justify-content:center;color:white;font-size:15px;font-weight:800;flex-shrink:0}
.sb-brand-title{font-size:13px;font-weight:800;color:#111827;line-height:1.2}
.sb-brand-sub{font-size:11px;color:#9ca3af;margin-top:2px}
.sb-section{font-size:10px;color:#9ca3af;font-weight:800;text-transform:uppercase;letter-spacing:.08em;margin:17px 18px 8px 18px}
section[data-testid="stSidebar"] [data-testid="stRadio"]{padding:0 14px !important}
section[data-testid="stSidebar"] [role="radiogroup"]{gap:6px !important}
section[data-testid="stSidebar"] [role="radiogroup"] label{border-radius:12px !important;padding:10px 12px !important;margin:0 0 5px 0 !important;background:transparent !important;color:#374151 !important;font-size:13px !important;font-weight:600 !important}
section[data-testid="stSidebar"] [role="radiogroup"] label:hover{background:#f3f4f6 !important}
section[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked){background:#eeebff !important;color:#4f46e5 !important;font-weight:800 !important}
section[data-testid="stSidebar"] [role="radiogroup"] label>div:first-child{display:none !important}
.qa-hero{background:linear-gradient(135deg,#4839d8 0%,#7c3aed 55%,#d067da 100%);border-radius:22px;padding:30px 32px;margin-bottom:24px;color:#fff;display:flex;align-items:flex-start;justify-content:space-between;box-shadow:0 18px 35px rgba(79,70,229,.18);min-height:138px;position:relative;overflow:hidden}
.qa-hero-badge{display:inline-block;background:rgba(255,255,255,.16);color:#fff;border-radius:999px;padding:5px 12px;font-size:11px;font-weight:800;margin-bottom:12px}
.qa-hero h1{font-size:30px;font-weight:900;margin:0 0 10px 0;color:#fff;line-height:1.15}
.qa-hero p{font-size:13px;line-height:1.6;color:rgba(255,255,255,.88);margin:0;max-width:520px}
.qa-hero-icon{width:66px;height:66px;border-radius:18px;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.15);display:flex;align-items:center;justify-content:center;font-size:28px;flex-shrink:0}
div[class*="stTextInput"]>label,div[class*="stSelectbox"]>label,div[class*="stFileUploader"]>label{font-size:12px !important;font-weight:800 !important;color:#374151 !important;margin-bottom:6px !important}
[data-testid="stTextInput"] input{border-radius:12px !important;border:1px solid #d9e0ea !important;padding:10px 13px !important;font-size:13px !important;background:#fff !important;box-shadow:none !important}
[data-testid="stTextInput"] input:focus{border-color:#8b5cf6 !important;box-shadow:0 0 0 3px rgba(139,92,246,.10) !important}
[data-baseweb="select"]>div{border-radius:12px !important;border:1px solid #d9e0ea !important;min-height:42px !important;box-shadow:none !important}
[data-testid="stForm"] div[role="radiogroup"]{display:inline-flex !important;flex-direction:row !important;gap:0 !important;background:#f1f5f9 !important;border:1px solid #dbe3ee !important;border-radius:999px !important;padding:3px !important;width:fit-content !important}
[data-testid="stForm"] div[role="radiogroup"] label{margin:0 !important;min-height:32px !important;min-width:76px !important;padding:7px 17px !important;border-radius:999px !important;border:none !important;background:transparent !important;color:#64748b !important;font-size:12px !important;font-weight:700 !important;display:flex !important;align-items:center !important;justify-content:center !important;box-shadow:none !important}
[data-testid="stForm"] div[role="radiogroup"] label:has(input:checked){background:#10b981 !important;color:#fff !important;box-shadow:0 6px 13px rgba(16,185,129,.25) !important}
[data-testid="stForm"] div[role="radiogroup"] label>div:first-child{display:none !important}
[data-testid="stForm"] div[data-testid="stRadio"]{display:inline-block !important;vertical-align:middle !important;margin-bottom:14px !important}
[data-testid="stForm"] div[data-testid="stRadio"]>label{display:none !important}
[data-testid="stFileUploader"]>label{display:none !important}
[data-testid="stFileUploader"] section,[data-testid="stFileUploaderDropzone"],[data-testid="stFileUploadDropzone"]{min-height:152px !important;border:1.5px dashed #9fa8ff !important;border-radius:18px !important;background:#f7f8ff !important;position:relative !important;display:flex !important;align-items:center !important;justify-content:center !important;padding:0 !important;overflow:hidden !important}
[data-testid="stFileUploader"] section button,[data-testid="stFileUploader"] section svg,[data-testid="stFileUploader"] section small,[data-testid="stFileUploader"] section span,[data-testid="stFileUploader"] section p{opacity:0 !important;visibility:hidden !important}
[data-testid="stFileUploader"] section input[type="file"]{position:absolute !important;inset:0 !important;width:100% !important;height:100% !important;opacity:0 !important;cursor:pointer !important;z-index:20 !important}
[data-testid="stFileUploader"] section>div,[data-testid="stFileUploaderDropzone"]>div,[data-testid="stFileUploadDropzone"]>div{opacity:0 !important;visibility:hidden !important;pointer-events:none !important}
[data-testid="stFileUploader"] section::before{content:"⇧";position:absolute;left:50%;top:30px;transform:translateX(-50%);width:44px;height:44px;border-radius:16px;background:linear-gradient(135deg,#5b5ce2,#7c3aed);color:#fff;font-size:24px;font-weight:800;display:flex;align-items:center;justify-content:center;z-index:5;visibility:visible !important;opacity:1 !important}
[data-testid="stFileUploader"] section::after{content:"Click or drag a file to upload";position:absolute;left:50%;top:88px;transform:translateX(-50%);width:100%;text-align:center;font-size:13px;font-weight:800;color:#111827;z-index:5;visibility:visible !important;opacity:1 !important}
[data-testid="stFormSubmitButton"] button{background:linear-gradient(135deg,#4338ca,#7c3aed) !important;color:white !important;border:none !important;border-radius:12px !important;font-size:14px !important;font-weight:900 !important;padding:12px !important;width:100% !important;box-shadow:0 12px 24px rgba(124,58,237,.18) !important}
.score-hero{background:#fff;border:1px solid #e8eaf0;border-radius:14px;padding:22px 26px;margin-bottom:1rem}
.score-num{font-size:54px;font-weight:700;color:#4f46e5;line-height:1}
.score-den{font-size:17px;font-weight:400;color:#9ca3af}
.score-grade{font-size:13px;font-weight:600;margin-top:5px;color:#4f46e5}
.score-verdict{font-size:13px;color:#6b7280;line-height:1.65;margin-top:10px}
.breakdown-box{background:#f9fafb;border:1px solid #e8eaf0;border-radius:10px;padding:13px 15px;margin-top:13px;font-size:13px}
.ded-row{display:flex;justify-content:space-between;padding:5px 0;color:#dc2626;border-bottom:1px solid #fee2e2}
.base-row{display:flex;justify-content:space-between;padding:5px 0;color:#374151;border-bottom:1px solid #f3f4f6}
.ok-row{display:flex;justify-content:space-between;padding:5px 0;color:#9ca3af;border-bottom:1px solid #f3f4f6;font-size:12px}
.total-row{display:flex;justify-content:space-between;padding:7px 0 2px;font-weight:700;font-size:14px;color:#111827;border-top:2px solid #e8eaf0;margin-top:3px}
.cmt-card{background:#f0f2f9;border-left:3px solid #4f46e5;padding:9px 13px;margin-bottom:7px;border-radius:0 8px 8px 0;font-size:13px}
.cmt-author{font-weight:600;color:#4f46e5}
.cmt-deduct{font-size:11px;color:#dc2626;font-weight:500;margin-top:3px}
.cat-ref{font-size:10px;font-weight:500;padding:2px 7px;border-radius:20px;background:#ede9fe;color:#4f46e5;margin-left:6px}
.suggest-item{display:flex;gap:11px;align-items:flex-start;padding:9px 0;border-bottom:1px solid #f3f4f6;font-size:13px}
.suggest-num{width:22px;height:22px;border-radius:50%;background:#ede9fe;color:#4f46e5;font-size:10px;font-weight:600;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:1px}
.suggest-cat{font-size:11px;color:#9ca3af;margin-top:2px}
.tag-str{background:#d1fae5;color:#065f46;padding:3px 11px;border-radius:20px;font-size:12px;font-weight:500;display:inline-block;margin:2px}
.tag-imp{background:#fef3c7;color:#92400e;padding:3px 11px;border-radius:20px;font-size:12px;font-weight:500;display:inline-block;margin:2px}
.bdg{font-size:11px;font-weight:500;padding:3px 10px;border-radius:20px}
.bdg-bay{background:#d1fae5;color:#065f46}
.bdg-dub{background:#fee2e2;color:#b91c1c}
.no-cmt-notice{background:#f0f2f9;border:1px solid #e0e4f0;border-radius:8px;padding:11px 15px;font-size:13px;color:#6b7280;margin-bottom:10px}
.dash-stats-row{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin-bottom:24px}
.dash-stat{background:#fff;border:1px solid #e5e7eb;border-radius:16px;padding:16px 14px;text-align:center}
.dash-stat-num{font-size:28px;font-weight:900;color:#111827;line-height:1}
.dash-stat-lbl{font-size:11px;font-weight:700;color:#9ca3af;margin-top:4px;text-transform:uppercase;letter-spacing:.05em}
.dash-stat.green .dash-stat-num{color:#059669}
.dash-stat.amber .dash-stat-num{color:#d97706}
.dash-stat.red   .dash-stat-num{color:#dc2626}
.dash-stat.blue  .dash-stat-num{color:#4f46e5}
.article-card{background:#fff;border:1px solid #e5e7eb;border-radius:18px;padding:18px 20px;margin-bottom:12px;display:grid;grid-template-columns:1fr auto;gap:16px;align-items:start}
.article-card:hover{box-shadow:0 8px 24px rgba(17,24,39,.09)}
.article-card-left{display:flex;flex-direction:column;gap:6px}
.article-card-title{font-size:15px;font-weight:900;color:#111827;line-height:1.25}
.article-card-meta{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
.meta-chip{font-size:11px;font-weight:700;padding:3px 9px;border-radius:20px;background:#f1f5f9;color:#475569}
.meta-chip.bay{background:#d1fae5;color:#065f46}
.meta-chip.dub{background:#fee2e2;color:#b91c1c}
.meta-chip.eng{background:#ede9fe;color:#5b21b6}
.meta-chip.ara{background:#fef3c7;color:#92400e}
.meta-chip.gdoc{background:#e8f0fe;color:#1a56db}
.article-card-summary{font-size:12px;color:#6b7280;line-height:1.6;margin-top:2px}
.article-card-right{display:flex;flex-direction:column;align-items:flex-end;gap:8px}
.score-ring{width:64px;height:64px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:radial-gradient(circle closest-side,white 72%,transparent 74%),conic-gradient(var(--rc) var(--rv),#eef2f7 0);font-size:16px;font-weight:900;color:#111827;margin:0 auto 4px}
.score-ring-lbl{font-size:10px;font-weight:700;color:#9ca3af;text-align:center}
.dec-badge{font-size:11px;font-weight:800;padding:4px 10px;border-radius:999px;white-space:nowrap}
.dec-approve{background:#d1fae5;color:#065f46}
.dec-revise{background:#fef3c7;color:#92400e}
.dec-reject{background:#fee2e2;color:#991b1b}
.dec-pending{background:#f1f5f9;color:#64748b}
.stepper{display:flex;align-items:center;justify-content:center;gap:14px;width:100%;max-width:760px;margin:0 auto 24px auto}
.step-item{display:inline-flex;align-items:center;gap:10px;white-space:nowrap;font-size:13px;font-weight:800;color:#64748b}
.step-num{width:34px;height:34px;border-radius:50%;border:1px solid #dbe3ee;background:#fff;display:inline-flex;align-items:center;justify-content:center;font-size:13px;font-weight:900;color:#111827}
.step-item.active{color:#6d28d9}
.step-item.active .step-num{background:linear-gradient(135deg,#5b5ce2,#7c3aed);color:#fff;border-color:transparent}
.step-line{flex:0 0 64px;height:1px;border-top:1px dashed #d6dbe6}
.side-card{background:#fff;border:1px solid #e5e7eb;border-radius:18px;padding:20px;box-shadow:0 12px 30px rgba(17,24,39,.055);margin-bottom:14px}
.side-card-title{font-size:14px;font-weight:900;color:#111827;margin-bottom:14px}
.timeline-row{display:grid;grid-template-columns:30px 1fr;gap:10px;margin-bottom:16px}
.timeline-num{width:28px;height:28px;border-radius:50%;background:linear-gradient(135deg,#5b5ce2,#7c3aed);color:white;font-weight:900;font-size:12px;display:flex;align-items:center;justify-content:center}
.timeline-title{color:#111827;font-size:12px;font-weight:900;margin-bottom:2px}
.timeline-sub{color:#64748b;font-size:11px;line-height:1.35}
.tip-box{background:#ecfdf5;border:1px solid #d1fae5;color:#065f46;border-radius:14px;padding:14px;font-size:12px;line-height:1.45}
.tip-title{font-size:13px;font-weight:900;margin-bottom:4px}
.file-card{display:flex;align-items:center;gap:12px;background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:12px 14px;margin-top:12px}
.file-icon{width:34px;height:34px;border-radius:10px;background:#eef2ff;color:#4f46e5;display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0}
.file-title{font-size:13px;color:#111827;font-weight:800}
.file-meta{font-size:11px;color:#64748b;margin-top:2px}
.file-status{margin-left:auto;background:#dcfce7;color:#15803d;font-size:11px;font-weight:800;border-radius:999px;padding:5px 10px}
.precheck{display:grid;grid-template-columns:repeat(4,1fr);background:#fbfcff;border:1px solid #e5e7eb;border-radius:13px;padding:8px 10px;margin-top:12px}
.precheck-item{display:flex;align-items:center;gap:7px;color:#64748b;font-size:11px;font-weight:700}
.precheck-dot{width:15px;height:15px;border-radius:50%;background:#e5e7eb;color:#94a3b8;display:inline-flex;align-items:center;justify-content:center;font-size:10px;font-weight:900}
.precheck-item.done{color:#334155}
.precheck-item.done .precheck-dot{background:#22c55e;color:white}
.form-section-divider{height:1px;background:#e8edf5;margin:16px 0 14px 0}
.form-card-header{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:20px}
.form-card-title{font-size:16px;font-weight:900;color:#111827;margin-bottom:3px}
.form-card-sub{font-size:12px;color:#6b7280}
.ready-badge{display:inline-flex;align-items:center;gap:6px;background:#dcfce7;color:#15803d;font-size:11px;font-weight:800;padding:5px 11px;border-radius:999px}
.ready-dot{width:6px;height:6px;border-radius:50%;background:#16a34a;display:inline-block}
.gdoc-input-wrap{border:1.5px dashed #9fa8ff;border-radius:18px;background:#f7f8ff;padding:28px 24px;margin-top:8px;text-align:center}
.gdoc-icon{font-size:32px;margin-bottom:10px}
.gdoc-label{font-size:13px;font-weight:800;color:#111827;margin-bottom:4px}
.gdoc-sub{font-size:11px;color:#9ca3af;margin-bottom:14px}
.rev-card{background:#f8faff;border:1px solid #e0e7ff;border-radius:12px;padding:14px 16px;margin-bottom:10px}
.rev-round{display:flex;align-items:center;gap:10px;margin-bottom:6px}
.rev-badge{font-size:10px;font-weight:800;padding:3px 9px;border-radius:20px}
.rev-editor{background:#fee2e2;color:#991b1b}
.rev-writer{background:#d1fae5;color:#065f46}
div[data-testid="stVerticalBlockBorderWrapper"]>div{border-radius:22px !important}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stVerticalBlock"]{gap:0.9rem !important}
</style>
""", unsafe_allow_html=True)

# ── Groq AI ────────────────────────────────────────────────────────────────
GROQ_MODELS = ["llama-3.3-70b-versatile", "llama-3.1-8b-instant", "llama3-8b-8192", "gemma2-9b-it"]

def call_ai(prompt, max_retries=3):
    if not GROQ_OK:
        raise Exception("groq package not installed")
    client   = Groq(api_key=st.secrets["GROQ_API_KEY"])
    last_err = "No models attempted"
    for model in GROQ_MODELS:
        for attempt in range(max_retries):
            try:
                resp = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.3,
                    max_tokens=2000,
                )
                text = resp.choices[0].message.content
                if text and text.strip():
                    return text.strip()
                last_err = f"{model}: empty response (attempt {attempt + 1})"
            except Exception as e:
                last_err = f"{model}: {e}"
                if any(x in str(e).lower() for x in ["model", "not found", "decommission"]):
                    break
    raise Exception(f"All Groq models failed. Last: {last_err}")

def parse_json_response(raw):
    if not raw or not raw.strip():
        return None
    clean = re.sub(r"```json\s*|```\s*", "", raw).strip()
    m = re.search(r'\{.*\}', clean, re.DOTALL)
    if m:
        clean = m.group(0)
    try:
        return json.loads(clean)
    except Exception:
        try:
            return json.loads(raw.strip())
        except Exception:
            return None

# ── Google Doc integration ─────────────────────────────────────────────────
def get_google_services():
    if not GOOGLE_OK:
        raise Exception("google-api-python-client not installed")
    info  = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(
        info,
        scopes=[
            "https://www.googleapis.com/auth/documents.readonly",
            "https://www.googleapis.com/auth/drive.readonly",
            "https://www.googleapis.com/auth/drive.activity.readonly",
            "https://www.googleapis.com/auth/spreadsheets",
        ]
    )
    docs  = build("docs",  "v1", credentials=creds)
    drive = build("drive", "v3", credentials=creds)
    return docs, drive, creds

def extract_doc_id(url):
    """Extract a Google Doc ID from a full URL or a raw document ID."""
    value = str(url or "").strip()
    if not value:
        return None

    value = value.split("#", 1)[0].split("?", 1)[0]
    m = re.search(r"/document/d/([a-zA-Z0-9_-]+)", value)
    if m:
        return m.group(1)

    if re.fullmatch(r"[a-zA-Z0-9_-]{20,}", value):
        return value

    return None

def clean_google_doc_url(url):
    """Return a clean canonical Google Doc URL for display/debugging."""
    doc_id = extract_doc_id(url)
    return f"https://docs.google.com/document/d/{doc_id}/edit" if doc_id else str(url or "").strip()

def get_allowed_google_doc_domains():
    defaults = ["dubizzle.com", "bayut.com", "dubizzlegroup.com", "bayut.jo"]
    try:
        raw = st.secrets.get("ALLOWED_GOOGLE_DOC_DOMAINS", "")
        custom = [d.strip().lower().lstrip("@") for d in raw.split(",") if d.strip()]
        return custom or defaults
    except Exception:
        return defaults

def _email_domain(email):
    email = str(email or "").strip().lower()
    if "@" not in email:
        return ""
    return email.rsplit("@", 1)[-1]

def is_allowed_company_email(email):
    domain = _email_domain(email)
    allowed_domains = get_allowed_google_doc_domains()
    return bool(domain) and domain in allowed_domains

def is_content_system_email(email):
    """Return True when an email is the Content QA service account."""
    email = str(email or "").strip().lower()
    allowed = {
        SERVICE_ACCOUNT_EMAIL.lower(),
        get_service_account_email().lower(),
    }
    try:
        raw = st.secrets.get("ALLOWED_CONTENT_SYSTEM_EMAILS", "")
        allowed.update(e.strip().lower() for e in raw.split(",") if e.strip())
    except Exception:
        pass
    return email in allowed

def has_content_system_editor_access(permissions):
    """Allow the doc when the Content QA service account is shared as Editor."""
    for p in permissions or []:
        email = str(p.get("emailAddress") or "").strip().lower()
        role = str(p.get("role") or "").strip().lower()
        if is_content_system_email(email) and role in {"owner", "organizer", "fileorganizer", "writer"}:
            return True
    return False

def validate_dubizzle_group_google_doc(drive_svc, doc_id):
    allowed_domains = get_allowed_google_doc_domains()

    try:
        meta = drive_svc.files().get(
            fileId=doc_id,
            fields=(
                "id,name,mimeType,ownedByMe,owners(displayName,emailAddress),"
                "permissions(id,type,role,emailAddress,domain,allowFileDiscovery)"
            ),
            supportsAllDrives=True,
        ).execute()
    except HttpError as e:
        return False, friendly_google_api_error(e, doc_id)
    except Exception as e:
        return False, f"Could not verify document access/security. Share the doc with {get_service_account_email()} as Editor and try again. Details: {e}"

    if meta.get("mimeType") != "application/vnd.google-apps.document":
        return False, "Only Google Docs files are accepted. Please submit a Google Doc link."

    permissions = meta.get("permissions", []) or []
    owners = meta.get("owners", []) or []

    for p in permissions:
        if p.get("type") == "anyone":
            return False, "Public Google Docs are not allowed. Keep sharing restricted and share the doc with the Content QA service account as Editor."

    if has_content_system_editor_access(permissions):
        return True, ""

    for p in permissions:
        if p.get("type") == "domain":
            domain = str(p.get("domain") or "").strip().lower()
            if domain and domain not in allowed_domains:
                return False, f"This document is shared with an unapproved domain: {domain}. Allowed domains: {', '.join(allowed_domains)}."

    owner_emails = [o.get("emailAddress", "") for o in owners if o.get("emailAddress")]
    if owner_emails:
        if not any(is_allowed_company_email(email) for email in owner_emails):
            return False, f"This private doc is accessible, but the Content QA service account is not shared as Editor. Share it with {get_service_account_email()} as Editor and try again."
        return True, ""

    company_domain_permission = any(
        p.get("type") == "domain" and str(p.get("domain") or "").strip().lower() in allowed_domains
        for p in permissions
    )
    company_user_permission = any(
        p.get("type") == "user" and is_allowed_company_email(p.get("emailAddress", ""))
        for p in permissions
    )

    if company_domain_permission or company_user_permission:
        return True, ""

    return False, f"Could not confirm access. Keep the doc restricted and share it with {get_service_account_email()} as Editor."

def _execute_revision_get(drive_svc, doc_id, revision_id):
    base = dict(
        fileId=doc_id,
        revisionId=revision_id,
        fields="id,modifiedTime,lastModifyingUser,exportLinks",
    )
    try:
        return drive_svc.revisions().get(**base).execute()
    except TypeError:
        raise
    except Exception as first_err:
        try:
            return drive_svc.revisions().get(**base, supportsAllDrives=True).execute()
        except Exception:
            raise first_err

def export_revision_text(drive_svc, creds, doc_id, revision_id):
    """Export a specific Google Docs revision as plain text."""
    if not revision_id:
        return None
    try:
        import google.auth.transport.requests
        rev = _execute_revision_get(drive_svc, doc_id, revision_id)
        export_url = (rev.get("exportLinks") or {}).get("text/plain", "")
        if not export_url:
            return None

        auth_req = google.auth.transport.requests.Request()
        if not creds.valid:
            creds.refresh(auth_req)
        req = urllib.request.Request(
            export_url,
            headers={"Authorization": f"Bearer {creds.token}"}
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.read().decode("utf-8")
    except Exception:
        return None

def list_drive_revisions_safe(drive_svc, doc_id):
    revisions = []
    page_token = None
    while True:
        base = dict(
            fileId=doc_id,
            fields="nextPageToken,revisions(id,modifiedTime,lastModifyingUser,exportLinks)",
            pageSize=1000,
        )
        if page_token:
            base["pageToken"] = page_token

        try:
            resp = drive_svc.revisions().list(**base).execute()
        except Exception:
            try:
                resp = drive_svc.revisions().list(**base, supportsAllDrives=True).execute()
            except Exception:
                break

        revisions.extend(resp.get("revisions", []) or [])
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return revisions

def _safe_html(value):
    return html.escape(str(value or ""))

def normalize_for_compare(text):
    text = str(text or "")
    arabic_diacritics = re.compile(r"[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED]")
    text = arabic_diacritics.sub("", text)
    text = re.sub(r"[\u0640]", "", text)
    text = re.sub(r"[إأآا]", "ا", text)
    text = re.sub(r"ى", "ي", text)
    text = re.sub(r"ة", "ه", text)
    text = re.sub(r"[\s\u00A0]+", " ", text)
    text = re.sub(r"[.,;:!?؟،؛\"'“”‘’()\[\]{}<>]+", "", text)
    return text.strip().lower()

def split_sentences_smart(text):
    chunks = []
    for para in str(text or "").split("\n"):
        para = para.strip()
        if not para:
            continue
        parts = re.split(r'(?<=[.!?؟。؛])\s+|(?<=،)\s+(?=[\u0600-\u06FFA-Z0-9])', para)
        for part in parts:
            part = part.strip()
            if part:
                chunks.append(part)
    return chunks

def token_similarity(a, b):
    a_norm = normalize_for_compare(a)
    b_norm = normalize_for_compare(b)
    if not a_norm and not b_norm:
        return 1.0
    if not a_norm or not b_norm:
        return 0.0
    a_tokens = set(a_norm.split())
    b_tokens = set(b_norm.split())
    if not a_tokens or not b_tokens:
        return difflib.SequenceMatcher(None, a_norm, b_norm).ratio()
    overlap = len(a_tokens & b_tokens) / max(len(a_tokens), len(b_tokens))
    seq_ratio = difflib.SequenceMatcher(None, a_norm, b_norm).ratio()
    return max(overlap, seq_ratio)

def looks_like_formatting_only(original, revised):
    return normalize_for_compare(original) == normalize_for_compare(revised)

def _comment_artifact_terms():
    return [
        "الأفضل", "الافضل", "فالافضل", "فالأفضل", "عدلت", "تعديل", "ترجمتها", "ترجمت", "تُرجمت",
        "بتبلش", "بتبدأ", "ما الها داعي", "ما إلها داعي", "ملاحظة", "مكررة", "هون", "بالله",
        "نعدلها", "الأدق", "ادقق", "لما يكون", "بالالاف", "بالآلاف", "بالألاف", "X,000", "x,000", "comment", "note", "dining counters", "pre handover", "lap pool"
    ]

def _truncate_exported_comment_tail(raw_line):
    line = str(raw_line or "")
    lowered = line.lower()

    marker_tail_patterns = [
        r"\s+جزء\s+(?=dining counters|pre handover|lap pool|معظم|ما الها|ما إلها|تُ?رجمت|ترجمت|عدلت|كلمة|لما يكون|هون|الأفضل|الافضل)",
        r"\s+هون\s+(?=pre handover|dining counters|lap pool)",
    ]
    for pat in marker_tail_patterns:
        m = re.search(pat, line, flags=re.IGNORECASE)
        if m:
            return line[:m.start()].rstrip()

    first_pos = None
    for term in _comment_artifact_terms():
        pos = lowered.find(term.lower())
        if pos >= 0:
            if first_pos is None or pos < first_pos:
                first_pos = pos

    if first_pos is not None:
        kept = line[:first_pos].rstrip(" -–—:؛،,.\t")
        kept = re.sub(r"\s+(جزء|هون|أما|اما)$", "", kept).rstrip()
        if kept and len(kept.split()) >= 2:
            return kept
        return ""

    return line

def clean_google_doc_export_artifacts(value):
    text = str(value or "")
    text = text.replace("\ufeff", "")

    cleaned_lines = []
    comment_words = _comment_artifact_terms()

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            cleaned_lines.append(raw_line)
            continue

        markers = re.findall(r"\[[a-zA-Z]{1,3}\]", line)
        lower_line = line.lower()

        if markers and any(w.lower() in lower_line for w in comment_words):
            first_marker = re.search(r"\[[a-zA-Z]{1,3}\]", raw_line)
            if first_marker:
                kept = raw_line[:first_marker.start()].rstrip()
                kept = _truncate_exported_comment_tail(kept)
                if kept and len(kept.split()) >= 2:
                    cleaned_lines.append(kept)
                continue

        if len(markers) >= 2 and len(line.split()) > 6:
            first_marker = re.search(r"\[[a-zA-Z]{1,3}\]", raw_line)
            if first_marker:
                kept = raw_line[:first_marker.start()].rstrip()
                kept = _truncate_exported_comment_tail(kept)
                if kept and len(kept.split()) >= 2:
                    cleaned_lines.append(kept)
                continue

        raw_line = re.sub(r"\[[a-zA-Z]{1,3}\]", "", raw_line)
        raw_line = _truncate_exported_comment_tail(raw_line)
        if raw_line.strip():
            cleaned_lines.append(raw_line)

    text = "\n".join(cleaned_lines)
    text = re.sub(r"\[[a-zA-Z]{1,3}\]", "", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text.strip()

def sanitize_diff_side(value):
    value = clean_google_doc_export_artifacts(value)
    value = _truncate_exported_comment_tail(value)
    value = re.sub(r"\s+(جزء|هون|أما|اما)$", "", value).strip()
    value = re.sub(r"[ \t]{2,}", " ", value).strip()
    return value

def compute_diff(writer_text, editor_text):
    writer_text = clean_google_doc_export_artifacts(writer_text)
    editor_text = clean_google_doc_export_artifacts(editor_text)

    w_sents = split_sentences_smart(writer_text)
    e_sents = split_sentences_smart(editor_text)

    w_keys = [normalize_for_compare(x) for x in w_sents]
    e_keys = [normalize_for_compare(x) for x in e_sents]

    sm = difflib.SequenceMatcher(None, w_keys, e_keys, autojunk=False)
    changes = []

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        original = sanitize_diff_side(" ".join(w_sents[i1:i2]).strip())
        revised = sanitize_diff_side(" ".join(e_sents[j1:j2]).strip())
        if not original and not revised:
            continue
        if _looks_like_comment_artifact(original, revised):
            continue
        if not original or not revised:
            continue
        if looks_like_formatting_only(original, revised):
            continue
        sim = round(token_similarity(original, revised), 3)
        changes.append({
            "tag": tag, 
            "original": original[:700],
            "revised": revised[:700],
            "similarity": sim,
            "word_delta": len(revised.split()) - len(original.split()),
        })

    return changes

def _tokenize_for_micro_diff(text):
    text = str(text or "")
    return re.findall(r"[\u0600-\u06FFA-Za-z0-9]+(?:[-_/][\u0600-\u06FFA-Za-z0-9]+)*", text)

def _micro_context(tokens, start, end, window=5):
    left = max(0, start - window)
    right = min(len(tokens), end + window)
    return " ".join(tokens[left:right]).strip()

def _explode_change_to_micro_edits(change, lang):
    original = change.get("original", "") or ""
    revised = change.get("revised", "") or ""
    tag = change.get("tag", "")

    if tag != "replace":
        return [change]

    old_tokens = _tokenize_for_micro_diff(original)
    new_tokens = _tokenize_for_micro_diff(revised)

    if not old_tokens or not new_tokens:
        return [change]

    old_norm = [normalize_for_compare(t) for t in old_tokens]
    new_norm = [normalize_for_compare(t) for t in new_tokens]

    sm = difflib.SequenceMatcher(None, old_norm, new_norm, autojunk=False)

    edits = []

    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            continue

        old_part = " ".join(old_tokens[i1:i2]).strip()
        new_part = " ".join(new_tokens[j1:j2]).strip()

        if not old_part and not new_part:
            continue

        if looks_like_formatting_only(old_part, new_part):
            continue

        old_clean = normalize_for_compare(old_part)
        new_clean = normalize_for_compare(new_part)

        if not old_clean and not new_clean:
            continue

        if len(old_clean + new_clean) < 2:
            continue

        old_ctx = _micro_context(old_tokens, i1, i2)
        new_ctx = _micro_context(new_tokens, j1, j2)

        edits.append({
            **{
                k: v
                for k, v in change.items()
                if k not in {"original", "revised", "similarity", "word_delta", "tag"}
            },
            "tag": op,
            "original": old_part[:700],
            "revised": new_part[:700],
            "original_context": old_ctx[:700],
            "revised_context": new_ctx[:700],
            "similarity": round(token_similarity(old_part, new_part), 3),
            "word_delta": len(new_part.split()) - len(old_part.split()),
            "micro_edit": True,
        })

    return edits if edits else [change]

def explode_changes_to_micro_edits(changes, lang):
    exploded = []
    for ch in changes or []:
        exploded.extend(_explode_change_to_micro_edits(ch, lang))
    return exploded

def compute_document_level_token_edits(writer_text, editor_text, lang="Arabic"):
    writer_text = clean_google_doc_export_artifacts(writer_text)
    editor_text = clean_google_doc_export_artifacts(editor_text)

    old_tokens = _tokenize_for_micro_diff(writer_text)
    new_tokens = _tokenize_for_micro_diff(editor_text)
    if not old_tokens or not new_tokens:
        return []

    old_norm = [normalize_for_compare(t) for t in old_tokens]
    new_norm = [normalize_for_compare(t) for t in new_tokens]

    sm = difflib.SequenceMatcher(None, old_norm, new_norm, autojunk=False)
    edits = []

    def block_text(tokens):
        return " ".join([str(t or "").strip() for t in tokens if str(t or "").strip()]).strip()

    def add_block(op_name, i1, i2, j1, j2):
        old_value = block_text(old_tokens[i1:i2])
        new_value = block_text(new_tokens[j1:j2])

        if not old_value and not new_value:
            return
        if _looks_like_comment_artifact(old_value, new_value):
            return

        old_clean = normalize_for_compare(old_value)
        new_clean = normalize_for_compare(new_value)
        if not old_clean and not new_clean:
            return
        if old_clean == new_clean:
            return
        if len(old_clean + new_clean) < 2:
            return

        if looks_like_formatting_only(old_value, new_value):
            return

        old_start = max(0, min(i1, len(old_tokens)))
        new_start = max(0, min(j1, len(new_tokens)))
        old_end = max(old_start, min(i2, len(old_tokens)))
        new_end = max(new_start, min(j2, len(new_tokens)))

        old_ctx = _micro_context(old_tokens, old_start, old_end, window=7)
        new_ctx = _micro_context(new_tokens, new_start, new_end, window=7)
        if not old_ctx and not new_ctx:
            return

        edits.append({
            "tag": op_name,
            "original": old_value[:700],
            "revised": new_value[:700],
            "original_context": old_ctx[:700],
            "revised_context": new_ctx[:700],
            "parent_original": old_value[:700],
            "parent_revised": new_value[:700],
            "similarity": round(token_similarity(old_value, new_value), 3),
            "word_delta": len(new_value.split()) - len(old_value.split()),
            "micro_edit": True,
            "atomic_edit": False,
            "phrase_edit": True,
            "document_token_edit": True,
        })

    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            continue
        add_block(op, i1, i2, j1, j2)

    return edits

def _edit_occurrence_key(row):
    original = sanitize_diff_side(row.get("original", ""))
    revised = sanitize_diff_side(row.get("revised", ""))
    old_norm = normalize_for_compare(original)
    new_norm = normalize_for_compare(revised)
    old_ctx = normalize_for_compare(row.get("original_context", "") or row.get("parent_original", ""))
    new_ctx = normalize_for_compare(row.get("revised_context", "") or row.get("parent_revised", ""))

    rev_from = str(row.get("revision_from", "") or "")
    rev_to = str(row.get("revision_to", "") or "")
    rev_pair = str(row.get("revision_pair_number", "") or "")
    revision_key = (rev_from, rev_to, rev_pair) if (rev_from or rev_to or rev_pair) else ()

    if old_ctx or new_ctx:
        return (old_norm[:320], new_norm[:320], old_ctx[:900], new_ctx[:900]) + revision_key
    return (old_norm[:320], new_norm[:320]) + revision_key

def merge_diff_changes(*change_groups):
    merged = []
    seen = set()

    for group in change_groups:
        for ch in group or []:
            if not ch:
                continue
            original = sanitize_diff_side(ch.get("original", ""))
            revised = sanitize_diff_side(ch.get("revised", ""))
            if not original and not revised:
                continue
            if _looks_like_comment_artifact(original, revised):
                continue

            old_norm = normalize_for_compare(original)
            new_norm = normalize_for_compare(revised)
            if old_norm == new_norm:
                continue
            key = _edit_occurrence_key(ch)
            if key in seen:
                continue
            seen.add(key)

            nd = dict(ch)
            nd["original"] = original
            nd["revised"] = revised
            merged.append(nd)

    return merged

def changes_contain_arabic(changes):
    return any(
        re.search(r"[\u0600-\u06FF]", f"{ch.get('original', '')} {ch.get('revised', '')}")
        for ch in (changes or [])
    )

def should_use_arabic_micro_edits(changes, lang):
    return lang == "Arabic" or changes_contain_arabic(changes)

def effective_diff_language(changes, lang):
    return "Arabic" if changes_contain_arabic(changes) else lang

def classify_diff_changes(changes, platform, lang):
    if not changes:
        return []

    allowed = list(COMMENT_WEIGHTS.keys())
    items = "\n".join(
        f"[{i+1}] TAG: {c.get('tag','')} | SIMILARITY: {c.get('similarity','')} | WORD_DELTA: {c.get('word_delta','')} | MICRO: {c.get('micro_edit', False)}\n"
        f"ORIGINAL: {c.get('original','')[:500]}\nREVISED: {c.get('revised','')[:500]}\n"
        f"ORIGINAL_CONTEXT: {c.get('original_context','')[:500]}\nREVISED_CONTEXT: {c.get('revised_context','')[:500]}"
        for i, c in enumerate(changes[:80])
    )

    arabic_rules = """
Arabic-specific rules:
- Treat tashkeel, hamza style, punctuation, spacing, and light صياغة changes as grammar/arabic_language unless the meaning changed.
- If Arabic wording changes a real entity, location, developer, unit type, price, area, number, handover date, amenity, payment plan, road name, or source-backed detail, classify it as factual/source_alignment.
- If the editor deletes unsupported Arabic information, classify it as wrong_info_removed.
- If the editor adds a required source-backed detail, classify it as missing_info_added.
""" if lang == "Arabic" else ""

    prompt = f"""You are a senior editorial QA analyst for {platform}. Content language: {lang}.

An editor silently changed a writer's article without comments. Classify each change by WHY the editor likely made it.

Use ONLY these type values:
- "grammar" → grammar/spelling/minor phrasing, no factual meaning change
- "arabic_language" → Arabic grammar, إملاء, صياغة, علامات ترقيم, no factual meaning change
- "rephrase" → same facts, same meaning, smoother sentence
- "formatting" → formatting only: bullets, asterisks, line breaks, punctuation separators, spacing, list formatting; no wording/fact change
- "structural" → section moved, heading fixed, paragraph reorganized, large rewrite without clear fact correction
- "missing_info_added" → editor added important/source-based info writer missed
- "missing" → same as missing_info_added when the edit clearly adds required info
- "factual" → wrong fact corrected: number, project name, developer, price, date, location, unit type, amenity, URL, legal/source info
- "wrong_info_removed" → unsupported/wrong information deleted
- "source_alignment" → wording changed to match brochure/developer/source exactly in facts
- "contradiction_fixed" → edit fixed conflict between two parts

{arabic_rules}

Important:
- Do NOT over-penalize simple rewriting.
- Do NOT classify an edit as factual/source-related only because the paragraph contains prices, ages, tickets, AED, locations, or source-related terms.
- Classify as factual/source-related ONLY if the factual value actually changed, was removed, or was corrected.
- If tickets/prices/ages/locations remain the same and only wording changes, use rephrase/grammar.
- If the only change is adding bullets, asterisks, separators, line breaks, or punctuation around the same words, use formatting.
- If old and new facts are the same, use grammar/rephrase/arabic_language.
- If the fact changed or wrong info was removed, use factual/wrong_info_removed/source_alignment.
- For delete-only changes, decide whether it is wrong_info_removed, structural, or rephrase cleanup.
- For insert-only changes, decide whether it is missing_info_added, rephrase, structural, or grammar.

Changes:
{items}

Return ONLY raw JSON:
{{"changes": [
  {{"index": 1, "type": "factual", "severity": "high", "meaning_changed": true, "reason": "brief reason", "old_fact": "", "new_fact": ""}}
]}}

Every change must be classified. Use one of: {allowed}.
"""

    try:
        raw = call_ai(prompt)
        result = parse_json_response(raw)
        if result and "changes" in result:
            cls_map = {}
            for c in result.get("changes", []):
                try:
                    cls_map[int(c.get("index"))] = c
                except Exception:
                    continue
            classified = []
            for i, ch in enumerate(changes, 1):
                info = cls_map.get(i, {})
                ctype = str(info.get("type", "grammar")).strip()
                if ctype == "brand_voice":
                    ctype = "rephrase"
                if ctype not in COMMENT_WEIGHTS:
                    ctype = fallback_diff_type(ch, lang)
                ctype = _downgrade_false_factual_type(ch, ctype, lang)
                if ctype not in COMMENT_WEIGHTS:
                    ctype = "rephrase"
                w = COMMENT_WEIGHTS[ctype]
                severity = info.get("severity") or ("high" if ctype in HIGH_IMPACT_EDIT_TYPES else "low" if ctype in LOW_IMPACT_EDIT_TYPES else "medium")
                classified.append({
                    **ch,
                    "type": ctype,
                    "label": w["label"],
                    "deduction": w["deduction"],
                    "color": w["color"],
                    "tc": w["tc"],
                    "severity": severity,
                    "meaning_changed": bool(info.get("meaning_changed", ctype in HIGH_IMPACT_EDIT_TYPES)),
                    "reason": info.get("reason", ""),
                    "old_fact": info.get("old_fact", ""),
                    "new_fact": info.get("new_fact", ""),
                })
            return classified
    except Exception:
        pass

    classified = []
    for ch in changes:
        ctype = fallback_diff_type(ch, lang)
        ctype = _downgrade_false_factual_type(ch, ctype, lang)
        if ctype not in COMMENT_WEIGHTS:
            ctype = "rephrase"
        w = COMMENT_WEIGHTS[ctype]
        classified.append({
            **ch,
            "type": ctype,
            "label": w["label"],
            "deduction": w["deduction"],
            "color": w["color"],
            "tc": w["tc"],
            "severity": "high" if ctype in HIGH_IMPACT_EDIT_TYPES else "low" if ctype in LOW_IMPACT_EDIT_TYPES else "medium",
            "meaning_changed": ctype in HIGH_IMPACT_EDIT_TYPES,
            "reason": "Fallback classification based on edit pattern and keywords.",
            "old_fact": "",
            "new_fact": "",
        })
    return classified

def _extract_numbers_for_fact_check(value):
    return re.findall(r"\d+(?:[,.]\d+)?", str(value or ""))

def _number_sets_differ(a, b):
    return set(_extract_numbers_for_fact_check(a)) != set(_extract_numbers_for_fact_check(b))

def _semantic_tokens_for_formatting(value):
    value = str(value or "")
    value = clean_google_doc_export_artifacts(value)
    value = value.replace("•", " ").replace("*", " ").replace("|", " ")
    value = re.sub(r"\bImage\s*[-–—:]", "Image ", value, flags=re.IGNORECASE)
    return [normalize_for_compare(t) for t in _tokenize_for_micro_diff(value) if normalize_for_compare(t)]

def _only_formatting_or_separator_change(original, revised):
    old_tokens = _semantic_tokens_for_formatting(original)
    new_tokens = _semantic_tokens_for_formatting(revised)
    return bool(old_tokens or new_tokens) and old_tokens == new_tokens

def _urls_for_fact_check(value):
    return set(re.findall(r"https?://\S+|www\.\S+", str(value or "").lower()))

def _facts_meaningfully_changed(original, revised):
    original = str(original or "")
    revised = str(revised or "")
    if not original or not revised:
        return False
    if _number_sets_differ(original, revised):
        return True
    if _urls_for_fact_check(original) != _urls_for_fact_check(revised):
        return True

    old_tokens = set(_semantic_tokens_for_formatting(original))
    new_tokens = set(_semantic_tokens_for_formatting(revised))
    changed = (old_tokens ^ new_tokens)
    factual_trigger_words = {
        "aed", "price", "prices", "ticket", "tickets", "age", "ages", "criteria",
        "location", "developer", "handover", "payment", "bedroom", "bedrooms",
        "studio", "floor", "floors", "unit", "units", "villa", "villas",
        "apartment", "apartments", "sqft", "rera", "dld", "url"
    }
    combined = old_tokens | new_tokens
    if combined & factual_trigger_words and 0 < len(changed) <= 4:
        return True
    return False

def _downgrade_false_factual_type(ch, ctype, lang):
    original = ch.get("original", "") or ""
    revised = ch.get("revised", "") or ""
    tag = ch.get("tag", "")

    if _only_formatting_or_separator_change(original, revised):
        return "formatting"

    if ctype in {"factual", "source_alignment", "contradiction_fixed"}:
        if original and revised and not _facts_meaningfully_changed(original, revised):
            return fallback_diff_type(ch, lang, skip_formatting_check=True, strict_fact_guard=True)

    if ctype in {"wrong_info_removed", "missing", "missing_info_added"}:
        if original and revised and not _facts_meaningfully_changed(original, revised):
            return fallback_diff_type(ch, lang, skip_formatting_check=True, strict_fact_guard=True)

    return ctype

def _looks_like_comment_artifact(original, revised):
    both = f"{original} {revised}"
    lower = both.lower()
    markers = re.findall(r"\[[a-zA-Z]{1,3}\]", both)
    terms = _comment_artifact_terms()

    if markers and any(w.lower() in lower for w in terms):
        return True

    strong_phrases = [
        "dining counters", "pre handover", "lap pool", "بتبلش", "ما الها داعي", "ما إلها داعي",
        "نعدلها", "مكررة", "هون pre", "ترجمتها", "تُرجمت", "عدلت الترجمة", "كلمة استوديو", "لما يكون الرقم", "بالالاف", "بالآلاف", "الأفضل نوع", "افضل نوع"
    ]
    if any(p.lower() in lower for p in strong_phrases):
        return True

    return False

def fallback_diff_type(ch, lang, skip_formatting_check=False, strict_fact_guard=False):
    original = ch.get("original", "") or ""
    revised = ch.get("revised", "") or ""
    tag = ch.get("tag", "")
    both = f"{original} {revised}".lower()
    ar = re.search(r"[\u0600-\u06FF]", both) is not None
    sim = ch.get("similarity")
    try:
        sim = float(sim)
    except Exception:
        sim = token_similarity(original, revised)

    if _looks_like_comment_artifact(original, revised):
        return "formatting"

    if not skip_formatting_check and _only_formatting_or_separator_change(original, revised):
        return "formatting"

    source_keywords = ["source", "brochure", "developer", "official", "dld", "المصدر", "الكتيب", "المطور", "رسمي", "url", "http"]
    hard_fact_keywords = [
        "aed", "price", "handover", "payment", "floor", "floors", "sqft", "sq ft",
        "school", "clinic", "hospital", "metro", "mall", "airport", "minutes", "drive",
        "درهم", "السعر", "أسعار", "تسليم", "الدفع", "طابق", "قدم", "مربع",
        "مدرسة", "عيادة", "مستشفى", "مترو", "مول", "مطار", "دقيقة", "بالسيارة"
    ]

    if ar or lang == "Arabic":
        arabic_agreement_pairs = [
            ("يوجد", "توجد"), ("موجود", "موجودة"), ("متاح", "متاحة"),
            ("يشمل", "تشمل"), ("يضم", "تضم"), ("يكون", "تكون"),
            ("يعد", "تعد"), ("يعتبر", "تعتبر"),
        ]
        old_norm_text = normalize_for_compare(original)
        new_norm_text = normalize_for_compare(revised)
        for old_word, new_word in arabic_agreement_pairs:
            if old_word in old_norm_text.split() and new_word in new_norm_text.split():
                return "arabic_language"
            if new_word in old_norm_text.split() and old_word in new_norm_text.split():
                return "arabic_language"
        if tag == "delete" and len(original.split()) >= 8:
            if any(k in both for k in source_keywords):
                return "wrong_info_removed"
            if _number_sets_differ(original, revised) and any(k in both for k in hard_fact_keywords):
                return "wrong_info_removed"
            return "rephrase" if len(original.split()) < 25 else "structural"

        if tag == "insert" and len(revised.split()) >= 8:
            if any(k in both for k in source_keywords):
                return "missing_info_added"
            if _number_sets_differ(original, revised) and any(k in both for k in hard_fact_keywords):
                return "missing_info_added"
            return "rephrase" if len(revised.split()) < 25 else "rephrase"

        if any(k in both for k in source_keywords) and sim < 0.90 and _facts_meaningfully_changed(original, revised):
            return "source_alignment"

        if _facts_meaningfully_changed(original, revised) and any(k in both for k in hard_fact_keywords):
            return "factual"

        if sim >= 0.82:
            return "arabic_language"
        if sim >= 0.62:
            return "rephrase"
        if abs(len(revised.split()) - len(original.split())) > 25:
            return "structural"
        return "rephrase"

    fact_keywords = [
        "aed", "price", "handover", "developer", "location", "bedroom", "studio",
        "sqft", "sq ft", "payment", "floor", "floors", "amenity", "amenities",
        "dld", "rera", "unit", "units", "villa", "apartment"
    ]

    if tag == "delete" and len(original.split()) >= 6:
        if any(k in both for k in fact_keywords + source_keywords):
            return "wrong_info_removed"
        return "structural" if len(original.split()) > 25 else "rephrase"
    if tag == "insert" and len(revised.split()) >= 6:
        if any(k in both for k in fact_keywords + source_keywords):
            return "missing_info_added"
        return "rephrase" if len(revised.split()) > 18 else "rephrase"
    if any(k in both for k in source_keywords) and sim < 0.92 and _facts_meaningfully_changed(original, revised):
        return "source_alignment"
    if any(k in both for k in fact_keywords) and sim < 0.88 and _facts_meaningfully_changed(original, revised):
        return "factual"
    if sim >= 0.90:
        return "grammar"
    if abs(len(revised.split()) - len(original.split())) > 20:
        return "structural"
    return "rephrase"

def fetch_writer_and_editor_revisions(drive_svc, creds, doc_id, editor_name, revisions):
    if not revisions or len(revisions) < 2:
        return None, None, None, None
    editor_name_lower = editor_name.strip().lower() if editor_name else ""
    writer_rev = editor_rev = None
    if editor_name_lower:
        first_idx = None
        for i, r in enumerate(revisions):
            display = r.get("lastModifyingUser", {}).get("displayName", "").lower()
            if editor_name_lower in display or display in editor_name_lower:
                if first_idx is None: first_idx = i
                editor_rev = r
        if first_idx and first_idx > 0:
            writer_rev = revisions[first_idx - 1]
    if writer_rev is None or editor_rev is None:
        writer_rev, editor_rev = revisions[0], revisions[-1]
    w = export_revision_text(drive_svc, creds, doc_id, writer_rev["id"])
    e = export_revision_text(drive_svc, creds, doc_id, editor_rev["id"])
    return w, e, writer_rev, editor_rev

def _revision_user_label(revision):
    user = revision.get("lastModifyingUser", {}) or {}
    return " ".join([
        (user.get("displayName") or "").strip(),
        (user.get("emailAddress") or "").strip(),
    ]).strip().lower()

def _normalise_revision_name(value):
    value = str(value or "").strip().lower()
    value = value.replace("’", "'").replace("‘", "'").replace("`", "'")
    value = re.sub(r"[^\w\u0600-\u06FF@.\s-]+", " ", value)
    value = re.sub(r"[._@'\-]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value

def _revision_tokens(value):
    value = _normalise_revision_name(value)
    return [t for t in value.split() if t]

def _name_matches_revision_user(name, revision):
    needle = _normalise_revision_name(name)
    label_raw = _revision_user_label(revision)
    label = _normalise_revision_name(label_raw)
    if not needle or not label:
        return False

    if needle in label or label in needle:
        return True

    needle_tokens = _revision_tokens(name)
    label_tokens = _revision_tokens(label_raw)
    if not needle_tokens or not label_tokens:
        return False

    first_token = needle_tokens[0]

    def strong_token_match(a, b):
        return bool(a and b and len(a) >= 4 and len(b) >= 4 and (a == b or a in b or b in a))

    raw_lower = str(label_raw or "").lower()
    email_match = re.search(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+", raw_lower)
    if email_match:
        local_part = email_match.group(0).split("@", 1)[0]
        local_tokens = [t for t in re.split(r"[._+\-@]+", local_part) if t]
        if local_tokens:
            shared_local = sum(
                1 for nt in needle_tokens
                if any(nt == lt or nt in lt or lt in nt for lt in local_tokens)
            )
            if shared_local >= 2:
                return True
            if any(strong_token_match(first_token, lt) for lt in local_tokens):
                return True

    matched = 0
    for nt in needle_tokens:
        if any(nt == lt or nt in lt or lt in nt for label_tokens):
            matched += 1

    if any(strong_token_match(first_token, lt) for lt in label_tokens):
        return True

    if len(needle_tokens) >= 3:
        first_ok = any(needle_tokens[0] == lt or needle_tokens[0] in lt or lt in needle_tokens[0] for lt in label_tokens)
        last_ok = any(nt == lt or nt in lt or lt in nt for nt in needle_tokens[1:] for lt in label_tokens)
        if first_ok and last_ok:
            return True
        return matched >= 2

    return matched >= min(len(needle_tokens), 2)

def _current_file_matches_revision_user(name, current_file_meta):
    if not current_file_meta:
        return False
    user = current_file_meta.get("lastModifyingUser", {}) or {}
    if not user:
        return False
    fake_revision = {"lastModifyingUser": user}
    return _name_matches_revision_user(name, fake_revision)

def _current_file_revision_stub(current_file_meta):
    current_file_meta = current_file_meta or {}
    return {
        "id": "current_google_doc_text",
        "modifiedTime": current_file_meta.get("modifiedTime", "current"),
        "lastModifyingUser": current_file_meta.get("lastModifyingUser", {}) or {},
        "is_current_doc_text": True,
    }

def fetch_editor_handoff_revisions(drive_svc, creds, doc_id, writer_name, editor_name, revisions, current_text=None, current_file_meta=None):
    if not revisions:
        return None, None, None, None, "not_enough_revisions"

    ordered = sorted(revisions, key=_rev_sort_key_global)
    writer_matches = [i for i, r in enumerate(ordered) if _name_matches_revision_user(writer_name, r)]
    editor_matches = [i for i, r in enumerate(ordered) if _name_matches_revision_user(editor_name, r)]

    current_matches_editor = bool(current_text) and _current_file_matches_revision_user(editor_name, current_file_meta)
    current_can_be_editor_final = bool(current_text) and current_matches_editor

    if not writer_matches and not editor_matches and not current_can_be_editor_final:
        return None, None, None, None, "writer_and_editor_not_found_in_revisions"
    if not writer_matches:
        return None, None, None, None, "writer_not_found_in_revisions"

    if len(ordered) == 1 and current_text:
        writer_rev = ordered[writer_matches[-1]]
        writer_text = export_revision_text(drive_svc, creds, doc_id, writer_rev.get("id"))
        editor_text = current_text
        editor_rev = _current_file_revision_stub(current_file_meta)
        if writer_text and editor_text and normalize_for_compare(writer_text) != normalize_for_compare(editor_text):
            return writer_text, editor_text, writer_rev, editor_rev, "single_writer_revision_vs_current_doc"
        return None, None, writer_rev, editor_rev, "single_revision_same_as_current_doc"

    if current_can_be_editor_final:
        writer_handoff_idx = _select_writer_handoff_index_for_current_doc(ordered, writer_matches, current_file_meta)
        if writer_handoff_idx is None:
            return None, None, None, None, "writer_not_found_before_current_editor_doc"
        writer_rev = ordered[writer_handoff_idx]
        writer_text = export_revision_text(drive_svc, creds, doc_id, writer_rev["id"])
        editor_text = current_text
        editor_rev = _current_file_revision_stub(current_file_meta)
        if not writer_text or not editor_text:
            return None, None, writer_rev, editor_rev, "could_not_export_writer_revision_or_current_doc"
        if current_matches_editor:
            return writer_text, editor_text, writer_rev, editor_rev, "editor_session_writer_handoff_vs_current_editor_doc"
        return writer_text, editor_text, writer_rev, editor_rev, "editor_not_in_drive_revisions_used_current_doc_text"

    if not editor_matches:
        if current_text:
            writer_handoff_idx = _select_writer_handoff_index_for_current_doc(ordered, writer_matches, current_file_meta)
            if writer_handoff_idx is not None:
                writer_rev = ordered[writer_handoff_idx]
                writer_text = export_revision_text(drive_svc, creds, doc_id, writer_rev.get("id"))
                editor_text = current_text
                editor_rev = _current_file_revision_stub(current_file_meta)
                if writer_text and editor_text:
                    return writer_text, editor_text, writer_rev, editor_rev, "editor_not_in_drive_revisions_used_current_doc_proxy"
        return None, None, None, None, "editor_not_found_in_revisions"

    candidates = []

    for editor_start_idx in editor_matches:
        writer_before = [i for i in writer_matches if i < editor_start_idx]
        if not writer_before:
            continue

        writer_handoff_idx = writer_before[-1]
        next_writer_after_editor_start = next(
            (i for i in writer_matches if i > editor_start_idx),
            None,
        )

        if next_writer_after_editor_start is None:
            session_editor_indices = [i for i in editor_matches if i >= editor_start_idx]
        else:
            session_editor_indices = [
                i for i in editor_matches
                if editor_start_idx <= i < next_writer_after_editor_start
            ]

        if not session_editor_indices:
            continue

        editor_final_idx = session_editor_indices[-1]
        candidates.append({
            "writer_idx": writer_handoff_idx,
            "editor_start_idx": editor_start_idx,
            "editor_idx": editor_final_idx,
            "writer_return_idx": next_writer_after_editor_start,
        })

    if not candidates:
        return None, None, None, None, "no_writer_handoff_before_editor_session"

    chosen = sorted(candidates, key=lambda x: (x["editor_idx"], x["editor_start_idx"]))[-1]

    writer_rev = ordered[chosen["writer_idx"]]
    editor_rev = ordered[chosen["editor_idx"]]

    if writer_rev.get("id") == editor_rev.get("id"):
        return None, None, None, None, "same_writer_and_editor_revision"

    writer_text = export_revision_text(drive_svc, creds, doc_id, writer_rev["id"])
    editor_text = export_revision_text(drive_svc, creds, doc_id, editor_rev["id"])

    if not writer_text or not editor_text:
        return None, None, writer_rev, editor_rev, "could_not_export_writer_or_editor_revision"

    return writer_text, editor_text, writer_rev, editor_rev, "editor_session_writer_handoff_vs_editor_final"

def fetch_latest_editor_previous_revision(drive_svc, creds, doc_id, final_editor_name, revisions):
    if not revisions or len(revisions) < 2:
        return None, None, None, None, "not_enough_revisions_for_fallback"

    ordered = sorted(revisions, key=_rev_sort_key_global)

    final_matches = [i for i, r in enumerate(ordered) if _name_matches_revision_user(final_editor_name, r)]
    if final_matches:
        final_idx = final_matches[-1]
        if final_idx > 0:
            base_idx = final_idx - 1
            base_rev = ordered[base_idx]
            final_rev = ordered[final_idx]
            base_text = export_revision_text(drive_svc, creds, doc_id, base_rev["id"])
            final_text = export_revision_text(drive_svc, creds, doc_id, final_rev["id"])
            return base_text, final_text, base_rev, final_rev, "fallback_latest_editor_vs_previous_revision"

    base_rev = ordered[-2]
    final_rev = ordered[-1]
    base_text = export_revision_text(drive_svc, creds, doc_id, base_rev["id"])
    final_text = export_revision_text(drive_svc, creds, doc_id, final_rev["id"])
    if base_text and final_text and normalize_for_compare(base_text) != normalize_for_compare(final_text):
        return base_text, final_text, base_rev, final_rev, "fallback_latest_vs_previous_revision"

    base_rev = ordered[0]
    final_rev = ordered[-1]
    if base_rev.get("id") != final_rev.get("id"):
        base_text = export_revision_text(drive_svc, creds, doc_id, base_rev["id"])
        final_text = export_revision_text(drive_svc, creds, doc_id, final_rev["id"])
        return base_text, final_text, base_rev, final_rev, "fallback_oldest_vs_latest_revision"

    return None, None, None, None, "fallback_no_comparable_revisions"

def _revision_user_matches_editor(revision, editor_name):
    editor_name_lower = (editor_name or "").strip().lower()
    if not editor_name_lower:
        return True
    user = revision.get("lastModifyingUser", {}) or {}
    display = (user.get("displayName") or "").strip().lower()
    email = (user.get("emailAddress") or "").strip().lower()
    candidates = [display, email]
    return any(c and (editor_name_lower in c or c in editor_name_lower) for c in candidates)

def _dedupe_diff_changes(changes):
    cleaned = []
    seen = set()
    for ch in changes or []:
        original = sanitize_diff_side(ch.get("original", ""))
        revised = sanitize_diff_side(ch.get("revised", ""))
        if not original and not revised:
            continue
        if _looks_like_comment_artifact(original, revised):
            continue

        old_norm = normalize_for_compare(original)
        new_norm = normalize_for_compare(revised)
        if old_norm == new_norm:
            continue

        nd = dict(ch)
        nd["original"] = original
        nd["revised"] = revised
        key = _edit_occurrence_key(nd)
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(nd)
    return cleaned

def _dedupe_classified_diff_edits(diff_classified):
    cleaned = []
    seen = {}
    for d in diff_classified or []:
        if d.get("type") in EVENT_ONLY_EDIT_TYPES:
            continue

        original = sanitize_diff_side(d.get("original", ""))
        revised = sanitize_diff_side(d.get("revised", ""))
        if not original and not revised:
            continue
        if _looks_like_comment_artifact(original, revised):
            continue

        old_norm = normalize_for_compare(original)
        new_norm = normalize_for_compare(revised)
        if old_norm == new_norm:
            continue

        nd = dict(d)
        nd["original"] = original
        nd["revised"] = revised
        key = _edit_occurrence_key(nd)

        if key in seen:
            existing_index = seen[key]
            try:
                old_deduction = float(cleaned[existing_index].get("deduction", 0) or 0)
            except Exception:
                old_deduction = 0.0
            try:
                new_deduction = float(nd.get("deduction", 0) or 0)
            except Exception:
                new_deduction = 0.0
            if new_deduction > old_deduction:
                cleaned[existing_index] = nd
            continue

        seen[key] = len(cleaned)
        cleaned.append(nd)
    return cleaned

def _split_arabic_large_changes(changes):
    refined = []
    for ch in changes or []:
        original = ch.get("original", "") or ""
        revised = ch.get("revised", "") or ""
        has_arabic = re.search(r"[\u0600-\u06FF]", original + revised) is not None
        if ch.get("tag") != "replace" or not has_arabic:
            refined.append(ch)
            continue
        if max(len(original.split()), len(revised.split())) < 22:
            refined.append(ch)
            continue
        o_parts = [x.strip() for x in re.split(r"[،؛.؟!\n]+", original) if len(x.strip().split()) >= 3]
        r_parts = [x.strip() for x in re.split(r"[،؛.؟!\n]+", revised) if len(x.strip().split()) >= 3]
        if len(o_parts) <= 1 or len(r_parts) <= 1:
            refined.append(ch)
            continue
        sm = difflib.SequenceMatcher(
            None,
            [normalize_for_compare(x) for x in o_parts],
            [normalize_for_compare(x) for x in r_parts],
            autojunk=False,
        )
        local = []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                continue
            o = "، ".join(o_parts[i1:i2]).strip()
            r = "، ".join(r_parts[j1:j2]).strip()
            if not o and not r:
                continue
            if looks_like_formatting_only(o, r):
                continue
            local.append({
                "tag": tag,
                "original": o[:700],
                "revised": r[:700],
                "similarity": round(token_similarity(o, r), 3),
                "word_delta": len(r.split()) - len(o.split()),
            })
        refined.extend(local if local else [ch])
    return refined

def _rev_sort_key_global(item):
    return item.get("modifiedTime", "") or item.get("id", "") or ""

def _revision_display_name(revision):
    return ((revision.get("lastModifyingUser", {}) or {}).get("displayName", "") or "").strip()

def get_revision_activity_events(revisions, editor_name=""):
    if not revisions:
        return []
    editor_name_lower = (editor_name or "").strip().lower()
    ordered = sorted(revisions, key=_rev_sort_key_global)
    events = []
    for idx, r in enumerate(ordered, 1):
        display = _revision_display_name(r)
        display_lower = display.lower()
        if editor_name_lower:
            if not (editor_name_lower in display_lower or display_lower in editor_name_lower):
                continue
        events.append({
            "revision_id": r.get("id", ""),
            "revision_time": r.get("modifiedTime", ""),
            "revision_user": display,
            "revision_index": idx,
        })
    return events

def _activity_sort_key_global(item):
    return item.get("revision_time", "") or item.get("revision_id", "") or ""

def fetch_drive_activity_edit_events(creds, doc_id, editor_name=""):
    events = []
    try:
        activity = build("driveactivity", "v2", credentials=creds)
        page_token = None
        editor_name_lower = (editor_name or "").strip().lower()
        while True:
            body = {
                "itemName": f"items/{doc_id}",
                "pageSize": 100,
                "filter": "detail.action_detail_case:EDIT",
            }
            if page_token:
                body["pageToken"] = page_token
            resp = activity.activity().query(body=body).execute()
            for idx, act in enumerate(resp.get("activities", []), 1):
                detail = act.get("primaryActionDetail", {}) or {}
                if "edit" not in detail:
                    continue
                actor_names = []
                for actor in act.get("actors", []) or []:
                    user = actor.get("user", {}) or {}
                    known = user.get("knownUser", {}) or {}
                    name = (known.get("personName") or known.get("displayName") or "").strip()
                    if name:
                        actor_names.append(name)
                    elif user.get("unknownUser") is not None:
                        actor_names.append("Unknown user")
                    elif user.get("deletedUser") is not None:
                        actor_names.append("Deleted user")
                actor_label = ", ".join(actor_names) if actor_names else "Drive Activity editor"

                comparable = actor_label and actor_label != "Drive Activity editor"
                if editor_name_lower and comparable:
                    al = actor_label.lower()
                    if not (editor_name_lower in al or al in editor_name_lower):
                        continue

                when = act.get("timestamp") or (act.get("timeRange", {}) or {}).get("endTime") or (act.get("timeRange", {}) or {}).get("startTime") or ""
                events.append({
                    "revision_id": f"activity:{when}:{len(events)+1}",
                    "revision_time": when,
                    "revision_user": actor_label,
                    "revision_index": len(events)+1,
                    "activity_source": "drive_activity_api",
                })
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
    except Exception:
        return []

    seen = set()
    clean = []
    for ev in sorted(events, key=_activity_sort_key_global):
        key = (ev.get("revision_id"), ev.get("revision_time"), ev.get("revision_user"))
        if key in seen:
            continue
        seen.add(key)
        ev["revision_index"] = len(clean) + 1
        clean.append(ev)
    return clean

def parse_pasted_google_docs_version_history(raw_text, editor_name=""):
    if not raw_text or not str(raw_text).strip():
        return []

    import re

    editor_name_lower = (editor_name or "").strip().lower()
    raw_lines = [ln.strip() for ln in str(raw_text).replace("\u202f", " ").replace("\xa0", " ").splitlines()]
    lines = [ln for ln in raw_lines if ln]

    date_re = re.compile(
        r"^(?:[A-Za-z]+,?\s+)?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{1,2}:\d{2}\s*(?:AM|PM|am|pm)$"
    )

    skip_labels = {
        "current version", "version history", "all versions", "thursday", "friday",
        "saturday", "sunday", "monday", "tuesday", "wednesday", "today", "yesterday"
    }

    events = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if date_re.match(line):
            revision_time = line
            j = i + 1
            while j < len(lines) and lines[j].strip().lower() in skip_labels:
                j += 1
            if j < len(lines):
                user = lines[j].strip()
                ul = user.lower()
                if not date_re.match(user) and user.lower() not in skip_labels:
                    if not editor_name_lower or editor_name_lower in ul or ul in editor_name_lower:
                        events.append({
                            "revision_id": f"manual:{len(events)+1}:{revision_time}:{user}",
                            "revision_time": revision_time,
                            "revision_user": user,
                            "revision_index": len(events)+1,
                            "activity_source": "manual_google_docs_version_history_paste",
                        })
                    i = j + 1
                    continue
        i += 1

    return events

def make_manual_revision_count_events(count, editor_name="", source="manual_visible_revision_count_override"):
    try:
        count = int(count or 0)
    except Exception:
        count = 0
    if count <= 0:
        return []
    user = (editor_name or "Selected editor").strip() or "Selected editor"
    return [{
        "revision_id": f"manual_count:{i}",
        "revision_time": "visible Google Docs version-history row",
        "revision_user": user,
        "revision_index": i,
        "activity_source": source,
    } for i in range(1, count + 1)]

def count_editor_rows_from_paste(raw_text, editor_name=""):
    parsed_events = parse_pasted_google_docs_version_history(raw_text, editor_name)
    parsed_count = len(parsed_events)
    if parsed_count:
        return parsed_count, parsed_events

    import re
    raw = str(raw_text or "")
    name = (editor_name or "").strip()
    if not raw.strip() or not name:
        return 0, []
    pattern = re.compile(r"(?im)^\s*" + re.escape(name) + r"\s*$")
    count = len(pattern.findall(raw.replace("\u202f", " ").replace("\xa0", " ")))
    return count, make_manual_revision_count_events(count, name, "manual_google_docs_version_history_name_count")

def add_revision_event_visibility(diff_changes, revision_events):
    diff_changes = list(diff_changes or [])
    revision_events = list(revision_events or [])
    changed_to_ids = {str(ch.get("revision_to", "")) for ch in diff_changes if ch.get("revision_to")}
    existing_event_ids = {str(ch.get("revision_to", "")) for ch in diff_changes if ch.get("type") == "revision_event"}
    for ev in revision_events:
        rid = str(ev.get("revision_id", ""))
        if not rid or rid in changed_to_ids or rid in existing_event_ids:
            continue
        diff_changes.append({
            "tag": "revision_event",
            "type": "revision_event",
            "label": "Revision save event",
            "deduction": 0.0,
            "color": "#f8fafc",
            "tc": "#64748b",
            "severity": "event-only",
            "meaning_changed": False,
            "original": "",
            "revised": f"Revision saved by {ev.get('revision_user','Unknown')} at {ev.get('revision_time','')}. Google export did not expose a separate plain-text before/after change for this save.",
            "reason": "Counted from Google Docs version history. No score deduction because no exact text diff was available from the API export.",
            "revision_to": rid,
            "revision_time": ev.get("revision_time", ""),
            "revision_user": ev.get("revision_user", ""),
            "revision_pair_number": ev.get("revision_index", 0),
        })
    return diff_changes

def compute_consecutive_revision_diffs(drive_svc, creds, doc_id, editor_name, revisions, lang):
    if not revisions or len(revisions) < 2:
        return [], None, None

    def _rev_sort_key(item):
        return item.get("modifiedTime", "") or item.get("id", "") or ""

    ordered_revs = sorted(revisions, key=_rev_sort_key)

    exported = []
    for rev in ordered_revs:
        txt = export_revision_text(drive_svc, creds, doc_id, rev.get("id"))
        if txt and txt.strip():
            exported.append((rev, txt))

    if len(exported) < 2:
        return [], None, None

    all_changes = []
    changed_revision_pairs = 0

    for idx in range(1, len(exported)):
        prev_rev, prev_text = exported[idx - 1]
        cur_rev, cur_text = exported[idx]

        if normalize_for_compare(prev_text) == normalize_for_compare(cur_text):
            continue

        changes = compute_diff(prev_text, cur_text)
        if not changes:
            continue

        changed_revision_pairs += 1
        for ch in changes:
            ch["revision_from"] = prev_rev.get("id")
            ch["revision_to"] = cur_rev.get("id")
            ch["revision_time"] = cur_rev.get("modifiedTime", "")
            ch["revision_user"] = (cur_rev.get("lastModifyingUser", {}) or {}).get("displayName", "")
            ch["revision_pair_number"] = changed_revision_pairs
        all_changes.extend(changes)

    if should_use_arabic_micro_edits(all_changes, lang):
        all_changes = _split_arabic_large_changes(all_changes)
        all_changes = explode_changes_to_micro_edits(all_changes, "Arabic")

    all_changes = _dedupe_diff_changes(all_changes)

    if not all_changes:
        first_rev, first_text = exported[0]
        last_rev, last_text = exported[-1]
        fallback = compute_diff(first_text, last_text)
        if should_use_arabic_micro_edits(fallback, lang):
            fallback = _split_arabic_large_changes(fallback)
            fallback = explode_changes_to_micro_edits(fallback, "Arabic")
        return fallback, first_rev, last_rev

    return all_changes, exported[0][0], exported[-1][0]

def _revision_id(value):
    return str((value or {}).get("id", ""))

def _same_revision_id(a, b):
    return bool(a) and bool(b) and str(a) == str(b)

def _same_google_user(user_a, user_b):
    user_a = user_a or {}
    user_b = user_b or {}
    email_a = str(user_a.get("emailAddress", "")).strip().lower()
    email_b = str(user_b.get("emailAddress", "")).strip().lower()
    if email_a and email_b and email_a == email_b:
        return True
    name_a = _normalise_revision_name(user_a.get("displayName", ""))
    name_b = _normalise_revision_name(user_b.get("displayName", ""))
    return bool(name_a and name_b and (name_a == name_b or name_a in name_b or name_b in name_a))

def _current_revision_index_in_ordered(ordered, current_file_meta):
    if not ordered or not current_file_meta:
        return None

    current_modified = str(current_file_meta.get("modifiedTime", "") or "")
    if current_modified:
        for i in range(len(ordered) - 1, -1, -1):
            rev_modified = str(ordered[i].get("modifiedTime", "") or "")
            if rev_modified and rev_modified[:19] == current_modified[:19]:
                return i

    latest = ordered[-1]
    if _same_google_user(
        (latest.get("lastModifyingUser", {}) or {}),
        (current_file_meta.get("lastModifyingUser", {}) or {}),
    ):
        return len(ordered) - 1

    return None

def _select_writer_handoff_index_for_current_doc(ordered, writer_matches, current_file_meta):
    if not writer_matches:
        return None

    current_idx = _current_revision_index_in_ordered(ordered, current_file_meta)
    if current_idx is not None:
        before_current = [i for i in writer_matches if i < current_idx]
        if before_current:
            return before_current[-1]

    try:
        current_user = (current_file_meta or {}).get("lastModifyingUser", {}) or {}
        latest_writer_user = (ordered[writer_matches[-1]].get("lastModifyingUser", {}) or {})
        if len(writer_matches) >= 2 and _same_google_user(latest_writer_user, current_user):
            return writer_matches[-2]
    except Exception:
        pass

    return writer_matches[-1]

def _prepare_arabic_or_normal_changes(writer_text, editor_text, lang):
    token_changes = compute_document_level_token_edits(writer_text, editor_text, lang or "English")
    if token_changes:
        return _dedupe_diff_changes(token_changes)

    paragraph_changes = compute_diff(writer_text, editor_text)
    paragraph_micro_changes = explode_changes_to_micro_edits(paragraph_changes, lang or "English")
    return _dedupe_diff_changes(paragraph_micro_changes)

def compute_editor_session_revision_diffs(drive_svc, creds, doc_id, revisions, writer_rev, editor_rev, editor_final_text, lang):
    if not revisions or not writer_rev or not editor_rev:
        return []

    ordered = sorted(revisions, key=_rev_sort_key_global)
    writer_id = _revision_id(writer_rev)
    editor_id = _revision_id(editor_rev)

    start_idx = next((i for i, r in enumerate(ordered) if _same_revision_id(r.get("id"), writer_id)), None)
    if start_idx is None:
        return []

    editor_is_current_doc = bool(editor_rev.get("is_current_doc_text")) or editor_id == "current_google_doc_text"
    if editor_is_current_doc:
        end_idx = len(ordered) - 1
    else:
        end_idx = next((i for i, r in enumerate(ordered) if _same_revision_id(r.get("id"), editor_id)), None)
        if end_idx is None:
            return []

    if end_idx < start_idx:
        return []

    exported = []
    for rev in ordered[start_idx:end_idx + 1]:
        txt = export_revision_text(drive_svc, creds, doc_id, rev.get("id"))
        if txt and txt.strip():
            if not exported or normalize_for_compare(exported[-1][1]) != normalize_for_compare(txt):
                exported.append((rev, txt))

    if editor_is_current_doc and editor_final_text and editor_final_text.strip():
        if not exported or normalize_for_compare(exported[-1][1]) != normalize_for_compare(editor_final_text):
            exported.append((editor_rev, editor_final_text))

    if len(exported) < 2:
        return []

    all_changes = []
    pair_no = 0
    for idx in range(1, len(exported)):
        prev_rev, prev_text = exported[idx - 1]
        cur_rev, cur_text = exported[idx]
        if normalize_for_compare(prev_text) == normalize_for_compare(cur_text):
            continue

        pair_changes = _prepare_arabic_or_normal_changes(prev_text, cur_text, lang)
        if not pair_changes:
            continue

        pair_no += 1
        for ch in pair_changes:
            ch["revision_from"] = prev_rev.get("id")
            ch["revision_to"] = cur_rev.get("id")
            ch["revision_time"] = cur_rev.get("modifiedTime", "")
            ch["revision_user"] = (cur_rev.get("lastModifyingUser", {}) or {}).get("displayName", "")
            ch["revision_pair_number"] = pair_no
        all_changes.extend(pair_changes)

    return _dedupe_diff_changes(all_changes)

def extract_text_from_gdoc(doc):
    text, headings = [], []
    heading_map = {"HEADING_1": "H1", "HEADING_2": "H2", "HEADING_3": "H3"}
    for element in doc.get("body", {}).get("content", []):
        if "paragraph" not in element:
            continue
        para  = element["paragraph"]
        style = para.get("paragraphStyle", {}).get("namedStyleType", "")
        raw   = "".join(
            e.get("textRun", {}).get("content", "")
            for e in para.get("elements", [])
        ).strip()
        if not raw:
            continue
        text.append(raw)
        if style in heading_map:
            headings.append({"level": heading_map[style], "text": raw})
    return "\n".join(text), headings

def extract_suggestions_from_doc(doc):
    suggestions = []
    seen_ids    = set()

    for element in doc.get("body", {}).get("content", []):
        if "paragraph" not in element:
            continue
        for elem in element["paragraph"].get("elements", []):
            if "textRun" not in elem:
                continue
            text_run = elem["textRun"]
            content  = text_run.get("content", "").strip()
            if not content or content == "\n":
                continue

            insert_ids = list(elem.get("suggestedInsertionIds", []))
            delete_ids = list(elem.get("suggestedDeletionIds", []))

            for sid in insert_ids:
                if sid not in seen_ids and content:
                    seen_ids.add(sid)
                    suggestions.append({
                        "type":          "insert",
                        "text":          content,
                        "suggestion_id": sid,
                    })
            for sid in delete_ids:
                if sid not in seen_ids and content:
                    seen_ids.add(sid)
                    suggestions.append({
                        "type":          "delete",
                        "text":          content,
                        "suggestion_id": sid,
                    })

    return suggestions

def fetch_google_doc(url):
    doc_id = extract_doc_id(url)
    if not doc_id:
        return None, "Invalid Google Doc URL"

    try:
        docs_svc, drive_svc, svc_creds = get_google_services()
    except Exception as e:
        return None, f"Google auth error: {e}"

    allowed, validation_error = validate_dubizzle_group_google_doc(drive_svc, doc_id)
    if not allowed:
        return None, validation_error

    try:
        doc   = docs_svc.documents().get(documentId=doc_id).execute()
        title = doc.get("title", "Untitled")
        text, headings = extract_text_from_gdoc(doc)
        try:
            current_file_meta = drive_svc.files().get(
                fileId=doc_id,
                fields="id,name,modifiedTime,lastModifyingUser(displayName,emailAddress)",
                supportsAllDrives=True,
            ).execute()
        except Exception:
            current_file_meta = {}
    except HttpError as e:
        return None, friendly_google_api_error(e, doc_id)
    except Exception as e:
        return None, f"Could not read doc. Share it with {get_service_account_email()} as Editor and try again. Details: {e}"

    suggestions = extract_suggestions_from_doc(doc)

    comments = []
    comments_error = ""
    try:
        all_fetched = []
        page_token  = None
        while True:
            params = dict(
                fileId=doc_id,
                fields="comments(id,author,content,resolved,replies,quotedFileContent),nextPageToken",
                includeDeleted=False,
                pageSize=100,
            )
            if page_token:
                params["pageToken"] = page_token
            resp = drive_svc.comments().list(**params).execute()
            all_fetched.extend(resp.get("comments", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break

        for c in all_fetched:
            body   = c.get("content", "").strip()
            author = c.get("author", {}).get("displayName", "Editor")
            email  = c.get("author", {}).get("emailAddress", "")
            resolved = c.get("resolved", False)
            quoted = ((c.get("quotedFileContent") or {}).get("value") or "").strip()
            if body:
                comments.append({
                    "author":   author,
                    "email":    email,
                    "text":     body,
                    "quoted":   quoted,
                    "resolved": resolved,
                })
            for r in c.get("replies", []):
                rbody = r.get("content", "").strip()
                rauth = r.get("author", {}).get("displayName", "")
                if rbody and len(rbody) > 10:
                    comments.append({
                        "author":   rauth,
                        "email":    r.get("author", {}).get("emailAddress", ""),
                        "text":     rbody,
                        "quoted":   quoted,
                        "resolved": resolved,
                        "is_reply": True,
                    })
    except Exception as e:
        comments_error = str(e)

    revisions = list_drive_revisions_safe(drive_svc, doc_id)

    return {
        "title":           title,
        "text":            text,
        "headings":        headings,
        "links":           re.findall(r'https?://\S+', text),
        "comments":        comments,
        "comments_error":  comments_error,
        "suggestions":     suggestions,
        "revisions":       revisions,
        "current_file_meta": current_file_meta,
        "revision_count_from_api": len(revisions),
        "word_count":      len(text.split()),
        "drive_svc":       drive_svc,
        "svc_creds":       svc_creds,
        "error":           "",
    }, None

def get_editor_emails():
    try:
        raw = st.secrets.get("EDITOR_EMAILS", "")
        return [e.strip().lower() for e in raw.split(",") if e.strip()]
    except Exception:
        return []

def count_revision_rounds(revisions, editor_name):
    if not revisions:
        return 0, 0, []

    editor_name_lower = editor_name.strip().lower() if editor_name else ""

    annotated = []
    for r in revisions:
        display = r.get("lastModifyingUser", {}).get("displayName", "")
        email   = r.get("lastModifyingUser", {}).get("emailAddress", "")
        is_editor = (editor_name_lower and
                     (editor_name_lower in display.lower() or
                      display.lower() in editor_name_lower))
        who = "editor" if is_editor else "writer"
        annotated.append({
            "id":      r.get("id"),
            "time":    r.get("modifiedTime", ""),
            "email":   email or display,
            "display": display,
            "who":     who,
        })

    rounds = 0
    prev   = None
    for r in annotated:
        if r["who"] == "editor" and prev != "editor":
            rounds += 1
        prev = r["who"]

    return rounds, len(revisions), annotated

def _plain_comment_text(value):
    value = html.unescape(str(value or ""))
    value = re.sub(r"<[^>]+>", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value

def _is_url_only_comment(text):
    value = _plain_comment_text(text)
    if not value:
        return True
    urls = re.findall(r"https?://\S+", value)
    if not urls:
        return False
    remainder = value
    for u in urls:
        remainder = remainder.replace(u, "")
    remainder = re.sub(r"[\s\-–—_:،,.;؛()\[\]{}]+", "", remainder).strip()
    return remainder == ""

def _comment_signal_features(comment, lang=""):
    text = _plain_comment_text(comment.get("text", "") if isinstance(comment, dict) else comment)
    quoted = _plain_comment_text(comment.get("quoted", "") if isinstance(comment, dict) else "")
    combined = f"{text} {quoted}".strip()
    low = combined.lower()

    urls = re.findall(r"https?://\S+", text)
    url_only = _is_url_only_comment(text)

    has_money_or_number = bool(re.search(r"\b(?:aed|درهم)\b|\d", low))
    source_object = bool(re.search(
        r"\b(source|listing|brochure|sheet|spreadsheet|data|dataset|official|url|link|lpv|crm|bayut|dubizzle)\b|"
        r"المصدر|مصدر|اللستنج|الليستنج|الشيت|الجدول|الداتا|البيانات|الرابط|اللينك|الرابط|رابط|رسمي|حسب",
        low
    ))
    factual_object = bool(re.search(
        r"\b(price|prices|aed|age|ticket|tickets|date|handover|location|area|name|title|reason|bedroom|bedrooms|studio|floor|sqft|payment|permit|number|unit|units|developer|project)\b|"
        r"السعر|الأسعار|اسعار|الاسعار|رسوم|التكلفة|تذكرة|تذاكر|العمر|السن|التاريخ|التسليم|الموقع|المنطقة|الاسم|العنوان|السبب|غرف|غرفة|طابق|قدم|خطة|الدفع|رقم|وحدة|وحدات|المطور|المشروع|نوع|تصريح",
        low
    )) or has_money_or_number

    wrong_signal = bool(re.search(
        r"\b(wrong|incorrect|inaccurate|not correct|false|mismatch|doesn'?t match|should be|must be|correct)\b|"
        r"غلط|خطأ|خطا|غير صحيح|مش صحيح|ليس صحيح|غير دقيق|مش دقيق|الصحيح|الصح|الأدق|ادق|بدل",
        low
    ))
    missing_signal = bool(re.search(
        r"\b(missing|where is|where are|add|include|mention|not mentioned|required|needs|need to add|please add|lacks)\b|"
        r"وين|أين|فين|ناقص|ناقصة|مفقود|مفقودة|أضف|اضف|ضيف|ضف|نضيف|لازم نضيف|يجب إضافة|اذكر|أذكر|نذكر|اكتب|أكتب|مطلوب|مطلوبة",
        low
    ))
    removal_signal = bool(re.search(
        r"\b(remove|delete|cut|unsupported|not in source|not available|not found|no source|not listed|not on listing)\b|"
        r"غير موجود|غير موجودة|مش موجود|مش موجودة|غير مذكور|غير مذكورة|احذف|احذفي|شيل|شيلي|نشيل|بدون مصدر|ليس في المصدر|مش في المصدر",
        low
    ))
    structural_signal = bool(re.search(
        r"\b(rewrite|restructure|structure|paragraph|section|heading|header|table|bullet|list|move|arrange|reorder|format as paragraph)\b|"
        r"بارجراف|فقرة|فقره|هيكلة|ترتيب|رتب|نقسم|قسم|عنوان|هيدر|جدول|نقل|انقل|مكانه|صياغة القسم",
        low
    ))
    language_signal = bool(re.search(
        r"\b(grammar|spelling|typo|punctuation|wording|language|capitalize|capitalise)\b|"
        r"إملاء|املاء|نحو|لغوي|لغوية|لغة|تشكيل|همزة|ترقيم|الفاصلة|النقطة|صياغة لغوية",
        low
    ))
    style_signal = bool(re.search(
        r"\b(rephrase|tone|style|voice|brand|too generic|generic|sounds|flow|readability)\b|"
        r"أسلوب|ستايل|نبرة|عام|عامة|صياغة|ركيك|أجمل|أفضل صياغة",
        low
    ))

    question_signal = "?" in text or "؟" in text

    return {
        "text": text,
        "quoted": quoted,
        "combined": combined,
        "url_only": url_only,
        "has_url": bool(urls),
        "source_object": source_object,
        "factual_object": factual_object,
        "wrong_signal": wrong_signal,
        "missing_signal": missing_signal,
        "removal_signal": removal_signal,
        "structural_signal": structural_signal,
        "language_signal": language_signal,
        "style_signal": style_signal,
        "question_signal": question_signal,
    }

def _logic_comment_type(comment, lang=""):
    f = _comment_signal_features(comment, lang)

    if not f["text"] or f["url_only"]:
        return "formatting", "Bare URL/source reference only."

    if f["missing_signal"] and (f["factual_object"] or f["source_object"] or f["question_signal"]):
        return "missing", "Comment asks for required source/data/details that are missing."

    if f["removal_signal"] and (f["source_object"] or f["factual_object"]):
        return "wrong_info_removed", "Comment says source/data/prices/claims are not available or should be removed."

    if f["wrong_signal"] and (f["source_object"] or f["factual_object"]):
        return "factual", "Comment corrects a source/data/name/price/reason/location detail."

    if f["source_object"] and (f["wrong_signal"] or f["factual_object"]):
        return "factual", "Comment refers to a source/listing/sheet for a factual correction."

    if f["structural_signal"] and not (f["source_object"] or f["wrong_signal"] or f["removal_signal"] or f["missing_signal"]):
        return "structural", "Comment asks for structural/paragraph/section change."

    if f["language_signal"] and not (f["source_object"] or f["factual_object"] or f["wrong_signal"] or f["missing_signal"] or f["removal_signal"]):
        return "arabic_language" if lang == "Arabic" or re.search(r"[\u0600-\u06FF]", f["combined"]) else "grammar", "Language, spelling, grammar, or punctuation only."

    if f["style_signal"] and not (f["source_object"] or f["factual_object"] or f["wrong_signal"] or f["missing_signal"] or f["removal_signal"]):
        return "rephrase", "Style, tone, or wording improvement only."

    return "", "No strong rule-based decision."

def _normalise_comment_type(ctype):
    ctype = str(ctype or "").strip().lower()
    aliases = {
        "data accuracy": "factual",
        "source": "factual",
        "source_related": "factual",
        "source-related": "factual",
        "fact": "factual",
        "incorrect": "factual",
        "wrong": "factual",
        "missing_info": "missing",
        "missing info": "missing",
        "language": "grammar",
        "phrasing": "grammar",
        "format": "formatting",
        "formatting_only": "formatting",
        "link_only": "formatting",
    }
    ctype = aliases.get(ctype, ctype)
    return ctype if ctype in COMMENT_WEIGHTS else "grammar"

def classify_comments_ai(comments, platform, lang):
    if not comments:
        return []

    c_txt = "\n".join(
        f"  [{i+1}] COMMENT: {_plain_comment_text(c.get('text',''))}\n      QUOTED TEXT: {_plain_comment_text(c.get('quoted',''))[:300]}"
        for i, c in enumerate(comments)
    )
    prompt = f"""You are a strict editorial QA classifier for {platform}. Content language: {lang}.

Classify the editorial INTENT of each Google Doc comment. Use exactly one type:
- "factual" when the comment corrects wrong source/data, wrong name, wrong reason, wrong price/date/location, or asks to align with a listing/source/sheet.
- "wrong_info_removed" when the comment says claims/prices/details are not in the source or should be removed.
- "missing" when required information is missing or the editor asks where it is.
- "structural" when the section/paragraph/header/table structure must change, without a wrong fact.
- "arabic_language" for Arabic grammar/spelling/agreement/punctuation only, no source/data issue.
- "grammar" for English grammar/punctuation/minor phrasing only, no source/data issue.
- "rephrase" for style/tone/brand voice rewrite only.
- "formatting" for a bare source link or formatting note with no correction.

Important logic:
- Do not use grammar just because the comment is short.
- A comment about source, listing, sheet, data, prices, names, reasons, or missing required details is not grammar.
- If both structure and wrong/missing data appear, classify by the data problem first.
- Bare links alone are formatting.

Comments:
{c_txt}

Return ONLY raw JSON, no markdown:
{{"classifications": [{{"index": 1, "type": "factual"}}, {{"index": 2, "type": "missing"}}]}}
"""

    ai_map = {}
    try:
        raw = call_ai(prompt)
        result = parse_json_response(raw)
        if result and "classifications" in result:
            for item in result.get("classifications", []):
                try:
                    ai_map[int(item.get("index"))] = _normalise_comment_type(item.get("type"))
                except Exception:
                    continue
    except Exception:
        ai_map = {}

    classified = []
    for i, c in enumerate(comments, 1):
        logic_type, reason = _logic_comment_type(c, lang)
        ai_type = ai_map.get(i, "")

        if logic_type:
            ctype = logic_type
        else:
            ctype = ai_type or ("arabic_language" if lang == "Arabic" else "grammar")

        ctype = _normalise_comment_type(ctype)
        w = COMMENT_WEIGHTS[ctype]
        classified.append({
            "author":    c.get("author", ""),
            "email":     c.get("email", ""),
            "text":      _plain_comment_text(c.get("text", "")),
            "quoted":    _plain_comment_text(c.get("quoted", "")),
            "type":      ctype,
            "label":     w["label"],
            "deduction": w["deduction"],
            "color":     w["color"],
            "tc":        w["tc"],
            "reason":    reason or "AI classification after intent check.",
        })
    return classified

def apply_gdoc_deductions(classified_comments, editor_rounds):
    return apply_gdoc_deductions_full(classified_comments, [], editor_rounds)

def capped_low_impact_deduction(items):
    events = [d for d in items if d.get("type") in EVENT_ONLY_EDIT_TYPES]
    high = [d for d in items if d.get("type") in HIGH_IMPACT_EDIT_TYPES]
    low = [d for d in items if d.get("type") in LOW_IMPACT_
