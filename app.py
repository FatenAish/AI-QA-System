import streamlit as st
import json
import re
import os
import urllib.request
from datetime import datetime
from io import BytesIO
import difflib
import html

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
# Silent edits are intentionally more detailed than comments so the system can
# separate harmless wording cleanup from factual corrections and source fixes.
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
}

LOW_IMPACT_EDIT_TYPES = {"rephrase", "grammar", "arabic_language"}
HIGH_IMPACT_EDIT_TYPES = {"factual", "wrong_info_removed", "source_alignment", "contradiction_fixed", "missing", "missing_info_added"}
EVENT_ONLY_EDIT_TYPES = {"revision_event"}
REVISION_ROUND_PENALTY = 0.7  # per extra round

RECORDS_FILE = "qa_records.json"

# ── Local persistence ──────────────────────────────────────────────────────
def load_records():
    if not os.path.exists(RECORDS_FILE):
        return []
    try:
        with open(RECORDS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def save_record(sub):
    records = load_records()
    key = (sub["writer"], sub["title"], sub["date"])
    for r in records:
        if (r["writer"], r["title"], r["date"]) == key:
            return
    records.append(_serialisable(sub))
    with open(RECORDS_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

def update_record_decision(sub):
    records = load_records()
    key = (sub["writer"], sub["title"], sub["date"])
    for r in records:
        if (r["writer"], r["title"], r["date"]) == key:
            r["editor_decision"] = sub.get("editor_decision", "")
            r["editor_notes"]    = sub.get("editor_notes", "")
            break
    with open(RECORDS_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

def _serialisable(obj):
    if isinstance(obj, dict):          return {k: _serialisable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)): return [_serialisable(i) for i in obj]
    if isinstance(obj, (str, int, float, bool)) or obj is None: return obj
    return str(obj)

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

    # Remove fragments/query strings first; they are not part of the file ID.
    value = value.split("#", 1)[0].split("?", 1)[0]

    # Standard Google Doc URL.
    m = re.search(r"/document/d/([a-zA-Z0-9_-]+)", value)
    if m:
        return m.group(1)

    # Sometimes users paste only the ID.
    if re.fullmatch(r"[a-zA-Z0-9_-]{20,}", value):
        return value

    return None

def clean_google_doc_url(url):
    """Return a clean canonical Google Doc URL for display/debugging."""
    doc_id = extract_doc_id(url)
    return f"https://docs.google.com/document/d/{doc_id}/edit" if doc_id else str(url or "").strip()



def get_allowed_google_doc_domains():
    """
    Allowed company Google domains for Google Doc submissions.
    Optional Streamlit secret:
    ALLOWED_GOOGLE_DOC_DOMAINS = "dubizzle.com,bayut.com,dubizzlegroup.com,bayut.jo"
    """
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
    """
    Security gate for Google Doc submissions.
    Blocks public docs, but allows private docs shared with the Content QA service account as Editor.
    """
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

    # 1) Block public docs: Anyone with the link / public web sharing.
    for p in permissions:
        if p.get("type") == "anyone":
            return False, "Public Google Docs are not allowed. Keep sharing restricted and share the doc with the Content QA service account as Editor."

    # 2) Main access rule: allow any private Google Doc shared with the Content QA service account as Editor.
    # This allows approved content-system submissions even when the owner domain is not Dubizzle/Bayut.
    if has_content_system_editor_access(permissions):
        return True, ""

    # 3) Optional company-domain fallback for older company-owned docs.
    for p in permissions:
        if p.get("type") == "domain":
            domain = str(p.get("domain") or "").strip().lower()
            if domain and domain not in allowed_domains:
                return False, f"This document is shared with an unapproved domain: {domain}. Allowed domains: {', '.join(allowed_domains)}."

    # 4) Company ownership fallback where owner email is visible.
    owner_emails = [o.get("emailAddress", "") for o in owners if o.get("emailAddress")]
    if owner_emails:
        if not any(is_allowed_company_email(email) for email in owner_emails):
            return False, f"This private doc is accessible, but the Content QA service account is not shared as Editor. Share it with {get_service_account_email()} as Editor and try again."
        return True, ""

    # 5) Shared drives can hide owners. In that case, allow only if a company-domain permission exists.
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

def export_revision_text(drive_svc, creds, doc_id, revision_id):
    """Export a specific revision as plain text."""
    try:
        import google.auth.transport.requests
        # Get export link for this revision
        rev = drive_svc.revisions().get(
            fileId=doc_id,
            revisionId=revision_id,
            fields="exportLinks"
        ).execute()
        export_url = rev.get("exportLinks", {}).get("text/plain", "")
        if not export_url:
            return None
        # Refresh credentials to get a valid token
        auth_req = google.auth.transport.requests.Request()
        if not creds.valid:
            creds.refresh(auth_req)
        req = urllib.request.Request(
            export_url,
            headers={"Authorization": f"Bearer {creds.token}"}
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.read().decode("utf-8")
    except Exception as e:
        return None

def _safe_html(value):
    return html.escape(str(value or ""))

def normalize_for_compare(text):
    """Normalize English/Arabic text for fair diff comparison."""
    text = str(text or "")
    arabic_diacritics = re.compile(r"[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED]")
    text = arabic_diacritics.sub("", text)
    text = re.sub(r"[\u0640]", "", text)  # tatweel
    text = re.sub(r"[إأآا]", "ا", text)
    text = re.sub(r"ى", "ي", text)
    text = re.sub(r"ة", "ه", text)
    text = re.sub(r"[\s\u00A0]+", " ", text)
    text = re.sub(r"[.,;:!?؟،؛\"'“”‘’()\[\]{}<>]+", "", text)
    return text.strip().lower()

def split_sentences_smart(text):
    """Sentence splitter that works for English and Arabic punctuation."""
    chunks = []
    for para in str(text or "").split("\n"):
        para = para.strip()
        if not para:
            continue
        # Handles English punctuation and Arabic full stop/question/comma-like boundaries.
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
    """Terms that usually appear only inside exported Google Docs comments/notes."""
    return [
        "الأفضل", "الافضل", "فالافضل", "فالأفضل", "عدلت", "تعديل", "ترجمتها", "ترجمت", "تُرجمت",
        "بتبلش", "بتبدأ", "ما الها داعي", "ما إلها داعي", "ملاحظة", "مكررة", "هون", "بالله",
        "نعدلها", "الأدق", "ادقق", "لما يكون", "بالالاف", "بالآلاف", "بالألاف", "X,000", "x,000", "comment", "note", "dining counters", "pre handover", "lap pool"
    ]


def _truncate_exported_comment_tail(raw_line):
    """
    Google Docs export can append inline comment text to the article line, sometimes
    after removing [a]/[b] markers. Keep the article content before the first clear
    comment phrase and drop the comment tail.
    """
    line = str(raw_line or "")
    lowered = line.lower()

    # Patterns where a removed marker leaves a connector word before the comment text,
    # e.g. "* دبي مارينا ووك جزء dining counters...".
    marker_tail_patterns = [
        r"\s+جزء\s+(?=dining counters|pre handover|lap pool|معظم|ما الها|ما إلها|تُ?رجمت|ترجمت|عدلت|كلمة|لما يكون|هون|الأفضل|الافضل)",
        r"\s+هون\s+(?=pre handover|dining counters|lap pool)",
    ]
    for pat in marker_tail_patterns:
        m = re.search(pat, line, flags=re.IGNORECASE)
        if m:
            return line[:m.start()].rstrip()

    # Direct comment trigger terms.
    first_pos = None
    for term in _comment_artifact_terms():
        pos = lowered.find(term.lower())
        if pos >= 0:
            if first_pos is None or pos < first_pos:
                first_pos = pos

    if first_pos is not None:
        # Keep only meaningful article text before the comment text.
        kept = line[:first_pos].rstrip(" -–—:؛،,.\t")
        # Remove orphan connector words that often remain after marker removal.
        kept = re.sub(r"\s+(جزء|هون|أما|اما)$", "", kept).rstrip()
        if kept and len(kept.split()) >= 2:
            return kept
        return ""

    return line


def clean_google_doc_export_artifacts(value):
    """
    Remove Google Docs exported comment/reference artifacts before diffing.
    Google exports sometimes inline comments as [a], [b] markers and appends
    editor notes after the final article text. These are not silent text edits.
    """
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

        # Full exported comment/note line, not article content. Keep only article
        # text before the first marker, then run tail truncation on that piece too.
        if markers and any(w.lower() in lower_line for w in comment_words):
            first_marker = re.search(r"\[[a-zA-Z]{1,3}\]", raw_line)
            if first_marker:
                kept = raw_line[:first_marker.start()].rstrip()
                kept = _truncate_exported_comment_tail(kept)
                if kept and len(kept.split()) >= 2:
                    cleaned_lines.append(kept)
                continue

        # A dense marker line is usually an exported comment block.
        if len(markers) >= 2 and len(line.split()) > 6:
            first_marker = re.search(r"\[[a-zA-Z]{1,3}\]", raw_line)
            if first_marker:
                kept = raw_line[:first_marker.start()].rstrip()
                kept = _truncate_exported_comment_tail(kept)
                if kept and len(kept.split()) >= 2:
                    cleaned_lines.append(kept)
                continue

        # Normal article line: remove reference markers only, then remove any
        # trailing comment text that survived without markers.
        raw_line = re.sub(r"\[[a-zA-Z]{1,3}\]", "", raw_line)
        raw_line = _truncate_exported_comment_tail(raw_line)
        if raw_line.strip():
            cleaned_lines.append(raw_line)

    text = "\n".join(cleaned_lines)
    text = re.sub(r"\[[a-zA-Z]{1,3}\]", "", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text.strip()

def sanitize_diff_side(value):
    """
    Clean one already-split diff side. This is stricter than document-level cleanup
    because Google export can attach comment text to a single sentence/bullet after
    sentence splitting.
    """
    value = clean_google_doc_export_artifacts(value)
    value = _truncate_exported_comment_tail(value)
    # Remove dangling connector words left by comment-marker cleanup.
    value = re.sub(r"\s+(جزء|هون|أما|اما)$", "", value).strip()
    value = re.sub(r"[ \t]{2,}", " ", value).strip()
    return value

def compute_diff(writer_text, editor_text):
    """
    Diff two versions of the doc at sentence/paragraph level.
    Returns list of {tag, original, revised, similarity} dicts.
    Arabic-aware: ignores tashkeel, tatweel and punctuation-only changes.
    """
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
            # Exported Google Docs comment text must never be counted as a silent edit.
            continue
        # If the only thing left on one side was a comment artifact, do not create
        # a fake delete/insert edit from comment cleanup.
        if not original or not revised:
            continue
        if looks_like_formatting_only(original, revised):
            # Formatting-only/tashkeel-only changes are kept only if meaningful enough.
            continue
        sim = round(token_similarity(original, revised), 3)
        changes.append({
            "tag": tag,  # replace / delete / insert
            "original": original[:700],
            "revised": revised[:700],
            "similarity": sim,
            "word_delta": len(revised.split()) - len(original.split()),
        })

    return changes


def _tokenize_for_micro_diff(text):
    """Tokenize Arabic/English text for micro silent-edit detection."""
    text = str(text or "")
    # Keep words/numbers as tokens and keep punctuation only as separators.
    # Arabic range + English letters + numbers. This intentionally ignores punctuation-only edits.
    return re.findall(r"[\u0600-\u06FFA-Za-z0-9]+(?:[-_/][\u0600-\u06FFA-Za-z0-9]+)*", text)


def _micro_context(tokens, start, end, window=5):
    """Small readable context around a token-level edit."""
    left = max(0, start - window)
    right = min(len(tokens), end + window)
    return " ".join(tokens[left:right]).strip()


def _explode_change_to_micro_edits(change, lang):
    """
    Break Arabic paragraph-level replace blocks into individual word/phrase edits.

    The old version was too conservative and could keep short Arabic paragraph
    replacements grouped as one edit. This version explodes every valid Arabic
    replace block into token-level edits whenever a clean micro edit is detected.
    """
    original = change.get("original", "") or ""
    revised = change.get("revised", "") or ""
    tag = change.get("tag", "")

    has_arabic = re.search(r"[\u0600-\u06FF]", original + revised) is not None

    # Only explode replace blocks for Arabic content.
    if tag != "replace" or (lang != "Arabic" and not has_arabic):
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

        # Ignore empty/very tiny noise after normalization.
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

    # Important: if even one valid micro edit is found, use it.
    # Do not fall back to paragraph-level grouping unless no micro edits exist.
    return edits if edits else [change]

def explode_changes_to_micro_edits(changes, lang):
    """Apply micro-edit splitting to a list of diff changes."""
    exploded = []
    for ch in changes or []:
        exploded.extend(_explode_change_to_micro_edits(ch, lang))
    return exploded


def changes_contain_arabic(changes):
    """Return True when any diff change contains Arabic text, regardless of UI language selection."""
    return any(
        re.search(r"[\u0600-\u06FF]", f"{ch.get('original', '')} {ch.get('revised', '')}")
        for ch in (changes or [])
    )

def should_use_arabic_micro_edits(changes, lang):
    """Use Arabic micro splitting when the selected language is Arabic OR the actual diff text is Arabic."""
    return lang == "Arabic" or changes_contain_arabic(changes)

def effective_diff_language(changes, lang):
    """Classify Arabic diffs with Arabic rules even if the UI language was accidentally set to English."""
    return "Arabic" if changes_contain_arabic(changes) else lang

def classify_diff_changes(changes, platform, lang):
    """
    Classify silent edits by editorial reason, not just by size.
    Returns each change with label, deduction, severity and explanation.
    """
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
                if ctype == "formatting":
                    ctype = "grammar"
                if ctype not in COMMENT_WEIGHTS:
                    ctype = fallback_diff_type(ch, lang)
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

def _looks_like_comment_artifact(original, revised):
    both = f"{original} {revised}"
    lower = both.lower()
    markers = re.findall(r"\[[a-zA-Z]{1,3}\]", both)
    terms = _comment_artifact_terms()

    # Marker + comment term is definitely exported comment material.
    if markers and any(w.lower() in lower for w in terms):
        return True

    # Even without markers, exported notes often survive as English/Arabic editor talk.
    strong_phrases = [
        "dining counters", "pre handover", "lap pool", "بتبلش", "ما الها داعي", "ما إلها داعي",
        "نعدلها", "مكررة", "هون pre", "ترجمتها", "تُرجمت", "عدلت الترجمة", "كلمة استوديو", "لما يكون الرقم", "بالالاف", "بالآلاف", "الأفضل نوع", "افضل نوع"
    ]
    if any(p.lower() in lower for p in strong_phrases):
        return True

    return False

def fallback_diff_type(ch, lang):
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
        return "grammar"

    source_keywords = ["source", "brochure", "developer", "official", "dld", "المصدر", "الكتيب", "المطور", "رسمي", "url", "http"]
    hard_fact_keywords = [
        "aed", "price", "handover", "payment", "floor", "floors", "sqft", "sq ft",
        "school", "clinic", "hospital", "metro", "mall", "airport", "minutes", "drive",
        "درهم", "السعر", "أسعار", "تسليم", "الدفع", "طابق", "قدم", "مربع",
        "مدرسة", "عيادة", "مستشفى", "مترو", "مول", "مطار", "دقيقة", "بالسيارة"
    ]

    # Arabic edits are usually translation/phrasing unless a measurable/source-backed fact changes.
    if ar or lang == "Arabic":
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

        if any(k in both for k in source_keywords) and sim < 0.90:
            return "source_alignment"

        # Only mark factual when numbers or clearly measurable details changed.
        if _number_sets_differ(original, revised) and any(k in both for k in hard_fact_keywords):
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
    if any(k in both for k in source_keywords) and sim < 0.92:
        return "source_alignment"
    if any(k in both for k in fact_keywords) and sim < 0.88:
        return "factual"
    if sim >= 0.90:
        return "grammar"
    if abs(len(revised.split()) - len(original.split())) > 20:
        return "structural"
    return "rephrase"

def fetch_writer_and_editor_revisions(drive_svc, creds, doc_id, editor_name, revisions):
    """Diff writer vs editor version. Matches by display name, falls back to first vs last."""
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
    """Return the best available user label from a Google Drive revision."""
    user = revision.get("lastModifyingUser", {}) or {}
    return " ".join([
        (user.get("displayName") or "").strip(),
        (user.get("emailAddress") or "").strip(),
    ]).strip().lower()


def _name_matches_revision_user(name, revision):
    """Tolerant matching for display names/emails in Drive revisions."""
    needle = (name or "").strip().lower()
    if not needle:
        return False
    label = _revision_user_label(revision)
    if not label:
        return False
    needle_tokens = [t for t in re.split(r"[\s._@-]+", needle) if t]
    label_tokens = [t for t in re.split(r"[\s._@-]+", label) if t]
    if needle in label or label in needle:
        return True
    # Allow partial Arabic/English names such as "Aya" or "Mohammad Salem".
    return bool(needle_tokens) and all(any(nt in lt or lt in nt for lt in label_tokens) for nt in needle_tokens)


def fetch_editor_handoff_revisions(drive_svc, creds, doc_id, base_editor_name, final_editor_name, revisions):
    """
    Compare the last version by the writer/base editor with the last version by
    the final editor/reviewer.

    This is more reliable for QA than counting every autosave because it answers:
    what changed between the writer's final handoff and the editor's final version?
    """
    if not revisions or len(revisions) < 2:
        return None, None, None, None, "not_enough_revisions"

    ordered = sorted(revisions, key=_rev_sort_key_global)
    final_matches = [i for i, r in enumerate(ordered) if _name_matches_revision_user(final_editor_name, r)]
    base_matches = [i for i, r in enumerate(ordered) if _name_matches_revision_user(base_editor_name, r)]

    if not final_matches or not base_matches:
        return None, None, None, None, "editor_names_not_found_in_revisions"

    final_idx = final_matches[-1]
    # Prefer the last base/writer revision that happened before or at the first final editor edit.
    first_final_idx = final_matches[0]
    base_before_final = [i for i in base_matches if i < first_final_idx]
    if not base_before_final:
        # Fallback: last base revision before final latest revision.
        base_before_final = [i for i in base_matches if i < final_idx]
    if not base_before_final:
        return None, None, None, None, "no_base_revision_before_final_editor"

    base_idx = base_before_final[-1]
    base_rev = ordered[base_idx]
    final_rev = ordered[final_idx]

    if base_rev.get("id") == final_rev.get("id"):
        return None, None, None, None, "same_base_and_final_revision"

    base_text = export_revision_text(drive_svc, creds, doc_id, base_rev["id"])
    final_text = export_revision_text(drive_svc, creds, doc_id, final_rev["id"])
    return base_text, final_text, base_rev, final_rev, "editor_handoff"




def fetch_latest_editor_previous_revision(drive_svc, creds, doc_id, final_editor_name, revisions):
    """
    Fallback comparison for silent edits.

    Why this is needed:
    Google Docs version history often does not contain the original writer's name.
    Sometimes the article is pasted/uploaded by the editor, or Drive API only exposes
    revision rows for the last modifying user. In that case, strict writer-vs-editor
    matching fails and the app incorrectly returns 100/100.

    Fallback order:
    1) latest revision by selected editor vs the revision immediately before it
    2) latest revision vs previous revision, regardless of names
    3) latest revision vs oldest available revision, when only broad snapshots exist
    """
    if not revisions or len(revisions) < 2:
        return None, None, None, None, "not_enough_revisions_for_fallback"

    ordered = sorted(revisions, key=_rev_sort_key_global)

    # Prefer the latest revision saved by the selected editor.
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

    # If the editor name does not appear in Drive API revisions, compare the last two revisions.
    base_rev = ordered[-2]
    final_rev = ordered[-1]
    base_text = export_revision_text(drive_svc, creds, doc_id, base_rev["id"])
    final_text = export_revision_text(drive_svc, creds, doc_id, final_rev["id"])
    if base_text and final_text and normalize_for_compare(base_text) != normalize_for_compare(final_text):
        return base_text, final_text, base_rev, final_rev, "fallback_latest_vs_previous_revision"

    # Last resort: oldest available revision vs latest available revision.
    base_rev = ordered[0]
    final_rev = ordered[-1]
    if base_rev.get("id") != final_rev.get("id"):
        base_text = export_revision_text(drive_svc, creds, doc_id, base_rev["id"])
        final_text = export_revision_text(drive_svc, creds, doc_id, final_rev["id"])
        return base_text, final_text, base_rev, final_rev, "fallback_oldest_vs_latest_revision"

    return None, None, None, None, "fallback_no_comparable_revisions"

def _revision_user_matches_editor(revision, editor_name):
    """Return True when the Google revision user looks like the selected editor."""
    editor_name_lower = (editor_name or "").strip().lower()
    if not editor_name_lower:
        return True
    user = revision.get("lastModifyingUser", {}) or {}
    display = (user.get("displayName") or "").strip().lower()
    email = (user.get("emailAddress") or "").strip().lower()
    candidates = [display, email]
    return any(c and (editor_name_lower in c or c in editor_name_lower) for c in candidates)


def _dedupe_diff_changes(changes):
    """Remove duplicate autosave / repeated revision changes."""
    unique = []
    seen = set()
    for ch in changes or []:
        key = (
            ch.get("tag", ""),
            normalize_for_compare(ch.get("original", ""))[:250],
            normalize_for_compare(ch.get("revised", ""))[:250],
        )
        if key in seen:
            continue
        seen.add(key)
        unique.append(ch)
    return unique


def _split_arabic_large_changes(changes):
    """
    Arabic paragraphs can be returned as one large replace block.
    This breaks big Arabic rewrite blocks into smaller phrase-level units when possible.
    """
    refined = []
    for ch in changes or []:
        original = ch.get("original", "") or ""
        revised = ch.get("revised", "") or ""
        has_arabic = re.search(r"[\u0600-\u06FF]", original + revised) is not None
        if ch.get("tag") != "replace" or not has_arabic:
            refined.append(ch)
            continue
        # Only split clearly large blocks. Small changes stay as-is.
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
    """
    Count every Google revision-history row returned by Drive API.

    Important: Google Docs UI may show many saves where exporting the document text
    produces the same plain-text result. Those rows are still editorial activity and
    should be visible to the user, even if they cannot be converted into a textual
    before/after diff.
    """
    if not revisions:
        return []
    editor_name_lower = (editor_name or "").strip().lower()
    ordered = sorted(revisions, key=_rev_sort_key_global)
    events = []
    for idx, r in enumerate(ordered, 1):
        display = _revision_display_name(r)
        display_lower = display.lower()
        if editor_name_lower:
            # Match selected editor name, but stay tolerant of partial names.
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
    """
    Fallback/upgrade for Google Docs version-history counting.

    Drive `revisions().list()` often does NOT expose every visible Google Docs
    version-history save. The Drive Activity API can expose edit activity events
    that are closer to the rows users see in the Google Docs Version history UI.

    Notes:
    - This needs the scope: https://www.googleapis.com/auth/drive.activity.readonly
    - The Drive Activity API may not always return a human display name for each
      actor, so when names are not exposed we count all edit activity for the doc
      rather than returning 0.
    - These are visibility/count events only; they should not create deductions
      unless a real text diff is also available.
    """
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

                # Only filter by editor name when Drive Activity actually exposes
                # comparable actor names. Otherwise do not throw the event away.
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

    # Preserve order and avoid exact duplicate activity ids/times.
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
    """
    Parse a manual paste from Google Docs Version history sidebar.

    Why this exists:
    Google Docs UI can show many version-history rows that are NOT exposed by
    Drive revisions().list() or Drive Activity API. When the user pastes the
    visible version-history list, this parser counts every visible save row for
    the selected editor. These rows are event-only activity, not extra text
    deductions.
    """
    if not raw_text or not str(raw_text).strip():
        return []

    import re

    editor_name_lower = (editor_name or "").strip().lower()
    raw_lines = [ln.strip() for ln in str(raw_text).replace("\u202f", " ").replace("\xa0", " ").splitlines()]
    lines = [ln for ln in raw_lines if ln]

    # Examples accepted:
    # May 7, 3:40 PM
    # May 7, 3:40PM
    # Thursday
    # Current version
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
            # Skip UI labels between timestamp and author.
            while j < len(lines) and lines[j].strip().lower() in skip_labels:
                j += 1
            if j < len(lines):
                user = lines[j].strip()
                ul = user.lower()
                # Avoid accidentally treating another timestamp as user.
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
    """
    Create zero-penalty revision-save events from a manual visible count.
    This is used when Google Docs UI shows many revision rows but Google APIs expose fewer.
    """
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
    """
    Robust visible-history counter.
    Primary method: parse timestamp + editor rows.
    Fallback: count occurrences of the selected editor name in the pasted text.
    """
    parsed_events = parse_pasted_google_docs_version_history(raw_text, editor_name)
    parsed_count = len(parsed_events)
    if parsed_count:
        return parsed_count, parsed_events

    # Fallback for pasted text that loses timestamps/layout but keeps names.
    import re
    raw = str(raw_text or "")
    name = (editor_name or "").strip()
    if not raw.strip() or not name:
        return 0, []
    pattern = re.compile(r"(?im)^\s*" + re.escape(name) + r"\s*$")
    count = len(pattern.findall(raw.replace("\u202f", " ").replace("\xa0", " ")))
    return count, make_manual_revision_count_events(count, name, "manual_google_docs_version_history_name_count")

def add_revision_event_visibility(diff_changes, revision_events):
    """
    Keep textual diff changes, but also add zero-penalty event-only rows for revision
    saves that produced no detectable plain-text diff. This makes the report reflect
    all Google Docs version-history activity instead of hiding saves that Google export
    cannot diff.
    """
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
    """
    Strict silent-edit mode:
    - Compare ALL consecutive Google Doc revisions returned by Drive API.
    - Do NOT filter by editor name here, because many docs have writer/editor activity under
      the same visible Google account name, and filtering can hide real silent edits.
    - Do NOT dedupe across revision pairs by default, because the user wants every edit event,
      not only unique final changed blocks.
    - For Arabic, split large paragraph replacements into micro token-level edits.
    """
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

    # Keep repeated edits from different revision pairs. Only remove exact empty/identical artifacts.
    cleaned = []
    for ch in all_changes:
        old_norm = normalize_for_compare(ch.get("original", ""))
        new_norm = normalize_for_compare(ch.get("revised", ""))
        if not old_norm and not new_norm:
            continue
        if old_norm == new_norm:
            continue
        cleaned.append(ch)
    all_changes = cleaned

    # If consecutive export still found nothing, fall back to first vs final.
    if not all_changes:
        first_rev, first_text = exported[0]
        last_rev, last_text = exported[-1]
        fallback = compute_diff(first_text, last_text)
        if should_use_arabic_micro_edits(fallback, lang):
            fallback = _split_arabic_large_changes(fallback)
            fallback = explode_changes_to_micro_edits(fallback, "Arabic")
        return fallback, first_rev, last_rev

    return all_changes, exported[0][0], exported[-1][0]

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
    """
    Extract tracked changes (suggestions) from a Google Docs API response.
    Works when editor used Suggesting mode instead of direct editing.
    Returns list of {type, text, suggestion_id}
    """
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
    """Fetch content, comments, suggestions, and revision history from a private Google Doc shared with the Content QA service account."""
    doc_id = extract_doc_id(url)
    if not doc_id:
        return None, "Invalid Google Doc URL"

    try:
        docs_svc, drive_svc, svc_creds = get_google_services()
    except Exception as e:
        return None, f"Google auth error: {e}"

    # ── Google Doc access gate ─────────────────────────────────
    # This blocks public "anyone with the link" docs and personal Gmail docs
    # before the app reads/scans/scoring the content.
    allowed, validation_error = validate_dubizzle_group_google_doc(drive_svc, doc_id)
    if not allowed:
        return None, validation_error

    try:
        doc   = docs_svc.documents().get(documentId=doc_id).execute()
        title = doc.get("title", "Untitled")
        text, headings = extract_text_from_gdoc(doc)
    except HttpError as e:
        return None, friendly_google_api_error(e, doc_id)
    except Exception as e:
        return None, f"Could not read doc. Share it with {get_service_account_email()} as Editor and try again. Details: {e}"

    # ── Suggestions (tracked changes in Suggesting mode) ───────────────────
    # These are editor edits made in suggestion mode — no extra API call needed.
    suggestions = extract_suggestions_from_doc(doc)

    # ── Comments ───────────────────────────────────────────────────────────
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
            if body:
                comments.append({
                    "author":   author,
                    "email":    email,
                    "text":     body,
                    "resolved": resolved,
                })
            # Also collect replies (editor follow-up comments)
            for r in c.get("replies", []):
                rbody = r.get("content", "").strip()
                rauth = r.get("author", {}).get("displayName", "")
                if rbody and len(rbody) > 10:
                    comments.append({
                        "author":   rauth,
                        "email":    r.get("author", {}).get("emailAddress", ""),
                        "text":     rbody,
                        "resolved": resolved,
                        "is_reply": True,
                    })
    except Exception as e:
        comments_error = str(e)

    # ── Revision history ───────────────────────────────────────────────────
    revisions = []
    try:
        page_token = None
        while True:
            params = dict(
                fileId=doc_id,
                fields="nextPageToken,revisions(id,modifiedTime,lastModifyingUser,exportLinks)",
                pageSize=200,
            )
            if page_token:
                params["pageToken"] = page_token
            resp = drive_svc.revisions().list(**params).execute()
            revisions.extend(resp.get("revisions", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
    except Exception:
        pass

    return {
        "title":           title,
        "text":            text,
        "headings":        headings,
        "links":           re.findall(r'https?://\S+', text),
        "comments":        comments,
        "comments_error":  comments_error,
        "suggestions":     suggestions,
        "revisions":       revisions,
        "revision_count_from_api": len(revisions),
        "word_count":      len(text.split()),
        "drive_svc":       drive_svc,
        "svc_creds":       svc_creds,
        "error":           "",
    }, None

def get_editor_emails():
    """Get editor emails from secrets or return empty list."""
    try:
        raw = st.secrets.get("EDITOR_EMAILS", "")
        return [e.strip().lower() for e in raw.split(",") if e.strip()]
    except Exception:
        return []

def count_revision_rounds(revisions, editor_name):
    """
    Count editor rounds by matching lastModifyingUser.displayName
    against the editor name entered in the form.
    """
    if not revisions:
        return 0, 0, []

    editor_name_lower = editor_name.strip().lower() if editor_name else ""

    annotated = []
    for r in revisions:
        display = r.get("lastModifyingUser", {}).get("displayName", "")
        email   = r.get("lastModifyingUser", {}).get("emailAddress", "")
        # Match if name contains editor name or vice versa
        is_editor = (editor_name_lower and
                     (editor_name_lower in display.lower() or
                      display.lower() in editor_name_lower))
        who = "editor" if is_editor else "writer"
        annotated.append({
            "id":      r.get("id"),
            "time":    r.get("modifiedTime", ""),
            "email":   email or display,  # show display name if no email
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

def classify_comments_ai(comments, platform, lang):
    """Use AI to classify each comment with weighted type."""
    if not comments:
        return []

    c_txt = "\n".join(f"  [{i+1}] {c['text']}" for i, c in enumerate(comments))
    prompt = f"""You are a content QA classifier for {platform} ({lang} content).

Classify each editor comment into exactly one type:
- "factual"     → Wrong data, incorrect facts, wrong names/prices/dates/locations
- "missing"     → Missing critical information that must be added
- "structural"  → Section needs rewriting, wrong structure, copied from source
- "grammar"     → Grammar, phrasing, minor wording fixes

Comments:
{c_txt}

Return ONLY raw JSON, no markdown:
{{"classifications": [{{"index": 1, "type": "factual"}}, {{"index": 2, "type": "grammar"}}]}}

Rules:
- Every comment must be classified
- Use only the 5 types listed above
- When in doubt between factual and structural, pick factual if wrong info is involved"""

    try:
        raw    = call_ai(prompt)
        result = parse_json_response(raw)
        if result and "classifications" in result:
            cls_map = {c["index"]: c["type"] for c in result["classifications"]}
            classified = []
            for i, c in enumerate(comments, 1):
                ctype = cls_map.get(i, "grammar")
                if ctype not in COMMENT_WEIGHTS:
                    ctype = "grammar"
                w = COMMENT_WEIGHTS[ctype]
                classified.append({
                    "author":    c["author"],
                    "email":     c.get("email", ""),
                    "text":      c["text"],
                    "type":      ctype,
                    "label":     w["label"],
                    "deduction": w["deduction"],
                    "color":     w["color"],
                    "tc":        w["tc"],
                })
            return classified
    except Exception:
        pass

    # Fallback: keyword-based
    classified = []
    for c in comments:
        low = c["text"].lower()
        if any(k in low for k in ["wrong","incorrect","not correct","inaccurate","source","from google",
                                   "copied","no apartments","under construction","url goes","link goes",
                                   "mins away","minutes away","data","fact","taken from","from lpv","lpv",
                                   "payment plan","off-plan","off plan","price","aed","sqft","sq ft",
                                   "bedroom","studio","floor","percentage","it is","it's","should be",
                                   "in the source","the source","from the brochure","from the source"]):
            ctype = "factual"
        elif any(k in low for k in ["missing","please add","include","mention","not mentioned",
                                     "should mention","we need","please mention","go through",
                                     "please write","notable","specific","more details","lacks",
                                     "branch","skip this","use another","variation","another link",
                                     "another variation","extensively"]):
            ctype = "missing"
        elif any(k in low for k in ["rewrite","restructure","section should","reorganize",
                                     "wrong section","wrong place","belongs","move this","header"]):
            ctype = "structural"
        elif any(k in low for k in ["brand","tone","style","voice","too general","sounds","generic"]):
            ctype = "rephrase"
        else:
            ctype = "grammar"
        w = COMMENT_WEIGHTS[ctype]
        classified.append({
            "author":    c["author"],
            "email":     c.get("email", ""),
            "text":      c["text"],
            "type":      ctype,
            "label":     w["label"],
            "deduction": w["deduction"],
            "color":     w["color"],
            "tc":        w["tc"],
        })
    return classified

def apply_gdoc_deductions(classified_comments, editor_rounds):
    """Legacy — kept for backward compat."""
    return apply_gdoc_deductions_full(classified_comments, [], editor_rounds)

def capped_low_impact_deduction(items):
    """
    Count every detected silent text edit in the score.

    Earlier versions capped low-impact edits at 8 points, which meant the edit
    count could increase while the final score stayed the same. For this QA
    system, the score must reflect the actual number of detected word/phrase
    edits, so grammar/rephrase/Arabic-language edits are no longer capped.
    """
    events = [d for d in items if d.get("type") in EVENT_ONLY_EDIT_TYPES]
    high = [d for d in items if d.get("type") in HIGH_IMPACT_EDIT_TYPES]
    medium = [d for d in items if d.get("type") not in HIGH_IMPACT_EDIT_TYPES and d.get("type") not in LOW_IMPACT_EDIT_TYPES and d.get("type") not in EVENT_ONLY_EDIT_TYPES]
    low = [d for d in items if d.get("type") in LOW_IMPACT_EDIT_TYPES]

    high_total = sum(float(d.get("deduction", 0)) for d in high)
    medium_total = sum(float(d.get("deduction", 0)) for d in medium)
    raw_low_total = sum(float(d.get("deduction", 0)) for d in low)

    # No cap: if the system detects more real edits, the score changes.
    low_total = raw_low_total

    return high_total + medium_total + low_total, {
        "event_count": len(events),
        "high_count": len(high),
        "medium_count": len(medium),
        "low_count": len(low),
        "raw_low_deduction": round(raw_low_total, 1),
        "low_cap_applied": False,
        "low_capped_deduction": round(low_total, 1),
        "low_uncapped_deduction": round(low_total, 1),
    }

def apply_gdoc_deductions_full(classified_comments, diff_classified, editor_rounds):
    """
    Score = 100 − comment deductions − silent-edit deductions − rounds penalty.
    Every detected silent text edit is counted; low-impact edits are not capped.
    """
    comment_deduction = sum(float(c.get("deduction", 0)) for c in classified_comments)

    # Diff deductions: factual/source changes and low-impact grammar/rephrase edits all count.
    diff_deduction, diff_summary = capped_low_impact_deduction(diff_classified)

    # Rounds penalty
    rounds_penalty = max(0, (editor_rounds - 1)) * REVISION_ROUND_PENALTY

    total_deduction = comment_deduction + diff_deduction + rounds_penalty
    final = max(0, round(100 - total_deduction, 1))

    by_type = {}
    for c in classified_comments:
        by_type.setdefault(c.get("type", "grammar"), []).append(c)

    diff_by_type = {}
    for d in diff_classified:
        diff_by_type.setdefault(d.get("type", "grammar"), []).append(d)

    return final, {
        "base_score":         100,
        "comment_count":      len(classified_comments),
        "comment_deduction":  round(comment_deduction, 1),
        "diff_count":         len(diff_classified),
        "diff_deduction":     round(diff_deduction, 1),
        "diff_summary":       diff_summary,
        "by_type":            by_type,
        "diff_by_type":       diff_by_type,
        "editor_rounds":      editor_rounds,
        "rounds_penalty":     rounds_penalty,
        "final_score":        final,
    }

# ── AI feedback (text only) ────────────────────────────────────────────────
def run_qa_feedback(title, content, writer, ctype, lang, platform, headings, links, comments):
    if not comments:
        scores = {cat: {"score": mx, "feedback": "No editor comments. Full marks awarded.", "comment_refs": []}
                  for cat, mx in CAT_MAX.items()}
        return {"scores": scores, "total": sum(CAT_MAX.values()),
                "overall_feedback": "No editor comments found. All categories awarded full marks.",
                "key_strengths": [], "areas_for_improvement": [], "suggestions": []}

    c_txt = "\n".join(f"  Comment {i+1} [{c['author']}]: {c['text']}" for i, c in enumerate(comments))
    prompt = f"""You are a content QA evaluator for {platform} ({lang}).
Article: "{title}" by {writer} ({ctype})

Editor comments:
{c_txt}

Article excerpt:
{content[:2000]}

Return ONLY a raw JSON object — no markdown, no explanation.

{{
  "scores": {{
    "Content Quality":    {{"score": <0-25>, "feedback": "<brief>", "comment_refs": [<nums>]}},
    "SEO & Structure":    {{"score": <0-20>, "feedback": "<brief>", "comment_refs": [<nums>]}},
    "Language & Grammar": {{"score": <0-20>, "feedback": "<brief>", "comment_refs": [<nums>]}},
    "Brand Voice":        {{"score": <0-15>, "feedback": "<brief>", "comment_refs": [<nums>]}},
    "Readability & Flow": {{"score": <0-10>, "feedback": "<brief>", "comment_refs": [<nums>]}},
    "Originality":        {{"score": <0-10>, "feedback": "<brief>", "comment_refs": [<nums>]}}
  }},
  "total": <sum>,
  "overall_feedback": "<2-3 sentence summary>",
  "key_strengths": ["<strength>"],
  "areas_for_improvement": ["<area>"],
  "suggestions": [
    {{"number": 1, "action": "<fix>", "category": "<category>"}},
    {{"number": 2, "action": "<fix>", "category": "<category>"}}
  ]
}}

Rules: Score based ONLY on the editor comments. Categories not mentioned keep their maximum score."""

    try:
        raw    = call_ai(prompt)
        result = parse_json_response(raw)
        if result and "scores" in result:
            for cat, mx in CAT_MAX.items():
                if cat not in result["scores"]:
                    result["scores"][cat] = {"score": mx, "feedback": "No issues flagged.", "comment_refs": []}
            return result
    except Exception:
        pass

    scores = {cat: {"score": mx, "feedback": "Manual review required — AI unavailable.", "comment_refs": []}
              for cat, mx in CAT_MAX.items()}
    return {"scores": scores, "total": sum(CAT_MAX.values()),
            "overall_feedback": f"AI feedback unavailable. {len(comments)} editor comment(s) found.",
            "key_strengths": [], "areas_for_improvement": [c["text"][:80] for c in comments[:3]], "suggestions": []}

# ── File parsers (existing) ────────────────────────────────────────────────
def extract_docx(raw):
    if not DOCX_OK:
        return {"text": "", "headings": [], "links": [], "comments": [], "word_count": 0, "error": "python-docx not installed"}
    import zipfile
    from lxml import etree as _etree
    doc = Document(BytesIO(raw))
    text, headings, links = [], [], []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t: continue
        text.append(t)
        s = p.style.name
        if   s.startswith("Heading 1"): headings.append({"level": "H1", "text": t})
        elif s.startswith("Heading 2"): headings.append({"level": "H2", "text": t})
        elif s.startswith("Heading 3"): headings.append({"level": "H3", "text": t})
    for rel in doc.part.rels.values():
        if "hyperlink" in rel.reltype: links.append(rel._target)
    comments = []
    try:
        WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        with zipfile.ZipFile(BytesIO(raw)) as z:
            if "word/comments.xml" in z.namelist():
                root  = _etree.fromstring(z.read("word/comments.xml"))
                all_c = root.findall(f".//{{{WNS}}}comment")
                reply_ids = set()
                if "word/commentsExtended.xml" in z.namelist():
                    W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
                    er  = _etree.fromstring(z.read("word/commentsExtended.xml"))
                    for ext in er.findall(f".//{{{W15}}}commentEx"):
                        if ext.get(f"{{{W15}}}paraIdParent"):
                            cid = ext.get(f"{{{W15}}}id", "")
                            if cid: reply_ids.add(cid)
                for c in all_c:
                    cid    = c.get(f"{{{WNS}}}id", "")
                    author = c.get(f"{{{WNS}}}author", "Editor")
                    body   = " ".join(c.itertext()).strip()
                    if body and cid not in reply_ids:
                        comments.append({"author": author, "text": body})
    except Exception:
        pass
    full = "\n".join(text)
    return {"text": full, "headings": headings, "links": links, "comments": comments,
            "word_count": len(full.split()), "error": ""}

def extract_pdf(raw):
    if not PDF_OK:
        return {"text": "", "headings": [], "links": [], "comments": [], "word_count": 0, "error": "pdfplumber not installed"}
    parts, links = [], []
    with pdfplumber.open(BytesIO(raw)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t: parts.append(t)
            for a in (page.annots or []):
                u = a.get("uri")
                if u: links.append(u)
    full = "\n".join(parts)
    return {"text": full, "headings": [], "links": links, "comments": [], "word_count": len(full.split()), "error": ""}

def extract_txt(raw):
    full = raw.decode("utf-8", errors="ignore")
    return {"text": full, "headings": [], "links": re.findall(r'https?://\S+', full),
            "comments": [], "word_count": len(full.split()), "error": ""}

def parse_file(f):
    raw  = f.getvalue(); name = f.name.lower()
    if   name.endswith(".docx"): return extract_docx(raw)
    elif name.endswith(".pdf"):  return extract_pdf(raw)
    else:                        return extract_txt(raw)

# ── Deterministic scoring (file mode) ─────────────────────────────────────
def classify_comment(text):
    low = text.lower()
    for kw in ["wrong","incorrect","not correct","inaccurate","error","should be","it is","it's",
               "the source","in the source","copied","from google","from maps","url goes","link goes",
               "apartments","no apartments","mins away","minutes away","under construction","off-plan","data","fact"]:
        if kw in low: return "Data accuracy", 1.2
    for kw in ["missing","add","please add","include","mention","not mentioned","should mention",
               "we need","please mention","go through","available","please write","notable projects",
               "specific","more details","lacks","header","section"]:
        if kw in low: return "Missing info", 1.2
    return "Grammar / rephrasing", 0.8

def apply_deductions(comments):
    classified        = []
    comment_deduction = 0.0
    for c in comments:
        ctype, pts = classify_comment(c["text"])
        classified.append({"author": c["author"], "text": c["text"], "type": ctype, "deduction": pts})
        comment_deduction += pts
    final = max(0, round(100 - comment_deduction, 1))
    return final, {
        "base_score":        100,
        "comment_count":     len(comments),
        "comment_deduction": round(comment_deduction, 1),
        "classified":        classified,
        "final_score":       final,
    }

def get_recommendation(score):
    return "approve" if score >= 80 else "reject" if score < 60 else "revise"

def get_grade(score):
    for t, label in GRADE_MAP:
        if score >= t: return label
    return GRADE_MAP[-1][1]

def sidebar():
    with st.sidebar:
        st.markdown('<div class="sb-brand"><div class="sb-brand-icon">✦</div><div><div class="sb-brand-title">Content QA</div><div class="sb-brand-sub">Editorial review</div></div></div>', unsafe_allow_html=True)
        st.markdown('<div class="sb-section">Navigation</div>', unsafe_allow_html=True)
        page = st.radio("Navigation",
                        ["📝  New evaluation",
                         "◫  Dashboard"],
                        label_visibility="collapsed", key="sidebar_navigation")
        st.markdown('<div class="sb-section">Deduction rules</div>', unsafe_allow_html=True)
        st.markdown("""
| Rule | Pts |
|---|---:|
| Factual/source correction | −3 |
| Wrong info removed | −2 |
| Missing info added | −1.2 to −1.5 |
| Structural rewrite | −1.2 |
| Arabic/grammar fix | −0.5 to −0.6 |
| Rephrase only | −0.3 |
| Extra revision round | −0.7 |
""")
        st.markdown(
            "<style>section[data-testid='stSidebar'] table{width:100%;font-size:12px;border-collapse:collapse}"
            "section[data-testid='stSidebar'] td,section[data-testid='stSidebar'] th{padding:7px 10px;border-bottom:1px solid #f0f0f0}"
            "section[data-testid='stSidebar'] td:last-child{color:#ef4444;font-weight:700;text-align:right}"
            "section[data-testid='stSidebar'] thead{display:none}</style>",
            unsafe_allow_html=True)

        if "Dashboard" in page: return "dashboard"
        return "gdoc"

# ── Submit page (file upload — existing) ──────────────────────────────────
def page_submit():
    inject_css()
    st.markdown('<div class="qa-hero"><div><div class="qa-hero-badge">✦ File upload mode</div><h1>Submit Article</h1><p>Upload a .docx with editor comments for automated scoring.</p></div><div class="qa-hero-icon">☑</div></div>', unsafe_allow_html=True)
    service_email = get_service_account_email()
    main_col, side_col = st.columns([3.1, 1.05], gap="large")

    with main_col:
        with st.container(border=True):
            st.markdown('<div class="form-card-header"><div><div class="form-card-title">New submission</div><div class="form-card-sub">Fill in the details and upload the article file.</div></div><div class="ready-badge"><span class="ready-dot"></span> Ready to submit</div></div>', unsafe_allow_html=True)
            with st.form("qa_form"):
                c1, c2 = st.columns(2)
                writer      = c1.text_input("Writer name",        placeholder="e.g. Sarah Ahmed")
                editor_name = c2.text_input("Subeditor / editor", placeholder="e.g. Mohamed Ali")
                c3, c4 = st.columns(2)
                title = c3.text_input("Article title", placeholder="e.g. Everything About Mortgages")
                ctype = c4.selectbox("Content type", CONTENT_TYPES)
                c5, _ = st.columns(2)
                lang  = c5.selectbox("Language", LANGUAGES)
                st.markdown('<span style="font-size:12px;font-weight:800;color:#374151;margin-right:8px">Platform</span>', unsafe_allow_html=True)
                platform = st.radio("Platform", PLATFORMS, horizontal=True,
                                    label_visibility="collapsed", key="platform_choice")
                st.markdown('<div class="form-section-divider"></div>', unsafe_allow_html=True)
                st.markdown('<div style="font-size:12px;font-weight:800;color:#374151;margin-bottom:6px">Upload article file</div>', unsafe_allow_html=True)
                upload = st.file_uploader("Upload", type=["docx", "pdf", "txt"],
                                          label_visibility="collapsed")
                if upload:
                    size_mb = upload.size / (1024 * 1024)
                    ext     = upload.name.split(".")[-1].upper()
                    st.markdown(f'<div class="file-card"><div class="file-icon">▤</div><div><div class="file-title">{upload.name}</div><div class="file-meta">{ext} · {size_mb:.1f} MB</div></div><div class="file-status">● Uploaded</div></div>', unsafe_allow_html=True)
                st.markdown(f"""
<div class="precheck">
  <div class="precheck-item {'done' if writer.strip() else ''}"><span class="precheck-dot">✓</span><span>Writer name</span></div>
  <div class="precheck-item {'done' if editor_name.strip() else ''}"><span class="precheck-dot">✓</span><span>Editor name</span></div>
  <div class="precheck-item {'done' if title.strip() else ''}"><span class="precheck-dot">✓</span><span>Article title</span></div>
  <div class="precheck-item {'done' if upload else ''}"><span class="precheck-dot">✓</span><span>File uploaded</span></div>
</div>""", unsafe_allow_html=True)
                go = st.form_submit_button("✦  Run full evaluation", use_container_width=True, type="primary")

    with side_col:
        st.markdown(f"""<div class="side-card">
  <div class="side-card-title">File upload mode</div>
  <div class="timeline-row"><div class="timeline-num">1</div><div><div class="timeline-title">Upload .docx</div><div class="timeline-sub">With editor comments inside.</div></div></div>
  <div class="timeline-row"><div class="timeline-num">2</div><div><div class="timeline-title">Score calculated</div><div class="timeline-sub">Based on comment types.</div></div></div>
  <div class="timeline-row" style="margin-bottom:0"><div class="timeline-num">3</div><div><div class="timeline-title">Get report</div><div class="timeline-sub">Confirm editor decision.</div></div></div>
</div>
<div class="side-card"><div class="tip-box"><div class="tip-title">💡 Try Google Doc mode</div>For richer scoring with comments and silent edit detection, use the Google Doc submission option in the sidebar.</div></div>""", unsafe_allow_html=True)

    if not go: return
    if not writer or not title or not upload:
        st.error("Please fill in writer name, title and upload a file."); return

    with st.spinner("Reading file…"):
        parsed = parse_file(upload)
    if not parsed["text"] or len(parsed["text"]) < 30:
        st.error(f"Could not read text. {parsed.get('error', '')}"); return

    reply_kw = ["fixed","done","added","removed","replaced","updated","changed","edited",
                "deleted","corrected","revised","@","noted","ok ","okay","sure","will do"]
    def is_reply(txt):
        low = txt.lower().strip()
        if low.startswith("@"): return True
        if len(low) < 30:
            for kw in reply_kw:
                if low.startswith(kw) or kw in low[:20]: return True
        return False
    parsed["comments"] = [c for c in parsed["comments"] if not is_reply(c["text"])]

    prog = st.progress(0, text="Starting…")
    prog.progress(20, text="Calculating score…")
    final_score, deductions = apply_deductions(parsed["comments"])
    recommendation          = get_recommendation(final_score)
    prog.progress(50, text="Getting AI feedback…")
    qa = run_qa_feedback(title, parsed["text"], writer, ctype, lang, platform,
                         parsed["headings"], parsed["links"], parsed["comments"])
    prog.progress(100, text="Done."); prog.empty()

    sub = {
        "mode":            "file",
        "date":            datetime.now().strftime("%d %b %Y %H:%M"),
        "platform":        platform,
        "writer":          writer,
        "editor_name":     editor_name,
        "title":           title,
        "content_type":    ctype,
        "language":        lang,
        "word_count":      parsed["word_count"],
        "headings":        parsed["headings"],
        "links":           parsed["links"],
        "comments":        parsed["comments"],
        "qa":              qa,
        "deductions":      deductions,
        "qa_score":        final_score,
        "recommendation":  recommendation,
        "editor_decision": "",
        "editor_notes":    "",
    }
    st.session_state.submissions.append(sub)
    save_record(sub)
    st.success(f"Evaluation complete — Final score: **{final_score} / 100**")
    render_report(sub)

# ── Google Doc submit page ─────────────────────────────────────────────────
def page_gdoc_submit():
    inject_css()
    st.markdown('<div class="qa-hero"><div><div class="qa-hero-badge">✦ Editorial QA Engine</div><h1>Content QA System</h1><p>Submit an article for automated review — editor comments and silent edits are scored automatically.</p></div><div class="qa-hero-icon">☑</div></div>', unsafe_allow_html=True)

    if not GOOGLE_OK:
        st.error("Google API libraries not installed. Add `google-api-python-client` and `google-auth` to requirements.txt")
        return

    service_email = get_service_account_email()
    main_col, side_col = st.columns([3.1, 1.05], gap="large")

    with main_col:
        with st.container(border=True):
            st.markdown('<div class="form-card-header"><div><div class="form-card-title">New submission</div><div class="form-card-sub">Fill in the details and paste the Google Doc link.</div></div><div class="ready-badge"><span class="ready-dot"></span> Ready to submit</div></div>', unsafe_allow_html=True)
            with st.form("gdoc_form"):
                c1, c2 = st.columns(2)
                writer      = c1.text_input("Writer name",        placeholder="e.g. Sarah Ahmed")
                editor_name = c2.text_input("Subeditor / editor", placeholder="e.g. Mohamed Ali")
                c3, c4 = st.columns(2)
                ctype    = c3.selectbox("Content type", CONTENT_TYPES)
                lang     = c4.selectbox("Language", LANGUAGES)
                st.markdown('<span style="font-size:12px;font-weight:800;color:#374151;margin-right:8px">Platform</span>', unsafe_allow_html=True)
                platform = st.radio("Platform", PLATFORMS, horizontal=True,
                                    label_visibility="collapsed", key="gdoc_platform")
                st.markdown('<div class="form-section-divider"></div>', unsafe_allow_html=True)
                st.markdown('<div style="font-size:12px;font-weight:800;color:#374151;margin-bottom:6px">Google Doc link</div>', unsafe_allow_html=True)
                doc_url = st.text_input("Google Doc URL",
                                        placeholder="https://docs.google.com/document/d/...",
                                        label_visibility="collapsed")
                st.markdown(f'<div style="font-size:11px;color:#9ca3af;margin-top:4px">⚠️ Keep the Google Doc restricted and share it with: <strong>{get_service_account_email()}</strong> as Editor</div>', unsafe_allow_html=True)

                # Hidden: version-history paste/count override removed from UI.
                # The system now focuses on the actual handoff comparison:
                # writer's latest available version vs editor's latest available version.
                manual_revision_history = ""
                manual_visible_revision_count = 0

                # Silent edit logic is fixed to editor handoff only:
                # writer's latest available version vs editor's latest available version.
                silent_compare_mode = "Editor handoff: last writer version vs last editor version"

                st.markdown(f"""
<div class="precheck">
  <div class="precheck-item {'done' if writer.strip() else ''}"><span class="precheck-dot">✓</span><span>Writer name</span></div>
  <div class="precheck-item {'done' if editor_name.strip() else ''}"><span class="precheck-dot">✓</span><span>Editor name</span></div>
  <div class="precheck-item {'done' if doc_url.strip() else ''}"><span class="precheck-dot">✓</span><span>Doc link</span></div>
  <div class="precheck-item done"><span class="precheck-dot">✓</span><span>Ready</span></div>
</div>""", unsafe_allow_html=True)
                go = st.form_submit_button("✦  Run full evaluation", use_container_width=True, type="primary")

    with side_col:
        st.markdown(f"""<div class="side-card">
  <div class="side-card-title">How scoring works</div>
  <div class="timeline-row"><div class="timeline-num">1</div><div><div class="timeline-title">Pull doc content</div><div class="timeline-sub">Text, comments and version history.</div></div></div>
  <div class="timeline-row"><div class="timeline-num">2</div><div><div class="timeline-title">AI classifies every issue</div><div class="timeline-sub">Fact/source −3 · Wrong info removed −2 · Missing −1.2/−1.5 · Rephrase −0.3</div></div></div>
  <div class="timeline-row" style="margin-bottom:0"><div class="timeline-num">3</div><div><div class="timeline-title">Latest writer vs editor version scored</div><div class="timeline-sub">Compares and scores actual text differences automatically.</div></div></div>
</div>
<div class="side-card"><div class="tip-box"><div class="tip-title">Before submitting</div>Use a private/restricted Google Doc. Public “Anyone with the link” docs will be rejected. Share the doc with <strong>{service_email}</strong> as Editor for revision export.</div></div>""", unsafe_allow_html=True)

    if not go: return
    if not writer or not doc_url:
        st.error("Please fill in writer name and Google Doc URL."); return

    doc_url = clean_google_doc_url(doc_url)

    with st.spinner("Fetching Google Doc…"):
        parsed, err = fetch_google_doc(doc_url)
    if err:
        st.error(f"Error: {err}"); return
    if not parsed["text"] or len(parsed["text"]) < 30:
        st.error("Could not extract text from the document."); return

    prog = st.progress(0, text="Starting…")

    # In Google Doc mode the API already returns only top-level comments —
    # not nested replies. Apply only a minimal filter for obvious one-word acks.
    CLEAR_REPLIES = {"done","fixed","noted","ok","okay","sure","thanks"}
    def is_obvious_reply(txt):
        t = txt.strip().lower().rstrip(".,!")
        return t in CLEAR_REPLIES or (len(txt.strip()) < 15 and txt.strip().startswith("@"))
    comments = [c for c in parsed["comments"]
                if not is_obvious_reply(c["text"]) and len(c["text"].strip()) > 3]

    raw_total = len(parsed["comments"])
    filtered  = raw_total - len(comments)
    with st.expander(f"📋 Fetched from Google Doc — {raw_total} comments found · {filtered} auto-filtered · {len(comments)} counted"):
        if parsed["comments"]:
            for i, c in enumerate(parsed["comments"], 1):
                kept = not is_obvious_reply(c["text"]) and len(c["text"].strip()) > 3
                color = "#dcfce7" if kept else "#f3f4f6"
                label = "✅ counted" if kept else "⏭ filtered"
                st.markdown(f'<div style="background:{color};border-radius:8px;padding:7px 11px;margin-bottom:5px;font-size:12px"><strong>{c["author"]}</strong> <span style="color:#9ca3af">{label}</span><br>{c["text"][:200]}</div>', unsafe_allow_html=True)
        else:
            st.caption("No comments found. Make sure the doc is shared with the service account and comments are open.")

    prog.progress(20, text="Analyzing editor edits…")
    editor_rounds, total_revs, annotated_revs = count_revision_rounds(parsed["revisions"], editor_name)

    prog.progress(35, text="Finding latest writer and editor versions…")

    diff_changes    = []
    diff_classified = []
    diff_source     = "editor_handoff_last_writer_vs_last_editor"
    revision_activity_events = []
    revision_activity_source = "editor_handoff_only"
    google_api_revision_activity_count = 0
    pasted_revision_count = 0
    writer_rev = None
    editor_rev = None

    # Silent edit logic:
    # 1) Try strict handoff: latest writer version → latest editor version.
    # 2) If the writer name is not exposed in Google revision history, fall back to
    #    latest editor revision → previous revision. This prevents false 100/100 scores.
    handoff_status = "not_run"
    if parsed.get("revisions"):
        prog.progress(50, text="Comparing latest writer version vs latest editor version…")
        handoff_writer_text, handoff_editor_text, handoff_writer_rev, handoff_editor_rev, handoff_status = fetch_editor_handoff_revisions(
            parsed["drive_svc"], parsed["svc_creds"],
            extract_doc_id(doc_url), writer, editor_name, parsed["revisions"]
        )

        # Fallback when the writer name is not present in Drive API revisions.
        if not (handoff_writer_text and handoff_editor_text):
            prog.progress(52, text="Writer/editor handoff not found; comparing latest editor revision with previous revision…")
            handoff_writer_text, handoff_editor_text, handoff_writer_rev, handoff_editor_rev, handoff_status = fetch_latest_editor_previous_revision(
                parsed["drive_svc"], parsed["svc_creds"],
                extract_doc_id(doc_url), editor_name, parsed["revisions"]
            )

        writer_rev, editor_rev = handoff_writer_rev, handoff_editor_rev

        if handoff_writer_text and handoff_editor_text:
            diff_changes = compute_diff(handoff_writer_text, handoff_editor_text)
            if should_use_arabic_micro_edits(diff_changes, lang):
                diff_changes = _split_arabic_large_changes(diff_changes)
                diff_changes = explode_changes_to_micro_edits(diff_changes, "Arabic")
                diff_changes = _dedupe_diff_changes(diff_changes)
            if diff_changes:
                prog.progress(60, text="Classifying and scoring silent edits…")
                diff_classified = classify_diff_changes(
                    diff_changes,
                    platform,
                    effective_diff_language(diff_changes, lang),
                )
            else:
                diff_classified = []
        else:
            diff_classified = []
    else:
        handoff_status = "not_enough_revisions"

    if not diff_classified and handoff_status not in {"editor_handoff", "fallback_latest_editor_vs_previous_revision", "fallback_latest_vs_previous_revision", "fallback_oldest_vs_latest_revision"}:
        st.warning(
            "Could not compare silent edits from Google Docs revision history. "
            "The service account can access the document, but Google did not expose enough exportable revisions. "
            "Try using the exact names from version history, or make one clear saved version before and after editing."
        )
    elif handoff_status != "editor_handoff":
        st.info(
            "Writer/editor handoff names were not fully available in Google revision history, "
            "so the app used fallback comparison: latest editor revision vs the previous available revision."
        )

    prog.progress(65, text="Classifying editor comments with AI…")
    classified = classify_comments_ai(comments, platform, lang)

    prog.progress(72, text="Calculating score…")
    final_score, deductions = apply_gdoc_deductions_full(
        classified, diff_classified, editor_rounds
    )
    recommendation = get_recommendation(final_score)

    prog.progress(88, text="Getting AI feedback…")
    qa = run_qa_feedback(parsed["title"], parsed["text"], writer, ctype, lang, platform,
                         parsed["headings"], parsed["links"], comments)

    prog.progress(100, text="Done."); prog.empty()

    sub = {
        "mode":            "gdoc",
        "doc_url":         doc_url,
        "date":            datetime.now().strftime("%d %b %Y %H:%M"),
        "platform":        platform,
        "writer":          writer,
        "editor_name":     editor_name,
        "title":           parsed["title"],
        "content_type":    ctype,
        "language":        lang,
        "word_count":      parsed["word_count"],
        "headings":        parsed["headings"],
        "links":           parsed["links"],
        "comments":        comments,
        "classified":      classified,
        "diff_classified": diff_classified,
        "diff_source":     diff_source,
        "silent_compare_mode": silent_compare_mode,
        "handoff_status":   locals().get("handoff_status", "not_run"),
        "revision_activity_count": len(revision_activity_events),
        "revision_activity_source": revision_activity_source,
        "revision_activity_events": revision_activity_events,
        "google_api_revision_activity_count": google_api_revision_activity_count,
        "manual_visible_revision_count": int(manual_visible_revision_count or 0),
        "pasted_revision_count": pasted_revision_count,
        "suggestions_raw": parsed.get("suggestions", []),
        "revisions":       annotated_revs,
        "editor_rounds":   editor_rounds,
        "total_revisions": total_revs,
        "writer_rev":      writer_rev,
        "editor_rev":      editor_rev,
        "qa":              qa,
        "deductions":      deductions,
        "qa_score":        final_score,
        "recommendation":  recommendation,
        "editor_decision": "",
        "editor_notes":    "",
    }
    st.session_state.submissions.append(sub)
    save_record(sub)
    st.success(f"Evaluation complete — Final score: **{final_score} / 100**")
    render_gdoc_report(sub)


# ── Human-readable silent edit report helpers ───────────────────────────────
def _short_clean_for_edit_report(value, limit=1200):
    """Return clean text for the editorial edit report."""
    value = sanitize_diff_side(value or "")
    value = re.sub(r"\s+", " ", value).strip()
    value = re.sub(r"^[*\-•]\s*", "", value).strip()
    if len(value) > limit:
        return value[:limit].rstrip() + "…"
    return value


def _edit_report_type(d):
    """Human-friendly edit type matching the user's preferred format."""
    label = (d.get("label") or "").strip()
    if label:
        return label
    typ = d.get("type", "")
    return COMMENT_WEIGHTS.get(typ, COMMENT_WEIGHTS.get("rephrase", {})).get("label", "Edit")


def _edit_report_title(d, idx):
    """Short heading for each edit item."""
    typ = (d.get("type") or "").strip()
    label = _edit_report_type(d)
    if typ in {"factual", "source_alignment", "wrong_info_removed", "missing_info_added", "contradiction_fixed"}:
        return f"{idx}. Factual / source-related edit"
    if typ == "structural":
        return f"{idx}. Structural edit"
    if typ == "arabic_language":
        return f"{idx}. Arabic language edit"
    if typ == "grammar":
        return f"{idx}. Grammar / phrasing edit"
    if typ == "formatting":
        return f"{idx}. Grammar/phrasing edit"
    return f"{idx}. {label}"


def _filter_real_text_edits_for_report(diff_classified):
    """Remove revision events and leaked comment artifacts from displayed edits."""
    rows = []
    for d in diff_classified or []:
        if d.get("type") == "revision_event":
            continue
        original = _short_clean_for_edit_report(d.get("original", ""))
        revised = _short_clean_for_edit_report(d.get("revised", ""))
        if not original and not revised:
            continue
        if _looks_like_comment_artifact(original, revised):
            continue
        nd = dict(d)
        nd["original"] = original
        nd["revised"] = revised
        rows.append(nd)
    return rows

# ── Google Doc Report ──────────────────────────────────────────────────────
def render_gdoc_report(sub):
    inject_css()
    qa    = sub["qa"]; ded = sub["deductions"]
    score = sub["qa_score"]; grade = get_grade(score); rec = sub["recommendation"]
    classified = sub.get("classified", [])

    st.divider()
    bdg_class   = "bdg-bay" if sub["platform"] == "Bayut" else "bdg-dub"
    editor_html = f"&nbsp; 👤 <strong>{sub['editor_name']}</strong>" if sub.get("editor_name") else ""
    mode_chip   = '<span style="font-size:11px;background:#e8f0fe;color:#1a56db;padding:3px 9px;border-radius:20px;font-weight:700">🔗 Google Doc</span>'
    st.markdown(
        f"**{sub['writer']}** &nbsp; <span class='bdg {bdg_class}'>{sub['platform']}</span>"
        f" &nbsp; {mode_chip} &nbsp; `{sub['content_type']}` &nbsp; `{sub['language']}` &nbsp; `{sub['word_count']} words`"
        f"{editor_html} &nbsp; `{sub['date']}`", unsafe_allow_html=True)

    rec_labels = {"approve": ("Approve","#d1fae5","#065f46"),
                  "revise":  ("Request revision","#fef3c7","#92400e"),
                  "reject":  ("Reject","#fee2e2","#991b1b")}
    rl, rbg, rtc = rec_labels.get(rec, rec_labels["revise"])

    def brow(cls, label, val):
        return f'<div class="{cls}"><span>{label}</span><span>{val}</span></div>'

    # Build breakdown by type
    by_type      = ded.get("by_type", {})
    diff_by_type = ded.get("diff_by_type", {})
    comment_rows = ""
    for ctype, items in by_type.items():
        w   = COMMENT_WEIGHTS.get(ctype, COMMENT_WEIGHTS["grammar"])
        tot = sum(i["deduction"] for i in items)
        comment_rows += brow("ded-row", f'Comments — {w["label"]} ({len(items)} × {w["deduction"]} pts)', f'−{round(tot,1)} pts')
    if not by_type:
        comment_rows = brow("ok-row", "No comment deductions", "")

    diff_rows = ""
    diff_summary = ded.get("diff_summary", {})
    for ctype, items in diff_by_type.items():
        w   = COMMENT_WEIGHTS.get(ctype, COMMENT_WEIGHTS["grammar"])
        raw_tot = sum(float(i.get("deduction", 0)) for i in items)
        row_cls = "ded-row" if ctype in HIGH_IMPACT_EDIT_TYPES else "base-row"
        diff_rows += brow(row_cls, f'Silent edits — {w["label"]} ({len(items)} found)', f'−{round(raw_tot,1)} raw')
    if diff_summary.get("low_count", 0) and not diff_summary.get("low_cap_applied"):
        diff_rows += brow("ded-row", f'Low-impact silent edits counted ({diff_summary.get("low_count",0)} grammar/rephrase edits)', f'counted −{diff_summary.get("low_uncapped_deduction", diff_summary.get("raw_low_deduction",0))} pts')
    elif diff_summary.get("low_cap_applied"):
        diff_rows += brow("ok-row", f'Low-impact silent edits capped ({diff_summary.get("low_count",0)} grammar/rephrase edits)', f'counted −{diff_summary.get("low_capped_deduction",0)} pts')
    if not diff_rows and ded.get("diff_count", 0) == 0:
        diff_rows = brow("ok-row", "No silent edits detected", "")

    rounds_row    = ""
    rounds_penalty = ded.get("rounds_penalty", 0)
    editor_rounds  = ded.get("editor_rounds", 0)
    if rounds_penalty > 0:
        rounds_row = brow("ded-row", f'Revision rounds ({editor_rounds} rounds, {editor_rounds-1} extra × 1 pt)', f'−{rounds_penalty} pts')
    else:
        rounds_row = brow("ok-row", f'Revision rounds ({editor_rounds} round{"s" if editor_rounds!=1 else ""}) — no extra penalty', "")

    bd = (brow("base-row","Base score","100 / 100") +
          comment_rows + diff_rows + rounds_row +
          brow("total-row","Final score",f"{score} / 100"))

    st.markdown(
        f'<div class="score-hero"><div class="score-num">{score}<span class="score-den"> / 100</span></div>'
        f'<div class="score-grade">{grade}</div>'
        f'<div style="display:inline-block;margin:6px 0 8px;padding:3px 12px;border-radius:20px;background:{rbg};color:{rtc};font-size:11px;font-weight:500">{rl}</div>'
        f'<div class="score-verdict">{qa.get("overall_feedback","")}</div>'
        f'<div class="breakdown-box">{bd}</div></div>', unsafe_allow_html=True)

    # Silent editor edits (diff)
    diff_classified = sub.get("diff_classified", [])
    if diff_classified:
        st.divider()
        diff_summary = ded.get("diff_summary", {})
        event_count = diff_summary.get("event_count", 0)
        text_change_count = len([d for d in diff_classified if d.get("type") != "revision_event"])
        activity_count = sub.get("revision_activity_count", 0)
        activity_source = sub.get("revision_activity_source", "drive_revisions_api")
        manual_sources = {"manual_google_docs_version_history_paste", "manual_google_docs_version_history_name_count", "manual_visible_revision_count_override"}
        count_label = "visible revision saves" if activity_source in manual_sources else "API revision saves"
        is_handoff = sub.get("diff_source") == "editor_handoff_last_writer_vs_last_editor"
        mode_label = "Editor edits to writer's last version" if is_handoff else "Silent editor activity"
        if is_handoff:
            st.markdown(f"#### {mode_label} — {text_change_count} detected edits")
        else:
            st.markdown(f"#### {mode_label} — {activity_count or event_count} {count_label} · {text_change_count} text changes found")
        google_api_count = sub.get("google_api_revision_activity_count", 0)
        pasted_count = sub.get("pasted_revision_count", 0)
        manual_count = sub.get("manual_visible_revision_count", 0)
        # Internal API/revision details are intentionally hidden from the report UI.
        if diff_summary.get("low_cap_applied"):
            st.info(
                f"Low-impact edits were capped: raw low-impact deduction was "
                f"{diff_summary.get('raw_low_deduction')} pts, counted as "
                f"{diff_summary.get('low_capped_deduction')} pts."
            )
        elif diff_summary.get("low_count", 0):
            st.info(
                f"Low-impact edits are not capped: all "
                f"{diff_summary.get('low_count')} grammar/rephrase edits were counted "
                f"for −{diff_summary.get('low_uncapped_deduction', diff_summary.get('raw_low_deduction'))} pts."
            )

        # Human-readable edit report: writer version vs editor version.
        report_edits = _filter_real_text_edits_for_report(diff_classified)
        writer_label = (sub.get("writer") or "Writer").strip() or "Writer"
        editor_label = (sub.get("editor_name") or "Editor").strip() or "Editor"

        with st.expander(f"View all editor edits — {len(report_edits)} detected edits", expanded=False):
            st.markdown("### Total")
            st.markdown(f"**{len(report_edits)} detected text edits**")
            st.caption("These are the actual text differences between the writer's latest version and the editor's latest version. The system may count several small changes inside one paragraph as separate edits.")

            st.markdown("### All edits")
            for idx, d in enumerate(report_edits, 1):
                original = d.get("original", "")
                revised = d.get("revised", "")
                edit_type = _edit_report_type(d)
                severity = (d.get("severity") or "").strip()
                meaning = "Meaning changed" if d.get("meaning_changed") else "No factual meaning change"
                deduction = d.get("deduction", 0)
                reason = (d.get("reason") or "").strip()

                st.markdown(f"#### {_edit_report_title(d, idx)}")
                if original:
                    st.markdown(f"**{writer_label}:** {original}")
                if revised:
                    st.markdown(f"**{editor_label}:** {revised}")
                st.markdown(f"**Edit type:** {edit_type}")
                st.caption(f"Impact: {severity or 'N/A'} · {meaning} · Raw deduction: −{deduction} pts")
                if reason:
                    st.caption(reason)
                st.markdown("")
    elif sub.get("writer_rev") is None and editor_rounds > 0:
        st.divider()
        st.info("Silent edit scoring unavailable — the revision export could not be retrieved. Make sure the doc is shared with the service account as an Editor, then resubmit.")

    # Classified comments
    if classified:
        st.divider()
        with st.expander(f"View editor comments — {len(classified)} found", expanded=False):
            for idx, c in enumerate(classified, 1):
                st.markdown(
                    f'<div class="cmt-card" style="border-left-color:{c["color"]}">'
                    f'<span class="cmt-author">{c["author"]}</span>'
                    f'<span style="font-size:10px;font-weight:500;padding:1px 8px;border-radius:20px;background:{c["color"]};color:{c["tc"]};margin-left:8px">{c["label"]}</span>'
                    f'<br>{c["text"]}<div class="cmt-deduct">−{c["deduction"]} pts deducted</div></div>',
                    unsafe_allow_html=True)

    # Strengths / improvements
    col_s, col_i = st.columns(2)
    with col_s:
        st.markdown("#### Strengths")
        for s in qa.get("key_strengths", []): st.markdown(f'<span class="tag-str">{s}</span>', unsafe_allow_html=True)
        if not qa.get("key_strengths"): st.caption("None identified.")
    with col_i:
        st.markdown("#### Required improvements")
        for imp in qa.get("areas_for_improvement", []): st.markdown(f'<span class="tag-imp">{imp}</span>', unsafe_allow_html=True)

    if qa.get("suggestions"):
        st.divider(); st.markdown("#### Suggestions")
        for sug in qa["suggestions"]:
            st.markdown(f'<div class="suggest-item"><div class="suggest-num">{sug.get("number","")}</div><div><div>{sug.get("action","")}</div><div class="suggest-cat">Addresses: {sug.get("category","")}</div></div></div>', unsafe_allow_html=True)

    # Decision
    st.divider(); st.markdown("#### Editor decision")
    st.caption("The AI recommendation is a guide. You make the final call.")
    rec_idx  = {"approve": 0, "revise": 1, "reject": 2}
    decision = st.radio("Decision", ["Approve", "Request revision", "Reject"],
                        index=rec_idx.get(rec, 1), horizontal=True,
                        key=f"gdec_{sub['title']}_{sub['date']}")
    notes = st.text_area("Notes for writer", height=90,
                         placeholder="Tell the writer exactly what to fix.",
                         key=f"gnotes_{sub['title']}_{sub['date']}")
    if st.button("Confirm decision", type="primary", use_container_width=True,
                 key=f"gconf_{sub['title']}_{sub['date']}"):
        if decision in ("Request revision", "Reject") and not notes.strip():
            st.error("Please add notes before confirming.")
        else:
            sub["editor_decision"] = decision; sub["editor_notes"] = notes
            update_record_decision(sub)
            st.success(f"Decision saved: {decision}")
            if notes: st.info(f"Notes for {sub['writer']}: {notes}")

    if sub.get("doc_url"):
        st.markdown(f"[🔗 Open original Google Doc]({sub['doc_url']})")
    st.caption(f"Content QA System — {sub['platform']} — Google Doc mode — {sub['date']}")

# ── File upload report ─────────────────────────────────────────────────────
def render_report(sub):
    inject_css()
    if sub.get("mode") == "gdoc":
        render_gdoc_report(sub); return

    qa    = sub["qa"]; ded = sub["deductions"]
    score = sub["qa_score"]; grade = get_grade(score); rec = sub["recommendation"]
    st.divider()
    bdg_class   = "bdg-bay" if sub["platform"] == "Bayut" else "bdg-dub"
    editor_html = f"&nbsp; 👤 <strong>{sub['editor_name']}</strong>" if sub.get("editor_name") else ""
    st.markdown(
        f"**{sub['writer']}** &nbsp; <span class='bdg {bdg_class}'>{sub['platform']}</span>"
        f" &nbsp; `{sub['content_type']}` &nbsp; `{sub['language']}` &nbsp; `{sub['word_count']} words`"
        f"{editor_html} &nbsp; `{sub['date']}`", unsafe_allow_html=True)

    rec_labels = {"approve": ("Approve","#d1fae5","#065f46"),
                  "revise":  ("Request revision","#fef3c7","#92400e"),
                  "reject":  ("Reject","#fee2e2","#991b1b")}
    rl, rbg, rtc = rec_labels.get(rec, rec_labels["revise"])

    def brow(cls, label, val):
        return f'<div class="{cls}"><span>{label}</span><span>{val}</span></div>'

    classified = ded.get("classified", [])
    data_acc   = [c for c in classified if c["type"] == "Data accuracy"]
    missing    = [c for c in classified if c["type"] == "Missing info"]
    grammar    = [c for c in classified if c["type"] == "Grammar / rephrasing"]
    comment_rows = ""
    if data_acc: comment_rows += brow("ded-row", f'Data accuracy ({len(data_acc)} × 1.5 pts)', f'−{round(len(data_acc)*1.5,1)} pts')
    if missing:  comment_rows += brow("ded-row", f'Missing info ({len(missing)} × 1.5 pts)',    f'−{round(len(missing)*1.5,1)} pts')
    if grammar:  comment_rows += brow("ded-row", f'Grammar / rephrasing ({len(grammar)} × 1 pt)', f'−{len(grammar)} pts')
    if not classified: comment_rows = brow("ok-row", "Editor comments", "no deduction")
    bd = brow("base-row","Base score","100 / 100") + comment_rows + brow("total-row","Final score",f"{score} / 100")

    st.markdown(
        f'<div class="score-hero"><div class="score-num">{score}<span class="score-den"> / 100</span></div>'
        f'<div class="score-grade">{grade}</div>'
        f'<div style="display:inline-block;margin:6px 0 8px;padding:3px 12px;border-radius:20px;background:{rbg};color:{rtc};font-size:11px;font-weight:500">{rl}</div>'
        f'<div class="score-verdict">{qa.get("overall_feedback","")}</div>'
        f'<div class="breakdown-box">{bd}</div></div>', unsafe_allow_html=True)

    col_s, col_i = st.columns(2)
    with col_s:
        st.markdown("#### Strengths")
        for s in qa.get("key_strengths", []): st.markdown(f'<span class="tag-str">{s}</span>', unsafe_allow_html=True)
        if not qa.get("key_strengths"): st.caption("None identified.")
    with col_i:
        st.markdown("#### Required improvements")
        for imp in qa.get("areas_for_improvement", []): st.markdown(f'<span class="tag-imp">{imp}</span>', unsafe_allow_html=True)

    if qa.get("suggestions"):
        st.divider(); st.markdown("#### Suggestions")
        for sug in qa["suggestions"]:
            st.markdown(f'<div class="suggest-item"><div class="suggest-num">{sug.get("number","")}</div><div><div>{sug.get("action","")}</div><div class="suggest-cat">Addresses: {sug.get("category","")}</div></div></div>', unsafe_allow_html=True)

    st.divider(); st.markdown("#### Editor decision")
    st.caption("The AI recommendation is a guide. You make the final call.")
    rec_idx  = {"approve": 0, "revise": 1, "reject": 2}
    decision = st.radio("Decision", ["Approve", "Request revision", "Reject"],
                        index=rec_idx.get(rec, 1), horizontal=True,
                        key=f"dec_{sub['title']}_{sub['date']}")
    notes = st.text_area("Notes for writer", height=90,
                         placeholder="Tell the writer exactly what to fix.",
                         key=f"notes_{sub['title']}_{sub['date']}")
    if st.button("Confirm decision", type="primary", use_container_width=True,
                 key=f"conf_{sub['title']}_{sub['date']}"):
        if decision in ("Request revision", "Reject") and not notes.strip():
            st.error("Please add notes before confirming.")
        else:
            sub["editor_decision"] = decision; sub["editor_notes"] = notes
            update_record_decision(sub)
            st.success(f"Decision saved: {decision}")
            if notes: st.info(f"Notes for {sub['writer']}: {notes}")
    st.caption(f"Content QA System — {sub['platform']} — {sub['date']}")

# ── Dashboard ──────────────────────────────────────────────────────────────
def _score_color(s): return "#059669" if s >= 80 else "#d97706" if s >= 60 else "#dc2626"
def _dec_class(d):   return {"Approve":"dec-approve","Request revision":"dec-revise","Reject":"dec-reject"}.get(d,"dec-pending")

def page_dashboard():
    inject_css()
    st.markdown('<div class="qa-hero"><div><div class="qa-hero-badge">Overview</div><h1>Dashboard</h1><p>All evaluation records — persisted across sessions.</p></div><div class="qa-hero-icon">📊</div></div>', unsafe_allow_html=True)
    all_subs = st.session_state.get("submissions", [])
    if not all_subs:
        st.info("No evaluations yet. Submit an article to get started."); return

    total     = len(all_subs)
    approved  = sum(1 for s in all_subs if s.get("editor_decision") == "Approve")
    revision  = sum(1 for s in all_subs if s.get("editor_decision") == "Request revision")
    rejected  = sum(1 for s in all_subs if s.get("editor_decision") == "Reject")
    pending   = sum(1 for s in all_subs if not s.get("editor_decision"))
    avg_score = round(sum(s.get("qa_score", 0) for s in all_subs) / max(total, 1), 1)

    st.markdown(f'<div class="dash-stats-row"><div class="dash-stat blue"><div class="dash-stat-num">{total}</div><div class="dash-stat-lbl">Total</div></div><div class="dash-stat green"><div class="dash-stat-num">{approved}</div><div class="dash-stat-lbl">Approved</div></div><div class="dash-stat amber"><div class="dash-stat-num">{revision}</div><div class="dash-stat-lbl">Revision</div></div><div class="dash-stat red"><div class="dash-stat-num">{rejected}</div><div class="dash-stat-lbl">Rejected</div></div><div class="dash-stat"><div class="dash-stat-num">{pending}</div><div class="dash-stat-lbl">Pending</div></div><div class="dash-stat blue"><div class="dash-stat-num">{avg_score}</div><div class="dash-stat-lbl">Avg score</div></div></div>', unsafe_allow_html=True)

    all_writers = sorted(set(s["writer"]             for s in all_subs if s.get("writer")))
    all_editors = sorted(set(s.get("editor_name","") for s in all_subs if s.get("editor_name")))
    fc1, fc2, fc3, fc4, fc5, fc6 = st.columns(6)
    wf = fc1.selectbox("Writer",       ["All"] + all_writers,   key="dash_writer")
    pf = fc2.selectbox("Platform",     ["All"] + PLATFORMS,     key="dash_platform")
    ef = fc3.selectbox("Editor",       ["All"] + all_editors,   key="dash_editor")
    tf = fc4.selectbox("Content type", ["All"] + CONTENT_TYPES, key="dash_type")
    lf = fc5.selectbox("Language",     ["All"] + LANGUAGES,     key="dash_lang")
    sf = fc6.selectbox("Decision",     ["All","Pending","Approve","Request revision","Reject"], key="dash_status")

    filtered = all_subs
    if wf != "All": filtered = [s for s in filtered if s.get("writer") == wf]
    if pf != "All": filtered = [s for s in filtered if s.get("platform") == pf]
    if ef != "All": filtered = [s for s in filtered if s.get("editor_name") == ef]
    if tf != "All": filtered = [s for s in filtered if s.get("content_type") == tf]
    if lf != "All": filtered = [s for s in filtered if s.get("language") == lf]
    if sf != "All":
        if sf == "Pending": filtered = [s for s in filtered if not s.get("editor_decision")]
        else:               filtered = [s for s in filtered if s.get("editor_decision") == sf]

    st.markdown(f"**{len(filtered)} submission{'s' if len(filtered) != 1 else ''}**")
    st.markdown("")

    for sub in reversed(filtered):
        score     = sub.get("qa_score", 0); dec = sub.get("editor_decision") or "Pending"
        ded       = sub.get("deductions", {})
        cmt_count = ded.get("comment_count", 0); cmt_ded = ded.get("comment_deduction", 0)
        rounds    = sub.get("editor_rounds", 0)
        overall_fb = sub.get("qa", {}).get("overall_feedback", "")
        mode      = sub.get("mode", "file")

        parts = []
        if cmt_count: parts.append(f"{cmt_count} comment{'s' if cmt_count!=1 else ''} (−{cmt_ded} pts)")
        if rounds > 1: parts.append(f"{rounds} revision rounds (−{(rounds-1)*2} pts)")
        if not parts: parts = ["No deductions"]
        score_brief = " · ".join(parts)
        fb_preview  = (overall_fb[:160] + "…") if len(overall_fb) > 160 else overall_fb
        plat_cls    = "bay" if sub.get("platform") == "Bayut" else "dub"
        lang_cls    = "eng" if sub.get("language") == "English" else "ara"
        grade_short = get_grade(score).split(" — ")[-1]
        editor_chip = f'<span class="meta-chip">👤 {sub["editor_name"]}</span>' if sub.get("editor_name") else ""
        mode_chip   = '<span class="meta-chip gdoc">🔗 Google Doc</span>' if mode == "gdoc" else '<span class="meta-chip">📄 File</span>'

        st.markdown(f"""
<div class="article-card">
  <div class="article-card-left">
    <div class="article-card-title">{sub.get('title','Untitled')}</div>
    <div class="article-card-meta">
      <span class="meta-chip">✍️ {sub.get('writer','—')}</span>
      {editor_chip}
      {mode_chip}
      <span class="meta-chip {plat_cls}">{sub.get('platform','')}</span>
      <span class="meta-chip {lang_cls}">{sub.get('language','')}</span>
      <span class="meta-chip">{sub.get('content_type','')}</span>
      <span class="meta-chip">{sub.get('word_count',0)} words</span>
      <span class="meta-chip">🗓 {sub.get('date','')}</span>
    </div>
    <div class="article-card-summary">
      <strong style="color:#374151">{score_brief}</strong>
      {"<br><span style='color:#9ca3af'>" + fb_preview + "</span>" if fb_preview else ""}
    </div>
  </div>
  <div class="article-card-right">
    <div>
      <div class="score-ring" style="--rc:{_score_color(score)};--rv:{int(score)}%">{int(score)}</div>
      <div class="score-ring-lbl">{grade_short}</div>
    </div>
    <span class="dec-badge {_dec_class(dec)}">{dec}</span>
  </div>
</div>""", unsafe_allow_html=True)

        with st.expander("View full report"):
            render_report(sub)

# ── Main ───────────────────────────────────────────────────────────────────
def main():
    if "submissions" not in st.session_state:
        st.session_state.submissions = load_records()
    page = sidebar()
    if page == "dashboard": page_dashboard()
    else:                   page_gdoc_submit()

if __name__ == "__main__":
    main()
