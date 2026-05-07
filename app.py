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
    GOOGLE_OK = True
except ImportError:
    GOOGLE_OK = False

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
    "factual":            {"label": "Factual correction",        "deduction": 2.0, "color": "#fee2e2", "tc": "#991b1b"},
    "wrong_info_removed": {"label": "Wrong info removed",        "deduction": 2.0, "color": "#fee2e2", "tc": "#991b1b"},
    "source_alignment":   {"label": "Source alignment",          "deduction": 2.0, "color": "#fee2e2", "tc": "#991b1b"},
    "contradiction_fixed": {"label": "Contradiction fixed",      "deduction": 2.0, "color": "#fee2e2", "tc": "#991b1b"},
    "missing":            {"label": "Missing critical info",     "deduction": 2.0, "color": "#fef3c7", "tc": "#92400e"},
    "missing_info_added": {"label": "Missing info added",        "deduction": 1.5, "color": "#fef3c7", "tc": "#92400e"},
    "structural":         {"label": "Structural rewrite",        "deduction": 1.5, "color": "#fde8d8", "tc": "#9a3412"},
    "brand_voice":        {"label": "Brand voice / tone",        "deduction": 1.2, "color": "#ede9fe", "tc": "#5b21b6"},
    "arabic_language":    {"label": "Arabic language correction", "deduction": 0.8, "color": "#e0f2fe", "tc": "#075985"},
    "grammar":            {"label": "Grammar / phrasing",        "deduction": 0.7, "color": "#f0f4ff", "tc": "#2D4A8A"},
    "rephrase":           {"label": "Rephrase only",             "deduction": 0.5, "color": "#f1f5f9", "tc": "#475569"},
    "formatting":         {"label": "Formatting / punctuation",  "deduction": 0.2, "color": "#f8fafc", "tc": "#64748b"},
}

LOW_IMPACT_EDIT_TYPES = {"formatting", "rephrase", "grammar", "arabic_language"}
HIGH_IMPACT_EDIT_TYPES = {"factual", "wrong_info_removed", "source_alignment", "contradiction_fixed", "missing", "missing_info_added"}
REVISION_ROUND_PENALTY = 1.0  # per extra round

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
        ]
    )
    docs  = build("docs",  "v1", credentials=creds)
    drive = build("drive", "v3", credentials=creds)
    return docs, drive, creds

def extract_doc_id(url):
    m = re.search(r'/document/d/([a-zA-Z0-9-_]+)', url)
    return m.group(1) if m else None

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

def compute_diff(writer_text, editor_text):
    """
    Diff two versions of the doc at sentence/paragraph level.
    Returns list of {tag, original, revised, similarity} dicts.
    Arabic-aware: ignores tashkeel, tatweel and punctuation-only changes.
    """
    w_sents = split_sentences_smart(writer_text)
    e_sents = split_sentences_smart(editor_text)

    w_keys = [normalize_for_compare(x) for x in w_sents]
    e_keys = [normalize_for_compare(x) for x in e_sents]

    sm = difflib.SequenceMatcher(None, w_keys, e_keys, autojunk=False)
    changes = []

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        original = " ".join(w_sents[i1:i2]).strip()
        revised = " ".join(e_sents[j1:j2]).strip()
        if not original and not revised:
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

def classify_diff_changes(changes, platform, lang):
    """
    Classify silent edits by editorial reason, not just by size.
    Returns each change with label, deduction, severity and explanation.
    """
    if not changes:
        return []

    allowed = list(COMMENT_WEIGHTS.keys())
    items = "\n".join(
        f"[{i+1}] TAG: {c.get('tag','')} | SIMILARITY: {c.get('similarity','')} | WORD_DELTA: {c.get('word_delta','')}\n"
        f"ORIGINAL: {c.get('original','')[:500]}\nREVISED: {c.get('revised','')[:500]}"
        for i, c in enumerate(changes[:80])
    )

    arabic_rules = """
Arabic-specific rules:
- Treat tashkeel, hamza style, punctuation, spacing, and light صياغة changes as grammar/arabic_language/formatting unless the meaning changed.
- If Arabic wording changes a real entity, location, developer, unit type, price, area, number, handover date, amenity, payment plan, road name, or source-backed detail, classify it as factual/source_alignment.
- If the editor deletes unsupported Arabic information, classify it as wrong_info_removed.
- If the editor adds a required source-backed detail, classify it as missing_info_added.
""" if lang == "Arabic" else ""

    prompt = f"""You are a senior editorial QA analyst for {platform}. Content language: {lang}.

An editor silently changed a writer's article without comments. Classify each change by WHY the editor likely made it.

Use ONLY these type values:
- "formatting" → punctuation, spacing, capitalization, tashkeel-only, no meaning change
- "grammar" → grammar/spelling/minor phrasing, no factual meaning change
- "arabic_language" → Arabic grammar, إملاء, صياغة, علامات ترقيم, no factual meaning change
- "rephrase" → same facts, same meaning, smoother sentence
- "brand_voice" → tone/platform voice improved, less generic, better marketing wording
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
- For insert-only changes, decide whether it is missing_info_added, brand_voice, or grammar.

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

    fact_keywords = [
        "aed", "price", "handover", "developer", "location", "bedroom", "studio",
        "sqft", "sq ft", "payment", "floor", "floors", "amenity", "amenities",
        "dubai", "abu dhabi", "dld", "rera", "unit", "units", "villa", "apartment",
        "درهم", "سعر", "أسعار", "المطور", "الموقع", "غرفة", "غرف", "استوديو",
        "قدم", "مربع", "خطة", "الدفع", "طابق", "مرافق", "دبي", "أبوظبي", "وحدة", "شقة", "فيلا"
    ]
    source_keywords = ["source", "brochure", "developer", "official", "dld", "المصدر", "الكتيب", "المطور", "رسمي"]

    if tag == "delete" and len(original.split()) >= 6:
        if any(k in both for k in fact_keywords + source_keywords):
            return "wrong_info_removed"
        return "structural" if len(original.split()) > 25 else "rephrase"
    if tag == "insert" and len(revised.split()) >= 6:
        if any(k in both for k in fact_keywords + source_keywords):
            return "missing_info_added"
        return "brand_voice" if len(revised.split()) > 18 else "rephrase"
    if any(k in both for k in source_keywords) and sim < 0.92:
        return "source_alignment"
    if any(k in both for k in fact_keywords) and sim < 0.88:
        return "factual"
    if ar or lang == "Arabic":
        return "arabic_language" if sim >= 0.82 else "rephrase"
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
    """Fetch content, comments, suggestions, and revision history from a Google Doc."""
    doc_id = extract_doc_id(url)
    if not doc_id:
        return None, "Invalid Google Doc URL"

    try:
        docs_svc, drive_svc, svc_creds = get_google_services()
    except Exception as e:
        return None, f"Google auth error: {e}"

    try:
        doc   = docs_svc.documents().get(documentId=doc_id).execute()
        title = doc.get("title", "Untitled")
        text, headings = extract_text_from_gdoc(doc)
    except Exception as e:
        return None, f"Could not read doc (is it shared with the service account?): {e}"

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
        resp = drive_svc.revisions().list(
            fileId=doc_id,
            fields="revisions(id,modifiedTime,lastModifyingUser,exportLinks)"
        ).execute()
        revisions = resp.get("revisions", [])
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
- "brand_voice" → Tone/style issues, too generic, not platform voice
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
            ctype = "brand_voice"
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
    Low-impact silent edits should not destroy the score.
    Factual/source edits are counted fully; grammar/rephrase/formatting has a soft cap.
    """
    high = [d for d in items if d.get("type") in HIGH_IMPACT_EDIT_TYPES]
    medium = [d for d in items if d.get("type") not in HIGH_IMPACT_EDIT_TYPES and d.get("type") not in LOW_IMPACT_EDIT_TYPES]
    low = [d for d in items if d.get("type") in LOW_IMPACT_EDIT_TYPES]

    high_total = sum(float(d.get("deduction", 0)) for d in high)
    medium_total = sum(float(d.get("deduction", 0)) for d in medium)
    raw_low_total = sum(float(d.get("deduction", 0)) for d in low)

    # Low-impact edits matter, but cap them so a polished article is not punished
    # like one with wrong facts. The first 10 count normally, then cap at 8 pts.
    low_total = min(raw_low_total, 8.0)
    return high_total + medium_total + low_total, {
        "high_count": len(high),
        "medium_count": len(medium),
        "low_count": len(low),
        "raw_low_deduction": round(raw_low_total, 1),
        "low_cap_applied": raw_low_total > low_total,
        "low_capped_deduction": round(low_total, 1),
    }

def apply_gdoc_deductions_full(classified_comments, diff_classified, editor_rounds):
    """
    Score = 100 − comment deductions − smart silent-edit deductions − rounds penalty.
    Silent edits use richer categories and a low-impact cap.
    """
    comment_deduction = sum(float(c.get("deduction", 0)) for c in classified_comments)

    # Diff deductions: factual/source changes count fully; grammar/rephrase is capped.
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
        if kw in low: return "Data accuracy", 1.5
    for kw in ["missing","add","please add","include","mention","not mentioned","should mention",
               "we need","please mention","go through","available","please write","notable projects",
               "specific","more details","lacks","header","section"]:
        if kw in low: return "Missing info", 1.5
    return "Grammar / rephrasing", 1.0

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
| Factual/source correction | −2 |
| Wrong info removed | −2 |
| Missing info added | −1.5 to −2 |
| Structural rewrite | −1.5 |
| Brand voice / tone | −1.2 |
| Arabic/grammar fix | −0.7 to −0.8 |
| Rephrase only | −0.5 |
| Formatting only | −0.2 |
| Extra revision round | −1 |
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
        st.markdown("""<div class="side-card">
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
                st.markdown('<div style="font-size:11px;color:#9ca3af;margin-top:4px">⚠️ Share the doc with: <strong>content-qa-bot@bayut-competitor-gap-analysis.iam.gserviceaccount.com</strong></div>', unsafe_allow_html=True)

                st.markdown(f"""
<div class="precheck">
  <div class="precheck-item {'done' if writer.strip() else ''}"><span class="precheck-dot">✓</span><span>Writer name</span></div>
  <div class="precheck-item {'done' if editor_name.strip() else ''}"><span class="precheck-dot">✓</span><span>Editor name</span></div>
  <div class="precheck-item {'done' if doc_url.strip() else ''}"><span class="precheck-dot">✓</span><span>Doc link</span></div>
  <div class="precheck-item done"><span class="precheck-dot">✓</span><span>Ready</span></div>
</div>""", unsafe_allow_html=True)
                go = st.form_submit_button("✦  Run full evaluation", use_container_width=True, type="primary")

    with side_col:
        st.markdown("""<div class="side-card">
  <div class="side-card-title">How scoring works</div>
  <div class="timeline-row"><div class="timeline-num">1</div><div><div class="timeline-title">Pull doc content</div><div class="timeline-sub">Text, comments and available edit signals.</div></div></div>
  <div class="timeline-row"><div class="timeline-num">2</div><div><div class="timeline-title">AI classifies every issue</div><div class="timeline-sub">Fact/source −2 · Missing −1.5/−2 · Rephrase −0.5</div></div></div>
  <div class="timeline-row" style="margin-bottom:0"><div class="timeline-num">3</div><div><div class="timeline-title">Silent edits scored too</div><div class="timeline-sub">Classifies grammar vs factual/source edits automatically.</div></div></div>
</div>
<div class="side-card"><div class="tip-box"><div class="tip-title">Before submitting</div>Share the Google Doc with the service account. Editor access is better for revision export; Viewer can still read content/comments.</div></div>""", unsafe_allow_html=True)

    if not go: return
    if not writer or not doc_url:
        st.error("Please fill in writer name and Google Doc URL."); return

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

    prog.progress(35, text="Exporting writer & editor versions…")
    writer_text, editor_text, writer_rev, editor_rev = fetch_writer_and_editor_revisions(
        parsed["drive_svc"], parsed["svc_creds"],
        extract_doc_id(doc_url), editor_name, parsed["revisions"]
    )

    diff_changes    = []
    diff_classified = []
    diff_source     = None

    if writer_text and editor_text:
        # Primary: revision export diff
        prog.progress(50, text="Computing diff between writer and editor versions…")
        diff_changes = compute_diff(writer_text, editor_text)
        prog.progress(60, text="Classifying editor edits with AI…")
        diff_classified = classify_diff_changes(diff_changes, platform, lang)
        diff_source = "revision_diff"

    elif parsed.get("suggestions"):
        # Fallback: use tracked suggestions (editor used Suggesting mode)
        prog.progress(55, text="Reading tracked suggestions…")
        suggestions = parsed["suggestions"]
        # Convert suggestions to diff-like changes for classification
        pseudo_changes = []
        for s in suggestions:
            if s["type"] == "insert":
                pseudo_changes.append({"tag": "insert", "original": "", "revised": s["text"]})
            elif s["type"] == "delete":
                pseudo_changes.append({"tag": "delete", "original": s["text"], "revised": ""})
        if pseudo_changes:
            prog.progress(60, text="Classifying tracked suggestions with AI…")
            diff_classified = classify_diff_changes(pseudo_changes, platform, lang)
            diff_changes    = pseudo_changes
            diff_source     = "suggestions"

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
    if diff_summary.get("low_cap_applied"):
        diff_rows += brow("ok-row", f'Low-impact silent edits capped ({diff_summary.get("low_count",0)} grammar/rephrase/formatting edits)', f'counted −{diff_summary.get("low_capped_deduction",0)} pts')
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
        st.markdown(f"#### Silent editor edits — {len(diff_classified)} changes found")
        st.caption(
            "The system separates low-impact wording cleanup from factual/source corrections. "
            f"High-impact: {diff_summary.get('high_count', 0)} · "
            f"Medium: {diff_summary.get('medium_count', 0)} · "
            f"Low-impact: {diff_summary.get('low_count', 0)}"
        )
        if diff_summary.get("low_cap_applied"):
            st.info(
                f"Low-impact edits were capped: raw low-impact deduction was "
                f"{diff_summary.get('raw_low_deduction')} pts, counted as "
                f"{diff_summary.get('low_capped_deduction')} pts."
            )
        for idx, d in enumerate(diff_classified, 1):
            tag_label = {"replace": "Rewritten", "delete": "Deleted", "insert": "Added"}.get(d.get("tag"), "Changed")
            original = _safe_html(d.get("original", "")[:300])
            revised = _safe_html(d.get("revised", "")[:300])
            reason = _safe_html(d.get("reason", ""))
            old_fact = _safe_html(d.get("old_fact", ""))
            new_fact = _safe_html(d.get("new_fact", ""))
            severity = _safe_html(d.get("severity", ""))
            meaning = "Meaning changed" if d.get("meaning_changed") else "No factual meaning change"
            fact_line = ""
            if old_fact or new_fact:
                fact_line = f'<br><span style="font-size:10px;color:#64748b"><strong>Fact shift:</strong> {old_fact} → {new_fact}</span>'
            st.markdown(
                f'<div class="cmt-card" style="border-left-color:{d.get("color", "#e5e7eb")}">'
                f'<span style="font-size:11px;font-weight:700;color:#374151">{idx}. {tag_label}</span>'
                f'<span style="font-size:10px;font-weight:500;padding:1px 8px;border-radius:20px;background:{d.get("color", "#f1f5f9")};color:{d.get("tc", "#475569")};margin-left:8px">{_safe_html(d.get("label", "Edit"))}</span>'
                f'<span style="font-size:10px;font-weight:700;color:#6b7280;margin-left:6px">{severity} · {meaning}</span>'
                f'{("<br><span style=\"font-size:11px;color:#dc2626;text-decoration:line-through\">" + original + "</span>") if original else ""}'
                f'{("<br><span style=\"font-size:11px;color:#059669\">" + revised + "</span>") if revised else ""}'
                f'{fact_line}'
                f'{("<br><span style=\"font-size:10px;color:#9ca3af\">" + reason + "</span>") if reason else ""}'
                f'<div class="cmt-deduct">−{d.get("deduction", 0)} pts raw</div></div>',
                unsafe_allow_html=True)
    elif sub.get("writer_rev") is None and editor_rounds > 0:
        st.divider()
        st.info("Silent edit scoring unavailable — the revision export could not be retrieved. Make sure the doc is shared with the service account as an Editor, then resubmit.")

    # Classified comments
    if classified:
        st.divider()
        st.markdown(f"#### Editor comments — {len(classified)} found")
        for idx, c in enumerate(classified, 1):
            st.markdown(
                f'<div class="cmt-card" style="border-left-color:{c["color"]}">'
                f'<span class="cmt-author">{c["author"]}</span>'
                f'<span style="font-size:10px;font-weight:500;padding:1px 8px;border-radius:20px;background:{c["color"]};color:{c["tc"]};margin-left:8px">{c["label"]}</span>'
                f'<br>{c["text"]}<div class="cmt-deduct">−{c["deduction"]} pts deducted</div></div>',
                unsafe_allow_html=True)

    # Category scores
    st.divider()
    st.markdown("#### Category scores")
    for cat, mx in CAT_MAX.items():
        data = qa["scores"].get(cat, {}); s = data.get("score", 0)
        fb   = data.get("feedback", ""); refs = data.get("comment_refs", [])
        ref_html = " ".join(f'<span class="cat-ref">Comment {r}</span>' for r in refs)
        ca, cb = st.columns([4, 1])
        ca.markdown(f"**{cat}**" + (f" &nbsp; {ref_html}" if ref_html else ""), unsafe_allow_html=True)
        ca.progress(s / mx); cb.markdown(f"**{s} / {mx}**"); st.caption(fb); st.markdown("")

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

    st.divider()
    st.markdown("#### Category scores")
    for cat, mx in CAT_MAX.items():
        data = qa["scores"].get(cat, {}); s = data.get("score", 0)
        fb   = data.get("feedback", ""); refs = data.get("comment_refs", [])
        ref_html = " ".join(f'<span class="cat-ref">Comment {r}</span>' for r in refs)
        ca, cb = st.columns([4, 1])
        ca.markdown(f"**{cat}**" + (f" &nbsp; {ref_html}" if ref_html else ""), unsafe_allow_html=True)
        ca.progress(s / mx); cb.markdown(f"**{s} / {mx}**"); st.caption(fb); st.markdown("")

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
