import streamlit as st
import json
import re
import math
import urllib.request
import urllib.error
from datetime import datetime
from io import BytesIO

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
    import gspread
    from google.oauth2.service_account import Credentials
    SHEETS_OK = True
except ImportError:
    SHEETS_OK = False

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

AI_PHRASES = [
    "in conclusion", "it is worth noting", "it is important to note",
    "delve into", "in the realm of", "furthermore", "moreover",
    "needless to say", "leverage", "utilize", "seamlessly",
    "it goes without saying", "in today's", "one such", "robust",
    "cutting-edge", "state-of-the-art", "at the end of the day",
    "effortlessly blends", "embody contemporary elegance",
    "functional vitality", "architectural lines", "expansive glazing",
]

BROCHURE_PHRASES = [
    "wellness-oriented", "highly anticipated", "masterplan",
    "effortlessly blends", "distinguished residential", "dynamic enclave",
    "lush landscaped buffers", "signature communal", "elevated everyday living",
    "embody contemporary elegance", "functional vitality", "architectural lines",
    "expansive glazing", "highly customisable aesthetic",
    "light and dark material finishes", "open-plan configurations",
    "smart-home integrations", "forward-looking environmental",
    "eco-living standards", "dark sky-compliant",
    "energy-efficient building methods", "smart irrigation",
    "pedestrian-friendly trails", "responsible, sustainable and healthy",
    "certainly. here are", "amenities mentioned in the brochure",
    "define the next chapter", "dynamic urban living",
    "tranquillity of expansive greenery", "wellness-oriented enclave",
    "eco-friendly spaces", "self-sustaining", "immersive community experience",
    "active design principles", "modern sanctuary", "lush landscape of parks",
    "culture, leisure and active", "fabric of daily life",
    "peaceful seclusion", "dynamic pulse", "world-class community amenities",
    "premium off-plan homes", "strategically located", "seamless connectivity",
    "metropolitan accessibility", "opulent master suite",
    "contemporary elegance", "smart-home integration",
    "lush landscape", "boasts", "unparalleled", "premium lifestyle",
    "setting the benchmark",
]

KNOWN_DOMAINS = [
    "emaar.com", "nakheel.com", "damac.com", "aldar.com", "meraas.com",
    "sobha.com", "omniyat.com", "ellington.ae", "azizi.ae",
]


# ── CSS ────────────────────────────────────────────────────────────────────
def inject_css():
    st.markdown("""
<style>
.stApp { background: #f5f6fb; }
.block-container { max-width: 1220px !important; padding-top: 1.8rem !important; padding-left: 2rem !important; padding-right: 2rem !important; padding-bottom: 3rem; }

[data-testid="stVerticalBlockBorderWrapper"] { border: 1px solid #e5e7eb !important; border-radius: 20px !important; box-shadow: 0 14px 35px rgba(17,24,39,.06) !important; background: #fff !important; }
[data-testid="stForm"] { border: none !important; padding: 0 !important; background: transparent !important; }

/* SIDEBAR */
[data-testid="stSidebar"] { background: #ffffff !important; border-right: 1px solid #e5e7eb !important; }
section[data-testid="stSidebar"] > div { padding: 0 !important; }
.sb-brand { display:flex;align-items:center;gap:10px;padding:22px 18px 18px 18px;border-bottom:1px solid #e5e7eb; }
.sb-brand-icon { width:34px;height:34px;border-radius:50%;background:linear-gradient(135deg,#5b5ce2,#7c3aed);display:flex;align-items:center;justify-content:center;color:white;font-size:15px;font-weight:800;flex-shrink:0; }
.sb-brand-title { font-size:13px;font-weight:800;color:#111827;line-height:1.2; }
.sb-brand-sub { font-size:11px;color:#9ca3af;margin-top:2px; }
.sb-section { font-size:10px;color:#9ca3af;font-weight:800;text-transform:uppercase;letter-spacing:.08em;margin:17px 18px 8px 18px; }
section[data-testid="stSidebar"] [data-testid="stRadio"] { padding: 0 14px !important; }
section[data-testid="stSidebar"] [role="radiogroup"] { gap: 6px !important; }
section[data-testid="stSidebar"] [role="radiogroup"] label { border-radius:12px !important;padding:10px 12px !important;margin:0 0 5px 0 !important;background:transparent !important;color:#374151 !important;font-size:13px !important;font-weight:600 !important; }
section[data-testid="stSidebar"] [role="radiogroup"] label:hover { background: #f3f4f6 !important; }
section[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) { background:#eeebff !important;color:#4f46e5 !important;font-weight:800 !important;box-shadow:0 6px 14px rgba(79,70,229,.08) !important; }
section[data-testid="stSidebar"] [role="radiogroup"] label > div:first-child { display: none !important; }
.sb-deduction-wrap { padding: 0 14px 0 14px; }
.sb-deduction-card { background:#fff;border:1px solid #dfe4ea;border-radius:16px;overflow:hidden; }
.sb-deduction-row { display:flex;align-items:center;justify-content:space-between;padding:11px 13px;border-bottom:1px solid #eef2f7;font-size:12px;color:#374151; }
.sb-deduction-row:last-child { border-bottom: none; }
.sb-pill { background:#fee2e2;color:#ef4444;font-size:11px;font-weight:800;border-radius:999px;padding:2px 9px;line-height:1.3; }

/* HERO */
.qa-hero { background:linear-gradient(135deg,#4839d8 0%,#7c3aed 55%,#d067da 100%);border-radius:22px;padding:30px 32px;margin-bottom:24px;color:#fff;display:flex;align-items:flex-start;justify-content:space-between;box-shadow:0 18px 35px rgba(79,70,229,.18);min-height:138px;position:relative;overflow:hidden; }
.qa-hero::after { content:"";position:absolute;right:160px;bottom:-60px;width:520px;height:220px;border-radius:50%;border:1px solid rgba(255,255,255,.12);transform:rotate(-12deg); }
.qa-hero-badge { display:inline-block;background:rgba(255,255,255,.16);color:#fff;border-radius:999px;padding:5px 12px;font-size:11px;font-weight:800;margin-bottom:12px; }
.qa-hero h1 { font-size:30px;font-weight:900;margin:0 0 10px 0;color:#fff;line-height:1.15; }
.qa-hero p { font-size:13px;line-height:1.6;color:rgba(255,255,255,.88);margin:0;max-width:520px; }
.qa-hero-icon { width:66px;height:66px;border-radius:18px;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.15);display:flex;align-items:center;justify-content:center;font-size:28px;flex-shrink:0; }

/* FORM */
.form-card-header { display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:20px; }
.form-card-title { font-size:16px;font-weight:900;color:#111827;margin-bottom:3px; }
.form-card-sub { font-size:12px;color:#6b7280; }
.ready-badge { display:inline-flex;align-items:center;gap:6px;background:#dcfce7;color:#15803d;font-size:11px;font-weight:800;padding:5px 11px;border-radius:999px;white-space:nowrap; }
.ready-dot { width:6px;height:6px;border-radius:50%;background:#16a34a;display:inline-block; }
div[class*="stTextInput"] > label, div[class*="stSelectbox"] > label, div[class*="stFileUploader"] > label { font-size:12px !important;font-weight:800 !important;color:#374151 !important;margin-bottom:6px !important; }
[data-testid="stTextInput"] input { border-radius:12px !important;border:1px solid #d9e0ea !important;padding:10px 13px !important;font-size:13px !important;background:#fff !important;box-shadow:none !important; }
[data-testid="stTextInput"] input:focus { border-color:#8b5cf6 !important;box-shadow:0 0 0 3px rgba(139,92,246,.10) !important; }
[data-baseweb="select"] > div { border-radius:12px !important;border:1px solid #d9e0ea !important;min-height:42px !important;box-shadow:none !important; }

/* Platform radio */
[data-testid="stForm"] div[role="radiogroup"] { display:inline-flex !important;flex-direction:row !important;gap:0 !important;background:#f1f5f9 !important;border:1px solid #dbe3ee !important;border-radius:999px !important;padding:3px !important;width:fit-content !important;box-shadow:inset 0 1px 2px rgba(15,23,42,.03) !important; }
[data-testid="stForm"] div[role="radiogroup"] label { margin:0 !important;min-height:32px !important;min-width:76px !important;padding:7px 17px !important;border-radius:999px !important;border:none !important;background:transparent !important;color:#64748b !important;font-size:12px !important;font-weight:700 !important;display:flex !important;align-items:center !important;justify-content:center !important;box-shadow:none !important; }
[data-testid="stForm"] div[role="radiogroup"] label:has(input:checked) { background:#10b981 !important;color:#ffffff !important;box-shadow:0 6px 13px rgba(16,185,129,.25) !important; }
[data-testid="stForm"] div[role="radiogroup"] label > div:first-child { display: none !important; }
[data-testid="stForm"] div[data-testid="stRadio"] { display:inline-block !important;vertical-align:middle !important;margin-bottom:14px !important; }
[data-testid="stForm"] div[data-testid="stRadio"] > label { display: none !important; }

/* Upload */
.upload-head { display:flex;align-items:center;justify-content:space-between;margin:10px 0 8px 0; }
.upload-head-label { font-size:12px;font-weight:800;color:#374151; }
.upload-head-help { width:18px;height:18px;border-radius:50%;border:1px solid #cbd5e1;color:#6b7280;font-size:11px;font-weight:800;display:flex;align-items:center;justify-content:center; }
[data-testid="stFileUploader"] > label { display: none !important; }
[data-testid="stFileUploader"] section, [data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"] { min-height:152px !important;border:1.5px dashed #9fa8ff !important;border-radius:18px !important;background:#f7f8ff !important;position:relative !important;display:flex !important;align-items:center !important;justify-content:center !important;padding:0 !important;overflow:hidden !important; }
[data-testid="stFileUploader"] section:hover, [data-testid="stFileUploaderDropzone"]:hover, [data-testid="stFileUploadDropzone"]:hover { border-color:#7c3aed !important;background:#f5f3ff !important; }
[data-testid="stFileUploader"] section button, [data-testid="stFileUploader"] section svg, [data-testid="stFileUploader"] section small, [data-testid="stFileUploader"] section span, [data-testid="stFileUploader"] section p, [data-testid="stFileUploader"] section [data-testid="stFileUploaderDropzoneInstructions"] { opacity:0 !important;visibility:hidden !important; }
[data-testid="stFileUploader"] section input[type="file"] { position:absolute !important;inset:0 !important;width:100% !important;height:100% !important;opacity:0 !important;cursor:pointer !important;z-index:20 !important; }
[data-testid="stFileUploader"] section::before, [data-testid="stFileUploaderDropzone"]::before, [data-testid="stFileUploadDropzone"]::before { content:"⇧";position:absolute;left:50%;top:30px;transform:translateX(-50%);width:44px;height:44px;border-radius:16px;background:linear-gradient(135deg,#5b5ce2,#7c3aed);color:#ffffff;font-size:24px;font-weight:800;display:flex;align-items:center;justify-content:center;box-shadow:0 10px 20px rgba(91,92,226,.20);z-index:5;visibility:visible !important;opacity:1 !important; }
[data-testid="stFileUploader"] section::after, [data-testid="stFileUploaderDropzone"]::after, [data-testid="stFileUploadDropzone"]::after { content:"Click or drag a file to upload";position:absolute;left:50%;top:88px;transform:translateX(-50%);width:100%;text-align:center;font-size:13px;line-height:1.4;font-weight:800;color:#111827;z-index:5;visibility:visible !important;opacity:1 !important; }
[data-testid="stFileUploader"] section > div::after, [data-testid="stFileUploaderDropzone"] > div::after, [data-testid="stFileUploadDropzone"] > div::after { content:"DOCX, PDF or TXT · up to 200MB per file";position:absolute;left:50%;top:113px;transform:translateX(-50%);width:100%;text-align:center;font-size:11px;color:#6b7280;z-index:5;visibility:visible !important;opacity:1 !important;pointer-events:none !important; }

/* Submit button */
[data-testid="stFormSubmitButton"] button { background:linear-gradient(135deg,#4338ca,#7c3aed) !important;color:white !important;border:none !important;border-radius:12px !important;font-size:14px !important;font-weight:900 !important;padding:12px !important;width:100% !important;box-shadow:0 12px 24px rgba(124,58,237,.18) !important; }

/* Alert */
[data-testid="stAlert"] { border-radius:16px !important;border:1px solid #d8dbe8 !important;background:#f3f4fb !important;padding:2px 4px !important; }
[data-testid="stAlert"] p { color:#4f46a5 !important;font-size:12px !important; }

/* REPORT */
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
.detect-card{border:1px solid #e8eaf0;border-radius:12px;padding:15px 17px;background:#fff}
.detect-title{font-size:13px;font-weight:600;color:#111827;margin-bottom:7px}
.detect-bar{height:5px;background:#f3f4f6;border-radius:3px;margin-bottom:9px}
.detect-bar-f{height:100%;border-radius:3px}
.detect-thresh{font-size:11px;font-weight:500;padding:4px 10px;border-radius:8px;display:inline-block;margin-bottom:9px}
.detect-note{font-size:12px;color:#6b7280;line-height:1.6}
.detect-split{display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-top:9px}
.detect-seg{text-align:center;background:#f9fafb;border-radius:8px;padding:7px}
.detect-seg-n{font-size:14px;font-weight:600}
.detect-seg-l{font-size:10px;color:#9ca3af;margin-top:2px}
.issue-block{background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:9px 11px;margin-top:9px}
.issue-block-title{font-size:10px;font-weight:700;color:#92400e;text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px}
.issue-snippet{background:#fff;border-left:3px solid #f59e0b;padding:5px 9px;margin-bottom:4px;border-radius:0 6px 6px 0;font-size:11px;color:#374151;line-height:1.5;font-style:italic}
.issue-snippet:last-child{margin-bottom:0}
.cmt-card{background:#f0f2f9;border-left:3px solid #4f46e5;padding:9px 13px;margin-bottom:7px;border-radius:0 8px 8px 0;font-size:13px}
.cmt-author{font-weight:600;color:#4f46e5}
.cmt-deduct{font-size:11px;color:#dc2626;font-weight:500;margin-top:3px}
.cat-ref{font-size:10px;font-weight:500;padding:2px 7px;border-radius:20px;background:#ede9fe;color:#4f46e5;margin-left:6px}
.suggest-item{display:flex;gap:11px;align-items:flex-start;padding:9px 0;border-bottom:1px solid #f3f4f6;font-size:13px}
.suggest-item:last-child{border:none;padding-bottom:0}
.suggest-num{width:22px;height:22px;border-radius:50%;background:#ede9fe;color:#4f46e5;font-size:10px;font-weight:600;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:1px}
.suggest-cat{font-size:11px;color:#9ca3af;margin-top:2px}
.tag-str{background:#d1fae5;color:#065f46;padding:3px 11px;border-radius:20px;font-size:12px;font-weight:500;display:inline-block;margin:2px}
.tag-imp{background:#fef3c7;color:#92400e;padding:3px 11px;border-radius:20px;font-size:12px;font-weight:500;display:inline-block;margin:2px}
.bdg{font-size:11px;font-weight:500;padding:3px 10px;border-radius:20px}
.bdg-bay{background:#d1fae5;color:#065f46}
.bdg-dub{background:#fee2e2;color:#b91c1c}
.no-cmt-notice{background:#f0f2f9;border:1px solid #e0e4f0;border-radius:8px;padding:11px 15px;font-size:13px;color:#6b7280;margin-bottom:10px}

/* DASHBOARD CARDS */
.dash-stats-row { display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin-bottom:24px; }
.dash-stat { background:#fff;border:1px solid #e5e7eb;border-radius:16px;padding:16px 14px;text-align:center; }
.dash-stat-num { font-size:28px;font-weight:900;color:#111827;line-height:1; }
.dash-stat-lbl { font-size:11px;font-weight:700;color:#9ca3af;margin-top:4px;text-transform:uppercase;letter-spacing:.05em; }
.dash-stat.green .dash-stat-num { color:#059669; }
.dash-stat.amber .dash-stat-num { color:#d97706; }
.dash-stat.red   .dash-stat-num { color:#dc2626; }
.dash-stat.blue  .dash-stat-num { color:#4f46e5; }

.article-card { background:#fff;border:1px solid #e5e7eb;border-radius:18px;padding:18px 20px;margin-bottom:12px;display:grid;grid-template-columns:1fr auto;gap:16px;align-items:start;transition:box-shadow .15s; }
.article-card:hover { box-shadow:0 8px 24px rgba(17,24,39,.09); }
.article-card-left { display:flex;flex-direction:column;gap:6px; }
.article-card-title { font-size:15px;font-weight:900;color:#111827;line-height:1.25; }
.article-card-meta { display:flex;flex-wrap:wrap;gap:6px;align-items:center; }
.meta-chip { font-size:11px;font-weight:700;padding:3px 9px;border-radius:20px;background:#f1f5f9;color:#475569; }
.meta-chip.bay { background:#d1fae5;color:#065f46; }
.meta-chip.dub { background:#fee2e2;color:#b91c1c; }
.meta-chip.eng { background:#ede9fe;color:#5b21b6; }
.meta-chip.ara { background:#fef3c7;color:#92400e; }
.article-card-summary { font-size:12px;color:#6b7280;line-height:1.6;margin-top:2px; }
.article-card-right { display:flex;flex-direction:column;align-items:flex-end;gap:8px; }
.score-ring-wrap { text-align:center; }
.score-ring { width:64px;height:64px;border-radius:50%;display:flex;align-items:center;justify-content:center;background: radial-gradient(circle closest-side, white 72%, transparent 74%), conic-gradient(var(--rc) var(--rv), #eef2f7 0);font-size:16px;font-weight:900;color:#111827;margin:0 auto 4px; }
.score-ring-lbl { font-size:10px;font-weight:700;color:#9ca3af;text-align:center; }
.dec-badge { font-size:11px;font-weight:800;padding:4px 10px;border-radius:999px;white-space:nowrap; }
.dec-approve { background:#d1fae5;color:#065f46; }
.dec-revise  { background:#fef3c7;color:#92400e; }
.dec-reject  { background:#fee2e2;color:#991b1b; }
.dec-pending { background:#f1f5f9;color:#64748b; }

.filter-bar { background:#fff;border:1px solid #e5e7eb;border-radius:16px;padding:14px 18px;margin-bottom:18px;display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end; }

/* Stepper */
.stepper { display:flex;align-items:center;justify-content:center;gap:14px;width:100%;max-width:760px;margin:0 auto 24px auto;flex-wrap:nowrap; }
.step-item { display:inline-flex;align-items:center;gap:10px;white-space:nowrap;font-size:13px;font-weight:800;color:#64748b; }
.step-num { width:34px;height:34px;border-radius:50%;border:1px solid #dbe3ee;background:#fff;display:inline-flex;align-items:center;justify-content:center;font-size:13px;font-weight:900;color:#111827;box-shadow:0 4px 10px rgba(15,23,42,.04); }
.step-item.active { color:#6d28d9; }
.step-item.active .step-num { background:linear-gradient(135deg,#5b5ce2,#7c3aed);color:#fff;border-color:transparent;box-shadow:0 10px 22px rgba(109,40,217,.18); }
.step-line { flex:0 0 64px;height:1px;border-top:1px dashed #d6dbe6; }

/* Side cards */
.side-card { background:#ffffff;border:1px solid #e5e7eb;border-radius:18px;padding:20px 20px;box-shadow:0 12px 30px rgba(17,24,39,.055);margin-bottom:14px; }
.side-card-title { font-size:14px;font-weight:900;color:#111827;margin-bottom:14px; }
.timeline-row { display:grid;grid-template-columns:30px 1fr;gap:10px;margin-bottom:16px; }
.timeline-num { width:28px;height:28px;border-radius:50%;background:linear-gradient(135deg,#5b5ce2,#7c3aed);color:white;font-weight:900;font-size:12px;display:flex;align-items:center;justify-content:center; }
.timeline-title { color:#111827;font-size:12px;font-weight:900;margin-bottom:2px; }
.timeline-sub { color:#64748b;font-size:11px;line-height:1.35; }
.preview-grid { display:grid;grid-template-columns:repeat(3,1fr);gap:12px; }
.ring { width:62px;height:62px;border-radius:50%;margin:0 auto 8px auto;display:flex;align-items:center;justify-content:center;background: radial-gradient(circle closest-side,white 72%,transparent 74%), conic-gradient(var(--ring-color) var(--ring-value),#eef2f7 0);color:#111827;font-size:13px;font-weight:900; }
.preview-label { text-align:center;color:#475569;font-size:11px;font-weight:700; }
.tip-box { background:#ecfdf5;border:1px solid #d1fae5;color:#065f46;border-radius:14px;padding:14px 14px;font-size:12px;line-height:1.45; }
.tip-title { font-size:13px;font-weight:900;margin-bottom:4px; }

.file-card { display:flex;align-items:center;gap:12px;background:#ffffff;border:1px solid #e5e7eb;border-radius:14px;padding:12px 14px;margin-top:12px; }
.file-icon { width:34px;height:34px;border-radius:10px;background:#eef2ff;color:#4f46e5;display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0; }
.file-title { font-size:13px;color:#111827;font-weight:800;line-height:1.2; }
.file-meta { font-size:11px;color:#64748b;margin-top:2px; }
.file-status { margin-left:auto;background:#dcfce7;color:#15803d;font-size:11px;font-weight:800;border-radius:999px;padding:5px 10px;white-space:nowrap; }

.precheck { display:grid;grid-template-columns:repeat(4,1fr);gap:0;background:#fbfcff;border:1px solid #e5e7eb;border-radius:13px;padding:8px 10px;margin-top:12px; }
.precheck-item { display:flex;align-items:center;gap:7px;color:#64748b;font-size:11px;font-weight:700;white-space:nowrap; }
.precheck-dot { width:15px;height:15px;border-radius:50%;background:#e5e7eb;color:#94a3b8;display:inline-flex;align-items:center;justify-content:center;font-size:10px;font-weight:900; }
.precheck-item.done { color:#334155; }
.precheck-item.done .precheck-dot { background:#22c55e;color:white; }
.form-section-divider { height:1px;background:#e8edf5;margin:16px 0 14px 0; }

div[data-testid="stVerticalBlockBorderWrapper"] > div { border-radius:22px !important; }
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stVerticalBlock"] { gap:0.9rem !important; }

[data-testid="stFileUploader"] section button,[data-testid="stFileUploader"] section svg,[data-testid="stFileUploader"] section small,[data-testid="stFileUploader"] section span,[data-testid="stFileUploader"] section p,[data-testid="stFileUploader"] section [data-testid="stFileUploaderDropzoneInstructions"],[data-testid="stFileUploaderDropzone"] button,[data-testid="stFileUploaderDropzone"] svg,[data-testid="stFileUploaderDropzone"] small,[data-testid="stFileUploaderDropzone"] span,[data-testid="stFileUploaderDropzone"] p,[data-testid="stFileUploadDropzone"] button,[data-testid="stFileUploadDropzone"] svg,[data-testid="stFileUploadDropzone"] small,[data-testid="stFileUploadDropzone"] span,[data-testid="stFileUploadDropzone"] p { opacity:0 !important;visibility:hidden !important; }

[data-testid="stFileUploader"] section::before,[data-testid="stFileUploaderDropzone"]::before,[data-testid="stFileUploadDropzone"]::before,[data-testid="stFileUploader"] section::after,[data-testid="stFileUploaderDropzone"]::after,[data-testid="stFileUploadDropzone"]::after,[data-testid="stFileUploader"] section > div::after,[data-testid="stFileUploaderDropzone"] > div::after,[data-testid="stFileUploadDropzone"] > div::after { opacity:1 !important;visibility:visible !important;pointer-events:none !important; }
</style>
""", unsafe_allow_html=True)


# ── Groq AI ────────────────────────────────────────────────────────────────
def call_ai(prompt):
    if not GROQ_OK:
        raise Exception("groq not installed")
    client = Groq(api_key=st.secrets["GROQ_API_KEY"])
    for model in ["llama-3.1-8b-instant", "llama3-8b-8192", "gemma2-9b-it"]:
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3, max_tokens=2000)
            return resp.choices[0].message.content.strip()
        except Exception as e:
            if any(x in str(e).lower() for x in ["model", "not found", "decommission"]):
                continue
            raise e
    raise Exception("All Groq models failed.")


# ── File parsers ───────────────────────────────────────────────────────────
def extract_docx(raw):
    if not DOCX_OK:
        return {"text": "", "headings": [], "links": [], "comments": [], "word_count": 0, "error": "python-docx not installed"}
    import zipfile
    from lxml import etree as _etree

    doc = Document(BytesIO(raw))
    text, headings, links = [], [], []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        text.append(t)
        s = p.style.name
        if   s.startswith("Heading 1"): headings.append({"level": "H1", "text": t})
        elif s.startswith("Heading 2"): headings.append({"level": "H2", "text": t})
        elif s.startswith("Heading 3"): headings.append({"level": "H3", "text": t})
    for rel in doc.part.rels.values():
        if "hyperlink" in rel.reltype:
            links.append(rel._target)

    comments = []
    try:
        WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        with zipfile.ZipFile(BytesIO(raw)) as z:
            if "word/comments.xml" in z.namelist():
                root = _etree.fromstring(z.read("word/comments.xml"))
                all_c = root.findall(f".//{{{WNS}}}comment")

                reply_comment_ids = set()
                if "word/commentsExtended.xml" in z.namelist():
                    ext_root = _etree.fromstring(z.read("word/commentsExtended.xml"))
                    W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
                    for ext in ext_root.findall(f".//{{{W15}}}commentEx"):
                        has_parent = ext.get(f"{{{W15}}}paraIdParent")
                        if has_parent:
                            cid = ext.get(f"{{{W15}}}id", "")
                            if cid:
                                reply_comment_ids.add(cid)

                for c in all_c:
                    cid    = c.get(f"{{{WNS}}}id", "")
                    author = c.get(f"{{{WNS}}}author", "Editor")
                    body   = " ".join(c.itertext()).strip()
                    if body and cid not in reply_comment_ids:
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
            if t:
                parts.append(t)
            for a in (page.annots or []):
                u = a.get("uri")
                if u:
                    links.append(u)
    full = "\n".join(parts)
    return {"text": full, "headings": [], "links": links, "comments": [], "word_count": len(full.split()), "error": ""}


def extract_txt(raw):
    full = raw.decode("utf-8", errors="ignore")
    return {"text": full, "headings": [], "links": re.findall(r'https?://\S+', full),
            "comments": [], "word_count": len(full.split()), "error": ""}


def parse_file(f):
    raw  = f.getvalue()
    name = f.name.lower()
    if   name.endswith(".docx"): return extract_docx(raw)
    elif name.endswith(".pdf"):  return extract_pdf(raw)
    else:                        return extract_txt(raw)


# ── Scoring ────────────────────────────────────────────────────────────────
def classify_comment(text):
    low = text.lower()
    data_accuracy_kw = ["wrong", "incorrect", "not correct", "inaccurate", "error",
        "should be", "it is", "it's", "the source", "in the source",
        "copied", "from google", "from maps", "url goes", "link goes",
        "apartments", "no apartments", "mins away", "minutes away",
        "under construction", "off-plan", "data", "fact"]
    missing_info_kw  = ["missing", "add", "please add", "include", "mention", "not mentioned",
        "should mention", "we need", "please mention", "go through",
        "available", "please write", "write the branch", "notable projects",
        "specific", "more details", "lacks", "header", "section"]
    grammar_kw       = ["grammar", "rephrase", "rewrite", "word", "sentence", "phrasing",
        "general", "too general", "vague", "unclear", "confusing",
        "brand voice", "tone", "style", "read", "sounds"]
    for kw in data_accuracy_kw:
        if kw in low: return "Data accuracy", 1.5
    for kw in missing_info_kw:
        if kw in low: return "Missing info", 1.5
    for kw in grammar_kw:
        if kw in low: return "Grammar / rephrasing", 1.0
    return "Grammar / rephrasing", 1.0


def apply_deductions(base_score, comments, plag_pct, ai_pct):
    classified       = []
    comment_deduction = 0.0
    for c in comments:
        ctype, pts = classify_comment(c["text"])
        classified.append({"author": c["author"], "text": c["text"], "type": ctype, "deduction": pts})
        comment_deduction += pts

    plag_brackets  = int(plag_pct // 20)
    plag_deduction = plag_brackets * 5
    ai_brackets    = int(ai_pct // 20)
    ai_deduction   = ai_brackets * 5
    final = max(0, round(100 - comment_deduction - plag_deduction - ai_deduction, 1))

    return final, {
        "base_score":        100,
        "comment_count":     len(comments),
        "comment_deduction": round(comment_deduction, 1),
        "classified":        classified,
        "plag_pct":          plag_pct,
        "plag_brackets":     plag_brackets,
        "plag_deduction":    plag_deduction,
        "ai_pct":            ai_pct,
        "ai_brackets":       ai_brackets,
        "ai_deduction":      ai_deduction,
        "final_score":       final,
    }


def get_recommendation(score):
    return "approve" if score >= 80 else "reject" if score < 60 else "revise"


def get_grade(score):
    for t, label in GRADE_MAP:
        if score >= t: return label
    return GRADE_MAP[-1][1]


# ── QA ─────────────────────────────────────────────────────────────────────
def run_qa(title, content, writer, ctype, lang, platform, headings, links, comments):
    h_txt = "\n".join(f"  [{h['level']}] {h['text']}" for h in headings) or "  None"
    l_txt = "\n".join(f"  - {l}" for l in links[:8])                      or "  None"

    if not comments:
        scores = {cat: {"score": mx, "feedback": "No editor comments. Full marks awarded.", "comment_refs": []}
                  for cat, mx in CAT_MAX.items()}
        return {"scores": scores, "total": sum(CAT_MAX.values()),
                "overall_feedback": "No editor comments found. All categories awarded full marks.",
                "key_strengths": [], "areas_for_improvement": [], "suggestions": []}

    c_txt = "\n".join(f"  Comment {i+1} [{c['author']}]: {c['text']}" for i, c in enumerate(comments))

    prompt = f"""You are a senior content QA evaluator for {platform}, a leading UAE real estate platform.
Evaluate this {ctype.lower()} written in {lang}.

TITLE: {title}
WRITER: {writer}
HEADINGS: {h_txt}
LINKS: {l_txt}
EDITOR COMMENTS: {c_txt}
ARTICLE (context only): {content[:3000]}

RULES:
1. Score ONLY based on editor comments. Do NOT evaluate content independently.
2. Map each comment to the most relevant category and reduce that score.
3. Categories with NO related comments get their MAXIMUM score.
4. Every comment must appear in at least one category's feedback.
5. comment_refs must list the comment numbers that affected each category.
6. Suggestions must directly address the comments.

Return ONLY valid JSON:
{{
  "scores": {{
    "Content Quality":    {{"score":<0-25>,"feedback":"<what comments flagged>","comment_refs":[]}},
    "SEO & Structure":    {{"score":<0-20>,"feedback":"<what comments flagged>","comment_refs":[]}},
    "Language & Grammar": {{"score":<0-20>,"feedback":"<what comments flagged>","comment_refs":[]}},
    "Brand Voice":        {{"score":<0-15>,"feedback":"<what comments flagged>","comment_refs":[]}},
    "Readability & Flow": {{"score":<0-10>,"feedback":"<what comments flagged>","comment_refs":[]}},
    "Originality":        {{"score":<0-10>,"feedback":"<what comments flagged>","comment_refs":[]}}
  }},
  "total":<sum>,
  "overall_feedback":"<3 sentence summary referencing specific comments>",
  "key_strengths":[],
  "areas_for_improvement":["<from comment 1>","<from comment 2>"],
  "suggestions":[
    {{"number":1,"action":"<specific fix from a comment>","category":"<category>"}},
    {{"number":2,"action":"<specific fix from a comment>","category":"<category>"}},
    {{"number":3,"action":"<specific fix from a comment>","category":"<category>"}}
  ]
}}"""

    raw   = call_ai(prompt)
    clean = re.sub(r"```json|```", "", raw).strip()
    m     = re.search(r'\{.*\}', clean, re.DOTALL)
    if m:
        clean = m.group(0)
    return json.loads(clean)


# ── Plagiarism ─────────────────────────────────────────────────────────────
def check_plagiarism(text, links):
    flagged_sources = [l for l in links if any(d in l for d in KNOWN_DOMAINS)]
    prompt = f"""You are a plagiarism detection expert for UAE real estate content.
Read the article below and identify sentences that appear to be copied or closely lifted from developer brochures or websites.
Return ONLY valid JSON:
{{
  "plagiarism_percentage": <0-100 integer>,
  "flagged_sentences": ["<exact copied sentence>"],
  "assessment": "<1 sentence summary>"
}}
Rules:
- Only flag sentences with clear brochure/developer marketing language
- Do NOT flag factual data (prices, sq ft, dates, distances)
- Return max 15 flagged sentences

ARTICLE:
{text[:5000]}"""

    try:
        raw    = call_ai(prompt)
        clean  = re.sub(r"```json|```", "", raw).strip()
        m      = re.search(r'{.*}', clean, re.DOTALL)
        if m:
            clean = m.group(0)
        result = json.loads(clean)
        pct    = min(int(result.get("plagiarism_percentage", 0)), 100)
        return {
            "percentage":        pct,
            "flagged_sources":   flagged_sources,
            "flagged_sentences": result.get("flagged_sentences", [])[:15],
            "hits":              len(result.get("flagged_sentences", [])),
            "source":            "Groq",
            "assessment":        result.get("assessment", ""),
            "status":            "danger" if pct > 20 else "warn" if pct > 10 else "safe",
        }
    except Exception:
        text_lower = text.lower()
        hits  = sum(1 for p in BROCHURE_PHRASES if p in text_lower)
        total = len(BROCHURE_PHRASES)
        base  = min(int((math.sqrt(hits) / math.sqrt(max(total, 1))) * 60), 60)
        bonus = min(len(flagged_sources) * 4, 15)
        pct   = min(base + bonus, 100)
        return {"percentage": pct, "flagged_sources": flagged_sources,
                "flagged_sentences": [], "hits": hits, "source": "heuristic",
                "status": "danger" if pct > 20 else "warn" if pct > 10 else "safe"}


def get_plag_snippets(text, links, plag_result=None):
    if plag_result and plag_result.get("flagged_sentences"):
        return plag_result.get("flagged_sources", []), plag_result.get("flagged_sentences", [])
    flagged_sources = [l for l in links if any(d in l for d in KNOWN_DOMAINS)]
    sentences = re.split(r'(?<=[.!?])\s+', text)
    flagged, seen = [], set()
    for sent in sentences:
        stripped = sent.strip()
        low = stripped.lower()
        if len(low) < 35 or stripped in seen:
            continue
        for phrase in BROCHURE_PHRASES:
            if phrase in low:
                seen.add(stripped)
                flagged.append(stripped[:280])
                break
    return flagged_sources, flagged[:8]


def highlight_plag(sentence):
    result = sentence
    for phrase in BROCHURE_PHRASES:
        pat = re.compile(re.escape(phrase), re.IGNORECASE)
        def mark(m): return ('<mark style="background:#fecaca;border-radius:3px;padding:0 2px;font-weight:500;color:#7f1d1d">' + m.group(0) + '</mark>')
        result = pat.sub(mark, result)
    return result


# ── AI detection ───────────────────────────────────────────────────────────
def check_ai(text):
    api_key = st.secrets.get("GPTZERO_API_KEY", "")
    if api_key and len(text.strip()) > 50:
        try:
            payload = json.dumps({"document": text[:10000], "version": "2025-01-09"}).encode("utf-8")
            req = urllib.request.Request(
                "https://api.gptzero.me/v2/predict/text",
                data=payload,
                headers={"Content-Type": "application/json", "Accept": "application/json", "x-api-key": api_key},
                method="POST")
            with urllib.request.urlopen(req, timeout=20) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            doc    = data.get("documents", [{}])[0]
            ai_pct = int(round(doc.get("completely_generated_prob", 0) * 100))
            sents  = doc.get("sentences", [])
            ai_sents = [s.get("sentence", "") for s in sents
                        if s.get("generated_prob", 0) > 0.5 and len(s.get("sentence", "")) > 30][:5]
            return {"ai_pct": ai_pct, "human_pct": 100 - ai_pct,
                    "status": "danger" if ai_pct > 20 else "warn" if ai_pct > 10 else "safe",
                    "source": "GPTZero", "ai_sentences": ai_sents}
        except Exception:
            pass

    hits     = sum(1 for p in AI_PHRASES if p in text.lower())
    pct      = min(hits * 5, 60)
    snippets = []
    for sent in re.split(r'(?<=[.!?])\s+', text):
        low = sent.strip().lower()
        if len(low) < 35:
            continue
        if any(p in low for p in AI_PHRASES):
            snippets.append(sent.strip()[:200])
        if len(snippets) >= 5:
            break
    return {"ai_pct": pct, "human_pct": 100 - pct,
            "status": "danger" if pct > 20 else "warn" if pct > 10 else "safe",
            "source": "heuristic", "ai_sentences": snippets}


def highlight_ai(sentence):
    result = sentence
    for phrase in AI_PHRASES:
        pat = re.compile(re.escape(phrase), re.IGNORECASE)
        def mark(m): return ('<mark style="background:#fef3c7;border-radius:3px;padding:0 2px;font-weight:500;color:#78350f">' + m.group(0) + '</mark>')
        result = pat.sub(mark, result)
    return result


# ── Google Sheets ──────────────────────────────────────────────────────────
SHEET_HEADERS = [
    "Date", "Platform", "Writer", "Editor", "Title", "Type", "Language",
    "Word Count", "Comment Count", "Comment Deduction",
    "Plagiarism Ded", "AI Ded", "Final Score",
    "Plagiarism%", "AI%", "Recommendation",
    "Editor Decision", "Notes", "Overall Feedback"
]


def _get_sheet():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc    = gspread.authorize(creds)
    return gc.open(st.secrets.get("SHEET_NAME", "Bayut QA Submissions")).sheet1


def log_to_sheets(row):
    if not SHEETS_OK:
        return
    try:
        sheet = _get_sheet()
        if not sheet.row_values(1):
            sheet.append_row(SHEET_HEADERS)
        d = row.get("deductions", {})
        sheet.append_row([
            row.get("date", ""),
            row.get("platform", ""),
            row.get("writer", ""),
            row.get("editor_name", ""),
            row.get("title", ""),
            row.get("content_type", ""),
            row.get("language", ""),
            row.get("word_count", 0),
            d.get("comment_count", 0),
            d.get("comment_deduction", 0),
            d.get("plag_deduction", 0),
            d.get("ai_deduction", 0),
            d.get("final_score", 0),
            row.get("plagiarism_pct", 0),
            row.get("ai_pct", 0),
            row.get("recommendation", ""),
            row.get("editor_decision", ""),
            row.get("editor_notes", ""),
            row.get("qa", {}).get("overall_feedback", ""),
        ])
    except Exception as e:
        st.warning(f"Google Sheets log failed: {e}")


def update_sheet_decision(sub):
    """Update the editor decision + notes on the existing row."""
    if not SHEETS_OK:
        return
    try:
        sheet = _get_sheet()
        records = sheet.get_all_records()
        for idx, r in enumerate(records, start=2):   # row 1 = header
            if (r.get("Writer") == sub.get("writer") and
                    r.get("Title") == sub.get("title") and
                    r.get("Date")  == sub.get("date")):
                # Column positions (1-indexed) for Editor Decision and Notes
                dec_col   = SHEET_HEADERS.index("Editor Decision") + 1
                notes_col = SHEET_HEADERS.index("Notes") + 1
                sheet.update_cell(idx, dec_col,   sub.get("editor_decision", ""))
                sheet.update_cell(idx, notes_col, sub.get("editor_notes", ""))
                return
    except Exception as e:
        st.warning(f"Could not update decision in Sheets: {e}")


def load_from_sheets():
    """Load all historical submissions from Google Sheets."""
    if not SHEETS_OK:
        return []
    try:
        sheet   = _get_sheet()
        records = sheet.get_all_records()
        loaded  = []
        for r in records:
            ai_pct   = int(r.get("AI%", 0) or 0)
            plag_pct = int(r.get("Plagiarism%", 0) or 0)
            score    = float(r.get("Final Score", 0) or 0)
            loaded.append({
                "date":          str(r.get("Date", "")),
                "platform":      str(r.get("Platform", "")),
                "writer":        str(r.get("Writer", "")),
                "editor_name":   str(r.get("Editor", "")),
                "title":         str(r.get("Title", "")),
                "content_type":  str(r.get("Type", "")),
                "language":      str(r.get("Language", "")),
                "word_count":    int(r.get("Word Count", 0) or 0),
                "headings":      [],
                "links":         [],
                "comments":      [],
                "qa": {
                    "scores":                {cat: {"score": mx, "feedback": "", "comment_refs": []} for cat, mx in CAT_MAX.items()},
                    "overall_feedback":      str(r.get("Overall Feedback", "")),
                    "key_strengths":         [],
                    "areas_for_improvement": [],
                    "suggestions":           [],
                },
                "plagiarism":     {"percentage": plag_pct, "source": "Sheets"},
                "plag_snippets":  [],
                "plag_sources":   [],
                "ai_detection":   {
                    "ai_pct":       ai_pct,
                    "human_pct":    100 - ai_pct,
                    "source":       "Sheets",
                    "ai_sentences": [],
                },
                "deductions": {
                    "base_score":        100,
                    "comment_count":     int(r.get("Comment Count", 0) or 0),
                    "comment_deduction": float(r.get("Comment Deduction", 0) or 0),
                    "plag_deduction":    float(r.get("Plagiarism Ded", 0) or 0),
                    "ai_deduction":      float(r.get("AI Ded", 0) or 0),
                    "plag_pct":          plag_pct,
                    "ai_pct":            ai_pct,
                    "plag_brackets":     plag_pct // 20,
                    "ai_brackets":       ai_pct // 20,
                    "final_score":       score,
                    "classified":        [],
                },
                "qa_score":        score,
                "plagiarism_pct":  plag_pct,
                "ai_pct":          ai_pct,
                "recommendation":  str(r.get("Recommendation", "")),
                "editor_decision": str(r.get("Editor Decision", "")),
                "editor_notes":    str(r.get("Notes", "")),
                "_from_sheets":    True,   # flag — no full report available
            })
        return loaded
    except Exception as e:
        st.warning(f"Could not load history from Sheets: {e}")
        return []


# ── Sidebar ────────────────────────────────────────────────────────────────
def sidebar():
    with st.sidebar:
        st.markdown("""
        <div class="sb-brand">
            <div class="sb-brand-icon">✦</div>
            <div>
                <div class="sb-brand-title">Content QA</div>
                <div class="sb-brand-sub">Editorial review</div>
            </div>
        </div>""", unsafe_allow_html=True)

        st.markdown('<div class="sb-section">Navigation</div>', unsafe_allow_html=True)
        page = st.radio("Navigation", ["📄  Submit article", "◫  Dashboard"],
                        label_visibility="collapsed", key="sidebar_navigation")

        st.markdown('<div class="sb-section">Deduction rules</div>', unsafe_allow_html=True)
        st.markdown("""
        <div class="sb-deduction-wrap">
            <div class="sb-deduction-card">
                <div class="sb-deduction-row"><span>Data accuracy comment</span><span class="sb-pill">-1.5</span></div>
                <div class="sb-deduction-row"><span>Missing info comment</span><span class="sb-pill">-1.5</span></div>
                <div class="sb-deduction-row"><span>Grammar / rephrasing</span><span class="sb-pill">-1</span></div>
                <div class="sb-deduction-row"><span>Plagiarism over 20%</span><span class="sb-pill">-5</span></div>
                <div class="sb-deduction-row"><span>AI content over 20%</span><span class="sb-pill">-5</span></div>
            </div>
        </div>""", unsafe_allow_html=True)

        if "Dashboard" in page:
            return "Dashboard"
        return "Submit article"


# ── Submit page ────────────────────────────────────────────────────────────
def page_submit():
    inject_css()
    st.markdown('<div class="qa-shell">', unsafe_allow_html=True)
    st.markdown("""
<div class="qa-hero">
  <div>
    <div class="qa-hero-badge">✦ Editorial QA Engine</div>
    <h1>Content QA System</h1>
    <p>Submit articles for automated review. We score editor comments, detect plagiarism, and flag AI-generated content.</p>
  </div>
  <div class="qa-hero-icon">☑</div>
</div>""", unsafe_allow_html=True)

    main_col, side_col = st.columns([3.1, 1.05], gap="large")

    with main_col:
        st.markdown("""
<div class="stepper">
  <div class="step-item active"><span class="step-num">1</span><span>Details</span></div>
  <div class="step-line"></div>
  <div class="step-item"><span class="step-num">2</span><span>Upload</span></div>
  <div class="step-line"></div>
  <div class="step-item"><span class="step-num">3</span><span>Evaluation</span></div>
  <div class="step-line"></div>
  <div class="step-item"><span class="step-num">4</span><span>Report</span></div>
</div>""", unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown("""
<div class="form-card-header">
  <div>
    <div class="form-card-title">New submission</div>
    <div class="form-card-sub">Fill in the details below and upload the article file.</div>
  </div>
  <div class="ready-badge"><span class="ready-dot"></span> Ready to submit</div>
</div>""", unsafe_allow_html=True)

            with st.form("qa_form"):
                c1, c2 = st.columns(2)
                writer      = c1.text_input("Writer name",         placeholder="e.g. Sarah Ahmed")
                editor_name = c2.text_input("Subeditor / editor",  placeholder="e.g. Mohamed Ali")

                c3, c4 = st.columns(2)
                title = c3.text_input("Article title", placeholder="e.g. Everything About Mortgages")
                ctype = c4.selectbox("Content type", CONTENT_TYPES)

                c5, c6 = st.columns(2)
                lang = c5.selectbox("Language", LANGUAGES)

                st.markdown('<span class="platform-label">Platform</span>', unsafe_allow_html=True)
                platform = st.radio("Platform", PLATFORMS, horizontal=True,
                                    label_visibility="collapsed", key="platform_choice")

                st.markdown('<div class="form-section-divider"></div>', unsafe_allow_html=True)
                st.markdown("""<div class="upload-head">
                    <div class="upload-head-label">Upload article file</div>
                    <div class="upload-head-help">?</div>
                </div>""", unsafe_allow_html=True)

                upload = st.file_uploader("Upload article file", type=["docx", "pdf", "txt"],
                                          label_visibility="collapsed")

                if upload:
                    size_mb  = upload.size / (1024 * 1024)
                    file_ext = upload.name.split(".")[-1].upper() if "." in upload.name else "FILE"
                    st.markdown(f"""
                    <div class="file-card">
                        <div class="file-icon">▤</div>
                        <div>
                            <div class="file-title">{upload.name}</div>
                            <div class="file-meta">{file_ext} • {size_mb:.1f} MB</div>
                        </div>
                        <div class="file-status">● File uploaded</div>
                    </div>""", unsafe_allow_html=True)

                writer_done = bool(writer.strip())
                editor_done = bool(editor_name.strip())
                title_done  = bool(title.strip())
                upload_done = bool(upload)

                st.markdown(f"""
                <div class="precheck">
                    <div class="precheck-item {'done' if writer_done else ''}">
                        <span class="precheck-dot">✓</span><span>Writer name</span>
                    </div>
                    <div class="precheck-item {'done' if editor_done else ''}">
                        <span class="precheck-dot">✓</span><span>Editor name</span>
                    </div>
                    <div class="precheck-item {'done' if title_done else ''}">
                        <span class="precheck-dot">✓</span><span>Article title</span>
                    </div>
                    <div class="precheck-item {'done' if upload_done else ''}">
                        <span class="precheck-dot">✓</span><span>File uploaded</span>
                    </div>
                </div>""", unsafe_allow_html=True)

                go = st.form_submit_button("✦  Run full evaluation",
                                           use_container_width=True, type="primary")

    with side_col:
        st.markdown("""
        <div class="side-card">
            <div class="side-card-title">What happens next?</div>
            <div class="timeline-row">
                <div class="timeline-num">1</div>
                <div><div class="timeline-title">We analyze your article</div>
                     <div class="timeline-sub">Checking editor comments, plagiarism and AI content.</div></div>
            </div>
            <div class="timeline-row">
                <div class="timeline-num">2</div>
                <div><div class="timeline-title">Calculate scores</div>
                     <div class="timeline-sub">Each category is scored automatically.</div></div>
            </div>
            <div class="timeline-row" style="margin-bottom:0">
                <div class="timeline-num">3</div>
                <div><div class="timeline-title">Get your report</div>
                     <div class="timeline-sub">Review results and recommendations.</div></div>
            </div>
        </div>
        <div class="side-card">
            <div class="side-card-title">Evaluation preview</div>
            <div class="preview-grid">
                <div><div class="ring" style="--ring-color:#5b5ce2;--ring-value:85%">85</div>
                     <div class="preview-label">Editor<br>Comments</div></div>
                <div><div class="ring" style="--ring-color:#ef4444;--ring-value:12%">12%</div>
                     <div class="preview-label">Plagiarism</div></div>
                <div><div class="ring" style="--ring-color:#4f46e5;--ring-value:8%">8%</div>
                     <div class="preview-label">AI Content</div></div>
            </div>
        </div>
        <div class="side-card">
            <div class="tip-box">
                <div class="tip-title">Tip</div>
                <div>Ensure your article is final and editor comments are saved inside the .docx file.</div>
            </div>
        </div>""", unsafe_allow_html=True)

    if not go:
        st.info("Upload a .docx file with editor comments. Scores are based entirely on those comments.")
        st.markdown('</div>', unsafe_allow_html=True)
        return

    if not writer or not title or not upload:
        st.error("Please fill in writer name, title and upload a file.")
        st.markdown('</div>', unsafe_allow_html=True)
        return

    with st.spinner("Reading file..."):
        parsed = parse_file(upload)

    if not parsed["text"] or len(parsed["text"]) < 30:
        st.error(f"Could not read text from file. {parsed.get('error', '')}")
        st.markdown('</div>', unsafe_allow_html=True)
        return

    writer_reply_keywords = ["fixed", "done", "added", "removed", "replaced", "updated",
        "changed", "edited", "deleted", "corrected", "revised", "@",
        "noted", "ok ", "okay", "sure", "will do"]

    def is_writer_reply(comment_text):
        low = comment_text.lower().strip()
        if low.startswith("@"):
            return True
        if len(low) < 30:
            for kw in writer_reply_keywords:
                if low.startswith(kw) or kw in low[:20]:
                    return True
        return False

    all_comments    = parsed["comments"]
    editor_comments = [c for c in all_comments if not is_writer_reply(c["text"])]
    parsed["comments"]     = editor_comments
    parsed["all_comments"] = all_comments

    with st.expander(f"Extracted — {len(parsed['headings'])} headings · {len(parsed['links'])} links · {len(parsed['comments'])} editor comments"):
        col_h, col_l, col_c = st.columns(3)
        with col_h:
            st.markdown("**Headings**")
            for h in parsed["headings"]:
                st.markdown(f"`{h['level']}` {h['text']}")
            if not parsed["headings"]:
                st.caption("None detected")
        with col_l:
            st.markdown("**Links**")
            for l in parsed["links"][:6]:
                st.markdown(f"- {l}")
            if not parsed["links"]:
                st.caption("None detected")
        with col_c:
            st.markdown("**Editor comments**")
            for idx, c in enumerate(parsed["comments"], 1):
                st.markdown(
                    f'<div class="cmt-card"><span class="cmt-author">Comment {idx} — {c["author"]}</span><br>'
                    f'{c["text"]}<div class="cmt-deduct">points deducted</div></div>',
                    unsafe_allow_html=True)
            if not parsed["comments"]:
                st.caption("No comments found")

    prog = st.progress(0, text="Starting...")

    try:
        prog.progress(15, text="Reading document and extracting editor comments...")
        qa = run_qa(title, parsed["text"], writer, ctype, lang, platform,
                    parsed["headings"], parsed["links"], parsed["comments"])
    except Exception as e:
        st.error(f"AI evaluation failed: {e}")
        st.info("Make sure GROQ_API_KEY is set in Streamlit Secrets.")
        st.markdown('</div>', unsafe_allow_html=True)
        return

    prog.progress(50, text="Checking plagiarism...")
    plag = check_plagiarism(parsed["text"], parsed["links"])
    plag_sources, plag_snippets = get_plag_snippets(parsed["text"], parsed["links"], plag)

    prog.progress(72, text="Checking AI-generated content...")
    ai = check_ai(parsed["text"])

    prog.progress(95, text="Calculating final score...")
    base_score  = qa.get("total", 0)
    final_score, deductions = apply_deductions(
        base_score, parsed["comments"], plag["percentage"], ai["ai_pct"])
    recommendation = get_recommendation(final_score)

    prog.progress(100, text="Done.")
    prog.empty()

    sub = {
        "date":           datetime.now().strftime("%d %b %Y %H:%M"),
        "platform":       platform,
        "writer":         writer,
        "editor_name":    editor_name,
        "title":          title,
        "content_type":   ctype,
        "language":       lang,
        "word_count":     parsed["word_count"],
        "headings":       parsed["headings"],
        "links":          parsed["links"],
        "comments":       parsed["comments"],
        "qa":             qa,
        "plagiarism":     plag,
        "plag_snippets":  plag_snippets,
        "plag_sources":   plag_sources,
        "ai_detection":   ai,
        "deductions":     deductions,
        "qa_score":       final_score,
        "plagiarism_pct": plag["percentage"],
        "ai_pct":         ai["ai_pct"],
        "recommendation": recommendation,
        "editor_decision": "",
        "editor_notes":    "",
        "_from_sheets":    False,
    }

    if "submissions" not in st.session_state:
        st.session_state.submissions = []
    st.session_state.submissions.append(sub)
    log_to_sheets(sub)

    st.success(f"Evaluation complete — Final score: **{final_score} / 100**")
    render_report(sub)
    st.markdown('</div>', unsafe_allow_html=True)


# ── Report ─────────────────────────────────────────────────────────────────
def render_report(sub):
    inject_css()
    qa    = sub["qa"]
    plag  = sub["plagiarism"]
    ai    = sub["ai_detection"]
    ded   = sub["deductions"]
    score = sub["qa_score"]
    grade = get_grade(score)
    rec   = sub["recommendation"]
    plag_snippets = sub.get("plag_snippets", [])
    plag_sources  = sub.get("plag_sources", [])

    st.divider()
    bdg_class = "bdg-bay" if sub["platform"] == "Bayut" else "bdg-dub"
    editor_html = f"&nbsp; 👤 <strong>{sub.get('editor_name','—')}</strong>" if sub.get('editor_name') else ""
    st.markdown(
        f"**{sub['writer']}** &nbsp; <span class='bdg {bdg_class}'>{sub['platform']}</span> &nbsp;"
        f"`{sub['content_type']}` &nbsp; `{sub['language']}` &nbsp; `{sub['word_count']} words`"
        f"{editor_html} &nbsp; `{sub['date']}`",
        unsafe_allow_html=True)
    st.markdown("")

    rec_labels = {
        "approve": ("Approve",           "#d1fae5", "#065f46"),
        "revise":  ("Request revision",  "#fef3c7", "#92400e"),
        "reject":  ("Reject",            "#fee2e2", "#991b1b"),
    }
    rl, rbg, rtc = rec_labels.get(rec, rec_labels["revise"])

    def brow(cls, label, val):
        return f'<div class="{cls}"><span>{label}</span><span>{val}</span></div>'

    classified    = ded.get("classified", [])
    data_acc_cmts = [c for c in classified if c["type"] == "Data accuracy"]
    missing_cmts  = [c for c in classified if c["type"] == "Missing info"]
    grammar_cmts  = [c for c in classified if c["type"] == "Grammar / rephrasing"]

    comment_rows = ""
    if data_acc_cmts:
        comment_rows += brow("ded-row", f'Data accuracy ({len(data_acc_cmts)} × 1.5 pts)', f'- {round(len(data_acc_cmts)*1.5,1)} pts')
    if missing_cmts:
        comment_rows += brow("ded-row", f'Missing info ({len(missing_cmts)} × 1.5 pts)',   f'- {round(len(missing_cmts)*1.5,1)} pts')
    if grammar_cmts:
        comment_rows += brow("ded-row", f'Grammar / rephrasing ({len(grammar_cmts)} × 1 pt)', f'- {len(grammar_cmts)} pts')
    if not classified:
        comment_rows = brow("ok-row", "Editor comments", "no deduction")

    plag_b   = ded.get("plag_brackets", 0)
    plag_row = (brow("ded-row",
        f'Plagiarism {ded["plag_pct"]}% ({plag_b} × 20% bracket × 5 pts)',
        f'- {ded["plag_deduction"]} pts')
        if ded["plag_deduction"] > 0
        else brow("ok-row", f'Plagiarism {ded["plag_pct"]}% — under 20%', "no deduction"))

    ai_b   = ded.get("ai_brackets", 0)
    ai_row = (brow("ded-row",
        f'AI content {ded["ai_pct"]}% ({ai_b} × 20% bracket × 5 pts)',
        f'- {ded["ai_deduction"]} pts')
        if ded["ai_deduction"] > 0
        else brow("ok-row", f'AI content {ded["ai_pct"]}% — under 20%', "no deduction"))

    bd = (brow("base-row", "Base score", f'{ded["base_score"]} / 100') +
          comment_rows + plag_row + ai_row +
          brow("total-row", "Final score", f"{score} / 100"))

    st.markdown(
        f'<div class="score-hero">'
        f'<div class="score-num">{score}<span class="score-den"> / 100</span></div>'
        f'<div class="score-grade">{grade}</div>'
        f'<div style="display:inline-block;margin:6px 0 8px;padding:3px 12px;border-radius:20px;'
        f'background:{rbg};color:{rtc};font-size:11px;font-weight:500">{rl}</div>'
        f'<div class="score-verdict">{qa.get("overall_feedback","")}</div>'
        f'<div class="breakdown-box">{bd}</div></div>',
        unsafe_allow_html=True)

    st.divider()
    st.markdown("#### Plagiarism and AI detection")
    pc1, pc2 = st.columns(2)

    with pc1:
        pp   = plag["percentage"]
        over = pp > 20
        col  = "#dc2626" if over else "#059669"
        thresh = (f'<span class="detect-thresh" style="background:#fee2e2;color:#991b1b">{pp}% — over 20% — 5 pts deducted</span>'
                  if over else
                  f'<span class="detect-thresh" style="background:#d1fae5;color:#065f46">{pp}% — under 20% — no deduction</span>')
        snip_html = ""
        if plag_snippets or plag_sources:
            lbl = "Copied content detected" if over else "Suspicious brochure language"
            snip_html = f'<div class="issue-block"><div class="issue-block-title">{lbl}</div>'
            for src in plag_sources[:3]:
                snip_html += f'<div class="issue-snippet"><strong style="color:#92400e">Source matched:</strong> {src}</div>'
            for s in plag_snippets[:8]:
                snip_html += f'<div class="issue-snippet">{highlight_plag(s)}</div>'
            snip_html += '</div>'
        st.markdown(
            f'<div class="detect-card">'
            f'<div class="detect-title">Plagiarism check <span style="font-size:10px;color:#888;margin-left:6px">via {plag.get("source","heuristic")}</span></div>'
            f'<div class="detect-bar"><div class="detect-bar-f" style="width:{min(pp,100)}%;background:{col}"></div></div>'
            f'{thresh}<div class="detect-note">{"Rewrite all flagged sections completely." if over else "Content is within acceptable range."}</div>'
            f'{snip_html}</div>', unsafe_allow_html=True)

    with pc2:
        ap      = ai["ai_pct"]
        ai_over = ap > 20
        a_col   = "#dc2626" if ai_over else "#059669"
        src_lbl = ai.get("source", "heuristic")
        a_thresh = (f'<span class="detect-thresh" style="background:#fee2e2;color:#991b1b">{ap}% — over 20% — 5 pts deducted</span>'
                    if ai_over else
                    f'<span class="detect-thresh" style="background:#d1fae5;color:#065f46">{ap}% — under 20% — no deduction</span>')
        ai_sents     = ai.get("ai_sentences", [])
        ai_snip_html = ""
        if ai_sents:
            ai_snip_html = '<div class="issue-block"><div class="issue-block-title">Flagged sentences</div>'
            for s in ai_sents[:5]:
                ai_snip_html += f'<div class="issue-snippet">{highlight_ai(s)}</div>'
            ai_snip_html += '</div>'
        st.markdown(
            f'<div class="detect-card">'
            f'<div class="detect-title">AI detection <span style="font-size:10px;color:#888;margin-left:6px">via {src_lbl}</span></div>'
            f'<div class="detect-bar"><div class="detect-bar-f" style="width:{min(ap,100)}%;background:{a_col}"></div></div>'
            f'{a_thresh}<div class="detect-note">{"High AI content detected." if ai_over else "Content appears mostly human-written."}</div>'
            f'<div class="detect-split">'
            f'<div class="detect-seg"><div class="detect-seg-n" style="color:#059669">{ai["human_pct"]}%</div><div class="detect-seg-l">Human</div></div>'
            f'<div class="detect-seg"><div class="detect-seg-n" style="color:{a_col}">{ap}%</div><div class="detect-seg-l">AI likely</div></div>'
            f'</div>{ai_snip_html}</div>', unsafe_allow_html=True)

    st.divider()
    st.markdown("#### Category scores")
    if not sub["comments"]:
        st.markdown('<div class="no-cmt-notice">No editor comments found. All categories awarded full marks.</div>', unsafe_allow_html=True)

    for cat, mx in CAT_MAX.items():
        data = qa["scores"].get(cat, {})
        s    = data.get("score", 0)
        fb   = data.get("feedback", "")
        refs = data.get("comment_refs", [])
        ref_html = " ".join(f'<span class="cat-ref">Comment {r}</span>' for r in refs)
        ca, cb = st.columns([4, 1])
        ca.markdown(f"**{cat}**" + (f" &nbsp; {ref_html}" if ref_html else ""), unsafe_allow_html=True)
        ca.progress(s / mx)
        cb.markdown(f"**{s} / {mx}**")
        st.caption(fb)
        st.markdown("")

    st.divider()
    st.markdown("#### Document structure")
    sc1, sc2, sc3 = st.columns(3)
    sc1.metric("Headings",    len(sub["headings"]))
    sc2.metric("Total links", len(sub["links"]))
    internal = [l for l in sub["links"] if sub["platform"].lower() in l.lower()]
    sc3.metric("Internal links", len(internal))

    if sub["comments"]:
        classified = ded.get("classified", [])
        cmap = {c["text"]: c for c in classified}
        type_colors = {
            "Data accuracy":       ("#fee2e2", "#991b1b", "- 1.5 pts"),
            "Missing info":        ("#fef3c7", "#92400e", "- 1.5 pts"),
            "Grammar / rephrasing":("#f0f4ff", "#2D4A8A", "- 1 pt"),
        }
        st.markdown(f"**Editor comments — {len(sub['comments'])} found**")
        for idx, c in enumerate(sub["comments"], 1):
            info         = cmap.get(c["text"], {})
            ctype_label  = info.get("type", "Grammar / rephrasing")
            bg, tc, pts  = type_colors.get(ctype_label, ("#f5f6fa", "#555", "- 1 pt"))
            st.markdown(
                f'<div class="cmt-card">'
                f'<span class="cmt-author">Comment {idx} — {c["author"]}</span>'
                f'<span style="font-size:10px;font-weight:500;padding:1px 8px;border-radius:20px;'
                f'background:{bg};color:{tc};margin-left:8px">{ctype_label}</span><br>'
                f'{c["text"]}'
                f'<div class="cmt-deduct">{pts} deducted</div></div>',
                unsafe_allow_html=True)

    st.divider()
    col_s, col_i = st.columns(2)
    with col_s:
        st.markdown("#### Strengths")
        strengths = qa.get("key_strengths", [])
        for s in strengths:
            st.markdown(f'<span class="tag-str">{s}</span>', unsafe_allow_html=True)
        if not strengths:
            st.caption("Add editor comments to see strengths.")
    with col_i:
        st.markdown("#### Required improvements")
        for imp in qa.get("areas_for_improvement", []):
            st.markdown(f'<span class="tag-imp">{imp}</span>', unsafe_allow_html=True)

    suggestions = qa.get("suggestions", [])
    if suggestions:
        st.divider()
        st.markdown("#### Suggestions to improve the article")
        for sug in suggestions:
            st.markdown(
                f'<div class="suggest-item"><div class="suggest-num">{sug.get("number","")}</div>'
                f'<div><div>{sug.get("action","")}</div>'
                f'<div class="suggest-cat">Addresses: {sug.get("category","")}</div></div></div>',
                unsafe_allow_html=True)

    st.divider()
    st.markdown("#### Editor decision")
    st.caption("The AI recommendation is a guide. You make the final call.")
    rec_idx  = {"approve": 0, "revise": 1, "reject": 2}
    decision = st.radio("Decision", ["Approve", "Request revision", "Reject"],
                        index=rec_idx.get(rec, 1), horizontal=True,
                        key=f"dec_{sub['title']}_{sub['date']}")
    notes = st.text_area("Notes for writer (required for revision and rejection)", height=90,
                         placeholder="Tell the writer exactly what to fix.",
                         key=f"notes_{sub['title']}_{sub['date']}")
    if st.button("Confirm decision", type="primary", use_container_width=True,
                 key=f"conf_{sub['title']}_{sub['date']}"):
        if decision in ("Request revision", "Reject") and not notes.strip():
            st.error("Please add notes before confirming.")
        else:
            sub["editor_decision"] = decision
            sub["editor_notes"]    = notes
            update_sheet_decision(sub)
            st.success(f"Decision saved: {decision}")
            if notes:
                st.info(f"Notes for {sub['writer']}: {notes}")

    st.caption(f"Content QA System — {sub['platform']} — Powered by Groq — {sub['date']}")


# ── Score colour helper ─────────────────────────────────────────────────────
def _score_color(score):
    if score >= 80: return "#059669"
    if score >= 60: return "#d97706"
    return "#dc2626"


def _dec_class(dec):
    if dec == "Approve":          return "dec-approve"
    if dec == "Request revision": return "dec-revise"
    if dec == "Reject":           return "dec-reject"
    return "dec-pending"


# ── Dashboard ──────────────────────────────────────────────────────────────
def page_dashboard():
    inject_css()
    st.markdown("""
<div class="qa-hero">
  <div>
    <div class="qa-hero-badge">Overview</div>
    <h1>Dashboard</h1>
    <p>All submissions across Bayut and Dubizzle — persisted from Google Sheets.</p>
  </div>
  <div class="qa-hero-icon">📊</div>
</div>""", unsafe_allow_html=True)

    all_subs = st.session_state.get("submissions", [])

    if not all_subs:
        st.info("No submissions yet. Submit an article to get started.")
        return

    # ── Stats row ──────────────────────────────────────────────────────────
    total    = len(all_subs)
    approved = sum(1 for s in all_subs if s.get("editor_decision") == "Approve")
    revision = sum(1 for s in all_subs if s.get("editor_decision") == "Request revision")
    rejected = sum(1 for s in all_subs if s.get("editor_decision") == "Reject")
    pending  = sum(1 for s in all_subs if not s.get("editor_decision"))
    avg_score = round(sum(s.get("qa_score", 0) for s in all_subs) / max(total, 1), 1)

    st.markdown(f"""
<div class="dash-stats-row">
  <div class="dash-stat blue"><div class="dash-stat-num">{total}</div><div class="dash-stat-lbl">Total</div></div>
  <div class="dash-stat green"><div class="dash-stat-num">{approved}</div><div class="dash-stat-lbl">Approved</div></div>
  <div class="dash-stat amber"><div class="dash-stat-num">{revision}</div><div class="dash-stat-lbl">Revision</div></div>
  <div class="dash-stat red"><div class="dash-stat-num">{rejected}</div><div class="dash-stat-lbl">Rejected</div></div>
  <div class="dash-stat"><div class="dash-stat-num">{pending}</div><div class="dash-stat-lbl">Pending</div></div>
  <div class="dash-stat blue"><div class="dash-stat-num">{avg_score}</div><div class="dash-stat-lbl">Avg Score</div></div>
</div>""", unsafe_allow_html=True)

    # ── Filters ────────────────────────────────────────────────────────────
    all_writers = sorted(set(s["writer"]      for s in all_subs if s.get("writer")))
    all_editors = sorted(set(s.get("editor_name", "") for s in all_subs if s.get("editor_name")))

    fc1, fc2, fc3, fc4, fc5, fc6 = st.columns(6)
    wf  = fc1.selectbox("Writer",           ["All"] + all_writers,          key="dash_writer")
    pf  = fc2.selectbox("Platform",         ["All"] + PLATFORMS,            key="dash_platform")
    ef  = fc3.selectbox("Editor / subeditor", ["All"] + all_editors,        key="dash_editor")
    tf  = fc4.selectbox("Content type",     ["All"] + CONTENT_TYPES,        key="dash_type")
    lf  = fc5.selectbox("Language",         ["All"] + LANGUAGES,            key="dash_lang")
    sf  = fc6.selectbox("Decision",         ["All", "Pending", "Approve", "Request revision", "Reject"], key="dash_status")

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

    # ── Article cards ──────────────────────────────────────────────────────
    for sub in reversed(filtered):
        score      = sub.get("qa_score", 0)
        dec        = sub.get("editor_decision") or "Pending"
        ring_color = _score_color(score)
        ring_val   = f"{int(score)}%"
        dec_cls    = _dec_class(dec)

        plat_cls = "bay" if sub.get("platform") == "Bayut" else "dub"
        lang_cls = "eng" if sub.get("language") == "English" else "ara"

        # Brief score summary
        ded          = sub.get("deductions", {})
        cmt_ded      = ded.get("comment_deduction", 0)
        plag_ded     = ded.get("plag_deduction", 0)
        ai_ded       = ded.get("ai_deduction", 0)
        cmt_count    = ded.get("comment_count", 0)
        plag_pct     = sub.get("plagiarism_pct", 0)
        ai_pct       = sub.get("ai_pct", 0)
        overall_fb   = sub.get("qa", {}).get("overall_feedback", "")

        # Build a compact 1-line summary
        parts = []
        if cmt_count:
            parts.append(f"{cmt_count} comment{'s' if cmt_count != 1 else ''} (−{cmt_ded} pts)")
        if plag_ded:
            parts.append(f"plagiarism {plag_pct}% (−{plag_ded} pts)")
        if ai_ded:
            parts.append(f"AI {ai_pct}% (−{ai_ded} pts)")
        if not parts:
            parts = ["No deductions applied"]
        score_brief = " · ".join(parts)

        # Truncated feedback
        fb_text = overall_fb[:160] + ("…" if len(overall_fb) > 160 else "") if overall_fb else score_brief

        editor_chip = f'<span class="meta-chip">👤 {sub["editor_name"]}</span>' if sub.get("editor_name") else ""
        grade_label = get_grade(score).split(" — ")[-1]  # e.g. "Good"

        card_html = f"""
<div class="article-card">
  <div class="article-card-left">
    <div class="article-card-title">{sub.get('title','Untitled')}</div>
    <div class="article-card-meta">
      <span class="meta-chip">✍️ {sub.get('writer','—')}</span>
      {editor_chip}
      <span class="meta-chip {plat_cls}">{sub.get('platform','')}</span>
      <span class="meta-chip {lang_cls}">{sub.get('language','')}</span>
      <span class="meta-chip">{sub.get('content_type','')}</span>
      <span class="meta-chip">{sub.get('word_count',0)} words</span>
      <span class="meta-chip">🗓 {sub.get('date','')}</span>
    </div>
    <div class="article-card-summary">
      <strong style="color:#374151">{score_brief}</strong>
      {"<br><span style='color:#9ca3af'>" + fb_text + "</span>" if overall_fb and overall_fb != score_brief else ""}
    </div>
  </div>
  <div class="article-card-right">
    <div class="score-ring-wrap">
      <div class="score-ring" style="--rc:{ring_color};--rv:{ring_val}">{int(score)}</div>
      <div class="score-ring-lbl">{grade_label}</div>
    </div>
    <span class="dec-badge {dec_cls}">{dec}</span>
  </div>
</div>"""
        st.markdown(card_html, unsafe_allow_html=True)

        # Expand full report only for in-session submissions (not Sheets-only rows)
        if not sub.get("_from_sheets", True):
            with st.expander("View full report"):
                render_report(sub)
        else:
            # Sheets-only: show a condensed deduction table
            with st.expander("View score breakdown"):
                st.markdown(f"""
| | |
|---|---|
| **Writer** | {sub.get('writer','—')} |
| **Editor** | {sub.get('editor_name','—')} |
| **Platform** | {sub.get('platform','—')} |
| **Final score** | **{score} / 100** |
| **Grade** | {get_grade(score)} |
| **Comment deduction** | −{cmt_ded} pts ({cmt_count} comments) |
| **Plagiarism deduction** | −{plag_ded} pts ({plag_pct}%) |
| **AI deduction** | −{ai_ded} pts ({ai_pct}%) |
| **Recommendation** | {sub.get('recommendation','').capitalize()} |
| **Decision** | {dec} |
| **Notes** | {sub.get('editor_notes','—')} |
""")
                if overall_fb:
                    st.caption(f"**AI feedback:** {overall_fb}")


# ── Main ───────────────────────────────────────────────────────────────────
def main():
    if "submissions" not in st.session_state:
        st.session_state.submissions = []

    # Load historical data from Sheets once per session
    if "sheets_loaded" not in st.session_state:
        st.session_state.sheets_loaded = False

    if not st.session_state.sheets_loaded:
        with st.spinner("Loading submission history from Google Sheets…"):
            historical = load_from_sheets()

        if historical:
            existing_keys = {
                (s["writer"], s["title"], s["date"])
                for s in st.session_state.submissions
            }
            added = 0
            for h in historical:
                key = (h["writer"], h["title"], h["date"])
                if key not in existing_keys:
                    st.session_state.submissions.append(h)
                    added += 1

        st.session_state.sheets_loaded = True

    page = sidebar()
    if "Submit" in page:
        page_submit()
    else:
        page_dashboard()


if __name__ == "__main__":
    main()
