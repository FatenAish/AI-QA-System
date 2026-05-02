import streamlit as st

import json

import re

import math

import os

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

    "in conclusion","it is worth noting","it is important to note","delve into",

    "in the realm of","furthermore","moreover","needless to say","leverage",

    "utilize","seamlessly","it goes without saying","in today's","one such","robust",

    "cutting-edge","state-of-the-art","at the end of the day","effortlessly blends",

    "embody contemporary elegance","functional vitality","architectural lines","expansive glazing",

]



BROCHURE_PHRASES = [

    "wellness-oriented","highly anticipated","masterplan","effortlessly blends",

    "distinguished residential","dynamic enclave","lush landscaped buffers",

    "signature communal","elevated everyday living","embody contemporary elegance",

    "functional vitality","architectural lines","expansive glazing",

    "highly customisable aesthetic","light and dark material finishes",

    "open-plan configurations","smart-home integrations","forward-looking environmental",

    "eco-living standards","dark sky-compliant","energy-efficient building methods",

    "smart irrigation","pedestrian-friendly trails","responsible, sustainable and healthy",

    "certainly. here are","amenities mentioned in the brochure","define the next chapter",

    "dynamic urban living","tranquillity of expansive greenery","wellness-oriented enclave",

    "eco-friendly spaces","self-sustaining","immersive community experience",

    "active design principles","modern sanctuary","lush landscape of parks",

    "culture, leisure and active","fabric of daily life","peaceful seclusion",

    "dynamic pulse","world-class community amenities","premium off-plan homes",

    "strategically located","seamless connectivity","metropolitan accessibility",

    "opulent master suite","contemporary elegance","smart-home integration",

    "lush landscape","boasts","unparalleled","premium lifestyle","setting the benchmark",

]



KNOWN_DOMAINS = [

    "emaar.com","nakheel.com","damac.com","aldar.com","meraas.com",

    "sobha.com","omniyat.com","ellington.ae","azizi.ae",

]



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

.sb-deduction-wrap{padding:0 14px}

.sb-deduction-card{background:#fff;border:1px solid #dfe4ea;border-radius:16px;overflow:hidden}

.sb-deduction-row{display:flex;align-items:center;justify-content:space-between;padding:11px 13px;border-bottom:1px solid #eef2f7;font-size:12px;color:#374151}

.sb-deduction-row:last-child{border-bottom:none}

.sb-pill{background:#fee2e2;color:#ef4444;font-size:11px;font-weight:800;border-radius:999px;padding:2px 9px}

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

div[data-testid="stVerticalBlockBorderWrapper"]>div{border-radius:22px !important}

div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stVerticalBlock"]{gap:0.9rem !important}

</style>

""", unsafe_allow_html=True)





# ── Groq AI ────────────────────────────────────────────────────────────────

GROQ_MODELS = ["llama-3.3-70b-versatile", "llama-3.1-8b-instant", "llama3-8b-8192", "gemma2-9b-it"]



def call_ai(prompt, max_retries=3):

    """Call Groq with model fallback and retry on empty responses."""

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

    """Extract and parse the first JSON object from a model response."""

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





# ── Deterministic scoring ──────────────────────────────────────────────────
# Score is based ONLY on editor comments — no AI/plagiarism deductions.

def classify_comment(text):

    """Keyword-based classification — always returns the same result for the same text."""

    low = text.lower()

    for kw in ["wrong","incorrect","not correct","inaccurate","error","should be","it is","it's",

               "the source","in the source","copied","from google","from maps","url goes","link goes",

               "apartments","no apartments","mins away","minutes away","under construction","off-plan","data","fact"]:

        if kw in low: return "Data accuracy", 1.5

    for kw in ["missing","add","please add","include","mention","not mentioned","should mention",

               "we need","please mention","go through","available","please write","notable projects",

               "specific","more details","lacks","header","section"]:

        if kw in low: return "Missing info", 1.5

    for kw in ["grammar","rephrase","rewrite","word","sentence","phrasing","general","too general",

               "vague","unclear","confusing","brand voice","tone","style","read","sounds"]:

        if kw in low: return "Grammar / rephrasing", 1.0

    return "Grammar / rephrasing", 1.0



def apply_deductions(comments):

    """

    Final score = 100 − comment_deductions only.

    No AI or plagiarism deductions.

    """

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





# ── AI — feedback text only (does not affect score) ────────────────────────

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

            "overall_feedback": f"AI feedback unavailable. {len(comments)} editor comment(s) found; deductions applied automatically.",

            "key_strengths": [], "areas_for_improvement": [c["text"][:80] for c in comments[:3]], "suggestions": []}



# ── Sidebar ────────────────────────────────────────────────────────────────

def sidebar():

    with st.sidebar:

        st.markdown('<div class="sb-brand"><div class="sb-brand-icon">✦</div><div><div class="sb-brand-title">Content QA</div><div class="sb-brand-sub">Editorial review</div></div></div>', unsafe_allow_html=True)

        st.markdown('<div class="sb-section">Navigation</div>', unsafe_allow_html=True)

        page = st.radio("Navigation", ["📄  Submit article", "◫  Dashboard"],

                        label_visibility="collapsed", key="sidebar_navigation")

        st.markdown('<div class="sb-section">Deduction rules</div>', unsafe_allow_html=True)

        st.markdown("""<div class="sb-deduction-wrap"><div class="sb-deduction-card">

            <div class="sb-deduction-row"><span>Data accuracy comment</span><span class="sb-pill">−1.5</span></div>

            <div class="sb-deduction-row"><span>Missing info comment</span><span class="sb-pill">−1.5</span></div>

            <div class="sb-deduction-row"><span>Grammar / rephrasing</span><span class="sb-pill">−1</span></div>

        </div></div>""", unsafe_allow_html=True)

        return "Dashboard" if "Dashboard" in page else "Submit article"



# ── Submit page ────────────────────────────────────────────────────────────

def page_submit():

    inject_css()

    st.markdown('<div class="qa-hero"><div><div class="qa-hero-badge">✦ Editorial QA Engine</div><h1>Content QA System</h1><p>Submit articles for automated review — editor comments drive the score.</p></div><div class="qa-hero-icon">☑</div></div>', unsafe_allow_html=True)

    main_col, side_col = st.columns([3.1, 1.05], gap="large")

    with main_col:

        st.markdown('<div class="stepper"><div class="step-item active"><span class="step-num">1</span><span>Details</span></div><div class="step-line"></div><div class="step-item"><span class="step-num">2</span><span>Upload</span></div><div class="step-line"></div><div class="step-item"><span class="step-num">3</span><span>Evaluation</span></div><div class="step-line"></div><div class="step-item"><span class="step-num">4</span><span>Report</span></div></div>', unsafe_allow_html=True)

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

                go = st.form_submit_button("✦  Run full evaluation",

                                           use_container_width=True, type="primary")

    with side_col:

        st.markdown("""<div class="side-card">

  <div class="side-card-title">What happens next?</div>

  <div class="timeline-row"><div class="timeline-num">1</div><div><div class="timeline-title">Analyze article</div><div class="timeline-sub">Editor comments extracted and classified.</div></div></div>

  <div class="timeline-row"><div class="timeline-num">2</div><div><div class="timeline-title">Calculate scores</div><div class="timeline-sub">Each category scored automatically.</div></div></div>

  <div class="timeline-row" style="margin-bottom:0"><div class="timeline-num">3</div><div><div class="timeline-title">Get your report</div><div class="timeline-sub">Review results and confirm decision.</div></div></div>

</div>

<div class="side-card"><div class="tip-box"><div class="tip-title">Tip</div>Save editor comments inside the .docx before uploading.</div></div>""", unsafe_allow_html=True)

    if not go:

        st.info("Upload a .docx with editor comments. Scores are based entirely on those comments.")

        return

    if not writer or not title or not upload:

        st.error("Please fill in writer name, title and upload a file.")

        return

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

    with st.expander(f"Extracted — {len(parsed['headings'])} headings · {len(parsed['links'])} links · {len(parsed['comments'])} editor comments"):

        col_h, col_l, col_c = st.columns(3)

        with col_h:

            st.markdown("**Headings**")

            for h in parsed["headings"]: st.markdown(f"`{h['level']}` {h['text']}")

            if not parsed["headings"]: st.caption("None")

        with col_l:

            st.markdown("**Links**")

            for l in parsed["links"][:6]: st.markdown(f"- {l}")

            if not parsed["links"]: st.caption("None")

        with col_c:

            st.markdown("**Editor comments**")

            for idx, c in enumerate(parsed["comments"], 1):

                st.markdown(f'<div class="cmt-card"><span class="cmt-author">Comment {idx} — {c["author"]}</span><br>{c["text"]}</div>', unsafe_allow_html=True)

            if not parsed["comments"]: st.caption("None")

    prog = st.progress(0, text="Starting…")

    prog.progress(20, text="Calculating score…")

    final_score, deductions = apply_deductions(parsed["comments"])

    recommendation          = get_recommendation(final_score)

    prog.progress(50, text="Getting AI feedback…")

    qa = run_qa_feedback(title, parsed["text"], writer, ctype, lang, platform,

                         parsed["headings"], parsed["links"], parsed["comments"])

    prog.progress(100, text="Done."); prog.empty()

    sub = {

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



# ── Report ─────────────────────────────────────────────────────────────────

def render_report(sub):

    inject_css()

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

    classified = ded.get("classified", []); cmap = {c["text"]: c for c in classified}

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

    if not sub["comments"]:

        st.markdown('<div class="no-cmt-notice">No editor comments — all categories awarded full marks.</div>', unsafe_allow_html=True)

    for cat, mx in CAT_MAX.items():

        data = qa["scores"].get(cat, {}); s = data.get("score", 0)

        fb   = data.get("feedback", ""); refs = data.get("comment_refs", [])

        ref_html = " ".join(f'<span class="cat-ref">Comment {r}</span>' for r in refs)

        ca, cb = st.columns([4, 1])

        ca.markdown(f"**{cat}**" + (f" &nbsp; {ref_html}" if ref_html else ""), unsafe_allow_html=True)

        ca.progress(s / mx); cb.markdown(f"**{s} / {mx}**"); st.caption(fb); st.markdown("")

    st.divider()

    sc1, sc2, sc3 = st.columns(3)

    sc1.metric("Headings",     len(sub["headings"]))

    sc2.metric("Total links",  len(sub["links"]))

    sc3.metric("Internal links", len([l for l in sub["links"] if sub["platform"].lower() in l.lower()]))

    if sub["comments"]:

        type_colors = {"Data accuracy":        ("#fee2e2","#991b1b","−1.5 pts"),

                       "Missing info":          ("#fef3c7","#92400e","−1.5 pts"),

                       "Grammar / rephrasing":  ("#f0f4ff","#2D4A8A","−1 pt")}

        st.markdown(f"**Editor comments — {len(sub['comments'])} found**")

        for idx, c in enumerate(sub["comments"], 1):

            info        = cmap.get(c["text"], {})

            ctype_label = info.get("type", "Grammar / rephrasing")

            bg, tc, pts = type_colors.get(ctype_label, ("#f5f6fa","#555","−1 pt"))

            st.markdown(

                f'<div class="cmt-card"><span class="cmt-author">Comment {idx} — {c["author"]}</span>'

                f'<span style="font-size:10px;font-weight:500;padding:1px 8px;border-radius:20px;background:{bg};color:{tc};margin-left:8px">{ctype_label}</span>'

                f'<br>{c["text"]}<div class="cmt-deduct">{pts} deducted</div></div>', unsafe_allow_html=True)

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

    st.caption(f"Content QA System — {sub['platform']} — Powered by Groq — {sub['date']}")



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

    wf = fc1.selectbox("Writer",             ["All"] + all_writers,   key="dash_writer")

    pf = fc2.selectbox("Platform",           ["All"] + PLATFORMS,     key="dash_platform")

    ef = fc3.selectbox("Editor / subeditor", ["All"] + all_editors,   key="dash_editor")

    tf = fc4.selectbox("Content type",       ["All"] + CONTENT_TYPES, key="dash_type")

    lf = fc5.selectbox("Language",           ["All"] + LANGUAGES,     key="dash_lang")

    sf = fc6.selectbox("Decision",           ["All","Pending","Approve","Request revision","Reject"], key="dash_status")

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

        overall_fb = sub.get("qa", {}).get("overall_feedback", "")

        parts = []

        if cmt_count: parts.append(f"{cmt_count} comment{'s' if cmt_count!=1 else ''} (−{cmt_ded} pts)")

        if not parts: parts = ["No deductions"]

        score_brief = " · ".join(parts)

        fb_preview  = (overall_fb[:160] + "…") if len(overall_fb) > 160 else overall_fb

        plat_cls    = "bay" if sub.get("platform") == "Bayut" else "dub"

        lang_cls    = "eng" if sub.get("language") == "English" else "ara"

        grade_short = get_grade(score).split(" — ")[-1]

        editor_chip = f'<span class="meta-chip">👤 {sub["editor_name"]}</span>' if sub.get("editor_name") else ""

        st.markdown(f"""

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

    if "Submit" in page: page_submit()

    else:                page_dashboard()



if __name__ == "__main__":

    main()
