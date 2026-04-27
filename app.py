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

st.set_page_config(page_title="Content QA System", page_icon="Q", layout="centered",
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
    st.markdown("""<style>
    /* global */
    [data-testid="stAppViewContainer"] > .main {background:#f0f2f9}
    [data-testid="stSidebar"]{background:#ffffff!important;border-right:1px solid #e8eaf0}
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p{font-size:13px;color:#374151}
    [data-testid="stSidebar"] h3{font-size:11px!important;font-weight:600!important;color:#9ca3af!important;text-transform:uppercase!important;letter-spacing:.08em!important}
    section[data-testid="stSidebar"] > div{padding-top:1rem}

    /* hero */
    .qa-hero{background:linear-gradient(135deg,#4f46e5 0%,#7c3aed 55%,#a855f7 100%);border-radius:16px;padding:28px 30px;margin-bottom:1.4rem;color:#fff;display:flex;align-items:flex-start;justify-content:space-between}
    .qa-hero-badge{background:rgba(255,255,255,.2);border-radius:20px;padding:4px 12px;font-size:11px;font-weight:500;color:#fff;margin-bottom:10px;display:inline-block}
    .qa-hero h1{font-size:24px;font-weight:700;color:#fff;margin:5px 0 8px;line-height:1.2}
    .qa-hero p{font-size:12px;color:rgba(255,255,255,.82);line-height:1.6;margin:0;max-width:340px}
    .qa-hero-icon{width:52px;height:52px;background:rgba(255,255,255,.15);border-radius:14px;display:flex;align-items:center;justify-content:center;font-size:24px;flex-shrink:0}

    /* form card */
    .form-card{background:#fff;border-radius:14px;padding:22px 26px;border:1px solid #e8eaf0;margin-bottom:1rem}
    .form-card-header{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:18px;padding-bottom:14px;border-bottom:1px solid #f3f4f6}
    .form-card-title{font-size:15px;font-weight:600;color:#111827;margin-bottom:3px}
    .form-card-sub{font-size:12px;color:#9ca3af}
    .ready-badge{background:#dcfce7;color:#15803d;font-size:11px;font-weight:500;padding:3px 10px;border-radius:20px;display:flex;align-items:center;gap:4px;white-space:nowrap}
    .ready-dot{width:6px;height:6px;border-radius:50%;background:#16a34a;display:inline-block}

    /* platform pills */
    .plat-wrap{display:flex;align-items:center;gap:10px;margin-bottom:14px}
    .plat-lbl{font-size:13px;color:#374151;font-weight:500}
    .plat-pills{display:flex;background:#f3f4f6;border-radius:50px;padding:3px;gap:2px}
    .pill-bay-on{background:#10b981;color:#fff;border-radius:50px;padding:6px 20px;font-size:13px;font-weight:500;border:none;cursor:pointer}
    .pill-off{background:transparent;color:#9ca3af;border-radius:50px;padding:6px 20px;font-size:13px;font-weight:500;border:none;cursor:pointer}

    /* score */
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

    /* detect */
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

    /* comments */
    .cmt-card{background:#f0f2f9;border-left:3px solid #4f46e5;padding:9px 13px;margin-bottom:7px;border-radius:0 8px 8px 0;font-size:13px}
    .cmt-author{font-weight:600;color:#4f46e5}
    .cmt-deduct{font-size:11px;color:#dc2626;font-weight:500;margin-top:3px}
    .cat-ref{font-size:10px;font-weight:500;padding:2px 7px;border-radius:20px;background:#ede9fe;color:#4f46e5;margin-left:6px}

    /* suggestions */
    .suggest-item{display:flex;gap:11px;align-items:flex-start;padding:9px 0;border-bottom:1px solid #f3f4f6;font-size:13px}
    .suggest-item:last-child{border:none;padding-bottom:0}
    .suggest-num{width:22px;height:22px;border-radius:50%;background:#ede9fe;color:#4f46e5;font-size:10px;font-weight:600;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:1px}
    .suggest-cat{font-size:11px;color:#9ca3af;margin-top:2px}

    /* tags */
    .tag-str{background:#d1fae5;color:#065f46;padding:3px 11px;border-radius:20px;font-size:12px;font-weight:500;display:inline-block;margin:2px}
    .tag-imp{background:#fef3c7;color:#92400e;padding:3px 11px;border-radius:20px;font-size:12px;font-weight:500;display:inline-block;margin:2px}
    .bdg{font-size:11px;font-weight:500;padding:3px 10px;border-radius:20px}
    .bdg-bay{background:#d1fae5;color:#065f46}
    .bdg-dub{background:#fee2e2;color:#b91c1c}
    .no-cmt-notice{background:#f0f2f9;border:1px solid #e0e4f0;border-radius:8px;padding:11px 15px;font-size:13px;color:#6b7280;margin-bottom:10px}

    /* streamlit overrides */
    div[data-testid="stProgress"]>div{background:#7c3aed!important}
    [data-testid="stFormSubmitButton"] button{
        background:linear-gradient(135deg,#3730a3,#6d28d9)!important;
        color:#fff!important;border:none!important;border-radius:10px!important;
        font-size:14px!important;font-weight:600!important;padding:13px!important;width:100%!important}
    [data-testid="stAlert"]{border-radius:10px!important;border:none!important;background:#ede9fe!important}
    [data-testid="stAlert"] p{color:#5b21b6!important;font-size:13px!important}
    [data-testid="stFileUploadDropzone"]{background:#ffffff!important;border:1.5px dashed #d1d5db!important;border-radius:12px!important;min-height:130px!important;padding:20px!important}
    [data-testid="stFileUploadDropzone"]:hover{border-color:#a78bfa!important;background:#faf5ff!important}
    [data-testid="stFileUploadDropzone"] > div{background:transparent!important;border:none!important}
    [data-testid="stFileUploaderDropzoneInstructions"] > div{background:transparent!important}
    [data-testid="stFileUploaderDropzoneInstructions"] svg{background:#7c3aed!important;border-radius:12px!important;padding:10px!important;color:#fff!important;width:40px!important;height:40px!important}
    /* platform pill toggle */
    [data-testid="stRadio"] > div{display:flex!important;flex-direction:row!important;background:#f3f4f6!important;border-radius:50px!important;padding:3px!important;gap:2px!important;width:fit-content!important}
    [data-testid="stRadio"] label{border-radius:50px!important;padding:7px 22px!important;font-size:13px!important;font-weight:500!important;cursor:pointer!important;margin:0!important;min-height:unset!important;display:flex!important;align-items:center!important;justify-content:center!important}
    [data-testid="stRadio"] label:has(input:checked){background:#10b981!important;color:#fff!important}
    [data-testid="stRadio"] label:not(:has(input:checked)){background:transparent!important;color:#6b7280!important}
    [data-testid="stRadio"] label:not(:has(input:checked)):hover{background:rgba(0,0,0,.04)!important;color:#374151!important}
    [data-testid="stRadio"] input[type="radio"]{position:fixed!important;opacity:0!important;width:0!important;height:0!important;pointer-events:none!important}
    [data-testid="stRadio"] p{font-size:13px!important;font-weight:500!important;margin:0!important;line-height:1!important;color:inherit!important}
    [data-testid="stTextInput"] input{border-radius:10px!important;border:1.5px solid #e5e7eb!important;padding:10px 14px!important;font-size:14px!important;background:#fff!important}
    [data-testid="stTextInput"] input:focus{border-color:#a78bfa!important;box-shadow:0 0 0 3px rgba(167,139,250,.15)!important}
    [data-testid="stSelectbox"] > div > div{border-radius:10px!important;border:1.5px solid #e5e7eb!important;background:#fff!important;padding:2px 4px!important}
    [data-testid="stSelectbox"] > div > div:focus-within{border-color:#a78bfa!important}
    [data-testid="stForm"]{border:none!important;padding:0!important;background:transparent!important}
    div[class*="stTextInput"] > label{font-size:13px!important;font-weight:500!important;color:#374151!important}
    div[class*="stSelectbox"] > label{font-size:13px!important;font-weight:500!important;color:#374151!important}
    </style>""", unsafe_allow_html=True)


# ── Groq AI ────────────────────────────────────────────────────────────────
def call_ai(prompt):
    if not GROQ_OK:
        raise Exception("groq not installed — check requirements.txt")
    client = Groq(api_key=st.secrets["GROQ_API_KEY"])
    for model in ["llama-3.1-8b-instant", "llama3-8b-8192", "gemma2-9b-it"]:
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[{"role":"user","content":prompt}],
                temperature=0.3, max_tokens=2000)
            return resp.choices[0].message.content.strip()
        except Exception as e:
            if any(x in str(e).lower() for x in ["model","not found","decommission"]):
                continue
            raise e
    raise Exception("All Groq models failed. Check GROQ_API_KEY in Secrets.")


# ── file parsers ───────────────────────────────────────────────────────────
def extract_docx(raw):
    if not DOCX_OK:
        return {"text":"","headings":[],"links":[],"comments":[],"word_count":0,"error":"python-docx not installed"}
    import zipfile
    from lxml import etree as _etree

    doc = Document(BytesIO(raw))
    text, headings, links = [], [], []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t: continue
        text.append(t)
        s = p.style.name
        if   s.startswith("Heading 1"): headings.append({"level":"H1","text":t})
        elif s.startswith("Heading 2"): headings.append({"level":"H2","text":t})
        elif s.startswith("Heading 3"): headings.append({"level":"H3","text":t})
    for rel in doc.part.rels.values():
        if "hyperlink" in rel.reltype:
            links.append(rel._target)

    # Extract comments directly from the zip — works reliably across all Word versions
    # Only keep the FIRST comment per thread (the editor's comment, not writer replies)
    comments = []
    try:
        WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        with zipfile.ZipFile(BytesIO(raw)) as z:
            if "word/comments.xml" in z.namelist():
                root = _etree.fromstring(z.read("word/comments.xml"))
                all_c = root.findall(f".//{{{WNS}}}comment")

                # Separate parent comments from replies
                # Replies have w:paraIdParent attribute or are in w:commentExtended
                # Simpler: the first author to comment on a thread is the editor
                # Collect all comment IDs that are replies
                reply_ids = set()
                # Check commentsExtended if available
                if "word/commentsExtended.xml" in z.namelist():
                    ext_root = _etree.fromstring(z.read("word/commentsExtended.xml"))
                    W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
                    for ext in ext_root.findall(f".//{{{W15}}}commentEx"):
                        parent_id = ext.get(f"{{{W15}}}paraIdParent")
                        if parent_id:
                            # This comment is a reply — get its ID
                            cid = ext.get(f"{{{W15}}}paraId")
                            reply_ids.add(cid)

                # Build para_id -> comment_id map from document body
                body_xml = z.read("word/document.xml")
                body_root = _etree.fromstring(body_xml)
                # Map paraId to comment IDs from commentsExtended
                para_to_cid = {}
                if "word/commentsExtended.xml" in z.namelist():
                    ext_root2 = _etree.fromstring(z.read("word/commentsExtended.xml"))
                    W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
                    for ext in ext_root2.findall(f".//{{{W15}}}commentEx"):
                        para_id  = ext.get(f"{{{W15}}}paraId","")
                        par_par  = ext.get(f"{{{W15}}}paraIdParent","")
                        cid      = ext.get(f"{{{W15}}}id","")
                        if par_par:
                            reply_ids.add(para_id)

                # Get all comment IDs that are replies by checking paraIdParent
                reply_comment_ids = set()
                if "word/commentsExtended.xml" in z.namelist():
                    ext_root3 = _etree.fromstring(z.read("word/commentsExtended.xml"))
                    W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
                    for ext in ext_root3.findall(f".//{{{W15}}}commentEx"):
                        has_parent = ext.get(f"{{{W15}}}paraIdParent")
                        if has_parent:
                            cid = ext.get(f"{{{W15}}}id","")
                            if cid: reply_comment_ids.add(cid)

                # Add comments — skip replies
                for c in all_c:
                    cid    = c.get(f"{{{WNS}}}id","")
                    author = c.get(f"{{{WNS}}}author","Editor")
                    body   = " ".join(c.itertext()).strip()
                    if body and cid not in reply_comment_ids:
                        comments.append({"author":author,"text":body})

    except Exception:
        pass

    full = "\n".join(text)
    return {"text":full,"headings":headings,"links":links,"comments":comments,"word_count":len(full.split()),"error":""}

def extract_pdf(raw):
    if not PDF_OK:
        return {"text":"","headings":[],"links":[],"comments":[],"word_count":0,"error":"pdfplumber not installed"}
    parts, links = [], []
    with pdfplumber.open(BytesIO(raw)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t: parts.append(t)
            for a in (page.annots or []):
                u = a.get("uri")
                if u: links.append(u)
    full = "\n".join(parts)
    return {"text":full,"headings":[],"links":links,"comments":[],"word_count":len(full.split()),"error":""}

def extract_txt(raw):
    full = raw.decode("utf-8", errors="ignore")
    return {"text":full,"headings":[],"links":re.findall(r'https?://\S+',full),"comments":[],"word_count":len(full.split()),"error":""}

def parse_file(f):
    raw  = f.getvalue()
    name = f.name.lower()
    if   name.endswith(".docx"): return extract_docx(raw)
    elif name.endswith(".pdf"):  return extract_pdf(raw)
    else:                        return extract_txt(raw)


# ── scoring ────────────────────────────────────────────────────────────────
def classify_comment(text):
    """Classify editor comment into type using keywords. Returns (type, deduction)."""
    low = text.lower()
    # Data accuracy — factual errors, wrong numbers, wrong info
    data_accuracy_kw = [
        "wrong", "incorrect", "not correct", "inaccurate", "error",
        "should be", "it is", "it's", "the source", "in the source",
        "copied", "from google", "from maps", "url goes", "link goes",
        "apartments", "no apartments", "mins away", "minutes away",
        "under construction", "off-plan", "data", "fact",
    ]
    # Missing info — content gaps
    missing_info_kw = [
        "missing", "add", "please add", "include", "mention", "not mentioned",
        "should mention", "we need", "please mention", "go through",
        "available", "please write", "write the branch", "notable projects",
        "specific", "more details", "lacks", "header", "section",
    ]
    # Grammar / rephrasing
    grammar_kw = [
        "grammar", "rephrase", "rewrite", "word", "sentence", "phrasing",
        "general", "too general", "vague", "unclear", "confusing",
        "brand voice", "tone", "style", "read", "sounds",
    ]

    for kw in data_accuracy_kw:
        if kw in low:
            return "Data accuracy", 1.5
    for kw in missing_info_kw:
        if kw in low:
            return "Missing info", 1.5
    for kw in grammar_kw:
        if kw in low:
            return "Grammar / rephrasing", 1.0
    # Default — treat as grammar/rephrasing
    return "Grammar / rephrasing", 1.0


def apply_deductions(base_score, comments, plag_pct, ai_pct):
    """
    Scoring rules:
    - Base: 100
    - Data accuracy comment: -1.5 pts each
    - Missing info comment:   -1.5 pts each
    - Grammar/rephrasing:     -1.0 pt each
    - Plagiarism: -5 pts per 20% (40%=-10, 60%=-15 etc.)
    - AI content: -5 pts per 20% (same scale)
    """
    base = 100

    # Classify each comment and sum deductions
    classified = []
    comment_deduction = 0.0
    for c in comments:
        ctype, pts = classify_comment(c["text"])
        classified.append({"author":c["author"],"text":c["text"],"type":ctype,"deduction":pts})
        comment_deduction += pts

    # Plagiarism: -5 per 20% bracket
    plag_brackets  = int(plag_pct // 20)
    plag_deduction = plag_brackets * 5

    # AI: -5 per 20% bracket
    ai_brackets  = int(ai_pct // 20)
    ai_deduction = ai_brackets * 5

    final = max(0, round(base - comment_deduction - plag_deduction - ai_deduction, 1))

    return final, {
        "base_score":        base,
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


# ── QA — comment-driven only ───────────────────────────────────────────────
def run_qa(title, content, writer, ctype, lang, platform, headings, links, comments):
    h_txt = "\n".join(f"  [{h['level']}] {h['text']}" for h in headings) or "  None"
    l_txt = "\n".join(f"  - {l}" for l in links[:8])                      or "  None"

    if not comments:
        scores = {cat:{"score":mx,"feedback":"No editor comments. Full marks awarded.","comment_refs":[]}
                  for cat,mx in CAT_MAX.items()}
        return {"scores":scores,"total":sum(CAT_MAX.values()),
                "overall_feedback":"No editor comments found. All categories awarded full marks.",
                "key_strengths":[],"areas_for_improvement":[],"suggestions":[]}

    c_txt = "\n".join(f"  Comment {i+1} [{c['author']}]: {c['text']}" for i,c in enumerate(comments))

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
    clean = re.sub(r"```json|```","",raw).strip()
    m     = re.search(r'\{.*\}', clean, re.DOTALL)
    if m: clean = m.group(0)
    return json.loads(clean)


# ── plagiarism ─────────────────────────────────────────────────────────────
def check_plagiarism(text, links):
    """
    Groq-powered plagiarism detection — free and unlimited.
    Asks Llama 3 to read the article and identify sentences
    that appear copied from developer brochures or websites,
    then returns a percentage and the flagged sentences.
    """
    flagged_sources = [l for l in links if any(d in l for d in KNOWN_DOMAINS)]

    prompt = f"""You are a plagiarism detection expert for UAE real estate content.

Read the article below and identify sentences that appear to be copied or closely lifted from:
- Developer brochures (Emaar, Aldar, Damac, Nakheel, Meraas, Sobha etc.)
- Developer websites
- Property listing descriptions
- AI-generated filler text (e.g. "Certainly. Here are the amenities...")

For each copied sentence, include the exact sentence from the article.

Return ONLY valid JSON:
{{
  "plagiarism_percentage": <0-100 integer, estimate what % of the article is copied>,
  "flagged_sentences": [
    "<exact copied sentence from article>",
    "<exact copied sentence from article>"
  ],
  "assessment": "<1 sentence summary of the plagiarism situation>"
}}

Rules:
- Only flag sentences that are clearly copied developer/brochure language
- Do NOT flag factual data (prices, sq ft, dates, distances) — those are always original
- Do NOT flag simple descriptive sentences that any writer might write
- Flag sentences with marketing phrases like "embody contemporary elegance", "seamlessly integrates", "wellness-oriented", "dynamic enclave", "fabric of daily life", "certainly. here are" etc.
- Return between 0 and 15 flagged sentences maximum

ARTICLE:
{text[:5000]}"""

    try:
        raw   = call_ai(prompt)
        clean = re.sub(r"```json|```", "", raw).strip()
        m     = re.search(r'{.*}', clean, re.DOTALL)
        if m: clean = m.group(0)
        result = json.loads(clean)

        pct       = min(int(result.get("plagiarism_percentage", 0)), 100)
        sentences = result.get("flagged_sentences", [])[:15]

        return {
            "percentage":       pct,
            "flagged_sources":  flagged_sources,
            "flagged_sentences":sentences,
            "hits":             len(sentences),
            "source":           "Groq",
            "assessment":       result.get("assessment",""),
            "status":           "danger" if pct>20 else "warn" if pct>10 else "safe",
        }
    except Exception:
        # Heuristic fallback if Groq call fails
        text_lower = text.lower()
        hits  = sum(1 for p in BROCHURE_PHRASES if p in text_lower)
        total = len(BROCHURE_PHRASES)
        base  = min(int((math.sqrt(hits)/math.sqrt(max(total,1)))*60), 60)
        bonus = min(len(flagged_sources)*4, 15)
        pct   = min(base+bonus, 100)
        return {"percentage":pct,"flagged_sources":flagged_sources,
                "flagged_sentences":[],"hits":hits,"source":"heuristic",
                "status":"danger" if pct>20 else "warn" if pct>10 else "safe"}

def get_plag_snippets(text, links, plag_result=None):
    """Return (sources, flagged_sentences). Uses Groq result if available."""
    if plag_result and plag_result.get("flagged_sentences"):
        return plag_result.get("flagged_sources",[]), plag_result.get("flagged_sentences",[])
    # Heuristic fallback
    flagged_sources = [l for l in links if any(d in l for d in KNOWN_DOMAINS)]
    sentences = re.split(r'(?<=[.!?])\s+', text)
    flagged, seen = [], set()
    for sent in sentences:
        stripped = sent.strip()
        low = stripped.lower()
        if len(low) < 35 or stripped in seen: continue
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
        def mark(m): return ('<mark style="background:#fecaca;border-radius:3px;padding:0 2px;font-weight:500;color:#7f1d1d">'+m.group(0)+'</mark>')
        result = pat.sub(mark, result)
    return result


# ── AI detection ───────────────────────────────────────────────────────────
def check_ai(text):
    """Real AI detection via GPTZero API. Falls back to heuristic if no key."""
    api_key = st.secrets.get("GPTZERO_API_KEY","")
    if api_key and len(text.strip()) > 50:
        try:
            payload = json.dumps({"document":text[:10000],"version":"2025-01-09"}).encode("utf-8")
            req     = urllib.request.Request(
                "https://api.gptzero.me/v2/predict/text",
                data=payload,
                headers={"Content-Type":"application/json","Accept":"application/json","x-api-key":api_key},
                method="POST")
            with urllib.request.urlopen(req, timeout=20) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            doc      = data.get("documents",[{}])[0]
            ai_pct   = int(round(doc.get("completely_generated_prob",0)*100))
            sents    = doc.get("sentences",[])
            ai_sents = [s.get("sentence","") for s in sents
                        if s.get("generated_prob",0)>0.5 and len(s.get("sentence",""))>30][:5]
            return {"ai_pct":ai_pct,"human_pct":100-ai_pct,
                    "status":"danger" if ai_pct>20 else "warn" if ai_pct>10 else "safe",
                    "source":"GPTZero","ai_sentences":ai_sents}
        except Exception:
            pass

    # Heuristic fallback
    hits = sum(1 for p in AI_PHRASES if p in text.lower())
    pct  = min(hits*5, 60)
    snippets = []
    for sent in re.split(r'(?<=[.!?])\s+', text):
        low = sent.strip().lower()
        if len(low)<35: continue
        if any(p in low for p in AI_PHRASES):
            snippets.append(sent.strip()[:200])
        if len(snippets)>=5: break
    return {"ai_pct":pct,"human_pct":100-pct,
            "status":"danger" if pct>20 else "warn" if pct>10 else "safe",
            "source":"heuristic","ai_sentences":snippets}

def highlight_ai(sentence):
    result = sentence
    for phrase in AI_PHRASES:
        pat = re.compile(re.escape(phrase), re.IGNORECASE)
        def mark(m): return ('<mark style="background:#fef3c7;border-radius:3px;padding:0 2px;font-weight:500;color:#78350f">'+m.group(0)+'</mark>')
        result = pat.sub(mark, result)
    return result


# ── Google Sheets ──────────────────────────────────────────────────────────
def log_to_sheets(row):
    if not SHEETS_OK: return
    try:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/spreadsheets"])
        gc    = gspread.authorize(creds)
        sheet = gc.open(st.secrets.get("SHEET_NAME","Bayut QA Submissions")).sheet1
        if not sheet.row_values(1):
            sheet.append_row(["Date","Platform","Writer","Title","Type","Language",
                              "Base Score","Comments","Plagiarism Ded","AI Ded",
                              "Final Score","Plagiarism%","AI%","Recommendation",
                              "Editor Decision","Notes"])
        d = row.get("deductions",{})
        sheet.append_row([row.get("date"),row.get("platform"),row.get("writer"),
                          row.get("title"),row.get("content_type"),row.get("language"),
                          d.get("base_score",0),d.get("comment_deduction",0),
                          d.get("plag_deduction",0),d.get("ai_deduction",0),
                          d.get("final_score",0),row.get("plagiarism_pct",0),
                          row.get("ai_pct",0),row.get("recommendation",""),
                          row.get("editor_decision",""),row.get("editor_notes","")])
    except Exception as e:
        st.warning(f"Google Sheets log failed: {e}")


# ── sidebar ────────────────────────────────────────────────────────────────
def sidebar():
    with st.sidebar:
        st.markdown("""
<div style="display:flex;align-items:center;gap:10px;padding:4px 0 16px">
  <div style="width:36px;height:36px;border-radius:50%;background:linear-gradient(135deg,#4f46e5,#7c3aed);display:flex;align-items:center;justify-content:center;color:#fff;font-size:13px;font-weight:700;flex-shrink:0">QA</div>
  <div><div style="font-size:13px;font-weight:600;color:#111827;line-height:1.2">Content QA</div><div style="font-size:11px;color:#9ca3af">Editorial review</div></div>
</div>""", unsafe_allow_html=True)
        st.divider()
        st.markdown('<p style="font-size:10px;font-weight:600;color:#9ca3af;letter-spacing:.08em;text-transform:uppercase;margin-bottom:4px">Navigation</p>', unsafe_allow_html=True)
        page = st.radio("Go to",["Submit article","Dashboard"],label_visibility="collapsed")
        st.divider()
        st.markdown('<p style="font-size:10px;font-weight:600;color:#9ca3af;letter-spacing:.08em;text-transform:uppercase;margin-bottom:4px">Deduction rules</p>', unsafe_allow_html=True)
        for label, pts in [("Data accuracy","-1.5"),("Missing info","-1.5"),("Grammar","-1"),("Plagiarism / 20%","-5"),("AI content / 20%","-5")]:
            st.markdown(f'<div style="display:flex;justify-content:space-between;align-items:center;padding:7px 0;border-bottom:1px solid #f3f4f6"><span style="font-size:13px;color:#374151">{label}</span><span style="background:#fee2e2;color:#dc2626;font-size:11px;font-weight:600;padding:2px 10px;border-radius:20px">{pts}</span></div>', unsafe_allow_html=True)
        st.markdown("""
<div style="margin-top:16px;background:#f0f2f9;border-radius:10px;padding:12px 14px">
  <div style="font-size:12px;font-weight:600;color:#374151;margin-bottom:4px">Enable AI detection</div>
  <div style="font-size:11px;color:#6b7280;line-height:1.5">Add your AI-content detection key to unlock real plagiarism & AI scoring.</div>
</div>""", unsafe_allow_html=True)
        return page


# ── submit page ────────────────────────────────────────────────────────────
def page_submit():
    inject_css()
    st.markdown(
        '<div class="qa-hero">'
        '<div>'
        '<div class="qa-hero-badge">Editorial QA Engine</div>'
        '<h1>Content QA System</h1>'
        '<p>Submit articles for automated review. We score editor comments,<br>detect plagiarism, and flag AI-generated content.</p>'
        '</div>'
        '<div class="qa-hero-icon">&#x1F4CB;</div>'
        '</div>', unsafe_allow_html=True)

    st.markdown(
        '<div class="form-card-header">'
        '<div><div class="form-card-title">New submission</div>'
        '<div class="form-card-sub">Fill in the details below and upload the article file.</div></div>'
        '<div class="ready-badge"><span class="ready-dot"></span> Ready</div>'
        '</div>', unsafe_allow_html=True)

    with st.form("qa_form"):
        c1,c2 = st.columns(2)
        writer = c1.text_input("Writer name",placeholder="e.g. Sarah Ahmed")
        title  = c2.text_input("Article title",placeholder="e.g. Everything About Mortgages")
        c3,c4  = st.columns(2)
        ctype  = c3.selectbox("Content type",CONTENT_TYPES)
        lang   = c4.selectbox("Language",LANGUAGES)
        st.markdown("**Platform**")
        platform = st.radio("Platform",PLATFORMS,horizontal=True,label_visibility="collapsed")
        bay = platform=="Bayut"
        upload = st.file_uploader("Upload article file",type=["docx","pdf","txt"],
                    help=".docx recommended — headings, links and editor comments are extracted automatically")
        go = st.form_submit_button("Run full evaluation",use_container_width=True,type="primary")

    if not go:
        st.info("Upload a .docx file. Category scores are based entirely on editor comments. If no comments exist, full marks are awarded.")
        return
    if not writer or not title or not upload:
        st.error("Please fill in writer name, title and upload a file.")
        return

    with st.spinner("Reading file..."):
        parsed = parse_file(upload)

    if not parsed["text"] or len(parsed["text"])<30:
        st.error(f"Could not read text from file. {parsed.get('error','')}")
        return

    # Auto-detect editor comments vs writer replies
    # Writer replies typically contain: "fixed", "done", "added", "removed",
    # "replaced", "updated", "@", "changed" — editor comments are questions/issues
    writer_reply_keywords = [
        "fixed", "done", "added", "removed", "replaced", "updated",
        "changed", "edited", "deleted", "corrected", "revised",
        "@", "noted", "ok ", "okay", "sure", "will do",
    ]

    def is_writer_reply(comment_text):
        low = comment_text.lower().strip()
        # Pure @mentions are always writer responses
        if low.startswith("@"): return True
        # Short replies like "Fixed." are writer responses
        if len(low) < 30:
            for kw in writer_reply_keywords:
                if low.startswith(kw) or kw in low[:20]:
                    return True
        return False

    all_comments    = parsed["comments"]
    editor_comments = [c for c in all_comments if not is_writer_reply(c["text"])]
    parsed["comments"]     = editor_comments
    parsed["all_comments"] = all_comments

    with st.expander(f"Extracted — {len(parsed['headings'])} headings, {len(parsed['links'])} links, {len(parsed['comments'])} editor comments"):
        col_h,col_l,col_c = st.columns(3)
        with col_h:
            st.markdown("**Headings**")
            for h in parsed["headings"]: st.markdown(f"`{h['level']}` {h['text']}")
            if not parsed["headings"]: st.caption("None detected")
        with col_l:
            st.markdown("**Links**")
            for l in parsed["links"][:6]: st.markdown(f"- {l}")
            if not parsed["links"]: st.caption("None detected")
        with col_c:
            st.markdown("**Editor comments**")
            for idx,c in enumerate(parsed["comments"],1):
                st.markdown(
                    f'<div class="cmt-card"><span class="cmt-author">Comment {idx} — {c["author"]}</span><br>'
                    f'{c["text"]}<div class="cmt-deduct">1 point deducted</div></div>',
                    unsafe_allow_html=True)
            if not parsed["comments"]: st.caption("No comments found")

    prog = st.progress(0,text="Starting...")

    try:
        prog.progress(15,text="Scoring categories from editor comments...")
        qa = run_qa(title,parsed["text"],writer,ctype,lang,platform,
                    parsed["headings"],parsed["links"],parsed["comments"])
    except Exception as e:
        st.error(f"AI evaluation failed: {e}")
        st.info("Make sure GROQ_API_KEY is set in Streamlit Secrets.")
        return

    prog.progress(50,text="Running plagiarism check...")
    plag = check_plagiarism(parsed["text"],parsed["links"])
    plag_sources,plag_snippets = get_plag_snippets(parsed["text"],parsed["links"],plag)

    prog.progress(72,text="Running AI detection...")
    ai = check_ai(parsed["text"])

    prog.progress(95,text="Calculating final score...")
    base_score = qa.get("total",0)
    final_score,deductions = apply_deductions(base_score,parsed["comments"],plag["percentage"],ai["ai_pct"])
    recommendation = get_recommendation(final_score)

    prog.progress(100,text="Done!")
    prog.empty()

    sub = {
        "date":           datetime.now().strftime("%d %b %Y %H:%M"),
        "platform":       platform,"writer":writer,"title":title,
        "content_type":   ctype,"language":lang,"word_count":parsed["word_count"],
        "headings":       parsed["headings"],"links":parsed["links"],"comments":parsed["comments"],
        "qa":             qa,"plagiarism":plag,"plag_snippets":plag_snippets,"plag_sources":plag_sources,
        "ai_detection":   ai,"deductions":deductions,"qa_score":final_score,
        "plagiarism_pct": plag["percentage"],"ai_pct":ai["ai_pct"],
        "recommendation": recommendation,"editor_decision":"","editor_notes":"",
    }
    if "submissions" not in st.session_state:
        st.session_state.submissions = []
    st.session_state.submissions.append(sub)
    render_report(sub)


# ── report ─────────────────────────────────────────────────────────────────
def render_report(sub):
    inject_css()
    qa    = sub["qa"]
    plag  = sub["plagiarism"]
    ai    = sub["ai_detection"]
    ded   = sub["deductions"]
    score = sub["qa_score"]
    grade = get_grade(score)
    rec   = sub["recommendation"]
    plag_snippets = sub.get("plag_snippets",[])
    plag_sources  = sub.get("plag_sources",[])

    st.divider()
    bdg_class = "bdg-bay" if sub["platform"]=="Bayut" else "bdg-dub"
    st.markdown(
        f"**{sub['writer']}** &nbsp; <span class='bdg {bdg_class}'>{sub['platform']}</span> &nbsp; "
        f"`{sub['content_type']}` &nbsp; `{sub['language']}` &nbsp; `{sub['word_count']} words` &nbsp; `{sub['date']}`",
        unsafe_allow_html=True)
    st.markdown("")

    rec_labels = {"approve":("Approve","#d1fae5","#065f46"),
                  "revise":("Request revision","#fef3c7","#92400e"),
                  "reject":("Reject","#fee2e2","#991b1b")}
    rl,rbg,rtc = rec_labels.get(rec,rec_labels["revise"])

    def brow(cls,label,val):
        return f'<div class="{cls}"><span>{label}</span><span>{val}</span></div>'

    # Build comment breakdown by type
    classified     = ded.get("classified", [])
    data_acc_cmts  = [c for c in classified if c["type"]=="Data accuracy"]
    missing_cmts   = [c for c in classified if c["type"]=="Missing info"]
    grammar_cmts   = [c for c in classified if c["type"]=="Grammar / rephrasing"]

    comment_rows = ""
    if data_acc_cmts:
        comment_rows += brow("ded-row",
            f'Data accuracy ({len(data_acc_cmts)} comments × 1.5 pts)',
            f'- {round(len(data_acc_cmts)*1.5,1)} pts')
    if missing_cmts:
        comment_rows += brow("ded-row",
            f'Missing info ({len(missing_cmts)} comments × 1.5 pts)',
            f'- {round(len(missing_cmts)*1.5,1)} pts')
    if grammar_cmts:
        comment_rows += brow("ded-row",
            f'Grammar / rephrasing ({len(grammar_cmts)} comments × 1 pt)',
            f'- {len(grammar_cmts)} pts')
    if not classified:
        comment_rows = brow("ok-row","Editor comments","no deduction")

    # Plagiarism brackets
    plag_b  = ded.get("plag_brackets",0)
    plag_row = (brow("ded-row",
        f'Plagiarism {ded["plag_pct"]}% ({plag_b} × 20% bracket × 5 pts)',
        f'- {ded["plag_deduction"]} pts')
        if ded["plag_deduction"]>0
        else brow("ok-row",f'Plagiarism {ded["plag_pct"]}% — under 20%',"no deduction"))

    # AI brackets
    ai_b  = ded.get("ai_brackets",0)
    ai_row = (brow("ded-row",
        f'AI content {ded["ai_pct"]}% ({ai_b} × 20% bracket × 5 pts)',
        f'- {ded["ai_deduction"]} pts')
        if ded["ai_deduction"]>0
        else brow("ok-row",f'AI content {ded["ai_pct"]}% — under 20%',"no deduction"))

    bd = (
        brow("base-row","Base score",f'{ded["base_score"]} / 100') +
        comment_rows + plag_row + ai_row +
        brow("total-row","Final score",f"{score} / 100")
    )

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
    pc1,pc2 = st.columns(2)

    with pc1:
        pp   = plag["percentage"]
        over = pp > 20
        col  = "#dc2626" if over else "#059669"
        thresh = (f'<span class="detect-thresh" style="background:#fee2e2;color:#991b1b">{pp}% — over 20% threshold — 5 points deducted</span>'
                  if over else
                  f'<span class="detect-thresh" style="background:#d1fae5;color:#065f46">{pp}% — under 20% threshold — no deduction</span>')
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
            f'<div class="detect-title">Plagiarism check <span style="font-size:10px;font-weight:400;color:#888;margin-left:6px">via {plag.get("source","heuristic")}</span></div>'
            f'<div class="detect-bar"><div class="detect-bar-f" style="width:{min(pp,100)}%;background:{col}"></div></div>'
            f'{thresh}<div class="detect-note">{"Rewrite all flagged sections completely." if over else "Content is within acceptable range."}</div>'
            f'{snip_html}</div>',unsafe_allow_html=True)


    with pc2:
        ap      = ai["ai_pct"]
        ai_over = ap > 20
        a_col   = "#dc2626" if ai_over else "#059669"
        src_lbl = "GPTZero" if ai.get("source")=="GPTZero" else "heuristic"
        a_thresh = (f'<span class="detect-thresh" style="background:#fee2e2;color:#991b1b">{ap}% — over 20% threshold — 5 points deducted</span>'
                    if ai_over else
                    f'<span class="detect-thresh" style="background:#d1fae5;color:#065f46">{ap}% — under 20% threshold — no deduction</span>')
        ai_sents    = ai.get("ai_sentences",[])
        ai_snip_html = ""
        if ai_sents:
            ai_snip_html = '<div class="issue-block"><div class="issue-block-title">Flagged sentences</div>'
            for s in ai_sents[:5]:
                ai_snip_html += f'<div class="issue-snippet">{highlight_ai(s)}</div>'
            ai_snip_html += '</div>'
        st.markdown(
            f'<div class="detect-card">'
            f'<div class="detect-title">AI detection <span style="font-size:10px;font-weight:400;color:#888;margin-left:6px">via {src_lbl}</span></div>'
            f'<div class="detect-bar"><div class="detect-bar-f" style="width:{min(ap,100)}%;background:{a_col}"></div></div>'
            f'{a_thresh}<div class="detect-note">{"High AI content detected." if ai_over else "Content appears mostly human-written."}</div>'
            f'<div class="detect-split">'
            f'<div class="detect-seg"><div class="detect-seg-n" style="color:#059669">{ai["human_pct"]}%</div><div class="detect-seg-l">Human</div></div>'
            f'<div class="detect-seg"><div class="detect-seg-n" style="color:{a_col}">{ap}%</div><div class="detect-seg-l">AI likely</div></div>'
            f'</div>{ai_snip_html}</div>',unsafe_allow_html=True)

    st.divider()
    st.markdown("#### Category scores")
    if not sub["comments"]:
        st.markdown('<div class="no-cmt-notice">No editor comments found. All categories awarded full marks. Add comments to the .docx file for a real evaluation.</div>',unsafe_allow_html=True)

    for cat,mx in CAT_MAX.items():
        data = qa["scores"].get(cat,{})
        s    = data.get("score",0)
        fb   = data.get("feedback","")
        refs = data.get("comment_refs",[])
        ref_html = " ".join(f'<span class="cat-ref">Comment {r}</span>' for r in refs)
        ca,cb = st.columns([4,1])
        ca.markdown(f"**{cat}**"+(f" &nbsp; {ref_html}" if ref_html else ""),unsafe_allow_html=True)
        ca.progress(s/mx)
        cb.markdown(f"**{s} / {mx}**")
        st.caption(fb)
        st.markdown("")

    st.divider()
    st.markdown("#### Document structure")
    sc1,sc2,sc3 = st.columns(3)
    sc1.metric("Headings",     len(sub["headings"]))
    sc2.metric("Total links",  len(sub["links"]))
    internal = [l for l in sub["links"] if sub["platform"].lower() in l.lower()]
    sc3.metric("Internal links",len(internal))

    if sub["comments"]:
        classified = ded.get("classified", [])
        cmap = {c["text"]: c for c in classified}
        type_colors = {
            "Data accuracy":      ("#fee2e2","#991b1b","- 1.5 pts"),
            "Missing info":       ("#fef3c7","#92400e","- 1.5 pts"),
            "Grammar / rephrasing":("#f0f4ff","#2D4A8A","- 1 pt"),
        }
        st.markdown(f"**Editor comments — {len(sub['comments'])} found**")
        for idx,c in enumerate(sub["comments"],1):
            info    = cmap.get(c["text"], {})
            ctype   = info.get("type","Grammar / rephrasing")
            bg,tc,pts = type_colors.get(ctype, ("#f5f6fa","#555","- 1 pt"))
            st.markdown(
                f'<div class="cmt-card">'
                f'<span class="cmt-author">Comment {idx} — {c["author"]}</span>'
                f'<span style="font-size:10px;font-weight:500;padding:1px 8px;border-radius:20px;'
                f'background:{bg};color:{tc};margin-left:8px">{ctype}</span><br>'
                f'{c["text"]}'
                f'<div class="cmt-deduct">{pts} deducted</div></div>',
                unsafe_allow_html=True)

    st.divider()
    col_s,col_i = st.columns(2)
    with col_s:
        st.markdown("#### Strengths")
        strengths = qa.get("key_strengths",[])
        for s in strengths: st.markdown(f'<span class="tag-str">{s}</span>',unsafe_allow_html=True)
        if not strengths: st.caption("Add editor comments to see strengths.")
    with col_i:
        st.markdown("#### Required improvements")
        for imp in qa.get("areas_for_improvement",[]): st.markdown(f'<span class="tag-imp">{imp}</span>',unsafe_allow_html=True)

    suggestions = qa.get("suggestions",[])
    if suggestions:
        st.divider()
        st.markdown("#### Suggestions to improve the article")
        st.caption("Specific actions to address each editor comment")
        for sug in suggestions:
            st.markdown(
                f'<div class="suggest-item"><div class="suggest-num">{sug.get("number","")}</div>'
                f'<div><div>{sug.get("action","")}</div>'
                f'<div class="suggest-cat">Addresses: {sug.get("category","")}</div></div></div>',
                unsafe_allow_html=True)

    st.divider()
    st.markdown("#### Editor decision")
    st.caption("The AI recommendation is a guide. You make the final call.")
    rec_idx = {"approve":0,"revise":1,"reject":2}
    decision = st.radio("Decision",["Approve","Request revision","Reject"],
                        index=rec_idx.get(rec,1),horizontal=True,
                        key=f"dec_{sub['title']}_{sub['date']}")
    notes = st.text_area("Notes for writer (required for revision and rejection)",height=90,
                         placeholder="Tell the writer exactly what to fix.",
                         key=f"notes_{sub['title']}_{sub['date']}")
    if st.button("Confirm decision",type="primary",use_container_width=True,
                 key=f"conf_{sub['title']}_{sub['date']}"):
        if decision in ("Request revision","Reject") and not notes.strip():
            st.error("Please add notes before confirming.")
        else:
            sub["editor_decision"] = decision
            sub["editor_notes"]    = notes
            log_to_sheets(sub)
            st.success(f"Decision saved: {decision}")
            if notes: st.info(f"Notes for {sub['writer']}: {notes}")

    st.caption(f"Content QA System — {sub['platform']} — Powered by Groq — {sub['date']}")


# ── dashboard ──────────────────────────────────────────────────────────────
def page_dashboard():
    inject_css()
    st.markdown(
        '<div class="qa-hero">'
        '<div>'
        '<div class="qa-hero-badge">Overview</div>'
        '<h1>Dashboard</h1>'
        '<p>All submissions across Bayut and Dubizzle</p>'
        '</div>'
        '<div class="qa-hero-icon">&#x1F4CA;</div>'
        '</div>', unsafe_allow_html=True)
    all_subs = st.session_state.get("submissions",[])
    if not all_subs:
        st.info("No submissions yet.")
        return

    approved = sum(1 for s in all_subs if s.get("editor_decision")=="Approve")
    revision = sum(1 for s in all_subs if s.get("editor_decision")=="Request revision")
    rejected = sum(1 for s in all_subs if s.get("editor_decision")=="Reject")
    pending  = sum(1 for s in all_subs if not s.get("editor_decision"))

    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric("Total",len(all_subs))
    m2.metric("Approved",approved)
    m3.metric("Revision",revision)
    m4.metric("Rejected",rejected)
    m5.metric("Pending",pending)
    st.divider()

    all_writers = sorted(set(s["writer"] for s in all_subs if s.get("writer")))
    f1,f2,f3,f4,f5 = st.columns(5)
    wf = f1.selectbox("Writer",      ["All"]+all_writers)
    pf = f2.selectbox("Platform",    ["All"]+PLATFORMS)
    tf = f3.selectbox("Content type",["All"]+CONTENT_TYPES)
    lf = f4.selectbox("Language",    ["All"]+LANGUAGES)
    sf = f5.selectbox("Status",      ["All","Pending","Approve","Request revision","Reject"])

    filtered = all_subs
    if wf!="All": filtered=[s for s in filtered if s.get("writer")==wf]
    if pf!="All": filtered=[s for s in filtered if s["platform"]==pf]
    if tf!="All": filtered=[s for s in filtered if s["content_type"]==tf]
    if lf!="All": filtered=[s for s in filtered if s["language"]==lf]
    if sf!="All":
        if sf=="Pending": filtered=[s for s in filtered if not s.get("editor_decision")]
        else:             filtered=[s for s in filtered if s.get("editor_decision")==sf]

    st.markdown(f"**{len(filtered)} submissions**")
    for sub in reversed(filtered):
        dec   = sub.get("editor_decision") or "Pending"
        flags = ("  High plagiarism" if sub.get("plagiarism_pct",0)>20 else "") + \
                ("  High AI" if sub.get("ai_pct",0)>20 else "")
        label = (f"{sub['writer']} ({sub['platform']}) — "
                 f"{sub['title'][:45]}{'...' if len(sub['title'])>45 else ''} "
                 f"| Score: {sub['qa_score']} | {dec}{flags} | {sub['date']}")
        with st.expander(label):
            render_report(sub)


# ── main ───────────────────────────────────────────────────────────────────
def main():
    if "submissions" not in st.session_state:
        st.session_state.submissions = []
    page = sidebar()
    if "Submit" in page: page_submit()
    else:                page_dashboard()

if __name__ == "__main__":
    main()
