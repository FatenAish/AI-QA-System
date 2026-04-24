import streamlit as st
import json
import re
from datetime import datetime
from io import BytesIO

# ── optional imports (graceful fallback if not installed) ──────────────────
try:
    from docx import Document
    from docx.oxml.ns import qn
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

try:
    import pdfplumber
    PDF_OK = True
except ImportError:
    PDF_OK = False

try:
    import google.generativeai as genai
    GENAI_OK = True
except ImportError:
    GENAI_OK = False

try:
    import gspread
    from google.oauth2.service_account import Credentials
    SHEETS_OK = True
except ImportError:
    SHEETS_OK = False

# ── page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Content QA System | Bayut & Dubizzle",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── brand config ───────────────────────────────────────────────────────────
BRANDS = {
    "Bayut": {
        "primary": "#e2231a",
        "light": "#fff0ef",
        "logo": "🏠",
        "url": "bayut.com",
    },
    "Dubizzle": {
        "primary": "#00a699",
        "light": "#e0f5f4",
        "logo": "🏢",
        "url": "dubizzle.com",
    },
}

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
    (90, "A — Excellent",  "🟢"),
    (80, "B — Good",       "🟢"),
    (70, "C — Needs revision", "🟡"),
    (60, "D — Major revision", "🟠"),
    (0,  "F — Reject",     "🔴"),
]

# ── CSS ────────────────────────────────────────────────────────────────────
def inject_css(brand_color: str, brand_light: str):
    st.markdown(f"""
    <style>
    :root {{
        --brand: {brand_color};
        --brand-light: {brand_light};
    }}
    .brand-header {{
        background: var(--brand);
        color: white;
        padding: 16px 20px;
        border-radius: 10px;
        margin-bottom: 1.25rem;
    }}
    .brand-header h1 {{ font-size: 20px; font-weight: 600; margin: 0 0 4px; }}
    .brand-header p  {{ font-size: 12px; opacity: 0.85; margin: 0; }}

    .score-card {{
        background: var(--brand-light);
        border: 1.5px solid var(--brand);
        border-radius: 10px;
        padding: 16px 20px;
        margin-bottom: 1rem;
    }}
    .score-big   {{ font-size: 52px; font-weight: 700; color: var(--brand); line-height: 1; }}
    .grade-text  {{ font-size: 14px; font-weight: 600; margin-top: 4px; }}
    .verdict     {{ font-size: 13px; color: #444; line-height: 1.6; margin-top: 8px; }}

    .cat-card {{
        background: #fafafa;
        border: 0.5px solid #e8e8e8;
        border-radius: 8px;
        padding: 12px 14px;
        margin-bottom: 10px;
    }}
    .cat-name  {{ font-size: 13px; font-weight: 600; color: #1a1d2e; }}
    .cat-score {{ font-size: 13px; color: #555; }}
    .cat-fb    {{ font-size: 12px; color: #666; margin-top: 6px; line-height: 1.6; }}

    .detect-box {{
        border: 0.5px solid #e8e8e8;
        border-radius: 8px;
        padding: 14px;
        background: #fafafa;
    }}
    .detect-title {{ font-size: 12px; font-weight: 600; margin-bottom: 8px; }}
    .pct-big      {{ font-size: 32px; font-weight: 700; }}

    .tag-str {{ background: #d1fae5; color: #065f46; padding: 3px 10px; border-radius: 20px; font-size: 11px; font-weight: 500; display: inline-block; margin: 2px; }}
    .tag-imp {{ background: #fef3c7; color: #92400e; padding: 3px 10px; border-radius: 20px; font-size: 11px; font-weight: 500; display: inline-block; margin: 2px; }}

    .comment-box {{
        background: #fff8f0;
        border-left: 3px solid var(--brand);
        padding: 8px 12px;
        margin-bottom: 6px;
        border-radius: 0 6px 6px 0;
        font-size: 12px;
    }}
    .comment-author {{ font-weight: 600; color: var(--brand); }}

    div[data-testid="stProgress"] > div {{ background-color: var(--brand) !important; }}
    .stButton>button {{ border-radius: 8px !important; font-weight: 500 !important; }}
    </style>
    """, unsafe_allow_html=True)


# ── file parsing helpers ───────────────────────────────────────────────────
def extract_docx(file_bytes: bytes) -> dict:
    """Extract text, headings, links and comments from a .docx file."""
    if not DOCX_OK:
        return {"text": "", "headings": [], "links": [], "comments": [], "error": "python-docx not installed"}
    
    doc  = Document(BytesIO(file_bytes))
    text, headings, links = [], [], []

    for para in doc.paragraphs:
        t = para.text.strip()
        if not t:
            continue
        text.append(t)
        style = para.style.name
        if style.startswith("Heading 1"):
            headings.append({"level": "H1", "text": t})
        elif style.startswith("Heading 2"):
            headings.append({"level": "H2", "text": t})
        elif style.startswith("Heading 3"):
            headings.append({"level": "H3", "text": t})

    # hyperlinks
    for rel in doc.part.rels.values():
        if "hyperlink" in rel.reltype:
            links.append(rel._target)

    # comments from XML
    comments = []
    try:
        comments_part = doc.part.package.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        )
        for c in comments_part._element.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comment"
        ):
            author = c.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author", "Editor")
            body   = " ".join(p.text for p in c.iter() if p.text)
            if body.strip():
                comments.append({"author": author, "text": body.strip()})
    except Exception:
        pass  # no comments part in this doc

    return {
        "text":     "\n".join(text),
        "headings": headings,
        "links":    links,
        "comments": comments,
        "word_count": len(" ".join(text).split()),
    }


def extract_pdf(file_bytes: bytes) -> dict:
    if not PDF_OK:
        return {"text": "", "headings": [], "links": [], "comments": [], "error": "pdfplumber not installed"}
    
    text_parts, links = [], []
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text_parts.append(t)
            for ann in (page.annots or []):
                uri = ann.get("uri")
                if uri:
                    links.append(uri)

    full_text = "\n".join(text_parts)
    # naive heading extraction: lines in ALL CAPS or short lines
    headings = []
    for line in full_text.split("\n"):
        stripped = line.strip()
        if stripped and len(stripped) < 80 and stripped == stripped.upper() and len(stripped) > 3:
            headings.append({"level": "H2", "text": stripped.title()})

    return {
        "text":       full_text,
        "headings":   headings,
        "links":      links,
        "comments":   [],
        "word_count": len(full_text.split()),
    }


def extract_txt(file_bytes: bytes) -> dict:
    text = file_bytes.decode("utf-8", errors="ignore")
    return {
        "text":       text,
        "headings":   [],
        "links":      re.findall(r'https?://\S+', text),
        "comments":   [],
        "word_count": len(text.split()),
    }


def parse_file(uploaded_file) -> dict:
    name = uploaded_file.name.lower()
    # use getvalue() — more reliable than read() in Streamlit
    raw  = uploaded_file.getvalue()
    if name.endswith(".docx"):
        result = extract_docx(raw)
    elif name.endswith(".pdf"):
        result = extract_pdf(raw)
    else:
        result = extract_txt(raw)
    return result


# ── Claude QA evaluation ───────────────────────────────────────────────────
def call_gemini(prompt: str) -> str:
    """Call Google Gemini API using the official SDK."""
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    model    = genai.GenerativeModel("gemini-1.5-flash")
    response = model.generate_content(prompt)
    return response.text.strip()


def run_qa_evaluation(
    title: str,
    content: str,
    writer: str,
    content_type: str,
    language: str,
    brand: str,
    headings: list,
    links: list,
    comments: list,
) -> dict:

    headings_txt = "\n".join(f"  [{h['level']}] {h['text']}" for h in headings) or "  None detected"
    links_txt    = "\n".join(f"  - {l}" for l in links[:10])                    or "  None detected"
    comments_txt = "\n".join(f"  [{c['author']}]: {c['text']}" for c in comments) or "  None"

    prompt = f"""You are a senior content QA evaluator for {brand}, a leading real estate platform in the UAE and Middle East.

Evaluate this {content_type.lower()} written in {language}.

ARTICLE TITLE: {title}
WRITER: {writer}

EXTRACTED HEADINGS:
{headings_txt}

LINKS FOUND:
{links_txt}

EDITOR COMMENTS IN FILE:
{comments_txt}

ARTICLE CONTENT:
{content[:5000]}

---
Score across EXACTLY these 6 dimensions. Return ONLY a valid JSON object — no markdown, no preamble, no explanation outside the JSON.

{{
  "scores": {{
    "Content Quality":    {{"score": 0-25, "feedback": "2-3 sentences"}},
    "SEO & Structure":    {{"score": 0-20, "feedback": "2-3 sentences"}},
    "Language & Grammar": {{"score": 0-20, "feedback": "2-3 sentences"}},
    "Brand Voice":        {{"score": 0-15, "feedback": "2-3 sentences"}},
    "Readability & Flow": {{"score": 0-10, "feedback": "2-3 sentences"}},
    "Originality":        {{"score": 0-10, "feedback": "2-3 sentences"}}
  }},
  "total": <sum of all scores>,
  "overall_feedback": "3-4 sentence executive summary",
  "key_strengths":         ["strength 1", "strength 2", "strength 3"],
  "areas_for_improvement": ["area 1", "area 2", "area 3", "area 4"],
  "recommendation": "approve or revise or reject"
}}

SCORING RUBRICS:
- Content Quality (25): accuracy, depth, relevance to UAE/Middle East real estate; buyer-benefit framing
- SEO & Structure (20): heading use, keyword density, internal links, no raw Sources sections
- Language & Grammar (20): correct grammar; flag developer shorthand like G+1/G+2 for general audiences
- Brand Voice (15): independent advisory tone matching {brand}, not developer marketing language
- Readability (10): paragraph flow, sentence variety, scannability
- Originality (10): unique editorial angle, not copy-pasted from developer brochures

RECOMMENDATION THRESHOLD: approve >= 80, revise 60-79, reject < 60

Factor editor comments into relevant category scores and feedback."""

    raw   = call_gemini(prompt)
    clean = re.sub(r"```json|```", "", raw).strip()
    return json.loads(clean)


# ── simple plagiarism heuristic (no paid API needed) ──────────────────────
def run_plagiarism_check(content: str, links: list) -> dict:
    """
    Lightweight heuristic — flags repeated phrases and known source domains.
    Replace with Copyleaks API call for production.
    """
    known_domains  = ["emaar.com", "nakheel.com", "damac.com", "aldar.com", "meraas.com"]
    flagged_sources = [l for l in links if any(d in l for d in known_domains)]
    
    # rough duplicate-phrase detection
    words      = content.lower().split()
    chunks     = [" ".join(words[i:i+8]) for i in range(0, len(words) - 8, 4)]
    seen, dups = set(), 0
    for c in chunks:
        if c in seen:
            dups += 1
        seen.add(c)
    
    pct = min(int((dups / max(len(chunks), 1)) * 100) + (len(flagged_sources) * 8), 100)

    return {
        "percentage":      pct,
        "flagged_sources": flagged_sources,
        "status":          "danger" if pct > 20 else "warn" if pct > 10 else "safe",
        "note":            "⚠ For production, connect Copyleaks API in secrets.toml for accurate web-wide plagiarism detection.",
    }


# ── simple AI detection heuristic ─────────────────────────────────────────
def run_ai_detection(content: str) -> dict:
    """
    Lightweight heuristic based on lexical patterns.
    Replace with GPTZero or Originality.ai API call for production.
    """
    ai_phrases = [
        "in conclusion", "it is worth noting", "it is important to note",
        "delve into", "in the realm of", "furthermore", "moreover",
        "it goes without saying", "needless to say", "at the end of the day",
        "in today's fast-paced", "leverage", "utilize", "seamlessly",
    ]
    lower = content.lower()
    hits  = sum(1 for p in ai_phrases if p in lower)
    pct   = min(hits * 6, 60)  # cap at 60%

    return {
        "ai_pct":    pct,
        "human_pct": 100 - pct,
        "status":    "danger" if pct > 30 else "warn" if pct > 15 else "safe",
        "note":      "⚠ For production, connect GPTZero or Originality.ai API in secrets.toml.",
    }


# ── Google Sheets logging ──────────────────────────────────────────────────
def log_to_sheets(row: dict):
    if not SHEETS_OK:
        return
    try:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
        gc      = gspread.authorize(creds)
        sheet   = gc.open(st.secrets.get("SHEET_NAME", "Bayut QA Submissions")).sheet1
        
        # add header row if empty
        if sheet.row_count < 1 or not sheet.row_values(1):
            sheet.append_row([
                "Date", "Brand", "Writer", "Title", "Type", "Language",
                "QA Score", "Plagiarism %", "AI %", "Recommendation", "Editor Decision", "Notes"
            ])
        
        sheet.append_row([
            row.get("date"),          row.get("brand"),
            row.get("writer"),        row.get("title"),
            row.get("content_type"),  row.get("language"),
            row.get("qa_score"),      row.get("plagiarism_pct"),
            row.get("ai_pct"),        row.get("recommendation"),
            row.get("editor_decision", ""),
            row.get("editor_notes",   ""),
        ])
    except Exception as e:
        st.warning(f"Could not log to Google Sheets: {e}")


# ── grade helper ───────────────────────────────────────────────────────────
def get_grade(score: int) -> tuple:
    for threshold, label, icon in GRADE_MAP:
        if score >= threshold:
            return label, icon
    return GRADE_MAP[-1][1], GRADE_MAP[-1][2]


# ── sidebar ────────────────────────────────────────────────────────────────
def render_sidebar():
    with st.sidebar:
        st.markdown("## ⚙️ Settings")
        brand = st.radio("Platform", list(BRANDS.keys()), horizontal=True)
        st.divider()
        st.markdown("### 📋 Navigation")
        page = st.radio(
            "Go to",
            ["📤 Submit article", "📊 Dashboard"],
            label_visibility="collapsed",
        )
        st.divider()
        st.markdown("### ℹ️ Score thresholds")
        st.markdown("""
| Score | Grade |
|-------|-------|
| 90–100 | ✅ Approve |
| 80–89  | ✅ Approve |
| 70–79  | 🟡 Revise  |
| 60–69  | 🟠 Revise  |
| < 60   | 🔴 Reject  |
        """)
        return brand, page


# ── submit page ────────────────────────────────────────────────────────────
def page_submit(brand: str):
    b = BRANDS[brand]
    inject_css(b["primary"], b["light"])

    st.markdown(
        f'<div class="brand-header"><h1>{b["logo"]} {brand} — Content QA System</h1>'
        f'<p>Upload article · AI evaluation · Plagiarism & AI detection · Editor approval</p></div>',
        unsafe_allow_html=True,
    )

    # ── form ──────────────────────────────────────────────────────────────
    with st.form("submit_form", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            writer       = st.text_input("Writer name *", placeholder="e.g. Rabia Ahmed")
        with c2:
            article_title = st.text_input("Article title *", placeholder="e.g. Everything About Montura 2")
        c3, c4 = st.columns(2)
        with c3:
            content_type = st.selectbox("Content type", CONTENT_TYPES)
        with c4:
            language = st.selectbox("Language", LANGUAGES)

        uploaded = st.file_uploader(
            "Upload article file *",
            type=["docx", "pdf", "txt"],
            help="Headings, links and editor comments will be extracted automatically from .docx files",
        )
        submitted = st.form_submit_button(
            f"🚀 Run full evaluation — QA + Plagiarism + AI detection",
            use_container_width=True,
            type="primary",
        )

    if not submitted:
        st.info(
            "📎 Upload a `.docx` file for best results — the system will automatically read "
            "your heading structure, all hyperlinks, and any editor comments written in the document."
        )
        return

    # ── validation ────────────────────────────────────────────────────────
    if not writer or not article_title or not uploaded:
        st.error("Please fill in writer name, article title and upload a file.")
        return

    # ── parse file ────────────────────────────────────────────────────────
    with st.spinner("Reading file and extracting structure…"):
        parsed = parse_file(uploaded)

    if not parsed.get("text") or len(parsed["text"]) < 20:
        st.error(f"Could not read text from the file. File size: {uploaded.size} bytes. Please make sure the file is not password protected and try again.")
        if parsed.get("error"):
            st.error(f"Detail: {parsed['error']}")
        return

    # ── show extracted metadata ───────────────────────────────────────────
    with st.expander(f"📄 Extracted from file — {len(parsed['headings'])} headings · {len(parsed['links'])} links · {len(parsed['comments'])} comments", expanded=False):
        tc1, tc2, tc3 = st.columns(3)
        with tc1:
            st.markdown("**Headings**")
            for h in parsed["headings"]:
                st.markdown(f"`{h['level']}` {h['text']}")
            if not parsed["headings"]:
                st.caption("None detected")
        with tc2:
            st.markdown("**Links**")
            for l in parsed["links"][:8]:
                st.markdown(f"→ {l}")
            if not parsed["links"]:
                st.caption("None detected")
        with tc3:
            st.markdown("**Editor comments**")
            for c in parsed["comments"]:
                st.markdown(
                    f'<div class="comment-box"><span class="comment-author">{c["author"]}:</span> {c["text"]}</div>',
                    unsafe_allow_html=True,
                )
            if not parsed["comments"]:
                st.caption("No comments found in file")

    # ── run all checks ────────────────────────────────────────────────────
    progress = st.progress(0, text="Starting evaluation…")

    with st.spinner("Running AI quality evaluation…"):
        progress.progress(20, text="Running QA evaluation…")
        result = run_qa_evaluation(
            title        = article_title,
            content      = parsed["text"],
            writer       = writer,
            content_type = content_type,
            language     = language,
            brand        = brand,
            headings     = parsed["headings"],
            links        = parsed["links"],
            comments     = parsed["comments"],
        )

    progress.progress(60, text="Running plagiarism check…")
    plag = run_plagiarism_check(parsed["text"], parsed["links"])

    progress.progress(80, text="Running AI detection…")
    ai_det = run_ai_detection(parsed["text"])

    progress.progress(100, text="Building report…")
    progress.empty()

    # ── store in session ──────────────────────────────────────────────────
    submission = {
        "date":           datetime.now().strftime("%d %b %Y %H:%M"),
        "brand":          brand,
        "writer":         writer,
        "title":          article_title,
        "content_type":   content_type,
        "language":       language,
        "word_count":     parsed["word_count"],
        "headings":       parsed["headings"],
        "links":          parsed["links"],
        "comments":       parsed["comments"],
        "qa":             result,
        "plagiarism":     plag,
        "ai_detection":   ai_det,
        "qa_score":       result.get("total", 0),
        "plagiarism_pct": plag["percentage"],
        "ai_pct":         ai_det["ai_pct"],
        "recommendation": result.get("recommendation", "revise"),
        "editor_decision": "",
        "editor_notes":    "",
    }

    if "submissions" not in st.session_state:
        st.session_state.submissions = []
    st.session_state.submissions.append(submission)
    st.session_state.current = submission

    render_report(submission)


# ── report renderer ────────────────────────────────────────────────────────
def render_report(sub: dict):
    b     = BRANDS[sub["brand"]]
    qa    = sub["qa"]
    plag  = sub["plagiarism"]
    ai    = sub["ai_detection"]
    score = sub["qa_score"]
    grade, grade_icon = get_grade(score)

    inject_css(b["primary"], b["light"])

    st.divider()
    st.markdown(f"## 📋 QA Report — {sub['writer']} · {sub['title']}")

    # meta pills
    st.markdown(
        f"`{sub['brand']}` &nbsp; `{sub['content_type']}` &nbsp; "
        f"`{sub['language']}` &nbsp; `~{sub['word_count']} words` &nbsp; `{sub['date']}`",
        unsafe_allow_html=True,
    )

    # ── score hero ────────────────────────────────────────────────────────
    st.markdown(
        f"""<div class="score-card">
            <div class="score-big">{score}<span style="font-size:20px;font-weight:400;color:#888"> / 100</span></div>
            <div class="grade-text">{grade_icon} {grade}</div>
            <div class="verdict">{qa.get('overall_feedback','')}</div>
        </div>""",
        unsafe_allow_html=True,
    )

    # ── top metrics ───────────────────────────────────────────────────────
    m1, m2, m3 = st.columns(3)
    plag_color = {"safe": "🟢", "warn": "🟡", "danger": "🔴"}[plag["status"]]
    ai_color   = {"safe": "🟢", "warn": "🟡", "danger": "🔴"}[ai["status"]]
    m1.metric("QA Score",           f"{score} / 100")
    m2.metric("Plagiarism detected", f"{plag_color} {plag['percentage']}%")
    m3.metric("AI-generated",        f"{ai_color} {ai['ai_pct']}%")

    st.divider()

    # ── category breakdown ────────────────────────────────────────────────
    st.markdown("### 📊 Category scores")
    for cat_name, max_val in CAT_MAX.items():
        cat_data = qa["scores"].get(cat_name, {})
        s        = cat_data.get("score", 0)
        pct      = int(s / max_val * 100)
        fb       = cat_data.get("feedback", "")

        col_a, col_b = st.columns([3, 1])
        with col_a:
            st.markdown(f"**{cat_name}**")
            st.progress(pct / 100)
        with col_b:
            st.markdown(f"**{s} / {max_val}**")
        st.caption(fb)
        st.markdown("")

    st.divider()

    # ── plagiarism + AI detection ─────────────────────────────────────────
    st.markdown("### 🔍 Automated checks")
    pc1, pc2 = st.columns(2)

    with pc1:
        plag_icon  = {"safe": "✅", "warn": "⚠️", "danger": "🚨"}[plag["status"]]
        plag_label = {"safe": "Low — acceptable", "warn": "Medium — review flagged sections", "danger": "High — rewrite required"}[plag["status"]]
        st.markdown(f"#### {plag_icon} Plagiarism check")
        st.metric("Matched content", f"{plag['percentage']}%", label_visibility="visible")
        st.progress(plag["percentage"] / 100)
        st.caption(plag_label)
        if plag["flagged_sources"]:
            st.markdown("**Flagged sources:**")
            for src in plag["flagged_sources"]:
                st.markdown(f"- `{src}`")
        st.caption(plag["note"])

    with pc2:
        ai_icon  = {"safe": "✅", "warn": "⚠️", "danger": "🚨"}[ai["status"]]
        ai_label = {"safe": "Mostly human-written", "warn": "Some AI phrasing detected", "danger": "High AI content — rewrite required"}[ai["status"]]
        st.markdown(f"#### {ai_icon} AI content detection")
        st.metric("AI-generated estimate", f"{ai['ai_pct']}%")
        st.progress(ai["ai_pct"] / 100)
        st.caption(ai_label)
        ac1, ac2 = st.columns(2)
        ac1.metric("Human", f"{ai['human_pct']}%")
        ac2.metric("AI likely", f"{ai['ai_pct']}%")
        st.caption(ai["note"])

    st.divider()

    # ── structure ─────────────────────────────────────────────────────────
    st.markdown("### 🗂 Document structure")
    sc1, sc2, sc3 = st.columns(3)
    sc1.metric("Headings", len(sub["headings"]))
    sc2.metric("Links (total)", len(sub["links"]))
    internal = [l for l in sub["links"] if sub["brand"].lower() in l.lower()]
    sc3.metric("Internal links", len(internal), delta="⚠ add more" if not internal else None)

    if sub["comments"]:
        st.markdown("**Editor comments from file:**")
        for c in sub["comments"]:
            st.markdown(
                f'<div class="comment-box"><span class="comment-author">{c["author"]}:</span> {c["text"]}</div>',
                unsafe_allow_html=True,
            )

    st.divider()

    # ── strengths / improvements ──────────────────────────────────────────
    col_s, col_i = st.columns(2)
    with col_s:
        st.markdown("### ✅ Strengths")
        for s in qa.get("key_strengths", []):
            st.markdown(f'<span class="tag-str">{s}</span>', unsafe_allow_html=True)
    with col_i:
        st.markdown("### 🔧 Required improvements")
        for imp in qa.get("areas_for_improvement", []):
            st.markdown(f'<span class="tag-imp">{imp}</span>', unsafe_allow_html=True)

    st.divider()

    # ── editor decision ───────────────────────────────────────────────────
    st.markdown("### ✍️ Editor decision")
    st.caption("The AI recommendation is a guide — you make the final call.")

    rec = qa.get("recommendation", "revise")
    dec_map = {"approve": 0, "revise": 1, "reject": 2}
    decision = st.radio(
        "Decision",
        ["✅ Approve", "↩ Request revision", "🚨 Reject"],
        index=dec_map.get(rec, 1),
        horizontal=True,
        key=f"decision_{sub['title']}",
    )
    notes = st.text_area(
        "Notes for writer (required for revision / rejection)",
        height=100,
        placeholder="Be specific — tell the writer exactly what to change.",
        key=f"notes_{sub['title']}",
    )

    if st.button("✅ Confirm decision & save", type="primary", use_container_width=True):
        if ("revision" in decision or "Reject" in decision) and not notes.strip():
            st.error("Please add notes for the writer before confirming.")
        else:
            sub["editor_decision"] = decision
            sub["editor_notes"]    = notes
            log_to_sheets(sub)
            st.success(f"Decision saved: **{decision}**")
            if notes:
                st.info(f"📬 Notes for {sub['writer']}: {notes}")

    # ── print button ──────────────────────────────────────────────────────
    st.divider()
    st.caption(f"Generated by {sub['brand']} Content QA System · Powered by Claude · {sub['date']}")


# ── dashboard page ─────────────────────────────────────────────────────────
def page_dashboard(brand: str):
    b = BRANDS[brand]
    inject_css(b["primary"], b["light"])

    st.markdown(
        f'<div class="brand-header"><h1>{b["logo"]} {brand} — Submissions dashboard</h1>'
        f'<p>All writer evaluations and editor decisions</p></div>',
        unsafe_allow_html=True,
    )

    subs = [s for s in st.session_state.get("submissions", []) if s["brand"] == brand]

    if not subs:
        st.info("No submissions yet for this platform. Go to **Submit article** to run the first evaluation.")
        return

    # metrics
    approved = sum(1 for s in subs if "Approve" in s.get("editor_decision", ""))
    revision = sum(1 for s in subs if "revision" in s.get("editor_decision", ""))
    rejected = sum(1 for s in subs if "Reject" in s.get("editor_decision", ""))
    pending  = sum(1 for s in subs if not s.get("editor_decision"))

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total",    len(subs))
    m2.metric("Approved", approved)
    m3.metric("Revision", revision)
    m4.metric("Rejected", rejected)
    m5.metric("Pending",  pending)

    st.divider()

    # filter
    f1, f2, f3 = st.columns(3)
    type_filter = f1.selectbox("Content type", ["All"] + CONTENT_TYPES)
    lang_filter = f2.selectbox("Language",     ["All"] + LANGUAGES)
    dec_filter  = f3.selectbox("Status",       ["All", "Pending", "Approved", "Revision", "Rejected"])

    filtered = subs
    if type_filter != "All":
        filtered = [s for s in filtered if s["content_type"] == type_filter]
    if lang_filter != "All":
        filtered = [s for s in filtered if s["language"] == lang_filter]
    if dec_filter != "All":
        if dec_filter == "Pending":
            filtered = [s for s in filtered if not s.get("editor_decision")]
        else:
            filtered = [s for s in filtered if dec_filter.lower() in s.get("editor_decision", "").lower()]

    # table
    st.markdown(f"**{len(filtered)} submissions**")
    for sub in reversed(filtered):
        grade, icon = get_grade(sub["qa_score"])
        exp_title = (
            f"{icon} **{sub['writer']}** · {sub['title'][:50]}{'…' if len(sub['title'])>50 else ''} "
            f"· Score: **{sub['qa_score']}** · {sub['date']}"
        )
        with st.expander(exp_title, expanded=False):
            render_report(sub)


# ── main ───────────────────────────────────────────────────────────────────
def main():
    brand, page = render_sidebar()
    if "submissions" not in st.session_state:
        st.session_state.submissions = []

    if "Submit" in page:
        page_submit(brand)
    else:
        page_dashboard(brand)


if __name__ == "__main__":
    main()
