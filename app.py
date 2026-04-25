import streamlit as st
import json
import re
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

st.set_page_config(
    page_title="Content QA System",
    page_icon="Q",
    layout="wide",
    initial_sidebar_state="expanded",
)

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
    (90, "A  Excellent",      "green"),
    (80, "B  Good",           "green"),
    (70, "C  Needs revision", "orange"),
    (60, "D  Major revision", "orange"),
    (0,  "F  Reject",         "red"),
]

# AI phrases used for detection highlighting
AI_PHRASES = [
    "in conclusion", "it is worth noting", "it is important to note",
    "delve into", "in the realm of", "furthermore", "moreover",
    "needless to say", "leverage", "utilize", "seamlessly",
    "it goes without saying", "in today's", "one such", "robust",
    "cutting-edge", "state-of-the-art", "at the end of the day",
]

KNOWN_DOMAINS = [
    "emaar.com", "nakheel.com", "damac.com", "aldar.com", "meraas.com",
    "sobha.com", "omniyat.com", "ellington.ae", "azizi.ae", "reportage.ae",
]

# ── CSS ────────────────────────────────────────────────────────────────────
def inject_css():
    st.markdown("""
    <style>
    :root { --qa:#2D4A8A; --qa-light:#EEF2FB; }

    .qa-header {
        background:#1C2B5E; color:#fff;
        padding:16px 20px; border-radius:10px; margin-bottom:1.2rem;
    }
    .qa-header h1 { font-size:18px; font-weight:500; margin:0; }

    .score-hero {
        background:var(--qa-light); border:1px solid var(--qa);
        border-radius:10px; padding:16px 20px; margin-bottom:1rem;
    }
    .score-num    { font-size:52px; font-weight:500; color:var(--qa); line-height:1; }
    .score-den    { font-size:16px; font-weight:400; color:#888; }
    .score-grade  { font-size:13px; font-weight:500; margin-top:4px; color:var(--qa); }
    .score-verdict { font-size:12px; color:#444; line-height:1.65; margin-top:8px; }

    .breakdown-box {
        background:#fff; border:0.5px solid #e0e0e0;
        border-radius:8px; padding:12px 14px; margin-top:10px; font-size:12px;
    }
    .ded-row   { display:flex; justify-content:space-between; padding:4px 0;
                 color:#991b1b; border-bottom:0.5px solid #fce; }
    .base-row  { display:flex; justify-content:space-between; padding:4px 0;
                 color:#555; border-bottom:0.5px solid #eee; }
    .ok-row    { display:flex; justify-content:space-between; padding:4px 0;
                 color:#888; border-bottom:0.5px solid #eee; font-size:11px; }
    .total-row { display:flex; justify-content:space-between; padding:6px 0 2px;
                 font-weight:600; font-size:13px; color:#1a1d2e;
                 border-top:1.5px solid #ccc; margin-top:2px; }

    .detect-card {
        border:0.5px solid #e0e0e0; border-radius:10px;
        padding:14px 16px; background:#fff; height:100%;
    }
    .detect-title   { font-size:12px; font-weight:600; color:#1a1d2e; margin-bottom:6px; }
    .detect-pct     { font-size:32px; font-weight:500; line-height:1; margin-bottom:4px; }
    .detect-bar     { height:6px; background:#eee; border-radius:3px; margin-bottom:8px; }
    .detect-bar-f   { height:100%; border-radius:3px; }
    .detect-thresh  { font-size:11px; font-weight:500; padding:4px 10px;
                      border-radius:6px; display:inline-block; margin-bottom:8px; }
    .detect-note    { font-size:11px; color:#666; line-height:1.6; }
    .detect-split   { display:grid; grid-template-columns:1fr 1fr; gap:6px; margin-top:8px; }
    .detect-seg     { text-align:center; background:#f5f6fa; border-radius:6px; padding:6px; }
    .detect-seg-n   { font-size:14px; font-weight:500; }
    .detect-seg-l   { font-size:10px; color:#888; margin-top:2px; }

    .issue-block {
        background:#fffbf0; border:0.5px solid #f0d080;
        border-radius:6px; padding:8px 10px; margin-top:8px; font-size:11px;
    }
    .issue-block-title { font-size:10px; font-weight:600; color:#92400e;
                         text-transform:uppercase; letter-spacing:0.05em; margin-bottom:5px; }
    .issue-snippet {
        background:#fff; border-left:3px solid #d97706;
        padding:5px 8px; margin-bottom:4px; border-radius:0 4px 4px 0;
        font-size:11px; color:#444; line-height:1.5; font-style:italic;
    }
    .issue-snippet:last-child { margin-bottom:0; }
    .ai-highlight { background:#fef3c7; border-radius:3px; padding:0 2px;
                    font-weight:500; color:#92400e; }

    .cmt-card {
        background:#fff8f0; border-left:3px solid #2D4A8A;
        padding:8px 12px; margin-bottom:6px;
        border-radius:0 6px 6px 0; font-size:12px;
    }
    .cmt-author { font-weight:600; color:#2D4A8A; }
    .cmt-deduct { font-size:10px; color:#991b1b; font-weight:500; margin-top:3px; }
    .cat-ref    { font-size:10px; font-weight:500; padding:1px 7px; border-radius:20px;
                  background:#EEF2FB; color:#2D4A8A; margin-left:6px; }

    .suggest-item {
        display:flex; gap:10px; align-items:flex-start;
        padding:8px 0; border-bottom:0.5px solid #eee; font-size:12px;
    }
    .suggest-item:last-child { border:none; padding-bottom:0; }
    .suggest-num {
        width:22px; height:22px; border-radius:50%;
        background:#EEF2FB; color:#2D4A8A;
        font-size:10px; font-weight:600;
        display:flex; align-items:center; justify-content:center;
        flex-shrink:0; margin-top:1px;
    }
    .suggest-cat { font-size:10px; color:#888; margin-top:3px; }

    .tag-str { background:#d1fae5; color:#065f46; padding:3px 10px;
               border-radius:20px; font-size:11px; font-weight:500;
               display:inline-block; margin:2px; }
    .tag-imp { background:#fef3c7; color:#92400e; padding:3px 10px;
               border-radius:20px; font-size:11px; font-weight:500;
               display:inline-block; margin:2px; }

    .bdg     { font-size:10px; font-weight:500; padding:2px 9px; border-radius:20px; }
    .bdg-bay { background:#e8f5e9; color:#1b5e20; }
    .bdg-dub { background:#fdecea; color:#b71c1c; }

    .plat-btn {
        display:inline-block; padding:7px 20px; border-radius:8px;
        font-size:13px; font-weight:500; cursor:pointer;
        border:2px solid transparent; margin-right:8px;
        transition:all 0.15s;
    }
    .no-comments-notice {
        background:#f5f6fa; border:0.5px solid #e0e0e0;
        border-radius:6px; padding:10px 14px; font-size:12px;
        color:#666; margin-bottom:8px;
    }
    </style>
    """, unsafe_allow_html=True)


# ── Groq AI call ───────────────────────────────────────────────────────────
def call_ai(prompt: str) -> str:
    if not GROQ_OK:
        raise Exception("groq package not installed — check requirements.txt")
    client = Groq(api_key=st.secrets["GROQ_API_KEY"])
    for model in ["llama-3.1-8b-instant", "llama3-8b-8192", "gemma2-9b-it", "mixtral-8x7b-32768"]:
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=2000,
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            err = str(e).lower()
            if "model" in err or "not found" in err or "decommission" in err:
                continue
            raise e
    raise Exception("All Groq models failed. Check your GROQ_API_KEY in Streamlit Secrets.")


# ── file parsers ───────────────────────────────────────────────────────────
def extract_docx(raw):
    if not DOCX_OK:
        return {"text": "", "headings": [], "links": [], "comments": [],
                "word_count": 0, "error": "python-docx not installed"}
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
        if "hyperlink" in rel.reltype:
            links.append(rel._target)
    comments = []
    try:
        cp = doc.part.package.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        )
        for c in cp._element.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comment"
        ):
            author = c.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author", "Editor")
            body   = " ".join(x.text for x in c.iter() if x.text).strip()
            if body:
                comments.append({"author": author, "text": body})
    except Exception:
        pass
    full = "\n".join(text)
    return {"text": full, "headings": headings, "links": links,
            "comments": comments, "word_count": len(full.split()), "error": ""}

def extract_pdf(raw):
    if not PDF_OK:
        return {"text": "", "headings": [], "links": [], "comments": [],
                "word_count": 0, "error": "pdfplumber not installed"}
    parts, links = [], []
    with pdfplumber.open(BytesIO(raw)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t: parts.append(t)
            for a in (page.annots or []):
                u = a.get("uri")
                if u: links.append(u)
    full = "\n".join(parts)
    return {"text": full, "headings": [], "links": links,
            "comments": [], "word_count": len(full.split()), "error": ""}

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


# ── content issue extraction ───────────────────────────────────────────────
def find_ai_snippets(text):
    """Return sentences that contain AI phrases."""
    sentences = re.split(r'(?<=[.!?])\s+', text)
    flagged = []
    for sent in sentences:
        low = sent.lower()
        for phrase in AI_PHRASES:
            if phrase in low:
                snippet = sent.strip()
                if len(snippet) > 20 and snippet not in flagged:
                    flagged.append(snippet[:200])
                break
    return flagged[:5]  # return up to 5 snippets

def find_plag_snippets(text, links):
    """Extract actual sentences from the article that look like developer brochure copy."""
    flagged_sources = [l for l in links if any(d in l for d in KNOWN_DOMAINS)]

    brochure_phrases = [
        "inspired by", "equestrian tradition", "lush green", "rich history",
        "sense of belonging", "world-class", "premium lifestyle", "master plan",
        "seamlessly integrates", "state-of-the-art", "landmark development",
        "self-sustaining", "eco-friendly", "all-encompassing",
        "strategically located", "boasts", "showcases",
        "meticulous attention to detail", "bespoke", "craftsmanship",
        "setting the benchmark", "unparalleled", "premier destination",
        "prestigious location", "luxury finishes", "fine stone",
        "hand-finished", "custom woodwork", "timeless feel",
        "refined environment", "off-plan", "gated environment",
        "highly anticipated", "transform urban", "freehold destination",
    ]

    sentences = re.split(r"(?<=[.!?])\s+", text)
    flagged_sentences = []

    for sent in sentences:
        low = sent.strip().lower()
        if len(low) < 30:
            continue
        for phrase in brochure_phrases:
            if phrase in low and sent.strip() not in flagged_sentences:
                flagged_sentences.append(sent.strip()[:230])
                break

    return flagged_sources, flagged_sentences[:5]


# ── scoring ────────────────────────────────────────────────────────────────
def apply_deductions(base_score, comments, plag_pct, ai_pct):
    comment_deduction = len(comments)
    plag_deduction    = 5 if plag_pct > 20 else 0
    ai_deduction      = 5 if ai_pct   > 20 else 0
    final             = max(0, base_score - comment_deduction - plag_deduction - ai_deduction)
    return final, {
        "base_score":        base_score,
        "comment_count":     len(comments),
        "comment_deduction": comment_deduction,
        "plag_pct":          plag_pct,
        "plag_deduction":    plag_deduction,
        "ai_pct":            ai_pct,
        "ai_deduction":      ai_deduction,
        "final_score":       final,
    }

def get_recommendation(score):
    if score >= 80: return "approve"
    if score >= 60: return "revise"
    return "reject"

def get_grade(score):
    for t, label, color in GRADE_MAP:
        if score >= t:
            return label, color
    return GRADE_MAP[-1][1], GRADE_MAP[-1][2]


# ── QA evaluation ── only scores from comments, max score if no comments ───
def run_qa(title, content, writer, ctype, lang, platform, headings, links, comments):
    h_txt = "\n".join(f"  [{h['level']}] {h['text']}" for h in headings) or "  None"
    l_txt = "\n".join(f"  - {l}" for l in links[:8])                      or "  None"

    if not comments:
        # No comments — give max scores across all categories, no AI opinion
        scores = {cat: {"score": mx, "feedback": "No editor comments found. Full marks awarded.", "comment_refs": []}
                  for cat, mx in CAT_MAX.items()}
        total = sum(CAT_MAX.values())
        return {
            "scores":               scores,
            "total":                total,
            "overall_feedback":     "No editor comments were found in the uploaded file. All categories have been awarded full marks. Scores will be adjusted if editor comments are added.",
            "key_strengths":        [],
            "areas_for_improvement":[],
            "suggestions":          [],
        }

    c_txt = "\n".join(
        f"  Comment {i+1} [{c['author']}]: {c['text']}"
        for i, c in enumerate(comments)
    )

    prompt = f"""You are a senior content QA evaluator for {platform}, a leading UAE real estate platform.

TITLE: {title}
WRITER: {writer}
CONTENT TYPE: {ctype}
LANGUAGE: {lang}

HEADINGS IN FILE:
{h_txt}

LINKS IN FILE:
{l_txt}

EDITOR COMMENTS FROM FILE:
{c_txt}

ARTICLE CONTENT (for context only — do NOT score independently):
{content[:3000]}

CRITICAL RULES:
1. Score ONLY based on the editor comments listed above. Do NOT independently evaluate the article content.
2. Map each comment to the most relevant category and reduce that category's score.
3. If no comments relate to a category, award that category its MAXIMUM score.
4. Every comment MUST be reflected in at least one category's score and feedback.
5. "comment_refs" must list the comment numbers (1, 2, 3...) that affected each category.
6. Suggestions must directly address the comments — do not invent new issues.

Return ONLY valid JSON:
{{
  "scores": {{
    "Content Quality":    {{"score": <0-25>, "feedback": "<what comments flagged>", "comment_refs": []}},
    "SEO & Structure":    {{"score": <0-20>, "feedback": "<what comments flagged>", "comment_refs": []}},
    "Language & Grammar": {{"score": <0-20>, "feedback": "<what comments flagged>", "comment_refs": []}},
    "Brand Voice":        {{"score": <0-15>, "feedback": "<what comments flagged>", "comment_refs": []}},
    "Readability & Flow": {{"score": <0-10>, "feedback": "<what comments flagged>", "comment_refs": []}},
    "Originality":        {{"score": <0-10>, "feedback": "<what comments flagged>", "comment_refs": []}}
  }},
  "total": <sum>,
  "overall_feedback": "<summary referencing the specific comments>",
  "key_strengths": [],
  "areas_for_improvement": ["<from comment 1>", "<from comment 2>"],
  "suggestions": [
    {{"number": 1, "action": "<specific fix from comment 1>", "category": "<category>"}},
    {{"number": 2, "action": "<specific fix from comment 2>", "category": "<category>"}},
    {{"number": 3, "action": "<specific fix from comment 3>", "category": "<category>"}}
  ]
}}"""

    raw   = call_ai(prompt)
    clean = re.sub(r"```json|```", "", raw).strip()
    match = re.search(r'\{.*\}', clean, re.DOTALL)
    if match:
        clean = match.group(0)
    return json.loads(clean)


# ── plagiarism heuristic ───────────────────────────────────────────────────
def check_plagiarism(text, links):
    flagged = [l for l in links if any(d in l for d in KNOWN_DOMAINS)]
    words   = text.lower().split()
    chunks  = [" ".join(words[i:i+8]) for i in range(0, max(len(words)-8, 1), 4)]
    seen, dups = set(), 0
    for c in chunks:
        if c in seen: dups += 1
        seen.add(c)
    pct = min(int((dups / max(len(chunks), 1)) * 100) + len(flagged) * 8, 100)
    return {
        "percentage":      pct,
        "flagged_sources": flagged,
        "status":          "danger" if pct > 20 else "warn" if pct > 10 else "safe",
    }

def check_ai(text):
    hits = sum(1 for p in AI_PHRASES if p in text.lower())
    pct  = min(hits * 5, 65)
    return {
        "ai_pct":    pct,
        "human_pct": 100 - pct,
        "status":    "danger" if pct > 20 else "warn" if pct > 10 else "safe",
    }


# ── Google Sheets logger ───────────────────────────────────────────────────
def log_to_sheets(row):
    if not SHEETS_OK:
        return
    try:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
        gc    = gspread.authorize(creds)
        sheet = gc.open(st.secrets.get("SHEET_NAME", "Bayut QA Submissions")).sheet1
        if not sheet.row_values(1):
            sheet.append_row([
                "Date","Platform","Writer","Title","Type","Language",
                "Base Score","Comments","Plagiarism Ded","AI Ded",
                "Final Score","Plagiarism%","AI%","Recommendation",
                "Editor Decision","Notes"
            ])
        d = row.get("deductions", {})
        sheet.append_row([
            row.get("date"), row.get("platform"), row.get("writer"),
            row.get("title"), row.get("content_type"), row.get("language"),
            d.get("base_score",0), d.get("comment_deduction",0),
            d.get("plag_deduction",0), d.get("ai_deduction",0),
            d.get("final_score",0), row.get("plagiarism_pct",0),
            row.get("ai_pct",0), row.get("recommendation",""),
            row.get("editor_decision",""), row.get("editor_notes",""),
        ])
    except Exception as e:
        st.warning(f"Could not log to Google Sheets: {e}")


# ── sidebar ────────────────────────────────────────────────────────────────
def sidebar():
    with st.sidebar:
        st.markdown("## Content QA System")
        st.divider()
        st.markdown("### Navigation")
        page = st.radio("Go to", ["Submit article", "Dashboard"],
                        label_visibility="collapsed")
        st.divider()
        st.markdown("### Score guide")
        st.markdown("""
| Score | Grade |
|---|---|
| 90 – 100 | Approve |
| 80 – 89  | Approve |
| 70 – 79  | Revise  |
| 60 – 69  | Revise  |
| Below 60 | Reject  |
        """)
        st.divider()
        st.markdown("### Deduction rules")
        st.markdown("""
| Rule | Points |
|---|---|
| Per editor comment | - 1 |
| Plagiarism over 20% | - 5 |
| AI content over 20% | - 5 |
        """)
        return page


# ── platform selector ──────────────────────────────────────────────────────
def platform_selector():
    """Coloured Bayut/Dubizzle buttons using session state."""
    if "platform" not in st.session_state:
        st.session_state.platform = "Bayut"

    st.markdown("**Platform**")
    col_b, col_d, col_rest = st.columns([1, 1, 4])

    bay_style = (
        "background:#2e7d32;color:#fff;border:2px solid #2e7d32;"
        if st.session_state.platform == "Bayut"
        else "background:#fff;color:#2e7d32;border:2px solid #2e7d32;"
    )
    dub_style = (
        "background:#c62828;color:#fff;border:2px solid #c62828;"
        if st.session_state.platform == "Dubizzle"
        else "background:#fff;color:#c62828;border:2px solid #c62828;"
    )

    st.markdown(
        f"""
        <div style="display:flex;gap:8px;margin-bottom:8px">
          <div onclick="" style="padding:7px 20px;border-radius:8px;font-size:13px;
               font-weight:500;cursor:pointer;{bay_style}">Bayut</div>
          <div onclick="" style="padding:7px 20px;border-radius:8px;font-size:13px;
               font-weight:500;cursor:pointer;{dub_style}">Dubizzle</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # actual functional buttons hidden behind radio
    choice = st.radio(
        "Platform select",
        ["Bayut", "Dubizzle"],
        horizontal=True,
        label_visibility="collapsed",
        key="platform",
    )
    return choice


# ── submit page ────────────────────────────────────────────────────────────
def page_submit():
    inject_css()
    st.markdown(
        '<div class="qa-header"><h1>Content QA System</h1></div>',
        unsafe_allow_html=True,
    )

    with st.form("qa_form"):
        c1, c2 = st.columns(2)
        writer = c1.text_input("Writer name", placeholder="e.g. Sarah Ahmed")
        title  = c2.text_input("Article title", placeholder="e.g. Everything About Montura 2")
        c3, c4 = st.columns(2)
        ctype  = c3.selectbox("Content type", CONTENT_TYPES)
        lang   = c4.selectbox("Language", LANGUAGES)

        st.markdown("**Platform**")
        platform = st.radio(
            "Platform",
            PLATFORMS,
            horizontal=True,
            label_visibility="collapsed",
        )

        # Colour the selected platform label via markdown
        bay_sel = platform == "Bayut"
        dub_sel = platform == "Dubizzle"
        st.markdown(
            f'<div style="margin-top:-8px;margin-bottom:8px;display:flex;gap:8px">'
            f'<span style="padding:4px 14px;border-radius:20px;font-size:12px;font-weight:500;'
            f'background:{"#2e7d32" if bay_sel else "#e8f5e9"};'
            f'color:{"#fff" if bay_sel else "#2e7d32"}">Bayut</span>'
            f'<span style="padding:4px 14px;border-radius:20px;font-size:12px;font-weight:500;'
            f'background:{"#c62828" if dub_sel else "#fdecea"};'
            f'color:{"#fff" if dub_sel else "#c62828"}">Dubizzle</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

        upload = st.file_uploader(
            "Upload article file",
            type=["docx", "pdf", "txt"],
            help=".docx recommended — headings, links and editor comments are extracted automatically",
        )
        go = st.form_submit_button("Run full evaluation", use_container_width=True, type="primary")

    if not go:
        st.info(
            "Upload a .docx file. Category scores are based entirely on editor comments "
            "found in the document. If no comments exist, full marks are awarded."
        )
        return

    if not writer or not title or not upload:
        st.error("Please fill in writer name, title and upload a file.")
        return

    with st.spinner("Reading file..."):
        parsed = parse_file(upload)

    if not parsed["text"] or len(parsed["text"]) < 30:
        st.error(f"Could not read text from file. {parsed.get('error','')}")
        return

    with st.expander(
        f"Extracted from file — {len(parsed['headings'])} headings, "
        f"{len(parsed['links'])} links, {len(parsed['comments'])} editor comments"
    ):
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
                    f'<div class="cmt-card">'
                    f'<span class="cmt-author">Comment {idx} — {c["author"]}</span><br>'
                    f'{c["text"]}'
                    f'<div class="cmt-deduct">1 point deducted from final score</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
            if not parsed["comments"]:
                st.caption("No comments found in file")

    prog = st.progress(0, text="Starting evaluation...")

    try:
        prog.progress(20, text="Scoring categories from editor comments...")
        qa = run_qa(
            title, parsed["text"], writer, ctype, lang, platform,
            parsed["headings"], parsed["links"], parsed["comments"],
        )
    except Exception as e:
        st.error(f"AI evaluation failed: {e}")
        st.info("Make sure GROQ_API_KEY is set in Streamlit Settings and Secrets.")
        return

    prog.progress(60, text="Running plagiarism check...")
    plag = check_plagiarism(parsed["text"], parsed["links"])

    prog.progress(78, text="Running AI detection...")
    ai = check_ai(parsed["text"])

    prog.progress(90, text="Finding content issues...")
    ai_snippets   = find_ai_snippets(parsed["text"])
    plag_sources, plag_snippets = find_plag_snippets(parsed["text"], parsed["links"])

    prog.progress(100, text="Done!")
    prog.empty()

    base_score  = qa.get("total", 0)
    final_score, deductions = apply_deductions(
        base_score, parsed["comments"], plag["percentage"], ai["ai_pct"]
    )
    recommendation = get_recommendation(final_score)

    sub = {
        "date":            datetime.now().strftime("%d %b %Y %H:%M"),
        "platform":        platform,
        "writer":          writer,
        "title":           title,
        "content_type":    ctype,
        "language":        lang,
        "word_count":      parsed["word_count"],
        "headings":        parsed["headings"],
        "links":           parsed["links"],
        "comments":        parsed["comments"],
        "qa":              qa,
        "plagiarism":      plag,
        "ai_detection":    ai,
        "ai_snippets":     ai_snippets,
        "plag_snippets":   plag_snippets,
        "plag_sources":    plag_sources,
        "deductions":      deductions,
        "qa_score":        final_score,
        "plagiarism_pct":  plag["percentage"],
        "ai_pct":          ai["ai_pct"],
        "recommendation":  recommendation,
        "editor_decision": "",
        "editor_notes":    "",
    }

    if "submissions" not in st.session_state:
        st.session_state.submissions = []
    st.session_state.submissions.append(sub)
    render_report(sub)


# ── report renderer ────────────────────────────────────────────────────────
def render_report(sub):
    inject_css()
    qa    = sub["qa"]
    plag  = sub["plagiarism"]
    ai    = sub["ai_detection"]
    ded   = sub["deductions"]
    score = sub["qa_score"]
    grade, _ = get_grade(score)
    rec   = sub["recommendation"]

    ai_snippets  = sub.get("ai_snippets", [])
    plag_snippets = sub.get("plag_snippets", [])
    plag_sources  = sub.get("plag_sources", [])

    st.divider()

    plat      = sub["platform"]
    bdg_class = "bdg-bay" if plat == "Bayut" else "bdg-dub"
    plat_html = f'<span class="bdg {bdg_class}">{plat}</span>'
    st.markdown(
        f"**{sub['writer']}** &nbsp; {plat_html} &nbsp; "
        f"`{sub['content_type']}` &nbsp; `{sub['language']}` &nbsp; "
        f"`{sub['word_count']} words` &nbsp; `{sub['date']}`",
        unsafe_allow_html=True,
    )
    st.markdown("")

    rec_labels = {
        "approve": ("Approve",          "#d1fae5", "#065f46"),
        "revise":  ("Request revision", "#fef3c7", "#92400e"),
        "reject":  ("Reject",           "#fee2e2", "#991b1b"),
    }
    rl, rbg, rtc = rec_labels.get(rec, rec_labels["revise"])

    def br(cls, label, val):
        return f'<div class="{cls}"><span>{label}</span><span>{val}</span></div>'

    breakdown = (
        br("base-row", "Base score (from editor comments)", f'{ded["base_score"]} / 100') +
        (br("ded-row", f'Editor comments ({ded["comment_count"]} comments, 1 pt each)', f'- {ded["comment_deduction"]} pts')
         if ded["comment_deduction"] > 0 else br("ok-row", "Editor comments", "no deduction")) +
        (br("ded-row", f'Plagiarism {ded["plag_pct"]}% — over 20% threshold', f'- {ded["plag_deduction"]} pts')
         if ded["plag_deduction"] > 0 else br("ok-row", f'Plagiarism {ded["plag_pct"]}% — under 20% threshold', "no deduction")) +
        (br("ded-row", f'AI content {ded["ai_pct"]}% — over 20% threshold', f'- {ded["ai_deduction"]} pts')
         if ded["ai_deduction"] > 0 else br("ok-row", f'AI content {ded["ai_pct"]}% — under 20% threshold', "no deduction")) +
        br("total-row", "Final score", f"{score} / 100")
    )

    st.markdown(
        f'<div class="score-hero">'
        f'<div class="score-num">{score}<span class="score-den"> / 100</span></div>'
        f'<div class="score-grade">{grade}</div>'
        f'<div style="display:inline-block;margin:6px 0 8px;padding:3px 12px;'
        f'border-radius:20px;background:{rbg};color:{rtc};font-size:11px;font-weight:500">{rl}</div>'
        f'<div class="score-verdict">{qa.get("overall_feedback","")}</div>'
        f'<div class="breakdown-box">{breakdown}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    st.divider()

    # ── plagiarism and AI cards with content snippets ──────────────────────
    st.markdown("#### Plagiarism and AI detection")
    pc1, pc2 = st.columns(2)

    with pc1:
        plag_pct   = plag["percentage"]
        plag_over  = plag_pct > 20
        p_color    = "#dc2626" if plag_over else "#059669"
        p_thresh   = (
            f'<span class="detect-thresh" style="background:#fee2e2;color:#991b1b">'
            f'{plag_pct}% — over 20% threshold — 5 points deducted</span>'
            if plag_over else
            f'<span class="detect-thresh" style="background:#d1fae5;color:#065f46">'
            f'{plag_pct}% — under 20% threshold — no deduction</span>'
        )
        # build snippets HTML — highlight brochure phrases inside copied sentences
        snip_html = ""
        if plag_over and (plag_snippets or plag_sources):
            snip_html = '<div class="issue-block"><div class="issue-block-title">Copied content detected</div>'
            for src in plag_sources[:2]:
                snip_html += f'<div class="issue-snippet"><strong style="color:#92400e">Source:</strong> {src}</div>'
            for s in plag_snippets[:3]:
                # highlight brochure phrases within the sentence
                highlighted = s
                for phrase in ["world-class","strategically located","setting the benchmark",
                               "luxury finishes","master plan","state-of-the-art",
                               "lush green","gated environment","off-plan","unparalleled",
                               "premium lifestyle","hand-finished","custom woodwork",
                               "seamlessly integrates","rich history","bespoke","craftsmanship",
                               "highly anticipated","premier destination","freehold destination"]:
                    if phrase in highlighted.lower():
                        import re as _re
                        pat = _re.compile(_re.escape(phrase), _re.IGNORECASE)
                        highlighted = pat.sub(
                            f'<span style="background:#fecaca;border-radius:3px;padding:0 3px;font-weight:500;color:#7f1d1d">{phrase}</span>',
                            highlighted)
                snip_html += f'<div class="issue-snippet">{highlighted}</div>'
            snip_html += '</div>'

        st.markdown(
            f'<div class="detect-card">'
            f'<div class="detect-title">Plagiarism check</div>'
            f'<div class="detect-bar"><div class="detect-bar-f" '
            f'style="width:{min(plag_pct,100)}%;background:{p_color}"></div></div>'
            f'{p_thresh}'
            f'<div class="detect-note">'
            f'{"Content matched external sources. Rewrite the flagged sections completely." if plag_over else "Content is within the acceptable range."}'
            f'</div>'
            f'{snip_html}'
            f'</div>',
            unsafe_allow_html=True,
        )

    with pc2:
        ai_pct   = ai["ai_pct"]
        ai_over  = ai_pct > 20
        a_color  = "#dc2626" if ai_over else "#059669"
        a_thresh = (
            f'<span class="detect-thresh" style="background:#fee2e2;color:#991b1b">'
            f'{ai_pct}% — over 20% threshold — 5 points deducted</span>'
            if ai_over else
            f'<span class="detect-thresh" style="background:#d1fae5;color:#065f46">'
            f'{ai_pct}% — under 20% threshold — no deduction</span>'
        )
        # highlight AI phrases in snippets
        snip_html = ""
        if ai_over and ai_snippets:
            snip_html = '<div class="issue-block"><div class="issue-block-title">Flagged sentences</div>'
            for s in ai_snippets[:3]:
                highlighted = s
                for phrase in AI_PHRASES:
                    if phrase in highlighted.lower():
                        pattern = re.compile(re.escape(phrase), re.IGNORECASE)
                        highlighted = pattern.sub(
                            f'<span class="ai-highlight">{phrase}</span>', highlighted
                        )
                snip_html += f'<div class="issue-snippet">{highlighted}</div>'
            snip_html += '</div>'

        st.markdown(
            f'<div class="detect-card">'
            f'<div class="detect-title">AI content detection</div>'
            f'<div class="detect-bar"><div class="detect-bar-f" '
            f'style="width:{min(ai_pct,100)}%;background:{a_color}"></div></div>'
            f'{a_thresh}'
            f'<div class="detect-note">'
            f'{"High AI content detected. Flagged sentences shown below." if ai_over else "Content appears mostly human-written."}'
            f'</div>'
            f'<div class="detect-split">'
            f'<div class="detect-seg"><div class="detect-seg-n" style="color:#059669">{ai["human_pct"]}%</div>'
            f'<div class="detect-seg-l">Human</div></div>'
            f'<div class="detect-seg"><div class="detect-seg-n" style="color:{a_color}">{ai_pct}%</div>'
            f'<div class="detect-seg-l">AI likely</div></div>'
            f'</div>'
            f'{snip_html}'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.divider()

    # ── category scores ────────────────────────────────────────────────────
    st.markdown("#### Category scores")
    if not sub["comments"]:
        st.markdown(
            '<div class="no-comments-notice">'
            'No editor comments were found in the uploaded file. '
            'All categories have been awarded full marks. '
            'Add editor comments to the document to get a meaningful score.'
            '</div>',
            unsafe_allow_html=True,
        )

    for cat, mx in CAT_MAX.items():
        data     = qa["scores"].get(cat, {})
        s        = data.get("score", 0)
        feedback = data.get("feedback", "")
        refs     = data.get("comment_refs", [])

        col_a, col_b = st.columns([4, 1])
        ref_html = ""
        if refs:
            ref_html = " ".join(
                f'<span class="cat-ref">Comment {r}</span>' for r in refs
            )
        col_a.markdown(
            f"**{cat}**" + (f' &nbsp; {ref_html}' if ref_html else ""),
            unsafe_allow_html=True,
        )
        col_a.progress(s / mx)
        col_b.markdown(f"**{s} / {mx}**")
        st.caption(feedback)
        st.markdown("")

    st.divider()

    # structure
    st.markdown("#### Document structure")
    sc1, sc2, sc3 = st.columns(3)
    sc1.metric("Headings",       len(sub["headings"]))
    sc2.metric("Total links",    len(sub["links"]))
    internal = [l for l in sub["links"] if sub["platform"].lower() in l.lower()]
    sc3.metric("Internal links", len(internal))

    if sub["comments"]:
        st.markdown(f"**Editor comments — {len(sub['comments'])} found**")
        for idx, c in enumerate(sub["comments"], 1):
            st.markdown(
                f'<div class="cmt-card">'
                f'<span class="cmt-author">Comment {idx} — {c["author"]}</span><br>'
                f'{c["text"]}'
                f'<div class="cmt-deduct">1 point deducted from final score</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    st.divider()

    col_s, col_i = st.columns(2)
    with col_s:
        st.markdown("#### Strengths")
        strengths = qa.get("key_strengths", [])
        if strengths:
            for s in strengths:
                st.markdown(f'<span class="tag-str">{s}</span>', unsafe_allow_html=True)
        else:
            st.caption("Strengths will appear once editor comments are provided.")
    with col_i:
        st.markdown("#### Required improvements")
        for imp in qa.get("areas_for_improvement", []):
            st.markdown(f'<span class="tag-imp">{imp}</span>', unsafe_allow_html=True)

    suggestions = qa.get("suggestions", [])
    if suggestions:
        st.divider()
        st.markdown("#### Suggestions to improve the article")
        st.caption("Specific actions to address each editor comment")
        for sug in suggestions:
            st.markdown(
                f'<div class="suggest-item">'
                f'<div class="suggest-num">{sug.get("number","")}</div>'
                f'<div><div>{sug.get("action","")}</div>'
                f'<div class="suggest-cat">Addresses: {sug.get("category","")}</div>'
                f'</div></div>',
                unsafe_allow_html=True,
            )

    st.divider()
    st.markdown("#### Editor decision")
    st.caption("The AI recommendation is a guide. You make the final call.")
    rec_idx  = {"approve": 0, "revise": 1, "reject": 2}
    decision = st.radio(
        "Decision",
        ["Approve", "Request revision", "Reject"],
        index=rec_idx.get(rec, 1),
        horizontal=True,
        key=f"dec_{sub['title']}_{sub['date']}",
    )
    notes = st.text_area(
        "Notes for writer (required for revision and rejection)",
        height=90,
        placeholder="Tell the writer exactly what to fix.",
        key=f"notes_{sub['title']}_{sub['date']}",
    )
    if st.button("Confirm decision", type="primary", use_container_width=True,
                 key=f"conf_{sub['title']}_{sub['date']}"):
        if decision in ("Request revision", "Reject") and not notes.strip():
            st.error("Please add notes for the writer before confirming.")
        else:
            sub["editor_decision"] = decision
            sub["editor_notes"]    = notes
            log_to_sheets(sub)
            st.success(f"Decision saved: {decision}")
            if notes:
                st.info(f"Notes for {sub['writer']}: {notes}")

    st.caption(f"Content QA System — {sub['platform']} — Powered by Groq — {sub['date']}")


# ── dashboard ──────────────────────────────────────────────────────────────
def page_dashboard():
    inject_css()
    st.markdown(
        '<div class="qa-header"><h1>Dashboard</h1></div>',
        unsafe_allow_html=True,
    )

    all_subs = st.session_state.get("submissions", [])
    if not all_subs:
        st.info("No submissions yet. Go to Submit article to start.")
        return

    approved = sum(1 for s in all_subs if s.get("editor_decision") == "Approve")
    revision = sum(1 for s in all_subs if s.get("editor_decision") == "Request revision")
    rejected = sum(1 for s in all_subs if s.get("editor_decision") == "Reject")
    pending  = sum(1 for s in all_subs if not s.get("editor_decision"))

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total",    len(all_subs))
    m2.metric("Approved", approved)
    m3.metric("Revision", revision)
    m4.metric("Rejected", rejected)
    m5.metric("Pending",  pending)

    st.divider()

    all_writers   = sorted(set(s["writer"] for s in all_subs if s.get("writer")))
    f1, f2, f3, f4, f5 = st.columns(5)
    writer_filter = f1.selectbox("Writer",       ["All"] + all_writers)
    plat_filter   = f2.selectbox("Platform",     ["All"] + PLATFORMS)
    type_filter   = f3.selectbox("Content type", ["All"] + CONTENT_TYPES)
    lang_filter   = f4.selectbox("Language",     ["All"] + LANGUAGES)
    status_filter = f5.selectbox("Status",       ["All", "Pending", "Approve", "Request revision", "Reject"])

    filtered = all_subs
    if writer_filter != "All": filtered = [s for s in filtered if s.get("writer")       == writer_filter]
    if plat_filter   != "All": filtered = [s for s in filtered if s["platform"]         == plat_filter]
    if type_filter   != "All": filtered = [s for s in filtered if s["content_type"]     == type_filter]
    if lang_filter   != "All": filtered = [s for s in filtered if s["language"]         == lang_filter]
    if status_filter != "All":
        if status_filter == "Pending":
            filtered = [s for s in filtered if not s.get("editor_decision")]
        else:
            filtered = [s for s in filtered if s.get("editor_decision") == status_filter]

    st.markdown(f"**{len(filtered)} submissions**")
    for sub in reversed(filtered):
        dec       = sub.get("editor_decision") or "Pending"
        plag_flag = "  High plagiarism" if sub.get("plagiarism_pct", 0) > 20 else ""
        ai_flag   = "  High AI content" if sub.get("ai_pct", 0) > 20 else ""
        label = (
            f"{sub['writer']} ({sub['platform']}) — "
            f"{sub['title'][:45]}{'...' if len(sub['title'])>45 else ''} "
            f"| Score: {sub['qa_score']} | {dec}{plag_flag}{ai_flag} | {sub['date']}"
        )
        with st.expander(label):
            render_report(sub)


# ── main ───────────────────────────────────────────────────────────────────
def main():
    if "submissions" not in st.session_state:
        st.session_state.submissions = []
    page = sidebar()
    if "Submit" in page:
        page_submit()
    else:
        page_dashboard()


if __name__ == "__main__":
    main()
