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

# ── config ─────────────────────────────────────────────────────────────────
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

# ── CSS ────────────────────────────────────────────────────────────────────
def inject_css():
    st.markdown("""
    <style>
    :root { --qa:#2D4A8A; --qa-light:#EEF2FB; }

    .qa-header {
        background:#1C2B5E; color:#fff;
        padding:16px 20px; border-radius:10px; margin-bottom:1.2rem;
    }
    .qa-header h1 { font-size:18px; font-weight:500; margin:0 0 3px; }
    .qa-header p  { font-size:11px; opacity:.65; margin:0; }

    .score-hero {
        background:var(--qa-light); border:1px solid var(--qa);
        border-radius:10px; padding:16px 20px; margin-bottom:1rem;
    }
    .score-num { font-size:52px; font-weight:500; color:var(--qa); line-height:1; }
    .score-den { font-size:16px; font-weight:400; color:#888; }
    .score-grade { font-size:13px; font-weight:500; margin-top:4px; color:var(--qa); }
    .score-verdict { font-size:12px; color:#444; line-height:1.65; margin-top:8px; }

    .breakdown-box {
        background:#fff; border:0.5px solid #e0e0e0;
        border-radius:8px; padding:12px 14px; margin-top:10px; font-size:12px;
    }
    .ded-row { display:flex; justify-content:space-between; padding:4px 0;
               color:#991b1b; border-bottom:0.5px solid #fce; }
    .base-row { display:flex; justify-content:space-between; padding:4px 0;
                color:#555; border-bottom:0.5px solid #eee; }
    .ok-row   { display:flex; justify-content:space-between; padding:4px 0;
                color:#888; border-bottom:0.5px solid #eee; font-size:11px; }
    .total-row { display:flex; justify-content:space-between; padding:6px 0 2px;
                 font-weight:600; font-size:13px; color:#1a1d2e;
                 border-top:1.5px solid #ccc; margin-top:2px; }

    .plat-card {
        border:1.5px solid #e0e0e0; border-radius:8px;
        padding:10px 14px; cursor:pointer; transition:all 0.15s;
        display:flex; align-items:center; gap:10px;
    }
    .plat-card.selected { border-color:var(--qa); background:var(--qa-light); }
    .plat-dot { width:10px; height:10px; border-radius:50%; flex-shrink:0; }
    .plat-name { font-size:13px; font-weight:500; }
    .plat-url  { font-size:10px; color:#888; margin-top:1px; }

    .cmt-card {
        background:#fff8f0; border-left:3px solid var(--qa);
        padding:8px 12px; margin-bottom:6px;
        border-radius:0 6px 6px 0; font-size:12px;
    }
    .cmt-author { font-weight:600; color:var(--qa); }
    .cmt-deduct { font-size:10px; color:#991b1b; font-weight:500; margin-top:2px; }

    .suggest-item {
        display:flex; gap:10px; align-items:flex-start;
        padding:8px 0; border-bottom:0.5px solid #eee; font-size:12px;
    }
    .suggest-item:last-child { border:none; padding-bottom:0; }
    .suggest-num {
        width:22px; height:22px; border-radius:50%;
        background:var(--qa-light); color:var(--qa);
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

    .bdg { font-size:10px; font-weight:500; padding:2px 9px; border-radius:20px; }
    .bdg-bay { background:#e8f5e9; color:#1b5e20; }
    .bdg-dub { background:#fdecea; color:#b71c1c; }
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


# ── scoring logic ──────────────────────────────────────────────────────────
def apply_deductions(base_score, comments, plag_pct, ai_pct):
    """
    Apply deductions to base AI score:
      - 1 point per editor comment
      - 5 points if plagiarism > 20%
      - 5 points if AI content > 20%
    Returns final score and breakdown dict.
    """
    comment_deduction = len(comments)
    plag_deduction    = 5 if plag_pct > 20 else 0
    ai_deduction      = 5 if ai_pct   > 20 else 0
    total_deduction   = comment_deduction + plag_deduction + ai_deduction
    final             = max(0, base_score - total_deduction)

    return final, {
        "base_score":         base_score,
        "comment_count":      len(comments),
        "comment_deduction":  comment_deduction,
        "plag_pct":           plag_pct,
        "plag_deduction":     plag_deduction,
        "ai_pct":             ai_pct,
        "ai_deduction":       ai_deduction,
        "total_deduction":    total_deduction,
        "final_score":        final,
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


# ── QA evaluation ──────────────────────────────────────────────────────────
def run_qa(title, content, writer, ctype, lang, platform, headings, links, comments):
    h_txt = "\n".join(f"  [{h['level']}] {h['text']}" for h in headings) or "  None"
    l_txt = "\n".join(f"  - {l}" for l in links[:8])                      or "  None"
    c_txt = "\n".join(f"  [{c['author']}]: {c['text']}" for c in comments) or "  None"

    prompt = f"""You are a senior content QA evaluator for {platform}, a leading UAE real estate platform.
Evaluate this {ctype.lower()} written in {lang}.

TITLE: {title}
WRITER: {writer}
HEADINGS FOUND IN FILE:
{h_txt}
LINKS FOUND IN FILE:
{l_txt}
EDITOR COMMENTS FROM FILE (these are real feedback issues):
{c_txt}

ARTICLE CONTENT:
{content[:4000]}

IMPORTANT: The editor comments above represent real issues flagged in the document.
Your evaluation must reflect and address each of these comments in the relevant category scores and feedback.

Return ONLY valid JSON — no markdown, no text outside JSON:
{{
  "scores": {{
    "Content Quality":    {{"score": <0-25>, "feedback": "<2 sentences>"}},
    "SEO & Structure":    {{"score": <0-20>, "feedback": "<2 sentences>"}},
    "Language & Grammar": {{"score": <0-20>, "feedback": "<2 sentences>"}},
    "Brand Voice":        {{"score": <0-15>, "feedback": "<2 sentences>"}},
    "Readability & Flow": {{"score": <0-10>, "feedback": "<2 sentences>"}},
    "Originality":        {{"score": <0-10>, "feedback": "<2 sentences>"}}
  }},
  "total": <sum of all 6 scores>,
  "overall_feedback": "<3 sentence summary>",
  "key_strengths": ["<s1>", "<s2>", "<s3>"],
  "areas_for_improvement": ["<a1>", "<a2>", "<a3>", "<a4>"],
  "suggestions": [
    {{"number": 1, "action": "<specific rewrite instruction>", "category": "<which score it fixes>"}},
    {{"number": 2, "action": "<specific rewrite instruction>", "category": "<which score it fixes>"}},
    {{"number": 3, "action": "<specific rewrite instruction>", "category": "<which score it fixes>"}},
    {{"number": 4, "action": "<specific rewrite instruction>", "category": "<which score it fixes>"}},
    {{"number": 5, "action": "<specific rewrite instruction>", "category": "<which score it fixes>"}}
  ]
}}

SCORING RUBRICS:
- Content Quality (25): accuracy, depth, UAE real estate relevance, buyer-benefit framing not just facts
- SEO & Structure (20): headings, keyword density, internal {platform} links, no raw Sources sections
- Language & Grammar (20): clean grammar, no developer shorthand like G+1 for general readers
- Brand Voice (15): {platform} advisory tone, not developer marketing copy
- Readability (10): flow, sentence variety, scannability
- Originality (10): unique angle, not copy-pasted from developer brochures

Suggestions must be specific and actionable — tell the writer exactly what to rewrite and how."""

    raw   = call_ai(prompt)
    clean = re.sub(r"```json|```", "", raw).strip()
    match = re.search(r'\{.*\}', clean, re.DOTALL)
    if match:
        clean = match.group(0)
    return json.loads(clean)


# ── plagiarism heuristic ───────────────────────────────────────────────────
def check_plagiarism(text, links):
    known = ["emaar.com","nakheel.com","damac.com","aldar.com","meraas.com","sobha.com",
             "omniyat.com","ellington.ae","azizi.ae","reportage.ae"]
    flagged = [l for l in links if any(d in l for d in known)]
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


# ── AI detection heuristic ─────────────────────────────────────────────────
def check_ai(text):
    phrases = [
        "in conclusion", "it is worth noting", "it is important to note",
        "delve into", "in the realm of", "furthermore", "moreover",
        "needless to say", "leverage", "utilize", "seamlessly",
        "it goes without saying", "in today's", "one such", "robust",
        "cutting-edge", "state-of-the-art", "at the end of the day",
    ]
    hits = sum(1 for p in phrases if p in text.lower())
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
                "Date", "Platform", "Writer", "Title", "Type", "Language",
                "Base Score", "Comments Deduction", "Plagiarism Deduction",
                "AI Deduction", "Final Score", "Plagiarism%", "AI%",
                "Recommendation", "Editor Decision", "Notes"
            ])
        d = row.get("deductions", {})
        sheet.append_row([
            row.get("date"), row.get("platform"), row.get("writer"),
            row.get("title"), row.get("content_type"), row.get("language"),
            d.get("base_score", 0), d.get("comment_deduction", 0),
            d.get("plag_deduction", 0), d.get("ai_deduction", 0),
            d.get("final_score", 0), row.get("plagiarism_pct", 0),
            row.get("ai_pct", 0), row.get("recommendation", ""),
            row.get("editor_decision", ""), row.get("editor_notes", ""),
        ])
    except Exception as e:
        st.warning(f"Could not log to Google Sheets: {e}")


# ── sidebar ────────────────────────────────────────────────────────────────
def sidebar():
    with st.sidebar:
        st.markdown("## Content QA System")
        st.markdown("Bayut and Dubizzle")
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
| Rule | Deduction |
|---|---|
| Per editor comment | - 1 pt |
| Plagiarism over 20% | - 5 pts |
| AI content over 20% | - 5 pts |
        """)
        return page


# ── submit page ────────────────────────────────────────────────────────────
def page_submit():
    inject_css()
    st.markdown(
        '<div class="qa-header"><h1>Content QA System</h1>'
        '<p>Bayut and Dubizzle — Writer evaluation platform</p></div>',
        unsafe_allow_html=True)

    with st.form("qa_form"):
        c1, c2 = st.columns(2)
        writer = c1.text_input("Writer name", placeholder="e.g. Sarah Ahmed")
        title  = c2.text_input("Article title", placeholder="e.g. Everything About Montura 2")

        c3, c4 = st.columns(2)
        ctype  = c3.selectbox("Content type", CONTENT_TYPES)
        lang   = c4.selectbox("Language", LANGUAGES)

        st.markdown("**Platform**")
        pc1, pc2 = st.columns(2)
        with pc1:
            bay_sel = st.form_submit_button("Select Bayut", use_container_width=True)
        with pc2:
            dub_sel = st.form_submit_button("Select Dubizzle", use_container_width=True)

        platform = st.selectbox(
            "Selected platform",
            PLATFORMS,
            label_visibility="collapsed",
        )

        upload = st.file_uploader(
            "Upload article file",
            type=["docx", "pdf", "txt"],
            help=".docx recommended — headings, links and editor comments are extracted automatically",
        )
        go = st.form_submit_button(
            "Run full evaluation",
            use_container_width=True,
            type="primary",
        )

    if not go:
        st.info(
            "Upload a .docx file for best results. "
            "The system reads all headings, links and editor comments written in the document automatically."
        )
        return

    if not writer or not title or not upload:
        st.error("Please fill in writer name, title, and upload a file.")
        return

    with st.spinner("Reading file..."):
        parsed = parse_file(upload)

    if not parsed["text"] or len(parsed["text"]) < 30:
        st.error(f"Could not read text from the file. {parsed.get('error', '')}")
        return

    # Show extracted content
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
                    f'<div class="cmt-author">{c["author"]}</div>'
                    f'{c["text"]}'
                    f'<div class="cmt-deduct">- 1 point will be deducted</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
            if not parsed["comments"]:
                st.caption("No comments found in file")

    prog = st.progress(0, text="Starting evaluation...")

    try:
        prog.progress(20, text="Running AI quality evaluation...")
        qa = run_qa(
            title, parsed["text"], writer, ctype, lang, platform,
            parsed["headings"], parsed["links"], parsed["comments"],
        )
    except Exception as e:
        st.error(f"AI evaluation failed: {e}")
        st.info("Make sure GROQ_API_KEY is set in Streamlit Settings and Secrets.")
        return

    prog.progress(65, text="Running plagiarism check...")
    plag = check_plagiarism(parsed["text"], parsed["links"])

    prog.progress(82, text="Running AI detection...")
    ai = check_ai(parsed["text"])

    prog.progress(95, text="Calculating final score...")
    base_score = qa.get("total", 0)
    final_score, deductions = apply_deductions(
        base_score, parsed["comments"], plag["percentage"], ai["ai_pct"]
    )
    recommendation = get_recommendation(final_score)

    prog.progress(100, text="Done!")
    prog.empty()

    sub = {
        "date":             datetime.now().strftime("%d %b %Y %H:%M"),
        "platform":         platform,
        "writer":           writer,
        "title":            title,
        "content_type":     ctype,
        "language":         lang,
        "word_count":       parsed["word_count"],
        "headings":         parsed["headings"],
        "links":            parsed["links"],
        "comments":         parsed["comments"],
        "qa":               qa,
        "plagiarism":       plag,
        "ai_detection":     ai,
        "deductions":       deductions,
        "qa_score":         final_score,
        "plagiarism_pct":   plag["percentage"],
        "ai_pct":           ai["ai_pct"],
        "recommendation":   recommendation,
        "editor_decision":  "",
        "editor_notes":     "",
    }

    if "submissions" not in st.session_state:
        st.session_state.submissions = []
    st.session_state.submissions.append(sub)
    render_report(sub)


# ── report renderer ────────────────────────────────────────────────────────
def render_report(sub):
    inject_css()
    qa   = sub["qa"]
    plag = sub["plagiarism"]
    ai   = sub["ai_detection"]
    ded  = sub["deductions"]
    score = sub["qa_score"]
    grade, grade_color = get_grade(score)
    rec   = sub["recommendation"]

    st.divider()

    # platform badge
    plat = sub["platform"]
    bdg_class = "bdg-bay" if plat == "Bayut" else "bdg-dub"
    plat_html = f'<span class="bdg {bdg_class}">{plat}</span>'

    st.markdown(
        f"**{sub['writer']}** &nbsp; {plat_html} &nbsp; "
        f"`{sub['content_type']}` &nbsp; `{sub['language']}` &nbsp; "
        f"`{sub['word_count']} words` &nbsp; `{sub['date']}`",
        unsafe_allow_html=True,
    )
    st.markdown("")

    # score hero with breakdown
    rec_labels = {
        "approve": ("Approve",          "#d1fae5", "#065f46"),
        "revise":  ("Request revision", "#fef3c7", "#92400e"),
        "reject":  ("Reject",           "#fee2e2", "#991b1b"),
    }
    rl, rbg, rtc = rec_labels.get(rec, rec_labels["revise"])

    breakdown_rows = (
        f'<div class="base-row"><span>AI base score</span><span>{ded["base_score"]} / 100</span></div>'
    )
    if ded["comment_deduction"] > 0:
        breakdown_rows += (
            f'<div class="ded-row"><span>Editor comments ({ded["comment_count"]} comments)</span>'
            f'<span>- {ded["comment_deduction"]} pts</span></div>'
        )
    else:
        breakdown_rows += '<div class="ok-row"><span>Editor comments</span><span>no deduction</span></div>'

    if ded["plag_deduction"] > 0:
        breakdown_rows += (
            f'<div class="ded-row"><span>Plagiarism {ded["plag_pct"]}% (over 20%)</span>'
            f'<span>- {ded["plag_deduction"]} pts</span></div>'
        )
    else:
        breakdown_rows += (
            f'<div class="ok-row"><span>Plagiarism {ded["plag_pct"]}% (under 20%)</span>'
            f'<span>no deduction</span></div>'
        )

    if ded["ai_deduction"] > 0:
        breakdown_rows += (
            f'<div class="ded-row"><span>AI content {ded["ai_pct"]}% (over 20%)</span>'
            f'<span>- {ded["ai_deduction"]} pts</span></div>'
        )
    else:
        breakdown_rows += (
            f'<div class="ok-row"><span>AI content {ded["ai_pct"]}% (under 20%)</span>'
            f'<span>no deduction</span></div>'
        )

    breakdown_rows += (
        f'<div class="total-row"><span>Final score</span><span>{score} / 100</span></div>'
    )

    st.markdown(
        f'<div class="score-hero">'
        f'<div class="score-num">{score}<span class="score-den"> / 100</span></div>'
        f'<div class="score-grade">{grade}</div>'
        f'<div style="display:inline-block;margin:6px 0 8px;padding:3px 12px;border-radius:20px;'
        f'background:{rbg};color:{rtc};font-size:11px;font-weight:500">{rl}</div>'
        f'<div class="score-verdict">{qa.get("overall_feedback", "")}</div>'
        f'<div class="breakdown-box">{breakdown_rows}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # top metrics
    m1, m2, m3 = st.columns(3)
    status_label = {"safe": "Safe", "warn": "Warning", "danger": "High"}
    m1.metric("Final score",           f"{score} / 100")
    m2.metric("Plagiarism",            f"{plag['percentage']}%",
              delta=f"Over 20% threshold — 5 pts deducted" if ded["plag_deduction"] else None,
              delta_color="inverse")
    m3.metric("AI content",            f"{ai['ai_pct']}%",
              delta=f"Over 20% threshold — 5 pts deducted" if ded["ai_deduction"] else None,
              delta_color="inverse")

    st.divider()

    # category scores
    st.markdown("#### Category scores")
    for cat, mx in CAT_MAX.items():
        data = qa["scores"].get(cat, {})
        s    = data.get("score", 0)
        ca, cb = st.columns([4, 1])
        ca.markdown(f"**{cat}**")
        ca.progress(s / mx)
        cb.markdown(f"**{s} / {mx}**")
        st.caption(data.get("feedback", ""))
        st.markdown("")

    st.divider()

    # automated checks
    st.markdown("#### Automated checks")
    pc1, pc2 = st.columns(2)

    with pc1:
        plag_status = "Over 20% — 5 pts deducted" if ded["plag_deduction"] else "Under 20% — no deduction"
        st.markdown("**Plagiarism**")
        st.metric("Matched content", f"{plag['percentage']}%")
        st.progress(min(plag["percentage"] / 100, 1.0))
        st.caption(plag_status)
        if plag["flagged_sources"]:
            st.markdown("Flagged sources:")
            for src in plag["flagged_sources"]:
                st.markdown(f"- `{src}`")

    with pc2:
        ai_status = "Over 20% — 5 pts deducted" if ded["ai_deduction"] else "Under 20% — no deduction"
        st.markdown("**AI detection**")
        st.metric("AI-generated estimate", f"{ai['ai_pct']}%")
        st.progress(min(ai["ai_pct"] / 100, 1.0))
        st.caption(ai_status)
        ac1, ac2 = st.columns(2)
        ac1.metric("Human", f"{ai['human_pct']}%")
        ac2.metric("AI",    f"{ai['ai_pct']}%")

    st.divider()

    # structure
    st.markdown("#### Document structure")
    sc1, sc2, sc3 = st.columns(3)
    sc1.metric("Headings", len(sub["headings"]))
    sc2.metric("Total links", len(sub["links"]))
    internal = [l for l in sub["links"] if sub["platform"].lower() in l.lower()]
    sc3.metric("Internal links", len(internal))

    # editor comments
    if sub["comments"]:
        st.markdown("**Editor comments from file** — 1 point deducted per comment")
        for c in sub["comments"]:
            st.markdown(
                f'<div class="cmt-card">'
                f'<span class="cmt-author">{c["author"]}:</span> {c["text"]}'
                f'<div class="cmt-deduct">- 1 point deducted</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    st.divider()

    # strengths and improvements
    col_s, col_i = st.columns(2)
    with col_s:
        st.markdown("#### Strengths")
        for s in qa.get("key_strengths", []):
            st.markdown(f'<span class="tag-str">{s}</span>', unsafe_allow_html=True)
    with col_i:
        st.markdown("#### Required improvements")
        for imp in qa.get("areas_for_improvement", []):
            st.markdown(f'<span class="tag-imp">{imp}</span>', unsafe_allow_html=True)

    st.divider()

    # suggestions
    suggestions = qa.get("suggestions", [])
    if suggestions:
        st.markdown("#### Suggestions to improve the article")
        st.caption("Specific actions the writer should take to raise the score")
        for sug in suggestions:
            st.markdown(
                f'<div class="suggest-item">'
                f'<div class="suggest-num">{sug.get("number", "")}</div>'
                f'<div>'
                f'<div>{sug.get("action", "")}</div>'
                f'<div class="suggest-cat">Addresses: {sug.get("category", "")}</div>'
                f'</div></div>',
                unsafe_allow_html=True,
            )

    st.divider()

    # editor decision
    st.markdown("#### Editor decision")
    st.caption("The AI recommendation is a guide. You make the final call.")

    rec_idx = {"approve": 0, "revise": 1, "reject": 2}
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
    if st.button(
        "Confirm decision",
        type="primary",
        use_container_width=True,
        key=f"conf_{sub['title']}_{sub['date']}",
    ):
        if decision in ("Request revision", "Reject") and not notes.strip():
            st.error("Please add notes for the writer before confirming.")
        else:
            sub["editor_decision"] = decision
            sub["editor_notes"]    = notes
            log_to_sheets(sub)
            st.success(f"Decision saved: {decision}")
            if notes:
                st.info(f"Notes for {sub['writer']}: {notes}")

    st.caption(
        f"Generated by Content QA System · {sub['platform']} · "
        f"Powered by Groq (Llama 3) · {sub['date']}"
    )


# ── dashboard ──────────────────────────────────────────────────────────────
def page_dashboard():
    inject_css()
    st.markdown(
        '<div class="qa-header"><h1>Dashboard</h1>'
        '<p>All submissions across Bayut and Dubizzle</p></div>',
        unsafe_allow_html=True,
    )

    all_subs = st.session_state.get("submissions", [])
    if not all_subs:
        st.info("No submissions yet. Go to Submit article to start.")
        return

    # summary metrics — all platforms
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

    # filters
    f1, f2, f3, f4 = st.columns(4)
    plat_filter   = f1.selectbox("Platform",     ["All"] + PLATFORMS)
    type_filter   = f2.selectbox("Content type", ["All"] + CONTENT_TYPES)
    lang_filter   = f3.selectbox("Language",     ["All"] + LANGUAGES)
    status_filter = f4.selectbox("Status",       ["All", "Pending", "Approve", "Request revision", "Reject"])

    filtered = all_subs
    if plat_filter   != "All": filtered = [s for s in filtered if s["platform"]      == plat_filter]
    if type_filter   != "All": filtered = [s for s in filtered if s["content_type"]  == type_filter]
    if lang_filter   != "All": filtered = [s for s in filtered if s["language"]      == lang_filter]
    if status_filter != "All":
        if status_filter == "Pending":
            filtered = [s for s in filtered if not s.get("editor_decision")]
        else:
            filtered = [s for s in filtered if s.get("editor_decision") == status_filter]

    st.markdown(f"**{len(filtered)} submissions**")

    for sub in reversed(filtered):
        grade, _ = get_grade(sub["qa_score"])
        plat     = sub["platform"]
        bdg      = f'<span class="bdg {"bdg-bay" if plat=="Bayut" else "bdg-dub"}">{plat}</span>'
        dec      = sub.get("editor_decision") or "Pending"
        exp_label = (
            f"{sub['writer']} — {sub['title'][:50]}{'...' if len(sub['title'])>50 else ''} "
            f"| Score: {sub['qa_score']} | {dec} | {sub['date']}"
        )
        with st.expander(exp_label):
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
