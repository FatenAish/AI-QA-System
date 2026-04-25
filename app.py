import streamlit as st
import json
import re
from datetime import datetime
from io import BytesIO

# ── optional imports ───────────────────────────────────────────────────────
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

# ── page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Content QA | Bayut & Dubizzle",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── brand config ───────────────────────────────────────────────────────────
BRANDS = {
    "Bayut":     {"primary": "#e2231a", "light": "#fff0ef", "logo": "🏠", "url": "bayut.com"},
    "Dubizzle":  {"primary": "#00a699", "light": "#e0f5f4", "logo": "🏢", "url": "dubizzle.com"},
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
    (90, "A — Excellent",          "🟢"),
    (80, "B — Good",               "🟢"),
    (70, "C — Needs revision",     "🟡"),
    (60, "D — Major revision",     "🟠"),
    (0,  "F — Reject",             "🔴"),
]

# ── CSS ────────────────────────────────────────────────────────────────────
def inject_css(primary, light):
    st.markdown(f"""
    <style>
    :root {{ --brand:{primary}; --light:{light}; }}
    .brand-hdr {{ background:var(--brand);color:#fff;padding:16px 20px;border-radius:10px;margin-bottom:1.2rem; }}
    .brand-hdr h1 {{ font-size:20px;font-weight:600;margin:0 0 3px; }}
    .brand-hdr p  {{ font-size:12px;opacity:.85;margin:0; }}
    .score-card {{ background:var(--light);border:1.5px solid var(--brand);border-radius:10px;padding:16px 20px;margin-bottom:1rem; }}
    .score-big  {{ font-size:52px;font-weight:700;color:var(--brand);line-height:1; }}
    .grade-txt  {{ font-size:14px;font-weight:600;margin-top:4px; }}
    .verdict    {{ font-size:13px;color:#444;line-height:1.6;margin-top:8px; }}
    .tag-str {{ background:#d1fae5;color:#065f46;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:500;display:inline-block;margin:2px; }}
    .tag-imp {{ background:#fef3c7;color:#92400e;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:500;display:inline-block;margin:2px; }}
    .cmt-box {{ background:#fff8f0;border-left:3px solid var(--brand);padding:8px 12px;margin-bottom:6px;border-radius:0 6px 6px 0;font-size:12px; }}
    .cmt-auth {{ font-weight:600;color:var(--brand); }}
    </style>
    """, unsafe_allow_html=True)

# ── AI call — Groq only, 100% free ────────────────────────────────────────
def call_ai(prompt: str) -> str:
    if not GROQ_OK:
        raise Exception("groq package not installed — check requirements.txt")

    client = Groq(api_key=st.secrets["GROQ_API_KEY"])

    # try models in order — all free on Groq
    for model in [
        "llama-3.1-8b-instant",
        "llama3-8b-8192",
        "gemma2-9b-it",
        "mixtral-8x7b-32768",
    ]:
        try:
            resp = client.chat.completions.create(
                model       = model,
                messages    = [{"role": "user", "content": prompt}],
                temperature = 0.3,
                max_tokens  = 1500,
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            err = str(e).lower()
            if "model" in err or "not found" in err or "decommission" in err:
                continue
            raise e

    raise Exception("All Groq models failed. Check your GROQ_API_KEY in Streamlit Secrets.")

# ── file parsers ───────────────────────────────────────────────────────────
def extract_docx(raw: bytes) -> dict:
    if not DOCX_OK:
        return {"text": "", "headings": [], "links": [], "comments": [], "word_count": 0,
                "error": "python-docx not installed"}
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

def extract_pdf(raw: bytes) -> dict:
    if not PDF_OK:
        return {"text": "", "headings": [], "links": [], "comments": [], "word_count": 0,
                "error": "pdfplumber not installed"}
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

def extract_txt(raw: bytes) -> dict:
    full = raw.decode("utf-8", errors="ignore")
    return {"text": full, "headings": [], "links": re.findall(r'https?://\S+', full),
            "comments": [], "word_count": len(full.split()), "error": ""}

def parse_file(f) -> dict:
    raw  = f.getvalue()
    name = f.name.lower()
    if   name.endswith(".docx"): return extract_docx(raw)
    elif name.endswith(".pdf"):  return extract_pdf(raw)
    else:                        return extract_txt(raw)

# ── QA evaluation ──────────────────────────────────────────────────────────
def run_qa(title, content, writer, ctype, lang, brand, headings, links, comments) -> dict:
    h_txt = "\n".join(f"  [{h['level']}] {h['text']}" for h in headings) or "  None"
    l_txt = "\n".join(f"  - {l}" for l in links[:8])                      or "  None"
    c_txt = "\n".join(f"  [{c['author']}]: {c['text']}" for c in comments) or "  None"

    prompt = f"""You are a senior content QA evaluator for {brand}, a leading UAE real estate platform.
Evaluate this {ctype.lower()} written in {lang}.

TITLE: {title}
WRITER: {writer}
HEADINGS: {h_txt}
LINKS: {l_txt}
EDITOR COMMENTS: {c_txt}

CONTENT:
{content[:4000]}

Return ONLY valid JSON — no markdown, no text outside the JSON:
{{
  "scores": {{
    "Content Quality":    {{"score": 0-25, "feedback": "2 sentences"}},
    "SEO & Structure":    {{"score": 0-20, "feedback": "2 sentences"}},
    "Language & Grammar": {{"score": 0-20, "feedback": "2 sentences"}},
    "Brand Voice":        {{"score": 0-15, "feedback": "2 sentences"}},
    "Readability & Flow": {{"score": 0-10, "feedback": "2 sentences"}},
    "Originality":        {{"score": 0-10, "feedback": "2 sentences"}}
  }},
  "total": <sum>,
  "overall_feedback": "3 sentence summary",
  "key_strengths": ["s1","s2","s3"],
  "areas_for_improvement": ["a1","a2","a3","a4"],
  "recommendation": "approve or revise or reject"
}}

RUBRICS:
- Content Quality(25): accuracy, depth, UAE real estate relevance, buyer framing
- SEO & Structure(20): headings, keywords, internal links, no raw Sources sections
- Language & Grammar(20): grammar, no developer shorthand like G+1 for general readers
- Brand Voice(15): {brand} advisory tone — not developer marketing copy
- Readability(10): flow, sentence variety, scannability
- Originality(10): unique angle, not copy-pasted from brochures

approve>=80, revise 60-79, reject<60. Factor editor comments into scores."""

    raw   = call_ai(prompt)
    clean = re.sub(r"```json|```", "", raw).strip()
    # sometimes the model wraps in extra text — extract just the JSON
    match = re.search(r'\{.*\}', clean, re.DOTALL)
    if match:
        clean = match.group(0)
    return json.loads(clean)

# ── plagiarism heuristic ───────────────────────────────────────────────────
def check_plagiarism(text, links):
    known = ["emaar.com","nakheel.com","damac.com","aldar.com","meraas.com","sobha.com"]
    flagged = [l for l in links if any(d in l for d in known)]
    words  = text.lower().split()
    chunks = [" ".join(words[i:i+8]) for i in range(0, max(len(words)-8,1), 4)]
    seen, dups = set(), 0
    for c in chunks:
        if c in seen: dups += 1
        seen.add(c)
    pct = min(int((dups / max(len(chunks),1)) * 100) + len(flagged)*8, 100)
    return {"percentage": pct, "flagged_sources": flagged,
            "status": "danger" if pct>20 else "warn" if pct>10 else "safe"}

# ── AI detection heuristic ─────────────────────────────────────────────────
def check_ai(text):
    phrases = ["in conclusion","it is worth noting","it is important to note","delve into",
               "in the realm of","furthermore","moreover","needless to say","leverage",
               "utilize","seamlessly","it goes without saying","in today's"]
    hits = sum(1 for p in phrases if p in text.lower())
    pct  = min(hits * 6, 60)
    return {"ai_pct": pct, "human_pct": 100-pct,
            "status": "danger" if pct>30 else "warn" if pct>15 else "safe"}

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
        sheet = gc.open(st.secrets.get("SHEET_NAME","Bayut QA Submissions")).sheet1
        if not sheet.row_values(1):
            sheet.append_row(["Date","Brand","Writer","Title","Type","Language",
                              "QA Score","Plagiarism%","AI%","Recommendation","Decision","Notes"])
        sheet.append_row([row.get("date"), row.get("brand"), row.get("writer"),
                          row.get("title"), row.get("content_type"), row.get("language"),
                          row.get("qa_score"), row.get("plagiarism_pct"), row.get("ai_pct"),
                          row.get("recommendation"), row.get("editor_decision",""),
                          row.get("editor_notes","")])
    except Exception as e:
        st.warning(f"Sheets log failed: {e}")

# ── grade helper ───────────────────────────────────────────────────────────
def get_grade(score):
    for t, l, i in GRADE_MAP:
        if score >= t:
            return l, i
    return GRADE_MAP[-1][1], GRADE_MAP[-1][2]

# ── sidebar ────────────────────────────────────────────────────────────────
def sidebar():
    with st.sidebar:
        st.markdown("## ⚙️ Settings")
        brand = st.radio("Platform", list(BRANDS.keys()), horizontal=True)
        st.divider()
        st.markdown("### 📋 Navigation")
        page  = st.radio("Go to", ["📤 Submit article","📊 Dashboard"],
                         label_visibility="collapsed")
        st.divider()
        st.markdown("### ℹ️ Score thresholds")
        st.markdown("""
| Score | Grade |
|---|---|
| 90–100 | ✅ Approve |
| 80–89  | ✅ Approve |
| 70–79  | 🟡 Revise  |
| 60–69  | 🟠 Revise  |
| < 60   | 🔴 Reject  |
        """)
        return brand, page

# ── submit page ────────────────────────────────────────────────────────────
def page_submit(brand):
    b = BRANDS[brand]
    inject_css(b["primary"], b["light"])
    st.markdown(
        f'<div class="brand-hdr"><h1>{b["logo"]} {brand} — Content QA</h1>'
        f'<p>Upload article · AI evaluation · Plagiarism & AI detection · Editor approval</p></div>',
        unsafe_allow_html=True)

    with st.form("qa_form"):
        c1, c2 = st.columns(2)
        writer = c1.text_input("Writer name *", placeholder="e.g. Rabia")
        title  = c2.text_input("Article title *", placeholder="e.g. Everything About Montura 2")
        c3, c4 = st.columns(2)
        ctype  = c3.selectbox("Content type", CONTENT_TYPES)
        lang   = c4.selectbox("Language", LANGUAGES)
        upload = st.file_uploader("Upload article file *", type=["docx","pdf","txt"],
                    help=".docx recommended — extracts headings, links and editor comments automatically")
        go     = st.form_submit_button("🚀 Run full evaluation", use_container_width=True, type="primary")

    if not go:
        st.info("📎 Upload a .docx file for best results — headings, links and editor comments are extracted automatically.")
        return

    if not writer or not title or not upload:
        st.error("Please fill in writer name, title and upload a file.")
        return

    with st.spinner("Reading file…"):
        parsed = parse_file(upload)

    if not parsed["text"] or len(parsed["text"]) < 30:
        st.error(f"Could not read text from file. {parsed.get('error','')}")
        return

    # show extracted metadata
    with st.expander(f"📄 Extracted — {len(parsed['headings'])} headings · {len(parsed['links'])} links · {len(parsed['comments'])} comments"):
        ec1, ec2, ec3 = st.columns(3)
        with ec1:
            st.markdown("**Headings**")
            for h in parsed["headings"]: st.markdown(f"`{h['level']}` {h['text']}")
            if not parsed["headings"]: st.caption("None detected")
        with ec2:
            st.markdown("**Links**")
            for l in parsed["links"][:6]: st.markdown(f"→ {l}")
            if not parsed["links"]: st.caption("None detected")
        with ec3:
            st.markdown("**Editor comments**")
            for c in parsed["comments"]:
                st.markdown(f'<div class="cmt-box"><span class="cmt-auth">{c["author"]}:</span> {c["text"]}</div>',
                            unsafe_allow_html=True)
            if not parsed["comments"]: st.caption("None found")

    prog = st.progress(0, text="Starting…")

    try:
        prog.progress(20, text="Running AI quality evaluation…")
        qa = run_qa(title, parsed["text"], writer, ctype, lang, brand,
                    parsed["headings"], parsed["links"], parsed["comments"])
    except Exception as e:
        st.error(f"AI evaluation failed: {e}")
        st.info("Make sure GROQ_API_KEY is set in Streamlit → Settings → Secrets")
        return

    prog.progress(65, text="Running plagiarism check…")
    plag = check_plagiarism(parsed["text"], parsed["links"])
    prog.progress(85, text="Running AI detection…")
    ai   = check_ai(parsed["text"])
    prog.progress(100, text="Done!")
    prog.empty()

    sub = {
        "date": datetime.now().strftime("%d %b %Y %H:%M"),
        "brand": brand, "writer": writer, "title": title,
        "content_type": ctype, "language": lang,
        "word_count": parsed["word_count"],
        "headings": parsed["headings"], "links": parsed["links"],
        "comments": parsed["comments"],
        "qa": qa, "plagiarism": plag, "ai_detection": ai,
        "qa_score": qa.get("total", 0),
        "plagiarism_pct": plag["percentage"],
        "ai_pct": ai["ai_pct"],
        "recommendation": qa.get("recommendation","revise"),
        "editor_decision": "", "editor_notes": "",
    }

    if "submissions" not in st.session_state:
        st.session_state.submissions = []
    st.session_state.submissions.append(sub)
    render_report(sub)

# ── report renderer ────────────────────────────────────────────────────────
def render_report(sub):
    b     = BRANDS[sub["brand"]]
    qa    = sub["qa"]
    plag  = sub["plagiarism"]
    ai    = sub["ai_detection"]
    score = sub["qa_score"]
    grade, gicon = get_grade(score)

    inject_css(b["primary"], b["light"])
    st.divider()
    st.markdown(f"## 📋 QA Report — {sub['writer']} · {sub['title']}")
    st.markdown(f"`{sub['brand']}` &nbsp; `{sub['content_type']}` &nbsp; `{sub['language']}` &nbsp; `~{sub['word_count']} words` &nbsp; `{sub['date']}`",
                unsafe_allow_html=True)

    # score hero
    st.markdown(
        f'<div class="score-card">'
        f'<div class="score-big">{score}<span style="font-size:20px;font-weight:400;color:#888"> / 100</span></div>'
        f'<div class="grade-txt">{gicon} {grade}</div>'
        f'<div class="verdict">{qa.get("overall_feedback","")}</div>'
        f'</div>', unsafe_allow_html=True)

    m1, m2, m3 = st.columns(3)
    icons = {"safe":"🟢","warn":"🟡","danger":"🔴"}
    m1.metric("QA Score",            f"{score} / 100")
    m2.metric("Plagiarism detected", f"{icons[plag['status']]} {plag['percentage']}%")
    m3.metric("AI-generated",        f"{icons[ai['status']]} {ai['ai_pct']}%")

    st.divider()
    st.markdown("### 📊 Category scores")
    for cat, mx in CAT_MAX.items():
        data = qa["scores"].get(cat, {})
        s    = data.get("score", 0)
        ca, cb = st.columns([3,1])
        ca.markdown(f"**{cat}**")
        ca.progress(s / mx)
        cb.markdown(f"**{s} / {mx}**")
        st.caption(data.get("feedback",""))
        st.markdown("")

    st.divider()
    st.markdown("### 🔍 Automated checks")
    pc1, pc2 = st.columns(2)
    with pc1:
        st.markdown(f"#### {icons[plag['status']]} Plagiarism")
        st.metric("Matched content", f"{plag['percentage']}%")
        st.progress(plag["percentage"] / 100)
        if plag["flagged_sources"]:
            for s in plag["flagged_sources"]: st.markdown(f"- `{s}`")
    with pc2:
        st.markdown(f"#### {icons[ai['status']]} AI detection")
        st.metric("AI-generated estimate", f"{ai['ai_pct']}%")
        st.progress(ai["ai_pct"] / 100)
        ac1, ac2 = st.columns(2)
        ac1.metric("Human", f"{ai['human_pct']}%")
        ac2.metric("AI",    f"{ai['ai_pct']}%")

    st.divider()
    st.markdown("### 🗂 Structure")
    sc1, sc2, sc3 = st.columns(3)
    sc1.metric("Headings", len(sub["headings"]))
    sc2.metric("Links",    len(sub["links"]))
    internal = [l for l in sub["links"] if sub["brand"].lower() in l.lower()]
    sc3.metric("Internal links", len(internal))

    if sub["comments"]:
        st.markdown("**Editor comments from file:**")
        for c in sub["comments"]:
            st.markdown(f'<div class="cmt-box"><span class="cmt-auth">{c["author"]}:</span> {c["text"]}</div>',
                        unsafe_allow_html=True)

    st.divider()
    col_s, col_i = st.columns(2)
    with col_s:
        st.markdown("### ✅ Strengths")
        for s in qa.get("key_strengths", []):
            st.markdown(f'<span class="tag-str">{s}</span>', unsafe_allow_html=True)
    with col_i:
        st.markdown("### 🔧 Improvements")
        for imp in qa.get("areas_for_improvement", []):
            st.markdown(f'<span class="tag-imp">{imp}</span>', unsafe_allow_html=True)

    st.divider()
    st.markdown("### ✍️ Editor decision")
    rec_map = {"approve":0, "revise":1, "reject":2}
    decision = st.radio("Decision",
        ["✅ Approve","↩ Request revision","🚨 Reject"],
        index=rec_map.get(sub["recommendation"],1),
        horizontal=True, key=f"dec_{sub['title']}_{sub['date']}")
    notes = st.text_area("Notes for writer (required for revision/rejection)", height=90,
                         placeholder="Tell the writer exactly what to fix.",
                         key=f"notes_{sub['title']}_{sub['date']}")
    if st.button("✅ Confirm & save", type="primary", use_container_width=True,
                 key=f"conf_{sub['title']}_{sub['date']}"):
        if ("revision" in decision or "Reject" in decision) and not notes.strip():
            st.error("Please add notes for the writer.")
        else:
            sub["editor_decision"] = decision
            sub["editor_notes"]    = notes
            log_to_sheets(sub)
            st.success(f"Decision saved: **{decision}**")
            if notes:
                st.info(f"📬 Notes for {sub['writer']}: {notes}")

    st.caption(f"Generated by {sub['brand']} Content QA · Powered by Groq (Llama 3) · {sub['date']}")

# ── dashboard ──────────────────────────────────────────────────────────────
def page_dashboard(brand):
    b = BRANDS[brand]
    inject_css(b["primary"], b["light"])
    st.markdown(
        f'<div class="brand-hdr"><h1>{b["logo"]} {brand} — Dashboard</h1>'
        f'<p>All writer evaluations and decisions</p></div>',
        unsafe_allow_html=True)

    subs = [s for s in st.session_state.get("submissions",[]) if s["brand"]==brand]
    if not subs:
        st.info("No submissions yet. Go to **Submit article** to start.")
        return

    approved = sum(1 for s in subs if "Approve" in s.get("editor_decision",""))
    revision = sum(1 for s in subs if "revision" in s.get("editor_decision",""))
    rejected = sum(1 for s in subs if "Reject" in s.get("editor_decision",""))
    pending  = sum(1 for s in subs if not s.get("editor_decision"))

    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric("Total",    len(subs))
    m2.metric("Approved", approved)
    m3.metric("Revision", revision)
    m4.metric("Rejected", rejected)
    m5.metric("Pending",  pending)

    st.divider()
    f1,f2,f3 = st.columns(3)
    tf = f1.selectbox("Type",   ["All"]+CONTENT_TYPES)
    lf = f2.selectbox("Language",["All"]+LANGUAGES)
    df = f3.selectbox("Status", ["All","Pending","Approved","Revision","Rejected"])

    filtered = subs
    if tf != "All": filtered = [s for s in filtered if s["content_type"]==tf]
    if lf != "All": filtered = [s for s in filtered if s["language"]==lf]
    if df != "All":
        if df == "Pending":
            filtered = [s for s in filtered if not s.get("editor_decision")]
        else:
            filtered = [s for s in filtered if df.lower() in s.get("editor_decision","").lower()]

    for sub in reversed(filtered):
        g, icon = get_grade(sub["qa_score"])
        with st.expander(f"{icon} **{sub['writer']}** · {sub['title'][:55]} · Score: **{sub['qa_score']}** · {sub['date']}"):
            render_report(sub)

# ── main ───────────────────────────────────────────────────────────────────
def main():
    if "submissions" not in st.session_state:
        st.session_state.submissions = []
    brand, page = sidebar()
    if "Submit" in page:
        page_submit(brand)
    else:
        page_dashboard(brand)

if __name__ == "__main__":
    main()
