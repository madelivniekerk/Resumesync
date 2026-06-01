"""
Resume Matcher Web App
Upload resume + paste job URL → Get compatibility score + recommendations
"""

import streamlit as st
import os
import re
import json
from datetime import datetime
from dotenv import load_dotenv
from anthropic import Anthropic
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pypdf
import io
import base64
import openpyxl
from openpyxl import load_workbook
try:
    from supabase import create_client as _supabase_create_client
    _SUPABASE_AVAILABLE = True
except Exception:
    _SUPABASE_AVAILABLE = False

# Load environment
load_dotenv()

# Configuration
MODEL_NAME = "claude-sonnet-4-6"
TRACKER_PATH = os.path.join(os.path.dirname(__file__), "job_applications.xlsx")
COVER_LETTERS_PATH = os.path.join(os.path.dirname(__file__), "cover_letters")

# Page config
st.set_page_config(
    page_title="ResumeSync - AI-Powered Job Matching",
    page_icon="🔄",
    layout="wide"
)

# Custom CSS — VisualizePro · Forest + Sage design system
st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,400;12..96,600;12..96,700;12..96,800&family=Lora:ital,wght@1,500;1,600&family=Space+Mono:wght@400;700&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500&display=swap">
<style>

    /* ── Design tokens ── */
    :root {
        --bg:          #071512;
        --bg-2:        #0a1b16;
        --panel:       #0c2019;
        --panel-2:     #0f271e;
        --ink:         #ecf4ee;
        --ink2:        #9fb6a8;
        --ink3:        #6e8a7b;
        --mint:        #7ad79f;
        --mint-bright: #94e6b1;
        --mint-deep:   #4fae7a;
        --mint-dim:    rgba(122,215,159,0.10);
        --mint-dim2:   rgba(122,215,159,0.16);
        --mint-border: rgba(122,215,159,0.22);
        --amber:       #e0a14a;
        --line:        rgba(159,182,168,0.12);
        --line-2:      rgba(159,182,168,0.07);
        /* aliases used in component CSS */
        --gold:        #7ad79f;
        --gold-soft:   #94e6b1;
        --muted:       #6e8a7b;
        --line-strong: rgba(122,215,159,0.22);
    }

    /* ── App background ── */
    .stApp, .main {
        background: var(--bg) !important;
        background-image:
            radial-gradient(ellipse 80% 60% at 15% 20%, rgba(109,193,138,0.07), transparent 60%),
            radial-gradient(ellipse 60% 50% at 95% 90%, rgba(122,215,159,0.04), transparent 60%) !important;
        font-family: 'DM Sans', system-ui, sans-serif !important;
        color: var(--ink) !important;
    }

    /* ── Sidebar ── */
    section[data-testid="stSidebar"] {
        background: var(--bg-2) !important;
        border-right: 1px solid var(--line) !important;
    }
    section[data-testid="stSidebar"] .stMarkdown { color: var(--ink) !important; }
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 { color: var(--ink) !important; }

    /* ── Typography ── */
    h1, h2, h3 {
        font-family: 'DM Sans', sans-serif !important;
        color: var(--ink) !important;
        font-weight: 700 !important;
    }
    .stMarkdown { font-family: 'DM Sans', sans-serif !important; color: var(--ink) !important; }
    p, li, label { color: var(--ink) !important; }

    /* ── Primary button — sage green fill ── */
    .stButton > button {
        background: var(--gold) !important;
        color: #0a1f17 !important;
        border: none !important;
        border-radius: 4px !important;
        padding: 16px 32px !important;
        font-weight: 600 !important;
        font-size: 15px !important;
        font-family: 'DM Sans', sans-serif !important;
        transition: transform .15s, background .2s !important;
        min-height: 52px !important;
    }
    .stButton > button:hover {
        background: var(--gold-soft) !important;
        transform: translateY(-1px) !important;
        box-shadow: none !important;
    }

    /* ── Download button — ghost style ── */
    .stDownloadButton > button {
        background: transparent !important;
        color: var(--ink) !important;
        border: 1px solid var(--line-strong) !important;
        border-radius: 4px !important;
        padding: 12px 28px !important;
        font-weight: 600 !important;
        font-family: 'DM Sans', sans-serif !important;
        transition: background .2s, color .2s !important;
    }
    .stDownloadButton > button:hover {
        background: var(--gold) !important;
        color: #0a1f17 !important;
        border-color: var(--gold) !important;
    }

    /* ── Form elements ── */
    .stFileUploader, .stTextArea, .stTextInput { font-family: 'DM Sans', sans-serif !important; }

    .stTextInput input, .stTextArea textarea {
        font-family: 'DM Sans', sans-serif !important;
        background: var(--bg-2) !important;
        color: #ffffff !important;
        border-radius: 4px !important;
        border-color: var(--line-strong) !important;
    }
    .stSelectbox [data-baseweb="select"] > div:first-child {
        background: var(--bg-2) !important;
        border-color: var(--line-strong) !important;
    }
    .stSelectbox [data-baseweb="select"] [data-baseweb="value-container"] *,
    .stSelectbox [data-baseweb="select"] [role="combobox"],
    .stSelectbox [data-baseweb="select"] input {
        color: #ffffff !important;
        font-family: 'DM Sans', sans-serif !important;
    }
    .stSelectbox [data-baseweb="select"] svg { fill: var(--muted) !important; }
    [data-baseweb="popover"] [data-baseweb="menu"] {
        background: var(--bg-2) !important;
    }
    [data-baseweb="popover"] [role="option"] {
        background: var(--bg-2) !important;
        color: #ffffff !important;
    }
    [data-baseweb="popover"] [role="option"]:hover,
    [data-baseweb="popover"] [aria-selected="true"] {
        background: rgba(109,193,138,0.15) !important;
        color: var(--gold) !important;
    }
    .stTextInput input::placeholder, .stTextArea textarea::placeholder {
        color: var(--muted) !important;
        opacity: 1 !important;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: var(--gold) !important;
        box-shadow: 0 0 0 1px var(--gold) !important;
    }

    /* ── File uploader — dark theme, consistent with other inputs ── */
    [data-testid="stFileUploaderDropzone"] {
        background: var(--bg-2) !important;
        border: 1px solid var(--line-strong) !important;
        border-radius: 4px !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: var(--gold) !important;
    }
    /* Hide the duplicate upload SVG icon + its label — keep only browse button + file-type hint */
    [data-testid="stFileUploaderDropzone"] svg { display: none !important; }
    [data-testid="stFileUploaderDropzoneInstructions"] > div > div:first-child { display: none !important; }
    [data-testid="stFileUploaderDropzoneInstructions"],
    [data-testid="stFileUploaderDropzoneInstructions"] *,
    [data-testid="stFileUploaderDropzone"] span,
    [data-testid="stFileUploaderDropzone"] p,
    [data-testid="stFileUploaderDropzone"] small {
        color: var(--muted) !important;
        font-family: 'DM Sans', sans-serif !important;
    }
    [data-testid="stFileUploaderDropzone"] button {
        background: rgba(122,215,159,0.10) !important;
        border: 1px solid var(--line-strong) !important;
        border-radius: 4px !important;
        min-width: 120px !important;
    }
    [data-testid="stFileUploaderDropzone"] button * {
        display: none !important;
    }
    [data-testid="stFileUploaderDropzone"] button::after {
        content: "Browse files";
        display: inline !important;
        font-size: 0.875rem !important;
        font-family: 'DM Sans', sans-serif !important;
        color: var(--gold) !important;
    }
    [data-testid="stFileUploaderDropzone"] button:hover {
        background: var(--gold) !important;
    }
    [data-testid="stFileUploaderDropzone"] button:hover::after {
        color: #0a1f17 !important;
    }
    /* Uploaded file pill */
    [data-testid="uploadedFile"] {
        background: rgba(122,215,159,0.07) !important;
        border: 1px solid var(--line) !important;
        border-radius: 4px !important;
    }
    [data-testid="uploadedFile"] *,
    [data-testid="uploadedFileData"] * { color: var(--ink) !important; }
    button[title="Remove file"] { color: var(--muted) !important; }

    /* ── Feedback boxes ── */
    .stSuccess {
        background-color: rgba(122,215,159,0.07) !important;
        color: var(--gold-soft) !important;
        border-left: 4px solid var(--gold) !important;
    }
    .stInfo {
        background-color: rgba(122,215,159,0.04) !important;
        border-left: 4px solid var(--line-strong) !important;
        color: var(--muted) !important;
    }
    .stError { font-family: 'DM Sans', sans-serif !important; }
    .stStatus { font-family: 'DM Sans', sans-serif !important; }

    /* ── Spinner ── */
    .stSpinner > div { border-top-color: var(--gold) !important; }

    /* ── Divider ── */
    hr { border-color: var(--line) !important; opacity: 1 !important; }

    /* ── Block container ── */
    .block-container, .main .block-container { padding-top: 0.75rem !important; }

    /* ── Remove Streamlit's top header bar entirely ── */
    [data-testid="stHeader"] { display: none !important; }
    #MainMenu { visibility: hidden !important; }
    footer { visibility: hidden !important; }

    /* ── Hide scrollbars ── */
    ::-webkit-scrollbar { display: none !important; }
    * { scrollbar-width: none !important; -ms-overflow-style: none !important; }

    /* ── Radio ── */
    .stRadio label { color: var(--muted) !important; }

    /* ── Checkbox ── */
    .stCheckbox label { color: var(--ink) !important; font-family: 'DM Sans', sans-serif !important; }

    /* ── Dataframe ── */
    .stDataFrame { border: 1px solid var(--line) !important; border-radius: 8px !important; }

    /* ── Cover letter white box — force dark ink ── */
    .cl-box, .cl-box * { color: #1b1b1b !important; font-family: 'DM Sans', sans-serif !important; }
</style>
""", unsafe_allow_html=True)


# ============= HELPERS =============

def _read_secret(key: str) -> str:
    """Read a secret: env var → st.secrets → local secrets.toml."""
    # 1. Environment variable
    val = os.getenv(key, "").strip()
    if val:
        return val
    # 2. Streamlit Cloud secrets manager
    try:
        val = str(st.secrets.get(key, "")).strip()
        if val:
            return val
    except Exception:
        pass
    # 3. Local .streamlit/secrets.toml (dev)
    secrets_path = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
    if os.path.exists(secrets_path):
        text = open(secrets_path, encoding="utf-8").read()
        m = re.search(rf'{key}\s*=\s*"([^"]+)"', text)
        if m:
            return m.group(1).strip()
    return ""


def get_client():
    api_key = _read_secret("ANTHROPIC_API_KEY")
    if not api_key or api_key == "your-api-key-here":
        return None
    return Anthropic(api_key=api_key)


@st.cache_resource
def get_supabase():
    if not _SUPABASE_AVAILABLE:
        return None
    url = _read_secret("SUPABASE_URL")
    key = _read_secret("SUPABASE_KEY")
    if url and key:
        try:
            return _supabase_create_client(url, key)
        except Exception:
            return None
    return None


def extract_text_from_pdf(file):
    try:
        pdf_reader = pypdf.PdfReader(file)
        return "".join(page.extract_text() for page in pdf_reader.pages)
    except Exception as e:
        return f"Error reading PDF: {str(e)}"


def extract_text_from_docx(file):
    try:
        doc = Document(file)
        return "\n".join(para.text for para in doc.paragraphs)
    except Exception as e:
        return f"Error reading Word document: {str(e)}"


def extract_text_from_txt(file):
    try:
        return file.read().decode('utf-8')
    except Exception as e:
        return f"Error reading text file: {str(e)}"


def extract_resume_text(uploaded_file):
    if uploaded_file is None:
        return "No file uploaded"
    file_type = uploaded_file.name.split('.')[-1].lower()
    if file_type == 'pdf':
        return extract_text_from_pdf(uploaded_file)
    elif file_type in ['docx', 'doc']:
        return extract_text_from_docx(uploaded_file)
    elif file_type == 'txt':
        return extract_text_from_txt(uploaded_file)
    return "Unsupported file type. Please upload PDF, DOCX, or TXT."



def scrape_job_url(url: str) -> dict:
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9',
            'DNT': '1',
            'Connection': 'keep-alive',
        }
        response = requests.get(url, headers=headers, timeout=15, allow_redirects=True)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        for tag in soup(["script", "style", "nav", "footer", "header"]):
            tag.decompose()
        lines = [l.strip() for l in soup.get_text(separator='\n', strip=True).splitlines() if l.strip()]
        return {'success': True, 'content': '\n'.join(lines)[:8000], 'url': url}
    except Exception as e:
        return {'success': False, 'error': str(e)}


# ============= ANALYSIS =============

def analyze_resume_vs_job(resume_text: str, job_content: str, job_url: str, client):
    """
    ResumeWorded-style deep analysis: keyword matching, hard/soft skills gap,
    section-by-section review, and ATS optimization.
    """
    prompt = f"""You are a professional resume expert and ATS specialist — apply the same rigorous methodology used by top resume review platforms like ResumeWorded.

**MY RESUME:**
{resume_text}

**JOB POSTING (from {job_url}):**
{job_content}

Perform a thorough, honest analysis using the format below. Be specific — name actual skills, tools, and phrases from both documents.

---

## COMPATIBILITY SCORE: [X]%
Score 0–100. Base it on: keyword overlap, required skills coverage, experience level match, and industry fit.
Provide one sentence explaining what drives this score.

## JOB DETAILS
- **Job Title:**
- **Company:**
- **Location:**
- **Employment Type:**

## KEYWORD MATCH ANALYSIS
List the most important keywords/phrases from the job posting, and whether each appears in the resume:

| Keyword / Phrase | In Resume? | Notes |
|------------------|-----------|-------|
| [keyword] | ✅ Yes / ❌ No | [brief note] |

Include at least 10 keywords. Focus on tools, technologies, skills, and role-specific terminology.

## HARD SKILLS GAP
**Matched hard skills:** List skills/tools the resume clearly demonstrates that the job requires.
**Missing hard skills:** List required technical skills, tools, or certifications not evident in the resume.

## SOFT SKILLS & EXPERIENCE GAP
**Matched:** Soft skills and experience types that align.
**Missing or weak:** Soft skills or experience the job emphasises that are not well demonstrated.

## RESUME SECTION REVIEW
Review each key section of the resume against what this role needs:

**Professional Summary:** [Is it tailored? Does it lead with the right strengths for this role?]
**Work Experience:** [Are the right achievements highlighted? Are bullets outcome-focused with metrics?]
**Skills Section:** [Does it reflect the job's requirements? What's missing?]
**Education/Certifications:** [Relevant? Any gaps?]

## SPECIFIC IMPROVEMENT RECOMMENDATIONS
Provide 4–6 concrete, actionable edits — reference exact bullet points or phrases to change:
1. [Specific change with before/after example where possible]
2. ...

## ATS OPTIMIZATION
List the exact keywords and phrases to add to pass ATS screening for this role. Group by category:
- **Must-add keywords:** (critical — likely filtered by ATS)
- **Nice-to-add keywords:** (boost ranking)

## COVER LETTER TALKING POINTS
Top 3 specific talking points tailored to this role and company.

Be direct and honest. If the match is weak, say so clearly and explain why.
"""

    try:
        message = client.messages.create(
            model=MODEL_NAME,
            max_tokens=3000,
            messages=[{"role": "user", "content": prompt}]
        )
        return {'success': True, 'analysis': message.content[0].text}
    except Exception as e:
        return {'success': False, 'error': str(e)}


def generate_cover_letter(resume_text: str, job_content: str, job_url: str, analysis_text: str, client,
                          tone: str = "Professional", length: str = "Standard (300–350 words)",
                          incorporate_recs: bool = True, prior_letter: str = None,
                          change_instructions: str = None, user_guidance: str = None):
    word_map = {"Brief (≈200 words)": "approximately 200", "Standard (300–350 words)": "300–350", "Detailed (≈450 words)": "approximately 450"}
    word_target = word_map.get(length, "300–350")
    recs_line = (
        "4. Subtly weave in 1–2 of the specific improvement recommendations from the match analysis — show you understand what the role demands."
        if incorporate_recs else
        "4. Address any apparent gaps using transferable skills and relevant context."
    )

    if prior_letter and change_instructions:
        prompt = f"""You are an expert career coach and professional cover letter writer.

You previously wrote the cover letter below. The user has requested specific changes. Apply them faithfully while keeping everything else intact.

**CURRENT COVER LETTER:**
{prior_letter}

**REQUESTED CHANGES:**
{change_instructions}

**CONTEXT — MY RESUME:**
{resume_text}

**CONTEXT — JOB POSTING (from {job_url}):**
{job_content}

**Keep these settings:**
- Tone: {tone}
- Length: {word_target} words
- Do NOT include placeholder text or generic phrases.

Return only the revised cover letter, starting directly with the salutation.
"""
    else:
        prompt = f"""You are an expert career coach and professional cover letter writer.

Based on this resume analysis and job posting, write a compelling cover letter.

**MY RESUME:**
{resume_text}

**JOB POSTING (from {job_url}):**
{job_content}

**MATCH ANALYSIS:**
{analysis_text}

**Instructions:**
1. Tone: {tone} — maintain this voice consistently throughout
2. Length: {word_target} words
3. Strong opening hook that immediately shows value — no "I am writing to apply" clichés
{recs_line}
5. Highlight 2–3 key achievements from the resume that directly match job requirements
6. Clear, confident call to action in the closing paragraph
7. Tailored to this specific job and company — reference the role and company by name
{f"8. Additional guidance from the user: {user_guidance}" if user_guidance and user_guidance.strip() else ""}

Do NOT include placeholder text or generic phrases.
Start directly with the salutation. Make it ready to copy-paste.
"""
    try:
        message = client.messages.create(
            model=MODEL_NAME,
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}]
        )
        return {'success': True, 'cover_letter': message.content[0].text}
    except Exception as e:
        return {'success': False, 'error': str(e)}


# ============= EXCEL TRACKER =============

def parse_analysis_fields(analysis_text: str) -> dict:
    """Extract job title, company, location and match % from analysis text."""
    fields = {'job_title': '', 'company': '', 'location': '', 'match_pct': ''}

    score_match = re.search(r'COMPATIBILITY SCORE[:\s]+(\d+)%', analysis_text, re.IGNORECASE)
    if score_match:
        fields['match_pct'] = score_match.group(1) + '%'

    title_match = re.search(r'\*{0,2}Job Title[:\*\s]+(.+)', analysis_text, re.IGNORECASE)
    if title_match:
        fields['job_title'] = title_match.group(1).strip().strip('*')

    company_match = re.search(r'\*{0,2}Company[:\*\s]+(.+)', analysis_text, re.IGNORECASE)
    if company_match:
        fields['company'] = company_match.group(1).strip().strip('*')

    location_match = re.search(r'\*{0,2}Location[:\*\s]+(.+)', analysis_text, re.IGNORECASE)
    if location_match:
        fields['location'] = location_match.group(1).strip().strip('*')

    return fields


def save_to_tracker(job_title: str, company: str, location: str,
                    resume_filename: str, match_pct: str, job_url: str,
                    cover_letter: str = '', cover_letter_path: str = ''):
    """Insert a job application row into Supabase."""
    sb = get_supabase()
    row = {
        "date":        datetime.now().strftime('%Y-%m-%d'),
        "job_title":   job_title,
        "company":     company,
        "location":    location,
        "resume_file": resume_filename,
        "match_pct":   match_pct,
        "job_url":     job_url if job_url != 'Manual Input' else '',
        "status":      "Applied",
        "notes":       "",
        "cover_letter": cover_letter_path if cover_letter_path else cover_letter,
    }
    if sb:
        sb.table("applications").insert(row).execute()
    else:
        raise RuntimeError("Supabase is not configured — cannot save application.")


@st.cache_data(ttl=30)
def load_tracker_data():
    """Load all applications from Supabase, newest first."""
    sb = get_supabase()
    if not sb:
        return []
    rows = sb.table("applications").select("*").order("created_at", desc=True).execute()
    return rows.data or []


def generate_tracker_excel(tracker_data: list) -> bytes:
    """Build an in-memory Excel workbook from tracker_data (list of dicts)."""
    headers  = ['Date', 'Job Title', 'Company', 'Location', 'Resume File', 'Match %', 'Job URL', 'Status', 'Notes', 'Cover Letter']
    col_keys = ['date', 'job_title', 'company', 'location', 'resume_file', 'match_pct', 'job_url', 'status', 'notes', 'cover_letter']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Applications"

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font      = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill      = openpyxl.styles.PatternFill(fill_type="solid", fgColor="230344")
        cell.alignment = openpyxl.styles.Alignment(horizontal="center")

    widths = [14, 30, 25, 20, 35, 10, 45, 15, 30, 80]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width

    for i, item in enumerate(tracker_data, 2):
        for col, key in enumerate(col_keys, 1):
            ws.cell(row=i, column=col, value=item.get(key) or '')
        ws.cell(row=i, column=10).alignment = openpyxl.styles.Alignment(wrap_text=True, vertical='top')
        if i % 2 == 0:
            fill = openpyxl.styles.PatternFill(fill_type="solid", fgColor="F3F0F8")
            for col in range(1, len(headers) + 1):
                ws.cell(row=i, column=col).fill = fill

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============= RESUME UPDATER =============

def generate_resume_updates(resume_text: str, analysis_text: str, client):
    """
    Ask Claude to return targeted find→replace pairs based on recommendations.
    Strictly forbidden from fabricating skills or experience not in the original resume.
    """
    prompt = f"""You are a professional resume editor with a strict honesty policy.

**CURRENT RESUME TEXT:**
{resume_text}

**ANALYSIS & RECOMMENDATIONS:**
{analysis_text}

Your task is to improve how existing experience is communicated — NOT to add experience that doesn't exist.

⚠️ HONESTY RULES (non-negotiable):
- NEVER add skills, tools, certifications, job titles, or responsibilities not already present in the resume
- NEVER invent metrics, numbers, or achievements — only use figures already stated in the resume
- NEVER add experience with software, frameworks, or industries not mentioned in the resume
- You MAY: use stronger action verbs, improve sentence structure, make achievements clearer, reframe existing content more powerfully, fix vague language
- If a recommendation requires adding something the candidate clearly doesn't have, SKIP that change entirely

Return ONLY a JSON array. Each item must have:
- "find": exact text from the resume (copy verbatim, including bullet characters)
- "replace": the improved version (same factual content, better expression)
- "description": one short sentence explaining what changed and why (e.g. "Replaced weak verb 'managed' with 'led'" or "Made outcome explicit from existing context")

```json
[
  {{
    "find": "exact text copied verbatim from the resume",
    "replace": "improved version — same facts, better expression",
    "description": "Short explanation of the change"
  }}
]
```

Rules:
- "find" MUST appear exactly in the resume — copy it verbatim including bullet prefix characters
- Keep the same bullet prefix/format as the original
- Make 5–10 targeted changes — not a full rewrite
- Do NOT change: job titles, company names, dates, education details, section headers

Return only the JSON array, no other text."""

    try:
        message = client.messages.create(
            model=MODEL_NAME,
            max_tokens=3000,
            messages=[{"role": "user", "content": prompt}]
        )
        response_text = message.content[0].text
        json_match = re.search(r'```json\s*(.*?)\s*```', response_text, re.DOTALL)
        json_str = json_match.group(1) if json_match else response_text.strip()
        updates = json.loads(json_str)
        return {'success': True, 'updates': updates}
    except Exception as e:
        return {'success': False, 'error': str(e)}


def apply_updates_to_docx(file_bytes: bytes, updates: list, original_filename: str) -> bytes:
    """
    Apply find→replace pairs to the original DOCX, preserving all formatting.
    Returns the updated DOCX as bytes.
    """
    doc = Document(io.BytesIO(file_bytes))
    applied = 0

    for update in updates:
        find_text = update.get('find', '').strip()
        replace_text = update.get('replace', '').strip()
        if not find_text or not replace_text:
            continue

        for para in doc.paragraphs:
            if find_text in para.text:
                # Preserve the run formatting of the first run
                if para.runs:
                    # Put all text into first run, clear the rest
                    para.runs[0].text = para.text.replace(find_text, replace_text)
                    for run in para.runs[1:]:
                        run.text = ''
                else:
                    para.text = para.text.replace(find_text, replace_text)
                applied += 1
                break

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue(), applied


# ============= DOCX BUILDERS =============

def _add_md_run(para, text: str):
    """Add runs to a paragraph, handling **bold** markers."""
    parts = re.split(r'(\*\*[^*]+\*\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            run = para.add_run(part[2:-2])
            run.bold = True
        elif part:
            para.add_run(part)


def _add_md_table(doc, table_lines: list):
    """Render pipe-delimited markdown table rows as a Word table."""
    rows = []
    for line in table_lines:
        if re.match(r'^\|[\s\-|:]+\|$', line.strip()):
            continue  # separator row
        cells = [c.strip() for c in line.strip().strip('|').split('|')]
        if cells:
            rows.append(cells)
    if not rows:
        return
    ncols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.style = 'Table Grid'
    for i, row in enumerate(rows):
        for j in range(ncols):
            cell_text = row[j] if j < len(row) else ''
            clean = re.sub(r'\*\*(.*?)\*\*', r'\1', cell_text)
            cell = table.rows[i].cells[j]
            cell.text = ''
            run = cell.paragraphs[0].add_run(clean)
            run.font.size = Pt(9)
            if i == 0:
                run.bold = True


def create_cover_letter_docx(cover_letter_text: str) -> bytes:
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin    = Inches(1.1)
    sec.bottom_margin = Inches(1.1)
    sec.left_margin   = Inches(1.25)
    sec.right_margin  = Inches(1.25)

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    for line in cover_letter_text.split('\n'):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0 if not line.strip() else 6)
        _add_md_run(p, line)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def create_analysis_docx(analysis_text: str, job_url: str, resume_filename: str) -> bytes:
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin    = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.left_margin   = Inches(1.15)
    sec.right_margin  = Inches(1.15)

    # Title block
    title = doc.add_heading('ResumeSync — Compatibility Analysis', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    meta = doc.add_paragraph()
    meta.add_run(f'Resume: ').bold = True
    meta.add_run(resume_filename)
    meta2 = doc.add_paragraph()
    meta2.add_run('Job: ').bold = True
    meta2.add_run(job_url)
    meta3 = doc.add_paragraph()
    meta3.add_run('Date: ').bold = True
    meta3.add_run(datetime.now().strftime('%Y-%m-%d'))
    doc.add_paragraph()

    lines = analysis_text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i]

        # Headings
        if line.startswith('### '):
            doc.add_heading(line[4:].strip(), level=3)
        elif line.startswith('## '):
            doc.add_heading(line[3:].strip(), level=2)
        elif line.startswith('# '):
            doc.add_heading(line[2:].strip(), level=1)

        # Table block — collect consecutive pipe lines
        elif line.strip().startswith('|'):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i])
                i += 1
            _add_md_table(doc, table_lines)
            doc.add_paragraph()
            continue

        # Bullet list
        elif re.match(r'^[-*]\s', line):
            p = doc.add_paragraph(style='List Bullet')
            _add_md_run(p, line[2:].strip())

        # Numbered list
        elif re.match(r'^\d+\.\s', line):
            p = doc.add_paragraph(style='List Number')
            _add_md_run(p, re.sub(r'^\d+\.\s', '', line))

        # Empty line
        elif not line.strip():
            doc.add_paragraph()

        # Normal paragraph (with inline bold)
        else:
            p = doc.add_paragraph()
            _add_md_run(p, line)

        i += 1

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============= LOGO =============

@st.cache_data
def get_vp_logo_b64():
    from PIL import Image
    import numpy as np
    logo_path = os.path.join(os.path.dirname(__file__), "assets", "logo.png")
    if not os.path.exists(logo_path):
        return None
    img = Image.open(logo_path).convert("RGB")
    # Resize to sidebar-friendly width
    w, h = img.size
    new_w = 200
    img = img.resize((new_w, int(h * new_w / w)), Image.LANCZOS)
    data = np.array(img)
    r, g, b = data[:, :, 0], data[:, :, 1], data[:, :, 2]
    # Purple background mask: #230344 ≈ R<80, G<30, B>30
    mask = (r < 80) & (g < 30) & (b > 30)
    # Replace background with sidebar green #0a1b16
    data[mask, 0] = 8
    data[mask, 1] = 32
    data[mask, 2] = 25
    # Gold/bronze icon mask: warm tone with high R, moderate G, low B
    r2, g2, b2 = data[:, :, 0], data[:, :, 1], data[:, :, 2]
    gold_mask = (r2 > 100) & (g2 > 60) & (b2 < 120) & (r2 > b2) & (~mask)
    # Replace gold with accent green #7ad79f
    data[gold_mask, 0] = 109
    data[gold_mask, 1] = 193
    data[gold_mask, 2] = 138
    result = Image.fromarray(data.astype(np.uint8))
    buf = io.BytesIO()
    result.save(buf, format="PNG")
    buf.seek(0)
    return base64.b64encode(buf.getvalue()).decode()


# ============= STREAMLIT UI =============

def main():
    # ── Sidebar ─────────────────────────────────────────────────────────────
    with st.sidebar:
        tracker_data = load_tracker_data()
        count = len(tracker_data)
        st.markdown(f"""
        <div style="padding: 1.5rem 1rem 0;">

          <!-- Brand mark -->
          <div style="display:flex;align-items:center;gap:11px;margin-bottom:1.4rem;">
            <div style="width:34px;height:34px;border-radius:9px;background:linear-gradient(150deg,#7ad79f,#4fae7a);display:grid;place-items:center;font-family:'Bricolage Grotesque',system-ui,sans-serif;font-weight:800;font-size:16px;color:#06140f;box-shadow:0 4px 12px rgba(122,215,159,0.28);flex-shrink:0;">R</div>
            <div>
              <div style="font-family:'Bricolage Grotesque',system-ui,sans-serif;font-weight:700;font-size:17px;letter-spacing:-0.02em;color:#ecf4ee;line-height:1.1;">ResumSync</div>
              <div style="font-family:'Space Mono',monospace;font-size:9px;letter-spacing:0.18em;text-transform:uppercase;color:#6e8a7b;margin-top:3px;">by VisualizePro</div>
            </div>
          </div>

          <!-- Divider -->
          <div style="border-top:1px solid rgba(159,182,168,0.12); margin-bottom:1.2rem;"></div>

          <!-- What you get -->
          <div style="border-top:1px solid rgba(159,182,168,0.12); padding-top:1rem; margin-bottom:1.2rem;">
            <p style="font-family:'Space Mono',monospace; font-size:9px; letter-spacing:0.16em; color:#7ad79f; text-transform:uppercase; margin:0 0 0.8rem;">What you get</p>
            <div style="display:flex; flex-direction:column; gap:0.6rem;">
              <div style="display:flex; gap:10px; align-items:flex-start;">
                <span style="color:#7ad79f; font-size:14px; flex-shrink:0; margin-top:1px;">📊</span>
                <span style="font-family:'DM Sans',sans-serif; font-size:0.8rem; color:#ecf4ee; line-height:1.4;"><b>Match % score</b> <span style="color:#9fb6a8;">— see your fit before you apply</span></span>
              </div>
              <div style="display:flex; gap:10px; align-items:flex-start;">
                <span style="color:#7ad79f; font-size:14px; flex-shrink:0; margin-top:1px;">🎯</span>
                <span style="font-family:'DM Sans',sans-serif; font-size:0.8rem; color:#ecf4ee; line-height:1.4;"><b>Skills &amp; keyword gaps</b> <span style="color:#9fb6a8;">— exactly what's missing</span></span>
              </div>
              <div style="display:flex; gap:10px; align-items:flex-start;">
                <span style="color:#7ad79f; font-size:14px; flex-shrink:0; margin-top:1px;">🤖</span>
                <span style="font-family:'DM Sans',sans-serif; font-size:0.8rem; color:#ecf4ee; line-height:1.4;"><b>ATS optimisation</b> <span style="color:#9fb6a8;">— get past the robot screeners</span></span>
              </div>
              <div style="display:flex; gap:10px; align-items:flex-start;">
                <span style="color:#7ad79f; font-size:14px; flex-shrink:0; margin-top:1px;">✍️</span>
                <span style="font-family:'DM Sans',sans-serif; font-size:0.8rem; color:#ecf4ee; line-height:1.4;"><b>Tailored cover letter</b> <span style="color:#9fb6a8;">— ready to send in seconds</span></span>
              </div>
              <div style="display:flex; gap:10px; align-items:flex-start;">
                <span style="color:#7ad79f; font-size:14px; flex-shrink:0; margin-top:1px;">✨</span>
                <span style="font-family:'DM Sans',sans-serif; font-size:0.8rem; color:#ecf4ee; line-height:1.4;"><b>Resume auto-updater</b> <span style="color:#9fb6a8;">— stronger words, same facts</span></span>
              </div>
            </div>
          </div>

          <!-- Job tracker count -->
          <div style="background:#0c2019;border:1px solid rgba(159,182,168,0.12);border-radius:12px;padding:16px;">
            <p style="font-family:'Space Mono',monospace; font-size:9.5px; letter-spacing:0.16em; text-transform:uppercase; color:#6e8a7b; margin:0 0 0.5rem;">Applications tracked</p>
            <p style="font-family:'Bricolage Grotesque',system-ui,sans-serif; font-size:2rem; font-weight:800; color:#7ad79f; margin:0 0 0.2rem; line-height:1;">
              {count}
            </p>
            <p style="font-family:'Space Mono',monospace; font-size:0.68rem; color:#6e8a7b; margin:0; line-height:1.4;">
              {'Track your first application below' if count == 0 else 'role(s) in your tracker →'}
            </p>
          </div>

        </div>
        """, unsafe_allow_html=True)

        if tracker_data:
            st.download_button(
                label="📥 Download Tracker",
                data=generate_tracker_excel(tracker_data),
                file_name="job_applications.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="sidebar_download_tracker",
                use_container_width=True
            )

    # ── Compact welcome topbar (matches handoff App design — lean workspace, not landing hero) ──
    st.markdown("""
    <div style="display:flex;align-items:center;justify-content:space-between;gap:20px;
                padding:20px 0 22px;">
      <div style="display:flex;align-items:center;gap:14px;">
        <div style="width:46px;height:46px;border-radius:11px;
                    background:linear-gradient(150deg,#1d3a31,#11251d);
                    border:1px solid rgba(122,215,159,0.22);
                    display:grid;place-items:center;
                    font-family:'Bricolage Grotesque',system-ui,sans-serif;
                    font-weight:700;font-size:18px;color:#7ad79f;flex-shrink:0;">M</div>
        <div>
          <div style="font-family:'Bricolage Grotesque',system-ui,sans-serif;font-weight:700;
                      font-size:22px;letter-spacing:-0.02em;color:#ecf4ee;line-height:1.1;">
            Let's find your next match.
          </div>
          <div style="font-size:13px;color:#9fb6a8;margin-top:3px;font-family:'DM Sans',sans-serif;">
            Upload your resume and a job posting to get your compatibility score.
          </div>
        </div>
      </div>
      <div style="display:flex;align-items:center;gap:7px;
                  font-family:'Space Mono',monospace;font-size:10px;
                  letter-spacing:0.12em;text-transform:uppercase;color:#6e8a7b;white-space:nowrap;">
        <span style="width:7px;height:7px;border-radius:50%;background:#7ad79f;
                     box-shadow:0 0 0 3px rgba(122,215,159,0.10);display:inline-block;"></span>
        <span style="width:7px;height:7px;border-radius:50%;background:rgba(159,182,168,0.22);display:inline-block;"></span>
        <span style="width:7px;height:7px;border-radius:50%;background:rgba(159,182,168,0.22);display:inline-block;"></span>
        <span style="margin-left:4px;">Step 1 of 3</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    client = get_client()
    if not client:
        st.error("⚠️ ANTHROPIC_API_KEY not found. On Streamlit Cloud: go to App Settings → Secrets and add your key in TOML format: ANTHROPIC_API_KEY = \"sk-ant-...\"")
        return

    # ── Workspace card ────────────────────────────────────────────────────────
    st.markdown("""
    <div style="background:#0c2019;border:1px solid rgba(159,182,168,0.12);border-radius:20px;
                padding:28px 28px 4px;box-shadow:0 24px 60px rgba(0,0,0,0.45),0 4px 14px rgba(0,0,0,0.35);
                margin-bottom:4px;">
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown("""
        <div style="display:flex;align-items:center;gap:9px;
                    font-family:'Space Mono',monospace;font-size:11px;letter-spacing:0.14em;
                    text-transform:uppercase;color:#7ad79f;margin-bottom:10px;">
          <span style="width:19px;height:19px;border-radius:50%;background:rgba(122,215,159,0.16);
                       display:grid;place-items:center;font-size:10px;color:#7ad79f;flex-shrink:0;
                       text-align:center;line-height:19px;">1</span>
          Upload your resume
        </div>
        """, unsafe_allow_html=True)
        resume_file = st.file_uploader(
            "Resume", type=['pdf', 'docx', 'doc', 'txt'], label_visibility="collapsed"
        )
        if resume_file:
            st.markdown(f'<p style="font-size:0.75rem;color:#7ad79f;margin:4px 0 0;font-family:DM Sans,sans-serif;">✅ {resume_file.name}</p>', unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div style="display:flex;align-items:center;gap:9px;
                    font-family:'Space Mono',monospace;font-size:11px;letter-spacing:0.14em;
                    text-transform:uppercase;color:#7ad79f;margin-bottom:10px;">
          <span style="width:19px;height:19px;border-radius:50%;background:rgba(122,215,159,0.16);
                       display:grid;place-items:center;font-size:10px;color:#7ad79f;flex-shrink:0;
                       text-align:center;line-height:19px;">2</span>
          Add the job posting
        </div>
        """, unsafe_allow_html=True)
        input_method = st.radio(
            "Input method", ["Paste URL · auto-scrape", "Paste manually"],
            horizontal=True, label_visibility="collapsed"
        )
        job_url = None
        manual_job_text = None
        if input_method == "Paste URL · auto-scrape":
            job_url = st.text_input("Job URL", placeholder="https://www.seek.com.au/job/...", label_visibility="collapsed")
            if job_url:
                st.markdown(f'<p style="font-size:0.75rem;color:#7ad79f;margin:4px 0 0;font-family:DM Sans,sans-serif;">✅ {job_url[:55]}</p>', unsafe_allow_html=True)
        else:
            manual_job_text = st.text_area("Job description", placeholder="Paste the full job description here...", height=100, label_visibility="collapsed")
            if manual_job_text:
                st.markdown(f'<p style="font-size:0.75rem;color:#7ad79f;margin:4px 0 0;font-family:DM Sans,sans-serif;">✅ {len(manual_job_text)} characters</p>', unsafe_allow_html=True)

    st.markdown("""
    <div style="border-top:1px solid rgba(159,182,168,0.07);margin:12px 0 4px;"></div>
    """, unsafe_allow_html=True)

    col_hint, col_btn = st.columns([2, 1])
    with col_hint:
        ready = resume_file and (job_url or manual_job_text)
        dot_color = "#7ad79f" if ready else "#6e8a7b"
        hint_text = "Ready to analyze" if ready else "Add both to analyze"
        st.markdown(f"""
        <div style="display:flex;align-items:center;gap:8px;font-size:12px;color:{dot_color};
                    font-family:'DM Sans',sans-serif;padding-top:10px;">
          <span style="width:6px;height:6px;border-radius:50%;background:{dot_color};
                       display:inline-block;flex-shrink:0;"></span>
          {hint_text}
        </div>
        """, unsafe_allow_html=True)
    with col_btn:
        analyze_button = st.button(
            "Analyze compatibility →",
            type="primary",
            use_container_width=True,
            disabled=not (resume_file and (job_url or manual_job_text))
        )

    if analyze_button:
        if not resume_file:
            st.error("❌ Please upload a resume file first.")
            return
        if not (job_url or manual_job_text):
            st.error("❌ Please add a job posting.")
            return

        with st.spinner("🔍 Analyzing your resume against the job posting..."):

            with st.status("Extracting text from resume...", expanded=True) as status:
                resume_text = extract_resume_text(resume_file)
                if "Error" in resume_text:
                    st.error(f"❌ {resume_text}")
                    return
                st.write(f"✅ Extracted {len(resume_text)} characters from resume")
                status.update(label="Resume extracted!", state="complete")

            if job_url:
                with st.status("Fetching job posting from URL...", expanded=True) as status:
                    job_data = scrape_job_url(job_url)
                    if not job_data['success']:
                        st.error(f"❌ Could not fetch job posting: {job_data['error']}")
                        st.info("💡 Switch to 'Paste job description manually'.")
                        return
                    job_content = job_data['content']
                    st.write(f"✅ Fetched {len(job_content)} characters from job posting")
                    status.update(label="Job posting fetched!", state="complete")
            else:
                job_content = manual_job_text
                job_url = "Manual Input"
                st.success(f"✅ Using manually pasted job description ({len(job_content)} characters)")

            with st.status("Analysing with Claude AI...", expanded=True) as status:
                result = analyze_resume_vs_job(resume_text, job_content, job_url, client)
                if not result['success']:
                    st.error(f"❌ Analysis failed: {result['error']}")
                    return
                status.update(label="Analysis complete!", state="complete")

        st.session_state['analysis_result'] = result
        st.session_state['resume_text'] = resume_text
        st.session_state['job_content'] = job_content
        st.session_state['job_url'] = job_url
        st.session_state['resume_filename'] = resume_file.name
        st.session_state['resume_file_bytes'] = resume_file.getvalue()
        st.session_state['resume_is_docx'] = resume_file.name.lower().endswith(('.docx', '.doc'))
        st.session_state['tracker_saved'] = False

    # Results
    if 'analysis_result' in st.session_state:
        result = st.session_state['analysis_result']
        resume_text = st.session_state['resume_text']
        job_content = st.session_state['job_content']
        job_url = st.session_state['job_url']
        resume_filename = st.session_state['resume_filename']

        st.divider()

        col_new = st.columns([3, 1])[1]
        with col_new:
            if st.button("🔄 New Analysis", key="new_analysis"):
                for key in ['analysis_result', 'resume_text', 'job_content', 'job_url',
                            'resume_filename', 'resume_file_bytes', 'resume_is_docx',
                            'cover_letter', 'tracker_saved', 'proposed_updates', 'updated_resume_bytes']:
                    st.session_state.pop(key, None)
                st.rerun()

        st.markdown('<h2 style="color:#ecf4ee; font-size:1.8rem; font-weight:700; text-align:center; margin:2rem 0; font-family:Bricolage Grotesque,serif; letter-spacing:-0.02em;">📋 Analysis Results</h2>', unsafe_allow_html=True)
        st.markdown(result['analysis'])

        # Build shared filename parts used across all downloads
        fields = parse_analysis_fields(result['analysis'])
        _fn_date = datetime.now().strftime('%Y-%m-%d')
        _fn_person = re.sub(r'[^\w\-]', '_', resume_filename.rsplit('.', 1)[0])[:25].strip('_')
        _fn_role = re.sub(r'[^\w\-]', '_', fields.get('job_title', 'Role').replace(' ', '_'))[:25].strip('_')

        st.divider()

        # ============= RESUME UPDATER =============
        st.markdown('<h2 style="color:#ecf4ee; font-size:1.6rem; font-weight:700; margin:1rem 0 0.5rem; font-family:\'Bricolage Grotesque\',system-ui,sans-serif; letter-spacing:-0.02em;">✨ Update My Resume</h2>', unsafe_allow_html=True)

        if st.session_state.get('resume_is_docx'):
            st.markdown(
                '<div style="background:rgba(122,215,159,0.06);padding:1rem 1.5rem;border-radius:12px;'
                'border-left:4px solid #7ad79f;margin-bottom:1rem;">'
                '<p style="color:#9fb6a8;font-size:0.9rem;margin:0;font-family:\'DM Sans\',sans-serif;line-height:1.6;">'
                'Improves how your existing experience is expressed — stronger verbs, clearer outcomes. '
                '<strong style="color:#ecf4ee;">Will never add skills or experience you don\'t have.</strong> '
                'You review and approve every change before it\'s applied.</p>'
                '</div>',
                unsafe_allow_html=True
            )

            col_upd = st.columns([1, 2, 1])[1]
            with col_upd:
                update_btn = st.button("✨ Propose Resume Changes", type="primary", key="update_resume_btn", use_container_width=True)

            if update_btn:
                with st.spinner("Generating proposed changes..."):
                    with st.status("Analysing what can be improved...", expanded=True) as upd_status:
                        upd_result = generate_resume_updates(resume_text, result['analysis'], client)
                        if not upd_result['success']:
                            st.error(f"❌ Could not generate changes: {upd_result['error']}")
                        else:
                            st.write(f"✅ {len(upd_result['updates'])} proposed changes ready for review")
                            upd_status.update(label="Proposals ready — please review below", state="complete")
                            st.session_state['proposed_updates'] = upd_result['updates']
                            st.session_state.pop('updated_resume_bytes', None)

            # ── Review step ────────────────────────────────────────────────
            if 'proposed_updates' in st.session_state and st.session_state['proposed_updates']:
                proposed = st.session_state['proposed_updates']

                st.markdown(
                    '<p style="font-family:\'DM Sans\',sans-serif;font-weight:600;color:#ecf4ee;'
                    'margin:1.5rem 0 0.5rem;">Review each proposed change — tick only the ones that are true for you:</p>',
                    unsafe_allow_html=True
                )

                selected = []
                for i, change in enumerate(proposed):
                    find = change.get('find', '')
                    replace = change.get('replace', '')
                    description = change.get('description', 'Improve phrasing')

                    with st.container():
                        checked = st.checkbox(f"**Change {i+1}:** {description}", value=True, key=f"chk_{i}")
                        col_b, col_a = st.columns(2)
                        with col_b:
                            st.markdown(
                                f'<div style="background:rgba(224,122,95,0.08);border-left:3px solid #e07a5f;'
                                f'padding:0.6rem 0.8rem;border-radius:10px;font-size:0.82rem;'
                                f'font-family:\'DM Sans\',sans-serif;color:#9fb6a8;">'
                                f'<strong style="color:#e07a5f;font-size:0.75rem;font-family:\'Space Mono\',monospace;'
                                f'letter-spacing:0.1em;text-transform:uppercase;">Before</strong><br>{find}</div>',
                                unsafe_allow_html=True
                            )
                        with col_a:
                            st.markdown(
                                f'<div style="background:rgba(122,215,159,0.08);border-left:3px solid #7ad79f;'
                                f'padding:0.6rem 0.8rem;border-radius:10px;font-size:0.82rem;'
                                f'font-family:\'DM Sans\',sans-serif;color:#ecf4ee;">'
                                f'<strong style="color:#7ad79f;font-size:0.75rem;font-family:\'Space Mono\',monospace;'
                                f'letter-spacing:0.1em;text-transform:uppercase;">After</strong><br>{replace}</div>',
                                unsafe_allow_html=True
                            )
                        st.markdown("<div style='margin-bottom:0.8rem'></div>", unsafe_allow_html=True)
                        if checked:
                            selected.append(change)

                n_selected = len(selected)
                col_apply = st.columns([1, 2, 1])[1]
                with col_apply:
                    apply_btn = st.button(
                        f"✅ Apply {n_selected} Selected Change{'s' if n_selected != 1 else ''}",
                        type="primary",
                        key="apply_selected_btn",
                        disabled=n_selected == 0
                    )

                if apply_btn and selected:
                    updated_bytes, applied = apply_updates_to_docx(
                        st.session_state['resume_file_bytes'],
                        selected,
                        resume_filename
                    )
                    st.session_state['updated_resume_bytes'] = updated_bytes
                    st.success(f"✅ {applied} change(s) applied to your resume.")

            if 'updated_resume_bytes' in st.session_state:
                new_filename = f"{_fn_person}_Updated_{_fn_role}_{_fn_date}.docx"
                col_dl_upd = st.columns([1, 2, 1])[1]
                with col_dl_upd:
                    st.download_button(
                        label="📥 Download Updated Resume",
                        data=st.session_state['updated_resume_bytes'],
                        file_name=new_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="download_updated_resume"
                    )
        else:
            st.markdown(
                '<div style="background:rgba(122,215,159,0.04);border:1px solid rgba(159,182,168,0.12);'
                'border-radius:10px;padding:0.75rem 1rem;">'
                '<p style="color:#6e8a7b;font-size:0.85rem;margin:0;font-family:\'DM Sans\',sans-serif;">'
                '💡 Upload a Word document (.docx) to enable automatic resume updating.</p>'
                '</div>',
                unsafe_allow_html=True
            )

        st.divider()

        # ============= COVER LETTER =============
        st.markdown('<h2 style="color:#ecf4ee; font-size:1.8rem; font-weight:700; text-align:center; margin:2rem 0; font-family:Bricolage Grotesque,serif; letter-spacing:-0.02em;">✍️ Generate Cover Letter</h2>', unsafe_allow_html=True)
        st.markdown(
            '<div style="background:rgba(109,193,138,0.06); padding:1.2rem 1.5rem; border-radius:8px; border-left:4px solid #7ad79f; margin-bottom:2rem;">'
            '<p style="color:#9fb6a8; font-size:0.95rem; margin:0; font-family:DM Sans,sans-serif;">🎯 Create a personalised cover letter tailored to this job based on your match analysis</p>'
            '</div>',
            unsafe_allow_html=True
        )

        cl_col1, cl_col2, cl_col3 = st.columns(3)
        with cl_col1:
            cl_tone = st.selectbox(
                "Tone",
                ["Professional", "Conversational", "Enthusiastic", "Formal", "Confident & Direct"],
                key="cl_tone"
            )
        with cl_col2:
            cl_length = st.selectbox(
                "Length",
                ["Brief (≈200 words)", "Standard (300–350 words)", "Detailed (≈450 words)"],
                index=1,
                key="cl_length"
            )
        with cl_col3:
            cl_incorporate = st.checkbox(
                "Incorporate improvement recommendations",
                value=True,
                key="cl_incorporate",
                help="Weaves 1–2 specific recommendations from the analysis into the letter"
            )

        cl_guidance = st.text_area(
            "Additional guidance *(optional)*",
            placeholder="e.g. Mention my leadership of the 2023 digital transformation project. Emphasise my Python skills. Keep it humble but confident.",
            height=90,
            key="cl_guidance"
        )

        col_gen = st.columns([1, 2, 1])[1]
        with col_gen:
            generate_cl_button = st.button("🚀 Generate Cover Letter", type="primary", key="generate_cover_letter", use_container_width=True)

        if generate_cl_button:
            with st.spinner("✍️ Writing your personalised cover letter..."):
                with st.status("Generating cover letter...", expanded=True) as status:
                    cl_result = generate_cover_letter(
                        resume_text, job_content, job_url, result['analysis'], client,
                        tone=cl_tone, length=cl_length, incorporate_recs=cl_incorporate,
                        user_guidance=cl_guidance
                    )
                    if not cl_result['success']:
                        st.error(f"❌ Cover letter generation failed: {cl_result['error']}")
                    else:
                        st.write("✅ Cover letter generated!")
                        status.update(label="Cover letter ready!", state="complete")
                        st.session_state['cover_letter'] = cl_result['cover_letter']

        if 'cover_letter' in st.session_state:
            st.markdown("---")
            st.markdown(
                '<div class="cl-box" style="background:white; padding:2rem; border-radius:8px; '
                'box-shadow:0 2px 8px rgba(0,0,0,0.1); border:1px solid rgb(220,220,220); margin:2rem 0;">'
                f'<div style="font-size:1rem; line-height:1.8; white-space:pre-wrap;">{st.session_state["cover_letter"]}</div>'
                '</div>',
                unsafe_allow_html=True
            )

            # ── Regenerate section ─────────────────────────────────────────
            st.markdown(
                '<p style="color:#9fb6a8; font-size:0.85rem; font-family:DM Sans,sans-serif; '
                'margin:0.5rem 0 0.3rem;">Want changes? Describe what to adjust and regenerate.</p>',
                unsafe_allow_html=True
            )
            cl_changes = st.text_area(
                "Proposed changes",
                placeholder="e.g. Make the opening more confident. Mention my Tableau experience earlier. Shorten the second paragraph. Use a more conversational tone in the closing.",
                height=100,
                key="cl_changes",
                label_visibility="collapsed"
            )

            regen_col, dl_col = st.columns([1, 1])
            with regen_col:
                regen_btn = st.button("🔄 Regenerate Cover Letter", key="regen_cover_letter", use_container_width=True)
            with dl_col:
                st.download_button(
                    label="💾 Download Cover Letter (.docx)",
                    data=create_cover_letter_docx(st.session_state['cover_letter']),
                    file_name=f"CoverLetter_{_fn_person}_{_fn_role}_{_fn_date}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download_cl",
                    use_container_width=True
                )

            if regen_btn:
                with st.spinner("✍️ Rewriting your cover letter..."):
                    with st.status("Applying your changes...", expanded=True) as regen_status:
                        cl_result = generate_cover_letter(
                            resume_text, job_content, job_url, result['analysis'], client,
                            tone=cl_tone, length=cl_length, incorporate_recs=cl_incorporate,
                            prior_letter=st.session_state['cover_letter'],
                            change_instructions=cl_changes if cl_changes.strip() else None,
                            user_guidance=cl_guidance
                        )
                        if not cl_result['success']:
                            st.error(f"❌ Regeneration failed: {cl_result['error']}")
                        else:
                            st.write("✅ Cover letter updated!")
                            regen_status.update(label="Done!", state="complete")
                            st.session_state['cover_letter'] = cl_result['cover_letter']
                            st.rerun()

        st.divider()

        # ============= JOB TRACKER =============
        st.markdown('<h2 style="color:#ecf4ee; font-size:1.8rem; font-weight:700; text-align:center; margin:2rem 0; font-family:Bricolage Grotesque,serif; letter-spacing:-0.02em;">📊 Save to Job Tracker</h2>', unsafe_allow_html=True)

        t1, t2, t3 = st.columns(3)
        with t1:
            job_title_input = st.text_input("Job Title", value=fields['job_title'], key="tracker_title")
        with t2:
            company_input = st.text_input("Company", value=fields['company'], key="tracker_company")
        with t3:
            location_input = st.text_input("Location", value=fields['location'], key="tracker_location")

        col_save = st.columns([1, 2, 1])[1]
        with col_save:
            if st.session_state.get('tracker_saved'):
                st.success(f"✅ Saved to tracker! ({TRACKER_PATH})")
            else:
                if st.button("💾 Save to Job Tracker", key="save_tracker", use_container_width=True):
                    cover_letter_text = st.session_state.get('cover_letter', '')
                    cl_path = ''
                    if cover_letter_text:
                        os.makedirs(COVER_LETTERS_PATH, exist_ok=True)
                        cl_filename = f"CoverLetter_{_fn_person}_{_fn_role}_{_fn_date}.docx"
                        cl_path = os.path.join(COVER_LETTERS_PATH, cl_filename)
                        with open(cl_path, 'wb') as f:
                            f.write(create_cover_letter_docx(cover_letter_text))
                    save_to_tracker(
                        job_title=job_title_input,
                        company=company_input,
                        location=location_input,
                        resume_filename=resume_filename,
                        match_pct=fields['match_pct'],
                        job_url=job_url,
                        cover_letter=cover_letter_text,
                        cover_letter_path=cl_path
                    )
                    st.session_state['tracker_saved'] = True
                    st.rerun()

        tracker_data = load_tracker_data()
        if tracker_data:
            with st.expander(f"📋 View Job Tracker ({len(tracker_data)} roles)", expanded=False):
                st.dataframe(tracker_data, use_container_width=True)
                st.download_button(
                    label="📥 Download Tracker (Excel)",
                    data=generate_tracker_excel(tracker_data),
                    file_name="job_applications.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_tracker"
                )

        st.divider()

        # Download full report
        col_download = st.columns([1, 2, 1])[1]
        with col_download:
            st.download_button(
                label="💾 Download Analysis Report (.docx)",
                data=create_analysis_docx(result['analysis'], job_url, resume_filename),
                file_name=f"ResumeAnalysis_{_fn_person}_{_fn_role}_{_fn_date}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_analysis",
                use_container_width=True
            )

    # Footer
    st.divider()
    st.markdown(
        '<div style="text-align:center;padding:2rem 0 1rem;font-family:\'DM Sans\',sans-serif;">'
        '<p style="color:#7ad79f; font-size:0.85rem; margin-bottom:0.4rem; font-family:Space Mono,monospace; letter-spacing:0.1em;">Powered by Claude AI</p>'
        '<p style="color:#9fb6a8; font-size:0.8rem; font-family:Space Mono,monospace; letter-spacing:0.08em;">© 2026 ResumeSync by VisualizePro · AI-Powered Career Intelligence</p>'
        '</div>',
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as _e:
        import traceback
        st.error(f"App error: {_e}")
        st.code(traceback.format_exc())
else:
    try:
        main()
    except Exception as _e:
        import traceback
        st.error(f"App error: {_e}")
        st.code(traceback.format_exc())
