"""
Resume Matcher Web App
Upload resume + paste job URL → Get compatibility score + recommendations
"""

import streamlit as st
import traceback as _tb

# Page config must be first Streamlit call
st.set_page_config(
    page_title="ResumeSync - AI-Powered Job Matching",
    page_icon="🔄",
    layout="wide"
)

try:
    import streamlit.components.v1 as _components
    import os
    import re
    import json
    import hashlib
    import urllib.parse
    from datetime import datetime
    from typing import Optional
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

    _IMPORT_ERROR = None
except Exception as _ie:
    _IMPORT_ERROR = _ie

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
    /* ── Analysis result headings (LLM markdown ## / ###) — keep compact ── */
    .stMarkdown h2 {
        font-size: 1.05rem !important;
        font-weight: 700 !important;
        margin: 1.4rem 0 0.35rem !important;
        padding-bottom: 0.3rem !important;
        border-bottom: 1px solid rgba(255,255,255,0.08) !important;
        letter-spacing: 0 !important;
    }
    .stMarkdown h3 {
        font-size: 0.92rem !important;
        font-weight: 600 !important;
        margin: 0.9rem 0 0.25rem !important;
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
    .stSelectbox [data-baseweb="select"] input,
    .stSelectbox [data-baseweb="single-select"],
    .stSelectbox [data-baseweb="single-select"] *,
    [data-baseweb="value-container"],
    [data-baseweb="value-container"] *,
    [data-baseweb="value-container"] > div,
    [data-baseweb="value-container"] > div > div {
        color: #ecf4ee !important;
        font-family: 'Bricolage Grotesque', system-ui, sans-serif !important;
        font-weight: 800 !important;
        font-size: 13px !important;
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


# ── Auth & subscription helpers ──────────────────────────────────────────────

FREE_CREDITS = 2

PACK_INFO = {
    "pack5":  {"name": "Starter Pack", "credits": 5,  "price": "A$9",  "amount": "9.00",  "color": "#6fb1e0"},
    "pack20": {"name": "Value Pack",   "credits": 20, "price": "A$15", "amount": "15.00", "color": "#7ad79f"},
    "pack30": {"name": "Pro Pack",     "credits": 30, "price": "A$20", "amount": "20.00", "color": "#a78bfa"},
    "pack50": {"name": "Max Pack",     "credits": 50, "price": "A$30", "amount": "30.00", "color": "#f59e0b"},
}

PAY_ADVANCED_URL = "https://pad.live/pa53992/pay"
APP_URL = "https://resumesync-adc8kjcujdwh8jnnhcgcpe.streamlit.app"


def _fresh_supabase():
    """Non-cached Supabase client used for auth calls."""
    if not _SUPABASE_AVAILABLE:
        return None
    url = _read_secret("SUPABASE_URL")
    key = _read_secret("SUPABASE_KEY")
    if url and key:
        try:
            return _supabase_create_client(url, key)
        except Exception:
            pass
    return None


def _auth_user_dict(resp) -> Optional[dict]:
    """Extract user dict from a Supabase auth response."""
    if resp and resp.user and resp.session:
        return {
            "user_id":       resp.user.id,
            "email":         resp.user.email or "",
            "access_token":  resp.session.access_token,
            "refresh_token": resp.session.refresh_token,
        }
    return None


def sign_in(email: str, password: str) -> tuple:
    """Sign in with email + password. Returns (user_dict | None, error_str)."""
    sb = _fresh_supabase()
    if not sb:
        return None, "Supabase not configured."
    try:
        resp = sb.auth.sign_in_with_password({"email": email, "password": password})
        user = _auth_user_dict(resp)
        return user, ("" if user else "Invalid email or password.")
    except Exception as e:
        msg = str(e)
        if "Invalid login credentials" in msg:
            return None, "Invalid email or password."
        return None, msg


def sign_up(email: str, password: str) -> tuple:
    """Create a pre-confirmed account using the service role key, then sign in."""
    if not _SUPABASE_AVAILABLE:
        return None, "Supabase not configured."
    url  = _read_secret("SUPABASE_URL")
    skey = _read_secret("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not skey:
        # Fall back to regular signup if no service role key
        return _sign_up_basic(email, password)
    try:
        admin_sb = _supabase_create_client(url, skey)
        # Create user as already confirmed — no email needed
        admin_sb.auth.admin.create_user({
            "email": email,
            "password": password,
            "email_confirm": True,
        })
    except Exception as e:
        msg = str(e)
        if "already been registered" in msg.lower() or "already registered" in msg.lower() or "already exists" in msg.lower():
            return None, "An account with that email already exists. Please sign in."
        return None, msg
    # Now sign in with the new credentials
    return sign_in(email, password)


def _sign_up_basic(email: str, password: str) -> tuple:
    """Fallback signup without service role key (requires email confirmation off in Supabase)."""
    sb = _fresh_supabase()
    if not sb:
        return None, "Supabase not configured."
    try:
        resp = sb.auth.sign_up({"email": email, "password": password})
        user = _auth_user_dict(resp)
        if user:
            return user, ""
        return None, "Account created — please check your inbox to confirm your email, then sign in."
    except Exception as e:
        msg = str(e)
        if "already registered" in msg.lower() or "already been registered" in msg.lower():
            return None, "An account with that email already exists. Please sign in."
        return None, msg


def get_or_create_profile(user_id: str, email: str) -> dict:
    """Load user profile from Supabase, creating it on first login."""
    sb = get_supabase()
    if not sb:
        return {"id": user_id, "email": email, "credits": FREE_CREDITS}
    try:
        res = sb.table("profiles").select("*").eq("id", user_id).execute()
        if res.data:
            return res.data[0]
        new_profile = {"id": user_id, "email": email, "credits": FREE_CREDITS}
        sb.table("profiles").insert(new_profile).execute()
        return new_profile
    except Exception:
        return {"id": user_id, "email": email, "credits": FREE_CREDITS}


def can_run_analysis(profile: dict) -> tuple:
    """Returns (allowed: bool, reason: str)."""
    credits = int(profile.get("credits", 0))
    if credits > 0:
        return True, ""
    return False, "You have no analyses left. Top up with a credit pack to continue."


def decrement_credits(user_id: str):
    """Deduct 1 credit after a successful analysis."""
    sb = get_supabase()
    if not sb or not user_id:
        return
    try:
        res = sb.table("profiles").select("credits").eq("id", user_id).execute()
        if res.data:
            new_credits = max(0, int(res.data[0].get("credits", 0)) - 1)
            sb.table("profiles").update({"credits": new_credits}).eq("id", user_id).execute()
            if "user_profile" in st.session_state:
                st.session_state["user_profile"]["credits"] = new_credits
    except Exception:
        pass


def add_credits(user_id: str, amount: int) -> bool:
    """Add credits to a user's account after a pack purchase."""
    sb = _fresh_supabase()
    if not sb:
        return False
    try:
        res = sb.table("profiles").select("credits").eq("id", user_id).execute()
        current = int(res.data[0].get("credits", 0)) if res.data else 0
        sb.table("profiles").update({"credits": current + amount}).eq("id", user_id).execute()
        return True
    except Exception:
        return False


def extract_text_from_pdf(file):
    try:
        pdf_reader = pypdf.PdfReader(file)
        return "".join(page.extract_text() for page in pdf_reader.pages)
    except Exception as e:
        return f"Error reading PDF: {str(e)}"


def extract_text_from_docx(file):
    try:
        file_bytes = file.getvalue() if hasattr(file, 'getvalue') else file.read()
        doc = Document(io.BytesIO(file_bytes))
        texts = []
        # Body paragraphs
        for para in doc.paragraphs:
            if para.text.strip():
                texts.append(para.text)
        # Table cells — many resume templates use tables for layout
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if para.text.strip():
                            texts.append(para.text)
        return "\n".join(texts)
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
    elif file_type == 'docx':
        return extract_text_from_docx(uploaded_file)
    elif file_type == 'doc':
        return "Error: Old .doc format is not supported. Please open your resume in Word, save it as .docx, then re-upload."
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

⚠️ KEYWORD EXACTNESS WARNING: Flag any case where the resume uses a vague or paraphrased term instead of the job ad's exact wording. ATS systems match literal strings — "SQL" ≠ "database querying", "Python" ≠ "scripting", "AWS" ≠ "cloud platforms", "Tableau" ≠ "visualisation tools". List each mismatch explicitly so the candidate knows to use the job ad's exact terminology.
- Flag missing semantic phrases where relevant: "end-to-end ownership", "data-driven decision making", "cross-functional collaboration", "stakeholder management", "scalable solutions", "production-ready systems", "business insights"
- Warn if resume shows signs of keyword stuffing or random keyword lists — modern ATS detects and penalises this
- Score bullets against the magic formula: Action Verb + Exact Skill + Quantified Impact + Business Outcome. Flag weak bullets that are missing impact or outcome.

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
- Always end the letter with:
  Kind regards,

  [applicant's full name from the resume]
  [applicant's phone number from the resume, if present]

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
End the letter with:
  Kind regards,

  [applicant's full name from the resume]
  [applicant's phone number from the resume, if present]
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


def extract_recommendations_summary(analysis_text: str) -> str:
    """Pull the SPECIFIC IMPROVEMENT RECOMMENDATIONS section out of the analysis."""
    m = re.search(
        r'##\s*SPECIFIC IMPROVEMENT RECOMMENDATIONS\s*\n(.*?)(?=\n##|\Z)',
        analysis_text, re.DOTALL | re.IGNORECASE
    )
    if m:
        return m.group(1).strip()[:1500]
    return ""


def update_application_status(record_id, new_status: str):
    """Update the status of a single application row in Supabase."""
    sb = get_supabase()
    if sb and record_id:
        sb.table("applications").update({"status": new_status}).eq("id", record_id).execute()
        load_tracker_data.clear()


def save_to_tracker(job_title: str, company: str, location: str,
                    resume_filename: str, match_pct: str, job_url: str,
                    cover_letter: str = '', cover_letter_path: str = '',
                    notes: str = '', updated_resume_file: str = '',
                    user_id: str = None):
    """Insert a job application row into Supabase."""
    sb = get_supabase()
    row = {
        "date":                datetime.now().strftime('%Y-%m-%d'),
        "job_title":           job_title,
        "company":             company,
        "location":            location,
        "resume_file":         resume_filename,
        "updated_resume_file": updated_resume_file,
        "match_pct":           match_pct,
        "job_url":             job_url if job_url != 'Manual Input' else '',
        "status":              "Applied",
        "notes":               notes,
        "cover_letter":        cover_letter_path if cover_letter_path else cover_letter,
        "user_id":             user_id,
    }
    if not sb:
        raise RuntimeError("Supabase is not configured — cannot save application.")
    result = sb.table("applications").insert(row).execute()
    if hasattr(result, 'error') and result.error:
        raise RuntimeError(f"Supabase insert failed: {result.error}")


@st.cache_data(ttl=30)
def load_tracker_data(user_id: str = None):
    """Load applications from Supabase for the given user, newest first."""
    sb = get_supabase()
    if not sb:
        return []
    try:
        query = sb.table("applications").select("*").order("created_at", desc=True)
        if user_id:
            query = query.eq("user_id", user_id)
        rows = query.execute()
        return rows.data or []
    except Exception:
        return []


def generate_tracker_excel(tracker_data: list) -> bytes:
    """Build an in-memory Excel workbook from tracker_data (list of dicts)."""
    headers  = ['Date', 'Job Title', 'Company', 'Location', 'Original Resume', 'Updated Resume', 'Match %', 'Job URL', 'Status', 'Cover Letter']
    col_keys = ['date', 'job_title', 'company', 'location', 'resume_file', 'updated_resume_file', 'match_pct', 'job_url', 'status', 'cover_letter']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Applications"

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font      = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill      = openpyxl.styles.PatternFill(fill_type="solid", fgColor="230344")
        cell.alignment = openpyxl.styles.Alignment(horizontal="center")

    widths = [14, 30, 25, 20, 35, 35, 10, 45, 15, 30]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width

    for i, item in enumerate(tracker_data, 2):
        for col, key in enumerate(col_keys, 1):
            ws.cell(row=i, column=col, value=item.get(key) or '')
        if i % 2 == 0:
            fill = openpyxl.styles.PatternFill(fill_type="solid", fgColor="F3F0F8")
            for col in range(1, len(headers) + 1):
                ws.cell(row=i, column=col).fill = fill

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============= RESUME UPDATER =============

def generate_resume_updates(resume_text: str, analysis_text: str, client, user_guidance: str = ""):
    """
    Ask Claude to return targeted find→replace pairs based on recommendations.
    Strictly forbidden from fabricating skills or experience not in the original resume.
    """
    guidance_block = ""
    if user_guidance and user_guidance.strip():
        guidance_block = (
            f"\n⭐ CANDIDATE GUIDANCE — apply this first, it takes priority over general rules:\n"
            f"{user_guidance.strip()}\n"
        )

    prompt = f"""You are a professional resume editor with a strict honesty policy.
{guidance_block}
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
- PRIORITISE these action verbs wherever they naturally fit existing content: Built, Designed, Developed, Automated, Optimised, Implemented, Led, Delivered, Reduced, Increased, Migrated, Integrated — replace weak or passive verbs with these where accurate
- If the user has provided metrics in the guidance (%, time saved, revenue, cost reduction, efficiency, scale), work them into the relevant bullet points
- Where a bullet point is missing a quantified result and one could logically exist, note this in the description field (e.g. "Consider adding a metric here — what % improvement or time saving resulted?") so the user knows where to strengthen further
- CRITICAL — technical keyword matching: if the job ad uses specific tool/technology names (e.g. "SQL", "Python", "Tableau", "AWS"), the resume must use those EXACT words — not synonyms or paraphrases. "SQL" not "database querying". "Python" not "scripting". "AWS" not "cloud platforms". "Tableau" not "visualisation tools". Where the candidate clearly has the skill but has used a vague term, replace it with the exact keyword from the job ad
- SEMANTIC PHRASES — where naturally supported by existing experience, weave in AI-scoring phrases: "end-to-end ownership", "data-driven decision making", "cross-functional collaboration", "stakeholder management", "scalable solutions", "production-ready systems", "business insights" — only where they accurately describe what the candidate did
- AVOID keyword stuffing, repeating the same keyword multiple times, or adding keyword lists — modern ATS systems detect and downgrade this
- THE MAGIC FORMULA for a strong bullet: Action Verb + Exact Skill/Tool + Quantified Impact + Business Outcome. Example: "Built automated Tableau dashboards using SQL and Python, reducing manual reporting time by 40% and improving stakeholder decision-making speed." Apply this structure to existing bullets wherever possible
- If additional guidance is provided, follow it only where it aligns with existing resume content
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

    def _sanitize_json(s: str) -> str:
        """Escape unescaped control characters inside JSON string values."""
        result = []
        in_string = False
        i = 0
        _escapes = {'\n': '\\n', '\r': '\\r', '\t': '\\t', '\b': '\\b', '\f': '\\f'}
        while i < len(s):
            c = s[i]
            if c == '\\' and in_string:
                result.append(c)
                i += 1
                if i < len(s):
                    result.append(s[i])
            elif c == '"':
                result.append(c)
                in_string = not in_string
            elif in_string and ord(c) < 0x20:
                result.append(_escapes.get(c, ''))
            else:
                result.append(c)
            i += 1
        return ''.join(result)

    try:
        message = client.messages.create(
            model=MODEL_NAME,
            max_tokens=3000,
            messages=[{"role": "user", "content": prompt}]
        )
        response_text = message.content[0].text
        json_match = re.search(r'```json\s*(.*?)\s*```', response_text, re.DOTALL)
        json_str = json_match.group(1) if json_match else response_text.strip()
        try:
            updates = json.loads(json_str)
        except json.JSONDecodeError:
            updates = json.loads(_sanitize_json(json_str))
        return {'success': True, 'updates': updates}
    except Exception as e:
        return {'success': False, 'error': str(e)}


def re_score_resume(updated_text: str, job_content: str, client) -> int | None:
    """Quick re-score of the updated resume. Returns integer % or None on failure."""
    try:
        message = client.messages.create(
            model=MODEL_NAME,
            max_tokens=10,
            messages=[{"role": "user", "content":
                f"Score this resume against the job posting. "
                f"Return ONLY a single integer 0-100 — the compatibility percentage. No other text.\n\n"
                f"RESUME:\n{updated_text[:4000]}\n\nJOB POSTING:\n{job_content[:4000]}"}]
        )
        val = re.search(r'\d+', message.content[0].text.strip())
        return min(int(val.group()), 100) if val else None
    except Exception:
        return None


def apply_updates_to_docx(file_bytes: bytes, updates: list, original_filename: str) -> bytes:
    """
    Apply find→replace pairs to the original DOCX, preserving all formatting.
    Searches both body paragraphs and table cells (many templates use tables for layout).
    Returns the updated DOCX as bytes.
    """
    doc = Document(io.BytesIO(file_bytes))
    applied = 0

    def _all_paragraphs(doc):
        yield from doc.paragraphs
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    yield from cell.paragraphs

    def _replace_in_para(para, find_text, replace_text):
        # Best case: find_text lives entirely within one run — replace only that run,
        # leaving every other run's formatting completely untouched.
        for run in para.runs:
            if find_text in run.text:
                run.text = run.text.replace(find_text, replace_text)
                return True

        # Fallback: text spans multiple runs — concatenate, replace, rebuild.
        if find_text not in para.text:
            return False

        new_text = para.text.replace(find_text, replace_text)

        # Pick the last non-empty content run as the formatting reference so we
        # inherit colour/size from the actual text, not from a structural run[0].
        ref_run = next(
            (r for r in reversed(para.runs) if r.text.strip()),
            para.runs[0] if para.runs else None,
        )

        if para.runs:
            tgt = para.runs[0]
            tgt.text = new_text
            if ref_run and ref_run is not tgt:
                try:
                    if ref_run.font.color.rgb is not None:
                        tgt.font.color.rgb = ref_run.font.color.rgb
                except Exception:
                    pass
                if ref_run.font.size:
                    tgt.font.size = ref_run.font.size
                if ref_run.font.name:
                    tgt.font.name = ref_run.font.name
                tgt.bold = ref_run.bold
                tgt.italic = ref_run.italic
            for run in para.runs[1:]:
                run.text = ''
        return True

    for update in updates:
        find_text = update.get('find', '').strip()
        replace_text = update.get('replace', '').strip()
        if not find_text or not replace_text:
            continue

        for para in _all_paragraphs(doc):
            if find_text in para.text:
                if _replace_in_para(para, find_text, replace_text):
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


def create_analysis_docx(analysis_text: str, job_url: str, resume_filename: str, job_content: str = '') -> bytes:
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
    meta.add_run('Resume: ').bold = True
    meta.add_run(resume_filename)
    meta2 = doc.add_paragraph()
    meta2.add_run('Job: ').bold = True
    meta2.add_run(job_url)
    meta3 = doc.add_paragraph()
    meta3.add_run('Date: ').bold = True
    meta3.add_run(datetime.now().strftime('%Y-%m-%d'))
    doc.add_paragraph()

    # Job description section
    if job_content and job_content.strip():
        doc.add_heading('Job Description', level=1)
        for para in job_content.strip().split('\n'):
            if para.strip():
                doc.add_paragraph(para.strip())
            else:
                doc.add_paragraph()
        doc.add_page_break()

    doc.add_heading('Analysis', level=1)
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


# ============= LANDING PAGE =============

def show_landing():
    """Marketing landing page — full-page iframe with sidebar trigger for navigation."""

    # Hide all Streamlit chrome and set dark background to match design
    st.markdown("""
<style>
section[data-testid="stSidebar"]{display:none!important;}
[data-testid="collapsedControl"]{display:none!important;}
header[data-testid="stHeader"]{display:none!important;}
#MainMenu{display:none!important;}
footer[data-testid="stFooter"]{display:none!important;}
.stApp,[data-testid="stApp"]{background:#071512!important;overflow:hidden!important;}
[data-testid="stAppViewContainer"],.stAppViewContainer{background:#071512!important;height:100vh!important;overflow:hidden!important;}
section.main,.main{background:#071512!important;padding:0!important;height:100vh!important;overflow:hidden!important;}
.block-container,[data-testid="stMainBlockContainer"]{background:#071512!important;padding:0!important;max-width:100%!important;margin:0!important;height:100vh!important;overflow:hidden!important;}
[data-testid="stVerticalBlock"]{height:100vh!important;gap:0!important;overflow:hidden!important;}
iframe[title="st.iframe"]{height:100vh!important;width:100vw!important;border:none!important;display:block!important;}
</style>""", unsafe_allow_html=True)

    # Hidden login trigger in sidebar (hidden by display:none CSS above).
    # Iframe JS finds this button by text and clicks it to trigger a Python-side rerun.
    with st.sidebar:
        if st.button("__rs_login__", key="rs_login_trigger"):
            st.session_state.show_login = True
            st.rerun()

    # Full landing page in isolated iframe — no Streamlit container constraints.
    _landing_doc = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>ResumeSync — know your match before you apply</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,400;12..96,600;12..96,700;12..96,800&family=Lora:ital,wght@1,500;1,600&family=Space+Mono:wght@400;700&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500&display=swap" rel="stylesheet">
<style>
:root{
--bg:#071512;--bg-2:#0a1b16;--panel:#0c2019;--panel-2:#0f271e;
--ink:#ecf4ee;--ink2:#9fb6a8;--ink3:#6e8a7b;
--mint:#7ad79f;--mint-bright:#94e6b1;--mint-deep:#4fae7a;
--mint-dim:rgba(122,215,159,0.10);--mint-dim2:rgba(122,215,159,0.16);
--mint-border:rgba(122,215,159,0.22);--amber:#e0a14a;
--line:rgba(159,182,168,0.12);--line-2:rgba(159,182,168,0.07);
--shadow-lg:0 24px 60px rgba(0,0,0,0.45),0 4px 14px rgba(0,0,0,0.35);
--mono:'Space Mono',ui-monospace,monospace;
--display:'Bricolage Grotesque',system-ui,sans-serif;
--serif:'Lora',Georgia,serif;--sans:'DM Sans',system-ui,sans-serif;
}
*{box-sizing:border-box;margin:0;padding:0;}
html{scroll-behavior:smooth;}
body{font-family:var(--sans);background:var(--bg);color:var(--ink);min-height:100vh;overflow-x:hidden;-webkit-font-smoothing:antialiased;position:relative;}
body::before{content:'';position:fixed;inset:0;z-index:0;pointer-events:none;background:radial-gradient(900px 600px at 78% -8%,rgba(122,215,159,0.10),transparent 60%),radial-gradient(700px 700px at 6% 4%,rgba(122,215,159,0.05),transparent 55%);}
body>*{position:relative;z-index:1;}
.eyebrow{font-family:var(--mono);font-size:12px;font-weight:400;letter-spacing:0.22em;text-transform:uppercase;color:var(--mint);display:inline-flex;align-items:center;gap:14px;}
.eyebrow::before,.eyebrow::after{content:'';width:26px;height:1px;background:var(--mint);opacity:0.5;}
.eyebrow.one::after{display:none;}
.nav{position:fixed;top:0;left:0;right:0;z-index:9999;display:flex;align-items:center;justify-content:space-between;padding:18px 56px;background:rgba(7,21,18,0.92);backdrop-filter:blur(14px);border-bottom:1px solid var(--line-2);}
body{padding-top:74px;}
.logo{display:flex;align-items:center;gap:12px;text-decoration:none;}
.logo-mark{width:38px;height:38px;border-radius:10px;background:linear-gradient(150deg,var(--mint),var(--mint-deep));display:grid;place-items:center;font-family:var(--display);font-size:18px;font-weight:800;color:#06140f;box-shadow:0 4px 14px rgba(122,215,159,0.30);}
.logo-txt{display:flex;flex-direction:column;line-height:1;}
.logo-name{font-family:var(--display);font-size:19px;font-weight:700;letter-spacing:-0.02em;color:var(--ink);}
.logo-name span{color:var(--mint);}
.logo-by{font-family:var(--mono);font-size:9.5px;letter-spacing:0.14em;text-transform:uppercase;color:var(--ink3);margin-top:3px;}
.nav-links{display:flex;align-items:center;gap:30px;}
.nav-link{font-size:14px;color:var(--ink2);text-decoration:none;font-weight:500;transition:color .15s;}
.nav-link:hover{color:var(--ink);}
.nav-cta{font-size:13.5px;font-weight:600;padding:10px 20px;border-radius:9px;background:var(--mint);color:#06140f;text-decoration:none;transition:background .15s,transform .15s;box-shadow:0 4px 14px rgba(122,215,159,0.22);}
.nav-cta:hover{background:var(--mint-bright);transform:translateY(-1px);}
.hero{padding:78px 56px 96px;max-width:1280px;margin:0 auto;display:grid;grid-template-columns:1fr 472px;gap:60px;align-items:center;}
h1{font-family:var(--display);font-weight:800;font-size:clamp(38px,4.1vw,60px);line-height:1.0;letter-spacing:-0.03em;color:var(--ink);margin:22px 0 24px;}
h1 .italic{font-family:var(--serif);font-weight:600;font-style:italic;color:var(--mint);letter-spacing:-0.01em;}
.hero-lede{font-size:18px;line-height:1.66;color:var(--ink2);max-width:50ch;margin-bottom:36px;}
.hero-lede strong{color:var(--ink);font-weight:500;}
.hero-btns{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:34px;}
.btn-primary{font-size:15px;font-weight:600;padding:15px 28px;border-radius:11px;background:var(--mint);color:#06140f;text-decoration:none;box-shadow:0 10px 30px rgba(122,215,159,0.26);transition:background .15s,transform .15s;display:inline-block;}
.btn-primary:hover{background:var(--mint-bright);transform:translateY(-1px);}
.btn-ghost{font-size:15px;font-weight:500;padding:15px 26px;border-radius:11px;border:1px solid var(--line);color:var(--ink);text-decoration:none;background:rgba(255,255,255,0.015);transition:border-color .15s,background .15s;display:inline-block;}
.btn-ghost:hover{border-color:var(--mint-border);background:var(--mint-dim);}
.hero-social{display:flex;align-items:center;gap:18px;flex-wrap:wrap;}
.social-item{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--ink3);}
.social-item b{color:var(--ink2);font-weight:600;}
.social-item .tick{color:var(--mint);}
.social-div{width:1px;height:14px;background:var(--line);}
.hero-visual{position:relative;}
.photo-frame{position:relative;border-radius:20px;overflow:hidden;border:1px solid var(--line);box-shadow:var(--shadow-lg);aspect-ratio:5/6;background:#0b1e18 center/cover no-repeat;}
.photo-frame img{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;display:block;}
.photo-frame::after{content:'';position:absolute;inset:0;pointer-events:none;background:linear-gradient(to top,rgba(6,17,13,0.92) 0%,rgba(6,17,13,0.55) 22%,rgba(6,17,13,0) 46%);z-index:2;}
.photo-badge{position:absolute;top:18px;left:18px;z-index:3;font-family:var(--mono);font-size:10.5px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;padding:7px 12px;border-radius:7px;background:rgba(6,17,13,0.78);color:var(--mint);border:1px solid var(--mint-border);backdrop-filter:blur(4px);display:inline-flex;align-items:center;gap:8px;}
.pulse{width:6px;height:6px;border-radius:50%;background:var(--mint);animation:pulse 2.2s infinite;}
@keyframes pulse{0%{box-shadow:0 0 0 0 rgba(122,215,159,0.5);}70%{box-shadow:0 0 0 8px rgba(122,215,159,0);}100%{box-shadow:0 0 0 0 rgba(122,215,159,0);}}
.photo-name{position:absolute;left:22px;bottom:20px;z-index:3;}
.photo-name .nm{font-family:var(--display);font-size:21px;font-weight:700;color:#fff;letter-spacing:-0.01em;}
.photo-name .role{font-family:var(--mono);font-size:10.5px;letter-spacing:0.12em;text-transform:uppercase;color:var(--mint);margin-top:6px;}
.float{position:absolute;z-index:5;background:rgba(10,27,21,0.82);border:1px solid var(--mint-border);border-radius:16px;box-shadow:0 12px 30px rgba(0,0,0,0.30);backdrop-filter:blur(10px);}
.float-score{top:-44px;right:-34px;padding:18px 20px 16px;width:172px;text-align:center;}
.fs-label{font-family:var(--mono);font-size:9.5px;letter-spacing:0.16em;text-transform:uppercase;color:var(--ink3);margin-bottom:10px;}
.ring-wrap{position:relative;width:108px;height:108px;margin:0 auto;}
.ring-wrap svg{transform:rotate(-90deg);}
.ring-num{position:absolute;inset:0;display:grid;place-items:center;font-family:var(--display);font-weight:800;font-size:30px;color:var(--mint);}
.ring-num small{font-size:14px;color:var(--ink2);margin-left:1px;}
.fs-fit{font-family:var(--mono);font-size:10px;letter-spacing:0.14em;text-transform:uppercase;color:var(--mint);margin-top:10px;}
.float-skills{bottom:64px;right:-44px;padding:16px 18px;width:228px;}
.fsk-label{font-family:var(--mono);font-size:9.5px;letter-spacing:0.16em;text-transform:uppercase;color:var(--ink3);margin-bottom:12px;}
.skill-row{display:grid;grid-template-columns:64px 1fr 26px;align-items:center;gap:10px;margin-bottom:9px;}
.skill-row:last-child{margin-bottom:0;}
.skill-name{font-size:11.5px;color:var(--ink2);font-weight:500;}
.skill-track{height:5px;border-radius:99px;background:rgba(159,182,168,0.14);overflow:hidden;}
.skill-fill{height:100%;border-radius:99px;background:linear-gradient(90deg,var(--mint-deep),var(--mint));}
.skill-val{font-family:var(--mono);font-size:11px;color:var(--mint);text-align:right;}
.section{max-width:1180px;margin:0 auto;padding:96px 56px;}
.band{background:var(--bg-2);border-top:1px solid var(--line-2);border-bottom:1px solid var(--line-2);}
.sec-head{text-align:center;margin-bottom:62px;}
h2{font-family:var(--display);font-weight:800;font-size:clamp(34px,3.8vw,52px);letter-spacing:-0.025em;line-height:1.04;margin:16px 0 0;}
h2 .italic{font-family:var(--serif);font-weight:600;font-style:italic;color:var(--mint);}
.sec-sub{color:var(--ink2);font-size:16.5px;max-width:54ch;margin:18px auto 0;line-height:1.6;}
.sec-head .eyebrow{justify-content:center;}
.features-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:18px;}
.feature-card{background:var(--panel);border:1px solid var(--line);border-radius:18px;padding:28px;transition:transform .2s,border-color .2s,background .2s;position:relative;overflow:hidden;}
.feature-card:hover{transform:translateY(-4px);border-color:var(--mint-border);background:var(--panel-2);}
.feature-num{font-family:var(--mono);font-size:11px;color:var(--ink3);letter-spacing:0.1em;margin-bottom:18px;}
.feature-card h3{font-family:var(--display);font-size:18px;font-weight:700;color:var(--ink);margin-bottom:10px;letter-spacing:-0.01em;}
.feature-card p{font-size:13.5px;line-height:1.62;color:var(--ink2);}
.feature-tag{display:inline-block;margin-top:16px;font-family:var(--mono);font-size:10px;letter-spacing:0.08em;text-transform:uppercase;padding:4px 10px;border-radius:99px;background:var(--mint-dim);color:var(--mint);border:1px solid var(--mint-border);}
.steps{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;}
.step{background:var(--panel);border:1px solid var(--line);border-radius:18px;padding:26px 22px;position:relative;transition:border-color .2s,background .2s;}
.step:hover{border-color:var(--mint-border);background:var(--panel-2);}
.step-num{font-family:var(--display);font-size:15px;font-weight:800;color:#06140f;width:34px;height:34px;border-radius:9px;background:var(--mint);display:grid;place-items:center;margin-bottom:18px;}
.step h3{font-family:var(--display);font-size:16px;font-weight:700;color:var(--ink);margin-bottom:8px;}
.step p{font-size:13px;line-height:1.58;color:var(--ink2);}
.pricing-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:18px;max-width:960px;margin:0 auto;align-items:start;}
.price-card{background:var(--panel);border:1px solid var(--line);border-radius:20px;padding:30px 26px;position:relative;}
.price-card.featured{background:linear-gradient(170deg,#103025,#0a1f18);border-color:var(--mint-border);box-shadow:0 24px 60px rgba(0,0,0,0.4),0 0 0 1px rgba(122,215,159,0.10) inset;transform:translateY(-10px);}
.price-badge{position:absolute;top:-12px;left:50%;transform:translateX(-50%);font-family:var(--mono);font-size:9.5px;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;background:var(--mint);color:#06140f;padding:5px 14px;border-radius:99px;white-space:nowrap;box-shadow:0 6px 16px rgba(122,215,159,0.3);}
.price-plan{font-family:var(--display);font-size:18px;font-weight:700;margin-bottom:6px;}
.price-desc{font-size:12.5px;color:var(--ink3);line-height:1.5;margin-bottom:22px;min-height:38px;}
.price-num{font-family:var(--display);font-size:50px;font-weight:800;line-height:1;display:flex;align-items:flex-start;gap:3px;}
.price-cur{font-size:20px;margin-top:9px;color:var(--ink2);font-weight:600;}
.price-period{font-family:var(--mono);font-size:11px;color:var(--ink3);letter-spacing:0.04em;margin:8px 0 24px;}
.price-cta{display:block;text-align:center;width:100%;padding:13px 0;border-radius:10px;font-size:14px;font-weight:600;text-decoration:none;cursor:pointer;margin-bottom:24px;transition:background .15s,transform .15s,border-color .15s;}
.cta-fill{background:var(--mint);color:#06140f;}
.cta-fill:hover{background:var(--mint-bright);transform:translateY(-1px);}
.cta-out{background:transparent;border:1px solid var(--line);color:var(--ink);}
.cta-out:hover{border-color:var(--mint-border);background:var(--mint-dim);}
.price-features{list-style:none;display:flex;flex-direction:column;gap:11px;}
.pf{display:flex;align-items:flex-start;gap:10px;font-size:13px;line-height:1.4;color:var(--ink2);}
.pf b{color:var(--ink);font-weight:600;}
.ck{flex:0 0 17px;width:17px;height:17px;border-radius:50%;display:grid;place-items:center;font-size:9px;margin-top:1px;}
.ck.y{background:var(--mint-dim2);color:var(--mint);}
.ck.n{background:rgba(159,182,168,0.08);color:var(--ink3);}
.cta-strip{max-width:1180px;margin:0 auto;padding:0 56px 96px;}
.cta-inner{background:linear-gradient(150deg,#11362a,#0a1f18);border:1px solid var(--mint-border);border-radius:24px;padding:56px;text-align:center;position:relative;overflow:hidden;}
.cta-inner::before{content:'';position:absolute;inset:0;background:radial-gradient(600px 300px at 50% -30%,rgba(122,215,159,0.16),transparent 70%);pointer-events:none;}
.cta-inner h2{position:relative;}
.cta-inner p{position:relative;color:var(--ink2);font-size:16px;margin:16px auto 30px;max-width:46ch;line-height:1.6;}
.cta-inner .hero-btns{justify-content:center;margin-bottom:0;position:relative;}
.rs-footer{border-top:1px solid var(--line-2);padding:38px 56px;max-width:1280px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;gap:20px;flex-wrap:wrap;}
.footer-logo{display:flex;align-items:center;gap:10px;font-family:var(--display);font-weight:700;font-size:15px;color:var(--ink);}
.footer-mark{width:28px;height:28px;border-radius:7px;background:linear-gradient(150deg,var(--mint),var(--mint-deep));display:grid;place-items:center;font-size:12px;font-weight:800;color:#06140f;}
.footer-note{font-family:var(--mono);font-size:11px;color:var(--ink3);letter-spacing:0.03em;}
.footer-note a{color:var(--mint);text-decoration:none;}
@media(max-width:1080px){.hero{grid-template-columns:1fr;gap:64px;padding-bottom:80px;}.hero-visual{max-width:460px;}.float-skills{right:-20px;}.float-score{right:-14px;}.features-grid{grid-template-columns:1fr 1fr;}.steps{grid-template-columns:1fr 1fr;}.pricing-grid{grid-template-columns:1fr;max-width:440px;}.price-card.featured{transform:none;}}
@media(max-width:640px){.nav{padding:14px 20px;}.nav-links{display:none;}.hero{padding:44px 20px 56px;}.section{padding:64px 20px;}.cta-strip{padding:0 20px 64px;}.cta-inner{padding:38px 24px;}.features-grid,.steps{grid-template-columns:1fr;}.float-score{top:-18px;right:-8px;transform:scale(0.88)}.float-skills{right:-8px;bottom:20px;transform:scale(0.9)}.rs-footer{padding:28px 20px;flex-direction:column;text-align:center;}}
</style>
</head>
<body>
<nav class="nav">
<a class="logo" href="#">
<div class="logo-mark">R</div>
<div class="logo-txt">
<div class="logo-name">Resume<span>Sync</span></div>
<div class="logo-by">by VisualizePro</div>
</div>
</a>
<div class="nav-links">
<a class="nav-link" href="#features">Features</a>
<a class="nav-link" href="#how">How it works</a>
<a class="nav-link" href="#pricing">Pricing</a>
</div>
<a class="nav-cta cta-login" id="nav-cta-btn" href="#">Log in &#8594;</a>
</nav>
<section class="hero">
<div class="hero-copy">
<div class="eyebrow">know before you apply</div>
<h1>Stop applying blindly.<br><span class="italic">Know your match</span><br>before you apply.</h1>
<p class="hero-lede">Upload your resume, paste any job posting &#8212; and in seconds know your exact <strong>match score</strong>, which skills are missing, which keywords get you <strong>past ATS</strong>, and walk away with a tailored cover letter and resume ready to send.</p>
<div class="hero-btns">
<a class="btn-primary cta-login" href="#">Start free &#8212; no card needed</a>
<a class="btn-ghost" href="#how">See how it works &#8594;</a>
</div>
<div class="hero-social">
<div class="social-item"><span class="tick">&#10003;</span><b>3</b> free applications</div>
<div class="social-div"></div>
<div class="social-item"><span class="tick">&#10003;</span><b>AI</b> resume rewrites</div>
<div class="social-div"></div>
<div class="social-item"><span class="tick">&#10003;</span>Cover letter <b>included</b></div>
</div>
</div>
<div class="hero-visual">
<div class="photo-frame">
<img src="https://images.unsplash.com/photo-1573496359142-b8d87734a5a2?w=900&q=80&auto=format&fit=crop" alt="Professional" loading="lazy"/>
<div class="photo-badge"><span class="pulse"></span>Now hiring &middot; You</div>
<div class="photo-name">
<div class="nm">Madeli, BI Specialist</div>
<div class="role">Currently matching to 12 open roles</div>
</div>
</div>
<div class="float float-score">
<div class="fs-label">Match Score</div>
<div class="ring-wrap">
<svg width="108" height="108" viewBox="0 0 108 108">
<circle cx="54" cy="54" r="48" fill="none" stroke="rgba(159,182,168,0.16)" stroke-width="8"/>
<circle cx="54" cy="54" r="48" fill="none" stroke="var(--mint)" stroke-width="8" stroke-linecap="round" stroke-dasharray="301.6" stroke-dashoffset="24.1"/>
</svg>
<div class="ring-num">92<small>%</small></div>
</div>
<div class="fs-fit">Strong fit</div>
</div>
<div class="float float-skills">
<div class="fsk-label">Skill alignment</div>
<div class="skill-row"><span class="skill-name">Tableau</span><span class="skill-track"><span class="skill-fill" style="width:96%"></span></span><span class="skill-val">96</span></div>
<div class="skill-row"><span class="skill-name">Power BI</span><span class="skill-track"><span class="skill-fill" style="width:82%"></span></span><span class="skill-val">82</span></div>
<div class="skill-row"><span class="skill-name">SQL</span><span class="skill-track"><span class="skill-fill" style="width:90%"></span></span><span class="skill-val">90</span></div>
<div class="skill-row"><span class="skill-name">Snowflake</span><span class="skill-track"><span class="skill-fill" style="width:74%"></span></span><span class="skill-val">74</span></div>
</div>
</div>
</section>
<div class="band">
<section class="section" id="features">
<div class="sec-head">
<div class="eyebrow one">what resumesync does</div>
<h2>Every tool you need to<br><span class="italic">land the job.</span></h2>
<p class="sec-sub">From pasting a job posting to sending your application &#8212; one workflow handles the score, the rewrite, the cover letter, and the tracking.</p>
</div>
<div class="features-grid">
<div class="feature-card"><div class="feature-num">01</div><h3>Match Score</h3><p>Paste any job description and get an instant percentage showing how well your resume fits &#8212; scored on skills, keywords, experience, and tone.</p><span class="feature-tag">Instant results</span></div>
<div class="feature-card"><div class="feature-num">02</div><h3>Skill &amp; keyword gaps</h3><p>See exactly what&#8217;s missing &#8212; which keywords to add, which experience to emphasise, which sections to rewrite. Specific, actionable, never vague.</p><span class="feature-tag">AI-powered</span></div>
<div class="feature-card"><div class="feature-num">03</div><h3>Resume rewriter</h3><p>One click tailors your resume to the job &#8212; sharpening language and lifting your ATS score without inventing anything or losing your voice.</p><span class="feature-tag">Claude AI</span></div>
<div class="feature-card"><div class="feature-num">04</div><h3>Cover letter generator</h3><p>A tailored cover letter built from the posting and your updated resume. Sounds human, not robotic &#8212; and you can refine it with your own prompts.</p><span class="feature-tag">Tailored per role</span></div>
<div class="feature-card"><div class="feature-num">05</div><h3>Prompt &amp; refine</h3><p>Not happy with the output? Tell it what to change. &#8220;Make it more senior.&#8221; &#8220;Emphasise leadership.&#8221; You always direct the final result.</p><span class="feature-tag">You direct the AI</span></div>
<div class="feature-card"><div class="feature-num">06</div><h3>Application tracker</h3><p>Every application in one place &#8212; role, company, match score, resume version, cover letter, and status. Your whole job search, organised.</p><span class="feature-tag">Never lose track</span></div>
</div>
</section>
</div>
<section class="section" id="how">
<div class="sec-head">
<div class="eyebrow one">how it works</div>
<h2>Four steps to a<br><span class="italic">stronger application.</span></h2>
<p class="sec-sub">From job posting to ready-to-send application in under ten minutes.</p>
</div>
<div class="steps">
<div class="step"><div class="step-num">1</div><h3>Paste the job</h3><p>Drop in a Seek or LinkedIn URL, or paste the description. We scrape the rest.</p></div>
<div class="step"><div class="step-num">2</div><h3>Get your score</h3><p>See your match percentage instantly, with a clear list of gaps to close.</p></div>
<div class="step"><div class="step-num">3</div><h3>Rewrite &amp; refine</h3><p>Let AI rewrite your resume and cover letter, then tweak with your own prompts.</p></div>
<div class="step"><div class="step-num">4</div><h3>Track &amp; apply</h3><p>Save it all to your tracker and send. Every document stored, every status logged.</p></div>
</div>
</section>
<div class="band">
<section class="section" id="pricing">
<div class="sec-head">
<div class="eyebrow one">simple pricing</div>
<h2>Buy what you need.<br><span class="italic">No subscription.</span></h2>
<p class="sec-sub">Start with 2 free analyses. Top up whenever you like &#8212; credits never expire.</p>
</div>
<div class="pricing-grid">
<div class="price-card">
<div class="price-plan">Free</div>
<div class="price-desc">Try it on your first two applications. No credit card needed.</div>
<div class="price-num"><span class="price-cur">A$</span>0</div>
<div class="price-period">2 analyses included</div>
<a class="price-cta cta-out cta-login" data-plan="free" href="#">Get started &#8594;</a>
<ul class="price-features">
<li class="pf"><span class="ck y">&#10003;</span><span>2 analyses included</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Match score + gap report</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Resume rewriter</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Cover letter generator</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Application tracker</span></li>
</ul>
</div>
<div class="price-card featured">
<div class="price-badge">Most popular</div>
<div class="price-plan">Value Pack</div>
<div class="price-desc">Great value for an active job search. Credits never expire.</div>
<div class="price-num"><span class="price-cur">A$</span>15</div>
<div class="price-period">20 analyses &middot; one-off</div>
<a class="price-cta cta-fill cta-login" data-plan="pack20" href="#">Buy 20 analyses &#8594;</a>
<ul class="price-features">
<li class="pf"><span class="ck y">&#10003;</span><span>20 analyses</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Match score + gap report</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Resume rewriter</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Cover letter generator</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Application tracker</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Credits never expire</span></li>
</ul>
</div>
<div class="price-card">
<div class="price-plan">Starter Pack</div>
<div class="price-desc">Enough for a focused sprint of applications.</div>
<div class="price-num"><span class="price-cur">A$</span>9</div>
<div class="price-period">5 analyses &middot; one-off</div>
<a class="price-cta cta-out cta-login" data-plan="pack5" href="#">Buy 5 analyses &#8594;</a>
<ul class="price-features">
<li class="pf"><span class="ck y">&#10003;</span><span>5 analyses</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Match score + gap report</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Resume rewriter</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Cover letter generator</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Application tracker</span></li>
<li class="pf"><span class="ck y">&#10003;</span><span>Credits never expire</span></li>
</ul>
</div>
</div>
<div style="text-align:center;margin-top:1.5rem;">
<p style="font-family:'DM Sans',sans-serif;font-size:14px;color:#6e8a7b;margin:0 0 0.75rem;">
  Need more? <strong style="color:#9fb6a8;">Pro Pack — 30 analyses for A$20</strong> &middot; <strong style="color:#9fb6a8;">Max Pack — 50 analyses for A$30</strong>
</p>
<a class="price-cta cta-out cta-login" data-plan="pack30" href="#" style="display:inline-block;margin:0 6px 6px;">Pro Pack A$20 &#8594;</a>
<a class="price-cta cta-out cta-login" data-plan="pack50" href="#" style="display:inline-block;margin:0 6px 6px;">Max Pack A$30 &#8594;</a>
</div>
</section>
</div>
<div class="cta-strip" id="cta">
<div class="cta-inner">
<div class="eyebrow one" style="justify-content:center;">start in 60 seconds</div>
<h2 style="margin-top:14px;">Know your match.<br><span class="italic">Then apply with confidence.</span></h2>
<p>Your first two applications are free. Upload a resume, paste a job, and see your score &#8212; no card, no commitment.</p>
<div class="hero-btns">
<a class="btn-primary cta-login" href="#">Start free &#8212; no card needed</a>
<a class="btn-ghost" href="#features">Explore features &#8594;</a>
</div>
</div>
</div>
<div class="rs-footer">
<div class="footer-logo"><div class="footer-mark">R</div>ResumeSync</div>
<div class="footer-note">Built by <a href="https://visualizepro.com.au" target="_blank">VisualizePro</a> &middot; Sydney, Australia &middot; Powered by Claude AI</div>
</div>
<script>
(function() {
  // Resize iframe to match parent viewport so no double-scrollbar confusion
  function resizeToViewport() {
    try {
      var frames = window.parent.document.querySelectorAll('iframe');
      for (var i = 0; i < frames.length; i++) {
        if (frames[i].contentWindow === window) {
          frames[i].style.height = (window.parent.innerHeight || 768) + 'px';
          break;
        }
      }
    } catch(e) {}
  }
  resizeToViewport();
  window.addEventListener('resize', resizeToViewport);

  // Navigate parent page to login, optionally with a plan
  function triggerLogin(plan) {
    var planParam = plan ? '&plan=' + encodeURIComponent(plan) : '';
    // Method 1: direct parent location (works for same-origin user-gesture in most browsers)
    try {
      window.parent.location.href = window.parent.location.pathname + '?login=1' + planParam;
      return;
    } catch(e1) {}
    // Method 2: click the hidden trigger button in the parent Streamlit DOM
    try {
      var btns = window.parent.document.querySelectorAll('button');
      for (var i = 0; i < btns.length; i++) {
        if (btns[i].textContent.includes('__rs_login__')) {
          btns[i].click();
          return;
        }
      }
    } catch(e2) {}
    // Method 3: navigate same tab as last resort
    try {
      window.location.href = '?login=1' + planParam;
    } catch(e3) {}
  }

  // Single event-delegation listener — handles all clicks in the document
  document.addEventListener('click', function(e) {
    // Walk up from clicked element to find the <a> tag
    var el = e.target;
    while (el && el !== document.body) {
      if (el.tagName === 'A') {
        var href = el.getAttribute('href') || '';
        if (el.classList.contains('cta-login')) {
          e.preventDefault();
          var plan = el.getAttribute('data-plan') || '';
          triggerLogin(plan);
          return;
        }
        if (href.startsWith('#') && href.length > 1) {
          e.preventDefault();
          var target = document.getElementById(href.slice(1));
          if (target) target.scrollIntoView({ behavior: 'smooth', block: 'start' });
          return;
        }
        break;
      }
      el = el.parentElement;
    }
  });
})();
</script>
</body>
</html>"""

    _components.html(_landing_doc, height=850, scrolling=True)



# ============= LOGIN PAGE =============

def show_login_page():
    """Email + password login / signup screen."""
    _, col, _ = st.columns([1, 1.6, 1])
    with col:
        # Toggle between sign-in and sign-up
        mode = st.session_state.get("login_mode", "signin")

        st.markdown(f"""
        <div style="padding:3rem 0 1.5rem;">
          <div style="display:flex;align-items:center;gap:12px;margin-bottom:2.5rem;">
            <div style="width:44px;height:44px;border-radius:12px;
                        background:linear-gradient(150deg,#7ad79f,#4fae7a);
                        display:grid;place-items:center;
                        font-family:'Bricolage Grotesque',sans-serif;
                        font-weight:800;font-size:22px;color:#06140f;
                        box-shadow:0 4px 16px rgba(122,215,159,0.30);">R</div>
            <div>
              <div style="font-family:'Bricolage Grotesque',sans-serif;font-weight:700;
                          font-size:20px;color:#ecf4ee;letter-spacing:-0.02em;">ResumeSync</div>
              <div style="font-family:'Space Mono',monospace;font-size:9px;
                          letter-spacing:0.18em;text-transform:uppercase;color:#6e8a7b;">by VisualizePro</div>
            </div>
          </div>
          <h1 style="font-family:'Bricolage Grotesque',sans-serif;font-weight:800;
                     font-size:30px;color:#ecf4ee;margin:0 0 0.4rem;letter-spacing:-0.02em;">
            {"Create account" if mode == "signup" else "Sign in"}
          </h1>
          <p style="font-family:'DM Sans',sans-serif;font-size:15px;color:#9fb6a8;margin:0 0 2rem;">
            {"Enter your email and choose a password." if mode == "signup" else "Welcome back — enter your email and password."}
          </p>
        </div>
        """, unsafe_allow_html=True)

        email_input = st.text_input(
            "Email address",
            placeholder="you@example.com",
            label_visibility="collapsed",
            key="login_email_input",
        )
        password_input = st.text_input(
            "Password",
            placeholder="Password (min 6 characters)",
            type="password",
            label_visibility="collapsed",
            key="login_password_input",
        )

        if mode == "signup":
            if st.button("Create account →", type="primary", use_container_width=True, key="signup_btn"):
                email_val    = (email_input or "").strip().lower()
                password_val = (password_input or "").strip()
                if not email_val or "@" not in email_val:
                    st.error("Please enter a valid email address.")
                elif len(password_val) < 6:
                    st.error("Password must be at least 6 characters.")
                else:
                    info, err = sign_up(email_val, password_val)
                    if info:
                        st.session_state["auth_user_id"]  = info["user_id"]
                        st.session_state["auth_email"]    = info["email"]
                        st.session_state["auth_token"]    = info["access_token"]
                        st.session_state["auth_refresh"]  = info["refresh_token"]
                        st.session_state["user_profile"]  = get_or_create_profile(info["user_id"], info["email"])
                        st.session_state.pop("login_mode", None)
                        st.rerun()
                    else:
                        st.error(err)
            st.markdown("<div style='height:0.5rem;'></div>", unsafe_allow_html=True)
            if st.button("Already have an account? Sign in", key="switch_to_signin"):
                st.session_state["login_mode"] = "signin"
                st.rerun()
        else:
            if st.button("Sign in →", type="primary", use_container_width=True, key="signin_btn"):
                email_val    = (email_input or "").strip().lower()
                password_val = (password_input or "").strip()
                if not email_val or "@" not in email_val:
                    st.error("Please enter a valid email address.")
                elif not password_val:
                    st.error("Please enter your password.")
                else:
                    info, err = sign_in(email_val, password_val)
                    if info:
                        st.session_state["auth_user_id"]  = info["user_id"]
                        st.session_state["auth_email"]    = info["email"]
                        st.session_state["auth_token"]    = info["access_token"]
                        st.session_state["auth_refresh"]  = info["refresh_token"]
                        st.session_state["user_profile"]  = get_or_create_profile(info["user_id"], info["email"])
                        st.session_state.pop("login_mode", None)
                        st.rerun()
                    else:
                        st.error(err)
            st.markdown("<div style='height:0.5rem;'></div>", unsafe_allow_html=True)
            if st.button("No account yet? Create one free", key="switch_to_signup"):
                st.session_state["login_mode"] = "signup"
                st.rerun()

        st.markdown("""
        <div style="margin-top:2.5rem;padding-top:1.5rem;border-top:1px solid rgba(159,182,168,0.12);">
          <p style="font-family:'DM Sans',sans-serif;font-size:13px;color:#6e8a7b;margin:0 0 1.2rem;">
            <b style="color:#9fb6a8;">Free:</b> 2 analyses included ·
            <b style="color:#9fb6a8;">Packs from A$9</b> — no subscription, never expires
          </p>
          <a href="https://madelivniekerk.github.io/Resumesync"
             style="font-family:'DM Sans',sans-serif;font-size:13px;color:#6e8a7b;
                    text-decoration:none;display:inline-flex;align-items:center;gap:4px;
                    transition:color 0.2s;"
             onmouseover="this.style.color='#9fb6a8'"
             onmouseout="this.style.color='#6e8a7b'">
            ← Back to main page
          </a>
        </div>
        """, unsafe_allow_html=True)


# ============= PAYMENT PAGES =============

def show_payment_redirect(pack: str):
    """Shown after login when user selected a pack — redirect them to Pay Advanced checkout."""
    info = PACK_INFO.get(pack, {})
    if not info:
        st.rerun()
        return

    auth_email = st.session_state.get("auth_email", "")
    _ref = f"{auth_email}|{pack}"[:50]
    pay_params = urllib.parse.urlencode({
        "paymentamount":      info["amount"],
        "paymentdescription": f"ResumeSync {info['name']} ({info['credits']} analyses)",
        "paymentref":         _ref,
        "email":              auth_email,
    })
    pay_url = PAY_ADVANCED_URL + "?" + pay_params

    st.markdown("""
<style>
[data-testid="stHeader"]{display:none!important;}
#MainMenu{visibility:hidden!important;}
footer{visibility:hidden!important;}
</style>""", unsafe_allow_html=True)

    _, col, _ = st.columns([1, 1.6, 1])
    with col:
        st.markdown(f"""
<div style="padding:3rem 0 1rem;">
  <div style="display:flex;align-items:center;gap:12px;margin-bottom:2.5rem;">
    <div style="width:44px;height:44px;border-radius:12px;
                background:linear-gradient(150deg,#7ad79f,#4fae7a);
                display:grid;place-items:center;
                font-family:'Bricolage Grotesque',sans-serif;
                font-weight:800;font-size:22px;color:#06140f;
                box-shadow:0 4px 16px rgba(122,215,159,0.30);">R</div>
    <div>
      <div style="font-family:'Bricolage Grotesque',sans-serif;font-weight:700;
                  font-size:20px;color:#ecf4ee;letter-spacing:-0.02em;">ResumeSync</div>
      <div style="font-family:'Space Mono',monospace;font-size:9px;
                  letter-spacing:0.18em;text-transform:uppercase;color:#6e8a7b;">by VisualizePro</div>
    </div>
  </div>
  <div style="background:#0c2019;border:1px solid rgba(159,182,168,0.12);
              border-radius:14px;padding:1.8rem;margin-bottom:1.5rem;">
    <div style="font-family:'Space Mono',monospace;font-size:10px;letter-spacing:0.14em;
                text-transform:uppercase;color:{info['color']};margin-bottom:0.5rem;">Selected pack</div>
    <div style="font-family:'Bricolage Grotesque',sans-serif;font-weight:800;
                font-size:26px;color:#ecf4ee;margin-bottom:0.25rem;">{info['name']}</div>
    <div style="font-family:'DM Sans',sans-serif;font-size:22px;font-weight:700;
                color:{info['color']};margin-bottom:0.5rem;">{info['price']} one-off</div>
    <div style="font-family:'DM Sans',sans-serif;font-size:14px;color:#9fb6a8;">{info['credits']} analyses · never expires · no subscription</div>
  </div>
  <p style="font-family:'DM Sans',sans-serif;font-size:14px;color:#9fb6a8;margin:0 0 1.5rem;">
    You're signed in as <strong style="color:#ecf4ee;">{auth_email}</strong>.
    Click below to complete your payment — you'll be redirected back automatically.
  </p>
</div>""", unsafe_allow_html=True)

        st.markdown(f"""
<a href="{pay_url}" target="_self" style="
  display:block;width:100%;padding:0.65rem 1rem;
  background:linear-gradient(135deg,#7ad79f,#4fae7a);
  color:#06140f;font-family:'DM Sans',sans-serif;font-weight:700;
  font-size:15px;text-align:center;border-radius:9px;
  text-decoration:none;box-sizing:border-box;
  box-shadow:0 4px 14px rgba(122,215,159,0.30);">
  Pay {info['price']} — get {info['credits']} analyses →
</a>""", unsafe_allow_html=True)
        st.markdown("<div style='height:0.75rem;'></div>", unsafe_allow_html=True)
        if st.button("Skip — use my free credits for now", key="skip_payment_btn", use_container_width=True):
            st.rerun()


def show_payment_pending():
    """Shown after returning from Pay Advanced while the webhook processes."""
    st.markdown("""
<style>
[data-testid="stHeader"]{display:none!important;}
#MainMenu{visibility:hidden!important;}
footer{visibility:hidden!important;}
</style>""", unsafe_allow_html=True)

    _, col, _ = st.columns([1, 1.6, 1])
    with col:
        st.markdown("""
<div style="padding:3rem 0 1rem;">
  <div style="display:flex;align-items:center;gap:12px;margin-bottom:2.5rem;">
    <div style="width:44px;height:44px;border-radius:12px;
                background:linear-gradient(150deg,#7ad79f,#4fae7a);
                display:grid;place-items:center;
                font-family:'Bricolage Grotesque',sans-serif;
                font-weight:800;font-size:22px;color:#06140f;
                box-shadow:0 4px 16px rgba(122,215,159,0.30);">R</div>
    <div>
      <div style="font-family:'Bricolage Grotesque',sans-serif;font-weight:700;
                  font-size:20px;color:#ecf4ee;letter-spacing:-0.02em;">ResumeSync</div>
      <div style="font-family:'Space Mono',monospace;font-size:9px;
                  letter-spacing:0.18em;text-transform:uppercase;color:#6e8a7b;">by VisualizePro</div>
    </div>
  </div>
  <div style="background:#0c2019;border:1px solid rgba(122,215,159,0.20);
              border-radius:14px;padding:1.8rem;margin-bottom:1.5rem;text-align:center;">
    <div style="font-size:36px;margin-bottom:0.75rem;">✓</div>
    <div style="font-family:'Bricolage Grotesque',sans-serif;font-weight:800;
                font-size:22px;color:#ecf4ee;margin-bottom:0.5rem;">Payment received!</div>
    <p style="font-family:'DM Sans',sans-serif;font-size:14px;color:#9fb6a8;margin:0;">
      Your credits are being added. It may take a moment to activate —
      click below to continue and your balance will reflect shortly.
    </p>
  </div>
</div>""", unsafe_allow_html=True)

        if st.button("Continue to app →", type="primary", use_container_width=True, key="payment_continue_btn"):
            # Refresh the profile from Supabase to pick up the webhook-updated tier
            uid = st.session_state.get("auth_user_id")
            email = st.session_state.get("auth_email", "")
            if uid:
                st.session_state["user_profile"] = get_or_create_profile(uid, email)
            st.rerun()


# ============= TRACKER PAGE =============

def show_tracker():
    """Dedicated Applications Tracker page."""

    # Handle status changes submitted via the HTML select + URL params
    _p = st.query_params
    _rid = _p.get('rid', '')
    _sc  = _p.get('sc', '')
    if _rid and _sc in ['Applied', 'Interview', 'Offer', 'Declined', 'Rejected', 'Draft']:
        update_application_status(_rid, _sc)
        st.query_params.clear()
        st.rerun()

    # Hide Streamlit header
    st.markdown("""
    <style>
    [data-testid="stHeader"]{display:none!important;}
    #MainMenu{visibility:hidden!important;}
    footer{visibility:hidden!important;}
    </style>
    """, unsafe_allow_html=True)

    _tracker_uid = st.session_state.get('auth_user_id')
    tracker_data = load_tracker_data(user_id=_tracker_uid)

    # ── Sidebar ──────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("""
        <div style="padding:1.5rem 1rem 0;">
          <div style="display:flex;align-items:center;gap:11px;margin-bottom:1.6rem;">
            <div style="width:34px;height:34px;border-radius:9px;
                        background:linear-gradient(150deg,#7ad79f,#4fae7a);
                        display:grid;place-items:center;
                        font-family:'Bricolage Grotesque',system-ui,sans-serif;
                        font-weight:800;font-size:16px;color:#06140f;
                        box-shadow:0 4px 12px rgba(122,215,159,0.28);flex-shrink:0;">R</div>
            <div>
              <div style="font-family:'Bricolage Grotesque',system-ui,sans-serif;font-weight:700;
                          font-size:17px;letter-spacing:-0.02em;color:#ecf4ee;line-height:1.1;">ResumeSync</div>
              <div style="font-family:'Space Mono',monospace;font-size:9px;letter-spacing:0.18em;
                          text-transform:uppercase;color:#6e8a7b;margin-top:3px;">by VisualizePro</div>
            </div>
          </div>

          <div style="border-top:1px solid rgba(159,182,168,0.12);margin-bottom:1.2rem;"></div>

          <p style="font-family:'Space Mono',monospace;font-size:9px;letter-spacing:0.16em;
                    text-transform:uppercase;color:#7ad79f;margin:0 0 0.8rem;">Workspace</p>
        </div>
        """, unsafe_allow_html=True)

        if st.button("✦ New Analysis", key="tracker_nav_new", use_container_width=True):
            st.session_state['page'] = 'app'
            st.rerun()

        # Applications — active/current page
        st.markdown("""
        <div style="background:rgba(122,215,159,0.10);border:1px solid rgba(122,215,159,0.22);
                    border-radius:6px;padding:10px 14px;margin:4px 0;
                    font-family:'DM Sans',sans-serif;font-size:0.875rem;font-weight:600;color:#7ad79f;">
          📋 Applications
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div style='margin-top:1.5rem;border-top:1px solid rgba(159,182,168,0.12);padding-top:1rem;'>
          <a href='https://madelivniekerk.github.io/Resumesync/' target='_self'
             style='display:block;text-align:center;font-family:DM Sans,sans-serif;font-size:13px;
                    color:#6e8a7b;text-decoration:none;padding:6px 0;'>
            ← Back to home
          </a>
        </div>""", unsafe_allow_html=True)

    # ── Stats computation ─────────────────────────────────────────────────────
    total = len(tracker_data)
    applied_count = sum(1 for r in tracker_data if r.get('status') == 'Applied')
    interview_count = sum(1 for r in tracker_data if r.get('status') == 'Interview')
    pct_vals = []
    for r in tracker_data:
        raw = str(r.get('match_pct', '') or '').replace('%', '').strip()
        if raw.isdigit():
            pct_vals.append(int(raw))
    avg_match = (sum(pct_vals) // len(pct_vals)) if pct_vals else 0

    # ── Page header ───────────────────────────────────────────────────────────
    hdr_l, hdr_r = st.columns([3, 1])
    with hdr_l:
        st.markdown("""
        <div style="padding:28px 0 8px;">
          <h1 style="font-family:'Bricolage Grotesque',system-ui,sans-serif;font-weight:700;
                     font-size:clamp(26px,3vw,38px);letter-spacing:-0.025em;color:#ecf4ee;margin:0 0 8px;">
            Applications
          </h1>
          <p style="font-family:'DM Sans',sans-serif;font-size:14px;color:#9fb6a8;margin:0;">
            Every role you've analysed — resume, cover letter, score and status in one place.
          </p>
        </div>
        """, unsafe_allow_html=True)
    with hdr_r:
        st.markdown("<div style='padding-top:32px;'>", unsafe_allow_html=True)
        if st.button("New analysis →", key="tracker_hdr_new", type="primary", use_container_width=True):
            st.session_state['page'] = 'app'
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Stats row ─────────────────────────────────────────────────────────────
    s1, s2, s3, s4 = st.columns(4)
    stat_cards = [
        (s1, str(total), "Tracked", "#7ad79f"),
        (s2, str(applied_count), "Applied", "#6fb1e0"),
        (s3, str(interview_count), "Interviews", "#e0a14a"),
        (s4, (str(avg_match) + "%" if pct_vals else "—"), "Avg Match", "#7ad79f"),
    ]
    for col, val, label, color in stat_cards:
        with col:
            st.markdown(
                '<div style="background:#0c2019;border:1px solid rgba(159,182,168,0.12);'
                'border-radius:14px;padding:20px 18px;text-align:center;margin-bottom:16px;">'
                '<div style="font-family:\'Bricolage Grotesque\',system-ui,sans-serif;font-weight:800;'
                'font-size:2rem;color:' + color + ';line-height:1;">' + val + '</div>'
                '<div style="font-family:\'Space Mono\',monospace;font-size:9px;letter-spacing:0.16em;'
                'text-transform:uppercase;color:#6e8a7b;margin-top:6px;">' + label + '</div>'
                '</div>',
                unsafe_allow_html=True
            )

    # ── Empty state ───────────────────────────────────────────────────────────
    if total == 0:
        st.markdown("""
        <div style="background:#0c2019;border:1px solid rgba(159,182,168,0.12);border-radius:16px;
                    padding:48px 24px;text-align:center;margin-top:8px;">
          <div style="font-size:2.4rem;margin-bottom:14px;">📋</div>
          <div style="font-family:'Bricolage Grotesque',system-ui,sans-serif;font-weight:700;
                      font-size:20px;color:#ecf4ee;margin-bottom:8px;">No applications yet</div>
          <div style="font-family:'DM Sans',sans-serif;font-size:14px;color:#9fb6a8;margin-bottom:24px;">
            Run your first resume analysis to start tracking your job applications.
          </div>
        </div>
        """, unsafe_allow_html=True)
        _, mid_col, _ = st.columns([1, 1, 1])
        with mid_col:
            if st.button("Run your first analysis →", key="tracker_empty_cta", type="primary", use_container_width=True):
                st.session_state.page = 'app'
                st.rerun()
    else:
        STATUS_OPTIONS = ['Applied', 'Interview', 'Offer', 'Declined', 'Rejected', 'Draft']
        STATUS_CONFIG = {
            'Applied':   {'color': '#6fb1e0', 'bg': 'rgba(111,177,224,0.12)', 'border': 'rgba(111,177,224,0.30)'},
            'Interview': {'color': '#e0a14a', 'bg': 'rgba(224,161,74,0.14)',  'border': 'rgba(224,161,74,0.34)'},
            'Offer':     {'color': '#7ad79f', 'bg': 'rgba(122,215,159,0.12)', 'border': 'rgba(122,215,159,0.30)'},
            'Declined':  {'color': '#e07a5f', 'bg': 'rgba(224,122,95,0.12)', 'border': 'rgba(224,122,95,0.30)'},
            'Rejected':  {'color': '#ef4444', 'bg': 'rgba(239,68,68,0.10)',  'border': 'rgba(239,68,68,0.30)'},
            'Draft':     {'color': '#6e8a7b', 'bg': 'rgba(110,138,123,0.12)','border': 'rgba(110,138,123,0.28)'},
        }
        logo_colors = ['#1d3a31', '#1a2f3e', '#2d2518', '#2a1f35', '#1a2e2a', '#2e1f1f']

        # ── Compact row styling ────────────────────────────────────────────────
        st.markdown("""<style>
        [data-testid="stHorizontalBlock"] .element-container {
            margin-bottom: 0 !important;
            padding-bottom: 0 !important;
        }
        [data-testid="stHorizontalBlock"] [data-testid="stVerticalBlockBorderWrapper"],
        [data-testid="stHorizontalBlock"] [data-testid="stVerticalBlock"] {
            gap: 0 !important;
        }
        /* Remove padding around component iframes used for HTML selects */
        [data-testid="stHorizontalBlock"] iframe {
            display: block !important;
        }
        </style>""", unsafe_allow_html=True)

        # ── Table card header ─────────────────────────────────────────────────
        st.markdown("""
        <div style="background:#0c2019;border:1px solid rgba(159,182,168,0.12);
                    border-radius:14px 14px 0 0;padding:16px 22px 14px;margin-top:4px;
                    border-bottom:1px solid rgba(159,182,168,0.09);">
          <span style="font-family:'Bricolage Grotesque',sans-serif;font-weight:700;
                       font-size:16px;color:#ecf4ee;">All applications</span>
        </div>
        """, unsafe_allow_html=True)

        # ── Column headers ────────────────────────────────────────────────────
        hc0, hc1, hc2, hc3 = st.columns([3.2, 0.9, 1.3, 2.4])
        hs = ("font-family:'Space Mono',monospace;font-size:9.5px;letter-spacing:0.14em;"
              "text-transform:uppercase;color:#6e8a7b;padding:10px 0 6px;")
        with hc0: st.markdown(f'<div style="{hs}">Role</div>', unsafe_allow_html=True)
        with hc1: st.markdown(f'<div style="{hs}text-align:center;">Match</div>', unsafe_allow_html=True)
        with hc2: st.markdown(f'<div style="{hs}">Applied</div>', unsafe_allow_html=True)
        with hc3: st.markdown(f'<div style="{hs}">Status</div>', unsafe_allow_html=True)

        # ── Rows ──────────────────────────────────────────────────────────────
        for idx, row in enumerate(tracker_data):
            record_id = row.get('id')
            company   = row.get('company', '') or ''
            job_title = row.get('job_title', '') or ''
            match_pct = str(row.get('match_pct', '') or '')
            status    = row.get('status', 'Draft') or 'Draft'
            date_val  = str(row.get('date', '') or '')
            initial   = company[0].upper() if company else '?'
            logo_bg   = logo_colors[idx % len(logo_colors)]
            cfg       = STATUS_CONFIG.get(status, STATUS_CONFIG['Draft'])

            st.markdown(
                '<div style="height:1px;background:rgba(159,182,168,0.08);"></div>',
                unsafe_allow_html=True
            )

            col_role, col_match, col_date, col_status = st.columns([3.2, 0.9, 1.3, 2.4])

            with col_role:
                st.markdown(
                    f'<div style="display:flex;align-items:center;gap:12px;padding:5px 0 4px;">'
                    f'<div style="width:36px;height:36px;border-radius:9px;flex-shrink:0;'
                    f'background:{logo_bg};border:1px solid rgba(255,255,255,0.07);'
                    f'display:grid;place-items:center;font-family:\'Bricolage Grotesque\',sans-serif;'
                    f'font-weight:800;font-size:14px;color:#ecf4ee;">{initial}</div>'
                    f'<div style="flex:1;min-width:0;">'
                    f'<div style="font-family:\'DM Sans\',sans-serif;font-weight:600;font-size:13.5px;'
                    f'color:#ecf4ee;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{job_title}</div>'
                    f'<div style="font-family:\'DM Sans\',sans-serif;font-size:11.5px;color:#6e8a7b;margin-top:2px;">{company}</div>'
                    f'</div></div>',
                    unsafe_allow_html=True
                )

            with col_match:
                st.markdown(
                    f'<div style="display:flex;align-items:center;justify-content:center;min-height:28px;">'
                    f'<span style="font-family:\'Bricolage Grotesque\',sans-serif;font-weight:800;'
                    f'font-size:18px;color:#ecf4ee;">{match_pct}</span>'
                    f'</div>',
                    unsafe_allow_html=True
                )

            with col_status:
                # Native HTML <select> — full color control, no Base Web conflicts
                opts_html = ''.join(
                    f'<option value="{o}" {"selected" if o == status else ""}>{o}</option>'
                    for o in STATUS_OPTIONS
                )
                _components.html(f"""
                <html><head><style>
                html,body{{margin:0;padding:0;background:transparent;overflow:hidden;}}
                select{{
                    font-family:'Bricolage Grotesque',system-ui,sans-serif;
                    font-weight:800;font-size:12px;
                    color:{cfg['color']};
                    background:{cfg['bg']};
                    border:1px solid {cfg['border']};
                    border-radius:20px;
                    padding:4px 8px 4px 12px;
                    cursor:pointer;outline:none;
                    width:100%;max-width:150px;
                    margin-top:5px;
                }}
                </style></head><body>
                <select onchange="try{{
                    var u=new URL(window.parent.location.href);
                    u.searchParams.set('sc',this.value);
                    u.searchParams.set('rid','{record_id}');
                    window.parent.location.href=u.href;
                }}catch(e){{}}">
                    {opts_html}
                </select>
                </body></html>
                """, height=38)

            with col_date:
                st.markdown(
                    f'<div style="display:flex;align-items:center;min-height:28px;">'
                    f'<span style="font-family:\'Space Mono\',monospace;font-size:11px;'
                    f'color:#6e8a7b;white-space:nowrap;">{date_val}</span>'
                    f'</div>',
                    unsafe_allow_html=True
                )

        st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

        # Download tracker Excel from tracker page too
        st.download_button(
            label="📥 Download Tracker (Excel)",
            data=generate_tracker_excel(tracker_data),
            file_name="job_applications.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="tracker_page_download",
            use_container_width=False
        )


# ============= STREAMLIT UI =============

def main():
    # ── Handle ?login=1[&plan=X] query params ────────────────────────────────
    if st.query_params.get('login'):
        plan = st.query_params.get('plan', '')
        st.query_params.clear()
        st.session_state.show_login = True
        if plan in ('free', 'pack5', 'pack20', 'pack30', 'pack50'):
            st.session_state.pending_plan = plan
        st.rerun()

    # ── Handle ?payment_return=1 (returning from Pay Advanced) ──────────────
    if st.query_params.get('payment_return'):
        st.query_params.clear()
        st.session_state.show_payment_pending = True
        # If session was lost during redirect, send to login first (not landing)
        if 'auth_user_id' not in st.session_state:
            st.session_state.show_login = True
        st.rerun()

    # ── Auth gate ─────────────────────────────────────────────────────────────
    if 'auth_user_id' not in st.session_state:
        if st.session_state.get('show_login'):
            show_login_page()
        else:
            show_landing()
        return

    # ── Post-login payment screens ────────────────────────────────────────────
    if st.session_state.pop('show_payment_pending', None):
        show_payment_pending()
        return

    pending_plan = st.session_state.get('pending_plan', '')
    if pending_plan in PACK_INFO:
        st.session_state.pop('pending_plan', None)
        show_payment_redirect(pending_plan)
        return
    st.session_state.pop('pending_plan', None)

    auth_user_id = st.session_state['auth_user_id']
    auth_email   = st.session_state.get('auth_email', '')
    profile      = st.session_state.get('user_profile') or get_or_create_profile(auth_user_id, auth_email)
    st.session_state['user_profile'] = profile

    # ── Route to tracker or app workspace ─────────────────────────────────────
    page = st.session_state.get('page', 'app')
    if page == 'tracker':
        show_tracker()
        return

    # ── Sidebar ─────────────────────────────────────────────────────────────
    with st.sidebar:
        tracker_data = load_tracker_data(user_id=auth_user_id)
        count = len(tracker_data)
        st.markdown(f"""
        <div style="padding: 1.5rem 1rem 0;">

          <!-- Brand mark -->
          <div style="display:flex;align-items:center;gap:11px;margin-bottom:1.4rem;">
            <div style="width:34px;height:34px;border-radius:9px;background:linear-gradient(150deg,#7ad79f,#4fae7a);display:grid;place-items:center;font-family:'Bricolage Grotesque',system-ui,sans-serif;font-weight:800;font-size:16px;color:#06140f;box-shadow:0 4px 12px rgba(122,215,159,0.28);flex-shrink:0;">R</div>
            <div>
              <div style="font-family:'Bricolage Grotesque',system-ui,sans-serif;font-weight:700;font-size:17px;letter-spacing:-0.02em;color:#ecf4ee;line-height:1.1;">ResumeSync</div>
              <div style="font-family:'Space Mono',monospace;font-size:9px;letter-spacing:0.18em;text-transform:uppercase;color:#6e8a7b;margin-top:3px;">by VisualizePro</div>
            </div>
          </div>

          <!-- Divider -->
          <div style="border-top:1px solid rgba(159,182,168,0.12); margin-bottom:1.2rem;"></div>

          <!-- User / credits block -->
          <div style="background:#0c2019;border:1px solid rgba(159,182,168,0.10);
                      border-radius:10px;padding:12px 14px;margin-bottom:1.2rem;">
            <div style="font-family:'DM Sans',sans-serif;font-size:12px;
                        color:#9fb6a8;white-space:nowrap;overflow:hidden;
                        text-overflow:ellipsis;margin-bottom:6px;">{auth_email}</div>
            <div style="display:flex;align-items:center;gap:8px;">
              <span style="font-family:'Bricolage Grotesque',sans-serif;font-weight:800;
                           font-size:22px;color:#7ad79f;line-height:1;">
                {int(profile.get('credits', 0))}
              </span>
              <span style="font-family:'DM Sans',sans-serif;font-size:11px;color:#6e8a7b;">
                {"analysis" if int(profile.get('credits',0)) == 1 else "analyses"} remaining
              </span>
            </div>
          </div>

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

        _has_analysis = 'analysis_result' in st.session_state
        if not st.session_state.get('_confirm_leave_tracker'):
            if st.button("📋 View Applications", key="sidebar_view_tracker", use_container_width=True):
                if _has_analysis:
                    st.session_state['_confirm_leave_tracker'] = True
                    st.rerun()
                else:
                    st.session_state.page = 'tracker'
                    st.rerun()
        else:
            st.markdown(
                '<div style="background:rgba(224,122,95,0.10);border:1px solid rgba(224,122,95,0.3);'
                'border-radius:8px;padding:0.75rem 0.9rem;margin-bottom:0.5rem;">'
                '<p style="color:#e07a5f;font-family:\'DM Sans\',sans-serif;font-size:0.82rem;'
                'font-weight:600;margin:0 0 0.3rem;">⚠️ Leave this analysis?</p>'
                '<p style="color:#9fb6a8;font-family:\'DM Sans\',sans-serif;font-size:0.78rem;margin:0;">'
                'Download your report first — it won\'t be here when you come back.</p>'
                '</div>',
                unsafe_allow_html=True
            )
            _res = st.session_state.get('analysis_result', {})
            _rc  = st.session_state.get('resume_text', '')
            _jc  = st.session_state.get('job_content', '')
            _jurl = st.session_state.get('job_url', '')
            _rfn  = st.session_state.get('resume_filename', 'resume')
            if _res:
                st.download_button(
                    label="💾 Download Report",
                    data=create_analysis_docx(_res.get('analysis', ''), _jurl, _rfn, _jc),
                    file_name=f"ResumeAnalysis_{_rfn.rsplit('.', 1)[0][:20]}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="sidebar_dl_before_leave",
                    use_container_width=True
                )
            leave_col, stay_col = st.columns(2)
            with leave_col:
                if st.button("Go anyway", key="confirm_leave_tracker", use_container_width=True):
                    st.session_state.pop('_confirm_leave_tracker', None)
                    st.session_state.page = 'tracker'
                    st.rerun()
            with stay_col:
                if st.button("Stay", key="cancel_leave_tracker", use_container_width=True):
                    st.session_state.pop('_confirm_leave_tracker', None)
                    st.rerun()

        st.markdown("""
        <a href='https://madelivniekerk.github.io/Resumesync/' target='_self'
           style='display:block;text-align:center;font-family:DM Sans,sans-serif;font-size:13px;
                  color:#6e8a7b;text-decoration:none;padding:6px 0;margin-bottom:6px;'>
          ← Back to home
        </a>""", unsafe_allow_html=True)

        if tracker_data:
            st.download_button(
                label="📥 Download Tracker",
                data=generate_tracker_excel(tracker_data),
                file_name="job_applications.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="sidebar_download_tracker",
                use_container_width=True
            )

        st.markdown("<div style='height:1rem;'></div>", unsafe_allow_html=True)
        if st.button("Sign out", key="sign_out_btn", use_container_width=True):
            for _k in ['auth_user_id','auth_email','auth_token','auth_refresh','user_profile','otp_sent_email']:
                st.session_state.pop(_k, None)
            st.rerun()

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
            # Cache bytes immediately so they survive all subsequent reruns
            st.session_state['_resume_bytes'] = resume_file.getvalue()
            st.session_state['_resume_name'] = resume_file.name
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

    has_resume = bool(st.session_state.get('_resume_bytes'))
    has_job = bool(job_url or manual_job_text)

    col_hint, col_btn = st.columns([2, 1])
    with col_hint:
        ready = has_resume and has_job
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
        if st.button(
            "Analyze compatibility →",
            type="primary",
            use_container_width=True,
        ):
            if not has_resume:
                st.error("❌ Please upload a resume file first.")
            elif not has_job:
                st.error("❌ Please add a job posting (URL or paste).")
            else:
                st.session_state['_do_analysis'] = True
                st.session_state['_job_url_snap'] = job_url
                st.session_state['_manual_job_snap'] = manual_job_text

    if st.session_state.get('_do_analysis'):
        st.session_state['_do_analysis'] = False
        resume_bytes = st.session_state.get('_resume_bytes')
        resume_name  = st.session_state.get('_resume_name', 'resume')
        job_url      = st.session_state.get('_job_url_snap') or job_url
        manual_job_text = st.session_state.get('_manual_job_snap') or manual_job_text
        if not resume_bytes:
            st.error("❌ Please upload a resume file first.")
            return
        if not (job_url or manual_job_text):
            st.error("❌ Please add a job posting.")
            return

        # ── Usage limit check ─────────────────────────────────────────────────
        _allowed, _reason = can_run_analysis(profile)
        if not _allowed:
            st.error(f"🔒 {_reason}")
            st.markdown("**Top up with a credit pack — one-off, no subscription, never expires.**")
            _pc1, _pc2 = st.columns(2)
            with _pc1:
                if st.button("5 analyses — A$9", key="buy_pack5", use_container_width=True):
                    st.session_state.pending_plan = 'pack5'
                    st.rerun()
                if st.button("30 analyses — A$20", key="buy_pack30", use_container_width=True):
                    st.session_state.pending_plan = 'pack30'
                    st.rerun()
            with _pc2:
                if st.button("20 analyses — A$15", key="buy_pack20", use_container_width=True):
                    st.session_state.pending_plan = 'pack20'
                    st.rerun()
                if st.button("50 analyses — A$30", key="buy_pack50", use_container_width=True):
                    st.session_state.pending_plan = 'pack50'
                    st.rerun()
            return

        with st.spinner("🔍 Analyzing your resume against the job posting..."):

            with st.status("Extracting text from resume...", expanded=True) as status:
                resume_file_obj = io.BytesIO(resume_bytes)
                resume_file_obj.name = resume_name
                resume_text = extract_resume_text(resume_file_obj)
                if "Error" in resume_text:
                    st.error(f"❌ {resume_text}")
                    return
                if len(resume_text.strip()) < 100:
                    st.error("❌ Could not read enough text from your resume. If it's a PDF, it may be image-based (scanned) — try saving as DOCX instead. If it's a DOCX, re-save it from Word and re-upload.")
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

        # Clear previous comparison results before storing new ones
        for _k in ['cover_letter', 'proposed_updates', 'updated_resume_bytes',
                   'updated_resume_name', 'updated_match_pct', 'tracker_saved',
                   '_confirm_leave_tracker', 'upd_guidance', '_upd_guidance_saved']:
            st.session_state.pop(_k, None)

        st.session_state['analysis_result'] = result
        st.session_state['resume_text'] = resume_text
        st.session_state['job_content'] = job_content
        st.session_state['job_url'] = job_url
        st.session_state['resume_filename'] = resume_name
        st.session_state['resume_file_bytes'] = resume_bytes
        st.session_state['resume_is_docx'] = resume_name.lower().endswith(('.docx', '.doc'))
        st.session_state['tracker_saved'] = False

        # Only charge one analysis per unique resume+job combination
        _analysis_key = hashlib.md5((resume_text + job_content).encode()).hexdigest()
        if st.session_state.get('_last_analysis_key') != _analysis_key:
            decrement_credits(auth_user_id)
            st.session_state['_last_analysis_key'] = _analysis_key

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
                            'cover_letter', 'tracker_saved', 'proposed_updates',
                            'updated_resume_bytes', 'updated_resume_name', 'updated_match_pct',
                            'upd_guidance', '_upd_guidance_saved']:
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

            st.markdown(
                '<div style="background:rgba(111,177,224,0.08);padding:1rem 1.5rem;border-radius:12px;'
                'border-left:4px solid #6fb1e0;margin-bottom:1rem;">'
                '<p style="color:#6fb1e0;font-family:\'Space Mono\',monospace;font-size:9px;'
                'letter-spacing:0.14em;text-transform:uppercase;margin:0 0 0.5rem;">Tip — quantified impact</p>'
                '<p style="color:#9fb6a8;font-size:0.88rem;margin:0 0 0.5rem;font-family:\'DM Sans\',sans-serif;line-height:1.6;">'
                'AI systems heavily favour measurable results. If you know any of these, add them to the guidance box below and the AI will work them in:</p>'
                '<p style="color:#ecf4ee;font-size:0.85rem;margin:0;font-family:\'DM Sans\',sans-serif;line-height:1.8;">'
                '📊 <b>%</b> &nbsp;·&nbsp; ⏱ <b>time saved</b> &nbsp;·&nbsp; 💰 <b>revenue / cost reduction</b> &nbsp;·&nbsp; ⚡ <b>efficiency gain</b> &nbsp;·&nbsp; 📈 <b>scale</b> (users, datasets, systems)<br>'
                '<span style="color:#6e8a7b;font-size:0.82rem;">'
                'e.g. "reduced processing time by 35%" &nbsp;·&nbsp; "increased accuracy by 20%" &nbsp;·&nbsp; "supported 5M+ records daily"'
                '</span></p>'
                '</div>',
                unsafe_allow_html=True
            )

            upd_guidance = st.text_area(
                "Additional guidance *(optional)*",
                placeholder="Add any metrics you know, e.g. 'reduced costs by 30%', 'managed a team of 8', 'processed 2M records daily'. Also: emphasise leadership, lead with Python skills, make it more senior in tone.",
                height=150,
                key="upd_guidance"
            )

            st.markdown('<div id="resume-updater-anchor"></div>', unsafe_allow_html=True)
            col_upd = st.columns([1, 2, 1])[1]
            with col_upd:
                update_btn = st.button("✨ Propose Resume Changes", type="primary", key="update_resume_btn", use_container_width=True)

            if update_btn:
                st.session_state['_upd_guidance_saved'] = upd_guidance
                with st.spinner("Generating proposed changes..."):
                    with st.status("Analysing what can be improved...", expanded=True) as upd_status:
                        upd_result = generate_resume_updates(resume_text, result['analysis'], client, user_guidance=upd_guidance)
                        if not upd_result['success']:
                            st.error(f"❌ Could not generate changes: {upd_result['error']}")
                        else:
                            st.write(f"✅ {len(upd_result['updates'])} proposed changes ready for review")
                            upd_status.update(label="Proposals ready — please review below", state="complete")
                            st.session_state['proposed_updates'] = upd_result['updates']
                            st.session_state.pop('updated_resume_bytes', None)
                # Scroll back to the Resume Updater section, not the bottom of the page
                _components.html(
                    '<script>'
                    'window.parent.document.getElementById("resume-updater-anchor")'
                    '.scrollIntoView({behavior:"smooth",block:"start"});'
                    '</script>',
                    height=0,
                )

            # ── Review step ────────────────────────────────────────────────
            if 'proposed_updates' in st.session_state and st.session_state['proposed_updates']:
                proposed = st.session_state['proposed_updates']

                _saved_guidance = st.session_state.get('_upd_guidance_saved', '').strip()
                if _saved_guidance:
                    st.markdown(
                        f'<div style="background:rgba(111,177,224,0.08);border-left:3px solid #6fb1e0;'
                        f'padding:0.5rem 0.9rem;border-radius:6px;margin-bottom:0.8rem;font-family:\'DM Sans\',sans-serif;">'
                        f'<span style="color:#6fb1e0;font-size:0.75rem;font-family:\'Space Mono\',monospace;'
                        f'text-transform:uppercase;letter-spacing:0.1em;">✓ Guidance applied: </span>'
                        f'<span style="color:#9fb6a8;font-size:0.82rem;">{_saved_guidance}</span>'
                        f'</div>',
                        unsafe_allow_html=True
                    )

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
                                f'<p style="color:#7ad79f;font-size:0.75rem;font-family:\'Space Mono\',monospace;'
                                f'letter-spacing:0.1em;text-transform:uppercase;margin:0 0 0.3rem;">After — edit if needed</p>',
                                unsafe_allow_html=True
                            )
                            edited_replace = st.text_area(
                                f"After {i+1}",
                                value=replace,
                                height=150,
                                key=f"edit_replace_{i}",
                                label_visibility="collapsed"
                            )
                        st.markdown("<div style='margin-bottom:0.8rem'></div>", unsafe_allow_html=True)
                        if checked:
                            selected.append({**change, 'replace': edited_replace})

                n_selected = len(selected)
                col_apply = st.columns([1, 2, 1])[1]
                with col_apply:
                    apply_btn = st.button(
                        "✅ Apply " + str(n_selected) + " Selected Change" + ("s" if n_selected != 1 else ""),
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
                    new_filename = f"{_fn_person}_Updated_{_fn_role}_{_fn_date}.docx"
                    st.session_state['updated_resume_bytes'] = updated_bytes
                    st.session_state['updated_resume_name'] = new_filename
                    st.success(f"✅ {applied} change(s) applied to your resume.")
                    # Re-score the updated resume
                    with st.spinner("Calculating updated compatibility score..."):
                        _upd_doc = Document(io.BytesIO(updated_bytes))
                        _upd_paras = list(_upd_doc.paragraphs)
                        for _t in _upd_doc.tables:
                            for _r in _t.rows:
                                for _c in _r.cells:
                                    _upd_paras.extend(_c.paragraphs)
                        updated_text = "\n".join(p.text for p in _upd_paras if p.text.strip())
                        new_score = re_score_resume(updated_text, job_content, client)
                        if new_score is not None:
                            st.session_state['updated_match_pct'] = f"{new_score}%"

            if 'updated_resume_bytes' in st.session_state:
                new_filename = st.session_state.get('updated_resume_name', f"{_fn_person}_Updated_{_fn_role}_{_fn_date}.docx")
                # Show score lift if re-score is available
                updated_pct = st.session_state.get('updated_match_pct')
                if updated_pct:
                    orig_pct = fields.get('match_pct', '')
                    orig_num = int(re.search(r'\d+', orig_pct).group()) if re.search(r'\d+', orig_pct) else None
                    new_num  = int(re.search(r'\d+', updated_pct).group()) if re.search(r'\d+', updated_pct) else None
                    if orig_num is not None and new_num is not None:
                        delta = new_num - orig_num
                        delta_str = f"+{delta}" if delta >= 0 else str(delta)
                        delta_color = "#7ad79f" if delta >= 0 else "#ef4444"
                        st.markdown(
                            f'<div style="text-align:center;margin:8px 0 4px;">'
                            f'<span style="font-family:\'Space Mono\',monospace;font-size:12px;color:#9fb6a8;">Compatibility: </span>'
                            f'<span style="font-family:\'Bricolage Grotesque\',sans-serif;font-weight:800;font-size:17px;color:#9fb6a8;text-decoration:line-through;">{orig_pct}</span>'
                            f'<span style="font-family:\'Bricolage Grotesque\',sans-serif;font-size:17px;color:#6e8a7b;margin:0 6px;">→</span>'
                            f'<span style="font-family:\'Bricolage Grotesque\',sans-serif;font-weight:800;font-size:17px;color:#ecf4ee;">{updated_pct}</span>'
                            f'<span style="font-family:\'Space Mono\',monospace;font-size:11px;color:{delta_color};'
                            f'background:{"rgba(122,215,159,0.12)" if delta >= 0 else "rgba(239,68,68,0.10)"};'
                            f'padding:2px 7px;border-radius:5px;margin-left:8px;">{delta_str} pts</span>'
                            f'</div>',
                            unsafe_allow_html=True
                        )
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
                '<p style="font-family:\'DM Sans\',sans-serif;font-size:0.8rem;color:#9fb6a8;'
                'margin:0 0 0.3rem;">Edit directly below — your changes are reflected in the download.</p>',
                unsafe_allow_html=True
            )
            cl_edited = st.text_area(
                "Cover letter",
                value=st.session_state['cover_letter'],
                height=420,
                key="cl_edit_area",
                label_visibility="collapsed"
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
                    data=create_cover_letter_docx(cl_edited),
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
                            prior_letter=cl_edited,
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
                st.success("✅ Saved to tracker!")
                if st.button("📋 View Applications →", key="goto_tracker_after_save", use_container_width=True):
                    st.session_state.page = 'tracker'
                    st.rerun()
            else:
                if st.button("💾 Save to Job Tracker", key="save_tracker", use_container_width=True):
                    try:
                        # Use edited cover letter text if available (from cl_edit_area widget)
                        cover_letter_text = (
                            st.session_state.get('cl_edit_area')
                            or st.session_state.get('cover_letter', '')
                        )
                        save_to_tracker(
                            job_title=job_title_input,
                            company=company_input,
                            location=location_input,
                            resume_filename=resume_filename,
                            match_pct=st.session_state.get('updated_match_pct') or fields['match_pct'],
                            job_url=job_url,
                            cover_letter=cover_letter_text,
                            cover_letter_path='',
                            notes=extract_recommendations_summary(result['analysis']),
                            updated_resume_file=st.session_state.get('updated_resume_name', ''),
                            user_id=st.session_state.get('auth_user_id'),
                        )
                        load_tracker_data.clear()
                        st.session_state['tracker_saved'] = True
                        st.rerun()
                    except Exception as _save_err:
                        st.error(f"❌ Could not save to tracker: {_save_err}")

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
                data=create_analysis_docx(result['analysis'], job_url, resume_filename, job_content),
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


if _IMPORT_ERROR:
    st.error(f"**Startup import error:** {_IMPORT_ERROR}")
    st.code(_tb.format_exc(), language="python")
    st.info("Check Streamlit Cloud → App Settings → Secrets for ANTHROPIC_API_KEY, SUPABASE_URL, SUPABASE_KEY.")
else:
    try:
        main()
    except Exception as _e:
        st.error(f"**App error:** {_e}")
        st.code(_tb.format_exc(), language="python")
        st.info("Check Streamlit Cloud → App Settings → Secrets for ANTHROPIC_API_KEY, SUPABASE_URL, SUPABASE_KEY.")
