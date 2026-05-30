# ResumeSync

AI-powered resume matcher by VisualizePro. Upload your resume and paste any job posting URL or description — get a compatibility score, skills gap analysis, ATS keyword optimisation, a tailored cover letter, and an AI-improved resume in seconds.

## Tech Stack

| Layer | Technology |
|---|---|
| Frontend | Streamlit |
| AI | Anthropic Claude API (claude-sonnet-4-5-20250929) |
| Job scraping | BeautifulSoup4, Requests |
| Resume parsing | PyPDF2, python-docx |
| Tracker | openpyxl |
| Assets | Pillow, NumPy |

## Setup

1. Install dependencies
   ```
   pip install -r requirements.txt
   ```
2. Create a `.env` file in `C:\Users\madel\.env` with:
   ```
   ANTHROPIC_API_KEY=your_key_here
   ```
3. Run the app
   ```
   streamlit run resume_matcher_app.py
   ```
4. Open http://localhost:8501

## Project Structure

```
resumesync/
├── resume_matcher_app.py        Main Streamlit app
├── requirements.txt             Python dependencies
├── assets/                      Logo and static assets
├── cover_letters/               Auto-saved cover letter DOCXs
├── job_applications.xlsx        Job application tracker
├── Resume App/                  Design handoff files (hero section)
├── resume_analyzer.py           Standalone resume analysis utility
├── extract_resume.py            Resume text extraction utility
├── job_scraper_demo.py          Job scraping demo/utility
└── README.md / PRD.md           Project documentation
```

## Features

- Resume upload (PDF, DOCX, TXT)
- Job posting input via URL (auto-scraped) or manual paste
- AI compatibility score (0–100%) with keyword match table
- Hard skills and soft skills gap analysis
- Section-by-section resume review
- ATS keyword optimisation report
- Cover letter generator (tone, length, custom guidance, regenerate with changes)
- AI resume updater — proposes and applies improvements, user approves each change
- Job application tracker saved to Excel with cover letter file path
- Download analysis report, cover letter, and updated resume as DOCX
