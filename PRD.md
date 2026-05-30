# ResumeSync — Product Requirements Document

## Overview

ResumeSync is an AI-powered career tool that helps job seekers understand how well their resume matches a specific job posting, identify gaps, and take action — all in one workflow.

## Goals

- Reduce wasted applications by showing fit before applying
- Help users pass ATS screening with targeted keyword recommendations
- Generate ready-to-send cover letters tailored to each role
- Track applications in one place

## Target Users

- Active job seekers across all industries
- Career changers wanting to assess transferable skills
- Recruiters or career coaches reviewing candidates

## User Stories

| As a... | I want to... | So that... |
|---|---|---|
| Job seeker | Upload my resume and paste a job URL | I can see my compatibility score instantly |
| Job seeker | See which keywords are missing | I can add them before applying |
| Job seeker | Generate a tailored cover letter | I don't have to write one from scratch each time |
| Job seeker | Download an improved resume | I have an updated DOCX ready to submit |
| Job seeker | Save the application to a tracker | I can keep track of every role I've applied for |
| Career coach | Paste a manual job description | I can review client fit without a live URL |

## Features

### Core (Complete)
- **Resume upload** — PDF, DOCX, TXT support
- **Job input** — URL scraping (BeautifulSoup) or manual paste
- **AI compatibility analysis** — match score, keyword table, hard/soft skills gap, section review, ATS tips, cover letter talking points
- **Cover letter generator** — tone, length, incorporate-recs toggle, additional guidance, regenerate with changes
- **AI resume updater** — proposes find→replace edits; user reviews and selects; outputs updated DOCX
- **Job tracker** — saves job title, company, location, match %, URL, cover letter DOCX path to Excel
- **Sidebar tracker download** — accessible from landing page without running an analysis

### Planned
- User authentication (Supabase)
- Subscription tiers via Stripe (Free / Starter / Pro)
- Hosted deployment on Railway (`streamlit run resume_matcher_app.py --server.port $PORT --server.address 0.0.0.0`)

## Non-Goals

- Does not fabricate skills or experience not present in the original resume
- Does not store personal data server-side (local-first)
- Does not apply to jobs on the user's behalf

## Tech Stack

Python · Streamlit · Anthropic Claude API · BeautifulSoup4 · PyPDF2 · python-docx · openpyxl · Pillow · NumPy

## Design System

VisualizePro Forest + Sage: `#0e2a23` background, `#082019` sidebar, `#6dc18a` accent. Fonts: Fraunces (headings), Manrope (body), JetBrains Mono (data labels).
