# Resume Matcher Web App - User Guide

## What Is This?

A **professional web-based application** that analyzes your resume against job postings and provides:
- ✅ **Compatibility Score (0-100%)**
- ✅ **Strengths & Weaknesses**
- ✅ **Resume Improvement Recommendations**
- ✅ **Cover Letter Talking Points**
- ✅ **ATS Optimization Tips**
- ✨ **NEW: AI-Generated Cover Letters**

## Features

### 1. Upload Resume
- Supports **PDF, Word (.docx), and Text (.txt)** files
- Automatically extracts text from any format
- No copy-pasting needed!

### 2. Paste Job URL or Description
- **Option A:** Paste job URL for auto-scraping
- **Option B:** Manually paste job description text
- Works with Seek, LinkedIn, Indeed, and any job board
- Manual paste option for sites that block scraping

### 3. AI-Powered Analysis
- Uses **Claude Sonnet 4.5** (latest AI model)
- Provides detailed, actionable feedback
- Specific recommendations for THIS job

### 4. Cover Letter Generation ✨ NEW!
- Click one button to generate a personalized cover letter
- Tailored to the specific job and company
- 250-350 words, professional format
- Based on your resume and match analysis
- Ready to copy-paste or download

### 5. Download Results
- Save analysis as a text file
- Download cover letter separately
- Keep track of all applications
- Reference materials for interviews

---

## How to Run the App

### Method 1: Double-Click (Easiest)

1. Navigate to `C:\Users\madel\`
2. Double-click **`run_resume_matcher.bat`**
3. A browser window will open automatically
4. Start analyzing resumes!

### Method 2: Command Line

Open Command Prompt and run:
```bash
cd C:\Users\madel
streamlit run resume_matcher_app.py
```

---

## How to Use the App

### Step 1: Upload Your Resume
- Click "Browse files" button
- Select your resume (PDF, Word, or Text)
- You'll see a green checkmark when uploaded

### Step 2: Add Job Posting
**Option A: Paste URL (Auto-scrape)**
- Choose "Paste URL (auto-scrape)" radio button
- Go to Seek/LinkedIn/Indeed
- Copy the full job posting URL
  - Example: `https://www.seek.com.au/job/82345678`
- Paste it into the "Job posting URL" field

**Option B: Paste Description Manually**
- Choose "Paste job description manually" radio button
- Copy the full job description text from the website
- Paste into the text box
- Use this if URL scraping fails (some sites block it)

### Step 3: Click "Analyze Compatibility"
- The app will:
  1. Extract text from your resume
  2. Get the job posting (scrape or use manual paste)
  3. Send both to Claude AI for analysis
  4. Display comprehensive results

### Step 4: Review Results
You'll get:
- **Compatibility Score**: How well you match (0-100%)
- **Job Details**: Extracted title, company, location
- **Strengths**: What makes you a good fit
- **Gaps**: What you're missing or under-emphasizing
- **Recommendations**: Specific changes to make
- **Cover Letter Points**: Top 3 things to mention
- **ATS Tips**: Keywords to include

### Step 5: Generate Cover Letter ✨ NEW!
- Scroll down after viewing your analysis
- Click "🚀 Generate Cover Letter" button
- Wait 10-15 seconds for AI to write it
- Review the personalized cover letter
- Download it or copy-paste into your application

### Step 6: Download Reports (Optional)
- Click "Download Analysis Report" for the analysis
- Click "Download Cover Letter" for the letter
- Save both for reference when applying
- Click "🔄 New Analysis" to analyze a different job

---

## Cost Per Analysis

### Resume Analysis Only
Each analysis costs approximately **$0.025** (2.5 cents):
- Resume input: ~500 tokens × $3/million = $0.0015
- Job posting: ~500 tokens × $3/million = $0.0015
- Claude response: ~1500 tokens × $15/million = $0.0225
- **Total: ~$0.025** per analysis

### Analysis + Cover Letter
Combined cost approximately **$0.045-0.05** (4.5-5 cents):
- Resume analysis: ~$0.025
- Cover letter generation: ~$0.02
- **Total: ~$0.045** per complete session

### Your API Credits
**With $10 credit**:
- ~400 analyses OR
- ~220 analyses + cover letters

**With $20 credit**:
- ~800 analyses OR
- ~440 analyses + cover letters

---

## Example Workflow

### Scenario: Applying for BI Specialist Role

1. **Open the app**: Double-click `run_resume_matcher.bat`

2. **Upload resume**: Select `Madeli_van_Niekerk_Resume.docx`

3. **Paste job URL**:
   ```
   https://www.seek.com.au/job/82345678?type=standard
   ```

4. **Click "Analyze"**

5. **Get results**:
   ```
   COMPATIBILITY SCORE: 85%

   KEY STRENGTHS:
   - 10+ years in BI and visualization
   - Financial services experience (Pepper Money)
   - Strong stakeholder engagement
   - Tableau expertise transferable to Power BI

   GAPS & WEAKNESSES:
   - Power BI not emphasized as primary tool
   - Limited Python projects mentioned
   - No DAX skills highlighted

   RECOMMENDATIONS:
   1. Add "Certified in Power BI" to summary
   2. Create a "Technical Skills" section mentioning DAX
   3. Emphasize AI integration work (Claude dashboards)
   4. Highlight financial services governance experience

   COVER LETTER POINTS:
   - Lead with Pepper Money financial services work
   - Emphasize learning agility (Tableau → Power BI)
   - Showcase AI integration expertise

   ATS KEYWORDS TO ADD:
   Power BI, DAX, data modeling, financial services,
   stakeholder management, Python, AI integration
   ```

6. **Download** and save the analysis

7. **Update resume** based on recommendations

8. **Write cover letter** using the talking points

---

## Troubleshooting

### "API Key not configured"
- Check that `.env` file exists at `C:\Users\madel\.env`
- Verify it contains your API key
- Make sure no extra spaces or quotes

### "Could not fetch job posting"
**Possible reasons:**
- Website blocks automated scraping (common)
- URL is incorrect or expired
- Site requires login

**Solutions:**
1. Try a different job URL
2. Copy job description manually and we can add a "paste mode"
3. Some sites work better than others (Seek usually works well)

### "Credit balance too low"
- Go to https://console.anthropic.com/settings/billing
- Add $10-20 in credits
- Start analyzing!

### Page won't load
- Check if port 8501 is already in use
- Close other instances of Streamlit
- Restart the app

---

## Technical Details

### What Happens Behind the Scenes

1. **Resume Upload** → Text extraction using PyPDF2/python-docx
2. **Job URL** → Web scraping using BeautifulSoup + requests
3. **Both sent to Claude** → Intelligent analysis
4. **Results displayed** → Formatted in Streamlit interface

### Files Created

| File | Purpose |
|------|---------|
| `resume_matcher_app.py` | Main web application code |
| `run_resume_matcher.bat` | Easy launcher |
| `requirements.txt` | Package dependencies |
| `.env` | API key configuration |
| `RESUME_MATCHER_GUIDE.md` | This guide! |

### Technologies Used

- **Streamlit**: Web framework
- **Claude AI**: Resume analysis
- **BeautifulSoup**: Job scraping
- **PyPDF2**: PDF text extraction
- **python-docx**: Word document reading

---

## Next Steps

### For Personal Use
1. Add $10-20 in Anthropic credits
2. Analyze all your target jobs
3. Improve your resume based on feedback
4. Track which jobs you match best

### For Building a Business
1. **Test with 10+ users** to validate demand
2. **Add user accounts** for tracking history
3. **Implement payment** (Stripe integration)
4. **Deploy to cloud** (Railway, Render, Heroku)
5. **Market it** on LinkedIn, Reddit, job forums

### Potential Features to Add
- ✅ User login / account system
- ✅ Save analysis history
- ✅ Compare multiple jobs side-by-side
- ✅ Email analysis reports
- ✅ Batch processing (upload multiple jobs at once)
- ✅ Resume templates and suggestions
- ✅ LinkedIn profile optimization
- ✅ Interview question generator based on job

---

## Important Reminders

### Security
- **NEVER share your API key** publicly
- **Regenerate your API key** after sharing it earlier
- Keep `.env` file private (don't commit to GitHub)

### API Usage
- Each analysis costs ~$0.025
- Monitor usage in Anthropic console
- Set budget limits if needed

### Commercialization
- ✅ You own this code
- ✅ You can sell/monetize it
- ✅ No attribution required
- Consider trademarking your brand name

---

## Support

**Created by:** You (Madeli van Niekerk)
**Powered by:** Claude AI (Anthropic)
**Built with:** Claude Code assistance

**For questions:**
- Check this guide first
- Review error messages carefully
- Test with different job URLs
- Verify API credit balance

---

## Success Tips

1. **Test with jobs you're actually applying to**
2. **Compare analyses** to find your best matches
3. **Use recommendations** to update your resume
4. **Track results** - which jobs score highest?
5. **Iterate** - re-analyze after making changes

**Good luck with your job search!** 🚀

---

_Last updated: April 28, 2026_
