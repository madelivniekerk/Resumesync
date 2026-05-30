# Changelog — ResumeSync

## [1.2] — 2026-05-18

### Added
- Cover letters auto-saved to `cover_letters/` folder when saving to tracker
- Cover letter file path stored in Excel instead of full text
- Sidebar "Download Tracker" button visible on landing page without requiring an analysis
- `@st.cache_data(ttl=30)` on `load_tracker_data()` to reduce repeated disk reads

### Changed
- Google Fonts loading changed from blocking `@import` to non-blocking `<link rel="stylesheet">`
- Unsplash hero image given `loading="lazy"` attribute
- Project files moved into dedicated `resumesync/` folder

## [1.1] — 2026-05-16

### Added
- AI resume updater with find→replace review step (user approves each change)
- Cover letter regeneration with requested changes
- DOCX download for cover letter and full analysis report
- Additional guidance text area for cover letter generation
- Dark theme applied to all Streamlit form elements (dropdowns, textareas, file uploader)
- Job application tracker with Excel export

### Changed
- All file paths moved to `os.path.dirname(__file__)` for deployment compatibility

## [1.0] — Initial release

- Resume upload (PDF, DOCX, TXT)
- Job URL scraping and manual paste fallback
- AI compatibility analysis with score, keyword table, skills gap, ATS tips
- Basic cover letter generation
- VisualizePro Forest + Sage design system
