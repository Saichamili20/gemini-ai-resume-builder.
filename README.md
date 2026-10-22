# AI Resume & Career Assistant

A Gradio web app powered by Gemini. It combines ATS-friendly resume generation with job-match analysis and career preparation tools.

## Features

1. **Resume Builder** — generates a Word resume from user-provided facts.
2. **ATS Resume Analyzer** — accepts PDF or DOCX resumes (or pasted text) and a job description; returns an estimated match score, score breakdown, matched/missing keywords, skill gaps, and improvement suggestions.
3. **Skill Advisor** — suggests role-related hard and soft skills, explains why they matter, and offers truthful resume phrasing examples.
4. **Projects & Courses** — suggests portfolio projects, learning resources, and a starter roadmap. Verify links and course availability before relying on them.
5. **Resume Improver** — improves summaries and bullet points without inventing experience or achievements.
6. **Interview Coach** — creates role-specific technical, scenario, and HR practice questions.
7. **Career Dashboard** — turns resume evidence and a target role into a practical preparation plan.

## Local setup

Requires Python 3.10 or newer.

```bash
pip install -r requirements.txt
```

Set the Gemini key as an environment variable, then run the app.

**Windows PowerShell:**
```powershell
$env:GEMINI_API_KEY="your-key-here"
python app.py
```

**macOS/Linux:**
```bash
export GEMINI_API_KEY="your-key-here"
python app.py
```

Open the local URL shown in the terminal.

## Render deployment

- Build command: `pip install -r requirements.txt`
- Start command: `python app.py`
- Add environment variable `GEMINI_API_KEY` in the Render service's Environment settings.
- Keep the key out of source code and GitHub.
- Commit `app.py`, `requirements.txt`, and `README.md` to the same GitHub repository/branch connected to Render. Render should redeploy after the commit.

## Notes and limitations

- The ATS score is an AI-generated estimate of alignment with the supplied job description, not an official ATS score or a hiring prediction.
- Scanned/image-only PDFs may not contain extractable text. Use a text-based PDF or paste the resume text into the fallback field.
- AI-generated course links and recommendations should be checked before use; availability and pricing can change.
- Resume content is sent to the configured Gemini API when the user requests analysis or generation. Do not upload sensitive information unnecessarily.
- No persistent user account, database, or progress storage is implemented in this version; the Career Dashboard generates an action plan but does not save checklists between visits.
