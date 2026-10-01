# ResumeIQ

Static HTML page (`docs/`) + FastAPI backend (`backend/`).

## How it works

1. **Analyze**: Groq extracts hard-skill keywords from the job description (cached per JD). Your resume is matched against them, so the score is the share of keywords found. Groq also suggests up to 6 single-line replacements and 5 interview questions. Suggested lines that don't exist in your resume are dropped.
2. **Apply**: only the ticked lines are replaced. No AI call, no full rewrite.
   - **PDF upload**: edited directly. The old line is covered and the new text is written at the same position, so the layout stays identical. Edited lines use a standard font (Helvetica/Times/Courier style), so they may look slightly different from a custom font.
   - **Word (.docx) upload**: paragraphs are edited with python-docx, keeping formatting. If LibreOffice is installed you also get a PDF; otherwise you get the .docx (open it in Word and Save as PDF).
3. The result is re-scored with the same keywords (before and after).

## Run locally

```bash
cd backend
pip install -r requirements.txt
set GROQ_API_KEY=gsk_...          # Windows (use export on macOS/Linux), or paste a key in the page
uvicorn main:app --reload
```

Open `docs/index.html` in your browser (it talks to `http://localhost:8000`). LibreOffice is only needed for the optional PDF export of .docx uploads.

## Notes

- 5 MB upload limit. Text-based files only (no scanned images).
- Suggestions that use a missing keyword are only offered when the line already shows that skill. Skills you tick are added to the end of the skills line that matches most keywords.
- Set `GROQ_MODEL` to change the model. Set `ALLOWED_ORIGINS` if you host the page somewhere.
- To deploy: backend via the included Dockerfile (Render / Hugging Face Space), page via GitHub Pages `/docs` after setting the backend URL in `index.html`.