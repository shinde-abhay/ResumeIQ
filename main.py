import os, json
from fastapi import FastAPI, UploadFile, File, Form, Header, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from starlette.concurrency import run_in_threadpool
import core

app = FastAPI(title="ResumeIQ API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=os.getenv("ALLOWED_ORIGINS", "*").split(","),   # set to your GitHub Pages origin in production
    allow_methods=["*"],
    allow_headers=["*"],
)


async def _read_pdf(file: UploadFile) -> bytes:
    if not (file.filename or "").lower().endswith((".pdf", ".docx")):
        raise HTTPException(400, "Upload a PDF or Word (.docx) file.")
    data = await file.read()
    if len(data) > core.MAX_BYTES:
        raise HTTPException(413, "PDF is larger than 5 MB.")
    return data


async def _run(fn, *args):
    try:
        return await run_in_threadpool(fn, *args)
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(502, f"Processing failed: {e}")


@app.get("/")
def health():
    return {"ok": True}


@app.post("/api/analyze")
async def analyze(file: UploadFile = File(...), jd: str = Form(...), x_groq_key: str | None = Header(None)):
    return await _run(core.analyze, await _read_pdf(file), jd, x_groq_key)


@app.post("/api/improve")
async def improve(
    file: UploadFile = File(...),
    jd: str = Form(...),
    edits: str = Form("[]"),
    extra_skills: str = Form("[]"),
    x_groq_key: str | None = Header(None),
):
    try:
        skills = [s for s in json.loads(extra_skills) if isinstance(s, str)][:25]
        chosen = [e for e in json.loads(edits)[:20]
                  if isinstance(e, dict) and isinstance(e.get("original"), str) and isinstance(e.get("suggested"), str)]
    except (json.JSONDecodeError, TypeError):
        skills, chosen = [], []
    return await _run(core.improve, await _read_pdf(file), jd, chosen, skills, x_groq_key)