"""
Resume Reboot - FastAPI Backend
"""

import os
import uuid
import json
import re
from pathlib import Path

import anthropic
import stripe
from fastapi.responses import Response
from docx_builder import build_resume_docx

# Fail loud at startup if required secrets are missing.
stripe.api_key = os.environ["STRIPE_SECRET_KEY"]
STRIPE_PUBLISHABLE_KEY = os.environ.get("STRIPE_PUBLISHABLE_KEY", "")
STRIPE_WEBHOOK_SECRET = os.environ["STRIPE_WEBHOOK_SECRET"]

_rewrite_cache: dict = {}
# Tracks redeemed Stripe checkout session IDs to block replay/reuse.
_used_stripe_sessions: set[str] = set()

import pdfplumber
from docx import Document
from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded

app = FastAPI(title="Resume Reboot")
limiter = Limiter(key_func=get_remote_address)
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")

UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

MAX_UPLOAD_BYTES = 10 * 1024 * 1024  # 10 MB

# Strict UUID4 pattern used to validate any session_id before it touches
# filesystem globs or dict lookups.
_UUID4_RE = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$"
)

client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])


def _safe_filename(filename: str) -> str:
    """Return a filename stripped of directory components and non-safe characters."""
    name = Path(filename).name          # strips any ../.. traversal
    name = re.sub(r"[^\w.\-]", "_", name)
    return name or "upload"


def extract_text_from_pdf(path: Path) -> str:
    text = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text.append(t)
    return "\n".join(text)


def extract_text_from_docx(path: Path) -> str:
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def parse_resume(path: Path, filename: str) -> str:
    ext = filename.lower().split(".")[-1]
    if ext == "pdf":
        return extract_text_from_pdf(path)
    elif ext in ("docx", "doc"):
        return extract_text_from_docx(path)
    else:
        raise ValueError("Unsupported file type. Upload PDF or DOCX.")


SCORE_PROMPT = """You are a senior technical recruiter and résumé coach with 15 years of experience.

Evaluate the following résumé for the target role of: {job_title}

Score three dimensions from 0-100:
1. ATS Compatibility: keyword presence, section headers, formatting signals, no tables/columns
2. Narrative Clarity: action verbs, quantified outcomes, concise bullets, readability
3. Role Fit: how well the candidate's experience maps to typical {job_title} requirements

For each score, provide exactly 2-3 specific observations — not generic advice.

Return ONLY valid JSON in this exact structure, no preamble, no markdown:
{{
  "ats_score": 0,
  "ats_observations": ["obs1", "obs2"],
  "clarity_score": 0,
  "clarity_observations": ["obs1", "obs2"],
  "role_fit_score": 0,
  "role_fit_observations": ["obs1", "obs2"],
  "overall_summary": "2-3 sentence honest assessment"
}}

RÉSUMÉ:
{resume_text}"""

REWRITE_PROMPT = """You are a professional résumé writer specializing in {industry} roles.

Rewrite the following résumé to target the role of: {job_title}

Rules:
- Lead every bullet with a strong past-tense action verb
- Quantify outcomes wherever the original hints at scale
- Align language to {job_title} role requirements without fabricating experience
- Preserve ALL factual content — never invent companies, titles, or dates
- Sharpen the summary/objective to speak directly to {job_title} hiring managers
- Remove weak filler phrases ("responsible for", "helped with", "worked on")

Return ONLY valid JSON in this exact structure, no preamble, no markdown:
{{
  "contact": "name and contact info as-is",
  "summary": "rewritten summary paragraph",
  "experience": [
    {{
      "company": "company name",
      "title": "job title",
      "dates": "date range",
      "bullets": ["rewritten bullet 1", "rewritten bullet 2"]
    }}
  ],
  "education": "education section as-is or lightly cleaned",
  "skills": ["skill1", "skill2"]
}}

RÉSUMÉ:
{resume_text}"""

KEYWORDS_PROMPT = """You are an ATS optimization specialist.

Compare this résumé against high-value ATS keywords commonly required for {job_title} roles.

Identify the 8-10 most impactful MISSING keywords. For each, suggest one natural place to incorporate it.

Return ONLY valid JSON in this exact structure, no preamble, no markdown:
{{
  "missing_keywords": [
    {{"keyword": "keyword here", "suggestion": "Where/how to add it naturally"}}
  ]
}}

RÉSUMÉ:
{resume_text}"""


def claude_call(prompt: str) -> dict:
    message = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )
    raw = message.content[0].text.strip()
    raw = re.sub(r"^```json\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw)


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/scan")
@limiter.limit("10/hour")
async def scan_resume(
    request: Request,
    file: UploadFile = File(...),
    job_title: str = Form(...),
):
    if not file.filename.lower().endswith((".pdf", ".docx")):
        raise HTTPException(400, "Upload PDF or DOCX only.")

    session_id = str(uuid.uuid4())

    # FIX 1: sanitize filename before constructing the save path to prevent
    # path traversal (e.g. "../../main.py.pdf" → "....main.py.pdf").
    safe_name = _safe_filename(file.filename)
    save_path = UPLOAD_DIR / f"{session_id}_{safe_name}"

    content = await file.read(MAX_UPLOAD_BYTES + 1)
    if len(content) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, "File exceeds 10 MB limit.")
    save_path.write_bytes(content)

    try:
        resume_text = parse_resume(save_path, file.filename)
    except Exception:
        raise HTTPException(400, "Could not parse the uploaded file.")

    if len(resume_text.strip()) < 100:
        raise HTTPException(400, "Could not extract text from file. Try a text-based PDF.")

    try:
        scores = claude_call(SCORE_PROMPT.format(
            job_title=job_title,
            resume_text=resume_text[:6000]
        ))
    except Exception:
        raise HTTPException(500, "Scoring failed. Please try again.")

    return JSONResponse({
        "session_id": session_id,
        "job_title": job_title,
        "scores": scores
    })


# FIX 3: /rewrite endpoint removed. It was unauthenticated and unpaid —
# any caller who knew their session_id could bypass Stripe entirely.
# All rewrites now flow exclusively through /success after webhook-verified payment.


@app.post("/create-checkout-session")
@limiter.limit("20/hour")
async def create_checkout_session(
    request: Request,
    session_id: str = Form(...),
    job_title: str = Form(...),
    industry: str = Form(default="technology"),
):
    if not _UUID4_RE.match(session_id):
        raise HTTPException(400, "Invalid session.")

    try:
        checkout_session = stripe.checkout.Session.create(
            payment_method_types=["card"],
            line_items=[{
                "price_data": {
                    "currency": "usd",
                    "product_data": {
                        "name": "Resume Reboot — Full Rewrite",
                        "description": f"AI-powered résumé rewrite targeting: {job_title}",
                    },
                    "unit_amount": 700,  # $7.00
                },
                "quantity": 1,
            }],
            mode="payment",
            # FIX 4a: session_id, job_title, and industry are stored in Stripe-side
            # metadata rather than in the success URL query string. /success reads
            # them back from Stripe, so they cannot be tampered with by the client.
            metadata={
                "resume_session_id": session_id,
                "job_title": job_title,
                "industry": industry,
            },
            # Only the Stripe-generated checkout session ID travels in the URL now.
            success_url="https://resumereboot.io/success?stripe_session_id={CHECKOUT_SESSION_ID}",
            cancel_url="https://resumereboot.io/?cancelled=true",
        )
        return JSONResponse({"checkout_url": checkout_session.url})
    except Exception:
        raise HTTPException(500, "Checkout failed. Please try again.")


# FIX 4b: Webhook endpoint with Stripe signature verification.
# Register this URL in the Stripe dashboard and set STRIPE_WEBHOOK_SECRET.
# The raw request body must be read before any parsing — do not use a Pydantic
# model or Body() here, as that would decode it and break signature verification.
@app.post("/stripe-webhook")
async def stripe_webhook(request: Request):
    payload = await request.body()
    sig_header = request.headers.get("stripe-signature", "")
    try:
        event = stripe.Webhook.construct_event(payload, sig_header, STRIPE_WEBHOOK_SECRET)
    except (ValueError, stripe.error.SignatureVerificationError):
        raise HTTPException(400, "Invalid webhook signature.")

    # Log confirmed payments for audit / future async worker flows.
    # /success is still the rewrite trigger; this provides a cryptographically
    # verified record that does not depend on the client redirect completing.
    if event["type"] == "checkout.session.completed":
        cs = event["data"]["object"]
        resume_session_id = (cs.get("metadata") or {}).get("resume_session_id", "")
        if resume_session_id and _UUID4_RE.match(resume_session_id):
            # Idempotency marker — safe to call multiple times for the same event.
            _used_stripe_sessions.discard(f"pending_{resume_session_id}")

    return JSONResponse({"received": True})


@app.get("/success")
@limiter.limit("5/hour")
async def success(request: Request, stripe_session_id: str):
    # --- Verify payment ---
    try:
        checkout_session = stripe.checkout.Session.retrieve(stripe_session_id)
    except Exception:
        raise HTTPException(402, "Payment verification failed.")

    if checkout_session.payment_status != "paid":
        raise HTTPException(402, "Payment not completed.")

    # --- Pull session params from Stripe metadata (client cannot alter these) ---
    metadata = (checkout_session.get("metadata") or {})
    session_id = metadata.get("resume_session_id", "")
    job_title = metadata.get("job_title", "")
    industry = metadata.get("industry", "technology")

    if not session_id or not _UUID4_RE.match(session_id):
        raise HTTPException(400, "Invalid payment session.")

    # FIX 4c: Prevent reuse of the same Stripe checkout session ID.
    # Once redeemed, a second hit to /success with the same stripe_session_id
    # (e.g. browser back, shared link) is rejected rather than running another
    # free Claude rewrite.
    if stripe_session_id in _used_stripe_sessions:
        raise HTTPException(409, "This payment has already been redeemed.")
    _used_stripe_sessions.add(stripe_session_id)

    # --- Run the rewrite ---
    matches = list(UPLOAD_DIR.glob(f"{session_id}_*"))
    if not matches:
        # Release the lock so support can reprocess if needed.
        _used_stripe_sessions.discard(stripe_session_id)
        raise HTTPException(404, "Session expired. Please re-upload your résumé.")

    try:
        resume_text = parse_resume(matches[0], matches[0].name)
        rewritten = claude_call(REWRITE_PROMPT.format(
            job_title=job_title,
            industry=industry,
            resume_text=resume_text[:6000]
        ))
        keywords = claude_call(KEYWORDS_PROMPT.format(
            job_title=job_title,
            resume_text=resume_text[:6000]
        ))
    except Exception:
        _used_stripe_sessions.discard(stripe_session_id)
        raise HTTPException(500, "Rewrite failed. Please contact support with your order ID.")

    _rewrite_cache[session_id] = {"rewritten": rewritten, "job_title": job_title}

    # FIX 2: Escape all server-derived values before embedding in a <script> block.
    # json.dumps produces a quoted, JSON-escaped JS value for strings; replacing
    # '</' with '<\/' prevents the literal token </script> from terminating the
    # script tag if it appears inside a string value.
    data_json = json.dumps(
        {"rewritten": rewritten, "keyword_gaps": keywords.get("missing_keywords", []), "session_id": session_id},
        ensure_ascii=True,
    ).replace("</", "<\\/")
    job_title_js = json.dumps(job_title, ensure_ascii=True).replace("</", "<\\/")

    template = templates.get_template("index.html")
    html = template.render({"request": request})

    inject = f"""
    <script>
    window.addEventListener('load', function() {{
        const data = {data_json};
        window._sessionId = data.session_id;
        window._jobTitle = {job_title_js};
        sessionId = data.session_id;
        jobTitle = {job_title_js};
        renderRewrite(data);
        document.getElementById('results').classList.add('show');
        document.getElementById('paywall').style.display = 'none';
        document.getElementById('rewrite-results').classList.add('show');
        document.getElementById('upload-card').style.display = 'none';
    }});
    </script>
    """
    html = html.replace("</body>", inject + "</body>")
    return HTMLResponse(html)


@app.get("/download/{session_id}")
async def download_docx(session_id: str):
    if not _UUID4_RE.match(session_id):
        raise HTTPException(400, "Invalid session.")
    if session_id not in _rewrite_cache:
        raise HTTPException(404, "No rewrite found for this session. Please rewrite first.")
    data = _rewrite_cache[session_id]
    docx_bytes = build_resume_docx(data, data.get("job_title", ""))
    return Response(
        content=docx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=resume_reboot.docx"}
    )


@app.get("/health")
async def health():
    return {"status": "ok"}
