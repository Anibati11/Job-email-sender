#!/usr/bin/env python3
"""
email_script.py
───────────────
Reads job_desc.txt and resume.txt, calls the Anthropic API to:
  1. Tweak the resume (keyword injection + location update)
  2. Generate an email subject and body
Then:
  3. Builds a polished .docx of the tweaked resume
  4. Opens Gmail compose window in your browser with subject, body,
     recipient, and the resume attached (base64-encoded via mailto / Gmail API)

You hit Send. The script never sends on your behalf.
"""

import os, sys, re, base64, json, webbrowser, subprocess, textwrap, time
from json_repair import repair_json
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from google import genai
from openai import OpenAI

# ── third-party ──────────────────────────────────────────────────────────────
try:
    import anthropic
except ImportError:
    sys.exit("Missing dependency: pip install anthropic")

import requests

try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    sys.exit("Missing dependency: pip install python-docx")

try:
    from google.oauth2.credentials      import Credentials
    from google_auth_oauthlib.flow      import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery      import build
    from googleapiclient.errors         import HttpError
except ImportError:
    sys.exit("Missing dependency: pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client")

from config import (
    ANTHROPIC_API_KEY, ANTHROPIC_MODEL,
    GOOGLE_API_KEY, GOOGLE_API_KEY_2, GEMINI_MODEL, LLM_PROVIDER,
    HUGGINGFACE_API_KEY, HUGGINGFACE_MODEL,
    SENDER_GMAIL, GOOGLE_CREDENTIALS_FILE, GOOGLE_TOKEN_FILE,
    GMAIL_SCOPES, OUTPUT_RESUME_PATH,
)

# ─────────────────────────────────────────────────────────────────────────────
# 1.  File helpers
# ─────────────────────────────────────────────────────────────────────────────

def read_file(path: str) -> str:
    p = Path(path)
    if not p.exists():
        sys.exit(f"File not found: {path}")
    return p.read_text(encoding="utf-8").strip()


def parse_job_desc(raw: str) -> dict:
    system = textwrap.dedent("""
        You are an expert at parsing job descriptions.
        Extract the following fields and return ONLY valid JSON, no markdown fences:
        {
          "RECRUITER_NAME":  "full name of the recruiter/sender, or empty string if not found",
          "RECRUITER_EMAIL": "recruiter email address, or empty string if not found",
          "COMPANY":         "hiring company name, or recruiter's company if client company not mentioned",
          "LOCATION":        "job location, e.g. Dallas, TX or Remote",
          "JD_BODY":         "the full job description text verbatim"
        }
    """)
    raw_json = call_llm(system, raw)
    return json.loads(repair_json(extract_json(raw_json)))


# ─────────────────────────────────────────────────────────────────────────────
# 2.  LLM helpers
# ─────────────────────────────────────────────────────────────────────────────

def extract_json(raw: str) -> str:
    """Strip think blocks, markdown fences, and extract the first JSON object."""
    # Remove <tool_call>...<tool_call> blocks (Qwen/reasoning models)
    raw = re.sub(r"<tool_call>.*?</tool_call>", "", raw, flags=re.DOTALL).strip()
    # Remove markdown fences
    raw = re.sub(r"^```[a-z]*\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)
    # Extract first {...} block in case model added surrounding text
    match = re.search(r"\{.*\}", raw, re.DOTALL)
    if match:
        return match.group(0)
    return raw

def call_llm(system: str, user: str) -> str:
    if LLM_PROVIDER == "anthropic":
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        msg = client.messages.create(
            model      = ANTHROPIC_MODEL,
            max_tokens = 4096,
            system     = system,
            messages   = [{"role": "user", "content": user}],
        )
        return msg.content[0].text.strip()
    elif LLM_PROVIDER == "huggingface":
        client = OpenAI(
                        base_url="https://router.huggingface.co/v1",
                        api_key=HUGGINGFACE_API_KEY,
                        )
        for attempt in range(3):
            try:
                response = client.chat.completions.create(
                    model=HUGGINGFACE_MODEL,
                    messages=[
                        {"role": "system", "content": system},
                        {"role": "user",   "content": user},
                    ],
                )
                return response.choices[0].message.content
            except Exception as e:
                if "429" in str(e) and attempt < 2:
                    wait = 10 * (attempt + 1)
                    print(f"   HuggingFace rate limited — retrying in {wait}s …")
                    time.sleep(wait)
                else:
                    raise
    else:
        for api_key in [GOOGLE_API_KEY, GOOGLE_API_KEY_2]:
            try:
                client = genai.Client(api_key=api_key)
                response = client.models.generate_content(
                    model    = GEMINI_MODEL,
                    contents = user,
                    config   = genai.types.GenerateContentConfig(system_instruction=system),
                )
                return response.text.strip()
            except (genai.errors.ClientError, genai.errors.ServerError) as e:
                if e.code in (429, 503):
                    continue
                raise
        sys.exit("All Gemini API keys exhausted or unavailable. Try again later.")


def tweak_resume(resume_raw: str, jd: dict) -> dict:
    """
    Asks Claude to:
      - Inject relevant keywords from the JD naturally
      - Update the location to match the job
      - Return structured JSON with sections
    """
    system = textwrap.dedent("""
        You are an expert resume writer and ATS optimizer.
        Given a resume and a job description, you will:
        1. Naturally weave in missing keywords/technologies from the JD (never fabricate experience).
        2. Update the candidate's listed location to match the job location.
        3. Keep all bullet points truthful and specific.
        4. Return ONLY valid JSON with this exact structure – no markdown fences, no extra text.
           IMPORTANT: use exactly these key names, do not rename or add keys:
        {
          "name":          "...",
          "contact": {
            "email":    "...",
            "phone":    "...",
            "linkedin": "...",
            "github":   "...",
            "location": "..."
          },
          "summary": ["bullet1", "bullet2", "..."],
          "experience": [
            {
              "title":    "...",
              "company":  "...",
              "dates":    "...",
              "location": "...",
              "bullets":  ["...", "..."]
            }
          ],
          "education": [
            {"degree": "...", "school": "...", "year": "..."}
          ],
          "skills": {
            "cloud_platforms":    "...",
            "languages":          "...",
            "databases":          "...",
            "data_processing":    "...",
            "ml_data_science":    "...",
            "containerization":   "...",
            "version_control":    "...",
            "operating_systems":  "..."
          },
          "certifications": ["..."]
        }
        IMPORTANT REQUIREMENTS:
        - "summary" must be a JSON array of exactly 20 concise bullet strings (no leading dashes).
        - Each experience entry's "bullets" array must contain AT LEAST 16 items.
    """)

    user = f"""
JOB LOCATION: {jd.get('LOCATION', '')}
COMPANY: {jd.get('COMPANY', '')}

===JOB DESCRIPTION===
{jd['JD_BODY']}

===CURRENT RESUME===
{resume_raw}
"""
    raw_json = call_llm(system, user)
    return json.loads(repair_json(extract_json(raw_json)))


def generate_email(resume_data: dict, jd: dict) -> dict:
    """Returns {"subject": "...", "body": "..."}"""
    system = textwrap.dedent("""
        You are a professional job-application email writer.
        Write a concise, compelling cold-application email from a candidate to a recruiter.
        Rules:
        - Subject line: punchy, role-specific, under 12 words.
        - Body: 3–4 short paragraphs. No fluff. Highlight 2-3 specific technical strengths
          that match the JD. Mention the resume is attached.
        - End with a professional sign-off using the candidate's name.
        - Return ONLY valid JSON, no markdown fences:
          {"subject": "...", "body": "..."}
    """)

    user = f"""
CANDIDATE: {resume_data['name']}
ROLE: {jd.get('JD_BODY', '')[:400]}
COMPANY: {jd.get('COMPANY', '')}
RECRUITER: {jd.get('RECRUITER_NAME', 'Hiring Manager')}
CANDIDATE SUMMARY: {resume_data['summary']}
TOP SKILLS: {resume_data['skills'].get('big_data','')} | {resume_data['skills'].get('cloud','')}
"""
    raw = call_llm(system, user)
    return json.loads(repair_json(extract_json(raw)))


# ─────────────────────────────────────────────────────────────────────────────
# 3.  Build .docx resume
# ─────────────────────────────────────────────────────────────────────────────

ACCENT = RGBColor(0x1F, 0x5C, 0x99)   # professional dark-blue

def _heading(doc, text, size=13, bold=True, color=ACCENT, space_before=8, space_after=2):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    run = p.add_run(text.upper())
    run.bold      = bold
    run.font.size = Pt(size)
    run.font.color.rgb = color
    # Underline divider via bottom border (paragraph style)
    from docx.oxml.ns import qn
    from docx.oxml    import OxmlElement
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'),   'single')
    bottom.set(qn('w:sz'),    '4')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '1F5C99')
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def _bullet(doc, text, indent=0.25):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent  = Inches(indent)
    p.paragraph_format.space_after  = Pt(1)
    run = p.add_run(text)
    run.font.size = Pt(10)
    return p


def build_docx(data: dict, output_path: str):
    doc = Document()

    # ── Page margins ──
    for section in doc.sections:
        section.top_margin    = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin   = Inches(0.85)
        section.right_margin  = Inches(0.85)

    # ── Name ──
    name_p = doc.add_paragraph()
    name_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_r = name_p.add_run(data['name'])
    name_r.bold           = True
    name_r.font.size      = Pt(18)
    name_r.font.color.rgb = ACCENT

    # ── Contact line ──
    c = data['contact']
    contact_line = " | ".join(filter(None, [
        c.get('location',''), c.get('email',''),
        c.get('phone',''),    c.get('linkedin',''), c.get('github',''),
    ]))
    cp = doc.add_paragraph(contact_line)
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.runs[0].font.size = Pt(9)
    cp.paragraph_format.space_after = Pt(4)

    # ── Summary ──
    _heading(doc, "Professional Summary")
    summary = data['summary']
    if isinstance(summary, list):
        for point in summary:
            _bullet(doc, point)
    else:
        # fallback: split on sentence boundaries if LLM returned a string
        for point in re.split(r'(?<=[.!?])\s+', summary.strip()):
            if point:
                _bullet(doc, point)

    # ── Experience ──
    _heading(doc, "Experience")
    for job in data.get('experience', []):
        # Role | Company | Dates | Location
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after  = Pt(1)
        r1 = p.add_run(f"{job.get('title','')}  —  {job.get('company','')}")
        r1.bold = True; r1.font.size = Pt(10.5)
        r2 = p.add_run(f"   {job.get('dates','')}  |  {job.get('location','')}")
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        for bullet in job.get('bullets', []):
            _bullet(doc, bullet)

    # ── Education ──
    _heading(doc, "Education")
    for edu in data.get('education', []):
        ep = doc.add_paragraph()
        ep.paragraph_format.space_before = Pt(3)
        ep.paragraph_format.space_after  = Pt(1)
        r = ep.add_run(f"{edu.get('degree','')}  —  {edu.get('school','')}  ({edu.get('year','')})")
        r.font.size = Pt(10)

    # ── Skills ──
    _heading(doc, "Technical Skills")
    skills = data.get('skills', {})
    skill_map = {
        "Cloud Platforms":                  skills.get('cloud_platforms', ''),
        "Programming Languages":            skills.get('languages', ''),
        "Databases":                        skills.get('databases', ''),
        "Data Processing & Integration":    skills.get('data_processing', ''),
        "Machine Learning & Data Science":  skills.get('ml_data_science', ''),
        "Containerization & Orchestration": skills.get('containerization', ''),
        "Version Control & CI/CD":          skills.get('version_control', ''),
        "Operating Systems":                skills.get('operating_systems', ''),
    }
    for label, value in skill_map.items():
        if value:
            sp2 = doc.add_paragraph()
            sp2.paragraph_format.space_after = Pt(1)
            bold_r = sp2.add_run(f"{label}: ")
            bold_r.bold = True; bold_r.font.size = Pt(10)
            sp2.add_run(value).font.size = Pt(10)

    # ── Certifications ──
    certs = [c for c in data.get('certifications', []) if c and c.strip() and c.strip() != '...']
    if certs:
        _heading(doc, "Certifications")
        for cert in certs:
            _bullet(doc, cert)

    doc.save(output_path)
    print(f"✅ Resume saved → {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
# 4.  Gmail API – open compose window (you hit Send)
# ─────────────────────────────────────────────────────────────────────────────

def get_gmail_service():
    creds = None
    if Path(GOOGLE_TOKEN_FILE).exists():
        creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_FILE, GMAIL_SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                GOOGLE_CREDENTIALS_FILE, GMAIL_SCOPES
            )
            creds = flow.run_local_server(port=0)
        Path(GOOGLE_TOKEN_FILE).write_text(creds.to_json())
    return build("gmail", "v1", credentials=creds)


def create_draft(service, to: str, subject: str, body: str, attachment_path: str) -> str:
    """Creates a Gmail draft (with attachment) and opens it in the browser."""
    msg = MIMEMultipart()
    msg["to"]      = to
    msg["from"]    = SENDER_GMAIL
    msg["subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    # Attach the .docx
    with open(attachment_path, "rb") as f:
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.wordprocessingml.document")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{Path(attachment_path).name}"')
    msg.attach(part)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    draft = service.users().drafts().create(
        userId="me",
        body={"message": {"raw": raw}}
    ).execute()

    draft_id = draft["id"]
    print(f"✅ Gmail draft created  (id: {draft_id})")

    # Build the Gmail compose URL and open it
    compose_url = f"https://mail.google.com/mail/u/0/#drafts/{draft['message']['id']}"
    print(f"🌐 Opening browser → {compose_url}")
    webbrowser.open(compose_url)
    return draft_id


# ─────────────────────────────────────────────────────────────────────────────
# 5.  Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("\n📄 Reading files …")
    jd_raw     = read_file("job_desc.txt")
    resume_raw = read_file("resume.txt")
    jd         = parse_job_desc(jd_raw)

    print(f"   Company:    {jd.get('COMPANY','?')}")
    print(f"   Location:   {jd.get('LOCATION','?')}")
    print(f"   Recruiter:  {jd.get('RECRUITER_NAME','?')}  <{jd.get('RECRUITER_EMAIL','?')}>")

    print("\n🤖 Calling Gemini to tweak resume …")
    resume_data = tweak_resume(resume_raw, jd)
    print(f"   Name:       {resume_data['name']}")
    print(f"   Location:   {resume_data['contact'].get('location', resume_data['contact'].get('city', '?'))}")

    print("\n🤖 Calling Gemini to draft email …")
    email_data = generate_email(resume_data, jd)
    print(f"   Subject:    {email_data['subject']}")

    print("\n📝 Building .docx resume …")
    build_docx(resume_data, OUTPUT_RESUME_PATH)

    print("\n🔐 Authenticating with Gmail …")
    service = get_gmail_service()

    print("\n📬 Creating Gmail draft …")
    create_draft(
        service,
        to              = jd["RECRUITER_EMAIL"],
        subject         = email_data["subject"],
        body            = email_data["body"],
        attachment_path = OUTPUT_RESUME_PATH,
    )

    print("\n✨ Done! Review the draft in your browser and hit Send when ready.\n")


if __name__ == "__main__":
    main()
