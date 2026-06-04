# Email-sender — Job Application Automation

Reads a job description, tweaks your resume with targeted keywords, drafts a cold-application email, builds a polished `.docx`, and opens a pre-filled Gmail compose window in your browser. **You hit Send — the script never sends on your behalf.**

---

## Project Structure

```
Email-sender/
├── job_desc.txt          ← Edit this each time you apply somewhere
├── resume.txt            ← Your master resume (plain text template)
├── config.py             ← API keys & settings (fill in your own values)
├── email_script.py       ← Main script
├── Email_test.py         ← Test/debug script
├── requirements.txt      ← Python dependencies
├── .gitignore
├── client_secret.json    ← (you create this – see Gmail setup below)
└── token.json            ← (auto-generated on first run)
```

---

## One-time Setup

### 1 — Python environment

```bash
cd Email-sender
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

---

### 2 — Configure API keys in `config.py`

Open `config.py` and fill in the values for whichever LLM provider(s) you want to use. Set `LLM_PROVIDER` to `"anthropic"`, `"gemini"`, or `"huggingface"`.

#### Anthropic (Claude)

1. Go to [https://console.anthropic.com/](https://console.anthropic.com/) and sign in
2. Navigate to **API Keys** → **Create Key**
3. Copy the key and paste it into `config.py`:
   ```python
   ANTHROPIC_API_KEY = "sk-ant-..."
   ```
4. Choose your model (see [Changing the AI Model](#changing-the-ai-model) below)

#### Google Gemini

1. Go to [https://aistudio.google.com/apikey](https://aistudio.google.com/apikey) and sign in
2. Click **Create API key** → select or create a project
3. Copy the key and paste it into `config.py`:
   ```python
   GOOGLE_API_KEY   = "your-key-here"
   GOOGLE_API_KEY_2 = "your-backup-key-here"   # optional fallback
   ```

#### HuggingFace

1. Go to [https://huggingface.co/settings/tokens](https://huggingface.co/settings/tokens) and sign in
2. Click **New token** → select **Read** access → **Create token**
3. Copy the token and paste it into `config.py`:
   ```python
   HUGGINGFACE_API_KEY = "hf_..."
   ```

#### Gmail sender address

Set your Gmail address in `config.py`:
```python
SENDER_GMAIL = "your-email@gmail.com"
```

#### Output resume filename

Set the output filename for your tweaked resume `.docx`:
```python
OUTPUT_RESUME_PATH = "FirstName_LastName_Resume.docx"
```

---

### 3 — Google Gmail API credentials

This is a one-time setup that lets the script open a pre-filled Gmail compose window in your browser.

#### 3a. Create a Google Cloud project

1. Go to [https://console.cloud.google.com/](https://console.cloud.google.com/)
2. Click **New Project** → name it (e.g. `email-sender`) → **Create**
3. Make sure the new project is selected in the top dropdown

#### 3b. Enable the Gmail API

1. In the left sidebar: **APIs & Services → Library**
2. Search for `Gmail API` → click it → **Enable**

#### 3c. Configure the OAuth consent screen

1. **APIs & Services → OAuth consent screen**
2. User Type: **External** → **Create**
3. Fill in:
   - App name: `Email Sender`
   - User support email: your Gmail
   - Developer contact: your Gmail
4. Click **Save and Continue** through all steps (Scopes, Test users)
5. On **Test users**: click **Add users** → add your Gmail address
6. "Testing" mode is fine for personal use — **Publish App** is not required

#### 3d. Create OAuth client credentials

1. **APIs & Services → Credentials → Create Credentials → OAuth client ID**
2. Application type: **Desktop app**
3. Name: `email-sender-desktop`
4. **Create** → click **Download JSON**
5. Rename the downloaded file to `client_secret.json`
6. Move it into the `Email-sender/` folder

#### 3e. First run (browser auth)

The first time you run the script, a browser window will open asking you to sign in to Google and grant permission. After approval, a `token.json` file is created and cached — you won't need to do this again unless you revoke access.

> **Note:** `client_secret.json` and `token.json` are listed in `.gitignore` and should never be committed.

---

## Daily Usage

### 1 — Edit `job_desc.txt`

Fill in the header fields at the top:

```
RECRUITER_EMAIL: recruiter@company.com
RECRUITER_NAME: Jane Smith
COMPANY: Amazon Web Services
LOCATION: Seattle, WA

---JOB DESCRIPTION BELOW---
Paste the full JD here...
```

### 2 — Run the script

```bash
cd Email-sender
source venv/bin/activate          # Windows: venv\Scripts\activate
python email_script.py
```

### 3 — Review & Send

- The script opens Gmail in your browser with the draft pre-loaded
- Subject, body, recipient, and your tweaked resume are already filled in
- **Read the email carefully, then hit Send**

---

## Changing the AI Model

Edit `LLM_PROVIDER` and the corresponding model in `config.py`:

```python
# Use Anthropic Claude
LLM_PROVIDER  = "anthropic"
ANTHROPIC_MODEL = "claude-opus-4-8"    # Best quality (slower, higher cost)
ANTHROPIC_MODEL = "claude-sonnet-4-6"  # Balanced (recommended)
ANTHROPIC_MODEL = "claude-haiku-4-5"   # Fastest / cheapest

# Use Google Gemini
LLM_PROVIDER  = "gemini"
GEMINI_MODEL  = "gemini-2.5-flash"

# Use HuggingFace
LLM_PROVIDER      = "huggingface"
HUGGINGFACE_MODEL = "openai/gpt-oss-120b:cerebras"
```

---

## Tips

- **Keep `resume.txt` as your truthful master** — the AI only adds/rewords based on what you already have. Review the `.docx` before every send.
- **Rotate `job_desc.txt`** — just overwrite it with the next JD and re-run.
- **Never commit `client_secret.json` or `token.json`** — they give access to your Gmail. They are already covered by `.gitignore`.
- If the resume JSON parse fails, the model output will be printed so you can debug the prompt.

---

## Troubleshooting

| Error | Fix |
|-------|-----|
| `ModuleNotFoundError` | Run `pip install -r requirements.txt` inside your venv |
| `invalid_client` (Google) | Re-download `client_secret.json` — old one may be stale |
| `Token has been expired` | Delete `token.json` and re-run to re-authenticate |
| `JSONDecodeError` | Model returned malformed JSON; re-run or switch to a smarter model |
| Draft opens but no attachment | Check that the output `.docx` was created in the same folder |
