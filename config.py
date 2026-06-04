# ============================================================
#  config.py  –  Edit these values before running the script
# ============================================================

# -----------------------------------------------------------
# Anthropic API
# -----------------------------------------------------------
ANTHROPIC_API_KEY = "your-anthropic-api-key-here"   # https://console.anthropic.com/
ANTHROPIC_MODEL   = "claude-opus-4-5"

GOOGLE_API_KEY    = "your-google-gemini-api-key"     # https://aistudio.google.com/apikey
GOOGLE_API_KEY_2 = "your-google-gemini-api-key-2"   # optional second key for fallback
GEMINI_MODEL      = "gemini-2.5-flash"

HUGGINGFACE_API_KEY   = "your-huggingface-api-key"  # https://huggingface.co/settings/tokens
HUGGINGFACE_MODEL     = "openai/gpt-oss-120b:cerebras"

# "gemini", "anthropic", or "huggingface"
LLM_PROVIDER      = "huggingface"

# -----------------------------------------------------------
# Gmail / Google OAuth
# -----------------------------------------------------------
# The Gmail account you want to send FROM
SENDER_GMAIL = "your-email@gmail.com"

# Path to the credentials JSON you downloaded from Google Cloud Console
# (See README.md for setup instructions)
GOOGLE_CREDENTIALS_FILE = "client_secret.json"

# Token cache – created automatically after first OAuth login
GOOGLE_TOKEN_FILE = "token.json"

# Gmail OAuth scopes needed  (compose only – script never reads your mail)
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.compose"]

# -----------------------------------------------------------
# Output
# -----------------------------------------------------------
# Where the tweaked resume .docx will be saved
OUTPUT_RESUME_PATH = "Your_Name_Resume.docx"
