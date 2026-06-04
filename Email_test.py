import os
from openai import OpenAI
from config import (
    ANTHROPIC_API_KEY, ANTHROPIC_MODEL,
    GOOGLE_API_KEY, GOOGLE_API_KEY_2, GEMINI_MODEL, LLM_PROVIDER,
    HUGGINGFACE_API_KEY, HUGGINGFACE_MODEL,
    SENDER_GMAIL, GOOGLE_CREDENTIALS_FILE, GOOGLE_TOKEN_FILE,
    GMAIL_SCOPES, OUTPUT_RESUME_PATH,
)

client = OpenAI(
    base_url="https://router.huggingface.co/v1",
    api_key= HUGGINGFACE_API_KEY,
)

response = client.chat.completions.create(
    model="openai/gpt-oss-120b:cerebras",
    messages=[{"role": "user", "content": "Tell me a fun fact about the Eiffel Tower."}],
)

print(response.choices[0].message.content)

