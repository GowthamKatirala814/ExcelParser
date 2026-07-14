import os
import re
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent

# Loads backend/.env if present (git-ignored - see .env.example for the
# expected keys). Real environment variables set outside the process
# (shell, systemd, Docker, etc.) always take precedence over the file.
load_dotenv(BASE_DIR / ".env")

MONGO_URI = os.environ.get("MONGO_URI", "mongodb://localhost:27017")
MONGO_DB_NAME = os.environ.get("MONGO_DB_NAME", "excel_color_extractor")

STORAGE_DIR = Path(os.environ.get("STORAGE_DIR", BASE_DIR / "storage"))
STORAGE_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_EXTENSIONS = {".xlsx", ".xls"}

# --- LLM structuring stage ---
# Provider is selected here rather than hardcoded at the call site, so a
# different LLM can be swapped in later by adding a provider class and
# pointing LLM_PROVIDER at it - no changes needed anywhere else.
LLM_PROVIDER = os.environ.get("LLM_PROVIDER", "gemini")


def _collect_gemini_keys():
    """
    Gathers every configured Gemini API key from whichever of these forms is
    present in the environment - all of them can be combined, in this order:
      - GEMINI_API_KEY           a single key (backward compatible)
      - GEMINI_API_KEYS          a comma-separated list of keys
      - GEMINI_API_KEY_1, _2, _3, ... any number of numbered keys
    Adding another key later means adding one more line to backend/.env -
    never a code change. Duplicates are dropped, order is preserved.
    """
    keys = []

    single = os.environ.get("GEMINI_API_KEY")
    if single and single.strip():
        keys.append(single.strip())

    csv_list = os.environ.get("GEMINI_API_KEYS")
    if csv_list:
        keys.extend(part.strip() for part in csv_list.split(",") if part.strip())

    numbered = []
    for name, value in os.environ.items():
        match = re.fullmatch(r"GEMINI_API_KEY_(\d+)", name)
        if match and value.strip():
            numbered.append((int(match.group(1)), value.strip()))
    numbered.sort(key=lambda pair: pair[0])
    keys.extend(value for _, value in numbered)

    seen = set()
    unique_keys = []
    for key in keys:
        if key not in seen:
            seen.add(key)
            unique_keys.append(key)
    return unique_keys


# All configured keys, in rotation order. Empty if none are set.
GEMINI_API_KEYS = _collect_gemini_keys()

# "gemini-2.5-flash" is blocked on some API keys/projects ("no longer
# available to new users" - a per-project rollout restriction, not a quota
# or key-validity issue) even though it's still listed in the model
# catalog. "gemini-flash-latest" is a stable alias that isn't subject to
# that restriction and is confirmed working; override via GEMINI_MODEL if
# your key needs a specific pinned version instead of the "latest" alias.
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-flash-latest")
GEMINI_API_BASE = os.environ.get("GEMINI_API_BASE", "https://generativelanguage.googleapis.com/v1beta")

# How long a key that just failed with a quota/rate-limit/auth error is
# skipped before being tried again.
GEMINI_KEY_COOLDOWN_SECONDS = int(os.environ.get("GEMINI_KEY_COOLDOWN_SECONDS", 60))

# Hard cap on the size of the raw JSON payload sent to the LLM, to fail fast
# with a clear error instead of an opaque request-too-large response.
LLM_MAX_INPUT_BYTES = int(os.environ.get("LLM_MAX_INPUT_BYTES", 15_000_000))
