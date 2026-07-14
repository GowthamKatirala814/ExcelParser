"""
Single place that decides which LLM backs the structuring stage. Swapping
providers means adding a new module (implementing StructuringProvider) and a
branch here - nothing in the router or structuring_service changes.
"""
from ... import config
from .base import ERROR_KIND_INVALID_KEY, StructuringError, StructuringProvider
from .gemini_provider import GeminiStructuringProvider
from .key_pool import get_key_pool


def get_structuring_provider() -> StructuringProvider:
    provider = (config.LLM_PROVIDER or "").lower()

    if provider == "gemini":
        if not config.GEMINI_API_KEYS:
            raise StructuringError(
                "No Gemini API key is configured",
                kind=ERROR_KIND_INVALID_KEY,
            )
        key_pool = get_key_pool(config.GEMINI_API_KEYS, config.GEMINI_KEY_COOLDOWN_SECONDS)
        return GeminiStructuringProvider(
            key_pool=key_pool,
            model=config.GEMINI_MODEL,
            api_base=config.GEMINI_API_BASE,
            max_input_bytes=config.LLM_MAX_INPUT_BYTES,
        )

    raise ValueError(f"Unknown LLM_PROVIDER '{config.LLM_PROVIDER}'")
