"""
Provider-agnostic contract for the structuring stage. Any LLM backend
(Gemini today, something else later) implements StructuringProvider; nothing
outside this package needs to know which one is in use.
"""
from abc import ABC, abstractmethod

# Coarse, UI-displayable categories for why structuring failed. Kept small
# and provider-agnostic so the router/UI can show a specific message
# ("token limit reached", "daily quota exceeded", ...) without knowing
# anything about Gemini's particular error shapes.
ERROR_KIND_TOKEN_LIMIT = "token_limit"
ERROR_KIND_QUOTA_EXCEEDED = "quota_exceeded"
ERROR_KIND_RATE_LIMITED = "rate_limited"
ERROR_KIND_TIMEOUT = "timeout"
ERROR_KIND_INVALID_JSON = "invalid_json"
ERROR_KIND_MALFORMED_RESPONSE = "malformed_response"
ERROR_KIND_INVALID_KEY = "invalid_key"
ERROR_KIND_NETWORK = "network_error"
ERROR_KIND_MODEL_UNAVAILABLE = "model_unavailable"
ERROR_KIND_OTHER = "other"


class StructuringError(Exception):
    """Raised for any failure to produce a structured JSON - network, auth,
    rate limit, or a response that isn't valid/parseable JSON. The raw
    extraction must remain unaffected by this; callers catch it and record
    the failure instead of letting it propagate."""

    def __init__(self, message, kind=ERROR_KIND_OTHER):
        super().__init__(message)
        self.kind = kind


class StructuringProvider(ABC):
    @abstractmethod
    async def structure(self, raw_json: dict) -> dict:
        """Converts a raw extraction JSON into a structured JSON dict, or
        raises StructuringError with a human-readable reason."""
        raise NotImplementedError
