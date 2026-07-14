"""
Orchestrates the structuring stage: raw JSON in, structured JSON out (or a
StructuringError), independent of which LLM provider is configured.
"""
from .llm.base import StructuringError
from .llm.factory import get_structuring_provider

__all__ = ["StructuringError", "generate_structured_json"]


async def generate_structured_json(raw_json: dict) -> dict:
    provider = get_structuring_provider()
    return await provider.structure(raw_json)
