"""
Process-wide pool of API keys with cooldown tracking. A fresh
StructuringProvider instance is built on every request (see factory.py), so
this module - not the provider - is where cooldown state actually has to
live for "skip a key that just failed" and "resume with remaining keys
across sheets" to work at all.

Safe without explicit locking: every method here is synchronous (no `await`
inside), so within Python's single-threaded asyncio event loop each call
runs to completion without interleaving - no torn reads/writes are possible
even when multiple workbooks are being structured concurrently.
"""
import logging
import time

logger = logging.getLogger("app.structuring.gemini.keys")


def _mask(key):
    return f"...{key[-4:]}" if len(key) >= 4 else "***"


class ApiKeyPool:
    def __init__(self, keys, cooldown_seconds):
        self._keys = list(dict.fromkeys(keys))  # de-dup, preserve order
        self._cooldown_seconds = cooldown_seconds
        self._unavailable_until = {}  # key -> monotonic seconds

    def total_keys(self):
        return len(self._keys)

    def available_keys(self):
        """Keys not currently in cooldown, in configured order."""
        now = time.monotonic()
        return [k for k in self._keys if self._unavailable_until.get(k, 0.0) <= now]

    def mark_cooldown(self, key, reason):
        self._unavailable_until[key] = time.monotonic() + self._cooldown_seconds
        logger.warning(
            "Gemini API key %s put on %ds cooldown (%s)", _mask(key), self._cooldown_seconds, reason
        )


_pool_instance = None
_pool_signature = None


def get_key_pool(keys, cooldown_seconds):
    """
    Lazily builds one process-wide pool and reuses it across every call -
    cooldown state must persist between sheets and between separately
    constructed provider instances. Rebuilt only if the configured
    keys/cooldown actually change (e.g. .env edited and the process reloaded).
    """
    global _pool_instance, _pool_signature
    signature = (tuple(keys), cooldown_seconds)
    if _pool_instance is None or _pool_signature != signature:
        _pool_instance = ApiKeyPool(keys, cooldown_seconds)
        _pool_signature = signature
    return _pool_instance
