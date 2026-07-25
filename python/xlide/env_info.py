"""Runtime environment introspection for the XLIDE backend."""
from __future__ import annotations

from importlib import metadata
from typing import Any

# The runtime dependencies XLIDE installs from python/requirements.txt.
_TRACKED_PACKAGES = ("pyOpenVBA", "openpyxl")


def get_package_versions() -> dict[str, Any]:
    """Installed versions of XLIDE's Python dependencies.

    Returns {"packages": {"pyOpenVBA": "3.0.1", "openpyxl": "3.1.5"}}. A package
    that cannot be resolved is simply omitted so the caller treats it as
    unknown; this method must never fail the bridge.
    """
    packages: dict[str, str] = {}
    for name in _TRACKED_PACKAGES:
        try:
            packages[name] = metadata.version(name)
        except Exception:  # noqa: BLE001 - introspection is strictly best-effort
            continue
    return {"packages": packages}
