"""Runtime build identity for Structures Tools.

Resolution order:
1. _build_info.py  — baked in at PyInstaller freeze time
2. git rev-parse   — live SHA during development (appends -dirty if needed)
3. "unknown"       — last-resort fallback
"""
import subprocess


def _get_build() -> str:
    try:
        from _build_info import __build__
        return __build__
    except ImportError:
        pass

    try:
        sha = subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"],
            stderr=subprocess.DEVNULL,
        ).decode().strip()
        dirty = subprocess.call(
            ["git", "diff", "--quiet"],
            stderr=subprocess.DEVNULL,
        ) != 0
        return sha + ("-dirty" if dirty else "")
    except Exception:
        pass

    return "unknown"


__build__ = _get_build()
