"""Auto-Resume: Automated resume/form filler for Word documents."""

from __future__ import annotations

import sys
from pathlib import Path

__version__ = "1.0.0"


def get_resource_path(relative_path: str | Path) -> Path:
    """Get the absolute path to a bundled resource.

    Works both in development (running from source) and when packaged
    as a standalone executable via PyInstaller.

    Args:
        relative_path: Path relative to the package directory (e.g. "templates").

    Returns:
        Absolute Path to the resource.
    """
    if hasattr(sys, "_MEIPASS"):
        # Running as PyInstaller executable
        base = Path(sys._MEIPASS) / "auto_resume"
    else:
        # Running from source
        base = Path(__file__).parent
    return base / relative_path
