"""
Erebus.Helper Package

Standalone post-payload processor for Erebus framework
Handles LNK creation, MSI backdooring, and Windows plugin execution
"""

__version__ = "0.1.0"
__author__ = "Lavender-exe, hunterino-sec, Whispergate"

from pathlib import Path

# Package metadata
PACKAGE_DIR = Path(__file__).parent.resolve()
MODULES_DIR = PACKAGE_DIR / "modules"

__all__ = [
    "PACKAGE_DIR",
    "MODULES_DIR",
]
