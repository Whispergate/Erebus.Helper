"""
Erebus.Helper - CHM Compilation Module

Compiles a CHM (Compiled HTML Help) project using hhc.exe - the HTML Help
Compiler bundled with HTML Help Workshop (free Microsoft download).

Called by main.py `chm` subcommand:
    python erebus_helper.py chm --project-dir <dir> --output document.chm

hhc.exe resolution order:
    1. PATH (where hhc.exe)
    2. Default HTML Help Workshop install: C:\\Program Files (x86)\\HTML Help Workshop\\hhc.exe

Note: hhc.exe exits with code 1 even on success when there are only warnings.
We treat exit code 1 as success if the output file exists.
"""

import logging
import os
import shutil
import subprocess
from pathlib import Path
from typing import Optional, Tuple

logger = logging.getLogger(__name__)

# Default HTML Help Workshop install path
_HHW_DEFAULT = r"C:\Program Files (x86)\HTML Help Workshop\hhc.exe"


def find_hhc() -> Optional[str]:
    """Locate hhc.exe on the system."""
    # 1. PATH
    hhc = shutil.which("hhc.exe") or shutil.which("hhc")
    if hhc:
        return hhc
    # 2. Default install
    if os.path.exists(_HHW_DEFAULT):
        return _HHW_DEFAULT
    return None


def compile_chm(
    project_dir: str,
    output_path: str,
    hhp_file: str = "project.hhp",
) -> Tuple[bool, str]:
    """Compile a CHM project using hhc.exe.

    Args:
        project_dir:  Directory containing project.hhp, toc.hhc, default.html.
        output_path:  Destination for the compiled .chm file.
        hhp_file:     Name of the .hhp project file (default: project.hhp).

    Returns:
        (success: bool, message: str)
    """
    project_dir = Path(project_dir)
    output_path = Path(output_path)

    hhp_path = project_dir / hhp_file
    if not hhp_path.exists():
        return False, f"HHP project file not found: {hhp_path}"

    hhc = find_hhc()
    if not hhc:
        return False, (
            "hhc.exe not found. Install HTML Help Workshop: "
            "https://www.microsoft.com/en-us/download/details.aspx?id=21138"
        )

    logger.info(f"Compiling CHM: {hhp_path} using {hhc}")

    try:
        result = subprocess.run(
            [hhc, str(hhp_path)],
            capture_output=True,
            text=True,
            cwd=str(project_dir),
            timeout=60,
        )
    except subprocess.TimeoutExpired:
        return False, "hhc.exe timed out after 60 seconds"
    except Exception as e:
        return False, f"hhc.exe subprocess error: {e}"

    # hhc.exe exits 1 on warnings (very common) and only truly fails on errors.
    # Check for output file existence as the definitive success signal.
    chm_name = hhp_path.stem + ".chm"
    chm_built = project_dir / chm_name

    if not chm_built.exists():
        stderr = result.stderr.strip() or result.stdout.strip()
        return False, f"hhc.exe failed (rc={result.returncode}): {stderr}"

    # Move to caller-specified output path
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(chm_built), str(output_path))

    size = output_path.stat().st_size
    return True, f"CHM compiled successfully: {output_path.name} ({size} bytes)"
