"""
Erebus.Helper — Windows-side Electron fake-installer container build.

Invoked by the deferred build runbook (build_electron.bat) produced by the
Linux builder. Runs npm install, then electron-builder to produce the
Windows NSIS installer, and copies the result to the requested output path.

The Electron project is now a pure JS wrapper (no native addon): it ships
with a staged `build/resources/payload/` tree that the wizard copies to
`%TEMP%\\inst-<uuid>` and spawns at install time. All cross-compile /
node-gyp / electron-rebuild complexity has been removed.

Node.js 20 LTS and an internet-reachable npm registry are required on the
host running this.
"""

import logging
import os
import shutil
import subprocess
from pathlib import Path
from typing import Optional, Tuple

logger = logging.getLogger(__name__)


def _run(cmd: list, cwd: Path, env: Optional[dict] = None) -> Tuple[int, str, str]:
    logger.info(f"$ {' '.join(cmd)}  (cwd={cwd})")
    proc = subprocess.run(
        cmd,
        cwd=str(cwd),
        env=env,
        capture_output=True,
        text=True,
        shell=False,
    )
    if proc.stdout:
        logger.debug(proc.stdout)
    if proc.stderr:
        logger.debug(proc.stderr)
    return proc.returncode, proc.stdout or "", proc.stderr or ""


def _find_nsis_output(dist_dir: Path) -> Optional[Path]:
    """Return the freshest *.exe in dist/ (prefer filenames containing 'Setup')."""
    if not dist_dir.exists():
        return None
    candidates = sorted(
        (p for p in dist_dir.glob("*.exe") if "setup" in p.name.lower()),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        candidates = sorted(
            dist_dir.glob("*.exe"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
    return candidates[0] if candidates else None


def compile_electron(
    project_dir: str,
    output_path: str,
    arch: str = "x64",
    skip_install: bool = False,
    sign_cert: Optional[str] = None,
    sign_password: Optional[str] = None,
) -> Tuple[bool, str]:
    """
    Build the staged Erebus.Electron project into a Windows NSIS installer.

    Steps:
      1. `npm install`                       (unless --skip-install)
      2. `npx @electron/rebuild -f -w erebus_rt --arch <arch>`
      3. `npx electron-builder --win nsis --<arch> --publish never`
      4. Copy dist/*Setup.exe to output_path

    Args:
        project_dir: Path to the staged Erebus.Electron project.
        output_path: Destination for the NSIS installer exe.
        arch: "x64" or "ia32".
        skip_install: Skip step 1 (node_modules already present).
        sign_cert: Optional .pfx cert path (sets CSC_LINK).
        sign_password: Optional cert password (sets CSC_KEY_PASSWORD).

    Returns:
        (success, message)
    """
    proj = Path(project_dir).resolve()
    out = Path(output_path).resolve()

    if not proj.exists() or not proj.is_dir():
        return False, f"project-dir does not exist: {proj}"
    if not (proj / "package.json").exists():
        return False, f"package.json not found in {proj}"

    # Sanity-check: builder must have rendered src/config.js and staged
    # the loader payload under build/resources/payload/ (picked up by
    # electron-builder's extraResources directive).
    required = ("src/config.js", "build/resources/payload")
    missing = [p for p in required if not (proj / p).exists()]
    if missing:
        return False, f"missing prerequisites: {', '.join(missing)}"

    if shutil.which("node") is None:
        return False, "node not found on PATH (requires Node.js 20 LTS)"
    if shutil.which("npm") is None:
        return False, "npm not found on PATH"

    env = os.environ.copy()
    if sign_cert:
        env["CSC_LINK"] = str(Path(sign_cert).resolve())
        if sign_password:
            env["CSC_KEY_PASSWORD"] = sign_password
    else:
        # Skip signing — electron-builder would otherwise probe the Windows cert store.
        env["CSC_IDENTITY_AUTO_DISCOVERY"] = "false"

    # Step 1: npm install
    if not skip_install:
        rc, _out, err = _run(["npm", "install", "--no-audit", "--no-fund", "--loglevel=error"], cwd=proj, env=env)
        if rc != 0:
            return False, f"npm install failed ({rc}): {err.strip() or _out.strip()}"

    # Step 2: electron-builder produces the NSIS installer.
    # No native addon rebuild needed — the Electron container is pure JS
    # (fake-installer wizard that extract-and-spawns the embedded loader).
    arch_flag = f"--{arch}" if arch in ("x64", "ia32") else "--x64"
    rc, _out, err = _run(
        ["npx", "--no-install", "electron-builder", "--win", "nsis", arch_flag, "--publish", "never"],
        cwd=proj,
        env=env,
    )
    if rc != 0:
        return False, f"electron-builder failed ({rc}): {err.strip() or _out.strip()}"

    # Step 4: locate the NSIS output and copy to the caller's output path
    nsis = _find_nsis_output(proj / "dist")
    if nsis is None:
        return False, f"electron-builder reported success but no *.exe found in {proj / 'dist'}"

    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(str(nsis), str(out))
    logger.info(f"Electron installer built: {out} ({out.stat().st_size} bytes)")
    return True, f"built {out.name} from {nsis.name}"
