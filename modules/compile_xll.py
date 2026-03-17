"""
XLL Compilation Module for Erebus Helper
Compiles C/C++ source files into Excel Add-In (.xll) DLLs on a Windows host.

Supported tool-chains
---------------------
* MSVC  - cl.exe located via standard Visual Studio installation layout
* MinGW - gcc found either in well-known directories or on PATH
* Clang - clang found on PATH (cross-compiling to Windows)

The module exposes a single high-level entry-point:

    result = compile_xll(source_file, output_path, ...)

and a lower-level ``XllCompiler`` class for callers that need finer control.
"""

import logging
import os
import pathlib
import shlex
import shutil
import subprocess
import sys
import tempfile
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger("ErebusHelper.XLL")

# ---------------------------------------------------------------------------
# Well-known compiler search paths
# ---------------------------------------------------------------------------

_MSVC_BASE_PATHS: List[str] = [
    r"C:\Program Files\Microsoft Visual Studio\18\Community\VC\Tools\MSVC",
    r"C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Tools\MSVC",
    r"C:\Program Files\Microsoft Visual Studio\2022\Professional\VC\Tools\MSVC",
    r"C:\Program Files\Microsoft Visual Studio\2022\Enterprise\VC\Tools\MSVC",
    r"C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Tools\MSVC",
    r"C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\VC\Tools\MSVC",
    r"C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\VC\Tools\MSVC",
    r"C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\VC\Tools\MSVC",
    r"C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\VC\Tools\MSVC",
    r"C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\VC\Tools\MSVC",
]

# Search roots for vcvarsall.bat (VS installer layout)
_VS_VCVARSALL_PATHS: List[str] = [
    r"C:\Program Files\Microsoft Visual Studio\18\Community\VC\Auxiliary\Build\vcvarsall.bat",
    r"C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat",
    r"C:\Program Files\Microsoft Visual Studio\2022\Professional\VC\Auxiliary\Build\vcvarsall.bat",
    r"C:\Program Files\Microsoft Visual Studio\2022\Enterprise\VC\Auxiliary\Build\vcvarsall.bat",
    r"C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvarsall.bat",
    r"C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\VC\Auxiliary\Build\vcvarsall.bat",
    r"C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\VC\Auxiliary\Build\vcvarsall.bat",
    r"C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\VC\Auxiliary\Build\vcvarsall.bat",
    r"C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\VC\Auxiliary\Build\vcvarsall.bat",
    r"C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\VC\Auxiliary\Build\vcvarsall.bat",
]

_MINGW_BASE_PATHS: List[str] = [
    r"C:\mingw64",
    r"C:\mingw32",
    r"C:\msys64\mingw64",
    r"C:\msys64\mingw32",
    r"C:\tools\mingw64",
]

# MSVC flags that produce a Windows DLL suitable for use as an XLL
_MSVC_XLL_FLAGS: List[str] = [
    "/D_WINDOWS", "/DWIN32", "/D_USRDLL", "/D_WINDLL",
    "/W3", "/nologo", "/EHsc", "/LD",
]

# MinGW / Clang equivalent flags
_GCC_XLL_FLAGS: List[str] = [
    "-shared", "-fPIC", "-DWIN32", "-D_WINDOWS", "-D_USRDLL", "-Wall",
]


# ============================================================================
# Compiler discovery helpers
# ============================================================================

def _find_vcvarsall() -> Optional[str]:
    """Return the path to vcvarsall.bat or None if not found."""
    for path in _VS_VCVARSALL_PATHS:
        if pathlib.Path(path).exists():
            return path
    return None


def _msvc_environment(architecture: str = "x64") -> Optional[dict]:
    """
    Run vcvarsall.bat and capture the resulting environment variables.

    Returns a dict suitable for passing as ``env=`` to subprocess, or
    ``None`` if vcvarsall.bat cannot be found or fails.
    """
    vcvarsall = _find_vcvarsall()
    if not vcvarsall:
        logger.warning("vcvarsall.bat not found - cl.exe may fail without SDK headers")
        return None

    arch_arg = "x64" if architecture == "x64" else "x86"
    # Run vcvarsall then print all env vars with a known sentinel
    cmd = f'cmd.exe /c ""{vcvarsall}" {arch_arg} && set"'
    try:
        result = subprocess.run(
            cmd,
            shell=True,
            capture_output=True,
            text=True,
            timeout=60,
        )
    except Exception as exc:
        logger.warning(f"Failed to run vcvarsall.bat: {exc}")
        return None

    if result.returncode != 0:
        logger.warning(f"vcvarsall.bat exited {result.returncode}: {result.stderr.strip()}")
        return None

    env: dict = {}
    for line in result.stdout.splitlines():
        if "=" in line:
            key, _, value = line.partition("=")
            env[key] = value
    logger.debug(f"Captured MSVC environment ({len(env)} vars) from {vcvarsall}")
    return env


def _find_cl_exe(architecture: str = "x64") -> Optional[str]:
    """
    Locate ``cl.exe`` inside the Visual Studio MSVC directory tree.

    Searches the known VS installation paths and falls back to ``PATH``
    if none of those directories are present on the current machine.
    """
    arch_folder = "x64" if architecture == "x64" else "x86"

    for base in _MSVC_BASE_PATHS:
        base_p = pathlib.Path(base)
        if not base_p.exists():
            continue
        # Each MSVC base dir contains one or more versioned sub-directories;
        # pick the latest (lexicographically greatest).
        host_dirs = sorted(base_p.glob("*/bin/Host*"), reverse=True)
        for host_dir in host_dirs:
            candidate = host_dir / arch_folder / "cl.exe"
            if candidate.exists():
                logger.debug(f"Found cl.exe at {candidate}")
                return str(candidate)

    # Last resort - hope it is on PATH
    found = shutil.which("cl.exe")
    if found:
        logger.debug(f"Found cl.exe on PATH: {found}")
    return found


def _find_gcc() -> Optional[str]:
    """Return path to the MinGW gcc executable or ``None``."""
    for base in _MINGW_BASE_PATHS:
        candidate = pathlib.Path(base) / "bin" / "gcc.exe"
        if candidate.exists():
            return str(candidate)
    return shutil.which("gcc") or shutil.which("gcc.exe")


def _find_clang() -> Optional[str]:
    """Return path to the clang executable or ``None``."""
    return shutil.which("clang") or shutil.which("clang.exe")


# ============================================================================
# XllCompiler class
# ============================================================================

class XllCompiler:
    """
    Compiles a C/C++ source file into an XLL (Excel Add-In DLL).

    Parameters
    ----------
    compiler : str
        Tool-chain name: ``"MSVC"``, ``"MinGW"``, or ``"Clang"``.
    architecture : str
        Target architecture: ``"x64"`` or ``"x86"``.
    optimization : str
        Optimisation level: ``"Ox"``, ``"O2"``, ``"O1"``, or ``"Od"``
        (MSVC style - the module converts to GCC style automatically).
    extra_flags : list[str], optional
        Additional raw compiler flags appended to the default set.
    timeout : int
        Maximum number of seconds to wait for the compiler process.
    """

    def __init__(
        self,
        compiler: str = "MSVC",
        architecture: str = "x64",
        optimization: str = "Ox",
        extra_flags: Optional[List[str]] = None,
        timeout: int = 300,
    ):
        self.compiler = compiler.upper()
        self.architecture = architecture
        self.optimization = optimization
        self.extra_flags: List[str] = extra_flags or []
        self.timeout = timeout

        if self.compiler not in ("MSVC", "MINGW", "CLANG"):
            raise ValueError(f"Unsupported compiler: {compiler!r}")

    # ------------------------------------------------------------------
    # Public entry-point
    # ------------------------------------------------------------------

    def compile(
        self,
        source_file: str,
        output_path: str,
        extra_libs: Optional[List[str]] = None,
        defines: Optional[Dict[str, Optional[str]]] = None,
        include_dirs: Optional[List[str]] = None,
    ) -> Tuple[bool, str]:
        """
        Compile *source_file* into an XLL at *output_path*.

        Parameters
        ----------
        source_file : str
            Path to the ``.c`` or ``.cpp`` translation unit.
        output_path : str
            Desired output path.  The ``.xll`` extension is enforced
            automatically if the caller omits it.
        extra_libs : list[str], optional
            Additional ``-l<lib>`` / ``<lib>.lib`` link inputs.
        defines : dict, optional
            Preprocessor defines.  A value of ``None`` means a bare
            ``-DNAME`` (MSVC: ``/DNAME``); a string value produces
            ``-DNAME=VALUE``.
        include_dirs : list[str], optional
            Extra ``-I`` / ``/I`` include search paths.

        Returns
        -------
        (success: bool, message: str)
            ``True`` + empty string on success; ``False`` + error description
            on failure.
        """
        src = pathlib.Path(source_file)
        if not src.exists():
            return False, f"Source file not found: {source_file}"

        # Enforce .xll extension
        out = pathlib.Path(output_path)
        if out.suffix.lower() != ".xll":
            out = out.with_suffix(".xll")
        out.parent.mkdir(parents=True, exist_ok=True)

        libs: List[str] = list(extra_libs or [])
        defs: Dict[str, Optional[str]] = defines or {}
        incs: List[str] = list(include_dirs or [])

        logger.info(
            f"Compiling XLL: {src.name} → {out.name}  "
            f"[{self.compiler} / {self.architecture} / /{self.optimization}]"
        )

        dispatch = {
            "MSVC":  self._compile_msvc,
            "MINGW": self._compile_mingw,
            "CLANG": self._compile_clang,
        }
        return dispatch[self.compiler](str(src), str(out), libs, defs, incs)

    # ------------------------------------------------------------------
    # MSVC back-end
    # ------------------------------------------------------------------

    def _compile_msvc(
        self,
        source: str,
        output: str,
        libs: List[str],
        defines: Dict[str, Optional[str]],
        include_dirs: List[str],
    ) -> Tuple[bool, str]:
        cl = _find_cl_exe(self.architecture)
        if not cl:
            return False, "cl.exe not found - install Visual C++ Build Tools"

        logger.info(f"Using cl.exe: {cl}")

        msvc_env = _msvc_environment(self.architecture)

        cmd: List[str] = [cl] + _MSVC_XLL_FLAGS + [f"/{self.optimization}"]

        for name, value in defines.items():
            cmd.append(f"/D{name}" if value is None else f"/D{name}={value}")
        for inc in include_dirs:
            cmd.append(f"/I{inc}")

        cmd += [f"/Fe{output}", source] + libs + self.extra_flags

        return self._run(cmd, output, env=msvc_env)

    # ------------------------------------------------------------------
    # MinGW back-end
    # ------------------------------------------------------------------

    def _compile_mingw(
        self,
        source: str,
        output: str,
        libs: List[str],
        defines: Dict[str, Optional[str]],
        include_dirs: List[str],
    ) -> Tuple[bool, str]:
        gcc = _find_gcc()
        if not gcc:
            return False, "gcc not found - install MinGW-w64"

        logger.info(f"Using gcc: {gcc}")

        arch_flag = "-m64" if self.architecture == "x64" else "-m32"
        opt_flag = f"-{self.optimization}"

        cmd: List[str] = [gcc] + _GCC_XLL_FLAGS + [arch_flag, opt_flag]

        for name, value in defines.items():
            cmd.append(f"-D{name}" if value is None else f"-D{name}={value}")
        for inc in include_dirs:
            cmd.append(f"-I{inc}")

        cmd += ["-o", output, source] + libs + self.extra_flags

        return self._run(cmd, output)

    # ------------------------------------------------------------------
    # Clang back-end
    # ------------------------------------------------------------------

    def _compile_clang(
        self,
        source: str,
        output: str,
        libs: List[str],
        defines: Dict[str, Optional[str]],
        include_dirs: List[str],
    ) -> Tuple[bool, str]:
        clang = _find_clang()
        if not clang:
            return False, "clang not found - install LLVM"

        logger.info(f"Using clang: {clang}")

        if self.architecture == "x64":
            target_flags = ["-m64", "-target", "x86_64-pc-windows-msvc"]
        else:
            target_flags = ["-m32", "-target", "i686-pc-windows-msvc"]

        opt_flag = f"-{self.optimization}"

        cmd: List[str] = [clang] + _GCC_XLL_FLAGS + target_flags + [opt_flag]

        for name, value in defines.items():
            cmd.append(f"-D{name}" if value is None else f"-D{name}={value}")
        for inc in include_dirs:
            cmd.append(f"-I{inc}")

        cmd += ["-o", output, source] + libs + self.extra_flags

        return self._run(cmd, output)

    # ------------------------------------------------------------------
    # Shared subprocess runner
    # ------------------------------------------------------------------

    def _run(self, cmd: List[str], output_path: str, env: Optional[dict] = None) -> Tuple[bool, str]:
        logger.debug("Running: " + " ".join(cmd))
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=self.timeout,
                env=env,
            )
        except subprocess.TimeoutExpired:
            return False, f"Compiler timed out after {self.timeout}s"
        except FileNotFoundError as exc:
            return False, f"Compiler binary not found: {exc}"

        if result.returncode != 0:
            combined = (result.stderr or "") + (result.stdout or "")
            logger.error(f"Compilation failed (exit {result.returncode}):\n{combined}")
            return False, combined.strip()

        if not pathlib.Path(output_path).exists():
            return False, "Compiler exited cleanly but output file was not produced"

        size = pathlib.Path(output_path).stat().st_size
        logger.info(f"XLL created: {output_path} ({size:,} bytes)")
        return True, ""

    # ------------------------------------------------------------------
    # PE / MZ header verification
    # ------------------------------------------------------------------

    @staticmethod
    def verify_xll(path: str) -> Tuple[bool, str]:
        """
        Quick sanity-check that *path* is a valid PE (MZ) file.

        Returns ``(True, "")`` on success, ``(False, reason)`` on failure.
        """
        p = pathlib.Path(path)
        if not p.exists():
            return False, "File does not exist"
        if p.stat().st_size < 512:
            return False, "File too small to be a valid PE"
        try:
            with open(path, "rb") as fh:
                if fh.read(2) != b"MZ":
                    return False, "Missing MZ signature - not a valid PE"
        except OSError as exc:
            return False, f"Cannot read file: {exc}"
        return True, ""


# ============================================================================
# Module-level convenience wrapper
# ============================================================================

def compile_xll(
    source_file: str,
    output_path: str,
    compiler: str = "MSVC",
    architecture: str = "x64",
    optimization: str = "Ox",
    extra_libs: Optional[List[str]] = None,
    defines: Optional[Dict[str, Optional[str]]] = None,
    include_dirs: Optional[List[str]] = None,
    extra_flags: Optional[List[str]] = None,
    timeout: int = 300,
    verify: bool = True,
) -> Tuple[bool, str]:
    """
    Compile a C/C++ source file into a Windows XLL (Excel Add-In DLL).

    This is the primary entry-point for callers that do not need to
    instantiate ``XllCompiler`` directly.

    Parameters
    ----------
    source_file : str
        Path to the ``.c`` / ``.cpp`` source.
    output_path : str
        Desired output path (``".xll"`` extension added if missing).
    compiler : str
        ``"MSVC"``, ``"MinGW"``, or ``"Clang"``.
    architecture : str
        ``"x64"`` or ``"x86"``.
    optimization : str
        MSVC-style optimisation level (``"Ox"``, ``"O2"``, ``"O1"``, ``"Od"``).
    extra_libs : list[str], optional
        Additional link inputs.
    defines : dict, optional
        Preprocessor defines.
    include_dirs : list[str], optional
        Extra include search paths.
    extra_flags : list[str], optional
        Raw compiler flags appended unconditionally.
    timeout : int
        Compiler process timeout in seconds.
    verify : bool
        When ``True``, perform a PE header sanity check after compilation.

    Returns
    -------
    (success: bool, message: str)
    """
    xc = XllCompiler(
        compiler=compiler,
        architecture=architecture,
        optimization=optimization,
        extra_flags=extra_flags,
        timeout=timeout,
    )
    ok, msg = xc.compile(
        source_file=source_file,
        output_path=output_path,
        extra_libs=extra_libs,
        defines=defines,
        include_dirs=include_dirs,
    )
    if not ok:
        return False, msg

    if verify:
        # output_path may have had .xll appended
        out = pathlib.Path(output_path)
        if out.suffix.lower() != ".xll":
            out = out.with_suffix(".xll")
        ok, msg = XllCompiler.verify_xll(str(out))
        if not ok:
            return False, f"Post-compile verification failed: {msg}"

    return True, ""
