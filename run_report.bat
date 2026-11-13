@echo off
rem ---------------------------------------------------------------------------
rem  Monthly RPP / DuitNow transaction volume report - Windows launcher
rem
rem  Works on machines that have "py" (Python Launcher) or "python" on PATH.
rem
rem  Usage:
rem    Double-click this file  -> asks for the 01_Results folder,
rem                               saves the report on your Desktop.
rem    Or from a terminal:
rem      run_report.bat "C:\...\Nov 2026\01_Results" -o "C:\Reports"
rem    Any option of the tool can be added (e.g. --month OCT_2026, --no-hints).
rem ---------------------------------------------------------------------------
setlocal
cd /d "%~dp0"

rem 1. Find a working Python. "python" may be the Microsoft Store placeholder,
rem    so each candidate is actually run before it is used.
set "PY="
py --version >nul 2>&1
if not errorlevel 1 set "PY=py"
if not defined PY (
    python --version >nul 2>&1
    if not errorlevel 1 set "PY=python"
)
if not defined PY (
    echo Python was not found.
    echo Install it from https://www.python.org/downloads/ and run this file again.
    goto :end_error
)
echo Using Python command: %PY%
%PY% --version

rem 2. Install the libraries the first time only.
%PY% -c "import pandas, openpyxl, xlrd" >nul 2>&1
if errorlevel 1 (
    echo Installing required libraries...
    %PY% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Installing the libraries failed. See the messages above.
        goto :end_error
    )
)

rem 3. Run the report.
if "%~1"=="" goto :ask
%PY% -m txn_report %*
goto :end

:ask
echo.
set "RESULTS="
set /p "RESULTS=Paste the full path of the 01_Results folder and press Enter: "
if not defined RESULTS goto :end_error
set "RESULTS=%RESULTS:"=%"
%PY% -m txn_report "%RESULTS%" -o "%USERPROFILE%\Desktop"
set "RC=%errorlevel%"
echo.
pause
exit /b %RC%

:end_error
echo.
pause
exit /b 1

:end
exit /b %errorlevel%
