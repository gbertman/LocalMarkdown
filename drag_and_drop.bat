@echo off
REM ============================================================================
REM  LocalMarkdown — drag-and-drop helper (Windows)
REM
REM  Put a SHORTCUT to this .bat on your Desktop, then drag a file or folder
REM  onto it. Every dropped path is converted to Markdown in markdown_output\.
REM
REM  You can drop multiple items at once. The window stays open on error so you
REM  can read any messages.
REM ============================================================================

setlocal
set "SCRIPT_DIR=%~dp0"

REM ---- Use the project's virtual environment if one exists -------------------
if exist "%SCRIPT_DIR%.venv\Scripts\python.exe" (
    set "PY=%SCRIPT_DIR%.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

if "%~1"=="" (
    echo Drag one or more files or folders onto this shortcut.
    pause
    exit /b 1
)

REM ---- Pass every dropped path to the `process` sub-command ------------------
"%PY%" "%SCRIPT_DIR%markdown_mcp_server.py" process %*

if errorlevel 1 (
    echo.
    echo Something went wrong. See messages above.
    pause
)
endlocal
