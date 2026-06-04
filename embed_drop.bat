@echo off
REM ============================================================================
REM  embed_drop.bat — drag-and-drop embedding into a local AnythingLLM
REM
REM  Put a SHORTCUT to this .bat on your Desktop, then drag .md / .markdown / .txt
REM  files or folders onto it. It lists your AnythingLLM workspaces, you pick one
REM  (by number, or type a new name), and the dropped files are embedded there.
REM
REM  One-time setup (so you don't have to type the key each time) — in a terminal:
REM      setx ALLM_KEY  "<your AnythingLLM Developer API key>"
REM      setx ALLM_URL  "http://localhost:3001"
REM  (Open a NEW terminal/Explorer after setx for it to take effect.)
REM ============================================================================

setlocal
set "SCRIPT_DIR=%~dp0"

REM ---- Use the project's virtual environment if one exists -------------------
if exist "%SCRIPT_DIR%.venv\Scripts\python.exe" (
    set "PY=%SCRIPT_DIR%.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

if "%ALLM_URL%"=="" set "ALLM_URL=http://localhost:3001"

if "%ALLM_KEY%"=="" (
    echo [!] ALLM_KEY is not set. Run:  setx ALLM_KEY "your-api-key"
    echo     ^(AnythingLLM: Settings -^> Tools -^> Developer API -^> Generate^)
    echo.
)

REM ---- No files dropped: just show the available workspaces ------------------
if "%~1"=="" (
    echo Drag .md / .markdown / .txt files or folders onto this shortcut to embed them.
    echo.
    echo Current AnythingLLM workspaces:
    "%PY%" "%SCRIPT_DIR%allm_sync.py" --list-workspaces
    echo.
    pause
    exit /b 0
)

REM ---- Embed the dropped paths (interactive workspace picker) ----------------
"%PY%" "%SCRIPT_DIR%allm_sync.py" %*

echo.
pause
endlocal
