@echo off
echo ===================================
echo Outlook MCP Server - Setup Script
echo ===================================
echo.

REM Try to find Python 3.11 or 3.12
set PYTHON_CMD=

REM Check for py launcher with 3.11
py -3.11 --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=py -3.11
    goto :found
)

REM Check for py launcher with 3.12
py -3.12 --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=py -3.12
    goto :found
)

REM Check for python311 in PATH
python3.11 --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python3.11
    goto :found
)

REM Check for python312 in PATH
python3.12 --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python3.12
    goto :found
)

echo ERROR: Python 3.11 or 3.12 not found!
echo Please install Python 3.11 or 3.12 from https://www.python.org/downloads/
echo Make sure to check "Add Python to PATH" during installation.
pause
exit /b 1

:found
echo Found Python: %PYTHON_CMD%
%PYTHON_CMD% --version
echo.

REM Remove old venv if exists
if exist .venv (
    echo Removing old virtual environment...
    rmdir /s /q .venv
)

REM Create new venv
echo Creating virtual environment...
%PYTHON_CMD% -m venv .venv
if %errorlevel% neq 0 (
    echo ERROR: Failed to create virtual environment
    pause
    exit /b 1
)

REM Activate and install
echo.
echo Installing dependencies (this may take a few minutes)...
call .venv\Scripts\activate.bat

REM Upgrade pip first
python -m pip install --upgrade pip

REM Install the package
pip install -e .

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Installation failed!
    echo Try running: pip install -e . --verbose
    pause
    exit /b 1
)

echo.
echo ===================================
echo Installation complete!
echo ===================================
echo.
echo To use the Outlook MCP server:
echo 1. Restart Claude Code (if running)
echo 2. The "outlook" MCP server should now be available
echo.
echo First-time setup:
echo - Open Outlook before using email search
echo - Run "email_index_rebuild" to index your emails
echo.
pause
