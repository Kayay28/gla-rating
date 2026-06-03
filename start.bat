@echo off
title GLA Rating System - Word Server
echo ================================================
echo   GLA Teacher Performance Rating System
echo   Starting Word Generation Server...
echo ================================================
echo.

cd /d "%~dp0"

echo Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python from python.org
    pause
    exit
)

echo Installing requirements...
pip install flask flask-cors python-docx --quiet --break-system-packages 2>nul
pip install flask flask-cors python-docx --quiet 2>nul

echo.
echo Starting server and opening browser...
python server.py

pause
