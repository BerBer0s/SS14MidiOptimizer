@echo off
title MIDI Optimizer SS14 - Build
echo ============================================
echo   MIDI Optimizer for SS14 - Build EXE
echo   (PySide6 edition)
echo ============================================
echo.

py -3.11 --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python 3.11 not found.
    echo Download: https://www.python.org/downloads/
    pause & exit /b 1
)

echo [1/3] Installing dependencies...
py -3.11 -m pip install --upgrade pip -q 2>nul
py -3.11 -m pip install PySide6 mido pyinstaller pillow -q
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    pause & exit /b 1
)

if exist build rmdir /s /q build 2>nul
if exist dist rmdir /s /q dist 2>nul
if exist MIDI_Optimizer_SS14.spec del MIDI_Optimizer_SS14.spec 2>nul

echo.
echo [2/3] Building EXE (1-3 minutes)...
echo.

py -3.11 -m PyInstaller --noconfirm --onefile --windowed --name "MIDI_Optimizer_SS14" --icon "icon.ico" --add-data "icon.ico;." --hidden-import "mido" --hidden-import "mido.backends.rtmidi" --hidden-import "PIL" midi_optimizer_gui.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed. See log above.
    pause & exit /b 1
)

echo.
echo [3/3] Done!
echo ============================================
echo   EXE: dist\MIDI_Optimizer_SS14.exe
echo ============================================
echo.
pause
