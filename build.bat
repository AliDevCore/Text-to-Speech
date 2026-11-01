@echo off
setlocal enabledelayedexpansion

echo =========================================================
echo  Building Text to Speech Batch Standalone Executable
echo =========================================================

:: Check if Python is installed
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python was not found in PATH. Please install Python 3.11+ and try again.
    pause
    exit /b 1
)

echo [1/4] Checking and installing dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to install required dependencies.
    pause
    exit /b 1
)

echo [2/4] Running unit test suite...
python -m unittest test_tts.py -v
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Unit tests failed. Aborting build.
    pause
    exit /b 1
)

echo [3/4] Packaging directory build (lightweight exe with _internal folder)...
python -m PyInstaller --onedir --windowed --name TextToSpeechBatch --noconfirm --collect-all google.genai --collect-all lameenc --collect-all docx --clean tts_gui.py
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] PyInstaller packaging failed.
    pause
    exit /b 1
)

echo [4/4] Verifying output executable...
if exist "dist\TextToSpeechBatch\TextToSpeechBatch.exe" (
    echo =========================================================
    echo [SUCCESS] Build succeeded!
    echo Application executable created at:
    echo %CD%\dist\TextToSpeechBatch\TextToSpeechBatch.exe
    echo =========================================================
) else (
    echo [ERROR] dist\TextToSpeechBatch\TextToSpeechBatch.exe was not found.
    pause
    exit /b 1
)

endlocal
