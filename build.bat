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

echo [1/5] Checking and installing pinned dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to install required dependencies.
    pause
    exit /b 1
)

echo [2/5] Running automated unit test suite...
python -m unittest test_tts.py -v
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Unit tests failed. Aborting build.
    pause
    exit /b 1
)

echo [3/5] Packaging standalone single-file executable with PyInstaller...
python -m PyInstaller --onefile --windowed --name TextToSpeechBatch --noconfirm --collect-all google.genai --collect-all lameenc --collect-all docx --clean tts_gui.py
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] PyInstaller packaging failed.
    pause
    exit /b 1
)

echo [4/5] Verifying output executable exists...
if not exist "dist\TextToSpeechBatch.exe" (
    echo [ERROR] dist\TextToSpeechBatch.exe was not found.
    pause
    exit /b 1
)

echo [5/5] Executing self-test verification on built executable...
dist\TextToSpeechBatch.exe --selftest
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Standalone executable self-test failed! Check %%TEMP%%\tts_selftest.log for details.
    pause
    exit /b 1
)

echo =========================================================
echo [SUCCESS] Build and verification succeeded!
echo Standalone executable created at:
echo %CD%\dist\TextToSpeechBatch.exe
echo =========================================================

endlocal
