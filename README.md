# Text to Speech Batch

A Windows desktop GUI application for batch-converting text files (`.txt`) and Word documents (`.docx`) to high-quality MP3 audio files using the Gemini TTS API (`gemini-3.8-flash-tts`).

The deliverable is a single standalone `.exe` that runs out of the box on Windows 10 and Windows 11 with zero Python or library dependencies required on the target machine.

---

## Features

- **Standalone Single File Executable**: Bundles Python runtime, GUI, Google GenAI SDK, and the `lameenc` MP3 encoder without relying on external tools like ffmpeg.
- **Input Formats**:
  - Text files (`.txt`): Automatically supports multiple encodings (`utf-8-sig`, `utf-8`, `utf-16`, `cp1256`).
  - Word documents (`.docx`): Accurately iterates document body in natural reading order, extracting paragraphs and table cells while preserving Zero-Width Non-Joiner (ZWNJ) characters.
  - Legacy `.doc` protection: Detects legacy `.doc` files and instructs users to save them as `.docx`.
  - Automatic filtering: Ignores temporary Word lock files (`~$*`).
- **Flexible Voice & Style Presets**:
  - 33 high-fidelity Gemini TTS voices available via a fixed dropdown.
  - 12 built-in style presets for tone modulation (cheerful, sad, calm, excited, formal, narration, whisper, promotional, etc.), with full support for custom free-text style instructions.
- **Security & Key Persistence**:
  - API key is protected using the Windows Data Protection API (DPAPI via `CryptProtectData`/`CryptUnprotectData`) and saved locally at `%APPDATA%\TextToSpeechBatch\config.json`.
- **Intelligent Chunking & Queue Management**:
  - Automatically splits large documents into chunks of up to 3000 characters by blank-line paragraphs and sentence boundaries.
  - Output name collision detection: Blocks processing and warns if two queued files would produce the same `.mp3` destination.
  - Atomic writes: Audio is written to a temporary file first and atomically renamed to prevent corrupt or incomplete MP3s.
- **Resilient Background Execution**:
  - Asynchronous background processing with full thread-safe Tkinter UI updates via `queue.Queue`.
  - Immediate responsive stop button.
  - Automatic retry with exponential backoff on transient errors and immediate fail-fast reporting on invalid API keys.
- **Offline Mock Testing Mode**:
  - Set environment variable `TTS_MOCK=1` to simulate conversions and test the entire pipeline without network access or API key.

---

## Requirements

- **Runtime**: Windows 10 or 11 (64-bit). No installation needed when using `TextToSpeechBatch.exe`.
- **Development & Building**:
  - Python 3.11+ (64-bit)
  - `requirements.txt` packages (`google-genai`, `lameenc`, `python-docx`, `pyinstaller`)

---

## Quick Start (Running from Source)

1. Clone or open the project folder.
2. Install the required dependencies:
   ```cmd
   pip install -r requirements.txt
   ```
3. Run the application:
   ```cmd
   python tts_gui.py
   ```

---

## Building the Standalone Executable

### Option A: Using the automated build script

Run `build.bat` in the project root:
```cmd
build.bat
```
This script will verify your Python environment, run the automated test suite, build the single-file executable using PyInstaller, and verify the resulting `dist\TextToSpeechBatch.exe`.

### Option B: Running PyInstaller manually

Execute the following command in Command Prompt or PowerShell:
```cmd
pyinstaller --onefile --windowed --name TextToSpeechBatch --collect-all google.genai --collect-all lameenc --collect-all docx --clean tts_gui.py
```
The output executable will be placed in the `dist` folder:
```
dist\TextToSpeechBatch.exe
```

---

## Automated CI/CD Workflow

A GitHub Actions workflow is provided at `.github/workflows/build.yml`. On every push or pull request to `main` or `master`, the workflow:
1. Provisions a clean `windows-latest` runner.
2. Installs Python 3.11 and dependencies.
3. Executes the full unit test suite (`test_tts.py`).
4. Builds `TextToSpeechBatch.exe` with PyInstaller.
5. Uploads the executable as a downloadable workflow artifact (`TextToSpeechBatch-Windows`).

---

## Running Automated Tests

To run the unit test suite covering chunking, output path collisions, docx extraction, prompt formatting, and offline mock pipeline:
```cmd
python -m unittest test_tts.py -v
```

To run the GUI in offline mock mode (no network, 1 second of silence per chunk):
```cmd
set TTS_MOCK=1
python tts_gui.py
```

