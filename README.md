# Text to Speech Batch

A Windows desktop GUI application for batch-converting text files (`.txt`) and Word documents (`.docx`) to high-quality MP3 audio files using the Gemini TTS API (`gemini-3.8-flash-tts`).

The deliverable is a standalone `.exe` (`TextToSpeechBatch.exe`) that runs on any Windows 10 or Windows 11 PC without requiring Python or external libraries (such as ffmpeg).

---

## Key Features

- **Single Standalone Executable**: Bundles Python runtime, tkinter GUI, Google GenAI SDK, and the `lameenc` MP3 encoder without relying on external dependencies.
- **Structured Error Classification**:
  - Differentiates `auth` (401, 403 `PERMISSION_DENIED`), `region` (400 `FAILED_PRECONDITION` location unsupported), `model_not_found` (404), `quota_daily` (429 per-day), `rate_limit` (429 per-minute), `bad_request` (400 invalid argument/voice), and `transient` (5xx, timeouts, connection resets).
  - Fatal account/environment errors (`auth`, `region`, `model_not_found`, `quota_daily`) halt the batch cleanly; single-file errors (`bad_request`) fail only the offending file while continuing the rest of the queue.
  - Server errors containing numbers in trace IDs (e.g. `5403211`) are correctly classified as transient and never falsely treated as authentication failures.
- **Adaptive Rate Limiting & Timeout**:
  - Automatically respects server retry hints (e.g. `retry in 23.5s`) or applies backoff (`20s, 40s, 60s, 60s`) with up to 5 attempts.
  - Enforces a minimum 20-second interval between consecutive requests across the queue after any 429 response.
  - Client timeout configured via `types.HttpOptions(timeout=180000)` (3 minutes).
- **Exact Substring Chunking**:
  - Splits documents into chunks of up to 3000 characters without altering numbers (e.g. `46.5`), times (`10:30`), URLs (`https://example.com/a.b`), or newlines.
  - Boundaries: sentence enders (`. ! ? : \u061f \u061b`) followed by whitespace, single line breaks, and paragraph breaks.
  - Retains original whitespace and separators; chunks are exact substrings of the input text.
- **Encoding & Extraction Robustness**:
  - `.txt`: Validates BOM (UTF-8-SIG, UTF-16-LE, UTF-16-BE), strict UTF-8, and strict CP1256 (Windows-1256). Even-length CP1256 files are never erroneously decoded as UTF-16.
  - `.docx`: Natural body traversal, recursive `w:sdt` content control parsing, and merged cell deduplication (`seen_tcs`) ensuring merged cells are read only once.
  - Legacy `.doc` protection: Prompts users to save legacy files as `.docx`.
  - Automatic filtering: Skips Word temporary lock files (`~$*`) and summarizes any unsupported file selections.
- **Chunk Caching & Resume on Stop**:
  - Synthesized chunks are saved to `%TEMP%\TextToSpeechBatch\cache\<hash>\<chunk_index>.pcm`.
  - Stopping and restarting resumes previously completed chunks instantly from cache without re-requesting the API.
  - Injects `0.3s` silence between chunks of the same file.
  - File cache is automatically cleaned upon successful MP3 generation, and orphaned caches older than 7 days are pruned on launch.
- **Responsive In-Place Grid Editing**:
  - Click or double-click the Voice or Style column directly in the file table to change settings per file.
  - Settings panel comboboxes automatically apply to selected files on selection or Enter.
  - Universal clipboard support (right-click context menu and Ctrl+V / Keycode 86 across all keyboard layouts).
- **Built-in CLI Self-Test**:
  - Run `TextToSpeechBatch.exe --selftest` to verify DPAPI encryption, LAME encoder, docx parser, chunking, and error classification in headless mode.

---

## Voice and Style Presets

- **33 Gemini TTS Voices**: Bodi, Fola, Sola, Zephyr, Puck, Charon, Kore, Fenrir, Leda, Orus, Aoede, Callirrhoe, Autonoe, Enceladus, Iapetus, Umbriel, Algieba, Despina, Erinome, Algenib, Rasalgethi, Laomedeia, Achernar, Alnilam, Schedar, Gacrux, Pulcherrima, Achird, Zubenelgenubi, Vindemiatrix, Sadachbia, Sadaltager, Sulafat.
- **12 Style Presets**: Cheerful, Sad, Calm, Excited, Serious/Formal, Warm/Friendly, Narration, News Anchor, Whisper, Promotional, Angry, and Neutral (no instruction).
- **Free-Text Style**: Enter any custom prompt in the Style field.

---

## Installation & Development

### Requirements
- Windows 10 or 11 (64-bit)
- Python 3.11+ (for development)

### Quick Start
```cmd
pip install -r requirements.txt
python tts_gui.py
```

### Run Unit Tests
```cmd
python -m unittest test_tts.py -v
```

### Run Headless Self-Test
```cmd
python tts_gui.py --selftest
```

---

## Building the Executable

Run the automated build script:
```cmd
build.bat
```
This script will:
1. Install pinned requirements.
2. Run the automated unit test suite.
3. Package `dist\TextToSpeechBatch.exe` with PyInstaller (`--onefile --windowed`).
4. Execute `dist\TextToSpeechBatch.exe --selftest` to verify the standalone binary.

---

## CI/CD Workflow

A GitHub Actions workflow is provided at `.github/workflows/build.yml`.
- Runs tests and builds `TextToSpeechBatch.exe` on every push to `main` / `master` and pull requests.
- Uploads the executable as a workflow artifact (`TextToSpeechBatch-Windows`).
- Publishes an official GitHub Release whenever a tag matching `v*` is pushed.
