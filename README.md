# Text to Speech Batch

A Windows desktop GUI application for batch-converting text files (`.txt`) and Word documents (`.docx`) to high-quality MP3 audio files using the Gemini TTS API (`gemini-3.8-flash-tts` and `gemini-3.8-flash-lite-tts`).

The deliverable is a standalone `.exe` (`TextToSpeechBatch.exe`) that runs on any Windows 10 or Windows 11 PC without requiring Python or external libraries (such as ffmpeg).

---

## Key Features

- **Single Standalone Executable & Directory Distribution**:
  - Bundles Python runtime, tkinter GUI, Google GenAI SDK, and the `lameenc` MP3 encoder without relying on external dependencies.
  - GitHub Actions automatically packages a portable standalone single-file `.exe` release, while local builds provide a fast-launching directory distribution (`dist\TextToSpeechBatch`).
- **Multi-Key API Pool (Up to 10 Keys) & Round-Robin Rotation**:
  - Define up to 10 Gemini API keys in the "Manage API Pool..." modal window.
  - Chunks rotate across all configured keys in round-robin order to balance quota and rate limits.
  - Convenient bulk management: individual Paste/Clear buttons, show/mask toggle, and a "Paste Multi-line (Auto-fill)" button to paste an entire list of keys from the clipboard at once.
  - All keys are securely encrypted on disk using Windows DPAPI (`CryptProtectData`/`CryptUnprotectData`).
- **Two-Stage Automatic Model Fallback (`gemini-3.8-flash-tts` -> `gemini-3.8-flash-lite-tts`)**:
  - **Stage 1**: Chunks execute round-robin across all active API keys using `gemini-3.8-flash-tts`.
  - **Per-Key Failover**: If a key encounters `quota_daily` (429 per-day quota), it is marked as exhausted for that model and the chunk immediately shifts to the next available key.
  - **Stage 2**: Once all keys exhaust quota on `gemini-3.8-flash-tts`, the system automatically shifts to `gemini-3.8-flash-lite-tts` across all keys in the pool without stopping the batch.
  - Halts the queue only when all keys are completely exhausted on both models.
- **Collision-Free Output File Naming**:
  - Automatically detects existing `.mp3` files in the target directory and assigns unique names (e.g., `input1.mp3`, `input2.mp3`) instead of overwriting existing audio files.
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
- **Fully Editable Voice & Style with Universal Clipboard**:
  - Both Voice and Style comboboxes allow free typing and pasting (e.g. custom voices like "Zali", or custom prompt styles).
  - Dedicated Paste buttons and full clipboard keyboard shortcuts (`Ctrl+V`, `Shift+Insert`, keycode 86) across all keyboard layouts.
  - Click or double-click the Voice or Style column directly in the file table to edit in-place.
- **Built-in CLI Self-Test**:
  - Run `TextToSpeechBatch.exe --selftest` to verify DPAPI encryption, LAME encoder, docx parser, chunking, pool rotation, and error classification in headless mode.

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
