"""Text to Speech Batch - Windows desktop application for converting text and Word files to MP3 using Gemini TTS.

All source code, identifiers, comments, log messages, and UI labels are written in English.
Style preset labels in Persian are represented strictly using \\uXXXX escape sequences.
Features modern dark styling, direct in-place grid editing, structured error classification,
exact substring chunking, robust encoding/docx extraction, non-spoken vocal styling,
automatic duplicate output naming, copyable activity logs, and CLI selftest.
"""

from dataclasses import dataclass
import os
import sys
import re
import json
import uuid
import time
import queue
import base64
import ctypes
from ctypes import wintypes
import threading
import hashlib
import shutil
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

import lameenc
import docx
from docx.text.paragraph import Paragraph
from docx.table import Table

from google import genai
from google.genai import types
from google.genai import errors


# ============================================================================
# Constants & Modern Dark Theme Palette
# ============================================================================

APP_NAME = "Text to Speech Batch"
APP_SUBTITLE = "AI Audio Studio"

PRIMARY_MODEL = "gemini-3.8-flash-tts"
FALLBACK_MODEL = "gemini-3.8-flash-lite-tts"
MODEL_NAME = PRIMARY_MODEL
MAX_API_KEYS = 10
MAX_CHARS = 3000
MP3_BITRATE_KBPS = 64
MP3_QUALITY = 2
DEFAULT_SAMPLE_RATE = 24000
CHUNK_GAP_SECONDS = 0.3
CACHE_EXPIRY_DAYS = 7
MIN_RATE_LIMIT_INTERVAL = 20.0

# Color Palette (Modern Dark Windows Studio)
COLOR_WINDOW_BG = "#1e1e24"
COLOR_CARD_BG = "#252836"
COLOR_INPUT_BG = "#181a20"
COLOR_BORDER = "#334155"
COLOR_BORDER_FOCUS = "#38bdf8"
COLOR_TEXT_PRIMARY = "#f8fafc"
COLOR_TEXT_SECONDARY = "#cbd5e1"
COLOR_TEXT_MUTED = "#94a3b8"
COLOR_ACCENT_BLUE = "#0284c7"
COLOR_ACCENT_HOVER = "#0369a1"
COLOR_ACCENT_CYAN = "#38bdf8"
COLOR_DANGER = "#dc2626"
COLOR_DANGER_HOVER = "#b91c1c"
COLOR_SECONDARY = "#334155"
COLOR_SECONDARY_HOVER = "#475569"
COLOR_DISABLED_BG = "#27272a"
COLOR_DISABLED_FG = "#71717a"
COLOR_TREE_HEADER = "#1e293b"
COLOR_SUCCESS = "#22c55e"
COLOR_WARN = "#f59e0b"

# Fixed list of supported Gemini TTS voices
VOICES = [
    "Bodi", "Fola", "Sola", "Zephyr", "Puck", "Charon", "Kore", "Fenrir",
    "Leda", "Orus", "Aoede", "Callirrhoe", "Autonoe", "Enceladus", "Iapetus",
    "Umbriel", "Algieba", "Despina", "Erinome", "Algenib", "Rasalgethi",
    "Laomedeia", "Achernar", "Alnilam", "Schedar", "Gacrux", "Pulcherrima",
    "Achird", "Zubenelgenubi", "Vindemiatrix", "Sadachbia", "Sadaltager",
    "Sulafat",
]

# Style Presets: English label shown in UI -> English vocal style metadata
STYLE_PRESETS = {
    "Neutral (No style)": "",
    "Cheerful & Upbeat": "cheerful, upbeat and energetic tone",
    "Sad & Melancholic": "sad, melancholic tone, slowly and softly",
    "Calm & Soothing": "calm, gentle and soothing tone",
    "Excited & Enthusiastic": "excited and enthusiastic tone",
    "Serious & Formal": "serious, formal and professional tone",
    "Warm & Friendly": "warm, friendly tone with natural micro-pauses",
    "Narration & Storytelling": "narration style with rich vocal modulation and dramatic pauses",
    "News Anchor": "news anchor, clear, neutral and steady",
    "Whisper": "whisper quietly",
    "Promotional & High-Energy": "high-energy promotional tone with elongated vowels",
    "Angry & Intense": "angry, intense tone",
}

DEFAULT_STYLE_LABEL = list(STYLE_PRESETS.keys())[0]


# ============================================================================
# Windows DPAPI and Config Storage
# ============================================================================

class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_byte)),
    ]


def protect_data(plain_bytes: bytes) -> bytes:
    """Encrypt bytes using Windows DPAPI CryptProtectData."""
    if sys.platform != "win32":
        raise OSError("DPAPI is only supported on Windows")
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32

    in_blob = DATA_BLOB(
        len(plain_bytes),
        ctypes.cast(ctypes.create_string_buffer(plain_bytes), ctypes.POINTER(ctypes.c_byte)),
    )
    out_blob = DATA_BLOB()
    if not crypt32.CryptProtectData(
        ctypes.byref(in_blob), None, None, None, None, 0, ctypes.byref(out_blob)
    ):
        raise ctypes.WinError()
    try:
        return ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        kernel32.LocalFree(out_blob.pbData)


def unprotect_data(cipher_bytes: bytes) -> bytes:
    """Decrypt bytes using Windows DPAPI CryptUnprotectData."""
    if sys.platform != "win32":
        raise OSError("DPAPI is only supported on Windows")
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32

    in_blob = DATA_BLOB(
        len(cipher_bytes),
        ctypes.cast(ctypes.create_string_buffer(cipher_bytes), ctypes.POINTER(ctypes.c_byte)),
    )
    out_blob = DATA_BLOB()
    if not crypt32.CryptUnprotectData(
        ctypes.byref(in_blob), None, None, None, None, 0, ctypes.byref(out_blob)
    ):
        raise ctypes.WinError()
    try:
        return ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        kernel32.LocalFree(out_blob.pbData)


def get_config_path() -> str:
    """Return the configuration file path in %APPDATA%\\TextToSpeechBatch\\config.json."""
    appdata = os.environ.get("APPDATA")
    if not appdata:
        appdata = os.path.expanduser("~")
    dir_path = os.path.join(appdata, "TextToSpeechBatch")
    try:
        os.makedirs(dir_path, exist_ok=True)
    except Exception:
        pass
    return os.path.join(dir_path, "config.json")


def save_api_keys(keys: list[str]) -> bool:
    """Save up to 10 API keys securely with DPAPI, falling back to plain storage."""
    config_path = get_config_path()
    cleaned = [k.strip() for k in keys if k and k.strip()][:MAX_API_KEYS]

    stored_keys = []
    for k in cleaned:
        try:
            encrypted_bytes = protect_data(k.encode("utf-8"))
            b64_str = base64.b64encode(encrypted_bytes).decode("ascii")
            stored_keys.append({"key": b64_str, "encrypted": True})
        except Exception:
            stored_keys.append({"key": k, "encrypted": False})

    data_to_store = {
        "api_keys": stored_keys,
        "api_key": stored_keys[0]["key"] if stored_keys else "",
        "encrypted": stored_keys[0]["encrypted"] if stored_keys else False,
    }

    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(data_to_store, f, indent=2)
        return True
    except Exception:
        return False


def load_api_keys() -> list[str]:
    """Load up to 10 API keys from config.json, decrypting if encrypted."""
    config_path = get_config_path()
    if not os.path.isfile(config_path):
        return []
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        keys: list[str] = []
        if "api_keys" in data and isinstance(data["api_keys"], list):
            for item in data["api_keys"]:
                if isinstance(item, dict):
                    raw_k = item.get("key", "")
                    if not raw_k:
                        continue
                    if item.get("encrypted", False):
                        try:
                            cipher_bytes = base64.b64decode(raw_k.encode("ascii"))
                            decrypted_bytes = unprotect_data(cipher_bytes)
                            keys.append(decrypted_bytes.decode("utf-8"))
                        except Exception:
                            keys.append(raw_k)
                    else:
                        keys.append(raw_k)
                elif isinstance(item, str) and item.strip():
                    keys.append(item.strip())

        # Fallback to single key if api_keys was empty
        if not keys:
            raw_key = data.get("api_key", "")
            if raw_key:
                if data.get("encrypted", False):
                    try:
                        cipher_bytes = base64.b64decode(raw_key.encode("ascii"))
                        decrypted_bytes = unprotect_data(cipher_bytes)
                        keys.append(decrypted_bytes.decode("utf-8"))
                    except Exception:
                        keys.append(raw_key)
                else:
                    keys.append(raw_key)

        seen = set()
        deduped = []
        for k in keys:
            cl = k.strip()
            if cl and cl not in seen:
                seen.add(cl)
                deduped.append(cl)
                if len(deduped) >= MAX_API_KEYS:
                    break
        return deduped
    except Exception:
        return []


def save_api_key(key: str) -> bool:
    """Save single API key for backwards compatibility."""
    return save_api_keys([key] if key.strip() else [])


def load_api_key() -> str:
    """Load primary API key for backwards compatibility."""
    keys = load_api_keys()
    return keys[0] if keys else ""


# ============================================================================
# Output Naming & Collision Avoidance
# ============================================================================

def get_unique_output_path(target_path: str, reserved_paths: set[str] | None = None) -> str:
    """Return a unique output path, appending numeric suffix (e.g. input1.mp3) if file exists or is reserved."""
    norm_reserved = {os.path.normcase(os.path.abspath(p)) for p in reserved_paths} if reserved_paths else set()
    norm_target = os.path.normcase(os.path.abspath(target_path))
    if not os.path.exists(target_path) and norm_target not in norm_reserved:
        return target_path

    directory = os.path.dirname(target_path)
    stem, ext = os.path.splitext(os.path.basename(target_path))
    counter = 1
    while True:
        candidate_name = f"{stem}{counter}{ext}"
        candidate_path = os.path.join(directory, candidate_name) if directory else candidate_name
        norm_cand = os.path.normcase(os.path.abspath(candidate_path))
        if not os.path.exists(candidate_path) and norm_cand not in norm_reserved:
            return candidate_path
        counter += 1


def check_output_collisions(file_paths: list[str]) -> dict[str, list[str]]:
    """Check if any input files would produce the same base output MP3 path."""
    dest_map: dict[str, list[str]] = {}
    for p in file_paths:
        out_path = os.path.normcase(os.path.abspath(os.path.splitext(p)[0] + ".mp3"))
        dest_map.setdefault(out_path, []).append(p)
    return {out_path: files for out_path, files in dest_map.items() if len(files) > 1}


# ============================================================================
# Error Classification
# ============================================================================

@dataclass
class ErrorInfo:
    kind: str  # 'auth', 'region', 'model_not_found', 'quota_daily', 'rate_limit', 'bad_request', 'transient'
    user_message: str


def classify_error(exc: Exception) -> ErrorInfo:
    """Classify an API or network exception into a structured ErrorInfo."""
    code = getattr(exc, "code", None) or getattr(exc, "status_code", None)
    status = str(getattr(exc, "status", "") or "").upper()
    raw_message = getattr(exc, "message", None) or str(exc)
    msg_lower = (str(raw_message) + " " + str(exc)).lower()

    # 1. auth: HTTP 401, or 403 with status PERMISSION_DENIED, or message containing "API key not valid" / "API_KEY_INVALID"
    if code == 401:
        return ErrorInfo("auth", "Invalid API key or authentication error.")
    if code == 403 and (
        "permission_denied" in status.lower()
        or "permission" in msg_lower
        or "api key not valid" in msg_lower
        or "api_key_invalid" in msg_lower
        or not status
    ):
        return ErrorInfo("auth", "Invalid API key or permission denied.")
    if "api key not valid" in msg_lower or "api_key_invalid" in msg_lower:
        return ErrorInfo("auth", "Invalid API key or authentication error.")

    # 2. region: 400 FAILED_PRECONDITION or message containing "location is not supported"
    if (
        code == 400
        and ("failed_precondition" in status.lower() or "location is not supported" in msg_lower or "user location is not supported" in msg_lower)
    ) or "location is not supported" in msg_lower or "user location is not supported" in msg_lower:
        return ErrorInfo(
            "region",
            "Gemini API reports that your location is not supported. Please check your network or proxy settings.",
        )

    # 3. model_not_found: HTTP 404
    if code == 404:
        return ErrorInfo("model_not_found", f"TTS model '{MODEL_NAME}' was not found (HTTP 404).")

    # 4. quota_daily: 429 whose message indicates a per-day quota
    if code == 429:
        if any(w in msg_lower for w in ("per day", "per_day", "daily", "quota exceeded for quota metric")):
            return ErrorInfo("quota_daily", "Daily quota exceeded for Gemini API.")
        return ErrorInfo("rate_limit", "Rate limit reached (HTTP 429).")

    # 5. bad_request: other 400 (for example an unsupported voice name)
    if code == 400:
        return ErrorInfo("bad_request", f"Bad request: {raw_message}")

    # 6. transient: 5xx, timeouts, connection errors
    if (
        (code is not None and 500 <= code < 600)
        or isinstance(exc, (TimeoutError, ConnectionError))
        or any(w in msg_lower for w in ("timeout", "timed out", "connection", "reset", "remoteprotocolerror"))
    ):
        return ErrorInfo("transient", f"Transient server or network error: {raw_message}")

    # Fallback client errors vs server errors
    if code is not None and 400 <= code < 500:
        return ErrorInfo("bad_request", f"Client error ({code}): {raw_message}")

    return ErrorInfo("transient", f"Network or API error: {raw_message}")


def extract_retry_delay(exc: Exception, attempt: int, default_sequence: list[float]) -> float:
    """Extract server retry hint if available, otherwise return backoff step."""
    msg = str(exc).lower()
    match = re.search(r"retry (?:in|after) ([\d\.]+)s", msg)
    if match:
        try:
            return float(match.group(1)) + 1.0
        except ValueError:
            pass
    if attempt < len(default_sequence):
        return default_sequence[attempt]
    return default_sequence[-1]


# ============================================================================
# API Key Pool & Round-Robin Rotation
# ============================================================================

class ApiKeyPool:
    """Manages round-robin rotation and per-model daily quota tracking across up to 10 Gemini API keys."""

    def __init__(self, keys: list[str]):
        seen = set()
        deduped = []
        for k in keys:
            cl = k.strip()
            if cl and cl not in seen:
                seen.add(cl)
                deduped.append(cl)
                if len(deduped) >= MAX_API_KEYS:
                    break
        self.keys: list[str] = deduped
        self.current_idx: int = 0
        self.exhausted_primary: set[str] = set()
        self.exhausted_fallback: set[str] = set()
        self.invalid_keys: set[str] = set()
        self.active_model: str = PRIMARY_MODEL
        self._clients: dict[str, genai.Client] = {}

    def get_client(self, api_key: str) -> genai.Client:
        if api_key not in self._clients:
            self._clients[api_key] = genai.Client(
                api_key=api_key,
                http_options=types.HttpOptions(timeout=180000),
            )
        return self._clients[api_key]

    def has_keys(self) -> bool:
        return len(self.keys) > 0

    def total_count(self) -> int:
        return len(self.keys)

    def valid_count(self) -> int:
        return len([k for k in self.keys if k not in self.invalid_keys])

    def get_active_model(self) -> str:
        return self.active_model

    def shift_to_fallback(self) -> None:
        self.active_model = FALLBACK_MODEL

    def is_all_exhausted(self) -> bool:
        valid_keys = [k for k in self.keys if k not in self.invalid_keys]
        if not valid_keys:
            return True
        return all(k in self.exhausted_primary for k in valid_keys) and all(k in self.exhausted_fallback for k in valid_keys)

    def key_index(self, api_key: str) -> int:
        try:
            return self.keys.index(api_key) + 1
        except ValueError:
            return 1

    def get_next_key_and_model(self) -> tuple[str, str, int] | None:
        """Return (api_key, model_name, key_index) for next request in round-robin order, or None if all exhausted."""
        valid_keys = [k for k in self.keys if k not in self.invalid_keys]
        if not valid_keys:
            return None

        # Stage 1: Round-robin across keys for PRIMARY_MODEL
        if self.active_model == PRIMARY_MODEL:
            primary_candidates = [k for k in valid_keys if k not in self.exhausted_primary]
            if primary_candidates:
                for _ in range(len(self.keys)):
                    candidate = self.keys[self.current_idx % len(self.keys)]
                    self.current_idx = (self.current_idx + 1) % len(self.keys)
                    if candidate in primary_candidates:
                        idx = self.keys.index(candidate) + 1
                        return candidate, PRIMARY_MODEL, idx

            # All valid keys exhausted for PRIMARY_MODEL -> shift to FALLBACK_MODEL!
            self.active_model = FALLBACK_MODEL

        # Stage 2: Round-robin across keys for FALLBACK_MODEL
        if self.active_model == FALLBACK_MODEL:
            fallback_candidates = [k for k in valid_keys if k not in self.exhausted_fallback]
            if fallback_candidates:
                for _ in range(len(self.keys)):
                    candidate = self.keys[self.current_idx % len(self.keys)]
                    self.current_idx = (self.current_idx + 1) % len(self.keys)
                    if candidate in fallback_candidates:
                        idx = self.keys.index(candidate) + 1
                        return candidate, FALLBACK_MODEL, idx

        # All keys exhausted on both models!
        return None

    def mark_quota_daily(self, api_key: str, model_name: str) -> tuple[bool, str]:
        """Mark daily quota exhausted for an API key on a model.
        
        Returns:
            (all_keys_exhausted_on_this_model, log_message)
        """
        valid_keys = [k for k in self.keys if k not in self.invalid_keys]
        key_num = self.keys.index(api_key) + 1 if api_key in self.keys else 1

        if model_name == PRIMARY_MODEL:
            self.exhausted_primary.add(api_key)
            all_primary = all(k in self.exhausted_primary for k in valid_keys)
            if all_primary:
                self.active_model = FALLBACK_MODEL
                return True, f"Daily quota exhausted across all {len(valid_keys)} API key(s) on primary model '{PRIMARY_MODEL}'."
            return False, f"API key #{key_num} exhausted daily quota for '{PRIMARY_MODEL}'. Shifting to next key..."
        else:
            self.exhausted_fallback.add(api_key)
            all_fallback = all(k in self.exhausted_fallback for k in valid_keys)
            if all_fallback:
                return True, f"Daily quota exhausted across all {len(valid_keys)} API key(s) on fallback model '{FALLBACK_MODEL}'."
            return False, f"API key #{key_num} exhausted daily quota for '{FALLBACK_MODEL}'. Shifting to next key..."

    def mark_invalid(self, api_key: str):
        """Mark an API key as invalid (e.g. auth error)."""
        self.invalid_keys.add(api_key)


# ============================================================================
# Cache Management
# ============================================================================

def get_cache_root_dir() -> str:
    """Return root directory for cached audio chunks in %TEMP%\\TextToSpeechBatch\\cache."""
    root = os.path.join(tempfile.gettempdir(), "TextToSpeechBatch", "cache")
    try:
        os.makedirs(root, exist_ok=True)
    except Exception:
        pass
    return root


def get_file_cache_dir(file_path: str, voice: str, resolved_style: str) -> str:
    """Compute deterministic cache directory for a file based on path, voice, style, and model."""
    key = f"{os.path.abspath(file_path)}|{voice}|{resolved_style}|{MODEL_NAME}"
    hash_str = hashlib.sha256(key.encode("utf-8")).hexdigest()
    file_cache_dir = os.path.join(get_cache_root_dir(), hash_str)
    try:
        os.makedirs(file_cache_dir, exist_ok=True)
    except Exception:
        pass
    return file_cache_dir


def cleanup_old_caches(max_age_days: int = CACHE_EXPIRY_DAYS):
    """Purge cached directories older than max_age_days at application startup."""
    root = get_cache_root_dir()
    if not os.path.isdir(root):
        return
    now = time.time()
    cutoff_seconds = max_age_days * 86400
    try:
        for entry in os.scandir(root):
            if entry.is_dir():
                try:
                    mtime = entry.stat().st_mtime
                    if now - mtime > cutoff_seconds:
                        shutil.rmtree(entry.path, ignore_errors=True)
                except Exception:
                    pass
    except Exception:
        pass


# ============================================================================
# Document Extraction & Chunking
# ============================================================================

def extract_text_from_txt(file_path: str) -> str:
    """Extract text from a .txt file strictly respecting BOM, strict UTF-8, and strict CP1256."""
    with open(file_path, "rb") as f:
        raw_bytes = f.read()

    # 1. BOM detection (never fall back to ungrounded UTF-16)
    if raw_bytes.startswith(b"\xef\xbb\xbf"):
        text = raw_bytes.decode("utf-8-sig")
        if not text.strip():
            raise ValueError("no text found")
        return text
    if raw_bytes.startswith(b"\xff\xfe") or raw_bytes.startswith(b"\xfe\xff"):
        text = raw_bytes.decode("utf-16")
        if not text.strip():
            raise ValueError("no text found")
        return text

    # 2. Strict UTF-8 without BOM
    try:
        text = raw_bytes.decode("utf-8")
        if text.strip():
            return text
        raise ValueError("no text found")
    except UnicodeDecodeError:
        pass

    # 3. Strict CP1256 without BOM (never try UTF-16 without BOM)
    try:
        text = raw_bytes.decode("cp1256")
        if text.strip():
            return text
        raise ValueError("no text found")
    except UnicodeDecodeError:
        pass

    # 4. Fallback UTF-8 with error replacement
    text = raw_bytes.decode("utf-8", errors="replace")
    if text.strip():
        return text
    raise ValueError("no text found")


def _extract_blocks_from_element(elem, doc: docx.Document) -> list[str]:
    """Recursively extract paragraphs, deduplicated tables, and SDT content controls."""
    blocks: list[str] = []
    for child in elem:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "p":
            p = Paragraph(child, doc)
            txt = p.text.strip()
            if txt:
                blocks.append(txt)
        elif tag == "tbl":
            tbl = Table(child, doc)
            table_rows = []
            seen_tcs = set()
            for row in tbl.rows:
                row_cells = []
                for cell in row.cells:
                    if cell._tc in seen_tcs:
                        continue
                    seen_tcs.add(cell._tc)
                    c_txt = cell.text.strip()
                    if c_txt:
                        row_cells.append(c_txt)
                if row_cells:
                    table_rows.append(" ".join(row_cells))
            if table_rows:
                blocks.append("\n".join(table_rows))
        elif tag == "sdt":
            content_elem = None
            for c in child:
                if c.tag.endswith("sdtContent"):
                    content_elem = c
                    break
            if content_elem is not None:
                sub_blocks = _extract_blocks_from_element(content_elem, doc)
                blocks.extend(sub_blocks)
    return blocks


def extract_text_from_docx(file_path: str) -> str:
    """Extract text from a .docx file in natural document order with merged cell deduplication."""
    try:
        doc = docx.Document(file_path)
    except Exception as exc:
        raise ValueError(f"Failed to read Word document: {exc}")

    body_blocks = _extract_blocks_from_element(doc.element.body, doc)
    full_text = "\n\n".join(body_blocks).strip()
    if not full_text:
        raise ValueError("no text found")
    return full_text


def extract_text(file_path: str) -> str:
    """Extract text from .txt or .docx file and reject empty content or legacy .doc."""
    lower_path = file_path.lower()
    if lower_path.endswith(".txt"):
        return extract_text_from_txt(file_path)
    if lower_path.endswith(".docx"):
        return extract_text_from_docx(file_path)
    if lower_path.endswith(".doc"):
        raise ValueError("Legacy .doc format is not supported. Please save as .docx.")
    raise ValueError(f"Unsupported file format: {os.path.splitext(file_path)[1]}")


def split_text_into_chunks(text: str, max_chars: int = MAX_CHARS) -> list[str]:
    """Split text into chunks of at most max_chars without altering text or punctuation."""
    clean_text = text.strip()
    if not clean_text:
        return []

    # Split boundaries: paragraph breaks, line breaks, or sentence enders followed by whitespace
    split_pattern = r"(\r?\n\s*\r?\n+|\r?\n|(?<=[.!?:\u061f\u061b])\s+)"
    parts = re.split(split_pattern, clean_text)

    tokens: list[tuple[str, str]] = []
    i = 0
    while i < len(parts):
        piece = parts[i]
        sep = parts[i + 1] if i + 1 < len(parts) else ""
        i += 2

        if not piece and not sep:
            continue

        if len(piece) > max_chars:
            start = 0
            while start < len(piece):
                remaining = piece[start:]
                if len(remaining) <= max_chars:
                    tokens.append((remaining, sep))
                    break
                sub = remaining[:max_chars]
                last_space = max(sub.rfind(" "), sub.rfind("\t"))
                if last_space > 0:
                    tokens.append((remaining[:last_space], remaining[last_space:last_space + 1]))
                    start += last_space + 1
                else:
                    tokens.append((sub, ""))
                    start += max_chars
        else:
            tokens.append((piece, sep))

    chunks: list[str] = []
    current_chunk = ""
    pending_sep = ""

    for piece, sep in tokens:
        if not piece:
            pending_sep += sep
            continue
        if not current_chunk:
            current_chunk = piece
            pending_sep = sep
            continue

        if len(current_chunk) + len(pending_sep) + len(piece) <= max_chars:
            current_chunk += pending_sep + piece
            pending_sep = sep
        else:
            chunks.append(current_chunk)
            current_chunk = piece
            pending_sep = sep

    if current_chunk:
        chunks.append(current_chunk)

    return chunks


# ============================================================================
# Prompt Construction & Audio Encoding
# ============================================================================

def resolve_style_instruction(style_value: str) -> str:
    """Convert preset label or user free text into a style description string."""
    text = (style_value or "").strip()
    if not text:
        return ""
    if text in STYLE_PRESETS:
        return STYLE_PRESETS[text]
    return text


def clean_style_instruction(style_value: str) -> str:
    """Extract clean style description without imperative prefixes (e.g. 'Say in a') or trailing colons."""
    raw = resolve_style_instruction(style_value).strip()
    clean = raw.rstrip(":").strip()
    prefixes = ("say in a ", "say in an ", "say in ", "say like a ", "say like ", "speak in a ", "speak in ")
    lower_c = clean.lower()
    for p in prefixes:
        if lower_c.startswith(p):
            return clean[len(p):].strip()
    return clean


def build_prompt(style_instruction: str, chunk: str) -> str:
    """Return verbatim text chunk for Gemini TTS synthesis, ensuring directions are never spoken."""
    return chunk.strip()


def parse_sample_rate(mime_type: str, default_rate: int = DEFAULT_SAMPLE_RATE) -> int:
    """Parse sample rate integer from audio mime type, e.g. audio/L16;rate=24000."""
    if not mime_type:
        return default_rate
    match = re.search(r"rate=(\d+)", mime_type, re.IGNORECASE)
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            pass
    return default_rate


def save_pcm_as_mp3(pcm_data: bytes, sample_rate: int, output_mp3_path: str):
    """Encode 16-bit mono PCM into MP3 (64 kbps, quality 2) and atomically write."""
    encoder = lameenc.Encoder()
    encoder.set_bit_rate(MP3_BITRATE_KBPS)
    encoder.set_in_sample_rate(sample_rate)
    encoder.set_channels(1)
    encoder.set_out_sample_rate(sample_rate)
    encoder.set_quality(MP3_QUALITY)

    mp3_bytes = encoder.encode(pcm_data)
    mp3_bytes += encoder.flush()

    out_dir = os.path.dirname(output_mp3_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    temp_path = os.path.join(out_dir, f".{os.path.basename(output_mp3_path)}.{uuid.uuid4().hex[:8]}.tmp")

    try:
        with open(temp_path, "wb") as f:
            f.write(mp3_bytes)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp_path, output_mp3_path)
    except Exception:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass
        raise


# ============================================================================
# Main Application GUI
# ============================================================================

class TTSBatchApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_NAME} — {APP_SUBTITLE}")
        self.root.geometry("940x780")
        self.root.minsize(820, 640)
        self.root.configure(bg=COLOR_WINDOW_BG)

        # State
        self.active_model = PRIMARY_MODEL
        self.api_keys_list: list[str] = []
        self.api_key_var = tk.StringVar()
        self.last_applied_voice = VOICES[0]
        self.last_applied_style = DEFAULT_STYLE_LABEL
        self.is_processing = False
        self.stop_event = threading.Event()
        self.queue: queue.Queue = queue.Queue()
        self.active_cell_editor: ttk.Combobox | None = None
        self._is_programmatic_update = False

        # Clean old caches from past sessions (> 7 days)
        cleanup_old_caches(CACHE_EXPIRY_DAYS)

        # Apply Modern Dark Stylesheet
        self._setup_dark_theme()

        # Build UI
        self._build_ui()

        # Initial greeting in the log
        self.log(f"{APP_NAME} initialized (Primary: {PRIMARY_MODEL}, Fallback: {FALLBACK_MODEL}). Ready for batch processing.", "ok")

        # Load persisted API keys (up to 10)
        self.api_keys_list = load_api_keys()
        if self.api_keys_list:
            self.api_key_var.set(self.api_keys_list[0])
            count = len(self.api_keys_list)
            self.manage_keys_btn.configure(text=f"Manage Pool ({count} active)")
            self.log(f"Loaded {count} saved API key(s) in pool from Windows DPAPI encrypted storage.", "ok")

        # Start queue polling
        self.root.after(100, self._process_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _setup_dark_theme(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # Global Frame and Card Styles
        style.configure("TFrame", background=COLOR_WINDOW_BG)
        style.configure("Card.TFrame", background=COLOR_CARD_BG)

        style.configure(
            "TLabelframe",
            background=COLOR_CARD_BG,
            bordercolor=COLOR_BORDER,
            borderwidth=1,
            relief="solid",
        )
        style.configure(
            "TLabelframe.Label",
            background=COLOR_CARD_BG,
            foreground=COLOR_ACCENT_CYAN,
            font=("Segoe UI", 10, "bold"),
        )

        # Labels
        style.configure("TLabel", background=COLOR_WINDOW_BG, foreground=COLOR_TEXT_PRIMARY, font=("Segoe UI", 9))
        style.configure("Card.TLabel", background=COLOR_CARD_BG, foreground=COLOR_TEXT_PRIMARY, font=("Segoe UI", 9))
        style.configure("HeaderTitle.TLabel", background=COLOR_WINDOW_BG, foreground=COLOR_TEXT_PRIMARY, font=("Segoe UI", 13, "bold"))
        style.configure("HeaderTag.TLabel", background=COLOR_WINDOW_BG, foreground=COLOR_ACCENT_CYAN, font=("Segoe UI", 10, "bold"))

        # Primary Action Button (Start)
        style.configure(
            "Primary.TButton",
            background=COLOR_ACCENT_BLUE,
            foreground="#ffffff",
            font=("Segoe UI", 9, "bold"),
            borderwidth=0,
            padding=6,
        )
        style.map(
            "Primary.TButton",
            background=[("active", COLOR_ACCENT_HOVER), ("disabled", COLOR_DISABLED_BG)],
            foreground=[("disabled", COLOR_DISABLED_FG)],
        )

        # Danger Button (Stop)
        style.configure(
            "Danger.TButton",
            background=COLOR_DANGER,
            foreground="#ffffff",
            font=("Segoe UI", 9, "bold"),
            borderwidth=0,
            padding=6,
        )
        style.map(
            "Danger.TButton",
            background=[("active", COLOR_DANGER_HOVER), ("disabled", COLOR_DISABLED_BG)],
            foreground=[("disabled", COLOR_DISABLED_FG)],
        )

        # Secondary Button
        style.configure(
            "Secondary.TButton",
            background=COLOR_SECONDARY,
            foreground=COLOR_TEXT_PRIMARY,
            font=("Segoe UI", 9),
            borderwidth=0,
            padding=5,
        )
        style.map(
            "Secondary.TButton",
            background=[("active", COLOR_SECONDARY_HOVER), ("disabled", COLOR_DISABLED_BG)],
            foreground=[("disabled", COLOR_DISABLED_FG)],
        )

        # Text Entry
        style.configure(
            "TEntry",
            fieldbackground=COLOR_INPUT_BG,
            foreground=COLOR_TEXT_PRIMARY,
            insertcolor=COLOR_ACCENT_CYAN,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            padding=5,
        )

        # Dropdowns (Combobox)
        style.configure(
            "TCombobox",
            fieldbackground=COLOR_INPUT_BG,
            foreground=COLOR_TEXT_PRIMARY,
            bordercolor=COLOR_BORDER,
            arrowcolor=COLOR_ACCENT_CYAN,
            padding=4,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", COLOR_INPUT_BG)],
            selectbackground=[("readonly", COLOR_INPUT_BG)],
            selectforeground=[("readonly", COLOR_TEXT_PRIMARY)],
        )
        self.root.option_add("*TCombobox*Listbox.background", COLOR_INPUT_BG)
        self.root.option_add("*TCombobox*Listbox.foreground", COLOR_TEXT_PRIMARY)
        self.root.option_add("*TCombobox*Listbox.selectBackground", COLOR_ACCENT_BLUE)
        self.root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")

        # Checkbutton
        style.configure(
            "TCheckbutton",
            background=COLOR_WINDOW_BG,
            foreground=COLOR_TEXT_PRIMARY,
            focuscolor=COLOR_WINDOW_BG,
            font=("Segoe UI", 9),
        )
        style.map("TCheckbutton", background=[("active", COLOR_WINDOW_BG)])

        # Treeview (File Queue)
        style.configure(
            "Treeview",
            background=COLOR_INPUT_BG,
            foreground=COLOR_TEXT_PRIMARY,
            fieldbackground=COLOR_INPUT_BG,
            bordercolor=COLOR_BORDER,
            borderwidth=1,
            rowheight=26,
            font=("Segoe UI", 9),
        )
        style.map(
            "Treeview",
            background=[("selected", COLOR_ACCENT_BLUE)],
            foreground=[("selected", "#ffffff")],
        )
        style.configure(
            "Treeview.Heading",
            background=COLOR_TREE_HEADER,
            foreground=COLOR_ACCENT_CYAN,
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padding=4,
        )
        style.map("Treeview.Heading", background=[("active", COLOR_SECONDARY)])

        # Progressbar
        style.configure(
            "Horizontal.TProgressbar",
            background=COLOR_ACCENT_BLUE,
            troughcolor=COLOR_INPUT_BG,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_ACCENT_BLUE,
            darkcolor=COLOR_ACCENT_BLUE,
        )

        # Scrollbars
        style.configure(
            "Vertical.TScrollbar",
            background=COLOR_SECONDARY,
            troughcolor=COLOR_WINDOW_BG,
            bordercolor=COLOR_WINDOW_BG,
            arrowcolor=COLOR_TEXT_SECONDARY,
        )

    def _build_ui(self):
        main_container = ttk.Frame(self.root, padding="14")
        main_container.pack(fill=tk.BOTH, expand=True)

        # --------------------------------------------------------------------
        # Top Title Bar
        # --------------------------------------------------------------------
        top_bar = ttk.Frame(main_container)
        top_bar.pack(fill=tk.X, pady=(0, 10))

        title_lbl = ttk.Label(top_bar, text=APP_NAME, style="HeaderTitle.TLabel")
        title_lbl.pack(side=tk.LEFT)

        tag_lbl = ttk.Label(top_bar, text=f" — {APP_SUBTITLE}", style="HeaderTag.TLabel")
        tag_lbl.pack(side=tk.LEFT)

        # --------------------------------------------------------------------
        # Row 1: File Queue Card (Treeview with direct in-place editing)
        # --------------------------------------------------------------------
        list_card = ttk.LabelFrame(
            main_container,
            text=" Batch Conversion Queue (Click or double-click Voice or Style in grid to edit directly) ",
            padding="8",
        )
        list_card.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        tree_container = ttk.Frame(list_card, style="Card.TFrame")
        tree_container.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        columns = ("filepath", "voice", "style", "status")
        self.tree = ttk.Treeview(
            tree_container, columns=columns, show="headings", selectmode="extended", height=6
        )

        self.tree.heading("filepath", text="File Path", anchor=tk.W)
        self.tree.heading("voice", text="Voice ▾", anchor=tk.CENTER)
        self.tree.heading("style", text="Style ▾", anchor=tk.CENTER)
        self.tree.heading("status", text="Status", anchor=tk.CENTER)

        self.tree.column("filepath", width=420, anchor=tk.W, stretch=True)
        self.tree.column("voice", width=120, anchor=tk.CENTER, stretch=False)
        self.tree.column("style", width=200, anchor=tk.CENTER, stretch=False)
        self.tree.column("status", width=160, anchor=tk.CENTER, stretch=False)

        tree_scroll = ttk.Scrollbar(
            tree_container, orient=tk.VERTICAL, command=self.tree.yview, style="Vertical.TScrollbar"
        )
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Selection and in-place cell editing bindings
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Button-1>", self._on_tree_cell_click)
        self.tree.bind("<Double-Button-1>", self._on_tree_cell_click)

        # Action buttons for the queue
        btn_bar = ttk.Frame(list_card, style="Card.TFrame")
        btn_bar.pack(fill=tk.X)

        self.add_btn = ttk.Button(btn_bar, text="+ Add files...", command=self._add_files, style="Secondary.TButton")
        self.add_btn.pack(side=tk.LEFT, padx=(0, 6))

        self.remove_btn = ttk.Button(btn_bar, text="- Remove selected", command=self._remove_selected, style="Secondary.TButton")
        self.remove_btn.pack(side=tk.LEFT, padx=(0, 6))

        self.clear_btn = ttk.Button(btn_bar, text="Clear list", command=self._clear_list, style="Secondary.TButton")
        self.clear_btn.pack(side=tk.LEFT)

        # --------------------------------------------------------------------
        # Row 3: Settings Panel Card
        # --------------------------------------------------------------------
        settings_card = ttk.LabelFrame(main_container, text=" Voice & Style Settings for Selected File(s) ", padding="8")
        settings_card.pack(fill=tk.X, pady=(0, 8))

        # Voice: editable Combobox + Paste button
        ttk.Label(settings_card, text="Voice:", style="Card.TLabel").grid(row=0, column=0, padx=(4, 6), sticky=tk.W)
        self.voice_var = tk.StringVar(value=self.last_applied_voice)
        self.voice_cb = ttk.Combobox(
            settings_card, textvariable=self.voice_var, values=VOICES, state="normal", width=14
        )
        self.voice_cb.grid(row=0, column=1, padx=(0, 4), sticky=tk.W)
        self.voice_cb.bind("<<ComboboxSelected>>", self._on_panel_voice_change)
        self.voice_cb.bind("<Return>", self._on_panel_voice_change)
        self.voice_cb.bind("<FocusOut>", self._on_panel_voice_focus_out)

        self.voice_paste_btn = ttk.Button(
            settings_card, text="Paste", command=self._paste_voice, style="Secondary.TButton"
        )
        self.voice_paste_btn.grid(row=0, column=2, padx=(0, 16), sticky=tk.W)

        # Style: editable Combobox + Paste button
        ttk.Label(settings_card, text="Style:", style="Card.TLabel").grid(row=0, column=3, padx=(0, 6), sticky=tk.W)
        self.style_var = tk.StringVar(value=self.last_applied_style)
        self.style_cb = ttk.Combobox(
            settings_card, textvariable=self.style_var, values=list(STYLE_PRESETS.keys()), state="normal", width=22
        )
        self.style_cb.grid(row=0, column=4, padx=(0, 4), sticky=tk.W)
        self.style_cb.bind("<<ComboboxSelected>>", self._on_panel_style_change)
        self.style_cb.bind("<Return>", self._on_panel_style_change)
        self.style_cb.bind("<FocusOut>", self._on_panel_style_focus_out)

        self.style_paste_btn = ttk.Button(
            settings_card, text="Paste", command=self._paste_style, style="Secondary.TButton"
        )
        self.style_paste_btn.grid(row=0, column=5, padx=(0, 16), sticky=tk.W)

        self.apply_btn = ttk.Button(settings_card, text="Apply to Selected", command=self._apply_settings, style="Secondary.TButton")
        self.apply_btn.grid(row=0, column=6, padx=(0, 6))

        self.apply_all_btn = ttk.Button(settings_card, text="Apply to All", command=self._apply_to_all, style="Secondary.TButton")
        self.apply_all_btn.grid(row=0, column=7)

        # --------------------------------------------------------------------
        # Row 3: Action & Progress Row
        # --------------------------------------------------------------------
        action_frame = ttk.Frame(main_container)
        action_frame.pack(fill=tk.X, pady=(0, 8))

        self.start_btn = ttk.Button(action_frame, text="▶ Start Batch", command=self._start_batch, style="Primary.TButton")
        self.start_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = ttk.Button(action_frame, text="■ Stop", command=self._stop_batch, state="disabled", style="Danger.TButton")
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 14))

        self.manage_keys_btn = ttk.Button(
            action_frame,
            text="Manage API Pool...",
            command=self._open_api_manager,
            style="Primary.TButton",
        )
        self.manage_keys_btn.pack(side=tk.RIGHT, padx=(8, 0))

        self.progress_label = ttk.Label(action_frame, text="Ready", width=20, anchor=tk.E)
        self.progress_label.pack(side=tk.RIGHT, padx=(0, 8))

        self.progress_bar = ttk.Progressbar(action_frame, orient=tk.HORIZONTAL, mode="determinate", style="Horizontal.TProgressbar")
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # --------------------------------------------------------------------
        # Row 5: Activity Log Card (Expansive, fully selectable & copyable console)
        # --------------------------------------------------------------------
        log_card = ttk.LabelFrame(main_container, text=" Activity Log ", padding="8")
        log_card.pack(fill=tk.BOTH, expand=True, pady=(0, 2))

        # Log Top Bar with Copy & Clear Actions
        log_top_bar = ttk.Frame(log_card, style="Card.TFrame")
        log_top_bar.pack(fill=tk.X, pady=(0, 4))

        log_lbl = ttk.Label(log_top_bar, text="Console Output:", style="Card.TLabel")
        log_lbl.pack(side=tk.LEFT)

        self.copy_log_btn = ttk.Button(
            log_top_bar, text="Copy Log", command=self._copy_all_logs, style="Secondary.TButton"
        )
        self.copy_log_btn.pack(side=tk.RIGHT, padx=(4, 0))

        self.clear_log_btn = ttk.Button(
            log_top_bar, text="Clear Log", command=self._clear_logs, style="Secondary.TButton"
        )
        self.clear_log_btn.pack(side=tk.RIGHT)

        self.log_text = ScrolledText(
            log_card,
            height=8,
            wrap=tk.WORD,
            state="disabled",
            bg=COLOR_INPUT_BG,
            fg=COLOR_TEXT_PRIMARY,
            insertbackground=COLOR_ACCENT_CYAN,
            selectbackground=COLOR_ACCENT_BLUE,
            selectforeground="#ffffff",
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_BORDER_FOCUS,
            font=("Consolas", 9),
            relief="flat",
        )
        self.log_text.tag_configure("time", foreground=COLOR_ACCENT_CYAN)
        self.log_text.tag_configure("ok", foreground=COLOR_SUCCESS)
        self.log_text.tag_configure("err", foreground=COLOR_DANGER)
        self.log_text.tag_configure("warn", foreground=COLOR_WARN)
        self.log_text.tag_configure("sub", foreground=COLOR_TEXT_MUTED)
        self.log_text.tag_configure("info", foreground=COLOR_TEXT_PRIMARY)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Attach universal clipboard handling (context menu & keycodes)
        self._attach_universal_clipboard(self.voice_cb, on_change=self._on_panel_voice_change)
        self._attach_universal_clipboard(self.style_cb, on_change=self._on_panel_style_change)
        self._attach_log_clipboard(self.log_text)

    # ------------------------------------------------------------------------
    # In-Place Grid Editing (Voice & Style)
    # ------------------------------------------------------------------------

    def _on_tree_cell_click(self, event):
        """Detect clicks on Voice or Style column to open in-place dropdown editor."""
        if self.is_processing:
            return
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return

        if column in ("#2", "#3"):
            self.root.after(20, lambda: self._open_cell_editor(row_id, column))

    def _open_cell_editor(self, row_id: str, column: str):
        if self.is_processing or not self.tree.exists(row_id):
            return

        if self.active_cell_editor and self.active_cell_editor.winfo_exists():
            self.active_cell_editor.destroy()
            self.active_cell_editor = None

        if column == "#2":
            self._edit_grid_cell(row_id, column, "voice", VOICES, is_readonly=False)
        elif column == "#3":
            self._edit_grid_cell(row_id, column, "style", list(STYLE_PRESETS.keys()), is_readonly=False)

    def _edit_grid_cell(self, row_id: str, col_id: str, col_name: str, values: list[str], is_readonly: bool):
        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, width, height = bbox
        current_val = self.tree.set(row_id, col_name)

        cb = ttk.Combobox(
            self.tree,
            values=values,
            state="readonly" if is_readonly else "normal",
        )
        cb.set(current_val)
        cb.place(x=x, y=y, width=width, height=height)
        cb.focus_set()
        if not is_readonly:
            cb.select_range(0, tk.END)
            cb.icursor(tk.END)
        self.active_cell_editor = cb
        self._attach_universal_clipboard(cb)

        def commit(evt=None):
            if not cb.winfo_exists():
                return
            new_val = cb.get().strip()
            if new_val:
                self.tree.set(row_id, col_name, new_val)
                file_name = os.path.basename(self.tree.set(row_id, "filepath"))
                self._is_programmatic_update = True
                try:
                    if col_name == "voice":
                        self.last_applied_voice = new_val
                        self.voice_var.set(new_val)
                        self.log(f"Grid edit: Set Voice of '{file_name}' to '{new_val}'.", "info")
                    elif col_name == "style":
                        self.last_applied_style = new_val
                        self.style_var.set(new_val)
                        self.log(f"Grid edit: Set Style of '{file_name}' to '{new_val}'.", "info")
                finally:
                    self._is_programmatic_update = False
            cb.destroy()
            self.active_cell_editor = None

        def cancel(evt=None):
            if cb.winfo_exists():
                cb.destroy()
            self.active_cell_editor = None

        cb.bind("<<ComboboxSelected>>", commit)
        cb.bind("<Return>", commit)
        cb.bind("<Escape>", cancel)

        def on_focus_out(event):
            def check():
                try:
                    fw = self.root.focus_get()
                    if fw and ("combobox" in str(fw).lower() or "listbox" in str(fw).lower()):
                        return
                    commit()
                except Exception:
                    pass
            self.root.after(120, check)

        cb.bind("<FocusOut>", on_focus_out)

    # ------------------------------------------------------------------------
    # Panel Auto-Apply Handlers
    # ------------------------------------------------------------------------

    def _on_panel_voice_change(self, event=None):
        if self._is_programmatic_update or self.is_processing:
            return
        selected = self.tree.selection()
        new_val = self.voice_var.get().strip() or VOICES[0]
        self.last_applied_voice = new_val
        if selected:
            for item in selected:
                self.tree.set(item, "voice", new_val)
            self.log(f"Auto-applied Voice '{new_val}' to {len(selected)} selected file(s).", "info")

    def _on_panel_voice_focus_out(self, event=None):
        if self._is_programmatic_update or self.is_processing:
            return
        def check():
            try:
                fw = self.root.focus_get()
                if fw and ("combobox" in str(fw).lower() or "listbox" in str(fw).lower()):
                    return
                self._on_panel_voice_change()
            except Exception:
                pass
        self.root.after(120, check)

    def _on_panel_style_change(self, event=None):
        if self._is_programmatic_update or self.is_processing:
            return
        selected = self.tree.selection()
        new_val = self.style_var.get().strip() or DEFAULT_STYLE_LABEL
        self.last_applied_style = new_val
        if selected:
            for item in selected:
                self.tree.set(item, "style", new_val)
            self.log(f"Auto-applied Style '{new_val}' to {len(selected)} selected file(s).", "info")

    def _on_panel_style_focus_out(self, event=None):
        if self._is_programmatic_update or self.is_processing:
            return
        def check():
            try:
                fw = self.root.focus_get()
                if fw and ("combobox" in str(fw).lower() or "listbox" in str(fw).lower()):
                    return
                self._on_panel_style_change()
            except Exception:
                pass
        self.root.after(120, check)

    # ------------------------------------------------------------------------
    # Clipboard & UI Handlers
    # ------------------------------------------------------------------------

    def _open_api_manager(self):
        """Open the API Key Pool manager dialog to configure up to 10 Gemini API keys."""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Gemini API Keys Pool — {APP_SUBTITLE}")
        dialog.geometry("720x620")
        dialog.minsize(640, 500)
        dialog.configure(bg=COLOR_WINDOW_BG)
        dialog.transient(self.root)
        dialog.grab_set()

        # Header Frame
        header_frame = ttk.Frame(dialog, padding="14")
        header_frame.pack(fill=tk.X)

        title_lbl = ttk.Label(
            header_frame,
            text="Gemini API Key Pool (Up to 10 Keys)",
            font=("Segoe UI", 12, "bold"),
            foreground=COLOR_ACCENT_CYAN,
        )
        title_lbl.pack(anchor=tk.W, pady=(0, 4))

        info_text = (
            "Define up to 10 Gemini API keys for seamless batch processing:\n"
            "• Stage 1: The engine rotates requests across all active keys in round-robin order using 'gemini-3.8-flash-tts'.\n"
            "• Stage 2: When daily quota is reached on all keys, it automatically shifts to 'gemini-3.8-flash-lite-tts' across all keys."
        )
        info_lbl = ttk.Label(
            header_frame,
            text=info_text,
            font=("Segoe UI", 9),
            foreground=COLOR_TEXT_SECONDARY,
            justify=tk.LEFT,
        )
        info_lbl.pack(anchor=tk.W)

        # Show/Hide Checkbox
        show_all_var = tk.BooleanVar(value=False)
        show_check = ttk.Checkbutton(
            header_frame,
            text="Show API Keys",
            variable=show_all_var,
        )
        show_check.configure(style="TCheckbutton")
        show_check.pack(anchor=tk.E, pady=(4, 0))

        # Scrollable Canvas container for keys
        list_container = ttk.Frame(dialog, padding="14")
        list_container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(list_container, bg=COLOR_CARD_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style="Card.TFrame", padding="8")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)

        canvas.bind("<Configure>", on_canvas_configure)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        key_vars: list[tk.StringVar] = []
        key_entries: list[ttk.Entry] = []

        existing_keys = list(self.api_keys_list)
        cur_ent = self.api_key_var.get().strip()
        if cur_ent and cur_ent not in existing_keys:
            existing_keys.insert(0, cur_ent)

        def update_show_mask():
            mask_char = "" if show_all_var.get() else "*"
            for ent in key_entries:
                ent.configure(show=mask_char)

        show_check.configure(command=update_show_mask)

        # Create 10 numbered slots
        for i in range(MAX_API_KEYS):
            row_frame = ttk.Frame(scrollable_frame, style="Card.TFrame", padding="3")
            row_frame.pack(fill=tk.X, pady=2)

            lbl = ttk.Label(
                row_frame,
                text=f"Key #{i+1:2d}:",
                width=8,
                font=("Segoe UI", 9, "bold"),
                style="Card.TLabel",
            )
            lbl.pack(side=tk.LEFT, padx=(4, 6))

            initial_val = existing_keys[i] if i < len(existing_keys) else ""
            var = tk.StringVar(value=initial_val)
            key_vars.append(var)

            ent = ttk.Entry(row_frame, textvariable=var, show="*" if not show_all_var.get() else "")
            ent.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
            key_entries.append(ent)
            self._attach_universal_clipboard(ent)

            def make_paste_fn(target_var=var):
                def fn():
                    try:
                        clip = dialog.clipboard_get().strip()
                        if clip:
                            first_line = clip.splitlines()[0].strip()
                            target_var.set(first_line)
                    except Exception:
                        pass
                return fn

            paste_btn = ttk.Button(row_frame, text="Paste", command=make_paste_fn(), style="Secondary.TButton")
            paste_btn.pack(side=tk.LEFT, padx=(0, 4))

            def make_clear_fn(target_var=var):
                def fn():
                    target_var.set("")
                return fn

            clr_btn = ttk.Button(row_frame, text="Clear", command=make_clear_fn(), style="Secondary.TButton")
            clr_btn.pack(side=tk.LEFT, padx=(0, 4))

        # Bottom Actions Bar
        bottom_bar = ttk.Frame(dialog, padding="14")
        bottom_bar.pack(fill=tk.X)

        def paste_multiline():
            try:
                clip = dialog.clipboard_get()
                lines = [line.strip() for line in clip.splitlines() if line.strip()]
                if not lines:
                    messagebox.showinfo("Clipboard Empty", "The clipboard does not contain any text lines.")
                    return
                for i in range(MAX_API_KEYS):
                    if i < len(lines):
                        key_vars[i].set(lines[i])
                    else:
                        key_vars[i].set("")
            except Exception as exc:
                messagebox.showinfo("Clipboard Notice", f"Unable to read clipboard: {exc}")

        paste_multi_btn = ttk.Button(
            bottom_bar,
            text="Paste Multi-line (Auto-fill)",
            command=paste_multiline,
            style="Secondary.TButton",
        )
        paste_multi_btn.pack(side=tk.LEFT, padx=(0, 8))

        def clear_all_dialog():
            for v in key_vars:
                v.set("")

        clear_all_dialog_btn = ttk.Button(
            bottom_bar,
            text="Clear All",
            command=clear_all_dialog,
            style="Secondary.TButton",
        )
        clear_all_dialog_btn.pack(side=tk.LEFT)

        def save_and_close():
            non_empty = [v.get().strip() for v in key_vars if v.get().strip()]
            seen_keys = set()
            clean_keys = []
            for k in non_empty:
                if k not in seen_keys:
                    seen_keys.add(k)
                    clean_keys.append(k)

            save_api_keys(clean_keys)
            self.api_keys_list = clean_keys
            if clean_keys:
                self.api_key_var.set(clean_keys[0])
                count = len(clean_keys)
                self.manage_keys_btn.configure(text=f"Manage Pool ({count} active)")
                self.log(f"Configured {count} Gemini API key(s) in round-robin pool.", "ok")
            else:
                self.api_key_var.set("")
                self.manage_keys_btn.configure(text="Manage API Pool...")
                self.log("Cleared all API keys from pool.", "info")

            dialog.destroy()

        save_btn = ttk.Button(
            bottom_bar,
            text="Save & Apply Pool",
            command=save_and_close,
            style="Primary.TButton",
        )
        save_btn.pack(side=tk.RIGHT, padx=(8, 0))

        cancel_btn = ttk.Button(
            bottom_bar,
            text="Cancel",
            command=dialog.destroy,
            style="Secondary.TButton",
        )
        cancel_btn.pack(side=tk.RIGHT)

    def _paste_api_key(self):
        try:
            clip = self.root.clipboard_get().strip()
            if clip:
                first_line = clip.splitlines()[0].strip()
                self.api_key_var.set(first_line)
                if first_line and first_line not in self.api_keys_list:
                    self.api_keys_list.insert(0, first_line)
                    save_api_keys(self.api_keys_list)
                    self.manage_keys_btn.configure(text=f"Manage Pool ({len(self.api_keys_list)} active)")
                self.log("Pasted API key from clipboard.", "ok")
            else:
                messagebox.showinfo("Clipboard Empty", "The clipboard does not contain any text.")
        except Exception as exc:
            messagebox.showinfo("Clipboard Notice", f"Unable to read clipboard: {exc}")

    def _paste_voice(self):
        try:
            clip = self.root.clipboard_get().strip()
            if clip:
                clean_voice = clip.splitlines()[0].strip()
                if clean_voice:
                    self.voice_var.set(clean_voice)
                    self._on_panel_voice_change()
                    self.log(f"Pasted Voice '{clean_voice}' from clipboard.", "ok")
            else:
                messagebox.showinfo("Clipboard Empty", "The clipboard does not contain any text.")
        except Exception as exc:
            messagebox.showinfo("Clipboard Notice", f"Unable to read clipboard: {exc}")

    def _paste_style(self):
        try:
            clip = self.root.clipboard_get().strip()
            if clip:
                clean_style = " ".join(clip.splitlines()).strip()
                if clean_style:
                    self.style_var.set(clean_style)
                    self._on_panel_style_change()
                    self.log(f"Pasted Style '{clean_style}' from clipboard.", "ok")
            else:
                messagebox.showinfo("Clipboard Empty", "The clipboard does not contain any text.")
        except Exception as exc:
            messagebox.showinfo("Clipboard Notice", f"Unable to read clipboard: {exc}")

    def _copy_all_logs(self):
        """Copy selected log text or entire log to clipboard."""
        try:
            try:
                sel = self.log_text.selection_get()
            except Exception:
                sel = ""
            if not sel:
                sel = self.log_text.get("1.0", tk.END).strip()
            if sel:
                self.root.clipboard_clear()
                self.root.clipboard_append(sel)
                self.log("Copied activity log to clipboard.", "ok")
            else:
                messagebox.showinfo("Log Empty", "The activity log is currently empty.")
        except Exception as exc:
            messagebox.showinfo("Clipboard Notice", f"Unable to copy log: {exc}")

    def _clear_logs(self):
        """Clear all contents from the activity log widget."""
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

    def _attach_log_clipboard(self, widget: ScrolledText):
        """Attach right-click context menu and universal copy/select shortcuts to log widget."""
        menu = tk.Menu(
            widget,
            tearoff=0,
            bg=COLOR_INPUT_BG,
            fg=COLOR_TEXT_PRIMARY,
            activebackground=COLOR_ACCENT_BLUE,
            activeforeground="#ffffff",
            bd=1,
            relief="solid",
        )

        def do_copy():
            try:
                try:
                    sel = widget.selection_get()
                except Exception:
                    sel = ""
                if not sel:
                    sel = widget.get("1.0", tk.END).strip()
                if sel:
                    self.root.clipboard_clear()
                    self.root.clipboard_append(sel)
            except Exception:
                pass

        def do_select_all():
            widget.tag_add(tk.SEL, "1.0", tk.END)

        def do_clear():
            self._clear_logs()

        menu.add_command(label="Copy", command=do_copy)
        menu.add_command(label="Select All", command=do_select_all)
        menu.add_separator()
        menu.add_command(label="Clear Log", command=do_clear)

        def show_popup(event):
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        widget.bind("<Button-3>", show_popup)

        def on_ctrl_key(event):
            if event.state & 4:
                # 67 = 'C', 65 = 'A' across all keyboard layouts
                if event.keycode == 67:
                    do_copy()
                    return "break"
                elif event.keycode == 65:
                    do_select_all()
                    return "break"

        widget.bind("<Control-KeyPress>", on_ctrl_key)

    def _attach_universal_clipboard(self, widget: tk.Widget, on_change=None):
        """Attach right-click context menu and keyboard-layout-independent shortcuts."""
        menu = tk.Menu(
            widget,
            tearoff=0,
            bg=COLOR_INPUT_BG,
            fg=COLOR_TEXT_PRIMARY,
            activebackground=COLOR_ACCENT_BLUE,
            activeforeground="#ffffff",
            bd=1,
            relief="solid",
        )

        def do_paste(event=None):
            try:
                clip = self.root.clipboard_get()
                if clip is not None:
                    clean_clip = clip.strip("\r\n")
                    try:
                        widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
                    except tk.TclError:
                        pass
                    widget.insert(tk.INSERT, clean_clip)
                    if on_change:
                        on_change()
            except Exception:
                pass
            return "break"

        def do_copy(event=None):
            try:
                sel = widget.selection_get()
                self.root.clipboard_clear()
                self.root.clipboard_append(sel)
            except Exception:
                pass
            return "break"

        def do_cut(event=None):
            try:
                sel = widget.selection_get()
                self.root.clipboard_clear()
                self.root.clipboard_append(sel)
                widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
                if on_change:
                    on_change()
            except Exception:
                pass
            return "break"

        def do_select_all(event=None):
            try:
                widget.select_range(0, tk.END)
                widget.icursor(tk.END)
            except Exception:
                pass
            return "break"

        def do_clear(event=None):
            try:
                widget.delete(0, tk.END)
                if on_change:
                    on_change()
            except Exception:
                pass
            return "break"

        menu.add_command(label="Paste", command=do_paste)
        menu.add_command(label="Copy", command=do_copy)
        menu.add_command(label="Cut", command=do_cut)
        menu.add_separator()
        menu.add_command(label="Select All", command=do_select_all)
        menu.add_command(label="Clear", command=do_clear)

        def show_popup(event):
            try:
                widget.focus_set()
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        widget.bind("<Button-3>", show_popup)

        widget.bind("<<Paste>>", do_paste)
        widget.bind("<<Copy>>", do_copy)
        widget.bind("<<Cut>>", do_cut)
        widget.bind("<<SelectAll>>", do_select_all)

        widget.bind("<Control-v>", do_paste)
        widget.bind("<Control-V>", do_paste)
        widget.bind("<Shift-Insert>", do_paste)
        widget.bind("<Control-c>", do_copy)
        widget.bind("<Control-C>", do_copy)
        widget.bind("<Control-x>", do_cut)
        widget.bind("<Control-X>", do_cut)
        widget.bind("<Control-a>", do_select_all)
        widget.bind("<Control-A>", do_select_all)

        def on_ctrl_key(event):
            if event.state & 4:
                k = getattr(event, "keysym", "").lower()
                code = getattr(event, "keycode", 0)
                # 86 = 'V', 67 = 'C', 88 = 'X', 65 = 'A' across Windows keyboards
                if code == 86 or k == "v":
                    return do_paste(event)
                elif code == 67 or k == "c":
                    return do_copy(event)
                elif code == 88 or k == "x":
                    return do_cut(event)
                elif code == 65 or k == "a":
                    return do_select_all(event)

        widget.bind("<Control-KeyPress>", on_ctrl_key)

    def _toggle_show_key(self):
        self.api_entry.configure(show="" if self.show_key_var.get() else "*")

    def _on_tree_select(self, _event=None):
        selected = self.tree.selection()
        if len(selected) >= 1:
            first_item = selected[0]
            voice_val = self.tree.set(first_item, "voice")
            style_val = self.tree.set(first_item, "style")
            self._is_programmatic_update = True
            try:
                if voice_val:
                    self.voice_var.set(voice_val)
                    self.last_applied_voice = voice_val
                if style_val:
                    self.style_var.set(style_val)
                    self.last_applied_style = style_val
            finally:
                self._is_programmatic_update = False

    def _add_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Select Text or Word Files",
            filetypes=[
                ("Text and Word files (*.txt *.docx)", "*.txt *.docx"),
                ("Text files (*.txt)", "*.txt"),
                ("Word documents (*.docx)", "*.docx"),
                ("All files (*.*)", "*.*"),
            ],
        )
        if not file_paths:
            return

        existing_paths = {
            os.path.normcase(os.path.abspath(self.tree.set(item, "filepath")))
            for item in self.tree.get_children()
        }

        added_count = 0
        skipped_unsupported: list[str] = []

        for path in file_paths:
            abs_path = os.path.abspath(path)
            base_name = os.path.basename(abs_path)

            if base_name.startswith("~$"):
                continue

            lower_ext = os.path.splitext(abs_path)[1].lower()
            if lower_ext not in (".txt", ".docx"):
                skipped_unsupported.append(base_name)
                continue

            norm_path = os.path.normcase(abs_path)
            if norm_path in existing_paths:
                continue

            existing_paths.add(norm_path)
            self.tree.insert(
                "",
                tk.END,
                values=(abs_path, self.last_applied_voice, self.last_applied_style, "Waiting"),
            )
            added_count += 1

        if added_count > 0:
            self.log(f"Added {added_count} file(s) to conversion queue.", "info")

        if skipped_unsupported:
            display_list = skipped_unsupported[:10]
            extra = f"\n...and {len(skipped_unsupported) - 10} more." if len(skipped_unsupported) > 10 else ""
            msg = (
                f"Skipped {len(skipped_unsupported)} unsupported file(s). "
                "Only .txt and .docx formats are supported:\n\n"
                + "\n".join(display_list)
                + extra
            )
            messagebox.showwarning("Unsupported Files Skipped", msg)

    def _remove_selected(self):
        if self.is_processing:
            messagebox.showinfo("Busy", "Cannot remove files while batch processing is active.")
            return
        selected = self.tree.selection()
        if not selected:
            return
        count = len(selected)
        for item in selected:
            self.tree.delete(item)
        self.log(f"Removed {count} file(s) from queue.", "info")

    def _clear_list(self):
        if self.is_processing:
            messagebox.showinfo("Busy", "Cannot clear files while batch processing is active.")
            return
        items = self.tree.get_children()
        if not items:
            return
        for item in items:
            self.tree.delete(item)
        self.log("Cleared all files from conversion queue.", "info")

    def _apply_settings(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Selection Required", "Please select one or more files in the list to apply settings.")
            return
        voice_val = self.voice_var.get().strip() or VOICES[0]
        style_val = self.style_var.get().strip() or DEFAULT_STYLE_LABEL

        for item in selected:
            self.tree.set(item, "voice", voice_val)
            self.tree.set(item, "style", style_val)

        self.last_applied_voice = voice_val
        self.last_applied_style = style_val
        self.log(f"Applied settings to {len(selected)} selected file(s): Voice={voice_val}, Style={style_val}", "info")

    def _apply_to_all(self):
        items = self.tree.get_children()
        if not items:
            messagebox.showinfo("Empty List", "No files in the list to apply settings to.")
            return
        voice_val = self.voice_var.get().strip() or VOICES[0]
        style_val = self.style_var.get().strip() or DEFAULT_STYLE_LABEL

        for item in items:
            self.tree.set(item, "voice", voice_val)
            self.tree.set(item, "style", style_val)

        self.last_applied_voice = voice_val
        self.last_applied_style = style_val
        self.log(f"Applied settings to all {len(items)} file(s): Voice={voice_val}, Style={style_val}", "info")

    def _start_batch(self):
        if self.is_processing:
            return

        pool_keys = [k for k in self.api_keys_list if k.strip()]
        is_mock = os.environ.get("TTS_MOCK") == "1"
        if not pool_keys and not is_mock:
            open_now = messagebox.askyesno(
                "API Keys Required",
                "No Gemini API keys have been configured in the pool.\n\nWould you like to open the API Key Pool manager now to configure your key(s)?",
            )
            if open_now:
                self._open_api_manager()
            return

        items = self.tree.get_children()
        if not items:
            messagebox.showwarning("Empty List", "Please add at least one .txt or .docx file to convert.")
            return

        file_records: list[tuple[str, str, str, str]] = []
        for item in items:
            f_path = self.tree.set(item, "filepath")
            f_voice = self.tree.set(item, "voice")
            f_style = self.tree.set(item, "style")
            if not os.path.isfile(f_path):
                messagebox.showerror("File Not Found", f"File does not exist:\n{f_path}\n\nPlease check or remove it.")
                return
            file_records.append((item, f_path, f_voice, f_style))

        self.is_processing = True
        self.stop_event.clear()
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.add_btn.configure(state="disabled")
        self.remove_btn.configure(state="disabled")
        self.clear_btn.configure(state="disabled")
        self.apply_btn.configure(state="disabled")
        self.apply_all_btn.configure(state="disabled")
        self.voice_paste_btn.configure(state="disabled")
        self.style_paste_btn.configure(state="disabled")
        self.manage_keys_btn.configure(state="disabled")

        self.progress_bar.configure(value=0, maximum=100)
        self.progress_label.configure(text="Starting...")
        self.log(f"Starting batch conversion of {len(file_records)} file(s)...", "info")

        pool = ApiKeyPool(pool_keys)

        worker = threading.Thread(
            target=self._worker_thread,
            args=(file_records, pool),
            daemon=True,
        )
        worker.start()

    def _stop_batch(self):
        if not self.is_processing:
            return
        self.stop_event.set()
        self.stop_btn.configure(state="disabled")
        self.log("Stop requested by user. Halting gracefully...", "warn")
        self.progress_label.configure(text="Stopping...")

    def _worker_thread(self, file_records: list[tuple[str, str, str, str]], pool: ApiKeyPool):
        """Sequential background processing worker with non-spoken styling, unique output naming, and caching."""
        try:
            self._worker_thread_internal(file_records, pool)
        except Exception as fatal_exc:
            import traceback
            err_trace = traceback.format_exc()
            self.queue.put(("log", f"FATAL WORKER ERROR: {fatal_exc}", "err"))
            self.queue.put(("log", f"{err_trace}", "err"))
            self.queue.put(("batch_finished", 0, len(file_records), 0, True))

    def _worker_thread_internal(self, file_records: list[tuple[str, str, str, str]], pool: ApiKeyPool):
        is_mock = os.environ.get("TTS_MOCK") == "1"

        if is_mock:
            self.queue.put(("log", "[MOCK MODE] Network calls bypassed. Generating silence per chunk.", "warn"))
        else:
            key_count = pool.total_count()
            self.queue.put(("log", f"Initialized Gemini API key pool with {key_count} active key(s).", "info"))
            self.queue.put((
                "log",
                f"Primary model: '{PRIMARY_MODEL}' | Fallback model: '{FALLBACK_MODEL}'.",
                "ok",
            ))

        # Pre-extract, inspect and chunk all files, assigning collision-free output paths
        prepared_files: list[dict] = []
        reserved_outputs: set[str] = set()
        total_chunks = 0

        for item_id, file_path, voice, style in file_records:
            base_name = os.path.basename(file_path)
            ext = os.path.splitext(base_name)[1].lower()
            self.queue.put(("log", f"Inspecting and extracting text from '{base_name}' ({ext})...", "info"))

            base_output_mp3 = os.path.splitext(file_path)[0] + ".mp3"
            unique_output_mp3 = get_unique_output_path(base_output_mp3, reserved_outputs)
            reserved_outputs.add(os.path.normcase(os.path.abspath(unique_output_mp3)))

            if os.path.normcase(os.path.abspath(unique_output_mp3)) != os.path.normcase(os.path.abspath(base_output_mp3)):
                self.queue.put((
                    "log",
                    f"  Existing file detected for '{os.path.basename(base_output_mp3)}'. Auto-assigned unique output: '{os.path.basename(unique_output_mp3)}'.",
                    "warn",
                ))

            try:
                text = extract_text(file_path)
                char_count = len(text)
                self.queue.put(("log", f"  Extracted {char_count:,} characters from '{base_name}'.", "sub"))

                chunks = split_text_into_chunks(text, MAX_CHARS)
                if not chunks:
                    raise ValueError("no text found")

                total_chunks += len(chunks)
                self.queue.put(("log", f"  Chunked into {len(chunks)} piece(s) (limit: {MAX_CHARS} chars/chunk):", "sub"))
                for c_idx, c in enumerate(chunks, start=1):
                    preview = c[:45].replace("\n", " ") + ("..." if len(c) > 45 else "")
                    self.queue.put(("log", f"    * Chunk {c_idx}/{len(chunks)}: {len(c):,} chars -> \"{preview}\"", "sub"))

                prepared_files.append({
                    "item_id": item_id,
                    "file_path": file_path,
                    "output_mp3": unique_output_mp3,
                    "voice": voice,
                    "style": style,
                    "chunks": chunks,
                    "error": None,
                })
            except Exception as exc:
                short_err = str(exc)
                self.queue.put(("log", f"  Extraction error on '{base_name}': {short_err}", "err"))
                prepared_files.append({
                    "item_id": item_id,
                    "file_path": file_path,
                    "output_mp3": unique_output_mp3,
                    "voice": voice,
                    "style": style,
                    "chunks": [],
                    "error": short_err,
                })

        self.queue.put(("progress_init", total_chunks))
        self.queue.put(("log", f"Total chunks across batch: {total_chunks}.", "info"))

        completed_chunks = 0
        batch_halted = False
        batch_halt_reason = ""
        last_request_time = 0.0
        min_request_interval = 0.0
        files_succeeded = 0
        files_failed = 0
        files_stopped = 0

        for file_info in prepared_files:
            item_id = file_info["item_id"]
            file_path = file_info["file_path"]
            output_mp3 = file_info["output_mp3"]
            voice = file_info["voice"]
            style = file_info["style"]
            chunks = file_info["chunks"]
            preload_error = file_info["error"]
            base_name = os.path.basename(file_path)

            if self.stop_event.is_set():
                files_stopped += 1
                self.queue.put(("item_status", item_id, "Stopped"))
                continue

            if batch_halted:
                files_stopped += 1
                self.queue.put(("item_status", item_id, f"Stopped: {batch_halt_reason}"))
                continue

            if preload_error:
                files_failed += 1
                self.queue.put(("item_status", item_id, f"Error: {preload_error}"))
                self.queue.put(("log", f"Skipping '{base_name}' due to error: {preload_error}", "err"))
                continue

            pcm_accumulator = bytearray()
            detected_sample_rate = DEFAULT_SAMPLE_RATE
            num_chunks = len(chunks)
            file_failed = False
            file_err_info: ErrorInfo | None = None
            resolved_style = resolve_style_instruction(style)
            clean_style = clean_style_instruction(resolved_style)
            file_cache_dir = get_file_cache_dir(file_path, voice, clean_style)

            style_desc = f"'{clean_style}'" if clean_style else "Neutral"
            self.queue.put(("log", f"Processing '{base_name}' ({num_chunks} chunks, Voice: '{voice}', Style: {style_desc})...", "info"))
            chunks_completed_for_file = 0

            for idx, chunk in enumerate(chunks, start=1):
                if self.stop_event.is_set():
                    self.queue.put(("item_status", item_id, f"Stopped ({chunks_completed_for_file}/{num_chunks} chunks)"))
                    self.queue.put(("log", f"Stopped while processing '{base_name}' ({chunks_completed_for_file}/{num_chunks} chunks completed).", "warn"))
                    break

                self.queue.put(("item_status", item_id, f"Processing (chunk {idx}/{num_chunks})"))

                cache_file = os.path.join(file_cache_dir, f"{idx}.pcm")

                # Check if this chunk is already cached from an earlier stopped run
                if os.path.isfile(cache_file) and os.path.getsize(cache_file) > 0:
                    try:
                        with open(cache_file, "rb") as cf:
                            chunk_pcm = cf.read()
                        chunk_rate = DEFAULT_SAMPLE_RATE
                        err_info = None
                        self.queue.put(("log", f"  Chunk {idx}/{num_chunks} loaded from local cache (resumed).", "ok"))
                    except Exception:
                        chunk_pcm = b""
                        chunk_rate = DEFAULT_SAMPLE_RATE
                        err_info = None
                else:
                    chunk_pcm = b""

                if not chunk_pcm:
                    # Enforce inter-request rate limit spacing if a 429 was seen
                    if min_request_interval > 0.0 and not is_mock:
                        elapsed = time.time() - last_request_time
                        if elapsed < min_request_interval:
                            wait_gap = min_request_interval - elapsed
                            gap_slept = 0.0
                            while gap_slept < wait_gap:
                                if self.stop_event.is_set():
                                    break
                                step = min(0.2, wait_gap - gap_slept)
                                time.sleep(step)
                                gap_slept += step

                    if self.stop_event.is_set():
                        self.queue.put(("item_status", item_id, f"Stopped ({chunks_completed_for_file}/{num_chunks} chunks)"))
                        break

                    last_request_time = time.time()
                    chunk_pcm, chunk_rate, err_info, was_rate_limited, active_model = self._fetch_audio_chunk_with_retry(
                        client=None,
                        chunk_text=chunk,
                        voice=voice,
                        clean_style=clean_style,
                        is_mock=is_mock,
                        chunk_idx=idx,
                        total_chunks=num_chunks,
                        active_model=pool.get_active_model(),
                        pool=pool,
                    )

                    if was_rate_limited:
                        min_request_interval = MIN_RATE_LIMIT_INTERVAL

                    if err_info:
                        file_failed = True
                        file_err_info = err_info
                        # Check halting kinds: auth, region, model_not_found, quota_daily
                        if err_info.kind in ("auth", "region", "model_not_found", "quota_daily"):
                            batch_halted = True
                            batch_halt_reason = err_info.kind
                        break

                    # Save chunk to cache
                    try:
                        with open(cache_file, "wb") as cf:
                            cf.write(chunk_pcm)
                    except Exception:
                        pass

                # Add 0.3s silence gap between chunks of the same file
                if pcm_accumulator:
                    gap_silence = b"\x00" * int(chunk_rate * 2 * CHUNK_GAP_SECONDS)
                    pcm_accumulator.extend(gap_silence)

                pcm_accumulator.extend(chunk_pcm)
                if chunk_rate:
                    detected_sample_rate = chunk_rate

                chunks_completed_for_file += 1
                completed_chunks += 1
                self.queue.put(("progress_update", completed_chunks, total_chunks))

            # Handle Stop / Failure / Success for this file
            if file_failed and file_err_info:
                files_failed += 1
                short_status = f"Error: {file_err_info.kind}"
                self.queue.put(("item_status", item_id, short_status))
                self.queue.put(("log", f"Failed '{base_name}': {file_err_info.user_message}", "err"))
                if batch_halted:
                    self.queue.put(("log", f"Critical error ({file_err_info.kind}). Halting remaining queue.", "err"))
                continue

            if self.stop_event.is_set():
                if chunks_completed_for_file < num_chunks:
                    files_stopped += 1
                    self.queue.put(("item_status", item_id, f"Stopped ({chunks_completed_for_file}/{num_chunks} chunks)"))
                    continue
                # If all chunks finished right as stop was clicked, still save the MP3!

            # All chunks for this file succeeded -> encode to MP3
            try:
                self.queue.put((
                    "log",
                    f"Encoding {len(pcm_accumulator):,} total PCM bytes into MP3 with lameenc (64 kbps mono)...",
                    "info",
                ))
                save_pcm_as_mp3(bytes(pcm_accumulator), detected_sample_rate, output_mp3)
                file_size_kb = os.path.getsize(output_mp3) / 1024.0
                files_succeeded += 1
                self.queue.put(("item_status", item_id, "Done"))
                self.queue.put(("log", f"Saved MP3: '{os.path.basename(output_mp3)}' ({file_size_kb:.1f} KB).", "ok"))

                # Clear this file's chunk cache on successful MP3 creation
                shutil.rmtree(file_cache_dir, ignore_errors=True)
            except Exception as save_exc:
                files_failed += 1
                self.queue.put(("item_status", item_id, "Error: save failed"))
                self.queue.put(("log", f"Failed to save MP3 for '{base_name}': {save_exc}", "err"))

        was_halted_or_stopped = self.stop_event.is_set() or batch_halted
        self.queue.put(("batch_finished", files_succeeded, files_failed, files_stopped, was_halted_or_stopped))

    def _fetch_audio_chunk_with_retry(
        self,
        client: genai.Client | None,
        chunk_text: str,
        voice: str,
        clean_style: str,
        is_mock: bool,
        chunk_idx: int,
        total_chunks: int,
        active_model: str = PRIMARY_MODEL,
        pool: ApiKeyPool | None = None,
    ) -> tuple[bytes, int, ErrorInfo | None, bool, str]:
        """Fetch TTS audio for a single chunk with round-robin pool rotation, vocal styling metadata, retry handling, and automatic model fallback."""
        if pool is None:
            pool = ApiKeyPool(["standalone_key"])
            if client is not None:
                pool._clients["standalone_key"] = client
            pool.active_model = active_model

        rate_limit_sequence = [20.0, 40.0, 60.0, 60.0]
        transient_sequence = [2.0, 4.0]
        max_rate_limit_attempts = 5
        max_transient_attempts = 3

        rate_limit_attempt = 0
        transient_attempt = 0
        saw_rate_limit = False

        while True:
            if self.stop_event.is_set():
                return b"", DEFAULT_SAMPLE_RATE, ErrorInfo("stopped", "Operation stopped by user"), saw_rate_limit, pool.get_active_model()

            candidate = pool.get_next_key_and_model()
            if candidate is None:
                return b"", DEFAULT_SAMPLE_RATE, ErrorInfo("quota_daily", "Daily quota exhausted across all configured API keys on both primary and fallback models."), saw_rate_limit, pool.get_active_model()

            current_api_key, current_model, key_idx = candidate
            key_tag = f"Key #{key_idx}" if pool.total_count() > 1 else "API Key"

            if is_mock:
                time.sleep(0.08)
                self.active_model = current_model
                return b"\x00" * (DEFAULT_SAMPLE_RATE * 2), DEFAULT_SAMPLE_RATE, None, False, current_model

            self.queue.put(("log", f"Sending chunk {chunk_idx}/{total_chunks} ({len(chunk_text):,} chars) to Gemini TTS using {key_tag} [{current_model}]...", "info"))
            if clean_style:
                self.queue.put(("log", f"  Applied vocal style metadata: \"{clean_style}\"", "sub"))

            try:
                active_client = pool.get_client(current_api_key)
                speech_metadata = types.SpeechMetadata(style=clean_style) if clean_style else None
                part = types.Part(
                    text=chunk_text,
                    speech_metadata=speech_metadata,
                )
                contents = [
                    types.Content(
                        role="user",
                        parts=[part],
                    )
                ]
                config = types.GenerateContentConfig(
                    temperature=1,
                    response_modalities=["audio"],
                    speech_config=types.SpeechConfig(
                        voice_config=types.VoiceConfig(
                            prebuilt_voice_config=types.PrebuiltVoiceConfig(voice_name=voice)
                        )
                    ),
                )

                pcm_chunks: list[bytes] = []
                detected_rate = DEFAULT_SAMPLE_RATE

                for resp_chunk in active_client.models.generate_content_stream(
                    model=current_model, contents=contents, config=config
                ):
                    if self.stop_event.is_set():
                        return b"", DEFAULT_SAMPLE_RATE, ErrorInfo("stopped", "Operation stopped by user"), saw_rate_limit, current_model

                    for p in (resp_chunk.parts or []):
                        if p.inline_data and p.inline_data.data:
                            raw = p.inline_data.data
                            if isinstance(raw, str):
                                raw = base64.b64decode(raw)
                            pcm_chunks.append(raw)
                            if p.inline_data.mime_type:
                                detected_rate = parse_sample_rate(p.inline_data.mime_type, detected_rate)

                full_chunk_pcm = b"".join(pcm_chunks)
                if not full_chunk_pcm:
                    raise RuntimeError("No audio data received from Gemini TTS API")

                self.active_model = current_model
                return full_chunk_pcm, detected_rate, None, saw_rate_limit, current_model

            except Exception as exc:
                if self.stop_event.is_set():
                    return b"", DEFAULT_SAMPLE_RATE, ErrorInfo("stopped", "Operation stopped by user"), saw_rate_limit, current_model

                err_info = classify_error(exc)

                # Automatic fallback from PRIMARY_MODEL to FALLBACK_MODEL on daily quota exhaustion or model unavailability
                if err_info.kind == "quota_daily":
                    all_exhausted, notice = pool.mark_quota_daily(current_api_key, current_model)
                    self.queue.put(("log", notice, "warn"))
                    if all_exhausted and current_model == PRIMARY_MODEL:
                        self.queue.put((
                            "log",
                            f"Daily quota reached on all keys for '{PRIMARY_MODEL}'. Automatically shifting to fallback model '{FALLBACK_MODEL}' across all keys...",
                            "warn",
                        ))
                    self.active_model = pool.get_active_model()
                    if pool.is_all_exhausted():
                        return b"", DEFAULT_SAMPLE_RATE, ErrorInfo("quota_daily", "Daily quota exhausted across all configured API keys on both primary and fallback models."), saw_rate_limit, pool.get_active_model()
                    time.sleep(0.3)
                    continue

                if err_info.kind == "model_not_found":
                    if current_model == PRIMARY_MODEL:
                        self.queue.put((
                            "log",
                            f"Primary model '{PRIMARY_MODEL}' not found. Shifting all keys to fallback model '{FALLBACK_MODEL}'...",
                            "warn",
                        ))
                        pool.shift_to_fallback()
                        self.active_model = FALLBACK_MODEL
                        time.sleep(0.3)
                        continue
                    else:
                        return b"", DEFAULT_SAMPLE_RATE, err_info, saw_rate_limit, current_model

                if err_info.kind == "auth":
                    pool.mark_invalid(current_api_key)
                    self.queue.put(("log", f"{key_tag} authorization failed ({err_info.user_message}). Marking key as inactive.", "warn"))
                    if pool.valid_count() == 0:
                        return b"", DEFAULT_SAMPLE_RATE, ErrorInfo("auth", f"All configured API keys failed authorization: {err_info.user_message}"), saw_rate_limit, current_model
                    time.sleep(0.3)
                    continue

                # Halt immediately on region
                if err_info.kind == "region":
                    return b"", DEFAULT_SAMPLE_RATE, err_info, saw_rate_limit, current_model

                # Fail-fast without retry on client 400 bad_request (e.g. invalid voice name)
                if err_info.kind == "bad_request":
                    return b"", DEFAULT_SAMPLE_RATE, err_info, saw_rate_limit, current_model

                # Handle 429 rate_limit
                if err_info.kind == "rate_limit":
                    saw_rate_limit = True
                    rate_limit_attempt += 1
                    if rate_limit_attempt >= max_rate_limit_attempts:
                        return b"", DEFAULT_SAMPLE_RATE, ErrorInfo("rate_limit", f"Rate limit exceeded after {max_rate_limit_attempts} attempts: {exc}"), True, current_model

                    delay = extract_retry_delay(exc, rate_limit_attempt - 1, rate_limit_sequence)
                    self.queue.put(("log", f"Rate limit (429) on chunk {chunk_idx}/{total_chunks} [{current_model}]. Waiting {delay:.0f}s before retry (attempt {rate_limit_attempt + 1}/{max_rate_limit_attempts})...", "warn"))

                    slept = 0.0
                    while slept < delay:
                        if self.stop_event.is_set():
                            return b"", DEFAULT_SAMPLE_RATE, ErrorInfo("stopped", "Operation stopped by user"), True, current_model
                        step = min(0.2, delay - slept)
                        time.sleep(step)
                        slept += step
                    continue

                # Handle transient 5xx or connection errors
                if err_info.kind == "transient":
                    transient_attempt += 1
                    if transient_attempt >= max_transient_attempts:
                        return b"", DEFAULT_SAMPLE_RATE, ErrorInfo("transient", f"Transient error failed after {max_transient_attempts} attempts: {exc}"), saw_rate_limit, current_model

                    delay = transient_sequence[min(transient_attempt - 1, len(transient_sequence) - 1)]
                    self.queue.put(("log", f"Transient error on chunk {chunk_idx} [{current_model}]: {exc}. Retrying in {delay:.0f}s (attempt {transient_attempt + 1}/{max_transient_attempts})...", "warn"))

                    slept = 0.0
                    while slept < delay:
                        if self.stop_event.is_set():
                            return b"", DEFAULT_SAMPLE_RATE, ErrorInfo("stopped", "Operation stopped by user"), saw_rate_limit, current_model
                        step = min(0.2, delay - slept)
                        time.sleep(step)
                        slept += step
                    continue

                # Any unclassified error: abort current file
                return b"", DEFAULT_SAMPLE_RATE, err_info, saw_rate_limit, current_model

    # ------------------------------------------------------------------------
    # Queue Processing & UI Dispatching
    # ------------------------------------------------------------------------

    def _process_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                msg_type = msg[0]
                args = msg[1:]
                if msg_type == "log":
                    text = args[0]
                    level = args[1] if len(args) > 1 else "info"
                    self.log(text, level)
                elif msg_type == "item_status":
                    item_id, status_val = args[0], args[1]
                    if self.tree.exists(item_id):
                        self.tree.set(item_id, "status", status_val)
                elif msg_type == "progress_init":
                    total = args[0]
                    self.progress_bar.configure(maximum=max(total, 1), value=0)
                    self.progress_label.configure(text=f"0/{total} chunks (0%)")
                elif msg_type == "progress_update":
                    current, total = args[0], args[1]
                    self.progress_bar.configure(value=current)
                    pct = int((current / total) * 100) if total > 0 else 0
                    self.progress_label.configure(text=f"{current}/{total} chunks ({pct}%)")
                elif msg_type == "batch_finished":
                    succ, fail, stop_cnt, was_stopped = args[0], args[1], args[2], args[3]
                    self._on_batch_finished(succ, fail, stop_cnt, was_stopped)
                elif msg_type == "batch_done":
                    self._on_batch_finished(1, 0, 0, False)
                elif msg_type == "batch_stopped":
                    self._on_batch_finished(0, 0, 1, True)
                elif msg_type == "auth_error":
                    self.log(f"ERROR: {args[0]}", "err")
                    messagebox.showerror("Authentication Error", args[0])
                    self._on_batch_finished(0, 1, 0, True)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._process_queue)

    def _on_batch_finished(self, succ: int, fail: int, stopped: int, was_stopped: bool):
        self.is_processing = False
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.add_btn.configure(state="normal")
        self.remove_btn.configure(state="normal")
        self.clear_btn.configure(state="normal")
        self.apply_btn.configure(state="normal")
        self.apply_all_btn.configure(state="normal")
        self.voice_paste_btn.configure(state="normal")
        self.style_paste_btn.configure(state="normal")
        self.manage_keys_btn.configure(state="normal")

        if was_stopped:
            self.progress_label.configure(text="Stopped")
            if fail > 0:
                self.log(f"Batch stopped with errors: {succ} succeeded, {fail} failed, {stopped} stopped.", "err")
            else:
                self.log(f"Batch processing stopped. ({succ} completed)", "warn")
        elif fail > 0:
            if succ > 0:
                self.progress_label.configure(text=f"Done ({fail} failed)")
                self.log(f"Batch completed with errors: {succ} file(s) succeeded, {fail} file(s) failed.", "warn")
            else:
                self.progress_label.configure(text="Failed")
                self.log(f"Batch failed: all {fail} file(s) encountered errors.", "err")
        else:
            self.progress_label.configure(text="Completed")
            self.log(f"All {succ} file(s) in batch completed successfully!", "ok")

    def log(self, message: str, level: str = "info"):
        """Append a message to the activity log text widget with syntax colors."""
        timestamp = time.strftime("[%H:%M:%S] ")
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, timestamp, "time")

        tag = level
        if tag not in ("info", "ok", "err", "warn", "sub"):
            tag = "info"

        self.log_text.insert(tk.END, message + "\n", tag)
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def _on_close(self):
        if self.api_keys_list:
            save_api_keys(self.api_keys_list)
        self.root.destroy()


# ============================================================================
# Self-Test Mode (--selftest CLI)
# ============================================================================

def run_selftest() -> int:
    """Run an automated self-test verifying DPAPI, audio encoder, chunking, and error classifier."""
    log_path = os.path.join(tempfile.gettempdir(), "tts_selftest.log")
    messages = []
    success = True

    def record(title: str, passed: bool, detail: str = ""):
        nonlocal success
        status_tag = "[PASS]" if passed else "[FAIL]"
        line = f"{status_tag} {title}" + (f": {detail}" if detail else "")
        messages.append(line)
        if not passed:
            success = False

    try:
        test_plain = b"TextToSpeechBatch_DPAPI_SelfTest_Payload"
        enc = protect_data(test_plain)
        dec = unprotect_data(enc)
        record("DPAPI protect & unprotect", dec == test_plain)
    except Exception as exc:
        record("DPAPI protect & unprotect", False, str(exc))

    try:
        encoder = lameenc.Encoder()
        encoder.set_bit_rate(MP3_BITRATE_KBPS)
        encoder.set_in_sample_rate(DEFAULT_SAMPLE_RATE)
        encoder.set_channels(1)
        encoder.set_out_sample_rate(DEFAULT_SAMPLE_RATE)
        encoder.set_quality(MP3_QUALITY)
        test_pcm = b"\x00" * 4800  # 0.1s silence
        mp3_out = encoder.encode(test_pcm) + encoder.flush()
        record("lameenc MP3 encoder initialization & flush", len(mp3_out) > 0)
    except Exception as exc:
        record("lameenc MP3 encoder initialization & flush", False, str(exc))

    try:
        sample_text = "Profit was 46.5 percent at 10:30 today."
        chunks = split_text_into_chunks(sample_text)
        record("Exact substring chunking preservation", len(chunks) == 1 and chunks[0] == sample_text)
    except Exception as exc:
        record("Exact substring chunking preservation", False, str(exc))

    try:
        from google.genai.errors import APIError
        err = APIError(401, {"error": {"code": 401, "message": "API_KEY_INVALID"}})
        info = classify_error(err)
        record("Structured error classification (401 auth)", info.kind == "auth")
    except Exception as exc:
        record("Structured error classification (401 auth)", False, str(exc))

    try:
        # Verify output renaming logic
        out_p = get_unique_output_path("test.mp3", {"test.mp3"})
        record("Output collision avoidance logic", out_p.endswith("test1.mp3"))
    except Exception as exc:
        record("Output collision avoidance logic", False, str(exc))

    try:
        pool = ApiKeyPool(["keyA", "keyB"])
        c1 = pool.get_next_key_and_model()
        c2 = pool.get_next_key_and_model()
        pool.mark_quota_daily("keyA", PRIMARY_MODEL)
        pool.mark_quota_daily("keyB", PRIMARY_MODEL)
        c3 = pool.get_next_key_and_model()
        record(
            "ApiKeyPool rotation and fallback shift",
            c1[0] == "keyA" and c2[0] == "keyB" and c3[1] == FALLBACK_MODEL,
        )
    except Exception as exc:
        record("ApiKeyPool rotation and fallback shift", False, str(exc))

    try:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(messages) + "\n")
            f.write(f"Result: {'SUCCESS' if success else 'FAILURE'}\n")
    except Exception:
        pass

    print("\n".join(messages))
    return 0 if success else 1


# ============================================================================
# Entry Point
# ============================================================================

def main():
    if "--selftest" in sys.argv:
        exit_code = run_selftest()
        sys.exit(exit_code)

    root = tk.Tk()
    app = TTSBatchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
