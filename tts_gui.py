"""Text to Speech Batch - Windows desktop application for converting text and Word files to MP3 using Gemini TTS.

All source code, identifiers, comments, and labels are written in English.
Style preset labels in Persian are represented strictly using \\uXXXX escape sequences.
Features a modern, high-contrast dark theme matching Online Subtitle Generator.
"""

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
import traceback
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

MODEL_NAME = "gemini-3.8-flash-tts"
MAX_CHARS = 3000
MP3_BITRATE_KBPS = 64
MP3_QUALITY = 2
DEFAULT_SAMPLE_RATE = 24000

# Color Palette (Matching Online Subtitle Generator / Modern Dark Windows)
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

# Fixed list of supported Gemini TTS voices
VOICES = [
    "Bodi", "Fola", "Sola", "Zephyr", "Puck", "Charon", "Kore", "Fenrir",
    "Leda", "Orus", "Aoede", "Callirrhoe", "Autonoe", "Enceladus", "Iapetus",
    "Umbriel", "Algieba", "Despina", "Erinome", "Algenib", "Rasalgethi",
    "Laomedeia", "Achernar", "Alnilam", "Schedar", "Gacrux", "Pulcherrima",
    "Achird", "Zubenelgenubi", "Vindemiatrix", "Sadachbia", "Sadaltager",
    "Sulafat",
]

# Style Presets: Persian label (escaped as \\uXXXX) -> English model instruction
STYLE_PRESETS = {
    # Without style (neutral) -> no instruction
    "\u0628\u062f\u0648\u0646 \u0627\u0633\u062a\u0627\u06cc\u0644 (\u062e\u0646\u062b\u06cc)": "",
    # Cheerful and energetic
    "\u0634\u0627\u062f \u0648 \u067e\u0631\u0627\u0646\u0631\u0698\u06cc": "Say in a cheerful, upbeat and energetic tone:",
    # Sad and depressed
    "\u063a\u0645\u06af\u06cc\u0646 \u0648 \u0627\u0641\u0633\u0631\u062f\u0647": "Say in a sad, melancholic tone, slowly and softly:",
    # Calm and gentle
    "\u0622\u0631\u0627\u0645 \u0648 \u0645\u0644\u0627\u06cc\u0645": "Say in a calm, gentle and soothing tone:",
    # Excited
    "\u0647\u06cc\u062c\u0627\u0646\u0632\u062f\u0647": "Say in an excited and enthusiastic tone:",
    # Serious and formal
    "\u062c\u062f\u06cc \u0648 \u0631\u0633\u0645\u06cc": "Say in a serious, formal and professional tone:",
    # Friendly and warm
    "\u062f\u0648\u0633\u062a\u0627\u0646\u0647 \u0648 \u06af\u0631\u0645": "Say in a warm, friendly tone with natural micro-pauses:",
    # Narration
    "\u0631\u0648\u0627\u06cc\u062a\u06af\u0631": "Say in a narration style with rich vocal modulation and dramatic pauses:",
    # News anchor
    "\u062e\u0628\u0631\u06cc": "Say like a news anchor, clear, neutral and steady:",
    # Whisper
    "\u0646\u062c\u0648\u0627": "Whisper quietly:",
    # Promotional
    "\u062a\u0628\u0644\u06cc\u063a\u0627\u062a\u06cc": "Say in a high-energy promotional tone with elongated vowels:",
    # Angry
    "\u0639\u0635\u0628\u0627\u0646\u06cc": "Say in an angry, intense tone:",
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


def save_api_key(key: str) -> bool:
    """Save the API key securely with DPAPI, falling back to plain storage."""
    config_path = get_config_path()
    trimmed = key.strip()
    data_to_store = {"api_key": "", "encrypted": False}

    if trimmed:
        try:
            encrypted_bytes = protect_data(trimmed.encode("utf-8"))
            b64_str = base64.b64encode(encrypted_bytes).decode("ascii")
            data_to_store = {"api_key": b64_str, "encrypted": True}
        except Exception:
            data_to_store = {"api_key": trimmed, "encrypted": False}

    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(data_to_store, f, indent=2)
        return True
    except Exception:
        return False


def load_api_key() -> str:
    """Load the API key from config.json, decrypting if encrypted."""
    config_path = get_config_path()
    if not os.path.isfile(config_path):
        return ""
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        raw_key = data.get("api_key", "")
        if not raw_key:
            return ""
        if data.get("encrypted", False):
            cipher_bytes = base64.b64decode(raw_key.encode("ascii"))
            decrypted_bytes = unprotect_data(cipher_bytes)
            return decrypted_bytes.decode("utf-8")
        return raw_key
    except Exception:
        return ""


# ============================================================================
# Document Extraction & Chunking
# ============================================================================

def extract_text_from_txt(file_path: str) -> str:
    """Extract text from a .txt file trying encodings: utf-8-sig, utf-8, utf-16, cp1256."""
    encodings = ["utf-8-sig", "utf-8", "utf-16", "cp1256"]
    with open(file_path, "rb") as f:
        raw_bytes = f.read()

    for enc in encodings:
        try:
            text = raw_bytes.decode(enc)
            if text.strip():
                return text
            raise ValueError("no text found")
        except UnicodeDecodeError:
            continue
    raise ValueError("Could not decode text file with supported encodings")


def extract_text_from_docx(file_path: str) -> str:
    """Extract text from a .docx file reading paragraphs and table cells in body order."""
    try:
        doc = docx.Document(file_path)
    except Exception as exc:
        raise ValueError(f"Failed to read Word document: {exc}")

    body_blocks = []
    for child in doc.element.body:
        tag = child.tag
        if tag.endswith("p"):
            p = Paragraph(child, doc)
            text = p.text.strip()
            if text:
                body_blocks.append(text)
        elif tag.endswith("tbl"):
            tbl = Table(child, doc)
            table_rows = []
            for row in tbl.rows:
                row_cells = [cell.text.strip() for cell in row.cells]
                row_text = " ".join(row_cells).strip()
                if row_text:
                    table_rows.append(row_text)
            if table_rows:
                body_blocks.append("\n".join(table_rows))

    full_text = "\n\n".join(body_blocks).strip()
    if not full_text:
        raise ValueError("no text found")
    return full_text


def extract_text(file_path: str) -> str:
    """Extract text from .txt or .docx file and reject empty content."""
    lower_path = file_path.lower()
    if lower_path.endswith(".txt"):
        return extract_text_from_txt(file_path)
    if lower_path.endswith(".docx"):
        return extract_text_from_docx(file_path)
    if lower_path.endswith(".doc"):
        raise ValueError("Legacy .doc format is not supported. Please save as .docx.")
    raise ValueError(f"Unsupported file format: {os.path.splitext(file_path)[1]}")


def split_text_into_chunks(text: str, max_chars: int = MAX_CHARS) -> list[str]:
    """Split text into chunks of at most max_chars."""
    clean_text = text.strip()
    if not clean_text:
        return []

    raw_paras = re.split(r"\n\s*\n+", clean_text)
    sentence_pattern = r"(?<=[.!?:\n\u061f\u061b])+(?![.!?:\n\u061f\u061b])\s*"

    atomic_pieces: list[tuple[str, bool]] = []

    for para in raw_paras:
        para_trimmed = para.strip()
        if not para_trimmed:
            continue

        raw_sentences = [s.strip() for s in re.split(sentence_pattern, para_trimmed) if s.strip()]
        if not raw_sentences:
            continue

        first_piece_in_para = True
        for sent in raw_sentences:
            if len(sent) > max_chars:
                for i in range(0, len(sent), max_chars):
                    sub = sent[i:i + max_chars]
                    atomic_pieces.append((sub, first_piece_in_para))
                    first_piece_in_para = False
            else:
                atomic_pieces.append((sent, first_piece_in_para))
                first_piece_in_para = False

    if not atomic_pieces:
        return []

    chunks: list[str] = []
    current_chunk = ""

    for piece, is_new_para in atomic_pieces:
        if not current_chunk:
            current_chunk = piece
            continue

        sep = "\n\n" if is_new_para else " "
        if len(current_chunk) + len(sep) + len(piece) <= max_chars:
            current_chunk += sep + piece
        else:
            chunks.append(current_chunk)
            current_chunk = piece

    if current_chunk:
        chunks.append(current_chunk)

    return chunks


# ============================================================================
# Prompt Construction & Audio Encoding
# ============================================================================

def resolve_style_instruction(style_value: str) -> str:
    """Convert preset label or user free text into a model instruction string."""
    text = (style_value or "").strip()
    if not text:
        return ""
    if text in STYLE_PRESETS:
        return STYLE_PRESETS[text]
    if not text.endswith(":"):
        return text + ":"
    return text


def build_prompt(style_instruction: str, chunk: str) -> str:
    """Build the request text for Gemini TTS in one isolated function."""
    instruction = (style_instruction or "").strip()
    if instruction:
        return f"{instruction}\n\n## Transcript:\n{chunk}"
    return f"## Transcript:\n{chunk}"


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


def check_output_collisions(file_paths: list[str]) -> dict[str, list[str]]:
    """Check if any input files would produce the same output MP3 path."""
    dest_map: dict[str, list[str]] = {}
    for p in file_paths:
        out_path = os.path.normcase(os.path.abspath(os.path.splitext(p)[0] + ".mp3"))
        dest_map.setdefault(out_path, []).append(p)
    return {out_path: files for out_path, files in dest_map.items() if len(files) > 1}


def is_auth_or_invalid_key_error(exc: Exception) -> bool:
    """Check if an exception represents an invalid API key or auth failure (HTTP 400/401/403)."""
    code = getattr(exc, "code", None) or getattr(exc, "status_code", None)
    if code in (400, 401, 403):
        return True
    msg = str(exc).lower()
    if any(k in msg for k in ("401", "403", "api key not valid", "permission denied", "invalid api key", "unauthenticated")):
        return True
    if "400" in msg and ("key" in msg or "invalid" in msg or "api_key" in msg):
        return True
    return False


def is_rate_limit_error(exc: Exception) -> bool:
    """Check if an exception represents rate limiting (HTTP 429)."""
    code = getattr(exc, "code", None) or getattr(exc, "status_code", None)
    if code == 429:
        return True
    msg = str(exc).lower()
    return "429" in msg or "resource_exhausted" in msg or "quota" in msg


# ============================================================================
# Main Application GUI (Modern Dark Theme)
# ============================================================================

class TTSBatchApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_NAME} — {APP_SUBTITLE}")
        self.root.geometry("900x670")
        self.root.minsize(780, 540)
        self.root.configure(bg=COLOR_WINDOW_BG)

        # State
        self.last_applied_voice = VOICES[0]
        self.last_applied_style = DEFAULT_STYLE_LABEL
        self.is_processing = False
        self.stop_event = threading.Event()
        self.queue: queue.Queue = queue.Queue()

        # Apply Modern Dark Stylesheet
        self._setup_dark_theme()

        # Build UI
        self._build_ui()

        # Load persisted API key
        saved_key = load_api_key()
        if saved_key:
            self.api_key_var.set(saved_key)
            self.log("Loaded saved API key from encrypted storage.")

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
        # Row 1: API Key Card
        # --------------------------------------------------------------------
        api_card = ttk.LabelFrame(main_container, text=" Gemini API Configuration ", padding="8")
        api_card.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(api_card, text="Gemini API Key:", style="Card.TLabel").pack(side=tk.LEFT, padx=(4, 8))

        self.api_key_var = tk.StringVar()
        self.api_entry = ttk.Entry(api_card, textvariable=self.api_key_var, show="*")
        self.api_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        self.show_key_var = tk.BooleanVar(value=False)
        self.show_check = ttk.Checkbutton(
            api_card, text="Show", variable=self.show_key_var, command=self._toggle_show_key
        )
        self.show_check.configure(style="TCheckbutton")
        self.show_check.pack(side=tk.LEFT, padx=(0, 4))

        # --------------------------------------------------------------------
        # Row 2: File Queue Card (Treeview)
        # --------------------------------------------------------------------
        list_card = ttk.LabelFrame(main_container, text=" Batch Conversion Queue ", padding="8")
        list_card.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        tree_container = ttk.Frame(list_card, style="Card.TFrame")
        tree_container.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        columns = ("filepath", "voice", "style", "status")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", selectmode="extended")

        self.tree.heading("filepath", text="File Path", anchor=tk.W)
        self.tree.heading("voice", text="Voice", anchor=tk.CENTER)
        self.tree.heading("style", text="Style", anchor=tk.CENTER)
        self.tree.heading("status", text="Status", anchor=tk.CENTER)

        self.tree.column("filepath", width=420, anchor=tk.W, stretch=True)
        self.tree.column("voice", width=110, anchor=tk.CENTER, stretch=False)
        self.tree.column("style", width=180, anchor=tk.CENTER, stretch=False)
        self.tree.column("status", width=160, anchor=tk.CENTER, stretch=False)

        tree_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview, style="Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

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
        settings_card = ttk.LabelFrame(main_container, text=" Voice & Style Settings for Selected File(s) ", padding="10")
        settings_card.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(settings_card, text="Voice:", style="Card.TLabel").grid(row=0, column=0, padx=(4, 6), sticky=tk.W)
        self.voice_var = tk.StringVar(value=self.last_applied_voice)
        self.voice_cb = ttk.Combobox(
            settings_card, textvariable=self.voice_var, values=VOICES, state="readonly", width=16
        )
        self.voice_cb.grid(row=0, column=1, padx=(0, 16), sticky=tk.W)

        ttk.Label(settings_card, text="Style:", style="Card.TLabel").grid(row=0, column=2, padx=(0, 6), sticky=tk.W)
        self.style_var = tk.StringVar(value=self.last_applied_style)
        self.style_cb = ttk.Combobox(
            settings_card, textvariable=self.style_var, values=list(STYLE_PRESETS.keys()), state="normal", width=28
        )
        self.style_cb.grid(row=0, column=3, padx=(0, 16), sticky=tk.W)

        self.apply_btn = ttk.Button(settings_card, text="Apply to Selected", command=self._apply_settings, style="Secondary.TButton")
        self.apply_btn.grid(row=0, column=4, padx=(0, 6))

        self.apply_all_btn = ttk.Button(settings_card, text="Apply to All", command=self._apply_to_all, style="Secondary.TButton")
        self.apply_all_btn.grid(row=0, column=5)

        # --------------------------------------------------------------------
        # Row 4: Action & Progress Row
        # --------------------------------------------------------------------
        action_frame = ttk.Frame(main_container)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        self.start_btn = ttk.Button(action_frame, text="▶ Start Batch", command=self._start_batch, style="Primary.TButton")
        self.start_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = ttk.Button(action_frame, text="■ Stop", command=self._stop_batch, state="disabled", style="Danger.TButton")
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 14))

        self.progress_bar = ttk.Progressbar(action_frame, orient=tk.HORIZONTAL, mode="determinate", style="Horizontal.TProgressbar")
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        self.progress_label = ttk.Label(action_frame, text="Ready", width=24, anchor=tk.E)
        self.progress_label.pack(side=tk.RIGHT)

        # --------------------------------------------------------------------
        # Row 5: Activity Log Card
        # --------------------------------------------------------------------
        log_card = ttk.LabelFrame(main_container, text=" Activity Log ", padding="8")
        log_card.pack(fill=tk.BOTH, expand=False)

        self.log_text = ScrolledText(
            log_card,
            height=6,
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
        self.log_text.tag_configure("err", foreground=COLOR_DANGER)
        self.log_text.tag_configure("ok", foreground=COLOR_SUCCESS)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # ------------------------------------------------------------------------
    # UI Callbacks & Handlers
    # ------------------------------------------------------------------------

    def _toggle_show_key(self):
        self.api_entry.configure(show="" if self.show_key_var.get() else "*")

    def _on_tree_select(self, _event=None):
        selected = self.tree.selection()
        if len(selected) >= 1:
            first_item = selected[0]
            voice_val = self.tree.set(first_item, "voice")
            style_val = self.tree.set(first_item, "style")
            if voice_val:
                self.voice_var.set(voice_val)
            if style_val:
                self.style_var.set(style_val)

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
        for path in file_paths:
            abs_path = os.path.abspath(path)
            base_name = os.path.basename(abs_path)

            if base_name.startswith("~$"):
                continue

            if abs_path.lower().endswith(".doc"):
                messagebox.showwarning(
                    "Unsupported Format",
                    f"Legacy .doc format is not supported for:\n{base_name}\n\n"
                    "Please open this document in Microsoft Word and save it as .docx format.",
                )
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
            self.log(f"Added {added_count} file(s) to queue.")

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
        self.log(f"Removed {count} file(s) from list.")

    def _clear_list(self):
        if self.is_processing:
            messagebox.showinfo("Busy", "Cannot clear files while batch processing is active.")
            return
        items = self.tree.get_children()
        if not items:
            return
        for item in items:
            self.tree.delete(item)
        self.log("Cleared file list.")

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
        self.log(f"Applied settings to {len(selected)} selected file(s): Voice={voice_val}, Style={style_val}")

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
        self.log(f"Applied settings to all {len(items)} file(s): Voice={voice_val}, Style={style_val}")

    def _start_batch(self):
        if self.is_processing:
            return

        api_key = self.api_key_var.get().strip()
        is_mock = os.environ.get("TTS_MOCK") == "1"
        if not api_key and not is_mock:
            messagebox.showwarning("API Key Required", "Please enter a valid Gemini API Key before starting.")
            self.api_entry.focus_set()
            return

        if api_key:
            save_api_key(api_key)

        items = self.tree.get_children()
        if not items:
            messagebox.showwarning("Empty List", "Please add at least one .txt or .docx file to convert.")
            return

        file_records: list[tuple[str, str, str, str]] = []
        file_paths = []
        for item in items:
            f_path = self.tree.set(item, "filepath")
            f_voice = self.tree.set(item, "voice")
            f_style = self.tree.set(item, "style")
            if not os.path.isfile(f_path):
                messagebox.showerror("File Not Found", f"File does not exist:\n{f_path}\n\nPlease check or remove it.")
                return
            file_records.append((item, f_path, f_voice, f_style))
            file_paths.append(f_path)

        collisions = check_output_collisions(file_paths)
        if collisions:
            collision_details = []
            for out_p, files in collisions.items():
                collision_details.append(f"Output: {out_p}\nConflicting input files:\n  - " + "\n  - ".join(files))
            msg = (
                "Output name collisions detected! The following files would overwrite each other:\n\n"
                + "\n\n".join(collision_details)
                + "\n\nPlease rename or remove conflicting files before starting."
            )
            messagebox.showerror("Output Path Collision", msg)
            return

        self.is_processing = True
        self.stop_event.clear()
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.add_btn.configure(state="disabled")
        self.remove_btn.configure(state="disabled")
        self.clear_btn.configure(state="disabled")
        self.apply_btn.configure(state="disabled")
        self.apply_all_btn.configure(state="disabled")

        self.progress_bar.configure(value=0, maximum=100)
        self.progress_label.configure(text="Starting...")
        self.log("Starting batch processing...")

        worker = threading.Thread(
            target=self._worker_thread,
            args=(file_records, api_key),
            daemon=True,
        )
        worker.start()

    def _stop_batch(self):
        if not self.is_processing:
            return
        self.stop_event.set()
        self.stop_btn.configure(state="disabled")
        self.log("Stop requested. Waiting for current chunk to finish...")
        self.progress_label.configure(text="Stopping...")

    def _worker_thread(self, file_records: list[tuple[str, str, str, str]], api_key: str):
        """Sequential background processing worker for files and chunks."""
        is_mock = os.environ.get("TTS_MOCK") == "1"
        client = None
        if not is_mock:
            try:
                client = genai.Client(api_key=api_key)
            except Exception as exc:
                self.queue.put(("auth_error", f"Failed to initialize Gemini client: {exc}"))
                return

        prepared_files: list[dict] = []
        total_chunks = 0

        for item_id, file_path, voice, style in file_records:
            try:
                text = extract_text(file_path)
                chunks = split_text_into_chunks(text, MAX_CHARS)
                if not chunks:
                    raise ValueError("no text found")
                total_chunks += len(chunks)
                prepared_files.append({
                    "item_id": item_id,
                    "file_path": file_path,
                    "voice": voice,
                    "style": style,
                    "chunks": chunks,
                    "error": None,
                })
            except Exception as exc:
                short_err = str(exc)
                prepared_files.append({
                    "item_id": item_id,
                    "file_path": file_path,
                    "voice": voice,
                    "style": style,
                    "chunks": [],
                    "error": short_err,
                })

        self.queue.put(("progress_init", total_chunks))

        completed_chunks = 0
        auth_failed = False

        for file_info in prepared_files:
            item_id = file_info["item_id"]
            file_path = file_info["file_path"]
            voice = file_info["voice"]
            style = file_info["style"]
            chunks = file_info["chunks"]
            preload_error = file_info["error"]

            if self.stop_event.is_set():
                self.queue.put(("item_status", item_id, "Stopped"))
                continue

            if auth_failed:
                self.queue.put(("item_status", item_id, "Stopped"))
                continue

            if preload_error:
                self.queue.put(("item_status", item_id, f"Error: {preload_error}"))
                self.queue.put(("log", f"Failed to process '{os.path.basename(file_path)}': {preload_error}"))
                continue

            output_mp3 = os.path.splitext(file_path)[0] + ".mp3"
            pcm_accumulator = bytearray()
            detected_sample_rate = DEFAULT_SAMPLE_RATE
            num_chunks = len(chunks)
            file_failed = False
            file_err_msg = ""

            self.queue.put(("log", f"Processing '{os.path.basename(file_path)}' ({num_chunks} chunks, Voice: {voice})..."))

            for idx, chunk in enumerate(chunks, start=1):
                if self.stop_event.is_set():
                    self.queue.put(("item_status", item_id, "Stopped"))
                    self.queue.put(("log", f"Stopped while processing '{os.path.basename(file_path)}'."))
                    break

                self.queue.put(("item_status", item_id, f"Processing (chunk {idx}/{num_chunks})"))

                style_instruction = resolve_style_instruction(style)
                prompt_text = build_prompt(style_instruction, chunk)

                chunk_pcm, chunk_rate, err = self._fetch_audio_chunk_with_retry(
                    client=client,
                    prompt_text=prompt_text,
                    voice=voice,
                    is_mock=is_mock,
                )

                if err:
                    file_failed = True
                    file_err_msg = err
                    if is_auth_or_invalid_key_error(Exception(err)):
                        auth_failed = True
                    break

                pcm_accumulator.extend(chunk_pcm)
                if chunk_rate:
                    detected_sample_rate = chunk_rate

                completed_chunks += 1
                self.queue.put(("progress_update", completed_chunks, total_chunks))

            if self.stop_event.is_set():
                self.queue.put(("item_status", item_id, "Stopped"))
                continue

            if file_failed:
                short_reason = file_err_msg.splitlines()[0]
                if len(short_reason) > 35:
                    short_reason = short_reason[:32] + "..."
                self.queue.put(("item_status", item_id, f"Error: {short_reason}"))
                self.queue.put(("log", f"Error on '{os.path.basename(file_path)}': {file_err_msg}"))
                if auth_failed:
                    self.queue.put(("log", "Authentication failure detected. Halting remaining files."))
                continue

            try:
                save_pcm_as_mp3(bytes(pcm_accumulator), detected_sample_rate, output_mp3)
                file_size_kb = os.path.getsize(output_mp3) / 1024.0
                self.queue.put(("item_status", item_id, "Done"))
                self.queue.put(("log", f"Finished '{os.path.basename(output_mp3)}' ({file_size_kb:.1f} KB)."))
            except Exception as save_exc:
                self.queue.put(("item_status", item_id, "Error: save failed"))
                self.queue.put(("log", f"Failed to save MP3 for '{os.path.basename(file_path)}': {save_exc}"))

        if self.stop_event.is_set():
            self.queue.put(("batch_stopped", None))
        else:
            self.queue.put(("batch_done", None))

    def _fetch_audio_chunk_with_retry(
        self, client: genai.Client | None, prompt_text: str, voice: str, is_mock: bool
    ) -> tuple[bytes, int, str | None]:
        """Fetch TTS audio for a single chunk with retry backoff (2s, 4s)."""
        if is_mock:
            time.sleep(0.08)
            return b"\x00" * (DEFAULT_SAMPLE_RATE * 2), DEFAULT_SAMPLE_RATE, None

        max_retries = 3
        backoff_delays = [2.0, 4.0]

        for attempt in range(max_retries):
            if self.stop_event.is_set():
                return b"", DEFAULT_SAMPLE_RATE, "Operation stopped by user"

            try:
                contents = [
                    types.Content(
                        role="user",
                        parts=[types.Part.from_text(text=prompt_text)],
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

                for resp_chunk in client.models.generate_content_stream(
                    model=MODEL_NAME, contents=contents, config=config
                ):
                    if self.stop_event.is_set():
                        return b"", DEFAULT_SAMPLE_RATE, "Operation stopped by user"

                    for part in (resp_chunk.parts or []):
                        if part.inline_data and part.inline_data.data:
                            raw = part.inline_data.data
                            if isinstance(raw, str):
                                raw = base64.b64decode(raw)
                            pcm_chunks.append(raw)
                            if part.inline_data.mime_type:
                                detected_rate = parse_sample_rate(part.inline_data.mime_type, detected_rate)

                full_chunk_pcm = b"".join(pcm_chunks)
                if not full_chunk_pcm:
                    raise RuntimeError("No audio data received from Gemini TTS API")
                return full_chunk_pcm, detected_rate, None

            except Exception as exc:
                if self.stop_event.is_set():
                    return b"", DEFAULT_SAMPLE_RATE, "Operation stopped by user"

                if is_auth_or_invalid_key_error(exc):
                    return b"", DEFAULT_SAMPLE_RATE, f"Invalid API Key or Permission Error: {exc}"

                if attempt == max_retries - 1:
                    return b"", DEFAULT_SAMPLE_RATE, f"Failed after {max_retries} attempts: {exc}"

                if is_rate_limit_error(exc):
                    delay = 7.0 * (attempt + 1)
                    self.queue.put(("log", f"Rate limit reached (429). Waiting {delay:.0f}s before retry..."))
                else:
                    delay = backoff_delays[attempt]
                    self.queue.put(("log", f"API warning ({exc}). Retrying in {delay:.0f}s (attempt {attempt + 2}/{max_retries})..."))

                sleep_passed = 0.0
                while sleep_passed < delay:
                    if self.stop_event.is_set():
                        return b"", DEFAULT_SAMPLE_RATE, "Operation stopped by user"
                    time.sleep(min(0.2, delay - sleep_passed))
                    sleep_passed += 0.2

        return b"", DEFAULT_SAMPLE_RATE, "Max retries exceeded"

    # ------------------------------------------------------------------------
    # Queue Processing & UI Dispatching
    # ------------------------------------------------------------------------

    def _process_queue(self):
        try:
            while True:
                msg_type, *args = self.queue.get_nowait()
                if msg_type == "log":
                    self.log(args[0])
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
                elif msg_type == "batch_done":
                    self._on_batch_finished(stopped=False)
                elif msg_type == "batch_stopped":
                    self._on_batch_finished(stopped=True)
                elif msg_type == "auth_error":
                    self.log(f"ERROR: {args[0]}")
                    messagebox.showerror("Authentication Error", args[0])
                    self._on_batch_finished(stopped=True)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._process_queue)

    def _on_batch_finished(self, stopped: bool):
        self.is_processing = False
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.add_btn.configure(state="normal")
        self.remove_btn.configure(state="normal")
        self.clear_btn.configure(state="normal")
        self.apply_btn.configure(state="normal")
        self.apply_all_btn.configure(state="normal")

        if stopped:
            self.progress_label.configure(text="Stopped")
            self.log("Batch processing stopped.")
        else:
            self.progress_label.configure(text="Completed")
            self.log("All files in batch completed.")

    def log(self, message: str):
        """Append a message to the activity log text widget."""
        timestamp = time.strftime("[%H:%M:%S] ")
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, timestamp, "time")
        if "error" in message.lower() or "failed" in message.lower():
            self.log_text.insert(tk.END, message + "\n", "err")
        elif "finished" in message.lower() or "completed" in message.lower():
            self.log_text.insert(tk.END, message + "\n", "ok")
        else:
            self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def _on_close(self):
        key = self.api_key_var.get().strip()
        if key:
            save_api_key(key)
        self.root.destroy()


# ============================================================================
# Entry Point
# ============================================================================

def main():
    root = tk.Tk()
    app = TTSBatchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
