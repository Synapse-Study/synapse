#!/usr/bin/env python3
"""
anki_study_suite.py
===================
Synapse – AI PowerPoint & MBBS Flashcard Studio
Modern Desktop App with PyWebView & HTML5/Tailwind GUI
Integrates: python-pptx • EasyOCR • Google Gemini • genanki • Pillow • pywebview • AnkiConnect

Calibrated MBBS Flashcard & Image Occlusion Suite
"""

import os
import sys

# ── macOS Process & Bundle Identity (Must run BEFORE WebKit/pywebview load) ──
if sys.platform == "darwin":
    try:
        from Foundation import NSBundle
        from Cocoa import NSProcessInfo
        bundle = NSBundle.mainBundle()
        if bundle:
            info = bundle.localizedInfoDictionary() or bundle.infoDictionary()
            if info is not None:
                info["CFBundleName"] = "Synapse"
                info["CFBundleDisplayName"] = "Synapse"
        NSProcessInfo.processInfo().setProcessName_("Synapse")
    except Exception:
        pass

import io
import json
import hashlib
import random
import threading
import queue
import subprocess
import textwrap
import traceback
import re
import uuid
import base64
import time
import shutil
import sqlite3
import webbrowser
import urllib.request
import urllib.error
import ssl
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, Any

try:
    import certifi
    os.environ.setdefault("SSL_CERT_FILE", certifi.where())
    os.environ.setdefault("REQUESTS_CA_BUNDLE", certifi.where())
except Exception:
    pass

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
import requests
import genanki
import webview
from webview.menu import Menu, MenuAction, MenuSeparator

from synapse_auth import (
    SUPABASE_URL,
    SUPABASE_PUBLISHABLE_KEY,
    QUESTIONS_APP_URL,
    sign_in_synapse,
    sign_up_synapse,
    clear_session_from_config,
    load_secure_config, save_secure_config, get_valid_session,
    accept_question_session, sign_out_synapse,
)

# Prevent pywebview from injecting duplicate default View/Edit menus & allow local file URLs
webview.settings["SHOW_DEFAULT_MENUS"] = False
webview.settings["ALLOW_FILE_URLS"] = True
webview.settings["OPEN_DEVTOOLS_IN_DEBUG"] = False

# Try importing Cocoa / PyObjC for native macOS integration
try:
    import AppKit
    from Cocoa import (
        NSApp,
        NSAppearance,
        NSAppearanceNameAqua,
        NSAppearanceNameDarkAqua,
        NSProcessInfo,
        NSObject
    )
    import PyObjCTools.AppHelper as AppHelper
    HAS_COCOA = True
except Exception:
    HAS_COCOA = False

# Try importing PDF rendering engines (fitz / PyMuPDF and pypdfium2)
try:
    import fitz  # type: ignore
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

try:
    import pypdfium2 as pdfium  # type: ignore
    HAS_PDFIUM = True
except ImportError:
    HAS_PDFIUM = False

# Global lock to serialize non-thread-safe C-extension operations (PDFium/fitz, PyTorch/EasyOCR, OpenCV)
# and prevent ARM64 libmalloc heap corruption across worker threads.
_C_EXTENSION_LOCK = threading.Lock()

# ─────────────────────────────────────────────────────────────────────────────
# 1. CONSTANTS, PATHS & DEFAULTS
# ─────────────────────────────────────────────────────────────────────────────
CURRENT_VERSION = "2.3.0"
GITHUB_REPO = "Synapse-Study/synapse"
LATEST_RELEASE_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
RELEASES_LIST_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases"

if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    BUNDLE_DIR = Path(sys._MEIPASS).resolve()
else:
    BUNDLE_DIR = Path(__file__).parent.resolve()
BASE_DIR = BUNDLE_DIR

def parse_version_tuple(v_str: str) -> tuple[int, ...]:
    """Parse version string like '2.2.1' or 'v2.2.1' into numeric tuple (2, 2, 1)."""
    nums = [int(x) for x in re.findall(r'\d+', str(v_str or ""))]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums[:3])

def check_github_update() -> dict:
    """Query GitHub releases API (with web redirect and SSL fallbacks) for latest Synapse version."""
    headers = {
        "User-Agent": f"SynapseApp/{CURRENT_VERSION}",
        "Accept": "application/vnd.github.v3+json",
    }
    web_headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }
    ssl_context = None
    try:
        import certifi
        ssl_context = ssl.create_default_context(cafile=certifi.where())
    except Exception:
        try:
            ssl_context = ssl.create_default_context()
        except Exception:
            pass

    def _fetch_url(url: str, h: dict = headers):
        req = urllib.request.Request(url, headers=h)
        try:
            return urllib.request.urlopen(req, timeout=12, context=ssl_context)
        except urllib.error.URLError as u_err:
            # Fallback to unverified SSL context if system certificates fail or proxy intercepts
            err_str = str(u_err).lower()
            if "certificate_verify_failed" in err_str or "certificate verify failed" in err_str:
                try:
                    unverified_ctx = ssl._create_unverified_context()
                    return urllib.request.urlopen(req, timeout=12, context=unverified_ctx)
                except Exception:
                    pass
            raise u_err

    try:
        data = None
        try:
            with _fetch_url(LATEST_RELEASE_API) as resp:
                if resp.status == 200:
                    data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as http_err:
            if http_err.code in (403, 404, 429):
                # Fallback 1: check releases list in case latest tag is not marked 'latest'
                try:
                    with _fetch_url(RELEASES_LIST_API) as resp2:
                        if resp2.status == 200:
                            releases = json.loads(resp2.read().decode("utf-8"))
                            if isinstance(releases, list) and releases:
                                data = releases[0]
                except Exception:
                    pass

                # Fallback 2: Check web redirect at https://github.com/Synapse-Study/synapse/releases/latest
                if not data:
                    try:
                        web_url = f"https://github.com/{GITHUB_REPO}/releases/latest"
                        with _fetch_url(web_url, h=web_headers) as web_resp:
                            final_url = web_resp.geturl()
                            m = re.search(r"/releases/tag/([^/?#]+)", final_url)
                            if m:
                                tag_name = m.group(1).strip()
                                latest_ver = parse_version_tuple(tag_name)
                                curr_ver = parse_version_tuple(CURRENT_VERSION)
                                has_update = latest_ver > curr_ver
                                return {
                                    "checked": True,
                                    "update_available": has_update,
                                    "current_version": CURRENT_VERSION,
                                    "latest_version": tag_name.lstrip("v") if tag_name else CURRENT_VERSION,
                                    "tag_name": tag_name,
                                    "release_name": f"Synapse {tag_name}",
                                    "release_notes": "See GitHub release page for details.",
                                    "html_url": final_url,
                                    "download_url": f"https://github.com/{GITHUB_REPO}/releases/download/{tag_name}/Synapse.dmg",
                                    "published_at": "",
                                }
                    except Exception:
                        pass

                if not data:
                    # Clean fallback: report up to date cleanly
                    return {
                        "checked": True,
                        "update_available": False,
                        "current_version": CURRENT_VERSION,
                        "latest_version": CURRENT_VERSION,
                        "tag_name": f"v{CURRENT_VERSION}",
                        "release_name": f"Synapse v{CURRENT_VERSION}",
                        "release_notes": "You are running the latest version of Synapse.",
                        "html_url": f"https://github.com/{GITHUB_REPO}",
                        "download_url": f"https://github.com/{GITHUB_REPO}/releases",
                        "published_at": "",
                    }
            else:
                raise http_err

        if data:
            tag_name = str(data.get("tag_name", "")).strip()
            latest_ver = parse_version_tuple(tag_name)
            curr_ver = parse_version_tuple(CURRENT_VERSION)
            has_update = latest_ver > curr_ver

            dmg_url = None
            for asset in data.get("assets", []):
                name = str(asset.get("name", "")).lower()
                if name.endswith(".dmg"):
                    dmg_url = asset.get("browser_download_url")
                    break

            html_url = data.get("html_url", f"https://github.com/{GITHUB_REPO}/releases/latest")
            return {
                "checked": True,
                "update_available": has_update,
                "current_version": CURRENT_VERSION,
                "latest_version": tag_name.lstrip("v") if tag_name else CURRENT_VERSION,
                "tag_name": tag_name,
                "release_name": data.get("name", tag_name),
                "release_notes": data.get("body", ""),
                "html_url": html_url,
                "download_url": dmg_url or html_url,
                "published_at": data.get("published_at", ""),
            }

        return {"checked": False, "update_available": False, "current_version": CURRENT_VERSION, "error": "No release information available"}
    except Exception as exc:
        err_msg = str(exc).strip()
        if not err_msg:
            err_msg = "Connection timed out" if (isinstance(exc, TimeoutError) or "timeout" in type(exc).__name__.lower()) else type(exc).__name__
        return {"checked": False, "update_available": False, "current_version": CURRENT_VERSION, "error": err_msg}

def _resolve_app_data_dir() -> Path:
    """Resolve per-user Synapse data directory (with optional SYNAPSE_DATA_DIR override)."""
    env_override = os.environ.get("SYNAPSE_DATA_DIR", "").strip()
    if env_override:
        return Path(env_override).expanduser().resolve()
    return Path.home() / "Library" / "Application Support" / "Synapse"

APP_DATA_DIR = _resolve_app_data_dir()
CONFIG_PATH = APP_DATA_DIR / "config.json"
LECTURES_CACHE_DB_PATH = APP_DATA_DIR / "kcl_lectures_cache.db"
EXPORTS_DIR = APP_DATA_DIR / "anki_exports"
SLIDE_CACHE_DIR = APP_DATA_DIR / "slide_cache"
TEMPLATES_DIR = APP_DATA_DIR / "templates"
SAVED_NOTES_DIR = APP_DATA_DIR / "saved_notes"
CACHE_DIR = SLIDE_CACHE_DIR

def _ensure_app_data_dirs() -> None:
    """Ensure per-user Application Support directories exist at runtime."""
    try:
        APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
        EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
        SLIDE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
        SAVED_NOTES_DIR.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass

def _extract_text_from_file(file_path: Path) -> str:
    """Extract clean text content from .md, .txt, .pdf, or .docx files."""
    if not file_path.exists():
        return ""
    ext = file_path.suffix.lower()
    if ext in (".md", ".txt", ".markdown"):
        try:
            return file_path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            return ""
    elif ext == ".pdf":
        text_parts = []
        if HAS_PDFIUM:
            try:
                doc = pdfium.PdfDocument(str(file_path))
                for page in doc:
                    text_page = page.get_textpage()
                    text_parts.append(text_page.get_text_range())
                if text_parts:
                    return "\n\n".join(text_parts)
            except Exception:
                pass
        if HAS_FITZ:
            try:
                import fitz
                doc = fitz.open(str(file_path))
                for page in doc:
                    text_parts.append(page.get_text())
                if text_parts:
                    return "\n\n".join(text_parts)
            except Exception:
                pass
        try:
            return file_path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            return ""
    else:
        try:
            return file_path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            return ""

# Ensure directories exist immediately at startup
_ensure_app_data_dirs()

_CONFIG_TEMPLATE: dict = {
    "THEME":                    "system",
    "ANKI_URL":                 "http://127.0.0.1:8765",
    "DEFAULT_PARENT_DECK":      "MBBS Year 1",
    "DEFAULT_OCC_MODE":         "hide_all",
    "LECTURES_DIR":             "",
    "EXPORTS_DIR":              str(EXPORTS_DIR),
    "OCR_CONFIDENCE_THRESHOLD": 0.25,
    "AUTO_RUN_OCR":             False,
    "GEMINI_MODEL":             "gemini-4-hydrogen",
    "GEMINI_NOTES_API_KEY":     "",
    "GEMINI_CHAT_API_KEY":      "",
    "GEMINI_API_KEY":           "",
    "NOTION_API_KEY":           "",
    "NOTION_PAGE_ID":           "",
    "TARGET_DECK":              "MBBS Year 1::Year 1 Stage 1::PAS::A - FAH::01. Anatomical Terminology",
    "AUTH_EMAIL":               "",
    "AUTH_ROLE":                "",
    "SUPABASE_ACCESS_TOKEN":    "",
    "SUPABASE_REFRESH_TOKEN":   "",
    "ONBOARDING_COMPLETE":      False,
    "USER_NAME":                "",
    "TIMETABLE_URL":            "",
    "APP_ACCENT":               "cobalt",
    "AUTO_PREPARE_NOTES":       True,
    "NOTES_STYLE_MODE":         "default",
    "CUSTOM_STYLE_INSTRUCTIONS": "",
    "CUSTOM_NOTE_TEMPLATES":     [],
    "DISCORD_FEEDBACK_WEBHOOK":  "https://discord.com/api/webhooks/1572979254625312929/3exv62rYExNtpoibQPrCCFOMkzipWFpRI5pZDzyyoprQbgNdTI3uP1M6nlH3DIKQf85K",
    "DISMISSED_ANNOUNCEMENTS":   [],
    "dismissed_announcements":   [],
}

_ENV_FALLBACKS: dict = {
    "GEMINI_NOTES_API_KEY": "GEMINI_NOTES_API_KEY",
    "GEMINI_CHAT_API_KEY":  "GEMINI_CHAT_API_KEY",
    "GEMINI_API_KEY":       "GEMINI_API_KEY",
    "NOTION_API_KEY":       "NOTION_API_KEY",
    "NOTION_PAGE_ID":       "NOTION_PAGE_ID",
    "TARGET_DECK":          "TARGET_DECK",
    "ANKI_URL":             "ANKI_URL",
    "DISCORD_FEEDBACK_WEBHOOK": "DISCORD_FEEDBACK_WEBHOOK",
}

def mask_api_key(key: str) -> str:
    """Safely mask API key for Terminal and status logging, showing only the tail."""
    k = (key or "").strip()
    if not k:
        return "<not set>"
    if len(k) <= 6:
        return "***"
    return f"...{k[-4:]}"

def pil_to_base64_data_uri(img: Image.Image, format: str = "PNG", quality: int = 90) -> str:
    """Converts a PIL Image to a base64 Data URI."""
    if img is None:
        return ""
    buf = io.BytesIO()
    fmt = format.upper()
    if fmt == "JPEG" and img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    save_kwargs = {}
    if fmt == "JPEG":
        save_kwargs["quality"] = quality
    img.save(buf, format=fmt, **save_kwargs)
    b64 = base64.b64encode(buf.getvalue()).decode("ascii")
    mime = "image/png" if fmt == "PNG" else "image/jpeg"
    return f"data:{mime};base64,{b64}"

def is_christmas_season() -> bool:
    """
    Seasonal check helper for Christmas theme and festive accents.
    Enforces the Dec 1 – Dec 26 window.
    """
    now = datetime.now()
    return (now.month == 12 and 1 <= now.day <= 26)

# ─────────────────────────────────────────────────────────────────────────────
# 2. CONFIG MANAGER
# ─────────────────────────────────────────────────────────────────────────────
class ConfigManager:
    def __init__(self, path: Path = CONFIG_PATH):
        self._path = path
        self._data: dict = {}
        self.load()

    def load(self) -> dict:
        _ensure_app_data_dirs()
        try:
            self._data = load_secure_config(self._path)
        except Exception:
            local_fallback = BASE_DIR / "config.json"
            if local_fallback.exists():
                try:
                    self._data = json.loads(local_fallback.read_text(encoding="utf-8"))
                except Exception:
                    self._data = {}
            else:
                self._data = {}

        for k, v in _CONFIG_TEMPLATE.items():
            if k not in self._data or self._data[k] is None:
                env_val = os.environ.get(_ENV_FALLBACKS.get(k, ""))
                self._data[k] = env_val if env_val else v

        # Clean legacy relative path aliases if present
        if str(self._data.get("LECTURES_DIR", "")).strip() in ("./Lectures_OneDrive", "Lectures_OneDrive"):
            self._data["LECTURES_DIR"] = ""
        if str(self._data.get("EXPORTS_DIR", "")).strip() in ("", "./anki_exports", "anki_exports"):
            self._data["EXPORTS_DIR"] = str(EXPORTS_DIR)

        # Christmas accent is strictly constrained to Dec 1 - Dec 26
        if self._data.get("APP_ACCENT") == "christmas" and not is_christmas_season():
            self._data["APP_ACCENT"] = "cobalt"

        # Backward compatibility and migration for split Gemini API keys
        legacy_key = (self._data.get("GEMINI_API_KEY") or self._data.get("gemini_api_key") or "").strip()
        notes_key = (self._data.get("GEMINI_NOTES_API_KEY") or self._data.get("gemini_notes_api_key") or "").strip()
        chat_key = (self._data.get("GEMINI_CHAT_API_KEY") or self._data.get("gemini_chat_api_key") or "").strip()

        if legacy_key:
            if not notes_key:
                self._data["GEMINI_NOTES_API_KEY"] = legacy_key
                notes_key = legacy_key
            if not chat_key:
                self._data["GEMINI_CHAT_API_KEY"] = legacy_key
                chat_key = legacy_key

        # Ensure a valid model setting
        if not self._data.get("GEMINI_MODEL") or self._data.get("GEMINI_MODEL") in ("gemini-3.1-flash-lite", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"):
            self._data["GEMINI_MODEL"] = "gemini-3-flash-preview"
            try:
                save_secure_config(self._data, self._path)
            except Exception:
                pass
        return self._data

    def save(self, updates: dict) -> None:
        _ensure_app_data_dirs()
        self._data.update(updates)
        try:
            save_secure_config(self._data, self._path)
        except Exception:
            try:
                self._path.write_text(json.dumps(self._data, indent=2), encoding="utf-8")
            except Exception:
                pass

    def get(self, key: str, default: Any = "") -> Any:
        return self._data.get(key, default)

    def set(self, key: str, value: Any) -> None:
        self._data[key] = value

    def get_notes_api_key(self) -> str:
        """Dedicated Gemini API Key for Notes & Flashcards generation.
        Falls back gracefully to Chat Key or legacy Key if only one is provided."""
        k = (self._data.get("GEMINI_NOTES_API_KEY") or 
             self._data.get("gemini_notes_api_key") or 
             os.environ.get("GEMINI_NOTES_API_KEY") or 
             self._data.get("GEMINI_CHAT_API_KEY") or 
             self._data.get("gemini_chat_api_key") or 
             os.environ.get("GEMINI_CHAT_API_KEY") or 
             self._data.get("GEMINI_API_KEY") or 
             self._data.get("gemini_api_key") or 
             os.environ.get("GEMINI_API_KEY") or "")
        return str(k).strip()

    def get_chat_api_key(self) -> str:
        """Dedicated Gemini API Key for AI Tutor / Chat.
        Falls back gracefully to Notes Key or legacy Key if only one is provided."""
        k = (self._data.get("GEMINI_CHAT_API_KEY") or 
             self._data.get("gemini_chat_api_key") or 
             os.environ.get("GEMINI_CHAT_API_KEY") or 
             self._data.get("GEMINI_NOTES_API_KEY") or 
             self._data.get("gemini_notes_api_key") or 
             os.environ.get("GEMINI_NOTES_API_KEY") or 
             self._data.get("GEMINI_API_KEY") or 
             self._data.get("gemini_api_key") or 
             os.environ.get("GEMINI_API_KEY") or "")
        return str(k).strip()

    def get_dismissed_announcements(self) -> list[str]:
        dismissed = self._data.get("DISMISSED_ANNOUNCEMENTS")
        if dismissed is None:
            dismissed = self._data.get("dismissed_announcements", [])
        if not isinstance(dismissed, list):
            dismissed = []
        return [str(x) for x in dismissed]

    def dismiss_announcement(self, announcement_id: str) -> None:
        ann_id = str(announcement_id).strip()
        if not ann_id:
            return
        curr = self.get_dismissed_announcements()
        if ann_id not in curr:
            curr.append(ann_id)
            self._data["DISMISSED_ANNOUNCEMENTS"] = curr
            self._data["dismissed_announcements"] = curr
            self.save({"DISMISSED_ANNOUNCEMENTS": curr, "dismissed_announcements": curr})

# ─────────────────────────────────────────────────────────────────────────────
# 3. ANKICONNECT BRIDGE (Port 8765)
# ─────────────────────────────────────────────────────────────────────────────
class AnkiBridge:
    ANKI_URL = "http://127.0.0.1:8765"

    @classmethod
    def invoke(cls, action: str, **params) -> dict:
        payload = {"action": action, "version": 6, "params": params}
        try:
            resp = requests.post(cls.ANKI_URL, json=payload, timeout=3.0)
            if resp.status_code == 200:
                data = resp.json()
                if data.get("error"):
                    return {"success": False, "error": data["error"]}
                return {"success": True, "result": data.get("result")}
            return {"success": False, "error": f"HTTP {resp.status_code}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @classmethod
    def test_connection(cls, url: str) -> dict:
        payload = {"action": "version", "version": 6}
        try:
            resp = requests.post(url, json=payload, timeout=3.0)
            if resp.status_code == 200:
                data = resp.json()
                if data.get("error"):
                    return {"connected": False, "error": data["error"]}
                return {"connected": True, "version": data.get("result")}
            return {"connected": False, "error": f"HTTP {resp.status_code}"}
        except Exception as e:
            return {"connected": False, "error": str(e)}

    @classmethod
    def is_connected(cls) -> bool:
        res = cls.invoke("version")
        return res.get("success", False)

    @classmethod
    def get_deck_names(cls) -> list[str]:
        res = cls.invoke("deckNames")
        if res.get("success") and isinstance(res.get("result"), list):
            return res["result"]
        return []

    @classmethod
    def create_deck_if_missing(cls, deck_name: str) -> bool:
        res = cls.invoke("createDeck", deck=deck_name)
        return res.get("success", False)

    @classmethod
    def store_media_file(cls, filename: str, data_bytes: bytes) -> bool:
        b64_data = base64.b64encode(data_bytes).decode("ascii")
        res = cls.invoke("storeMediaFile", filename=filename, data=b64_data)
        return res.get("success", False)

    @classmethod
    def add_notes(cls, notes: list[dict]) -> dict:
        return cls.invoke("addNotes", notes=notes)

    @classmethod
    def sync(cls) -> dict:
        return cls.invoke("sync")

# ─────────────────────────────────────────────────────────────────────────────
# 4. CENTRAL APPLICATION STATE
# ─────────────────────────────────────────────────────────────────────────────
class AppState:
    def __init__(self):
        self.pptx_path:          Optional[Path]        = None
        self._current_lecture_name: str                = ""
        self.presentation:       Optional[object]      = None
        self.total_slides:       int                   = 0
        self.current_slide_idx:  int                   = 0

        # Slide images & diagram render
        self.rendered_slide_image: Optional[Image.Image] = None
        self.composite_diagrams:   list[Image.Image]     = []
        self.slide_images:         list[Image.Image]     = []
        self.active_image:         Optional[Image.Image] = None
        self.slide_raw_text:       str                   = ""
        self.current_view_mode:    str                   = "Full Slide Diagram"
        self.raw_slide_views:      list[str]             = []

        # OCR detections on active image: list of {"id": int, "text": str, "box": (x1, y1, x2, y2), "prob": float, "checked": bool, "is_redaction": bool}
        self.ocr_results:          list[dict]            = []
        self.ocr_ready:            bool                  = False
        self.ocr_reader:           Optional[object]      = None

        # Selection state
        self.selected_indices:     set[int]              = set()

        # Undo / Redo history stacks
        self.undo_stack:           list[list[dict]]      = []
        self.redo_stack:           list[list[dict]]      = []

        # Slide OCR cache: (slide_idx, view_mode) -> list[dict]
        self.slide_ocr_cache:      dict[tuple[int, str], list[dict]] = {}

        # Flashcards & Compilation Queue
        self.normal_cards:         list[dict]            = []
        self.queued_occlusion_cards: list[dict]          = []
        self.slide_overview:       list[dict]            = []
        self.target_deck:          str                   = "MBBS Jassem A.::Year 1 Stage 1::PAS::A - FAH::01. Anatomical Terminology"
        self.available_decks:      list[str]             = []
        self.anki_connected:       bool                  = False

        # Lecture notes
        self.current_slide_notes:  str                   = ""

    @property
    def current_lecture_name(self) -> str:
        if self._current_lecture_name:
            return self._current_lecture_name
        return self.pptx_path.stem if self.pptx_path else ""

    @current_lecture_name.setter
    def current_lecture_name(self, value: str):
        self._current_lecture_name = value

# ─────────────────────────────────────────────────────────────────────────────
# 5. SLIDE RENDER ENGINE (High-Res Slide Diagrams)
# ─────────────────────────────────────────────────────────────────────────────
class SlideRenderEngine:
    """
    Renders PowerPoint slides into crisp, full-fidelity raster images
    (preserving all PowerPoint shapes, curves, diagrams, SmartArt, equations,
    callouts, and multi-image layouts).
    Stores rendered PDFs inside APP_DATA_DIR / 'slide_cache'.
    All PDF operations and slide exports are serialized to protect native C extensions.
    """
    def __init__(self, on_ready_callback=None):
        self._on_ready = on_ready_callback
        self._pdf = None
        self._fitz_doc = None
        self._total_pages: int = 0
        self._prs: Optional[Presentation] = None
        self._slide_to_pdf_page: dict[int, Optional[int]] = {}
        self._current_path: Optional[Path] = None
        self._cache_dir = APP_DATA_DIR / "slide_cache"
        try:
            self._cache_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        self._slide_cache: dict[int, Image.Image] = {}
        self._is_rendering: bool = False
        self._render_queue: queue.Queue = queue.Queue()
        self._worker_thread = threading.Thread(target=self._sequential_worker, daemon=True)
        self._worker_thread.start()

    @property
    def is_rendering(self) -> bool:
        return self._is_rendering

    def set_presentation(self, prs: Optional[Presentation]):
        self._prs = prs
        self._recompute_slide_mapping()

    def _recompute_slide_mapping(self):
        self._slide_to_pdf_page = {}
        if not self._prs:
            return

        total_slides = len(self._prs.slides)
        visible_indices = []
        for idx, slide in enumerate(self._prs.slides):
            is_hidden = False
            try:
                is_hidden = (slide._element.get("show") == "0")
            except Exception:
                pass
            if not is_hidden:
                visible_indices.append(idx)

        # Case A: PDF page count matches visible slide count exactly
        # (Standard behavior of PowerPoint & LibreOffice PDF exports omitting hidden slides)
        if self._total_pages > 0 and self._total_pages == len(visible_indices):
            v_map = {slide_idx: page_idx for page_idx, slide_idx in enumerate(visible_indices)}
            for idx in range(total_slides):
                self._slide_to_pdf_page[idx] = v_map.get(idx, None)
            print(f"[SlideRenderEngine] Mapped {len(visible_indices)} visible slides to {self._total_pages} PDF pages (handled {total_slides - len(visible_indices)} hidden slides) ✓")
            return

        # Case B: PDF page count matches total slides count (1:1 mapping)
        if self._total_pages == total_slides:
            for idx in range(total_slides):
                self._slide_to_pdf_page[idx] = idx
            return

        # Case C: Fallback direct mapping for valid ranges
        for idx in range(total_slides):
            self._slide_to_pdf_page[idx] = idx if idx < self._total_pages else None

    def _get_pdf_page_index(self, slide_idx: int) -> Optional[int]:
        if not hasattr(self, "_slide_to_pdf_page") or not self._slide_to_pdf_page:
            if 0 <= slide_idx < self._total_pages:
                return slide_idx
            return None
        return self._slide_to_pdf_page.get(slide_idx, None)

    def has_rendered_slide(self, slide_idx: int) -> bool:
        has_doc = (self._pdf is not None) or (self._fitz_doc is not None)
        if not has_doc:
            return False
        p_idx = self._get_pdf_page_index(slide_idx)
        return (p_idx is not None) and (0 <= p_idx < self._total_pages)

    def load_presentation(self, pptx_path: Path, prs: Optional[Presentation] = None):
        self._current_path = pptx_path
        self._prs = prs
        self._slide_to_pdf_page = {}
        with _C_EXTENSION_LOCK:
            self._slide_cache.clear()
            self._pdf = None
            self._fitz_doc = None
            self._total_pages = 0

        try:
            self._cache_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

        cache_pdf = self._get_cache_path(pptx_path)
        # Check legacy ~/.cache/anki_study_suite if not yet in APP_DATA_DIR / 'slide_cache'
        if not (cache_pdf.exists() and cache_pdf.stat().st_size > 0):
            legacy_pdf = Path.home() / ".cache" / "anki_study_suite" / cache_pdf.name
            if legacy_pdf.exists() and legacy_pdf.stat().st_size > 0:
                try:
                    import shutil
                    shutil.copy2(legacy_pdf, cache_pdf)
                except Exception:
                    cache_pdf = legacy_pdf

        if cache_pdf.exists() and cache_pdf.stat().st_size > 0:
            with _C_EXTENSION_LOCK:
                try:
                    if HAS_PDFIUM:
                        self._pdf = pdfium.PdfDocument(str(cache_pdf))
                        self._total_pages = len(self._pdf)
                    elif HAS_FITZ:
                        self._fitz_doc = fitz.open(str(cache_pdf))
                        self._total_pages = len(self._fitz_doc)
                    self._recompute_slide_mapping()
                    print(f"[SlideRenderEngine] Loaded cached slide PDF ({self._total_pages} slides) ✓")
                    if self._on_ready:
                        self._on_ready(success=True, message=f"{self._total_pages} slide diagrams loaded from cache")
                    return
                except Exception as exc:
                    print(f"[SlideRenderEngine] Failed to open cached PDF: {exc}")

        self._is_rendering = True
        self._render_queue.put((pptx_path, cache_pdf))

    def _get_cache_path(self, pptx_path: Path) -> Path:
        try:
            self._cache_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        mtime = pptx_path.stat().st_mtime if pptx_path.exists() else 0
        key = hashlib.md5(f"{pptx_path.resolve()}_{mtime}".encode()).hexdigest()
        return self._cache_dir / f"{key}.pdf"

    def _sequential_worker(self):
        while True:
            try:
                item = self._render_queue.get()
                if item is None:
                    break
                pptx_path, out_pdf = item
                self._export_worker(pptx_path, out_pdf)
            except Exception as e:
                print(f"[SlideRenderEngine] Sequential worker exception: {e}")
            finally:
                self._render_queue.task_done()

    def _export_worker(self, pptx_path: Path, out_pdf: Path):
        success = False
        try:
            out_pdf.parent.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

        # 1. Try Microsoft PowerPoint AppleScript (exact working script from anki_study_suite.py.bak)
        try:
            script = f'''
            tell application "Microsoft PowerPoint"
                open (POSIX file "{pptx_path.resolve()}")
                set prs to active presentation
                save prs in (POSIX file "{out_pdf.resolve()}") as save as PDF
                close prs saving no
            end tell
            '''
            res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=90)
            if res.returncode == 0 and out_pdf.exists() and out_pdf.stat().st_size > 0:
                success = True
                print("[SlideRenderEngine] Rendered slides via Microsoft PowerPoint ✓")
        except Exception as exc:
            print(f"[SlideRenderEngine] PowerPoint export error: {exc}")

        # 2. Try Keynote AppleScript if PowerPoint failed (restored from anki_study_suite.py.bak)
        if not success:
            try:
                keynote_app = Path("/Applications/Keynote.app")
                sys_keynote_app = Path("/System/Applications/Keynote.app")
                if keynote_app.exists() or sys_keynote_app.exists():
                    script = f'''
                    tell application "Keynote"
                        set doc to open (POSIX file "{pptx_path.resolve()}")
                        export doc to (POSIX file "{out_pdf.resolve()}") as PDF
                        close doc saving no
                    end tell
                    '''
                    res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=90)
                    if res.returncode == 0 and out_pdf.exists() and out_pdf.stat().st_size > 0:
                        success = True
                        print("[SlideRenderEngine] Rendered slides via Keynote ✓")
            except Exception as exc:
                print(f"[SlideRenderEngine] Keynote export error: {exc}")

        # 3. Try LibreOffice headless (soffice) if PowerPoint and Keynote failed
        if not success:
            for bin_name in ["soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"]:
                try:
                    subprocess.run(
                        [bin_name, "--headless", "--convert-to", "pdf", str(pptx_path.resolve()), "--outdir", str(out_pdf.parent)],
                        capture_output=True, text=True, timeout=90
                    )
                    conv_file = out_pdf.parent / f"{pptx_path.stem}.pdf"
                    if conv_file.exists() and conv_file.stat().st_size > 0:
                        if conv_file != out_pdf:
                            conv_file.rename(out_pdf)
                        success = True
                        print("[SlideRenderEngine] Rendered slides via LibreOffice ✓")
                        break
                except Exception:
                    continue

        self._is_rendering = False
        if success and out_pdf.exists() and out_pdf.stat().st_size > 0:
            with _C_EXTENSION_LOCK:
                try:
                    if HAS_PDFIUM:
                        self._pdf = pdfium.PdfDocument(str(out_pdf))
                        self._total_pages = len(self._pdf)
                    elif HAS_FITZ:
                        self._fitz_doc = fitz.open(str(out_pdf))
                        self._total_pages = len(self._fitz_doc)
                    self._recompute_slide_mapping()
                    print(f"[SlideRenderEngine] Ready with {self._total_pages} rendered slides ✓")
                except Exception as exc:
                    print(f"[SlideRenderEngine] Error loading exported PDF: {exc}")
                    success = False

            if success:
                if self._on_ready:
                    self._on_ready(success=True, message=f"{self._total_pages} slide diagrams ready")
                return

        if self._on_ready:
            self._on_ready(success=False, message="Could not render slide diagrams (falling back to extracted shapes)")

    def get_slide_image(self, slide_idx: int, scale: float = 2.0) -> Optional[Image.Image]:
        with _C_EXTENSION_LOCK:
            if not self.has_rendered_slide(slide_idx):
                return None
            if slide_idx in self._slide_cache:
                return self._slide_cache[slide_idx]
            p_idx = self._get_pdf_page_index(slide_idx)
            if p_idx is None:
                return None
            try:
                img = None
                if HAS_PDFIUM and self._pdf is not None:
                    page = self._pdf[p_idx]
                    img = page.render(scale=scale).to_pil().convert("RGBA")
                elif HAS_FITZ and getattr(self, "_fitz_doc", None) is not None:
                    page = self._fitz_doc[p_idx]
                    mat = fitz.Matrix(scale, scale)
                    pix = page.get_pixmap(matrix=mat, alpha=True)
                    img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)

                if img is not None:
                    if len(self._slide_cache) > 30:
                        oldest_k = next(iter(self._slide_cache))
                        del self._slide_cache[oldest_k]
                    self._slide_cache[slide_idx] = img
                    return img
                return None
            except Exception as exc:
                print(f"[SlideRenderEngine] Failed to render slide {slide_idx} (PDF page {p_idx}): {exc}")
                return None

# ─────────────────────────────────────────────────────────────────────────────
# 6. SLIDE PARSING ENGINE
# ─────────────────────────────────────────────────────────────────────────────
class SlideParsEngine:
    @staticmethod
    def extract_slide_text(slide) -> str:
        text_lines: list[str] = []
        for shape in SlideParsEngine._get_flat_shapes(slide.shapes):
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                if txt:
                    text_lines.append(txt)

            if shape.has_table:
                for row in shape.table.rows:
                    row_txt = [c.text.strip() for c in row.cells if c.text.strip()]
                    if row_txt:
                        text_lines.append(" | ".join(row_txt))
        return "\n".join(text_lines)

    @staticmethod
    def extract_slide_data(slide) -> tuple[list[Image.Image], str]:
        images: list[Image.Image] = []
        text_lines: list[str] = []

        for shape in SlideParsEngine._get_flat_shapes(slide.shapes):
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                if txt:
                    text_lines.append(txt)

            if shape.has_table:
                for row in shape.table.rows:
                    row_txt = [c.text.strip() for c in row.cells if c.text.strip()]
                    if row_txt:
                        text_lines.append(" | ".join(row_txt))

            # Shape type 13: Picture
            if shape.shape_type == 13:
                try:
                    img_bytes = shape.image.blob
                    pil_img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
                    w, h = pil_img.size
                    crop_l = getattr(shape, "crop_left", 0.0) or 0.0
                    crop_t = getattr(shape, "crop_top", 0.0) or 0.0
                    crop_r = getattr(shape, "crop_right", 0.0) or 0.0
                    crop_b = getattr(shape, "crop_bottom", 0.0) or 0.0
                    if crop_l > 0 or crop_t > 0 or crop_r > 0 or crop_b > 0:
                        x1 = max(0, int(crop_l * w))
                        y1 = max(0, int(crop_t * h))
                        x2 = min(w, int((1.0 - crop_r) * w))
                        y2 = min(h, int((1.0 - crop_b) * h))
                        if x2 > x1 and y2 > y1:
                            pil_img = pil_img.crop((x1, y1, x2, y2))
                    images.append(pil_img)
                except Exception:
                    pass

        return images, "\n".join(text_lines)

    @staticmethod
    def _get_flat_shapes(shapes) -> list:
        flat = []
        for s in shapes:
            flat.append(s)
            if getattr(s, "shape_type", None) == 6:  # Group
                try:
                    flat.extend(SlideParsEngine._get_flat_shapes(s.shapes))
                except Exception:
                    pass
        return flat

    @staticmethod
    def extract_composite_diagrams(prs: Presentation, slide_idx: int, rendered_slide: Optional[Image.Image]) -> list[dict]:
        if rendered_slide is None or not prs:
            return []
        try:
            slide = prs.slides[slide_idx]
            sw = prs.slide_width
            sh = prs.slide_height
            rw, rh = rendered_slide.size

            all_shapes = SlideParsEngine._get_flat_shapes(slide.shapes)

            cores = []
            for shape in all_shapes:
                st = getattr(shape, "shape_type", None)
                has_tf = getattr(shape, "has_text_frame", False)
                text_content = (shape.text_frame.text if has_tf else "").lower()

                is_img_or_group = (st in (13, 6)) or (has_tf and any(w in text_content for w in ["fig", "diagram", "pathway"]))
                if is_img_or_group:
                    bx1 = max(0.0, float(shape.left) / float(sw))
                    by1 = max(0.0, float(shape.top) / float(sh))
                    bx2 = min(1.0, float(shape.left + shape.width) / float(sw))
                    by2 = min(1.0, float(shape.top + shape.height) / float(sh))
                    w_norm = bx2 - bx1
                    h_norm = by2 - by1
                    if w_norm >= 0.05 and h_norm >= 0.05 and (w_norm * h_norm) >= 0.005:
                        # Skip top-corner logo on title slide
                        if slide_idx == 0 and by1 < 0.12 and bx2 > 0.8:
                            continue
                        cores.append([bx1, by1, bx2, by2])

            if not cores:
                return []

            diagram_boxes = []
            for core in cores:
                cx1, cy1, cx2, cy2 = core
                expanded = [cx1, cy1, cx2, cy2]

                for s in all_shapes:
                    st = getattr(s, "shape_type", None)
                    # Skip slide title and body text placeholders
                    if st == 14:
                        continue
                    # Skip top header area (Title and ILO markers)
                    if s.top < 0.12 * sh:
                        continue
                    # Skip tiny footer citations at very bottom edge
                    if s.top > 0.94 * sh and s.height < 0.06 * sh:
                        continue
                    # Skip wide multi-column body text
                    if float(s.width) / float(sw) > 0.40:
                        continue
                    # Must be label, drawing, or connector line
                    if st not in (1, 5, 9, 17):
                        continue

                    sx1 = max(0.0, float(s.left) / float(sw))
                    sy1 = max(0.0, float(s.top) / float(sh))
                    sx2 = min(1.0, float(s.left + s.width) / float(sw))
                    sy2 = min(1.0, float(s.top + s.height) / float(sh))

                    horiz_dist = max(0.0, max(cx1 - sx2, sx1 - cx2))
                    vert_dist  = max(0.0, max(cy1 - sy2, sy1 - cy2))

                    # Proximity check: label / line is near this core
                    if horiz_dist <= 0.18 and vert_dist <= 0.08:
                        expanded[0] = min(expanded[0], sx1)
                        expanded[1] = min(expanded[1], sy1)
                        expanded[2] = max(expanded[2], sx2)
                        expanded[3] = max(expanded[3], sy2)

                # Add 2.5% safe margin
                pad_x = 0.025
                pad_y = 0.025
                fx1 = max(0.0, expanded[0] - pad_x)
                fy1 = max(0.0, expanded[1] - pad_y)
                fx2 = min(1.0, expanded[2] + pad_x)
                fy2 = min(1.0, expanded[3] + pad_y)
                diagram_boxes.append([fx1, fy1, fx2, fy2])

            # Deduplicate heavily overlapping boxes (IoU > 0.5)
            merged_boxes = []
            for b in sorted(diagram_boxes, key=lambda x: (x[0], x[1])):
                overlap = False
                for mb in merged_boxes:
                    ix1 = max(b[0], mb[0])
                    iy1 = max(b[1], mb[1])
                    ix2 = min(b[2], mb[2])
                    iy2 = min(b[3], mb[3])
                    if ix2 > ix1 and iy2 > iy1:
                        inter_area = (ix2 - ix1) * (iy2 - iy1)
                        b_area = (b[2] - b[0]) * (b[3] - b[1])
                        mb_area = (mb[2] - mb[0]) * (mb[3] - mb[1])
                        iou = inter_area / float(b_area + mb_area - inter_area)
                        if iou > 0.5:
                            mb[0] = min(mb[0], b[0])
                            mb[1] = min(mb[1], b[1])
                            mb[2] = max(mb[2], b[2])
                            mb[3] = max(mb[3], b[3])
                            overlap = True
                            break
                if not overlap:
                    merged_boxes.append(b)

            diagrams = []
            for i, b in enumerate(merged_boxes):
                rx1 = int(b[0] * rw)
                ry1 = int(b[1] * rh)
                rx2 = int(b[2] * rw)
                ry2 = int(b[3] * rh)
                if rx2 > rx1 and ry2 > ry1:
                    crop = rendered_slide.crop((rx1, ry1, rx2, ry2))
                    diagrams.append({
                        "name": f"Diagram {i + 1} (with Labels & Arrows)",
                        "image": crop,
                        "box": (rx1, ry1, rx2, ry2),
                        "is_custom": False
                    })
            return diagrams
        except Exception:
            return []

    @staticmethod
    def should_skip_slide_for_cards(slide_idx: int, slide_text: str) -> tuple[bool, str]:
        if slide_idx == 0:
            return True, "Slide 1 / Title Slide"

        t_clean = slide_text.strip().lower()

        # Skip Objectives / ILOs slides
        ilo_markers = [
            "intended learning objective",
            "learning outcomes",
            "learning objectives",
            "by the end of this lecture",
            "session aims",
            "session objectives",
        ]
        if any(marker in t_clean for marker in ilo_markers):
            return True, "Objectives / ILOs Slide (Blueprint only)"

        # Skip administrative or closing slides
        admin_markers = [
            "reading list", "further reading", "recommended reading",
            "textbook recommendations", "any questions?", "questions?",
            "thank you for listening", "thank you!", "acknowledgements",
            "mentimeter", "menti.com", "feedback form",
        ]
        if any(marker in t_clean for marker in admin_markers):
            return True, "Administrative / Closing Slide"

        # Skip introductory justifications and disciplinary overview slides
        intro_markers = [
            "why study", "why do we study", "clinical and scientific reasons for studying",
            "disciplines that benefit", "clinical disciplines that benefit",
            "scientific disciplines that benefit", "reasons for studying",
            "importance of studying", "introduction to the module",
            "module overview", "course overview",
        ]
        if any(marker in t_clean for marker in intro_markers):
            return True, "Introductory Justification Slide"

        if len(t_clean) < 35:
            return True, "Insufficient content (< 35 chars)"

        return False, ""

# ─────────────────────────────────────────────────────────────────────────────
# 7. OCR ENGINE (EasyOCR)
# ─────────────────────────────────────────────────────────────────────────────
class OCREngine:
    def __init__(self, on_ready_callback=None, min_confidence: float = 0.25):
        self._reader = None
        self._ready  = False
        self._init_error = ""
        self._on_ready = on_ready_callback
        self.min_confidence = min_confidence
        threading.Thread(target=self._init_worker, daemon=True).start()

    def _init_worker(self):
        try:
            with _C_EXTENSION_LOCK:
                import easyocr

                # Check for bundled pre-trained models in BUNDLE_DIR / .easyocr_models
                bundled_models = BUNDLE_DIR / ".easyocr_models"
                if bundled_models.exists() and (bundled_models / "craft_mlt_25k.pth").exists():
                    models_dir = bundled_models
                else:
                    models_dir = APP_DATA_DIR / "ocr_models"
                models_dir.mkdir(parents=True, exist_ok=True)

                network_dir = APP_DATA_DIR / "ocr_user_network"
                network_dir.mkdir(parents=True, exist_ok=True)

                self._reader = easyocr.Reader(
                    ["en"],
                    gpu=False,
                    model_storage_directory=str(models_dir),
                    user_network_directory=str(network_dir),
                    verbose=False
                )
                self._ready = True
                self._init_error = ""
            print(f"[OCREngine] EasyOCR reader initialized successfully from {models_dir} ✓")
            if self._on_ready:
                self._on_ready(True, "EasyOCR initialized")
        except Exception as e:
            self._ready = False
            self._init_error = str(e)
            print(f"[OCREngine] EasyOCR Init Error: {e}")
            traceback.print_exc()
            if self._on_ready:
                self._on_ready(False, f"OCR Init Error: {e}")

    @property
    def is_ready(self) -> bool:
        return self._ready

    def detect_text_boxes(self, image: Image.Image) -> list[dict]:
        if not self._ready or self._reader is None:
            print(f"[OCREngine] detect_text_boxes skipped: OCR not ready (error: {self._init_error})")
            return []
        if image is None:
            return []

        with _C_EXTENSION_LOCK:
            rgb_img = image.convert("RGB")
            orig_w, orig_h = rgb_img.size

            # Upscale small cropped diagrams so fine font labels are easily resolved by CRAFT
            scale_factor = 1.0
            max_dim = max(orig_w, orig_h)
            if max_dim < 1400:
                scale_factor = min(2.5, 1400.0 / float(max_dim))
                target_w = int(orig_w * scale_factor)
                target_h = int(orig_h * scale_factor)
                ocr_img = rgb_img.resize((target_w, target_h), Image.Resampling.LANCZOS)
            else:
                ocr_img = rgb_img

            import numpy as np
            np_arr = np.array(ocr_img)

            try:
                raw_results = self._reader.readtext(
                    np_arr,
                    min_size=5,
                    text_threshold=0.35,
                    low_text=0.25,
                    link_threshold=0.3,
                    mag_ratio=1.5,
                    slope_ths=0.1,
                    height_ths=0.6,
                    width_ths=0.6,
                    add_margin=0.15,
                )
            except Exception as ocr_err:
                print(f"[OCREngine] readtext exception: {ocr_err}")
                traceback.print_exc()
                return []

        boxes = []
        for bbox, text, prob in raw_results:
            text = text.strip()
            if not text or len(text) < 2 or prob < self.min_confidence:
                continue

            xs = [pt[0] / scale_factor for pt in bbox]
            ys = [pt[1] / scale_factor for pt in bbox]
            x1, y1 = max(0, int(round(min(xs)))), max(0, int(round(min(ys))))
            x2, y2 = min(orig_w, int(round(max(xs)))), min(orig_h, int(round(max(ys))))

            # Expand bounds slightly to cover text border cleanly
            pad = 4
            x1 = max(0, x1 - pad)
            y1 = max(0, y1 - pad)
            x2 = min(orig_w, x2 + pad)
            y2 = min(orig_h, y2 + pad)

            if (x2 - x1) >= 8 and (y2 - y1) >= 8:
                boxes.append({
                    "text": text,
                    "box": (x1, y1, x2, y2),
                    "prob": round(float(prob), 2),
                    "checked": True,
                    "is_redaction": False,
                })

        # Automatically merge adjacent multi-line stacked labels and broken word fragments
        merged_boxes = self.merge_nearby_label_boxes(boxes)

        print(f"[OCREngine] Detected {len(boxes)} raw boxes -> Auto-grouped into {len(merged_boxes)} cohesive labels (image: {orig_w}x{orig_h}, scaled: {scale_factor:.2f}x)")
        return merged_boxes

    _OCR_ADJECTIVES = {
        'active', 'inactive', 'primary', 'secondary', 'tertiary',
        'anterior', 'posterior', 'superior', 'inferior', 'medial', 'lateral',
        'superficial', 'deep', 'internal', 'external', 'proximal', 'distal',
        'ascending', 'descending', 'afferent', 'efferent', 'common', 'greater', 'lesser',
        'simple', 'stratified', 'pseudostratified', 'squamous', 'cuboidal', 'columnar',
        'transitional', 'ciliated', 'mesenchymal', 'hematopoietic', 'haematopoietic',
        'smooth', 'skeletal', 'cardiac', 'sympathetic', 'parasympathetic',
        'myelinated', 'unmyelinated', 'cortical', 'medullary', 'visceral', 'parietal',
        'mature', 'immature', 'resting', 'proliferating', 'hypertrophic', 'calcified',
        'articular', 'compact', 'spongy', 'cancellous', 'trabecular', 'bone-lining',
        'inner', 'outer', 'dorsal', 'ventral', 'rostral', 'caudal', 'cranial',
        'pulmonary', 'systemic', 'coronary', 'hepatic', 'renal', 'splenic', 'gastric',
        'femoral', 'radial', 'ulnar', 'brachial', 'axillary', 'subclavian', 'carotid',
        'jugular', 'temporal', 'facial', 'occipital', 'frontal',
        'cervical', 'thoracic', 'lumbar', 'sacral', 'coccygeal', 'spinal',
        'left', 'right', 'middle', 'central', 'peripheral', 'apical', 'basal'
    }

    _OCR_CONTINUATION_NOUNS = {
        'cell', 'cells', 'stem cell', 'stem cells', 'progenitor', 'progenitors',
        'layer', 'layers', 'zone', 'zones', 'stratum', 'lamina', 'laminae',
        'membrane', 'membranes', 'sheath', 'sheaths', 'tissue', 'tissues',
        'fiber', 'fibers', 'fibre', 'fibres', 'matrix', 'substance',
        'artery', 'arteries', 'vein', 'veins', 'vessel', 'vessels', 'sinus', 'sinuses',
        'nerve', 'nerves', 'plexus', 'ganglion', 'ganglia', 'tract', 'tracts',
        'muscle', 'muscles', 'tendon', 'tendons', 'ligament', 'ligaments',
        'duct', 'ducts', 'canal', 'canals', 'tubule', 'tubules',
        'gland', 'glands', 'follicle', 'follicles', 'node', 'nodes', 'corpuscle', 'corpuscles',
        'process', 'tubercle', 'tuberosity', 'crest', 'spine', 'notch', 'groove', 'margin', 'border', 'pole',
        'complex', 'system', 'plate', 'cavity', 'foramen', 'fossa', 'nucleus', 'trunk',
        'arch', 'cord', 'root', 'ramus', 'head', 'neck', 'body', 'tail', 'end'
    }

    _OCR_CELL_SUFFIXES = ('blast', 'blasts', 'cyte', 'cytes', 'clast', 'clasts', 'phage', 'phages', 'phil', 'phils')

    @classmethod
    def merge_nearby_label_boxes(cls, input_boxes: list[dict]) -> list[dict]:
        """
        Intelligently groups and merges fragmented OCR bounding boxes for diagram labels:
        1. Multi-line stacked labels (tight vertical spacing with semantic/grammatical continuation
           such as adjectives, parentheticals, continuation nouns, or hyphens).
        2. Broken word fragments on the same line (tight horizontal spacing with vertical overlap,
           strictly excluding separate multi-word side-by-side labels).
        3. Overlapping or subsumed duplicate boxes.
        Iteratively merges until all cohesive label fragments form unified occlusion masks.
        """
        if not input_boxes or len(input_boxes) < 2:
            return input_boxes

        boxes = [dict(b) for b in input_boxes]
        merged = True
        iterations = 0
        max_iterations = 60

        while merged and iterations < max_iterations:
            merged = False
            iterations += 1
            n = len(boxes)

            for i in range(n):
                for j in range(i + 1, n):
                    b1 = boxes[i]["box"]
                    b2 = boxes[j]["box"]
                    w1, h1 = max(1, b1[2] - b1[0]), max(1, b1[3] - b1[1])
                    w2, h2 = max(1, b2[2] - b2[0]), max(1, b2[3] - b2[1])
                    min_w = min(w1, w2)
                    min_h = min(h1, h2)

                    # Overlaps in X and Y
                    x_overlap = min(b1[2], b2[2]) - max(b1[0], b2[0])
                    y_overlap = min(b1[3], b2[3]) - max(b1[1], b2[1])

                    # Gaps between boxes (0 if overlapping)
                    y_gap = max(0, max(b1[1], b2[1]) - min(b1[3], b2[3]))
                    x_gap = max(0, max(b1[0], b2[0]) - min(b1[2], b2[2]))

                    # Center coordinates
                    cx1 = (b1[0] + b1[2]) / 2.0
                    cx2 = (b2[0] + b2[2]) / 2.0
                    cy1 = (b1[1] + b1[3]) / 2.0
                    cy2 = (b2[1] + b2[3]) / 2.0
                    center_x_dist = abs(cx1 - cx2)
                    center_y_dist = abs(cy1 - cy2)

                    # Check if one box is primarily inside another (subsumed or duplicate)
                    area1 = w1 * h1
                    area2 = w2 * h2
                    inter_area = max(0, x_overlap) * max(0, y_overlap) if (x_overlap > 0 and y_overlap > 0) else 0
                    is_subsumed = (
                        inter_area > 0 and (
                            inter_area / float(area1) >= 0.70 or
                            inter_area / float(area2) >= 0.70 or
                            (inter_area / float(area1 + area2 - inter_area)) >= 0.50
                        )
                    )

                    left_item, right_item = (boxes[i], boxes[j]) if b1[0] <= b2[0] else (boxes[j], boxes[i])
                    t_left = left_item.get("text", "").strip()
                    t_right = right_item.get("text", "").strip()

                    top_item, bot_item = (boxes[i], boxes[j]) if b1[1] <= b2[1] else (boxes[j], boxes[i])
                    t_top = top_item.get("text", "").strip()
                    t_bot = bot_item.get("text", "").strip()

                    # Case 1: Horizontal continuation (broken word fragment on the exact same line)
                    is_horizontal = False
                    if not is_subsumed:
                        # STRICT check: Side-by-side labels must NEVER merge horizontally.
                        # Only merge if it is a broken word fragment (e.g. 'osteobla' + 'sts' or ends in hyphen/slash).
                        if (x_gap <= 3 or x_overlap >= 0) and y_overlap >= min_h * 0.75 and center_y_dist <= 6:
                            ends_punct = t_left.endswith(("-", "/", "•", "·", "–"))
                            starts_punct = t_right.startswith(("-", "/", "•", "·", "–"))
                            is_multiword = (" " in t_left) or (" " in t_right)
                            is_fragment = (len(t_left) <= 3 or len(t_right) <= 3 or min_w <= 35)

                            if ends_punct or starts_punct:
                                is_horizontal = True
                            elif not is_multiword and is_fragment and x_gap <= 2:
                                is_horizontal = True

                    # Case 2: Vertical stacking (lines of the same multi-line label)
                    is_vertical = False
                    if not is_subsumed and not is_horizontal:
                        # Vertical proximity: gap must be small (within standard line spacing)
                        is_vertically_tight = (y_gap <= 10) or (y_overlap > 0 and y_overlap <= min_h * 0.40)

                        if is_vertically_tight:
                            overlap_ratio = float(x_overlap) / float(min_w) if min_w > 0 else 0
                            h_ratio = h1 / float(h2)
                            has_similar_h = (0.22 <= h_ratio <= 4.5)

                            if has_similar_h and (overlap_ratio >= 0.35 or center_x_dist <= min_w * 0.45):
                                left_diff = abs(b1[0] - b2[0])
                                center_diff = center_x_dist
                                is_left_aligned = left_diff <= max(12, min_w * 0.15)
                                is_center_aligned = center_diff <= max(14, min_w * 0.15)

                                t_top_lower = t_top.lower()
                                t_bot_lower = t_bot.lower()
                                top_words = re.findall(r"\b\w+\b", t_top_lower)
                                bot_words = re.findall(r"\b\w+\b", t_bot_lower)
                                last_top_word = top_words[-1] if top_words else ""
                                first_bot_word = bot_words[0] if bot_words else ""

                                # 1. Subtitle or abbreviation in parentheses e.g. (periosteal cells), (GMP, CFU-GM)
                                is_parenthetical = t_bot.startswith("(") or t_bot.startswith("[")

                                # 2. Punctuation continuation: top line ends in -, /, ,, and, or, of
                                ends_continuation = t_top.endswith(("-", "/", ",")) or last_top_word in {"and", "or", "of", "the", "in", "to", "for", "with"}

                                # 3. Grammatical adjective continuation: top line ends in an adjective (active, inactive, etc.)
                                is_adj_continuation = last_top_word in cls._OCR_ADJECTIVES

                                # 4. Grammatical noun continuation: bottom line is a generic continuation noun (cells, progenitor, layer, etc.)
                                is_noun_continuation = (
                                    first_bot_word in cls._OCR_CONTINUATION_NOUNS or
                                    t_bot_lower in cls._OCR_CONTINUATION_NOUNS or
                                    (" ".join(bot_words[:2]) in cls._OCR_CONTINUATION_NOUNS if len(bot_words) >= 2 else False)
                                )

                                # 5. Short fragment: top or bot is a tiny word piece (<= 3 chars)
                                is_fragment_stack = (len(t_top) <= 3 or len(t_bot) <= 3)

                                has_semantic_continuation = (
                                    is_parenthetical or
                                    ends_continuation or
                                    is_adj_continuation or
                                    is_noun_continuation or
                                    is_fragment_stack
                                )

                                # CRITICAL NEGATIVE CONSTRAINT:
                                # If BOTH top and bottom lines end in distinct specific cell types
                                # (e.g. osteoblasts vs osteocytes, chondroblasts vs chondrocytes),
                                # they are NEVER the same label!
                                both_cell_types = (
                                    last_top_word.endswith(cls._OCR_CELL_SUFFIXES) and
                                    bot_words and bot_words[-1].endswith(cls._OCR_CELL_SUFFIXES)
                                )

                                if has_semantic_continuation and not both_cell_types:
                                    if is_left_aligned or is_center_aligned or overlap_ratio >= 0.60:
                                        is_vertical = True

                    if is_subsumed or is_horizontal or is_vertical:
                        nx1 = min(b1[0], b2[0])
                        ny1 = min(b1[1], b2[1])
                        nx2 = max(b1[2], b2[2])
                        ny2 = max(b1[3], b2[3])

                        if is_vertical:
                            t1 = t_top
                            t2 = t_bot
                            if t2.lower() in t1.lower():
                                joined_text = t1
                            elif t1.lower() in t2.lower():
                                joined_text = t2
                            else:
                                joined_text = f"{t1} {t2}"
                        else:
                            t1 = t_left
                            t2 = t_right
                            if t2.lower() in t1.lower():
                                joined_text = t1
                            elif t1.lower() in t2.lower():
                                joined_text = t2
                            elif t1.endswith("-"):
                                joined_text = f"{t1}{t2}"
                            else:
                                joined_text = f"{t1}{t2}" if (x_gap <= 2 and not t1.endswith(" ")) else f"{t1} {t2}"

                        new_item = {
                            "text": joined_text,
                            "box": (nx1, ny1, nx2, ny2),
                            "prob": max(boxes[i].get("prob", 1.0), boxes[j].get("prob", 1.0)),
                            "checked": True,
                            "is_redaction": boxes[i].get("is_redaction", False) or boxes[j].get("is_redaction", False)
                        }

                        boxes.pop(j)
                        boxes.pop(i)
                        boxes.append(new_item)
                        merged = True
                        break
                if merged:
                    break

        # Sort remaining boxes in standard diagram reading order
        boxes.sort(key=lambda b: (b["box"][1] // 40, b["box"][0]))
        return boxes

# ─────────────────────────────────────────────────────────────────────────────
# 8. GEMINI CARDS ENGINE & NOTES ENGINE
# ─────────────────────────────────────────────────────────────────────────────
class GeminiCardsEngine:
    def __init__(self, config: ConfigManager):
        self._config = config
        self.last_model_used = ""
        self._cancel_requested = False
        self._notes_cancel_requested = False
        self._cards_cancel_requested = False

    def _call_gemini_rest(
        self,
        prompt: str,
        timeout: int = 120,
        status_callback: Optional[Any] = None,
        preferred_model: Optional[str] = None,
        task_label: str = "Notes",
        restrict_to_flagship: bool = False
    ) -> str:
        api_key = self._config.get_notes_api_key()
        if not api_key:
            raise ValueError("Gemini Notes API Key is not set. Open Settings to configure your key.")

        configured_model = (self._config.get("GEMINI_MODEL", "gemini-4-hydrogen") or "gemini-4-hydrogen").strip()
        model = (preferred_model or configured_model).strip()
        masked = mask_api_key(api_key)
        print(f"[Gemini {task_label}] Request initiated | Key: Notes Key ({masked}) | Target Model: {model} | Flagship Only: {restrict_to_flagship}")
        
        gen_config = {"temperature": 0.2, "maxOutputTokens": 8192}
        if restrict_to_flagship:
            gen_config["thinkingConfig"] = {"thinkingBudget": 2048}

        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": gen_config
        }

        # Task-calibrated fallback cascade:
        # Notes synthesis strictly restricts to flagship models (gemini-4-hydrogen & gemini-3.8-flash) with thinking_level="medium".
        # It never downgrades to gemini-3.5-flash-lite on 503 errors.
        if restrict_to_flagship:
            cascade = [
                "gemini-4-hydrogen",
                "gemini-3.8-flash",
            ]
        elif preferred_model == "gemini-3-flash-preview":
            cascade = [
                "gemini-3-flash-preview",
                "gemini-3.5-flash-lite",
                "gemini-3.5-flash",
                "gemini-flash-lite-latest",
            ]
        elif preferred_model == "gemini-4-hydrogen":
            cascade = [
                "gemini-4-hydrogen",
                "gemini-3.8-flash",
                "gemini-3.5-flash-lite",
                "gemini-3.5-flash",
                "gemini-flash-lite-latest",
                "gemini-3-flash-preview",
            ]
        elif preferred_model == "gemini-3.8-flash":
            cascade = [
                "gemini-3.8-flash",
                "gemini-3.5-flash-lite",
                "gemini-3.5-flash",
                "gemini-flash-lite-latest",
                "gemini-3-flash-preview",
            ]
        else:
            cascade = [
                model,
                "gemini-3.5-flash-lite",
                "gemini-3.5-flash",
                "gemini-3.8-flash",
                "gemini-flash-lite-latest",
                "gemini-3-flash-preview",
            ]
        models_to_try = []
        for m in cascade:
            if m and m not in models_to_try:
                models_to_try.append(m)

        last_error = ""
        backoff_delays = [2.0, 4.0, 8.0]
        call_timeout = min(timeout, 50)

        for idx, attempt_model in enumerate(models_to_try):
            is_primary = (idx == 0)
            max_retries = 2 if is_primary else 1

            if status_callback:
                status_callback(f"Connecting with Notes Key ({masked}) to {attempt_model}...", model=attempt_model)

            url = f"https://generativelanguage.googleapis.com/v1beta/models/{attempt_model}:generateContent?key={api_key}"
            for retry in range(max_retries):
                is_notes = "note" in task_label.lower()
                is_cards = any(k in task_label.lower() for k in ["card", "flashcard"])
                if is_notes:
                    if getattr(self, "_notes_cancel_requested", False) or getattr(self, "_cancel_requested", False):
                        raise InterruptedError("Notes generation cancelled by user.")
                elif is_cards:
                    if getattr(self, "_cards_cancel_requested", False):
                        raise InterruptedError("Card generation cancelled by user.")
                elif getattr(self, "_cancel_requested", False):
                    raise InterruptedError("Generation cancelled by user.")

                try:
                    res = requests.post(url, json=payload, timeout=call_timeout)
                    if res.status_code == 200:
                        data = res.json()
                        try:
                            candidate = data["candidates"][0]
                            parts = candidate.get("content", {}).get("parts", [])
                            text = parts[0].get("text", "") if parts else ""
                            finish_reason = candidate.get("finishReason", "")

                            # If truncated by token limit, automatically perform seamless continuation
                            if finish_reason == "MAX_TOKENS":
                                cont_payload = {
                                    "contents": [
                                        {"role": "user", "parts": [{"text": prompt}]},
                                        {"role": "model", "parts": [{"text": text}]},
                                        {"role": "user", "parts": [{"text": "Continue immediately from the exact point where you stopped. Do not repeat previous sentences, just continue generating the rest of the lecture notes until the entire lecture has been fully covered:"}]}
                                    ],
                                    "generationConfig": {"temperature": 0.2, "maxOutputTokens": 8192}
                                }
                                try:
                                    cont_res = requests.post(url, json=cont_payload, timeout=call_timeout)
                                    if cont_res.status_code == 200:
                                        cont_data = cont_res.json()
                                        cont_parts = cont_data.get("candidates", [{}])[0].get("content", {}).get("parts", [])
                                        cont_text = cont_parts[0].get("text", "") if cont_parts else ""
                                        text = (text + "\n\n" + cont_text).strip() if text else cont_text.strip()
                                except Exception:
                                    pass

                            if not text.strip():
                                raise ValueError("Empty response text from model.")

                            self.last_model_used = attempt_model
                            if attempt_model != model:
                                print(f"[Gemini Notes] Fallback model '{attempt_model}' succeeded with Notes Key ({masked}).")
                            else:
                                print(f"[Gemini Notes] Primary model '{model}' succeeded with Notes Key ({masked}).")
                            return text
                        except Exception as e:
                            raise ValueError(f"Malformed Gemini response ({attempt_model}): {e} | data={data}")
                    elif res.status_code == 503:
                        last_error = f"Gemini Error 503 ({attempt_model}): {res.text}"
                        if restrict_to_flagship:
                            # Flagship restricted: NEVER downgrade to flash-lite! Cycle between 4-hydrogen & 3.8-flash
                            alt_idx = (idx + 1) % len(models_to_try)
                            next_model = models_to_try[alt_idx]
                            msg = f"⚠️ {attempt_model} flagship slot busy (503). Retrying with {next_model}..."
                            print(f"[Gemini Notes] {msg}")
                            if status_callback:
                                status_callback(msg, model=next_model)
                            time.sleep(3.0)
                            continue
                        # If model is overloaded on Google's servers, fail over immediately to next available model
                        if not is_primary or retry >= 1:
                            next_model = models_to_try[idx + 1] if idx + 1 < len(models_to_try) else "next candidate"
                            msg = f"⚠️ {attempt_model} is experiencing high demand (503). Switching immediately to {next_model}..."
                            print(f"[Gemini Notes] {msg}")
                            if status_callback:
                                status_callback(msg, model=next_model)
                            break
                        wait_time = 2.0
                        msg = f"{attempt_model} is experiencing high demand (503). Retrying once in {int(wait_time)}s..."
                        print(f"[Gemini Notes] {msg}")
                        if status_callback:
                            status_callback(msg, model=attempt_model)
                        time.sleep(wait_time)
                        continue
                    elif res.status_code == 400 and "thinkingConfig" in res.text:
                        payload["generationConfig"].pop("thinkingConfig", None)
                        continue
                    elif res.status_code == 429:
                        last_error = f"Gemini Error 429 ({attempt_model}): {res.text}"
                        is_daily_limit = any(term in res.text for term in [
                            "limit: 20",
                            "requests per day limit exceeded",
                            "GenerateRequestsPerDay",
                            "generate_content_free_tier_requests"
                        ])
                        if is_daily_limit:
                            next_model = models_to_try[idx + 1] if idx + 1 < len(models_to_try) else "next candidate"
                            msg = f"⚠️ {attempt_model} Free Tier daily limit reached (20 req/day). Switching to {next_model}..."
                            print(f"[Gemini Notes] {msg}")
                            if status_callback:
                                status_callback(msg, model=next_model)
                            break

                        # Extract retryDelay if returned by Gemini (e.g. "30s" or "30.415673054s")
                        retry_delay_match = re.search(r'"retryDelay":\s*"(\d+(?:\.\d+)?)s?"', res.text)
                        if retry_delay_match:
                            wait_time = float(retry_delay_match.group(1)) + 1.0
                        else:
                            wait_time = backoff_delays[min(retry, len(backoff_delays) - 1)]

                        # If rate limit wait is long and we have fallback models, switch immediately
                        if wait_time > 8.0 and idx + 1 < len(models_to_try):
                            next_model = models_to_try[idx + 1]
                            msg = f"⚠️ {attempt_model} rate-limited ({int(wait_time)}s reset). Switching immediately to {next_model}..."
                            print(f"[Gemini Notes] {msg}")
                            if status_callback:
                                status_callback(msg, model=next_model)
                            break

                        msg = f"{attempt_model} rate-limited. Pausing {int(wait_time)}s (attempt {retry + 1}/{max_retries})..."
                        print(f"[Gemini Notes] {msg}")
                        if status_callback:
                            status_callback(msg, model=attempt_model)
                        time.sleep(wait_time)
                        continue
                    elif res.status_code == 404:
                        last_error = f"Gemini Error 404: {attempt_model} not found on v1beta API"
                        print(f"[Gemini] {attempt_model} returned 404 Not Found. Skipping to next model in cascade...")
                        break
                    else:
                        last_error = f"Gemini API Error {res.status_code} ({attempt_model}): {res.text}"
                        print(f"[Gemini] {attempt_model} returned {res.status_code}. Trying next candidate...")
                        break
                except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
                    last_error = f"Timeout/Connection Error ({attempt_model}): {exc}"
                    next_model = models_to_try[idx + 1] if idx + 1 < len(models_to_try) else "next candidate"
                    msg = f"⚠️ {attempt_model} connection timed out. Failing over to {next_model}..."
                    print(f"[Gemini Notes] {msg}")
                    if status_callback:
                        status_callback(msg, model=next_model)
                    break
                except Exception as exc:
                    last_error = f"{type(exc).__name__}: {exc}"
                    time.sleep(1.0)
                    continue

        raise ValueError(last_error or "Gemini API failed after retries.")

    def build_lecture_blueprint(self, prs: Presentation, lecture_name: str) -> str:
        ilo_text_parts = []
        for i in range(min(5, len(prs.slides))):
            _, text = SlideParsEngine.extract_slide_data(prs.slides[i])
            t_low = text.lower()
            if any(m in t_low for m in ["objective", "outcome", "aim", "end of this"]):
                ilo_text_parts.append(text)
        return "\n".join(ilo_text_parts)

    def _clean_generated_cards(self, cards: list[dict], default_slide: int = 1) -> list[dict]:
        cleaned_cards = []
        for c in cards:
            c.setdefault("slide", default_slide)
            ctype = c.get("type", "1. Basic")
            front = c.get("front", "").strip()
            back = c.get("back", "").strip()

            front = re.sub(r"\s*\[Lecture:[^\]]*\]", "", front).strip()
            back = re.sub(r"\s*\[Lecture:[^\]]*\]", "", back).strip()

            f_low = front.lower()
            b_low = back.lower()
            combined_low = f"{f_low} {b_low}"
            if any(k in combined_low for k in [
                "why study", "why do we study", "benefit from the study",
                "disciplines that benefit", "reasons for studying",
                "benefit from studying", "benefit of studying",
                "clinical and scientific disciplines", "disciplines benefit",
                "clinical disciplines", "scientific disciplines",
                "learning objective", "learning outcome", "module code",
                "who won", "nobel prize"
            ]):
                continue

            # Reject anatomy part-listing / rote structure naming (tested via Image Occlusion)
            if any(k in f_low for k in [
                "list the parts", "what are the parts", "name the parts",
                "parts of the", "list the structures", "name the structures",
                "what are the structures that", "list the branches of",
                "name the branches of", "what are the branches of",
                "components of the", "list the components", "name the components",
                "list the muscles of", "name the muscles of", "what muscles make up",
                "list the bones of", "name the bones of"
            ]):
                continue

            has_cloze = bool(re.search(r"\{\{c\d+::", front))
            if ctype == "Cloze++" and not has_cloze:
                ctype = "1. Basic"
            elif ctype == "1. Basic" and has_cloze:
                ctype = "Cloze++"

            c["type"] = ctype
            c["front"] = front
            c["back"] = back
            cleaned_cards.append(c)
        return cleaned_cards

    def generate_cards_from_slides_batch(
        self,
        slides_batch: list[dict],
        lecture_name: str,
        blueprint: str = "",
        status_callback: Optional[Any] = None
    ) -> list[dict]:
        slides_content = []
        for s in slides_batch:
            slides_content.append(f"--- SLIDE {s['slide_num']} ---\n{s['text'].strip()}")
        combined_text = "\n\n".join(slides_content)

        prompt = f"""
You are an expert Medical Education Assistant generating high-yield MBBS flashcards.
Lecture: "{lecture_name}"
Content from slides:
\"\"\"{combined_text}\"\"\"

STRICT NEGATIVE CONSTRAINTS (NEVER generate cards for these):
1. Historical / Biographical Trivia:
   - NEVER ask who won a Nobel Prize or what year a technique was pioneered.
2. Administrative Metadata & Course Details:
   - NEVER ask about module codes (e.g. 4MBBS102), lecturer names, lecture titles, or dates.
3. Learning Objectives & Aims:
   - NEVER ask what the learning objectives or ILOs of the lecture are.
4. Disciplinary Justifications & 'Why Study' Questions:
   - NEVER ask why a medical subject or field is studied, what clinical/scientific disciplines benefit from studying it, or career/specialty overviews.
   - If a slide contains introductory justification or disciplinary benefits, DO NOT generate any cards for it.
5. Anatomy Topics — Strictly NO Listing of Anatomical Parts / Structures:
   - NEVER generate cards asking to list anatomical parts, structures, or components (e.g. 'List the parts of...', 'What are the parts of the pancreas?', 'Name the muscles of the anterior thigh', 'What are the 5 branches of...').
   - Pure anatomical structure naming and part identification is ALREADY tested via Image Occlusion in Anki.
   - If the slide is on anatomy, ONLY generate cards testing:
     * Spatial and anatomical relations (anterior/posterior/medial/lateral/deep/superficial).
     * Neurovascular supply (motor/sensory innervation, nerve roots, arterial supply, venous drainage).
     * Lymphatic drainage basins and fascial planes/compartments.
     * Functional biomechanics / muscle actions (synergists, antagonists).
     * Clinical correlates, entrapment neuropathies, fracture vulnerabilities, clinical tests, surgical landmarks.

EXAMINABLE MEDICAL FOCUS ONLY:
- Anatomical relations, neurovascular supply, lymphatics, fascial planes, clinical correlates (NO rote part-listing)
- Physiological mechanisms, timelines, hormones, receptor interactions
- Embryological milestones, germ layers, cell lineage, compaction/gastrulation
- Clinical correlates, congenital anomalies, teratogens, diagnostic signs

Note Type Distribution Rules:
- '1. Basic' (70% - 80%): Direct medical questions or categorical lists.
- 'Cloze++' (20% - 30%): Strictly for directional facts, conversions (A -> B), numbers, or reciprocal pairs: '{{{{c1::keyword}}}}'.

Rules:
- NEVER duplicate front question in back answer.
- Do NOT append lecture source tags or slide numbers (NEVER append '[Lecture: ..., Slide ...]'). Keep answers clean and concise.
- For each card, set "slide" to the exact slide number where the concept appears.

Return a valid JSON array of objects:
[
  {{
    "type": "1. Basic",
    "front": "What is ...?",
    "back": "Concise answer / mechanism.",
    "slide": {slides_batch[0]['slide_num']}
  }}
]
"""
        raw_text = self._call_gemini_rest(
            prompt,
            preferred_model="gemini-3-flash-preview",
            task_label="Card Creation (Batch)",
            status_callback=status_callback
        )
        m = re.search(r"\[.*\]", raw_text, re.DOTALL)
        if m:
            try:
                cards = json.loads(m.group(0))
                return self._clean_generated_cards(cards, default_slide=slides_batch[0]["slide_num"])
            except Exception:
                pass
        return []

    def generate_cards_from_text(
        self,
        slide_text: str,
        slide_num: int,
        lecture_name: str,
        blueprint: str = "",
        max_cards: int = 3,
        status_callback: Optional[Any] = None,
    ) -> list[dict]:
        prompt = f"""
You are an expert Medical Education Assistant generating high-yield MBBS flashcards.
Lecture: "{lecture_name}"
Slide Number: {slide_num}
Content:
\"\"\"{slide_text}\"\"\"

STRICT NEGATIVE CONSTRAINTS (NEVER generate cards for these):
1. Historical / Biographical Trivia:
   - NEVER ask who won a Nobel Prize or what year a technique was pioneered.
2. Administrative Metadata & Course Details:
   - NEVER ask about module codes (e.g. 4MBBS102), lecturer names, lecture titles, or dates.
3. Learning Objectives & Aims:
   - NEVER ask what the learning objectives or ILOs of the lecture are.
4. Disciplinary Justifications & 'Why Study' Questions:
   - NEVER ask why a medical subject or field is studied, what clinical/scientific disciplines benefit from studying it, or career/specialty overviews.
   - If a slide contains introductory justification or disciplinary benefits, DO NOT generate any cards for it; return an empty array [].
5. Anatomy Topics — Strictly NO Listing of Anatomical Parts / Structures:
   - NEVER generate cards asking to list anatomical parts, structures, or components (e.g. 'List the parts of...', 'What are the parts of the pancreas?', 'Name the muscles of the anterior thigh', 'What are the 5 branches of...').
   - Pure anatomical structure naming and part identification is ALREADY tested via Image Occlusion in Anki.
   - If the slide is on anatomy, ONLY generate cards testing:
     * Spatial and anatomical relations (anterior/posterior/medial/lateral/deep/superficial).
     * Neurovascular supply (motor/sensory innervation, nerve roots, arterial supply, venous drainage).
     * Lymphatic drainage basins and fascial planes/compartments.
     * Functional biomechanics / muscle actions (synergists, antagonists).
     * Clinical correlates, entrapment neuropathies, fracture vulnerabilities, clinical tests, surgical landmarks.

EXAMINABLE MEDICAL FOCUS ONLY:
- Anatomical relations, neurovascular supply, lymphatics, fascial planes, clinical correlates (NO rote part-listing)
- Physiological mechanisms, timelines, hormones, receptor interactions
- Embryological milestones, germ layers, cell lineage, compaction/gastrulation
- Clinical correlates, congenital anomalies, teratogens, diagnostic signs

Note Type Distribution Rules:
- '1. Basic' (70% - 80%): Direct medical questions or categorical lists.
- 'Cloze++' (20% - 30%): Strictly for directional facts, conversions (A -> B), numbers, or reciprocal pairs: '{{{{c1::keyword}}}}'.

Rules:
- NEVER duplicate front question in back answer.
- Do NOT append lecture source tags or slide numbers (NEVER append '[Lecture: ..., Slide ...]'). Keep answers clean and concise.

Return a valid JSON array of objects:
[
  {{
    "type": "1. Basic",
    "front": "What is ...?",
    "back": "Concise answer / mechanism.",
    "slide": {slide_num}
  }}
]
"""
        raw_text = self._call_gemini_rest(
            prompt,
            preferred_model="gemini-3-flash-preview",
            task_label="Card Creation",
            status_callback=status_callback
        )
        m = re.search(r"\[.*\]", raw_text, re.DOTALL)
        if m:
            try:
                cards = json.loads(m.group(0))
                return self._clean_generated_cards(cards, default_slide=slide_num)
            except Exception:
                pass
        return []

    def refine_cards_with_instruction(self, existing_cards: list[dict], instruction: str) -> list[dict]:
        prompt = f"""
You are an MBBS Flashcard Editor adhering to strict GEMINI.md rules.
Existing Cards JSON:
{json.dumps(existing_cards, indent=2)}

User Instruction / Correction:
"{instruction}"

Strict Rules:
- NEVER include source attribution tags like '[Lecture: ..., Slide ...]' anywhere.
- NEVER include questions about module codes, learning objectives (ILOs), Nobel prizes, or 'Why study [subject]' questions.
- NEVER generate cards asking to list anatomical parts, structures, or components (anatomical structure identification is strictly tested via Image Occlusion in Anki). Focus on relations, innervation, blood supply, fascial planes, and clinical correlates.
- Prefer '1. Basic' note type for standard Q&A. Use 'Cloze++' only for clozes.

Apply the requested changes and output the complete updated JSON array. Return ONLY the JSON array.
"""
        raw_text = self._call_gemini_rest(prompt, preferred_model="gemini-3-flash-preview", task_label="Card Creation")
        m = re.search(r"\[.*\]", raw_text, re.DOTALL)
        if m:
            try:
                cards = json.loads(m.group(0))
                cleaned = []
                for c in cards:
                    f = re.sub(r"\s*\[Lecture:[^\]]*\]", "", c.get("front", "")).strip()
                    b = re.sub(r"\s*\[Lecture:[^\]]*\]", "", c.get("back", "")).strip()
                    f_low = f.lower()
                    b_low = b.lower()
                    combined_low = f"{f_low} {b_low}"
                    if any(k in combined_low for k in [
                        "why study", "why do we study", "benefit from the study",
                        "disciplines that benefit", "reasons for studying",
                        "benefit from studying", "benefit of studying",
                        "clinical and scientific disciplines", "disciplines benefit",
                        "clinical disciplines", "scientific disciplines",
                        "learning objective", "learning outcome", "module code",
                        "who won", "nobel prize"
                    ]):
                        continue

                    # Reject anatomy part-listing / rote structure naming (tested via Image Occlusion)
                    if any(k in f_low for k in [
                        "list the parts", "what are the parts", "name the parts",
                        "parts of the", "list the structures", "name the structures",
                        "what are the structures that", "list the branches of",
                        "name the branches of", "what are the branches of",
                        "components of the", "list the components", "name the components",
                        "list the muscles of", "name the muscles of", "what muscles make up",
                        "list the bones of", "name the bones of"
                    ]):
                        continue

                    c["front"] = f
                    c["back"] = b
                    cleaned.append(c)
                return cleaned
            except Exception:
                pass
        return existing_cards

class GeminiNotesEngine:
    def __init__(self, state: AppState, config: ConfigManager, cards_engine: "GeminiCardsEngine" = None):
        self._state  = state
        self._config = config
        self._cards_engine = cards_engine

    def format_slide_notes(self, slide_text: str, slide_num: int) -> str:
        return f"### Slide {slide_num} Key Takeaways\n\n" + "\n".join(f"- {line}" for line in slide_text.split("\n") if line.strip())

    @staticmethod
    def _build_slides_summary_str(slides_slice: list[dict]) -> str:
        s_str = ""
        for s in slides_slice:
            is_title_or_admin = (s['slide_num'] == 1) or any(k in (s['title'] + " " + s['body']).lower() for k in [
                "intended learning", "learning outcome", "learning objective", "aims of this lecture", "aims",
                "table of contents", "outline", "reading list", "mentimeter", "acknowledgements"
            ])
            s_str += f"\n--- SLIDE {s['slide_num']}: {s['title']} ---\n"
            if s.get("has_diagram") and not is_title_or_admin:
                s_str += f"[Visual Content: {s['image_count']} diagram/atlas figure(s) on this slide. MANDATORY: Emit '(slide {s['slide_num']}: ...)' marker with caption and narration.]\n"
            if s['body']:
                s_str += f"Slide Text:\n{s['body']}\n"
            if s['notes']:
                s_str += f"Speaker Notes:\n{s['notes']}\n"
        return s_str

    def generate_lecture_notes(self, prs: Presentation, lecture_name: str, status_callback: Optional[Any] = None) -> tuple[str, str]:
        if not prs:
            raise ValueError("No presentation loaded to generate notes from.")

        # Extract all slide text and speaker notes
        slides_data = []
        ilo_text = ""
        for i, slide in enumerate(prs.slides):
            raw_imgs, text = SlideParsEngine.extract_slide_data(slide)
            title = ""
            try:
                if slide.shapes.title and slide.shapes.title.text:
                    title = slide.shapes.title.text.strip()
            except Exception:
                pass
            if not title and text:
                title = text.split("\n")[0].strip()
            if not title:
                title = f"Slide {i + 1}"

            notes_slide = ""
            try:
                if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                    notes_slide = slide.notes_slide.notes_text_frame.text.strip()
            except Exception:
                pass

            t_low = (title + " " + text).lower()
            if not ilo_text and any(k in t_low for k in ["intended learning", "learning outcome", "learning objective", "objectives", "aims of this lecture", "aims"]):
                ilo_text = text.strip()

            slides_data.append({
                "slide_num": i + 1,
                "title": title,
                "body": text.strip(),
                "notes": notes_slide,
                "has_diagram": len(raw_imgs) > 0,
                "image_count": len(raw_imgs)
            })

        total_slides_count = len(prs.slides)
        all_topics_list = [f"Slide {s['slide_num']}: {s['title']}" for s in slides_data if s['title']]
        topics_summary = "\n".join(f"- {t}" for t in all_topics_list)

        # Style Preference: Default MBBS Style vs. Personalised Notion Templates
        style_mode = self._config.get("NOTES_STYLE_MODE", "default")
        custom_templates = list(self._config.get("CUSTOM_NOTE_TEMPLATES", []) or [])
        custom_instructions = (self._config.get("CUSTOM_STYLE_INSTRUCTIONS") or "").strip()
        is_custom = (style_mode == "custom" and len(custom_templates) > 0)

        if is_custom:
            templates_block = ""
            for idx, tpl in enumerate(custom_templates[:3]):
                t_name = tpl.get("name") or tpl.get("filename") or f"Template {idx + 1}"
                t_content = tpl.get("content", "").strip()
                if t_content:
                    templates_block += f"\n--- USER REFERENCE TEMPLATE {idx + 1}: {t_name} ---\n{t_content[:15000]}\n"

            style_guide_section = f"""### PERSONALIZED USER NOTION STYLE (CUSTOM TEMPLATE MODE ACTIVE)
The student has provided {len(custom_templates[:3])} real lecture note templates showing EXACTLY how they format and produce notes in Notion:
{templates_block}

### USER'S PERSONAL STYLE GUIDELINES:
{custom_instructions if custom_instructions else 'Carefully replicate the layout, heading style, bullet density, callout conventions, and phrasing demonstrated in the user templates above.'}

### ADAPTATION RULES:
1. Replicate the structural framework, section layout, and visual flow from the user's template examples above instead of the default layout.
2. If the user uses specific section dividers, emojis, table summaries, bold/underline patterns, or callout formats in their templates, adopt them faithfully.
3. Crucial Diagram Capture: Wherever slides contain key diagrams or visual atlas figures, you MUST still output `(slide N: what the diagram shows)` on its own line followed by a 1-line user revision caption so Synapse can crop and embed the diagram in Notion."""
        else:
            style_guide = ""
            style_guide_path = BUNDLE_DIR / "my-lecture-notes.md"
            if style_guide_path.exists():
                try:
                    style_guide = style_guide_path.read_text(encoding="utf-8")
                except Exception:
                    pass

            examples_text = ""
            examples_path = BUNDLE_DIR / "references" / "style-examples.md"
            if examples_path.exists():
                try:
                    examples_text = examples_path.read_text(encoding="utf-8")
                except Exception:
                    pass

            ref_examples_block = f"Here are reference examples of the user's real notes:\n{examples_text}\n" if examples_text else ""

            style_guide_section = f"""Here is the exact style guide and rules you MUST follow (from my-lecture-notes.md):
{style_guide if style_guide else '''
Page layout:
## 🎯 Learning Objectives
(the slide's objectives, copied verbatim)

## 📝 Notes
(everything else)

Two modes, chosen by content:
- Mechanisms, pathways, sequences of steps become plain bullets that stay close to the slide's own wording, one step per bullet.
- Concepts and comparisons become short paragraphs in plain, slightly informal English, explaining cause and effect.

Explain the why: Where a slide states a result, add 1-2 sentences that make it make sense. Put this in the user's voice without marking it as an addition.
Explain complex words and phrases in brackets: Whenever you use or encounter complex, obscure, or hard-to-understand medical, biochemical, anatomical, or pharmacological words or phrases, provide a brief description ~1 to 2 sentences long in brackets (...) right after the term explaining what it means in plain English.
Describe diagrams in words: When a slide's point is a diagram, narrate what happens step by step in a paragraph. If you cannot read it or to anchor visual revision, output a diagram reference on its own line:
(slide N: what the diagram shows)
followed by user revision caption ("what to notice, what to ignore, what to memorize").
Crucial: Always write single slide numbers per reference line, e.g. (slide 3: ...) and on next line (slide 4: ...), never group multiple slides like (slide 3 and 4: ...).
Memory hooks: 1-2 per lecture (analogy, one-line rule, nickname in italics).
What to leave out: drop worked numerical examples, data tables, pure-recall strings, historical names/dates, derivations, context/title slides.
Formatting in Notion:
- Blank line between every block: heading, paragraph, bullet list, screenshot marker, caption.
- Before each major block (<b><u>...</u></b>), add an empty paragraph, written <p /> on its own line with a blank line either side.
- Topic headings inside Notes are underlined plain text (<u>Topic:</u>), not ## headings.
- Plain text formulas: V0 = Vmax[S] / Km + [S].
'''}

{ref_examples_block}"""

        shared_instructions = f"""
CRITICAL COMPLETENESS, THEORETICAL RIGOR & DIAGRAM CAPTURE INSTRUCTIONS:
1. MANDATORY FULL SEQUENTIAL LECTURE COVERAGE (NO SKIPPED SLIDES, NO BRIEF NOTES):
- You MUST step through the lecture following the exact natural order of the slides.
- DO NOT collapse multiple distinct slides into a brief 1-line bullet or general overview! Thoroughly cover the medical substance of all slides without skipping.
- Extract full medical details: anatomical boundaries, relationships, muscle attachments, neurovascular supply, physiological functions, and clinical correlates.
- Never write brief high-level summaries. Detailed, complete, and exhaustive for Year 1 MBBS examinations.

2. EXHAUSTIVE ANATOMICAL, HISTOLOGICAL & PHYSIOLOGICAL GRANULARITY:
- Explicitly detail every individual vessel and parent artery/vein (distinguish superior vs inferior surfaces, specific branches, drainage into azygos/IVC/internal thoracic).
- Explicitly list exact vertebral levels, boundaries, ligaments, and muscle crura.
- Detail every structure transmitted through every aperture, foramen, or hiatus.

3. MANDATORY DIAGRAM REFERENCE MARKERS FOR ALL DISTINCT VISUALS:
- Emit explicit diagram reference on its own line:
(slide N: what the diagram shows)
followed by a 1-line user revision caption on the next line ("what to notice, what to ignore, what to memorize").
- STRICT PROHIBITION: NEVER emit a diagram reference for Slide 1 (Title slide with university logos/module codes) or administrative/ILO slides. Only emit markers for content slides (Slide 2+).
- Visual deduplication rule: If consecutive slides repeat the exact same illustration, emit only ONE diagram marker under the section where first introduced.
- Anchor distinct diagrams individually under their relevant sections.

4. CLINICAL MECHANISMS, CAUSE-AND-EFFECT DEPTH & MEMORY HOOKS:
- Explain the functional cause-and-effect mechanism for every concept.
- Include 1-2 punchy memory hooks or mnemonics per major lecture section.

5. COMPLEX WORDS & PHRASES:
Whenever encountering any complex, technical, or obscure medical/anatomical words or phrases, provide a brief description ~1 to 2 sentences long in brackets `(...)` right after the term explaining what it means in simple English.

6. NOTION STRUCTURE & SPACING:
- Before each major block (<b><u>Major Block</u></b>), add an empty paragraph, written `<p />` on its own line with a blank line either side.
- Topic headings inside Notes are underlined plain text (`<u>Topic:</u>`), not ## headings.
- Use bold `**term:**` for bullet points to highlight key concepts cleanly.
"""

        engine = self._cards_engine or GeminiCardsEngine(self._config)

        # ── 2-PASS CHUNKED SYNTHESIS (FOR DECKS > 12 SLIDES) ───────────────────
        # Splits long decks into 2 sequential chunks to prevent 503 overload and maintain flagship quality.
        if total_slides_count > 12:
            half_idx = total_slides_count // 2
            slides_part1_str = self._build_slides_summary_str(slides_data[:half_idx])
            slides_part2_str = self._build_slides_summary_str(slides_data[half_idx:])

            # ── PASS 1: Learning Objectives + Slides 1 to N//2 ──
            if status_callback:
                status_callback(f"Generating Part 1/2: Learning Objectives & Slides 1–{half_idx}...", model="gemini-4-hydrogen")

            prompt_pass1 = f"""You are the student's personal medical study partner writing high-yield lecture notes for King's College London MBBS Year 1.
Your task is to turn these lecture slides into notes that read as if the user wrote them: lean, in their own words, built for revision rather than completeness.

{style_guide_section}

Lecture Title: "{lecture_name}"
Total Slides in Deck: {total_slides_count}
GENERATION PHASE: PART 1 OF 2 (Learning Objectives & Slides 1 through {half_idx})

COMPLETE LECTURE SYLLABUS:
{topics_summary}

Extracted Slide Content for Part 1 (Slides 1 through {half_idx}):
{slides_part1_str}

{shared_instructions}

PART 1 EXECUTION RULES:
1. Begin with the complete, unified ## 🎯 Learning Objectives for the ENTIRE lecture (synthesized directly from lecture slides & syllabus).
2. Follow with ## 📝 Notes covering Slides 1 through {half_idx} in strict sequential order.
3. Output ONLY raw markdown starting with ## 🎯 Learning Objectives.
"""

            raw_part1 = engine._call_gemini_rest(
                prompt_pass1,
                timeout=180,
                status_callback=status_callback,
                preferred_model="gemini-4-hydrogen",
                task_label="Notes Part 1/2",
                restrict_to_flagship=True
            ).strip()

            if getattr(engine, "_notes_cancel_requested", False) or getattr(engine, "_cancel_requested", False):
                raise InterruptedError("Generation cancelled by user.")

            # Extract unified Learning Objectives from Part 1
            ilo_match = re.search(r"##\s*(?:🎯\s*)?Learning Objectives(.*?)(?=##\s*(?:📝\s*)?Notes|\Z)", raw_part1, re.DOTALL | re.IGNORECASE)
            ilo_block = ilo_match.group(1).strip() if ilo_match else (ilo_text or "Understand the anatomical structures, pathways, and physiological mechanisms.")

            # Extract Part 1 Notes body
            notes_match = re.search(r"##\s*(?:📝\s*)?Notes(.*)", raw_part1, re.DOTALL | re.IGNORECASE)
            if notes_match:
                notes_part1 = notes_match.group(1).strip()
            else:
                notes_part1 = re.sub(r"^##\s*(?:🎯\s*)?Learning Objectives.*?\n\n", "", raw_part1, flags=re.DOTALL | re.IGNORECASE).strip()

            # Extract covered section headings from Part 1 to provide continuity context for Pass 2
            covered_headings = re.findall(r"(?:<b><u>|<u>Topic:)(.*?)(?:</u>|</u></b>|:)", notes_part1)
            part1_outline = ", ".join(h.strip() for h in covered_headings[:8] if h.strip()) or f"Slides 1 to {half_idx}"

            # ── PASS 2: Slides (N//2)+1 to N (Carrying Forward Context) ──
            if status_callback:
                status_callback(f"Generating Part 2/2: Slides {half_idx + 1}–{total_slides_count}...", model="gemini-4-hydrogen")

            prompt_pass2 = f"""You are the student's personal medical study partner writing high-yield lecture notes for King's College London MBBS Year 1.
Your task is to continue generating comprehensive lecture notes for the second half of this lecture, ensuring seamless continuity with Part 1.

{style_guide_section}

Lecture Title: "{lecture_name}"
Total Slides in Deck: {total_slides_count}
GENERATION PHASE: PART 2 OF 2 (Slides {half_idx + 1} through {total_slides_count})

ESTABLISHED LECTURE LEARNING OBJECTIVES (from Part 1):
{ilo_block}

TOPICS ALREADY COVERED IN PART 1 (Slides 1 to {half_idx}):
{part1_outline}

Extracted Slide Content for Part 2 (Slides {half_idx + 1} through {total_slides_count}):
{slides_part2_str}

{shared_instructions}

PART 2 CONTINUITY RULES:
1. STRICTLY DO NOT output `## 🎯 Learning Objectives` or re-state the learning objectives (already established in Part 1).
2. Continue directly with the Notes body covering Slides {half_idx + 1} through {total_slides_count} in sequential order without skipping.
3. Maintain the exact same style, clinical depth, `<p />` dividers, `<u>Topic:</u>` headings, bracketed term clarifications `(...)`, and single-slide diagram markers `(slide N: ...)`.
4. Output raw markdown.
"""

            raw_part2 = engine._call_gemini_rest(
                prompt_pass2,
                timeout=180,
                status_callback=status_callback,
                preferred_model="gemini-4-hydrogen",
                task_label="Notes Part 2/2",
                restrict_to_flagship=True
            ).strip()

            if getattr(engine, "_notes_cancel_requested", False) or getattr(engine, "_cancel_requested", False):
                raise InterruptedError("Generation cancelled by user.")

            # Clean Part 2 body (remove any stray ILO or Notes headers)
            notes_part2 = re.sub(r"^##\s*(?:🎯\s*)?Learning Objectives.*?(?=##\s*(?:📝\s*)?Notes|\n\n\w|\Z)", "", raw_part2, flags=re.DOTALL | re.IGNORECASE).strip()
            notes_part2 = re.sub(r"^##\s*(?:📝\s*)?Notes\s*", "", notes_part2, flags=re.IGNORECASE).strip()

            # Merge both parts under a single unified Learning Objectives header and Notes body
            merged_notes = f"## 🎯 Learning Objectives\n{ilo_block}\n\n## 📝 Notes\n{notes_part1}\n\n<p />\n\n{notes_part2}".strip()

            if not is_custom:
                merged_notes = re.sub(r"^##\s*(?:🎯\s*)?Learning Objectives", "## 🎯 Learning Objectives", merged_notes, flags=re.MULTILINE | re.IGNORECASE)
                merged_notes = re.sub(r"^##\s*(?:📝\s*)?Notes", "## 📝 Notes", merged_notes, flags=re.MULTILINE | re.IGNORECASE)
            return merged_notes, getattr(engine, "last_model_used", "")

        # ── SINGLE-PASS GENERATION (FOR DECKS <= 12 SLIDES) ───────────────────
        slides_summary_str = self._build_slides_summary_str(slides_data)
        prompt = f"""You are the student's personal medical study partner writing high-yield lecture notes for King's College London MBBS Year 1.
Your task is to turn these lecture slides into notes that read as if the user wrote them: lean, in their own words, built for revision rather than completeness.

{style_guide_section}

Lecture Title: "{lecture_name}"
Total Slides: {total_slides_count}

FULL LECTURE SYLLABUS TO COVER:
{topics_summary}

Extracted Lecture Content:
{slides_summary_str}

{shared_instructions}

Output ONLY raw markdown text starting with ## 🎯 Learning Objectives. Do not wrap in markdown code blocks.
"""

        if status_callback:
            status_callback("Synthesizing comprehensive lecture notes...", model="gemini-4-hydrogen")

        raw_notes = engine._call_gemini_rest(
            prompt,
            timeout=180,
            status_callback=status_callback,
            preferred_model="gemini-4-hydrogen",
            task_label="Notes Synthesis",
            restrict_to_flagship=True
        ).strip()

        # Deterministic emoji heading enforcement (applied in default style)
        if not is_custom:
            raw_notes = re.sub(r"^##\s*(?:🎯\s*)?Learning Objectives", "## 🎯 Learning Objectives", raw_notes, flags=re.MULTILINE | re.IGNORECASE)
            raw_notes = re.sub(r"^##\s*(?:📝\s*)?Notes", "## 📝 Notes", raw_notes, flags=re.MULTILINE | re.IGNORECASE)
        return raw_notes, getattr(engine, "last_model_used", "")

class NotionSyncEngine:
    def __init__(self, config: ConfigManager, state: AppState, render_engine: Optional["SlideRenderEngine"] = None):
        self._config = config
        self._state = state
        self._render_engine = render_engine

    def _get_headers(self) -> dict:
        token = self._config.get("NOTION_API_KEY", "").strip()
        if not token:
            raise ValueError("Notion API Key is not set. Please add it in Settings.")
        return {
            "Authorization": f"Bearer {token}",
            "Notion-Version": "2022-06-28"
        }

    @staticmethod
    def _extract_page_title(page_data: dict) -> str:
        props = page_data.get("properties", {})
        for p in props.values():
            if isinstance(p, dict) and p.get("type") == "title":
                parts = p.get("title", [])
                if parts:
                    return "".join(part.get("plain_text", "") for part in parts).strip()
        if "title" in props and isinstance(props["title"], list):
            return "".join(part.get("plain_text", "") for part in props["title"]).strip()
        return ""

    def search_all_notion_pages(self, query: str = "", limit: int = 25) -> list[dict]:
        headers = self._get_headers()
        seen: dict[str, dict] = {}

        # 1. Database query if database ID configured
        db_id = self._config.get("NOTION_PAGE_ID", "").strip()
        if db_id:
            clean_db_id = db_id.replace("-", "")
            try:
                db_payload = {"page_size": min(limit * 2, 100)}
                db_res = requests.post(
                    f"https://api.notion.com/v1/databases/{clean_db_id}/query",
                    headers={**headers, "Content-Type": "application/json"},
                    json=db_payload,
                    timeout=15
                )
                if db_res.status_code == 200:
                    for r in db_res.json().get("results", []):
                        if r.get("object") == "page":
                            pid = r.get("id")
                            title = self._extract_page_title(r)
                            if pid and title:
                                seen[pid] = {"id": pid, "title": title, "url": r.get("url", "")}
            except Exception:
                pass

        # 2. Workspace search
        try:
            search_payload: dict = {
                "filter": {"value": "page", "property": "object"},
                "page_size": min(limit * 2, 100)
            }
            if query and query.strip():
                search_payload["query"] = query.strip()
            res = requests.post(
                "https://api.notion.com/v1/search",
                headers={**headers, "Content-Type": "application/json"},
                json=search_payload,
                timeout=15
            )
            if res.status_code == 200:
                for r in res.json().get("results", []):
                    if r.get("object") == "page":
                        pid = r.get("id")
                        title = self._extract_page_title(r)
                        if pid and title:
                            seen[pid] = {"id": pid, "title": title, "url": r.get("url", "")}
        except Exception:
            pass

        results = list(seen.values())
        if query and query.strip():
            q_clean = query.strip().lower()
            def _rank(item):
                t_lower = item["title"].lower()
                if t_lower == q_clean:
                    return 0
                if t_lower.startswith(q_clean):
                    return 1
                if q_clean in t_lower:
                    return 2
                q_words = [w for w in q_clean.split() if len(w) > 1]
                matches = sum(1 for w in q_words if w in t_lower)
                return 10 - matches

            filtered = [
                it for it in results
                if q_clean in it["title"].lower() or any(w in it["title"].lower() for w in q_clean.split() if len(w) > 2)
            ]
            filtered.sort(key=_rank)
            return filtered[:limit]

        return results[:limit]

    def search_session_page(self, query: str, target_page_id: str = "") -> Optional[dict]:
        headers = self._get_headers()

        # If an explicit page ID is supplied, fetch it directly
        if target_page_id:
            target_clean = target_page_id.strip()
            try:
                res = requests.get(f"https://api.notion.com/v1/pages/{target_clean}", headers=headers, timeout=15)
                if res.status_code == 200:
                    data = res.json()
                    title = self._extract_page_title(data)
                    blk_res = requests.get(f"https://api.notion.com/v1/blocks/{target_clean}/children?page_size=5", headers=headers, timeout=10)
                    has_content = len(blk_res.json().get("results", [])) > 0 if blk_res.status_code == 200 else False
                    return {"id": target_clean, "title": title or "Notion Page", "has_content": has_content, "score": 1.0, "candidates": []}
            except Exception:
                pass
            return {"id": target_clean, "title": "Selected Notion Page", "has_content": True, "score": 1.0, "candidates": []}

        # 1. Clean filename & strip academic metadata, year prefixes, answer tags, and suffixes
        clean_q = re.sub(r'(\.pptx|\.pdf)$', '', query, flags=re.IGNORECASE).strip()

        # Remove academic years e.g. 2024, 2025, 2026, 26_27, 26-27, 2026-27
        clean_no_year = re.sub(r'\b(20\d\d|19\d\d|\d{2}[-_]\d{2}|\d{4}[-_]\d{2,4})\b', '', clean_q).strip(' _-')

        # Remove common student tags, answer tags, platform tags, tutorial tags
        clean_topic = re.sub(
            r'([-_ ]*(student[s]?([-_ ]*slides?)?|slides?|handout|annotated|printable|notes|final|v\d+'
            r'|with(out)?[-_ ]*answers?|no[-_ ]*answers?|keats|tutorial|workshop|cbd|case[-_ ]*based[-_ ]*discussion)\b|\([^\)]*\))',
            '', clean_no_year, flags=re.IGNORECASE
        ).strip(' _-')

        # Strip lecture prefixes e.g. "Lecture 05", "05.", "MCB-10", "ACRS03L"
        topic_no_num = re.sub(r'^(lecture[-_ ]*\d+[-_ ]*|\d+[-_ .]+|[A-Z]{2,5}[-_ ]*\d{1,3}[A-Z]?[-_ .]*)', '', clean_topic, flags=re.IGNORECASE).strip(' _-')
        if not topic_no_num:
            topic_no_num = clean_topic

        # Detect block from query, pptx_path, or target_deck
        detected_block = ""
        known_blocks = ['MCB', 'NBSS', 'NAM', 'DH', 'IAI', 'EAB', 'GEN', 'FAH', 'FPP', 'ACRS', 'PCRS', 'AGRS', 'PGRS', 'AMS', 'CRANIAL']
        path_str = str(getattr(self._state, "pptx_path", "") or "")
        deck_str = str(getattr(self._state, "target_deck", "") or "")
        full_context = f"{clean_q} {path_str} {deck_str}".upper()
        for b in known_blocks:
            if re.search(rf"(?<![A-Z]){b}(?![A-Z])", full_context):
                detected_block = b
                break

        # Detect specific code e.g. NBSS04, FAH01
        code_match = re.search(r'\b([A-Z]{3,5}[-_ ]*\d{1,3}[A-Z]?)\b', clean_q, re.IGNORECASE)
        code_str = code_match.group(1).upper() if code_match else ""

        # Extract title from presentation slide 0 if available
        slide1_title = ""
        prs = getattr(self._state, "presentation", None)
        if prs and len(prs.slides) > 0:
            s0 = prs.slides[0]
            try:
                if s0.shapes.title and s0.shapes.title.text:
                    slide1_title = s0.shapes.title.text.strip().replace("\x0b", " ").split("\n")[0].strip()
            except Exception:
                pass
            if not slide1_title:
                for sh in s0.shapes:
                    if sh.has_text_frame and sh.text_frame.text.strip():
                        slide1_title = sh.text_frame.text.strip().split("\n")[0].strip()
                        break

        # Build prioritized search terms to query Notion API
        search_terms = []
        if clean_topic:
            search_terms.append(clean_topic)
        if topic_no_num and topic_no_num not in search_terms:
            search_terms.append(topic_no_num)
        if slide1_title and slide1_title not in search_terms:
            s1_clean = re.sub(r'(\.pptx|\.pdf)$', '', slide1_title, flags=re.IGNORECASE).strip()
            if s1_clean:
                search_terms.append(s1_clean)
        if detected_block and clean_topic:
            search_terms.append(f"{detected_block} {clean_topic}")
        if code_str and code_str not in search_terms:
            search_terms.append(code_str)
        if clean_q not in search_terms:
            search_terms.append(clean_q)

        # Helper to normalize strings for comparison
        def _norm(s: str) -> str:
            return re.sub(r'[^a-z0-9 ]+', ' ', s.lower()).strip()

        stop_words = {
            'the', 'and', 'to', 'of', 'in', 'for', 'a', 'an', 'on', 'with', 'without',
            'answers', 'answer', 'student', 'slides', 'lecture', 'presentation', 'session',
            'part', 'keats', 'tutorial', 'workshop', 'cbd', 'case', 'discussion', 'notes',
            'intro', 'introduction', 'final', 'v1', 'v2'
        }
        core_query_words = [w for w in _norm(clean_topic).split() if w not in stop_words and len(w) > 1]
        if not core_query_words and topic_no_num:
            core_query_words = [w for w in _norm(topic_no_num).split() if w not in stop_words and len(w) > 1]
        if not core_query_words and slide1_title:
            core_query_words = [w for w in _norm(slide1_title).split() if w not in stop_words and len(w) > 1]

        GENERIC_MEDICAL_CATEGORIES = {
            'metabolism', 'physiology', 'anatomy', 'histology', 'pathology', 'pharmacology',
            'system', 'systems', 'introduction', 'intro', 'overview', 'principles', 'foundations',
            'basis', 'structure', 'function', 'tutorial', 'session', 'lecture', 'revision', 'summary'
        }

        # 2. Query Notion API
        seen_pages: dict[str, dict] = {}

        # 2a. If NOTION_PAGE_ID is configured and points to a database, query database directly with pagination
        db_id = self._config.get("NOTION_PAGE_ID", "").strip()
        if db_id:
            clean_db_id = db_id.replace("-", "")
            has_more = True
            start_cursor = None
            page_count = 0
            while has_more and page_count < 10:  # Fetch up to 1000 database items
                try:
                    payload = {"page_size": 100}
                    if start_cursor:
                        payload["start_cursor"] = start_cursor
                    db_res = requests.post(
                        f"https://api.notion.com/v1/databases/{clean_db_id}/query",
                        headers={**headers, "Content-Type": "application/json"},
                        json=payload,
                        timeout=15
                    )
                    if db_res.status_code == 200:
                        db_data = db_res.json()
                        for r in db_data.get("results", []):
                            if r.get("object") == "page":
                                seen_pages[r.get("id")] = r
                        has_more = db_data.get("has_more", False)
                        start_cursor = db_data.get("next_cursor")
                        page_count += 1
                    else:
                        break
                except Exception:
                    break

        # 2b. Run Notion global workspace search for candidate queries
        for term in search_terms[:4]:
            try:
                res = requests.post(
                    "https://api.notion.com/v1/search",
                    headers={**headers, "Content-Type": "application/json"},
                    json={"query": term, "filter": {"value": "page", "property": "object"}, "page_size": 25},
                    timeout=15
                )
                if res.status_code == 200:
                    for r in res.json().get("results", []):
                        if r.get("object") == "page":
                            seen_pages[r.get("id")] = r
            except Exception:
                continue

        if not seen_pages:
            return None

        # 3. Score every candidate page
        code_clean = re.sub(r'[^A-Z0-9]', '', code_str) if code_str else ''
        q_norm = _norm(clean_q)
        clean_topic_norm = _norm(clean_topic)
        topic_no_num_norm = _norm(topic_no_num)

        specific_query_words = [w for w in core_query_words if w not in GENERIC_MEDICAL_CATEGORIES and len(w) >= 3]

        scored_candidates = []
        for page_id, r in seen_pages.items():
            title = self._extract_page_title(r)
            if not title:
                continue

            title_norm = _norm(title)
            title_words = set(title_norm.split())

            # Detect candidate's own module/block code (e.g. ACRS07L, ACRS03L)
            cand_code_match = re.search(r'\b([A-Z]{3,5}[-_ ]*\d{1,3}[A-Z]?)\b', title, re.IGNORECASE)
            cand_clean = re.sub(r'[^A-Z0-9]', '', cand_code_match.group(1).upper()) if cand_code_match else ''

            # Keyword overlap
            matched_words = 0.0
            for cw in core_query_words:
                if cw in title_words:
                    matched_words += 1.0
                elif any((cw in tw or tw in cw) for tw in title_words if len(tw) >= 4 and len(cw) >= 4):
                    matched_words += 0.85

            overlap_ratio = (matched_words / len(core_query_words)) if core_query_words else 0.0
            score = overlap_ratio * 0.60

            # Substring match (bidirectional)
            if clean_topic_norm and (clean_topic_norm in title_norm or (len(title_norm) >= 5 and title_norm in clean_topic_norm)):
                score += 0.50
            elif topic_no_num_norm and (topic_no_num_norm in title_norm or (len(title_norm) >= 5 and title_norm in topic_no_num_norm)):
                score += 0.40
            elif slide1_title and (_norm(slide1_title) in title_norm or (len(title_norm) >= 5 and title_norm in _norm(slide1_title))):
                score += 0.35

            # Block match
            if detected_block and detected_block.lower() in title_norm:
                score += 0.25

            # Code match vs conflict
            if code_clean and cand_clean:
                if code_clean == cand_clean:
                    score += 0.60
                else:
                    score -= 0.80  # Direct code conflict (e.g. ACRS07L vs ACRS02L / ACRS03L)
            elif code_clean and code_clean.lower() in title_norm:
                score += 0.40

            # Distinguishing medical topic check:
            # If the lecture has specific distinguishing terms (e.g. 'anaerobic') but candidate matches NONE of them,
            # heavy penalty to avoid matching unrelated lectures that only share generic words like 'metabolism'.
            if specific_query_words:
                has_specific_match = any((sqw in title_words or any(sqw in tw for tw in title_words if len(tw) >= 4)) for sqw in specific_query_words)
                if not has_specific_match:
                    score -= 0.70

            # Penalize prominent contradictory specific topic keywords present in candidate but absent in query
            specific_title_words = [w for w in title_words if w not in GENERIC_MEDICAL_CATEGORIES and len(w) >= 3]
            contradictory_words = [tw for tw in specific_title_words if tw not in core_query_words and not any(tw in qw for qw in core_query_words)]
            if specific_query_words and not any(sqw in title_words for sqw in specific_query_words) and contradictory_words:
                score -= 0.35 * min(len(contradictory_words), 2)

            # Organ system mutual exclusivity (prevent Respiratory matching Cardiovascular)
            if ('respiratory' in q_norm or 'respiration' in q_norm) and 'cardiovascular' in title_norm and 'respiratory' not in title_norm:
                score -= 0.70
            if 'cardiovascular' in q_norm and 'respiratory' in title_norm and 'cardiovascular' not in title_norm:
                score -= 0.70
            if 'renal' in q_norm and 'gastrointestinal' in title_norm and 'renal' not in title_norm:
                score -= 0.70
            if 'gastrointestinal' in q_norm and 'renal' in title_norm and 'gastrointestinal' not in title_norm:
                score -= 0.70

            # Number match / mismatch
            q_nums = set(re.findall(r'\b\d+\b', clean_q))
            t_nums = set(re.findall(r'\b\d+\b', title))
            if q_nums and t_nums:
                if q_nums.intersection(t_nums):
                    score += 0.20
                else:
                    score -= 0.25

            if score >= 0.45:
                scored_candidates.append({
                    "id": page_id,
                    "title": title,
                    "score": score
                })

        if not scored_candidates:
            return None

        scored_candidates.sort(key=lambda x: x["score"], reverse=True)
        best = scored_candidates[0]

        has_content = False
        try:
            blk_res = requests.get(
                f"https://api.notion.com/v1/blocks/{best['id']}/children?page_size=5",
                headers=headers,
                timeout=10
            )
            if blk_res.status_code == 200:
                has_content = len(blk_res.json().get("results", [])) > 0
        except Exception:
            pass

        return {
            "id": best["id"],
            "title": best["title"],
            "has_content": has_content,
            "score": best["score"],
            "candidates": [
                {"id": c["id"], "title": c["title"], "score": round(c["score"], 2)}
                for c in scored_candidates[:10]
            ]
        }

    def upload_image(self, img_bytes: bytes, filename: str) -> str:
        headers = self._get_headers()
        init_res = requests.post(
            "https://api.notion.com/v1/file_uploads",
            headers={**headers, "Content-Type": "application/json"},
            json={"mode": "single_part"},
            timeout=20
        )
        if init_res.status_code != 200:
            raise ValueError(f"Notion file upload init failed: {init_res.text}")
        upload_id = init_res.json()["id"]

        files = {"file": (filename, img_bytes, "image/png")}
        send_res = requests.post(
            f"https://api.notion.com/v1/file_uploads/{upload_id}/send",
            headers=headers,
            files=files,
            timeout=40
        )
        if send_res.status_code != 200:
            raise ValueError(f"Notion file upload binary send failed: {send_res.text}")
        return upload_id

    def _get_diagram_crop_box(self, slide, slide_w: int, slide_h: int) -> Optional[tuple[float, float, float, float]]:
        """Intelligently crops slide to visual diagram area:
        Trims top slide title banner and bottom citation footer while preserving FULL HORIZONTAL SPAN (0.0 to 1.0).
        This guarantees multi-panel comparisons (A, B, C), side-by-side diagrams, CT scans, and right-aligned anatomical drawings are NEVER sliced.
        """
        title_shape = getattr(slide.shapes, "title", None)
        title_bot_norm = 0.10  # default top crop below slide title banner

        if title_shape and hasattr(title_shape, "top") and hasattr(title_shape, "height"):
            if title_shape.top < 0.20 * slide_h:
                title_bot_norm = min(0.20, max(0.08, (title_shape.top + title_shape.height) / slide_h + 0.008))
        else:
            # Check for top header / title text box
            for s in slide.shapes:
                if getattr(s, "top", 0) < 0.16 * slide_h and getattr(s, "has_text_frame", False):
                    words = len(s.text_frame.text.strip().split())
                    if 0 < words <= 12:
                        title_bot_norm = min(0.20, max(title_bot_norm, (s.top + s.height) / slide_h + 0.008))

        footer_top_norm = 1.0  # default bottom crop
        for s in slide.shapes:
            if getattr(s, "top", 0) >= 0.86 * slide_h and getattr(s, "has_text_frame", False):
                t_lower = s.text_frame.text.lower()
                if any(k in t_lower for k in [
                    "et al", "atlas", "schuenke", "moore", "standring", "gray", "netter",
                    "copyright", "source", "thieme", "elsevier", "springer", "mcgraw", "mcminn"
                ]):
                    footer_top_norm = min(footer_top_norm, max(0.85, s.top / slide_h - 0.005))

        # Check if any picture or drawing extends high up near the top or low near the bottom
        for s in slide.shapes:
            st = getattr(s, "shape_type", None)
            if st in (13, 6):  # Picture or Group
                if s.top < title_bot_norm * slide_h and (s.left + s.width) > 0.3 * slide_w:
                    # Diagram extends into upper area; reduce top crop to avoid clipping diagram top
                    title_bot_norm = max(0.0, min(title_bot_norm, s.top / slide_h - 0.005))
                if (s.top + s.height) > footer_top_norm * slide_h:
                    # Diagram extends into footer area; push bottom crop down to avoid clipping diagram bottom
                    footer_top_norm = min(1.0, max(footer_top_norm, (s.top + s.height) / slide_h + 0.005))

        return (0.0, max(0.0, title_bot_norm), 1.0, min(1.0, footer_top_norm))

    def get_best_slide_image(self, prs: Optional[Presentation], slide_num: int, pptx_path: Optional[Path] = None) -> Optional[bytes]:
        if slide_num <= 1:
            # Slide 1 is strictly the title/cover slide with university logos and metadata; never an examinable medical diagram
            return None
        slide_idx = slide_num - 1

        final_pil = None

        # 1. Primary: High-fidelity rendered slide image from SlideRenderEngine with intelligent diagram cropping
        # Strictly verify that SlideRenderEngine is currently holding THIS exact presentation (prevents cross-lecture diagram leakage)
        matching_render = False
        if self._render_engine:
            if prs and getattr(self._render_engine, "_prs", None) is prs:
                matching_render = True
            elif pptx_path and getattr(self._render_engine, "_current_path", None) == pptx_path:
                matching_render = True

        if matching_render and prs and slide_idx < len(prs.slides):
            rendered = self._render_engine.get_slide_image(slide_idx, scale=1.5)
            if rendered:
                final_pil = rendered.convert("RGB")
                slide = prs.slides[slide_idx]
                crop_norm = self._get_diagram_crop_box(slide, prs.slide_width, prs.slide_height)
                if crop_norm:
                    W, H = final_pil.size
                    l, t, r, b = crop_norm
                    box = (int(l * W), int(t * H), int(r * W), int(b * H))
                    if box[2] > box[0] + 50 and box[3] > box[1] + 50:
                        final_pil = final_pil.crop(box)
                else:
                    # No specific diagram box identified; remove the top title placeholder
                    W, H = final_pil.size
                    final_pil = final_pil.crop((0, int(0.12 * H), W, H))

        # If not matching the currently active render engine presentation, check if a cached PDF exists for pptx_path
        if final_pil is None and pptx_path and prs and slide_idx < len(prs.slides):
            try:
                mtime = pptx_path.stat().st_mtime if pptx_path.exists() else 0
                key = hashlib.md5(f"{pptx_path.resolve()}_{mtime}".encode()).hexdigest()
                cached_pdf = APP_DATA_DIR / "slide_cache" / f"{key}.pdf"
                if cached_pdf.exists() and cached_pdf.stat().st_size > 0:
                    with _C_EXTENSION_LOCK:
                        if HAS_PDFIUM:
                            doc = pdfium.PdfDocument(str(cached_pdf))
                            if slide_idx < len(doc):
                                final_pil = doc[slide_idx].render(scale=1.5).to_pil().convert("RGB")
                        elif HAS_FITZ:
                            doc = fitz.open(str(cached_pdf))
                            if slide_idx < len(doc):
                                page = doc[slide_idx]
                                mat = fitz.Matrix(1.5, 1.5)
                                pix = page.get_pixmap(matrix=mat, alpha=False)
                                final_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    if final_pil:
                        slide = prs.slides[slide_idx]
                        crop_norm = self._get_diagram_crop_box(slide, prs.slide_width, prs.slide_height)
                        if crop_norm:
                            W, H = final_pil.size
                            l, t, r, b = crop_norm
                            box = (int(l * W), int(t * H), int(r * W), int(b * H))
                            if box[2] > box[0] + 50 and box[3] > box[1] + 50:
                                final_pil = final_pil.crop(box)
            except Exception as e:
                print(f"[get_best_slide_image] Error rendering from cached PDF: {e}")

        # 2. Fallback: Extract raw pictures directly from prs slide
        if final_pil is None and prs and slide_idx < len(prs.slides):
            slide = prs.slides[slide_idx]
            extracted_imgs, _ = SlideParsEngine.extract_slide_data(slide)
            if extracted_imgs:
                extracted_imgs.sort(key=lambda img: img.size[0] * img.size[1], reverse=True)
                final_pil = extracted_imgs[0].convert("RGB")

        if final_pil is None:
            return None

        # Constrain dimensions to max width 960px, max height 540px to prevent oversized vertical images
        max_w, max_h = 960, 540
        w, h = final_pil.size
        scale = min(max_w / w, max_h / h, 1.0)
        if scale < 1.0:
            new_size = (int(w * scale), int(h * scale))
            final_pil = final_pil.resize(new_size, Image.Resampling.LANCZOS)

        out_buf = io.BytesIO()
        final_pil.save(out_buf, format="PNG", optimize=True)
        return out_buf.getvalue()

    @staticmethod
    def _parse_inline_rich_text(text: str) -> list[dict]:
        if not text:
            return []
        # Unescape backslashes before formatting chars e.g. \* -> *
        text = re.sub(r'\\([*_`~])', r'\1', text)
        
        # Normalize <span underline="true">...</span> -> <u>...</u>
        text = re.sub(r'<span\s+underline=["\']true["\']>(.*?)</span>', r'<u>\1</u>', text, flags=re.IGNORECASE)

        pattern = re.compile(
            r'(?P<bold_underline>(?:<b><u>|<u><b>)(.*?)(?:</u></b>|</b></u>))|'
            r'(?P<bold_tag><b>(.*?)</b>)|'
            r'(?P<under_tag><u>(.*?)</u>)|'
            r'(?P<italic_tag><i>(.*?)</i>)|'
            r'(?P<bold_md>\*\*(.*?)\*\*)|'
            r'(?P<code_md>`(.*?)`)|'
            r'(?P<italic_md>(?<!\*)\*(?!\*)(.*?)(?<!\*)\*(?!\*))'
        )

        results = []
        pos = 0

        def add_chunk(content: str, bold=False, italic=False, underline=False, code=False):
            if not content:
                return
            clean_content = re.sub(r'</?(?:b|u|i|span|strong|em)[^>]*>', '', content)
            if not clean_content:
                return
            # Safeguard against Notion's 2000-char per rich_text item limit
            for chunk_part in [clean_content[k:k+2000] for k in range(0, len(clean_content), 2000)]:
                results.append({
                    'type': 'text',
                    'text': {'content': chunk_part},
                    'annotations': {
                        'bold': bold,
                        'italic': italic,
                        'strikethrough': False,
                        'underline': underline,
                        'code': code,
                        'color': 'default'
                    }
                })

        for match in pattern.finditer(text):
            start, end = match.span()
            if start > pos:
                add_chunk(text[pos:start])

            if match.group('bold_underline'):
                content = match.group(2)
                add_chunk(content, bold=True, underline=True)
            elif match.group('bold_tag'):
                content = match.group(4)
                u_m = re.search(r'<u>(.*?)</u>', content)
                if u_m:
                    add_chunk(u_m.group(1), bold=True, underline=True)
                else:
                    add_chunk(content, bold=True)
            elif match.group('under_tag'):
                content = match.group(6)
                b_m = re.search(r'<b>(.*?)</b>|\*\*(.*?)\*\*', content)
                if b_m:
                    inner = b_m.group(1) or b_m.group(2)
                    add_chunk(inner, bold=True, underline=True)
                else:
                    add_chunk(content, underline=True)
            elif match.group('italic_tag'):
                content = match.group(8)
                add_chunk(content, italic=True)
            elif match.group('bold_md'):
                content = match.group(10)
                u_m = re.search(r'<u>(.*?)</u>', content)
                if u_m:
                    add_chunk(u_m.group(1), bold=True, underline=True)
                else:
                    add_chunk(content, bold=True)
            elif match.group('code_md'):
                content = match.group(12)
                add_chunk(content, code=True)
            elif match.group('italic_md'):
                content = match.group(14)
                add_chunk(content, italic=True)

            pos = end

        if pos < len(text):
            add_chunk(text[pos:])

        return results

    @classmethod
    def _is_table_separator(cls, line: str) -> bool:
        s = line.strip()
        if not s or "-" not in s:
            return False
        inner = s.strip("|")
        cols = [c.strip() for c in inner.split("|")]
        if not cols:
            return False
        return all(bool(re.match(r"^:?-+:?$", c)) for c in cols)

    @classmethod
    def _split_table_row(cls, row_str: str) -> list[str]:
        s = row_str.strip()
        if s.startswith("|"):
            s = s[1:]
        if s.endswith("|"):
            s = s[:-1]
        parts = re.split(r"(?<!\\)\|", s)
        return [p.strip().replace(r"\|", "|") for p in parts]

    def _parse_markdown_table_at(self, lines: list[str], start_idx: int) -> tuple[Optional[dict], int]:
        if start_idx + 1 >= len(lines):
            return None, start_idx
        header_line = lines[start_idx].strip()
        sep_line = lines[start_idx + 1].strip()

        if "|" not in header_line or not self._is_table_separator(sep_line):
            return None, start_idx

        header_cols = self._split_table_row(header_line)
        if not header_cols:
            return None, start_idx

        table_width = len(header_cols)
        row_blocks = []

        row_blocks.append({
            "object": "block",
            "type": "table_row",
            "table_row": {
                "cells": [
                    self._parse_inline_rich_text(col) if col else []
                    for col in header_cols
                ]
            }
        })

        curr_idx = start_idx + 2
        while curr_idx < len(lines):
            curr_line = lines[curr_idx].strip()
            if not curr_line or ("|" not in curr_line):
                break
            if curr_line.startswith(("#", "- ", "<p", "(")):
                break
            if self._is_table_separator(curr_line):
                curr_idx += 1
                continue

            row_cols = self._split_table_row(curr_line)
            if len(row_cols) < table_width:
                row_cols.extend([""] * (table_width - len(row_cols)))
            elif len(row_cols) > table_width:
                row_cols = row_cols[:table_width]

            row_blocks.append({
                "object": "block",
                "type": "table_row",
                "table_row": {
                    "cells": [
                        self._parse_inline_rich_text(c) if c else []
                        for c in row_cols
                    ]
                }
            })
            curr_idx += 1
            if len(row_blocks) >= 99:
                break

        table_block = {
            "object": "block",
            "type": "table",
            "table": {
                "table_width": table_width,
                "has_column_header": True,
                "has_row_header": False,
                "children": row_blocks
            }
        }
        return table_block, curr_idx

    def _parse_html_table_at(self, lines: list[str], start_idx: int) -> tuple[Optional[dict], int]:
        curr_line = lines[start_idx].strip()
        if not curr_line.lower().startswith("<table"):
            return None, start_idx
        table_html_parts = []
        curr_idx = start_idx
        while curr_idx < len(lines):
            l = lines[curr_idx]
            table_html_parts.append(l)
            curr_idx += 1
            if "</table>" in l.lower():
                break
        html_str = "\n".join(table_html_parts)
        row_matches = re.findall(r"<tr[^>]*>(.*?)</tr>", html_str, re.IGNORECASE | re.DOTALL)
        if not row_matches:
            return None, start_idx
        parsed_rows = []
        max_w = 0
        for r_html in row_matches:
            cells = re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", r_html, re.IGNORECASE | re.DOTALL)
            clean_cells = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
            if clean_cells:
                parsed_rows.append(clean_cells)
                if len(clean_cells) > max_w:
                    max_w = len(clean_cells)
        if max_w == 0 or not parsed_rows:
            return None, start_idx

        row_blocks = []
        for r in parsed_rows[:99]:
            if len(r) < max_w:
                r.extend([""] * (max_w - len(r)))
            elif len(r) > max_w:
                r = r[:max_w]
            row_blocks.append({
                "object": "block",
                "type": "table_row",
                "table_row": {
                    "cells": [
                        self._parse_inline_rich_text(c) if c else []
                        for c in r
                    ]
                }
            })

        table_block = {
            "object": "block",
            "type": "table",
            "table": {
                "table_width": max_w,
                "has_column_header": True,
                "has_row_header": False,
                "children": row_blocks
            }
        }
        return table_block, curr_idx

    def markdown_to_notion_blocks(self, md_text: str, prs: Optional[Presentation] = None) -> list[dict]:
        lines = md_text.split("\n")
        blocks = []
        seen_image_hashes: set[str] = set()
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line:
                i += 1
                continue

            # Check for Markdown table
            tbl_block, next_i = self._parse_markdown_table_at(lines, i)
            if tbl_block is not None:
                blocks.append(tbl_block)
                i = next_i
                continue

            # Check for HTML table
            html_tbl, next_i = self._parse_html_table_at(lines, i)
            if html_tbl is not None:
                blocks.append(html_tbl)
                i = next_i
                continue

            if line in ("<p />", "<p/>", "<p>", "<empty-block/>", "<empty-block />", "<br/>", "<br />"):
                blocks.append({"object": "block", "type": "paragraph", "paragraph": {"rich_text": []}})
                i += 1
                continue

            if line.startswith("## "):
                h_text = line[3:].strip()
                blocks.append({"object": "block", "type": "heading_2", "heading_2": {"rich_text": self._parse_inline_rich_text(h_text)}})
                i += 1
                continue
            elif line.startswith("# "):
                h_text = line[2:].strip()
                blocks.append({"object": "block", "type": "heading_1", "heading_1": {"rich_text": self._parse_inline_rich_text(h_text)}})
                i += 1
                continue
            elif line.startswith("### "):
                h_text = line[4:].strip()
                blocks.append({"object": "block", "type": "heading_3", "heading_3": {"rich_text": self._parse_inline_rich_text(h_text)}})
                i += 1
                continue

            if line.startswith("- "):
                bullet_text = line[2:].strip()
                blocks.append({"object": "block", "type": "bulleted_list_item", "bulleted_list_item": {"rich_text": self._parse_inline_rich_text(bullet_text)}})
                i += 1
                continue

            slide_match = re.match(r"^\(\s*slides?\s+([0-9\s,\-–and]+?):\s*(.*?)\s*\)$", line, re.IGNORECASE)
            if slide_match:
                slide_range_str = slide_match.group(1).strip()
                caption_title = slide_match.group(2).strip()
                extra_caption = ""
                if i + 1 < len(lines) and lines[i+1].strip() and not lines[i+1].startswith("- ") and not lines[i+1].startswith("<u>") and not lines[i+1].startswith("<b>") and not lines[i+1].startswith("<span") and not lines[i+1].startswith("##") and not lines[i+1].startswith("<p") and not lines[i+1].startswith("<empty") and not lines[i+1].startswith("("):
                    extra_caption = " " + lines[i+1].strip()
                    i += 1

                # Extract all slide numbers from range or enumeration
                slide_nums = []
                range_match = re.search(r'(\d+)\s*[-–]\s*(\d+)', slide_range_str)
                if range_match:
                    s_start, s_end = int(range_match.group(1)), int(range_match.group(2))
                    if s_start <= s_end and s_end - s_start <= 10:
                        slide_nums = list(range(s_start, s_end + 1))
                if not slide_nums:
                    nums = [int(n) for n in re.findall(r'\b\d+\b', slide_range_str)]
                    slide_nums = sorted(list(set(nums)))

                pushed_any_img = False
                for s_num in slide_nums:
                    img_blob = self.get_best_slide_image(prs, s_num)
                    if img_blob:
                        img_hash = hashlib.md5(img_blob).hexdigest()
                        # Only skip if the exact rendered diagram was already uploaded in this note
                        if img_hash in seen_image_hashes:
                            continue
                        seen_image_hashes.add(img_hash)
                        try:
                            upload_id = self.upload_image(img_blob, f"slide_{s_num}.png")
                            slide_cap = f"(slide {s_num}: {caption_title}){extra_caption}".strip()
                            blocks.append({
                                "object": "block",
                                "type": "image",
                                "image": {
                                    "type": "file_upload",
                                    "file_upload": {"id": upload_id},
                                    "caption": [{"type": "text", "text": {"content": slide_cap}}]
                                }
                            })
                            pushed_any_img = True
                        except Exception as exc:
                            print(f"[markdown_to_notion_blocks] Failed to upload diagram for slide {s_num}: {exc}")
                if not pushed_any_img:
                    full_caption = f"({line.strip('()')}){extra_caption}".strip()
                    blocks.append({"object": "block", "type": "paragraph", "paragraph": {"rich_text": self._parse_inline_rich_text(full_caption)}})
                i += 1
                continue

            # Major Title: <b><u>Title</u></b> or <span underline="true">**Title**</span>
            m_major = re.match(r"^(?:<b><u>|<u><b>|<span\s+underline=[\"']true[\"']>\*\*)(.*?)(?:</u></b>|</b></u>|\*\*</span>):?$", line)
            if m_major:
                content = m_major.group(1).strip()
                blocks.append({
                    "object": "block",
                    "type": "paragraph",
                    "paragraph": {
                        "rich_text": [{
                            "type": "text",
                            "text": {"content": content},
                            "annotations": {
                                "bold": True,
                                "underline": True,
                                "strikethrough": False,
                                "italic": False,
                                "code": False,
                                "color": "default"
                            }
                        }]
                    }
                })
                i += 1
                continue

            # Subtopic: <u>Title:</u> or <span underline="true">Title:</span>
            m_topic = re.match(r"^(?:<u>|<span\s+underline=[\"']true[\"']>)(.*?)(?:</u>|</span>):?$", line)
            if m_topic:
                content = m_topic.group(1).strip()
                if not content.endswith(":"):
                    content = content + ":"
                blocks.append({
                    "object": "block",
                    "type": "paragraph",
                    "paragraph": {
                        "rich_text": [{
                            "type": "text",
                            "text": {"content": content},
                            "annotations": {
                                "bold": False,
                                "underline": True,
                                "strikethrough": False,
                                "italic": False,
                                "code": False,
                                "color": "default"
                            }
                        }]
                    }
                })
                i += 1
                continue

            # Standard paragraph with inline rich text formatting
            blocks.append({
                "object": "block",
                "type": "paragraph",
                "paragraph": {
                    "rich_text": self._parse_inline_rich_text(line)
                }
            })
            i += 1
        return blocks

    def push_notes(self, lecture_name: str, md_notes: str, prs: Optional[Presentation] = None, target_page_id: str = "") -> dict:
        headers = self._get_headers()
        page_info = self.search_session_page(lecture_name, target_page_id=target_page_id)
        if not page_info:
            return {
                "success": False,
                "error": (
                    f"Could not find matching Notion page for '{lecture_name}'. "
                    f"Please ensure your 'Sessions' database in Notion is shared with your integration connection (click ••• in Notion > Connections > Add connections)."
                )
            }

        page_id = page_info["id"]
        blocks = self.markdown_to_notion_blocks(md_notes, prs)
        if not blocks:
            return {"success": False, "error": "No content blocks to push."}

        chunk_size = 50
        for c_idx in range(0, len(blocks), chunk_size):
            chunk = blocks[c_idx:c_idx+chunk_size]
            res = requests.patch(
                f"https://api.notion.com/v1/blocks/{page_id}/children",
                headers={**headers, "Content-Type": "application/json"},
                json={"children": chunk},
                timeout=30
            )
            if res.status_code != 200:
                return {"success": False, "error": f"Failed appending blocks to Notion: {res.status_code} - {res.text}"}

        return {
            "success": True,
            "page_title": page_info["title"],
            "page_id": page_id,
            "blocks_count": len(blocks),
            "had_content": page_info["has_content"]
        }

# ─────────────────────────────────────────────────────────────────────────────
# 8B. KCL ONEDRIVE ENGINE & AI CHATBOT ENGINE
# ─────────────────────────────────────────────────────────────────────────────
KCL_MODULES_INFO = {
    "FBS": "Foundations of Biomedical Sciences",
    "PAS": "Physiology and Anatomy of Systems",
    "IVBCP": "Introduction to Values-Based Clinical Practice",
    "FSTCP": "Foundations of Clinical Practice",
    "FSTCP 1": "Foundations of Clinical Practice 1"
}

KCL_BLOCKS_INFO = {
    "MCB": "Molecular & Cellular Biology",
    "NBSS": "Neural Basis of Systemic Structure",
    "NAM": "Nerve & Muscle",
    "DH": "Development & Homeostasis",
    "IAI": "Infection, Adaptation & Immunity",
    "EAB": "Elimination, Absorption & Balance",
    "GEN": "Genetics",
    "FAH": "Fundamental Anatomy & Histology",
    "FPP": "Foundations of Pharmacological Principles",
    "ACRS": "Anatomy of Cardiovascular & Respiratory Systems",
    "PCRS": "Physiology of Cardiovascular & Respiratory Systems",
    "AGRS": "Anatomy of Gastrointestinal & Renal Systems",
    "PGRS": "Physiology of Gastrointestinal & Renal Systems",
    "AMS": "Anatomy of Musculoskeletal System",
    "CRANIAL": "Cranial Nerves",
    "Clinical Skills": "Clinical Skills",
    "Human Values": "Human Values",
}

MEDICAL_ACRONYMS = {
    "ans": "autonomic nervous system",
    "pns": "peripheral nervous system",
    "cns": "central nervous system",
    "nmj": "neuromuscular junction",
    "ach": "acetylcholine",
    "ache": "acetylcholinesterase",
    "na": "noradrenaline",
    "ne": "norepinephrine",
    "adr": "adrenaline",
    "epi": "epinephrine",
    "ap": "action potential",
    "ecg": "electrocardiogram",
    "ekg": "electrocardiogram",
    "bp": "blood pressure",
    "cvs": "cardiovascular system",
    "copd": "chronic obstructive pulmonary disease",
    "gfr": "glomerular filtration rate",
    "raas": "renin angiotensin aldosterone system",
    "adh": "antidiuretic hormone",
    "av": "atrioventricular",
    "sa": "sinoatrial",
    "sv": "stroke volume",
    "co": "cardiac output",
    "tpr": "total peripheral resistance",
    "map": "mean arterial pressure",
    "edv": "end diastolic volume",
    "esv": "end systolic volume",
    "gi": "gastrointestinal",
    "git": "gastrointestinal tract",
    "pku": "phenylketonuria",
    "hb": "haemoglobin hemoglobin",
    "hbs": "sickle haemoglobin hemoglobin",
    "atp": "adenosine triphosphate",
    "adp": "adenosine diphosphate",
    "camp": "cyclic amp",
    "cgmp": "cyclic gmp",
    "ip3": "inositol trisphosphate",
    "dag": "diacylglycerol",
    "gpcrs": "g-protein coupled receptors",
    "gpcr": "g-protein coupled receptor",
    "ivf": "in vitro fertilisation fertilization",
    "pgt": "preimplantation genetic testing",
    "icsi": "intracytoplasmic sperm injection",
    "er": "endoplasmic reticulum",
    "rer": "rough endoplasmic reticulum",
    "ser": "smooth endoplasmic reticulum",
    "sr": "sarcoplasmic reticulum",
    "rbc": "red blood cell erythrocyte",
    "wbc": "white blood cell leukocyte",
    "cn": "cranial nerve",
    "cn1": "olfactory nerve",
    "cn2": "optic nerve",
    "cn3": "oculomotor nerve",
    "cn4": "trochlear nerve",
    "cn5": "trigeminal nerve",
    "cn6": "abducens nerve",
    "cn7": "facial nerve",
    "cn8": "vestibulocochlear nerve",
    "cn9": "glossopharyngeal nerve",
    "cn10": "vagus nerve",
    "cn11": "accessory nerve",
    "cn12": "hypoglossal nerve",
    "fah": "fundamental anatomy histology",
    "mcb": "molecular cellular biology",
    "fpp": "foundations pharmacological principles",
    "pcrs": "physiology cardiovascular respiratory systems",
    "acrs": "anatomy cardiovascular respiratory systems",
    "nbss": "neural basis systemic structure neuroscience",
    "fbs": "foundations biomedical sciences",
    "pas": "physiology anatomy systems",
}

COMMON_STOP_WORDS = {
    "a", "about", "above", "after", "again", "against", "all", "am", "an", "and", "any", "are",
    "as", "at", "be", "because", "been", "before", "being", "below", "between", "both", "but",
    "by", "can", "could", "did", "do", "does", "doing", "down", "during", "each", "few", "for",
    "from", "further", "had", "has", "have", "having", "he", "her", "here", "hers", "herself",
    "him", "himself", "his", "how", "i", "if", "in", "into", "is", "it", "its", "itself", "just",
    "me", "more", "most", "my", "myself", "no", "nor", "not", "now", "of", "off", "on", "once",
    "only", "or", "other", "our", "ours", "ourselves", "out", "over", "own", "same", "she",
    "should", "so", "some", "such", "than", "that", "the", "their", "theirs", "them",
    "themselves", "then", "there", "these", "they", "this", "those", "through", "to", "too",
    "under", "until", "up", "very", "was", "we", "were", "what", "when", "where", "which",
    "while", "who", "whom", "why", "will", "with", "would", "you", "your", "yours", "yourself",
    "yourselves", "tell", "explain", "describe", "show", "give", "list", "please", "help"
}

class KCLLectureIndexEngine:
    def __init__(self, db_path: Path, config: ConfigManager, on_update_callback=None):
        self.db_path = db_path
        self._config = config
        self._on_update = on_update_callback
        self._lock = threading.Lock()
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        con = sqlite3.connect(str(self.db_path), check_same_thread=False)
        con.row_factory = sqlite3.Row
        return con

    def _init_db(self):
        with self._lock:
            con = self._get_connection()
            cur = con.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS lectures (
                    file_path TEXT PRIMARY KEY,
                    filename TEXT NOT NULL,
                    module TEXT NOT NULL,
                    block TEXT NOT NULL,
                    mtime REAL NOT NULL,
                    slide_count INTEGER NOT NULL,
                    indexed_at TEXT NOT NULL
                );
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS slides (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_path TEXT NOT NULL,
                    filename TEXT NOT NULL,
                    module TEXT NOT NULL,
                    block TEXT NOT NULL,
                    slide_num INTEGER NOT NULL,
                    slide_title TEXT,
                    body_text TEXT,
                    speaker_notes TEXT,
                    FOREIGN KEY(file_path) REFERENCES lectures(file_path) ON DELETE CASCADE
                );
            """)
            cur.execute("CREATE INDEX IF NOT EXISTS idx_slides_file ON slides(file_path);")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_slides_block ON slides(block);")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_slides_module ON slides(module);")

            cur.execute("""
                CREATE VIRTUAL TABLE IF NOT EXISTS slides_fts USING fts5(
                    slide_title,
                    body_text,
                    speaker_notes,
                    filename,
                    module,
                    block,
                    slide_num UNINDEXED,
                    file_path UNINDEXED,
                    tokenize = 'porter unicode61'
                );
            """)
            con.commit()
            con.close()

    def _get_root_dir(self) -> Optional[Path]:
        configured = str(self._config.get("LECTURES_DIR", "") or "").strip()
        if configured:
            p = Path(configured).expanduser().resolve()
            if p.exists():
                return p
        fallback = Path.home() / "Library/CloudStorage/OneDrive-King'sCollegeLondon"
        if fallback.exists():
            return fallback
        return None

    def sync_index(self, force: bool = False) -> dict:
        with self._lock:
            root = self._get_root_dir()
            if root is None or not root.exists():
                configured = str(self._config.get("LECTURES_DIR", "") or "").strip()
                msg = f"Lectures folder not found: {configured}" if configured else "No KCL lecture folder selected yet. Select your lecture folder in Settings or Onboarding."
                return {"success": False, "error": msg, "lectures": 0, "slides": 0}

            found_files = {}
            for p in root.rglob("*.pptx"):
                if p.name.startswith("~$"):
                    continue
                try:
                    rel = p.resolve().relative_to(root.resolve()).parts
                except Exception:
                    rel = p.relative_to(root).parts
                block = rel[-2] if len(rel) > 2 else "General"
                module = rel[-3] if len(rel) > 3 else "Year 1"
                found_files[str(p.resolve())] = {
                    "path": p.resolve(),
                    "filename": p.name,
                    "module": module,
                    "block": block,
                    "mtime": p.stat().st_mtime
                }

            con = self._get_connection()
            cur = con.cursor()

            cur.execute("SELECT file_path, mtime FROM lectures")
            existing = {row["file_path"]: row["mtime"] for row in cur.fetchall()}

            # Delete removed files
            deleted = set(existing.keys()) - set(found_files.keys())
            for del_path in deleted:
                cur.execute("DELETE FROM lectures WHERE file_path = ?", (del_path,))
                cur.execute("DELETE FROM slides WHERE file_path = ?", (del_path,))
                cur.execute("DELETE FROM slides_fts WHERE file_path = ?", (del_path,))

            to_index = []
            for path_str, meta in found_files.items():
                if force or path_str not in existing or abs(existing[path_str] - meta["mtime"]) > 0.001:
                    to_index.append(meta)

            indexed_count = 0
            for meta in to_index:
                try:
                    self._index_file(con, meta)
                    indexed_count += 1
                except Exception as e:
                    print(f"Error indexing {meta['filename']}: {e}")

            con.commit()

            cur.execute("SELECT COUNT(*) as lc FROM lectures")
            lec_count = cur.fetchone()["lc"]
            cur.execute("SELECT COUNT(*) as sc FROM slides")
            slide_count = cur.fetchone()["sc"]
            con.close()

            res = {
                "success": True,
                "ready": True,
                "total_lectures": lec_count,
                "total_slides": slide_count,
                "reindexed": indexed_count,
                "deleted": len(deleted)
            }

            if (indexed_count > 0 or len(deleted) > 0) and self._on_update:
                try:
                    self._on_update(res)
                except Exception:
                    pass

            return res

    def _index_file(self, con: sqlite3.Connection, meta: dict):
        cur = con.cursor()
        p = meta["path"]
        path_str = str(p)

        cur.execute("DELETE FROM lectures WHERE file_path = ?", (path_str,))
        cur.execute("DELETE FROM slides WHERE file_path = ?", (path_str,))
        cur.execute("DELETE FROM slides_fts WHERE file_path = ?", (path_str,))

        prs = Presentation(str(p))
        slide_count = 0

        for idx, slide in enumerate(prs.slides):
            slide_num = idx + 1
            slide_count += 1

            slide_title = ""
            try:
                if slide.shapes.title and slide.shapes.title.text:
                    slide_title = slide.shapes.title.text.strip().replace("\x0b", " ")
            except Exception:
                pass

            body_parts = []
            for shape in SlideParsEngine._get_flat_shapes(slide.shapes):
                if shape.has_text_frame:
                    if slide_title and shape == getattr(slide.shapes, "title", None):
                        continue
                    t = shape.text_frame.text.strip().replace("\x0b", "\n")
                    if t and t != slide_title:
                        body_parts.append(t)
                if shape.has_table:
                    for row in shape.table.rows:
                        row_txt = [c.text.strip().replace("\x0b", " ") for c in row.cells if c.text.strip()]
                        if row_txt:
                            body_parts.append(" | ".join(row_txt))

            if not slide_title and body_parts:
                first_line = body_parts[0].split("\n")[0].strip()
                if len(first_line) < 100:
                    slide_title = first_line

            body_text = "\n\n".join(body_parts)

            speaker_notes = ""
            try:
                if slide.has_notes_slide:
                    speaker_notes = slide.notes_slide.notes_text_frame.text.strip().replace("\x0b", "\n")
            except Exception:
                pass

            cur.execute("""
                INSERT INTO slides (file_path, filename, module, block, slide_num, slide_title, body_text, speaker_notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (path_str, meta["filename"], meta["module"], meta["block"], slide_num, slide_title, body_text, speaker_notes))

            cur.execute("""
                INSERT INTO slides_fts (slide_title, body_text, speaker_notes, filename, module, block, slide_num, file_path)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (slide_title, body_text, speaker_notes, meta["filename"], meta["module"], meta["block"], slide_num, path_str))

        indexed_at = datetime.now().isoformat()
        cur.execute("""
            INSERT INTO lectures (file_path, filename, module, block, mtime, slide_count, indexed_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (path_str, meta["filename"], meta["module"], meta["block"], meta["mtime"], slide_count, indexed_at))

    def get_status(self) -> dict:
        try:
            con = self._get_connection()
            cur = con.cursor()
            cur.execute("SELECT COUNT(*) as lc FROM lectures")
            lc = cur.fetchone()["lc"]
            cur.execute("SELECT COUNT(*) as sc FROM slides")
            sc = cur.fetchone()["sc"]
            cur.execute("SELECT MAX(indexed_at) as last_idx FROM lectures")
            last_idx = cur.fetchone()["last_idx"] or ""
            con.close()
            return {
                "ready": lc > 0,
                "total_lectures": lc,
                "total_slides": sc,
                "last_indexed": last_idx
            }
        except Exception as e:
            return {"ready": False, "total_lectures": 0, "total_slides": 0, "error": str(e)}

    def find_lecture_by_filename(self, filename: str) -> Optional[dict]:
        if not filename:
            return None
        con = self._get_connection()
        cur = con.cursor()
        q = filename.strip()
        stem = re.sub(r"\.pptx$", "", q, flags=re.I).strip()
        cur.execute("""
            SELECT file_path, filename, module, block, slide_count
            FROM lectures
            WHERE filename = ? OR filename = ? OR filename LIKE ? OR file_path LIKE ?
            LIMIT 1
        """, (q, f"{stem}.pptx", f"{stem}%", f"%{stem}%"))
        row = cur.fetchone()
        if row:
            res = dict(row)
            con.close()
            return res

        # 2. Normalized whitespace & token-overlap match across all database lectures
        rows = cur.execute("SELECT file_path, filename, module, block, slide_count FROM lectures").fetchall()
        con.close()

        clean_q = re.sub(r"\s+", " ", stem).strip().lower()
        q_tokens = [t for t in re.split(r"[^a-z0-9]+", clean_q) if t]

        # 2a. Match normalized stem equality or substring
        for r in rows:
            fn_stem = re.sub(r"\.pptx$", "", r["filename"], flags=re.I).strip()
            clean_fn = re.sub(r"\s+", " ", fn_stem).strip().lower()
            if clean_q == clean_fn or clean_q in clean_fn or clean_fn in clean_q:
                return dict(r)

        # 2b. Match token overlap
        best_row = None
        best_score = 0
        for r in rows:
            fn_stem = re.sub(r"\.pptx$", "", r["filename"], flags=re.I).strip()
            clean_fn = re.sub(r"\s+", " ", fn_stem).strip().lower()
            fn_tokens = set(re.split(r"[^a-z0-9]+", clean_fn))
            overlap = sum(1 for t in q_tokens if t in fn_tokens)
            if overlap > best_score:
                best_score = overlap
                best_row = r

        if best_row and best_score >= max(1, len(q_tokens) * 0.5):
            return dict(best_row)

        # 3. Fallback: Search the OneDrive directory directly
        try:
            root = self._get_root_dir()
            if root.exists():
                for p in root.rglob("*.pptx"):
                    if p.name.startswith("~$"):
                        continue
                    p_stem = re.sub(r"\.pptx$", "", p.name, flags=re.I).strip()
                    clean_p = re.sub(r"\s+", " ", p_stem).strip().lower()
                    if clean_q == clean_p or (len(q_tokens) >= 2 and all(t in clean_p for t in q_tokens)):
                        return {
                            "file_path": str(p.resolve()),
                            "filename": p.name,
                            "module": "",
                            "block": "",
                            "slide_count": 0
                        }
        except Exception:
            pass

        return None

    def get_lecture_slides(self, file_path: str) -> list[dict]:
        con = self._get_connection()
        cur = con.cursor()
        s = str(file_path)
        cur.execute("""
            SELECT slide_num, slide_title, body_text, speaker_notes
            FROM slides
            WHERE file_path = ? OR filename = ? OR filename = ?
            ORDER BY slide_num ASC
        """, (s, s, Path(s).name))
        rows = cur.fetchall()
        con.close()
        return [dict(r) for r in rows]

    def search_slides(self, user_query: str, block: str = "", limit: int = 15) -> list[dict]:
        # Fast incremental mtime sync before search (<5ms if no changes)
        try:
            self.sync_index(force=False)
        except Exception:
            pass

        con = self._get_connection()
        cur = con.cursor()

        fts_query = self._build_fts_query(user_query)
        rows = []
        if fts_query:
            try:
                if block:
                    cur.execute("""
                        SELECT filename, file_path, module, block, slide_num, slide_title, body_text, speaker_notes,
                               bm25(slides_fts, 5.0, 2.0, 1.0, 0.5, 0.2, 0.2, 0.0, 0.0) as rank
                        FROM slides_fts
                        WHERE slides_fts MATCH ? AND block = ?
                        ORDER BY rank ASC
                        LIMIT ?
                    """, (fts_query, block, limit))
                else:
                    cur.execute("""
                        SELECT filename, file_path, module, block, slide_num, slide_title, body_text, speaker_notes,
                               bm25(slides_fts, 5.0, 2.0, 1.0, 0.5, 0.2, 0.2, 0.0, 0.0) as rank
                        FROM slides_fts
                        WHERE slides_fts MATCH ?
                        ORDER BY rank ASC
                        LIMIT ?
                    """, (fts_query, limit))
                rows = [dict(r) for r in cur.fetchall()]
            except Exception as e:
                print(f"FTS5 MATCH note: {e}")

        # Fallback if no rows returned from FTS MATCH
        if not rows:
            raw_tokens = [w for w in re.findall(r'[A-Za-z0-9]+', user_query.lower()) if len(w) > 2 and w not in COMMON_STOP_WORDS]
            if raw_tokens:
                like_clauses = " OR ".join(["body_text LIKE ? OR slide_title LIKE ?" for _ in raw_tokens[:3]])
                params = []
                for t in raw_tokens[:3]:
                    params.extend([f"%{t}%", f"%{t}%"])
                if block:
                    params.append(block)
                    sql = f"SELECT filename, file_path, module, block, slide_num, slide_title, body_text, speaker_notes, 0 as rank FROM slides WHERE ({like_clauses}) AND block = ? LIMIT {limit}"
                else:
                    sql = f"SELECT filename, file_path, module, block, slide_num, slide_title, body_text, speaker_notes, 0 as rank FROM slides WHERE {like_clauses} LIMIT {limit}"
                cur.execute(sql, params)
                rows = [dict(r) for r in cur.fetchall()]

        con.close()
        return rows

    def _build_fts_query(self, text: str) -> str:
        raw_tokens = re.findall(r'[A-Za-z0-9]+', text.lower())
        tokens = [t for t in raw_tokens if len(t) >= 2 and t not in COMMON_STOP_WORDS]
        if not tokens:
            tokens = [t for t in raw_tokens if len(t) >= 2]
        if not tokens:
            return ""

        parts = []
        for t in tokens:
            if t in MEDICAL_ACRONYMS:
                exp = MEDICAL_ACRONYMS[t]
                clean_exp = " ".join(re.findall(r'[A-Za-z0-9]+', exp))
                parts.append(f'("{t}" OR "{clean_exp}")')
            else:
                parts.append(f'"{t}"')

        return " OR ".join(parts)

class KCLOneDriveEngine:
    def __init__(self, config: ConfigManager, index_engine: Optional[KCLLectureIndexEngine] = None):
        self._config = config
        self._index = index_engine
        self._cache: dict = {}
        self._lectures_index: list = []

    def set_index_engine(self, index_engine: KCLLectureIndexEngine):
        self._index = index_engine

    def get_root_dir(self) -> Path:
        configured = str(self._config.get("LECTURES_DIR", "") or "").strip()
        if configured:
            p = Path(configured).expanduser().resolve()
            if p.exists():
                return p

        fallback = Path.home() / "Library/CloudStorage/OneDrive-King'sCollegeLondon"
        if fallback.exists():
            return fallback

        return APP_DATA_DIR

    def list_lectures(self) -> list[dict]:
        # Fast path: query persistent SQLite index if available
        if self._index:
            try:
                con = self._index._get_connection()
                cur = con.cursor()
                cur.execute("SELECT file_path, filename, module, block, slide_count FROM lectures ORDER BY module, block, filename")
                rows = cur.fetchall()
                con.close()
                if rows:
                    results = [{
                        "id": r["file_path"],
                        "name": Path(r["filename"]).stem,
                        "filename": r["filename"],
                        "path": r["file_path"],
                        "module": r["module"],
                        "block": r["block"],
                        "slide_count": r["slide_count"],
                        "size_kb": 0
                    } for r in rows]
                    self._lectures_index = results
                    return results
            except Exception as e:
                print(f"Index query note: {e}")

        # Fallback disk scan
        root = self.get_root_dir()
        if not root.exists():
            return []

        results = []
        try:
            for p in root.rglob("*.pptx"):
                if p.name.startswith("~$"):
                    continue
                try:
                    try:
                        rel = p.resolve().relative_to(root.resolve()).parts
                    except Exception:
                        rel = p.relative_to(root).parts
                    block = rel[-2] if len(rel) > 2 else "General"
                    module = rel[-3] if len(rel) > 3 else "Year 1"
                    results.append({
                        "id": str(p),
                        "name": p.stem,
                        "filename": p.name,
                        "path": str(p),
                        "module": module,
                        "block": block,
                        "size_kb": round(p.stat().st_size / 1024),
                        "slide_count": 0
                    })
                except Exception:
                    pass
        except Exception as exc:
            print(f"Error scanning KCL OneDrive: {exc}")

        results.sort(key=lambda x: (x["module"], x["block"], x["name"]))
        self._lectures_index = results
        return results

    def get_blocks_grouped(self) -> dict:
        lectures = self.list_lectures()
        grouped = {}
        for lec in lectures:
            mod = lec.get("module") or "General"
            blk = lec.get("block") or "General"
            if mod not in grouped:
                grouped[mod] = {
                    "module_code": mod,
                    "module_name": KCL_MODULES_INFO.get(mod, mod),
                    "blocks": {}
                }
            if blk not in grouped[mod]["blocks"]:
                grouped[mod]["blocks"][blk] = {
                    "block_code": blk,
                    "block_name": KCL_BLOCKS_INFO.get(blk, blk),
                    "count": 0,
                    "lectures": []
                }
            grouped[mod]["blocks"][blk]["count"] += 1
            grouped[mod]["blocks"][blk]["lectures"].append({
                "name": lec["name"],
                "path": lec["path"]
            })

        result = {}
        ordered_modules = ["FBS", "PAS"] + [m for m in sorted(grouped.keys()) if m not in ("FBS", "PAS")]
        for mod in ordered_modules:
            if mod in grouped:
                mod_data = grouped[mod]
                result[mod] = {
                    "module_code": mod_data["module_code"],
                    "module_name": mod_data["module_name"],
                    "blocks": sorted(
                        list(mod_data["blocks"].values()),
                        key=lambda b: b["block_code"]
                    )
                }
        return result

    def get_lecture_text(self, pptx_path: str) -> dict:
        if pptx_path in self._cache:
            return self._cache[pptx_path]

        p = Path(pptx_path)

        # Fast path: check SQLite cache
        if self._index:
            try:
                db_slides = self._index.get_lecture_slides(str(p.resolve()))
                if not db_slides:
                    db_slides = self._index.get_lecture_slides(str(p))
                if db_slides:
                    slides_text = []
                    full_text_parts = []
                    for s in db_slides:
                        s_num = s["slide_num"]
                        title = s.get("slide_title") or ""
                        body = s.get("body_text") or ""
                        notes = s.get("speaker_notes") or ""
                        parts = []
                        if title:
                            parts.append(f"Title: {title}")
                        if body:
                            parts.append(body)
                        if notes:
                            parts.append(f"Speaker Notes: {notes}")
                        combined = "\n\n".join(parts)
                        slides_text.append({"slide_num": s_num, "text": combined})
                        full_text_parts.append(f"[Slide {s_num}]\n{combined}")

                    res = {
                        "success": True,
                        "name": p.stem,
                        "filename": p.name,
                        "path": str(p),
                        "total_slides": len(db_slides),
                        "slides_text": slides_text,
                        "full_text": "\n\n".join(full_text_parts)
                    }
                    self._cache[pptx_path] = res
                    return res
            except Exception as e:
                print(f"SQLite slide read note: {e}")

        # Fallback: parse .pptx directly if not in database
        if not p.exists():
            return {"success": False, "error": f"File not found: {pptx_path}"}

        try:
            prs = Presentation(str(p))
            slides_text = []
            full_text_parts = []
            for idx, slide in enumerate(prs.slides):
                _, text = SlideParsEngine.extract_slide_data(slide)
                if text.strip():
                    slides_text.append({"slide_num": idx + 1, "text": text.strip()})
                    full_text_parts.append(f"[Slide {idx + 1}]\n{text.strip()}")

            res = {
                "success": True,
                "name": p.stem,
                "filename": p.name,
                "path": str(p),
                "total_slides": len(prs.slides),
                "slides_text": slides_text,
                "full_text": "\n\n".join(full_text_parts)
            }
            self._cache[pptx_path] = res
            return res
        except Exception as exc:
            return {"success": False, "error": str(exc)}

class GeminiChatbotEngine:
    def __init__(self, config: ConfigManager, onedrive_engine: KCLOneDriveEngine, state: AppState, index_engine: Optional[KCLLectureIndexEngine] = None):
        self._config = config
        self._onedrive = onedrive_engine
        self._state = state
        self._index = index_engine
        self._gemini_cache: dict[str, dict] = {}       # lecture_path -> {cache_name, expire_time, model}
        self._session_lecture_cache: dict[str, str] = {} # lecture_path -> full_text

    def set_index_engine(self, index_engine: KCLLectureIndexEngine):
        self._index = index_engine

    def _get_or_create_context_cache(self, api_key: str, model: str, lecture_name: str, full_text: str) -> Optional[str]:
        est_tokens = len(full_text) // 4
        # Gemini Context Caching API requires at least 32,768 tokens
        if est_tokens < 32768:
            return None

        now = time.time()
        cache_entry = self._gemini_cache.get(lecture_name)
        if cache_entry and cache_entry.get("expire_time", 0) > (now + 60) and cache_entry.get("model") == model:
            return cache_entry.get("cache_name")

        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/cachedContents?key={api_key}"
            clean_disp = re.sub(r'[^A-Za-z0-9_\-]', '_', lecture_name)[:40]
            payload = {
                "model": f"models/{model}",
                "displayName": f"LectureCache_{clean_disp}",
                "contents": [
                    {
                        "role": "user",
                        "parts": [{"text": f"KCL Lecture Content for {lecture_name}:\n\n{full_text}"}]
                    }
                ],
                "ttl": "600s"
            }
            res = requests.post(url, json=payload, timeout=25)
            if res.status_code in (200, 201):
                data = res.json()
                c_name = data.get("name")
                if c_name:
                    self._gemini_cache[lecture_name] = {
                        "cache_name": c_name,
                        "model": model,
                        "expire_time": now + 600
                    }
                    return c_name
        except Exception as e:
            print(f"Context caching note (falling back to session memory): {e}")
        return None

    def _get_banter_response(self) -> str:
        is_festive = is_christmas_season()
        if is_festive:
            roasts = [
                (
                    "🎄 **Even Santa checks his list twice, but you're not even checking your lecture slides!** 😭\n\n"
                    "Guy's and St Thomas' hospitals don't shut down for the holidays, and neither does preclinical medicine. Put down the mince pies, stop procrastinating, and let's actually revise."
                ),
                (
                    "❄️ **Diagnosis: Acute Holiday Revision Inertia.**\n\n"
                    "**Rx:** 500mg of *Stop Being Jobless* taken with a warm festive tea stat. The General Medical Council doesn't take Christmas off, and neither should your OSCE prep. Pick a slide and let's lock in!"
                ),
                (
                    "🦌 **Rudolph has better directional navigation than your anatomical relations right now.**\n\n"
                    "Stop asking random off-topic questions. You have exams around the corner—ask me about a nerve pathway, a receptor, or a slide!"
                )
            ]
            return random.choice(roasts)
        roasts = [
            "🚨 **Stop procrastinating, doctor.** The GMC isn't examining you on that! Drop the side quests and ask about your lecture slides, anatomical relations, or drug mechanisms.",
            "💀 **Diagnosis: Acute Procrastination.** Rx: 500mg of *Lock In* stat. Guy's and St Thomas' won't staff themselves—pick a slide and let's revise.",
            "📉 **API Quota Preserved.** I'm saving your tokens for high-yield medicine. Throw a real clinical pearl, pathway, or slide question at me.",
            "🛑 **Focus, doctor.** You have OSCEs looming and exams to conquer. What physiological mechanism or anatomical relation are we tackling right now?",
            "🩺 **The patient is crashing and you're chatting off-topic.** Stop procrastinating! Ask me about an anatomical relation or clinical sign so we can get back to work."
        ]
        return random.choice(roasts)

    def _check_lecture_relevance(self, message: str, slide_idx: int = -1, scope: str = "slide", block: str = "") -> tuple[bool, str]:
        """
        Evaluates whether the user's message is relevant to preclinical/medical lecture study
        BEFORE dispatching requests to Gemini API, preserving precious API prompt quota.
        """
        text = (message or "").strip()
        if not text:
            return False, "Please enter a question or topic about your lectures to get started."

        lower = text.lower()
        cleaned = re.sub(r"[^\w\s]", "", lower).strip()

        # 1. Polite greetings / single-word conversational checks (instant reply without wasting quota)
        greetings = {
            "hi", "hello", "hey", "hiya", "howdy", "sup", "yo", "good morning",
            "good afternoon", "good evening", "how are you", "who are you", "what are you"
        }
        if cleaned in greetings:
            is_festive = is_christmas_season()
            if is_festive:
                return False, (
                    "🎄 **Dr. Fledge on holiday ward rounds.** Grab a warm cup of coffee, stay focused, and let's lock in! "
                    "Ask about any slide, anatomical relation, pathway, or drug mechanism."
                )
            return False, (
                "👋 **Dr. Fledge on duty.** Ready to lock in? Ask about any slide, anatomical relation, pathway, or drug mechanism."
            )

        # 2. Canvas & Study Whitelist (Immediately allowed to pass through to Gemini)
        study_terms = [
            "slide", "diagram", "image", "label", "summarize", "summary", "explain",
            "mechanism", "pathway", "function", "innervation", "blood supply", "attachment",
            "relation", "relations", "difference between", "compare", "clinical pearl",
            "ilo", "learning objective", "flashcard", "notes", "test me", "question",
            "what does this mean", "what is this", "what are the", "why does", "how does",
            "pigeon", "wisdom", "lock in", "tuff pigeon", "dr pigeon", "fledge", "dr fledge",
            "medical joke", "doctor joke", "anatomy joke", "pharma joke", "hospital joke", "joke"
        ]
        if any(term in lower for term in study_terms):
            return True, ""

        # 3. Obvious off-topic conversational requests (time/date, weather, trivia, non-medical coding, tech support)
        off_topic_patterns = [
            r"\b(what day is it|what is today(?:'s)? date|what date is it|what's the date|what time is it|current time|current date)\b",
            r"\b(what(?:'s| is) the weather|weather forecast|is it raining|temperature outside)\b",
            r"\b(write a poem|write a song|write a story|sing a song|sing me a song)\b",
            r"\b(who won the|who is the president|who is the prime minister|capital of \w+)\b",
            r"\b(write a react|write javascript|write python script to scrape|build a website|coding help|fix my code|debug my code)\b",
            r"\b(raspberry pi|tailscale|wireguard|vpn|ssh|docker|kubernetes|homelab|home lab|linux|ubuntu|arch|debian|install windows|gpu|pc build|gaming|playstation|xbox|fortnite|minecraft|pokemon)\b",
            r"\b(crypto|bitcoin|ethereum|stocks|trading|forex|drop shipping|dropshipping|side hustle)\b",
            r"\b(netflix|movie recommendations|tv show|anime|dating advice|tinder|hinge|girlfriend|boyfriend)\b",
            r"\b(how do i (?:set up|install|configure|build|buy|fix|code|program|hack|root))\b",
            r"\b(what is 2\s*\+\s*2|calculate \d+\s*[\+\-\*\/]\s*\d+)\b",
        ]
        for pat in off_topic_patterns:
            if re.search(pat, lower):
                return False, self._get_banter_response()

        # 4. Medical / Preclinical Keywords & Stems
        medical_keywords = [
            "artery", "vein", "nerve", "muscle", "bone", "ligament", "tendon", "fascia",
            "heart", "lung", "brain", "liver", "kidney", "spleen", "pancreas", "stomach",
            "bowel", "colon", "duodenum", "jejunum", "ileum", "esophagus", "oesophagus",
            "trachea", "bronchus", "bronchi", "diaphragm", "larynx", "pharynx", "thyroid",
            "thymus", "adrenal", "pituitary", "ovary", "testis", "uterus", "prostate",
            "aorta", "vena cava", "atrium", "ventricle", "pericardium", "pleura", "peritoneum",
            "cranium", "cranial", "thoracic", "thorax", "abdomen", "abdominal", "pelvis",
            "crus", "crura", "hiatus", "foramen", "sulcus", "gyrus", "cortex", "medulla",
            "action potential", "depolarization", "repolarization", "synapse", "neurotransmitter",
            "receptor", "agonist", "antagonist", "enzyme", "substrate", "inhibitor", "atp",
            "adp", "camp", "phosphorylation", "glycolysis", "krebs", "insulin", "glucagon",
            "hormone", "cortisol", "aldosterone", "renin", "angiotensin", "cardiac output",
            "blood pressure", "stroke volume", "ventilation", "perfusion", "hypoxia",
            "hypercapnia", "acidosis", "alkalosis", "filtration", "reabsorption", "reflex",
            "epithelium", "squamous", "columnar", "cuboidal", "endothelium", "collagen",
            "fibroblast", "macrophage", "neutrophil", "lymphocyte", "erythrocyte", "platelet",
            "myelin", "axon", "dendrite", "sarcomere", "actin", "myosin", "troponin",
            "drug", "dose", "clearance", "half-life", "efficacy", "potency", "infection",
            "bacteria", "virus", "inflammation", "necrosis", "apoptosis", "edema", "oedema",
            "thrombosis", "embolism", "infarction", "ischemia", "ischaemia", "neoplasia",
            "carcinoma", "hyperplasia", "hypertrophy", "atrophy", "metaplasia", "dysplasia",
            "mutation", "syndrome", "disease", "disorder", "symptom", "sign", "diagnosis",
            "tract", "pain", "lesion", "deficit", "pathology", "physiology", "anatomy",
            "histology", "embryology", "pharmacology", "clinical", "patient", "presentation",
            "treatment", "therapy", "decussate", "decussation"
        ]
        if any(kw in lower for kw in medical_keywords):
            return True, ""

        med_affixes = r"\b\w*(itis|ectomy|otomy|ostomy|emia|aemia|uria|pathy|genic|lysis|genesis|trophy|plasia|oma|blast|cyte|ase|cardio|pulmo|gastro|nephro|neuro|osteo|myo|vasculo|angio)\w*\b"
        if re.search(med_affixes, lower):
            return True, ""

        # 5. Overlap with Active Slide Text or OCR Labels
        target_idx = slide_idx if slide_idx >= 0 else self._state.current_slide_idx
        if self._state.presentation and 0 <= target_idx < self._state.total_slides:
            slide = self._state.presentation.slides[target_idx]
            _, slide_text = SlideParsEngine.extract_slide_data(slide)
            slide_words = set(re.findall(r"\b[a-zA-Z]{4,}\b", (slide_text or "").lower()))
            ocr_words = set(re.findall(r"\b[a-zA-Z]{4,}\b", " ".join(m.get("text", "") for m in self._state.ocr_results).lower()))
            combined_slide = slide_words.union(ocr_words)
            query_words = set(re.findall(r"\b[a-zA-Z]{4,}\b", lower))
            if query_words.intersection(combined_slide):
                return True, ""

        # 6. Overlap with SQLite FTS5 Lecture Index (Only with substantive non-generic tokens)
        generic_query_words = {
            "want", "set", "like", "make", "know", "think", "get", "got", "go", "going",
            "take", "come", "use", "using", "used", "see", "look", "way", "need", "good",
            "new", "first", "last", "put", "also", "well", "much", "many", "let", "say",
            "said", "try", "tried", "ask", "asked", "work", "working", "one", "two", "three",
            "run", "running", "thing", "things", "pi", "something", "anything", "nothing",
            "someone", "anyone", "everyone", "really", "actually", "probably", "maybe",
            "still", "always", "never", "keep", "start", "stop", "help", "please", "can",
            "could", "would", "should", "might", "must", "give", "tell", "show", "find"
        }
        if self._index:
            try:
                substantive_tokens = [
                    w for w in re.findall(r"\b[a-zA-Z]{3,}\b", lower)
                    if w not in COMMON_STOP_WORDS and w not in generic_query_words
                ]
                if substantive_tokens:
                    search_str = " ".join(substantive_tokens[:5])
                    hits = self._index.search_slides(search_str, block=block if scope in ("block", "onedrive_block") else "", limit=1)
                    if hits:
                        return True, ""
            except Exception:
                pass

        # 7. Fallback: If no medical or lecture terms found, deliver banter and protect user prompt quota
        return False, self._get_banter_response()

    def chat(self, message: str, scope: str = "slide", lecture_path: str = "", slide_idx: int = -1, history: list = None, block: str = "", mode: str = "quick") -> dict:
        api_key = self._config.get_chat_api_key()
        if not api_key:
            return {"success": False, "error": "Gemini AI Tutor API Key is missing. Open Settings to set it."}

        mode = (mode or "quick").lower().strip()
        lower_msg = (message or "").lower()
        # If the user asks for Dr. Fledge wisdom or unsolicited wisdom, always route to Quick (gemini-3.5-flash-lite)
        # to ensure it NEVER burns the 20 req/day Complex model quota
        if "wisdom" in lower_msg or "pigeon" in lower_msg or "fledge" in lower_msg:
            mode = "quick"

        # Pre-filter check to prevent wasting user API prompt quota on non-lecture/off-topic questions
        is_relevant, filter_reply = self._check_lecture_relevance(
            message=message,
            slide_idx=slide_idx,
            scope=scope,
            block=block
        )
        if not is_relevant:
            print(f"[Gemini Chat / Pre-Filter] Intercepted non-lecture query ('{message[:40]}...'). Preserved API prompt quota.")
            return {
                "success": True,
                "reply": filter_reply,
                "answer": filter_reply,
                "model": "lecture-guardrail",
                "fallback_used": False,
                "mode": mode,
                "truncated": False,
                "filtered": True
            }
        if mode == "complex":
            primary_model = "gemini-4-hydrogen"
            cascade = [
                "gemini-4-hydrogen",
                "gemini-3.8-flash",
                "gemini-3-flash-preview",
                "gemini-3.5-flash",
                "gemini-3.5-flash-lite",
            ]
        else:
            primary_model = "gemini-3.5-flash-lite"
            cascade = [
                "gemini-3.5-flash-lite",
                "gemini-3-flash-preview",
                "gemini-flash-lite-latest",
            ]

        masked = mask_api_key(api_key)
        print(f"[Gemini Chat / AI Tutor] Request initiated | Key: AI Tutor Key ({masked}) | Mode: {mode.title()} | Target Model: {primary_model}")
        context_sections = []
        cached_content_id: Optional[str] = None

        # 1. Slide Context
        target_idx = slide_idx if slide_idx >= 0 else self._state.current_slide_idx
        if self._state.presentation and 0 <= target_idx < self._state.total_slides:
            slide = self._state.presentation.slides[target_idx]
            _, slide_text = SlideParsEngine.extract_slide_data(slide)
            active_masks = [m.get("text", "") for m in self._state.ocr_results if m.get("text") and not m.get("is_redaction")]
            diag_view = self._state.current_view_mode
            pres_name = self._state.pptx_path.name if self._state.pptx_path else "Presentation.pptx"

            context_sections.append(
                f"ACTIVE PRESENTATION: {pres_name}\n"
                f"ACTIVE SLIDE: [[{pres_name} | Slide {target_idx + 1}]]\n"
                f"ACTIVE DIAGRAM VIEW: {diag_view}\n"
                f"ACTIVE SLIDE TEXT:\n{slide_text}\n"
                f"OCR LABELS ON THIS SLIDE:\n{', '.join(active_masks) if active_masks else 'None'}"
            )

        # 2a. Specific KCL Block Scope (Fast FTS5 Search)
        if scope in ("block", "onedrive_block") and block:
            mod_name = "KCL MBBS Module"
            blk_name = KCL_BLOCKS_INFO.get(block.upper(), block)
            top_slides = []
            if self._index:
                top_slides = self._index.search_slides(message, block=block, limit=15)

            context_sections.append(
                f"SELECTED KCL MBBS BLOCK: {block} — {blk_name}\n"
                f"TOTAL RELEVANT SLIDES RETRIEVED VIA FTS5: {len(top_slides)}\n"
            )

            for s in top_slides:
                slide_snippet = (
                    f"[[{s['filename']} | Slide {s['slide_num']}]]\n"
                    f"Title: {s['slide_title'] or 'Slide ' + str(s['slide_num'])}\n"
                    f"Content:\n{s['body_text']}"
                )
                if s.get("speaker_notes"):
                    slide_snippet += f"\nSpeaker Notes: {s['speaker_notes']}"
                context_sections.append(slide_snippet)

        # 2b. Specific KCL OneDrive Lecture Context (Session Memory & Context Caching)
        elif scope == "onedrive_lecture" and lecture_path:
            p_str = str(lecture_path)
            if p_str in self._session_lecture_cache:
                full_text = self._session_lecture_cache[p_str]
                lec_name = Path(p_str).name
            else:
                lec_data = self._onedrive.get_lecture_text(p_str)
                if lec_data.get("success"):
                    full_text = lec_data.get("full_text", "")
                    lec_name = lec_data.get("filename", Path(p_str).name)
                    self._session_lecture_cache[p_str] = full_text
                else:
                    full_text = ""
                    lec_name = Path(p_str).name
                    context_sections.append(f"NOTE: Could not load lecture {lecture_path}: {lec_data.get('error')}")

            if full_text:
                cached_content_id = self._get_or_create_context_cache(api_key, primary_model, lec_name, full_text)
                if not cached_content_id:
                    context_sections.append(
                        f"SELECTED KCL ONEDRIVE LECTURE: {lec_name}\n"
                        f"LECTURE CONTENT (Cached in Session Memory):\n{full_text[:30000]}"
                    )

        # 2c. Active Lecture Scope (All Slides)
        elif scope == "lecture" and self._state.pptx_path:
            p_str = str(self._state.pptx_path)
            if p_str in self._session_lecture_cache:
                full_text = self._session_lecture_cache[p_str]
                lec_name = self._state.pptx_path.name
            else:
                lec_data = self._onedrive.get_lecture_text(p_str)
                if lec_data.get("success"):
                    full_text = lec_data.get("full_text", "")
                    lec_name = lec_data.get("filename", self._state.pptx_path.name)
                    self._session_lecture_cache[p_str] = full_text
                else:
                    full_text = ""
                    lec_name = self._state.pptx_path.name

            if full_text:
                cached_content_id = self._get_or_create_context_cache(api_key, primary_model, lec_name, full_text)
                if not cached_content_id:
                    context_sections.append(
                        f"ACTIVE LECTURE: {lec_name}\n"
                        f"FULL LECTURE CONTENT:\n{full_text[:30000]}"
                    )

        # 3. Global All KCL Lectures Search Context (Fast SQLite FTS5 BM25 Two-Stage Retrieval)
        elif scope == "all_onedrive":
            top_slides = []
            if self._index:
                top_slides = self._index.search_slides(message, limit=15)

            total_lec = 0
            if self._index:
                try:
                    st = self._index.get_status()
                    total_lec = st.get("total_lectures") or total_lec
                except Exception:
                    pass
            if not total_lec and self._onedrive:
                try:
                    total_lec = len(self._onedrive.list_all_lectures())
                except Exception:
                    pass

            lec_str = f"ALL {total_lec} " if total_lec else "ALL "
            context_sections.append(
                f"TOTAL RELEVANT SLIDES FOUND ACROSS {lec_str}KCL YEAR 1 LECTURES (SQLite FTS5): {len(top_slides)}\n"
            )

            for s in top_slides:
                slide_snippet = (
                    f"[[{s['filename']} | Slide {s['slide_num']}]]\n"
                    f"Module: {s['module']} | Block: {s['block']} | Slide {s['slide_num']}\n"
                    f"Title: {s['slide_title'] or 'Slide ' + str(s['slide_num'])}\n"
                    f"Content:\n{s['body_text']}"
                )
                if s.get("speaker_notes"):
                    slide_snippet += f"\nSpeaker Notes: {s['speaker_notes']}"
                context_sections.append(slide_snippet)

        active_lec_count = 0
        if self._index:
            try:
                st = self._index.get_status()
                active_lec_count = st.get("total_lectures") or active_lec_count
            except Exception:
                pass
        if not active_lec_count and self._onedrive:
            try:
                active_lec_count = len(self._onedrive.list_all_lectures())
            except Exception:
                pass

        lec_desc = f"all {active_lec_count} lectures in " if active_lec_count else "the lectures in "
        is_festive = is_christmas_season()
        festive_clause = ""
        if is_festive:
            festive_clause = (
                "\n7. FESTIVE / CHRISTMAS MODE (ACTIVE):\n"
                "   - It is the Christmas season! While strictly adhering to the brevity, high-yield facts, and zero-wasted-words rules, slightly sprinkle in witty British Christmas medical flair:\n"
                "     * Subtle seasonal/winter analogies (e.g. comparing anatomical structures or pathways to holiday tinsel, cold-induced vasoconstriction, winter ward shifts, mince pies).\n"
                "     * Tasteful holiday touches (e.g. 🎄, ❄️, 🩺) and warm, dry festive banter where appropriate.\n"
                "     * If asked for a medical joke or off-topic question, give a sharp, hilarious Christmas-medical themed punchline (under 50 words) before telling them to get back to revising."
            )

        system_instruction = (
            "You are Dr. Fledge (The Tuff Pigeon Doctor), Senior Registrar in Preclinical Medicine and AI Medical Tutor assisting an MBBS student at King's College London (KCL).\n"
            "Persona: Dr. Fledge—intellectually rigorous, clinically sharp, and direct. You have a delightfully dry, witty British medical style, but you value efficiency and brevity above all else. You deliver high-yield facts with zero wasted words.\n"
            f"You have direct access to active slides, OCR labels, and {lec_desc}their KCL OneDrive lecture repository (FBS, PAS, IVBCP modules).\n\n"
            "CORE BREVITY & CONCISENESS RULES (STRICT & ABSOLUTE):\n"
            "1. HEAVILY SHORTENED & DIRECT: Keep all responses heavily shortened down, punchy, and token-efficient. Strictly no fluff, no filler, no essay-length answers.\n"
            "2. NO THEATRICAL ROLEPLAY INTROS: Strictly NEVER include dramatic stage directions (e.g. '*adjusts stethoscope*', '*peers over spectacles*'), roleplay scene-setting, or patronizing preambles (e.g. 'Put down the coffee...', 'The GMC does not wait...'). Jump STRAIGHT into the answer with zero preamble.\n"
            "3. HIGH-YIELD BULLETS: Use concise bullet points, bold key terms, and short cause-and-effect explanations instead of long rambling paragraphs.\n"
            "4. CITATIONS: Cite slide sources concisely inline using double brackets: [[Filename.pptx | Slide X]].\n"
            "5. JOKES & IRRELEVANT QUESTIONS (KEEP EXTREMELY SHORT & TOKEN-EFFICIENT):\n"
            "   - If asked for a medical joke (e.g. 'tell me a medical joke'): Deliver ONE sharp, witty medical joke or punchline immediately (1-2 sentences), followed by a 1-sentence quick Dr. Fledge nudge to lock back in. Never write multi-paragraph dissertations, monologues, or theatrical scene-setting for jokes! Total length must be under 50 words.\n"
            "   - If asked completely off-topic questions (tech, coding, gaming, pop culture): Give a very brief, witty 1-2 sentence roast/reminder to focus on the MBBS curriculum, and stop. Do not elaborate.\n"
            "   - If asked for unsolicited wisdom: Deliver 1-3 short, memorable, high-yield clinical pearls or exam mnemonics.\n"
            "6. CLEAN FORMATTING: Use GitHub-flavored Markdown, concise notation (O2, CO2, Ca2+, Na+, ATP), and clean conclusions."
            f"{festive_clause}"
        )

        full_prompt = (
            f"{system_instruction}\n\n"
            f"=== GROUNDED LECTURE & SLIDE CONTEXT ===\n"
            f"{chr(10).join(context_sections)}\n"
            f"========================================\n\n"
            f"STUDENT QUESTION:\n{message}"
        )

        payload = {
            "contents": [{"parts": [{"text": full_prompt}]}],
            "generationConfig": {"temperature": 0.3, "maxOutputTokens": 1536}
        }

        # If Gemini Context Caching handle is available, attach it
        if cached_content_id:
            payload["cachedContent"] = cached_content_id
            payload["contents"] = [{"parts": [{"text": f"{system_instruction}\n\nSTUDENT QUESTION:\n{message}"}]}]

        models_to_try = []
        for m in cascade:
            if m and m not in models_to_try:
                models_to_try.append(m)

        last_error = ""
        for idx, attempt_model in enumerate(models_to_try):
            is_primary = (idx == 0)
            max_retries = 3 if is_primary else 2
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{attempt_model}:generateContent?key={api_key}"
            for retry in range(max_retries):
                try:
                    resp = requests.post(url, json=payload, timeout=90)
                    if resp.status_code == 200:
                        data = resp.json()
                        candidate = data.get("candidates", [{}])[0]
                        parts = candidate.get("content", {}).get("parts", [])
                        ans = parts[0].get("text", "") if parts else ""
                        finish_reason = candidate.get("finishReason", "")
                        truncated = False

                        # If response hit token limit, perform an automatic continuation call to conclude cleanly
                        if finish_reason == "MAX_TOKENS" and ans:
                            try:
                                cont_prompt = (
                                    f"{full_prompt}\n\n"
                                    f"--- PARTIAL ASSISTANT RESPONSE SO FAR ---\n"
                                    f"{ans[-1500:]}\n\n"
                                    f"--- INSTRUCTION ---\n"
                                    f"Continue directly from the exact word where you stopped above. "
                                    f"Do NOT repeat what was already written. Complete the remaining mechanisms, pathways, and clinical pearls cleanly."
                                )
                                cont_payload = {
                                    "contents": [{"parts": [{"text": cont_prompt}]}],
                                    "generationConfig": {"temperature": 0.3, "maxOutputTokens": 1536}
                                }
                                cont_resp = requests.post(url, json=cont_payload, timeout=70)
                                if cont_resp.status_code == 200:
                                    cont_data = cont_resp.json()
                                    cont_parts = cont_data.get("candidates", [{}])[0].get("content", {}).get("parts", [])
                                    cont_ans = cont_parts[0].get("text", "") if cont_parts else ""
                                    ans = ans.rstrip() + "\n\n" + cont_ans.lstrip()
                                    if cont_data.get("candidates", [{}])[0].get("finishReason") == "MAX_TOKENS":
                                        truncated = True
                                else:
                                    truncated = True
                            except Exception:
                                truncated = True

                        if attempt_model != primary_model:
                            print(f"[Gemini Chat / AI Tutor] Fallback model '{attempt_model}' succeeded with AI Tutor Key ({masked}).")
                        else:
                            print(f"[Gemini Chat / AI Tutor] Primary model '{primary_model}' succeeded with AI Tutor Key ({masked}).")

                        return {
                            "success": True,
                            "reply": ans,
                            "answer": ans,
                            "model": attempt_model,
                            "primary_model": primary_model,
                            "fallback_used": (attempt_model != primary_model),
                            "mode": mode,
                            "truncated": truncated
                        }
                    elif resp.status_code == 503:
                        last_error = f"Gemini Error 503 ({attempt_model}): {resp.text}"
                        wait_sec = 2.5 * (retry + 1)
                        if retry < max_retries - 1:
                            print(f"[Gemini Chat / AI Tutor] {attempt_model} experiencing high demand (503). Retrying in {wait_sec}s (attempt {retry + 1}/{max_retries}) [AI Tutor Key: {masked}]...")
                            time.sleep(wait_sec)
                            continue
                        else:
                            print(f"[Gemini Chat / AI Tutor] {attempt_model} high demand retries exhausted. Cascading to next candidate...")
                            break
                    elif resp.status_code == 429:
                        last_error = f"Gemini Error 429 ({attempt_model}): {resp.text}"
                        is_daily_limit = any(term in resp.text for term in [
                            "limit: 20",
                            "requests per day limit exceeded",
                            "GenerateRequestsPerDay",
                            "generate_content_free_tier_requests"
                        ])
                        if is_daily_limit:
                            print(f"[Gemini Chat / AI Tutor] ⚠️ {attempt_model} daily quota exhausted (20 req/day). Cascading to next candidate [AI Tutor Key: {masked}]...")
                            break

                        wait_sec = 2.0 * (retry + 1)
                        if retry < max_retries - 1:
                            print(f"[Gemini Chat / AI Tutor] {attempt_model} rate-limited (429). Retrying in {wait_sec}s (attempt {retry + 1}/{max_retries}) [AI Tutor Key: {masked}]...")
                            time.sleep(wait_sec)
                            continue
                        else:
                            print(f"[Gemini Chat / AI Tutor] {attempt_model} rate-limit retries exhausted. Cascading to next candidate...")
                            break
                    elif resp.status_code == 404:
                        last_error = f"Gemini Error 404: {attempt_model} not found on v1beta API"
                        print(f"[Gemini Chat / AI Tutor] {attempt_model} returned 404 Not Found. Skipping...")
                        break
                    else:
                        last_error = f"Gemini API Error {resp.status_code} ({attempt_model}): {resp.text}"
                        print(f"[Gemini Chat / AI Tutor] {attempt_model} returned {resp.status_code}. Cascading to next candidate...")
                        break
                except Exception as exc:
                    last_error = f"{type(exc).__name__}: {exc}"
                    time.sleep(1.5)
                    continue

        return {"success": False, "error": last_error}

# ─────────────────────────────────────────────────────────────────────────────
# 9. ANKI EXPORTER (Native Image Occlusion & .apkg Export)
# ─────────────────────────────────────────────────────────────────────────────
class AnkiExporter:
    @staticmethod
    def _stable_id(string: str) -> int:
        return int(hashlib.sha256(string.encode("utf-8")).hexdigest()[:8], 16)

    @classmethod
    def _build_basic_model(cls, model_id: int) -> genanki.Model:
        return genanki.Model(
            model_id,
            "1. Basic (MBBS Suite)",
            fields=[{"name": "Front"}, {"name": "Back"}],
            templates=[{
                "name": "Card 1",
                "qfmt": '<div class="card front-card"><div class="badge-tag">MBBS Flashcard</div><div class="card-question">{{Front}}</div></div>',
                "afmt": '{{FrontSide}}<hr id="answer"><div class="card back-card"><div class="card-answer">{{Back}}</div></div>',
            }],
            css="""
            .card { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; font-size: 16px; color: #1e293b; background: #faf9f6; padding: 24px; border-radius: 16px; }
            .badge-tag { font-size: 11px; font-weight: 700; text-transform: uppercase; color: #1d4ed8; letter-spacing: 0.05em; margin-bottom: 8px; }
            .card-question { font-size: 18px; font-weight: 600; line-height: 1.4; color: #0f172a; }
            .card-answer { font-size: 16px; line-height: 1.6; color: #334155; margin-top: 12px; }
            """
        )

    @classmethod
    def _build_cloze_model(cls, model_id: int) -> genanki.Model:
        return genanki.Model(
            model_id,
            "Cloze++ (MBBS Suite)",
            fields=[{"name": "Text"}, {"name": "Back Extra"}],
            templates=[{
                "name": "Cloze 1",
                "qfmt": '<div class="card front-card"><div class="badge-tag">MBBS Cloze Assertion</div><div class="card-question">{{cloze:Text}}</div></div>',
                "afmt": '<div class="card back-card"><div class="badge-tag">MBBS Cloze Assertion</div><div class="card-question">{{cloze:Text}}</div><hr id="answer"><div class="card-answer">{{Back Extra}}</div></div>',
            }],
            css="""
            .card { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; font-size: 16px; color: #1e293b; background: #faf9f6; padding: 24px; border-radius: 16px; }
            .badge-tag { font-size: 11px; font-weight: 700; text-transform: uppercase; color: #0f766e; letter-spacing: 0.05em; margin-bottom: 8px; }
            .card-question { font-size: 18px; font-weight: 600; line-height: 1.4; color: #0f172a; }
            .cloze { font-weight: 700; color: #1d4ed8; text-decoration: underline; }
            .card-answer { font-size: 15px; color: #475569; margin-top: 12px; font-style: italic; }
            """,
            model_type=genanki.Model.CLOZE
        )

    @classmethod
    def _build_native_io_model(cls, model_id: int) -> genanki.Model:
        return genanki.Model(
            model_id,
            "Image Occlusion++++++ (Native MBBS Suite)",
            fields=[
                {"name": "Occlusion"},
                {"name": "Image"},
                {"name": "Header"},
                {"name": "Back Extra"},
                {"name": "Comments"},
            ],
            templates=[{
                "name": "Image Occlusion",
                "qfmt": '<div class="io-header">{{Header}}</div><div class="io-wrapper">{{Image}}{{Occlusion}}</div>',
                "afmt": '<div class="io-header">{{Header}}</div><div class="io-wrapper">{{Image}}{{Occlusion}}</div><div class="io-extra">{{Back Extra}}</div>',
            }],
            css="""
            .io-wrapper { position: relative; display: inline-block; max-width: 100%; border-radius: 12px; overflow: hidden; }
            .io-header { font-size: 16px; font-weight: bold; margin-bottom: 8px; color: #0f172a; }
            .io-extra { font-size: 14px; color: #475569; margin-top: 12px; font-style: italic; }
            """,
            model_type=genanki.Model.CLOZE
        )

    @classmethod
    def compile_occlusion_cards(
        cls,
        base_img: Image.Image,
        ocr_boxes: list[dict],
        mode: str = "hide_all",
        lecture_name: str = "",
        slide_num: int = 1,
        header: str = "",
        notes: str = "",
    ) -> list[dict]:
        targets = [b for b in ocr_boxes if b.get("checked", True) and not b.get("is_redaction", False)]
        redactions = [b for b in ocr_boxes if b.get("is_redaction", False)]
        if not targets and not redactions:
            return []

        # If permanent redactions exist, black them out permanently on a copy of base_img
        # so spoilers/answers/clutter are never visible during review on any card
        if redactions:
            base_img = base_img.copy()
            draw = ImageDraw.Draw(base_img)
            for r in redactions:
                draw.rectangle(r["box"], fill="black")

        w, h = base_img.size
        for idx, item in enumerate(targets):
            item["cloze_num"] = 1 if mode == "hide_all" else (idx + 1)

        cloze_parts = []
        for idx, item in enumerate(targets):
            bx = item["box"]
            left = max(0.0, min(1.0, bx[0] / w))
            top = max(0.0, min(1.0, bx[1] / h))
            width = max(0.001, min(1.0, (bx[2] - bx[0]) / w))
            height = max(0.001, min(1.0, (bx[3] - bx[1]) / h))

            cloze_idx = item.get("cloze_num", idx + 1)
            oi = idx + 1
            cloze_parts.append(
                f"{{{{c{cloze_idx}::image-occlusion:rect:left={left:.4f}:top={top:.4f}:width={width:.4f}:height={height:.4f}:oi={oi}}}}}"
            )

        occlusion_field = "<br>".join(cloze_parts)
        card_count = len(set(item.get("cloze_num", idx + 1) for idx, item in enumerate(targets)))

        back_extra = ""
        if notes and not notes.startswith("Load a slide"):
            clean_notes = notes.strip().replace("\n", "<br>")[:350]
            back_extra = f"<div style='font-size:13px; color:#334155; padding:6px; font-style:italic;'>{clean_notes}</div>"

        return [{
            "type": "native_io",
            "base_image": base_img,
            "occlusion_field": occlusion_field,
            "header": header or "",
            "back_extra": back_extra,
            "card_count": card_count,
            "targets": [t["text"] for t in targets],
            "slide": slide_num,
            "lecture": lecture_name,
        }]

    @classmethod
    def export_full_deck(
        cls,
        deck_name: str,
        normal_cards: list[dict],
        occlusion_cards: list[dict],
        export_dir: Path,
    ) -> Path:
        deck_id   = cls._stable_id(deck_name)
        deck      = genanki.Deck(deck_id, deck_name)

        model_basic = cls._build_basic_model(cls._stable_id(deck_name + "::basic"))
        model_cloze = cls._build_cloze_model(cls._stable_id(deck_name + "::cloze"))
        model_io    = cls._build_native_io_model(cls._stable_id(deck_name + "::native_io"))

        media_files: list[str] = []
        uid = uuid.uuid4().hex[:8]

        for card in normal_cards:
            ctype = card.get("type", "Cloze++")
            front = re.sub(r"\s*\[Lecture:[^\]]*\]", "", card.get("front", "") or card.get("text", "")).strip()
            back = re.sub(r"\s*\[Lecture:[^\]]*\]", "", card.get("back", "") or card.get("extra", "")).strip()
            f_low = front.lower()
            if any(k in f_low for k in [
                "why study", "why do we study", "benefit from the study",
                "disciplines that benefit", "reasons for studying",
                "learning objective", "learning outcome", "module code",
                "who won", "nobel prize"
            ]):
                continue
            if ctype in ("Cloze++", "Cloze"):
                note = genanki.Note(model=model_cloze, fields=[front, back])
            else:
                note = genanki.Note(model=model_basic, fields=[front, back])
            deck.add_note(note)

        for idx, occ in enumerate(occlusion_cards):
            img_name = f"_io_{uid}_{idx:03d}.png"
            img_path = export_dir / img_name
            occ["base_image"].save(str(img_path), "PNG")
            media_files.append(str(img_path))

            note = genanki.Note(
                model=model_io,
                fields=[
                    occ["occlusion_field"],
                    f'<img src="{img_name}">',
                    occ.get("header", ""),
                    occ.get("back_extra", ""),
                    "",
                ],
            )
            deck.add_note(note)

        safe_deck = deck_name.replace("::", "_").replace(" ", "_")
        apkg_path = export_dir / f"{safe_deck}.apkg"
        package = genanki.Package(deck)
        package.media_files = media_files
        package.write_to_file(str(apkg_path))
        return apkg_path

# ─────────────────────────────────────────────────────────────────────────────
# 9.5 KCL TIMETABLE & DAILY FACT ENGINES
# ─────────────────────────────────────────────────────────────────────────────
class KCLTimetableManager:
    """Manages King's College London MyTimetable (.ics) parsing, caching, and querying."""
    def __init__(self, storage_path: Path):
        self.storage_path = storage_path

    @staticmethod
    def _unfold_ics_lines(raw_text: str) -> list[str]:
        lines = []
        for line in raw_text.splitlines():
            line_str = line.rstrip("\r\n")
            if not line_str:
                continue
            if line_str.startswith((" ", "\t")) and lines:
                lines[-1] += line_str[1:]
            else:
                lines.append(line_str)
        return lines

    @staticmethod
    def _parse_ics_datetime(dt_str: str) -> Optional[datetime]:
        """Parses iCalendar datetime string: YYYYMMDDTHHMMSSZ or YYYYMMDDTHHMMSS or YYYYMMDD."""
        if not dt_str:
            return None
        dt_clean = re.sub(r"^.*?:\s*", "", dt_str).strip()
        for fmt in ("%Y%m%dT%H%M%SZ", "%Y%m%dT%H%M%S", "%Y%m%d"):
            try:
                clean_target = dt_clean[:16] if "T" in dt_clean else dt_clean[:8]
                if clean_target.endswith("Z"):
                    return datetime.strptime(clean_target, "%Y%m%dT%H%M%SZ")
                elif "T" in clean_target:
                    return datetime.strptime(clean_target[:15], "%Y%m%dT%H%M%S")
                else:
                    return datetime.strptime(clean_target[:8], "%Y%m%d")
            except Exception:
                continue
        return None

    KCL_MODULE_MAP: dict[str, tuple[str, str]] = {
        "4MBBS101": ("FBS", "Foundations of Biomedical Sciences"),
        "4MBBS102": ("PAS", "Physiology and Anatomy of Systems"),
        "4MBBS103": ("CS", "Clinical Skills"),
        "4MBBS104": ("IVBCP", "Introduction to Values-Based Clinical Practice"),
        "4MBBS105": ("FBS", "Foundations of Biomedical Sciences (Sem 1)"),
        "4MBBS106": ("PAS", "Physiology and Anatomy of Systems (Sem 2)"),
        "4MBBS201": ("FBS2", "Foundations of Biomedical Sciences 2"),
        "4MBBS202": ("PAS2", "Physiology and Anatomy of Systems 2"),
    }

    @classmethod
    def clean_kcl_event(cls, ev: dict) -> dict:
        raw_summary = (ev.get("raw_summary") or ev.get("summary") or "").strip()
        raw_desc = (ev.get("raw_description") or ev.get("description") or "").strip()
        loc = (ev.get("location") or "").strip()

        # Parse Celcat MyTimetable summary pattern: e.g. "4MBBS102 PHYSIOLOGY AND SY 000001/Lecture40/01_JT"
        m = re.match(r"^([0-9A-Z]{7,8})\s+(.*?)\s+\d{6}/([A-Za-z0-9_\s\-]+)/([A-Za-z0-9_]+)$", raw_summary)
        module_code = ""
        module_short = "MBBS"
        module_name = ""
        activity_raw = ""
        group = ""

        if m:
            module_code = m.group(1).upper()
            raw_mod_name = m.group(2).strip()
            activity_raw = m.group(3).strip()
            group = m.group(4).strip()
            if module_code in cls.KCL_MODULE_MAP:
                module_short, module_name = cls.KCL_MODULE_MAP[module_code]
            else:
                module_short = module_code
                module_name = raw_mod_name
        else:
            code_m = re.search(r"\b(4MBBS\d{3}|[A-Z]{3,4}\d{2,3}[A-Za-z]?)\b", raw_summary)
            if code_m:
                module_code = code_m.group(1).upper()
                if module_code in cls.KCL_MODULE_MAP:
                    module_short, module_name = cls.KCL_MODULE_MAP[module_code]
            activity_raw = raw_summary

        act_lower = activity_raw.lower()
        activity_type = "Lecture"
        activity_title = activity_raw
        act_num = ""

        if "clinical" in act_lower:
            activity_type = "Clinical Skills"
            num_m = re.search(r"clinical\s*(\d+)", act_lower)
            if num_m:
                act_num = num_m.group(1)
                activity_title = f"Clinical Skills {act_num}"
            else:
                activity_title = "Clinical Skills"
        elif "tut" in act_lower:
            activity_type = "Tutorial"
            num_m = re.search(r"tut(?:orial)?\s*(\d+)", act_lower)
            if num_m:
                act_num = num_m.group(1)
                activity_title = f"Tutorial {act_num}"
            else:
                activity_title = "Tutorial"
        elif "study" in act_lower or "self" in act_lower:
            activity_type = "Self-Study"
            num_m = re.search(r"(?:study|self\s*study)\s*(\d+)", act_lower)
            if num_m:
                act_num = num_m.group(1)
                activity_title = f"Directed Self-Study {act_num}"
            else:
                activity_title = "Directed Self-Study"
        elif "dissection" in act_lower or "anatomy lab" in act_lower:
            activity_type = "Dissection"
            num_m = re.search(r"(?:dissection|lab)\s*(\d+)", act_lower)
            if num_m:
                act_num = num_m.group(1)
                activity_title = f"Dissection Lab {act_num}"
            else:
                activity_title = "Dissection Lab"
        elif "workshop" in act_lower:
            activity_type = "Workshop"
            num_m = re.search(r"workshop\s*(\d+)", act_lower)
            if num_m:
                act_num = num_m.group(1)
                activity_title = f"Workshop {act_num}"
            else:
                activity_title = "Workshop"
        elif "lecture" in act_lower:
            activity_type = "Lecture"
            num_m = re.search(r"lecture\s*(\d+)", act_lower)
            if num_m:
                act_num = num_m.group(1)
                activity_title = f"Lecture {act_num}"
            else:
                activity_title = "Lecture"

        # Celcat metadata parsing: unpack Location, Date, Group(s), Additional Information, Staff
        desc_single = re.sub(r"[\r\n]+", " ", raw_desc).strip()

        # Check if Celcat description embeds Event type (e.g. "Event type: Clinical Skill Description: ...")
        ev_type_m = re.search(r"^\s*Event\s*type\s*:\s*([^:]+?)\s*Description\s*:", desc_single, re.IGNORECASE)
        if ev_type_m:
            celcat_ev_type = ev_type_m.group(1).strip()
            if "clinical" in celcat_ev_type.lower():
                activity_type = "Clinical Skills"
                if not act_num:
                    activity_title = "Clinical Skills"
            elif "tut" in celcat_ev_type.lower():
                activity_type = "Tutorial"
                if not act_num:
                    activity_title = "Tutorial"
            elif "workshop" in celcat_ev_type.lower():
                activity_type = "Workshop"
                if not act_num:
                    activity_title = "Workshop"
            elif "lecture" in celcat_ev_type.lower():
                activity_type = "Lecture"
                if not act_num:
                    activity_title = "Lecture"

        desc_parts = re.split(r"\s*(?:Location|Date|Group\(s\)|Groups?|Additional Information|Staff|Lecturer):\s*", desc_single, maxsplit=1, flags=re.IGNORECASE)
        pure_topic = desc_parts[0].strip() if desc_parts else ""

        # Strip all Celcat boilerplate prefixes ("Event type: ... Description:", "Description:", "Event type:", etc.)
        pure_topic = re.sub(r"^\s*Event\s*type\s*:[^:]*Description\s*:\s*", "", pure_topic, flags=re.IGNORECASE).strip()
        pure_topic = re.sub(r"^\s*(?:Event\s*type|Title|Topic|Subject|Description)\s*:\s*", "", pure_topic, flags=re.IGNORECASE).strip()
        pure_topic = re.sub(r"^\s*Description\s*:\s*", "", pure_topic, flags=re.IGNORECASE).strip()

        loc_m = re.search(r"Location:\s*(.*?)(?=\s+(?:Date|Group\(s\)|Groups?|Additional Information|Staff|Lecturer):|$)", desc_single, re.IGNORECASE)
        if loc_m:
            parsed_loc = loc_m.group(1).strip()
            if parsed_loc and (not loc or loc == "Guy's Campus"):
                loc = parsed_loc

        staff_m = re.search(r"(?:Staff|Lecturer):\s*(.*?)(?=\s+(?:Location|Date|Group\(s\)|Groups?|Additional Information):|$)", desc_single, re.IGNORECASE)
        staff = staff_m.group(1).strip() if staff_m else ""

        grp_m = re.search(r"Group\(s\):\s*(.*?)(?=\s+(?:Location|Date|Additional Information|Staff|Lecturer):|$)", desc_single, re.IGNORECASE)
        if grp_m:
            group = grp_m.group(1).strip()

        add_m = re.search(r"Additional Information:\s*(.*?)(?=\s+(?:Location|Date|Group\(s\)|Groups?|Staff|Lecturer):|$)", desc_single, re.IGNORECASE)
        additional_info = add_m.group(1).strip() if add_m else ""

        # Specific Lecture Code extraction (e.g. ACRS04L, MCB07L, FAH01L) - avoid 4MBBS module codes
        lec_code = ""
        c_m = re.search(r"\b([A-Z]{3,5}[-_]?\d{1,3}[A-Za-z]?)\b", pure_topic or desc_single)
        if c_m:
            cand = c_m.group(1).upper().replace("-", "")
            if not cand.startswith("4MBBS"):
                lec_code = cand

        # The clean title for the card must strictly be the lecture topic name
        if pure_topic:
            clean_title = pure_topic
        elif activity_title:
            clean_title = f"{module_short} · {activity_title}"
        else:
            clean_title = raw_summary or "KCL Session"

        topic_title = pure_topic

        # Duration Calculation
        st = ev.get("start_time", "")
        et = ev.get("end_time", "")
        duration_str = ev.get("duration_str") or ""
        if not duration_str and st and et:
            try:
                t1 = datetime.strptime(st, "%H:%M")
                t2 = datetime.strptime(et, "%H:%M")
                mins = int((t2 - t1).total_seconds() / 60)
                if mins >= 60:
                    hrs = mins / 60.0
                    duration_str = f"{int(hrs)} hr" if hrs.is_integer() else f"{hrs:.1f} hrs"
                else:
                    duration_str = f"{mins} mins"
            except Exception:
                pass

        return {
            **ev,
            "code": lec_code or module_code,
            "lec_code": lec_code,
            "type": activity_type,
            "clean_title": clean_title,
            "module_code": module_code,
            "module_short": module_short,
            "module_name": module_name,
            "activity_type": activity_type,
            "activity_title": activity_title,
            "activity_num": act_num,
            "group": group,
            "staff": staff,
            "additional_info": additional_info,
            "topic_title": topic_title,
            "duration_str": duration_str,
            "location": loc or "Guy's Campus",
            "raw_summary": raw_summary,
            "raw_description": raw_desc
        }

    def parse_ics_content(self, ics_text: str, filename: str = "") -> list[dict]:
        lines = self._unfold_ics_lines(ics_text)
        events = []
        curr = None

        for line in lines:
            line_s = line.strip()
            if line_s == "BEGIN:VEVENT":
                curr = {}
            elif line_s == "END:VEVENT":
                if curr and curr.get("summary"):
                    events.append(curr)
                curr = None
            elif curr is not None:
                if ":" in line:
                    k_part, val = line.split(":", 1)
                    k = k_part.split(";")[0].strip().upper()
                    clean_val = val.replace("\\n", "\n").replace("\\;", ";").replace("\\,", ",").strip()
                    if k == "SUMMARY":
                        curr["summary"] = clean_val
                    elif k == "LOCATION":
                        curr["location"] = clean_val
                    elif k == "DESCRIPTION":
                        curr["description"] = clean_val
                    elif k == "DTSTART":
                        dt = self._parse_ics_datetime(val)
                        if dt:
                            curr["dtstart"] = dt.isoformat()
                            curr["date"] = dt.strftime("%Y-%m-%d")
                            curr["start_time"] = dt.strftime("%H:%M")
                    elif k == "DTEND":
                        dt = self._parse_ics_datetime(val)
                        if dt:
                            curr["dtend"] = dt.isoformat()
                            curr["end_time"] = dt.strftime("%H:%M")
                    elif k == "UID":
                        curr["uid"] = clean_val

        processed_events = []
        for ev in events:
            summary = ev.get("summary", "")
            code_match = re.search(r"\b([A-Z]{2,6}[-_]?\d{1,3}[A-Za-z]?)\b", summary)
            ev["code"] = code_match.group(1).replace("-", "") if code_match else ""
            cleaned = self.clean_kcl_event(ev)
            processed_events.append(cleaned)

        processed_events.sort(key=lambda x: x.get("dtstart", ""))
        self.save_timetable(processed_events, source_name=filename or "MyTimetable")
        return processed_events

    def save_timetable(self, events: list[dict], source_name: str = "", source_url: str = ""):
        try:
            self.storage_path.parent.mkdir(parents=True, exist_ok=True)
            payload = {
                "last_updated": datetime.now().isoformat(),
                "source_name": source_name,
                "source_url": source_url,
                "events": events
            }
            with open(self.storage_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=2)
        except Exception as e:
            print(f"[KCLTimetableManager] Error saving timetable: {e}")

    def sync_from_url(self, calendar_url: str) -> dict:
        url = (calendar_url or "").strip()
        if not url:
            raise ValueError("No calendar URL provided.")

        # Convert webcal:// or webcals:// to standard https://
        if url.lower().startswith("webcal://"):
            url = "https://" + url[9:]
        elif url.lower().startswith("webcals://"):
            url = "https://" + url[10:]
        elif not url.lower().startswith(("http://", "https://")):
            url = "https://" + url

        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Synapse/2.0"
        }
        res = requests.get(url, headers=headers, timeout=20)
        if res.status_code != 200:
            raise ValueError(f"Failed to fetch calendar feed (HTTP {res.status_code})")

        ics_text = res.text
        if "BEGIN:VCALENDAR" not in ics_text and "BEGIN:VEVENT" not in ics_text:
            raise ValueError("The URL responded, but does not contain valid iCalendar (.ics) data.")

        events = self.parse_ics_content(ics_text, filename="KCL Live Feed")
        self.save_timetable(events, source_name="KCL Live Feed", source_url=url)
        return {
            "success": True,
            "url": url,
            "total_events": len(events),
            "last_synced": datetime.now().isoformat()
        }

    def load_timetable(self) -> dict:
        if not self.storage_path.exists():
            return {"last_updated": None, "source_name": None, "events": []}
        try:
            with open(self.storage_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "events" in data and isinstance(data["events"], list):
                    data["events"] = [self.clean_kcl_event(e) for e in data["events"]]
                return data
        except Exception as e:
            print(f"[KCLTimetableManager] Error loading timetable: {e}")
            return {"last_updated": None, "source_name": None, "events": []}

    def get_sessions_for_view(self, view_mode: str = "today", target_date: str = "") -> dict:
        data = self.load_timetable()
        all_events = data.get("events", [])
        now = datetime.now()
        today_str = now.strftime("%Y-%m-%d")
        tomorrow_str = (now + timedelta(days=1)).strftime("%Y-%m-%d")

        week_dates = [(now + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(7)]
        week_grouped = []

        if view_mode == "date" and target_date:
            filtered = [e for e in all_events if e.get("date") == target_date]
            try:
                dt_obj = datetime.strptime(target_date, "%Y-%m-%d")
                label = dt_obj.strftime("%A, %d %B %Y")
            except Exception:
                label = f"Schedule for {target_date}"
        elif view_mode == "tomorrow":
            filtered = [e for e in all_events if e.get("date") == tomorrow_str]
            label = "Tomorrow's Schedule"
        elif view_mode == "week":
            filtered = [e for e in all_events if e.get("date") in set(week_dates)]
            label = "This Week's Schedule"
            for d in week_dates:
                day_sessions = [e for e in all_events if e.get("date") == d]
                try:
                    d_dt = datetime.strptime(d, "%Y-%m-%d")
                    day_name = d_dt.strftime("%A, %d %b")
                except Exception:
                    day_name = d
                week_grouped.append({
                    "date": d,
                    "date_formatted": day_name,
                    "is_today": (d == today_str),
                    "sessions": day_sessions
                })
        else: # "today"
            filtered = [e for e in all_events if e.get("date") == today_str]
            label = "Today's Schedule"

        distinct_dates = sorted(list({e.get("date") for e in all_events if e.get("date")}))

        return {
            "view_mode": view_mode,
            "label": label,
            "current_time": now.strftime("%H:%M"),
            "today_date": today_str,
            "target_date": target_date or (tomorrow_str if view_mode == "tomorrow" else today_str),
            "available_dates": distinct_dates,
            "last_updated": data.get("last_updated"),
            "source_name": data.get("source_name"),
            "source_url": data.get("source_url", ""),
            "total_imported": len(all_events),
            "sessions": filtered,
            "week_grouped": week_grouped
        }


CURATED_PIGEON_FACTS = [
    {
        "topic": "Anatomy · Neurovascular",
        "title": "Recurrent Laryngeal Nerve Asymmetry",
        "fact": "The left recurrent laryngeal nerve loops under the aortic arch (lateral to ligamentum arteriosum), while the right recurrent laryngeal nerve loops under the right subclavian artery.",
        "clinical_pearl": "Left recurrent laryngeal nerve palsy often signals mediastinal pathology (e.g. bronchogenic carcinoma of the left lung apex, aortic aneurysm, mediastinal lymphadenopathy), presenting with hoarseness and a bovine cough.",
        "pigeon_quote": "Left loops low around the aorta! Protect those vocal cords and don't get caught out on your anatomy viva. 🐦🩺",
        "tags": ["Anatomy", "Head & Neck", "High-Yield"]
    },
    {
        "topic": "Physiology · Cardiovascular",
        "title": "Frank-Starling Law of the Heart",
        "fact": "The stroke volume of the heart increases in response to an increase in end-diastolic volume (preload) when all other factors remain constant, optimizing myofilament actin-myosin overlap.",
        "clinical_pearl": "In severe systolic heart failure, the myocardium becomes overstretched beyond the optimal sarcomere length (2.2 µm), causing the curve to plateau or drop off, requiring inotropic support.",
        "pigeon_quote": "More stretch equals more snap! Until the heart fails, then it's pure congestion. 🐦🩺",
        "tags": ["Physiology", "Cardiovascular", "Preload"]
    },
    {
        "topic": "Biochemistry · Metabolism",
        "title": "Rate-Limiting Step of Glycolysis",
        "fact": "Phosphofructokinase-1 (PFK-1) irreversibly converts Fructose-6-Phosphate to Fructose-1,6-Bisphosphate and serves as the primary rate-limiting and committed step of glycolysis.",
        "clinical_pearl": "PFK-1 is allosterically inhibited by ATP and citrate (signals of cellular high energy), and strongly activated by AMP and Fructose-2,6-Bisphosphate (stimulated by insulin).",
        "pigeon_quote": "PFK-1 is the gatekeeper of your ATP budget. Respect the phosphates! 🐦🩺",
        "tags": ["Biochemistry", "Glycolysis", "Enzymes"]
    },
    {
        "topic": "Anatomy · Thorax",
        "title": "Apertures of the Thoracic Diaphragm",
        "fact": "The thoracic diaphragm has three major openings at vertebral levels T8, T10, and T12: IVC & right phrenic nerve (T8), Oesophagus & vagus nerves (T10), and Aorta, thoracic duct & azygos vein (T12).",
        "clinical_pearl": "Mnemonic: 'I 8 10 Eggs At 12' — IVC (8), Oesophagus (10), Aorta (12). Remember the phrenic nerve pierces the central tendon alongside the IVC.",
        "pigeon_quote": "I ate 10 eggs at 12! Never forget your vertebral levels on exam day. 🐦🩺",
        "tags": ["Anatomy", "Thorax", "Mnemonics"]
    },
    {
        "topic": "Pharmacology · Autonomics",
        "title": "Adrenergic Receptor G-Protein Coupling",
        "fact": "Adrenergic receptors couple to distinct G-proteins: Alpha-1 couples to Gq (activates PLC -> IP3/DAG -> increased intracellular Ca2+), Alpha-2 to Gi (inhibits adenylate cyclase -> decreased cAMP), and Beta-1, 2, 3 to Gs (stimulates adenylate cyclase -> increased cAMP).",
        "clinical_pearl": "Mnemonic: 'QISS' — Alpha-1 (q), Alpha-2 (i), Beta-1 (s), Beta-2 (s). Beta-1 agonists increase cardiac inotropy/chronotropy, while Beta-2 agonists induce smooth muscle bronchodilation.",
        "pigeon_quote": "QISS and tell! Master your G-proteins and autonomic pharmacology becomes a breeze. 🐦🩺",
        "tags": ["Pharmacology", "Autonomics", "G-Proteins"]
    },
    {
        "topic": "Anatomy · Pelvis & Perineum",
        "title": "Pudendal Nerve Pathway & Alcock's Canal",
        "fact": "The pudendal nerve (S2-S4) exits the pelvis via the greater sciatic foramen (inferior to piriformis), crosses the sacrospinous ligament / ischial spine, and re-enters via the lesser sciatic foramen into the pudendal (Alcock's) canal.",
        "clinical_pearl": "Pudendal nerve blocks for second-stage labor or perineal surgery palpate the ischial spine transvaginally as the landmark to infiltrate local anesthetic.",
        "pigeon_quote": "S2, 3, 4 keeps the pee and poo off the floor! Watch that ischial spine landmark. 🐦🩺",
        "tags": ["Anatomy", "Pelvis", "Nerve Roots"]
    },
    {
        "topic": "Physiology · Respiratory",
        "title": "Ventilation-Perfusion (V/Q) Mismatch",
        "fact": "Normal overall V/Q ratio is ~0.8. In the apex of the upright lung, V/Q is highest (~3.0, high PO2, low PCO2); at the base, perfusion outstrips ventilation so V/Q is lowest (~0.6).",
        "clinical_pearl": "Mycobacterium tuberculosis thrives in high oxygen tension and thus preferentially colonizes the lung apices. Pulmonary embolism creates a dead space (V/Q -> infinity).",
        "pigeon_quote": "Gravity pulls blood down, so the base is well-perfused, but TB loves the oxygen-rich apices! 🐦🩺",
        "tags": ["Physiology", "Respiratory", "V/Q"]
    }
]

# ─────────────────────────────────────────────────────────────────────────────
# 9B. IN-APP LECTURE NOTES MANAGER (MBBS MODULE & BLOCK TAXONOMY)
# ─────────────────────────────────────────────────────────────────────────────
class InAppNotesManager:
    """Manages local in-app lecture notes categorized by MBBS Year 1 Module and Block."""

    KCL_STRUCTURE: dict[str, dict] = {
        "FBS": {
            "name": "Foundations of Biomedical Sciences",
            "blocks": {
                "MCB": "Molecular and Cellular Biology",
                "NBSS": "Neural Basis of Systemic Structure",
                "NAM": "Nerve and Muscle",
                "DH": "Development and Homeostasis",
                "IAI": "Infection, Adaptation and Immunity",
                "EAB": "Elimination, Absorption and Balance",
                "GEN": "Genetics"
            }
        },
        "PAS": {
            "name": "Physiology and Anatomy of Systems",
            "blocks": {
                "FAH": "Fundamental Anatomy and Histology",
                "FPP": "Foundations of Pharmacological Principles",
                "ACRS": "Anatomy of Cardiovascular and Respiratory Systems",
                "PCRS": "Physiology of Cardiovascular and Respiratory Systems",
                "AGRS": "Anatomy of Gastrointestinal and Renal Systems",
                "PGRS": "Physiology of Gastrointestinal and Renal Systems",
                "AMS": "Anatomy of Musculoskeletal System",
                "Cranial Nerves": "Cranial Nerve Anatomy & Pathways"
            }
        },
        "IVBCP": {
            "name": "Introduction to Values-Based Clinical Practice",
            "blocks": {
                "Clinical Skills": "Clinical Skills & Patient Assessment",
                "Human Values": "Medical Ethics, Law & Professionalism"
            }
        },
        "General": {
            "name": "General Clinical & Science Notes",
            "blocks": {
                "General": "Other Revision Notes"
            }
        }
    }

    def __init__(self, base_dir: Path):
        self.base_dir = base_dir
        self.index_file = self.base_dir / "notes_index.json"

    def _ensure_storage(self):
        try:
            self.base_dir.mkdir(parents=True, exist_ok=True)
            for mod, data in self.KCL_STRUCTURE.items():
                for blk in data["blocks"].keys():
                    (self.base_dir / mod / blk.replace(" ", "_")).mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"[InAppNotesManager] _ensure_storage note: {e}")

    def _load_index(self) -> dict:
        try:
            if self.index_file.exists():
                return json.loads(self.index_file.read_text(encoding="utf-8"))
        except Exception as e:
            print(f"[InAppNotesManager] _load_index note: {e}")
        return {"notes": {}}

    def _save_index(self, data: dict):
        try:
            self.base_dir.mkdir(parents=True, exist_ok=True)
            self.index_file.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        except Exception as e:
            print(f"[InAppNotesManager] _save_index note: {e}")

    @classmethod
    def detect_module_and_block(cls, title: str = "", file_path: str = "") -> tuple[str, str]:
        combined = f"{file_path} {title}".upper()

        # 1. Path-based matching
        for mod, mod_data in cls.KCL_STRUCTURE.items():
            if f"/{mod}/" in file_path.upper() or f"\\{mod}\\" in file_path.upper():
                for blk in mod_data["blocks"].keys():
                    blk_clean = blk.replace(" ", "_")
                    if f"/{blk}/" in file_path.upper() or f"/{blk_clean}/" in file_path.upper():
                        return mod, blk
                # Check title for block keyword
                for blk in mod_data["blocks"].keys():
                    if re.search(rf"\b{blk}\b", combined):
                        return mod, blk
                return mod, list(mod_data["blocks"].keys())[0]

        # 2. Block prefix matching in title or file path
        if re.search(r"\b(ACRS\d*L?)\b", combined):
            return "PAS", "ACRS"
        if re.search(r"\b(PCRS\d*L?)\b", combined):
            return "PAS", "PCRS"
        if re.search(r"\b(FAH\d*L?)\b", combined):
            return "PAS", "FAH"
        if re.search(r"\b(FPP\d*L?|FPP[-_]?\d+)\b", combined):
            return "PAS", "FPP"
        if re.search(r"\b(AGRS\d*L?)\b", combined):
            return "PAS", "AGRS"
        if re.search(r"\b(PGRS\d*L?)\b", combined):
            return "PAS", "PGRS"
        if re.search(r"\b(AMS\d*L?)\b", combined):
            return "PAS", "AMS"
        if "CRANIAL" in combined:
            return "PAS", "Cranial Nerves"
        if re.search(r"\b(MCB\d*L?|MCB[-_]?\d+)\b", combined):
            return "FBS", "MCB"
        if re.search(r"\b(NBSS\d*L?)\b", combined):
            return "FBS", "NBSS"
        if re.search(r"\b(NAM\d*L?)\b", combined):
            return "FBS", "NAM"
        if re.search(r"\b(DH\d*L?)\b", combined):
            return "FBS", "DH"
        if re.search(r"\b(IAI\d*L?)\b", combined):
            return "FBS", "IAI"
        if re.search(r"\b(EAB\d*L?)\b", combined):
            return "FBS", "EAB"
        if re.search(r"\b(GEN\d*L?)\b", combined):
            return "FBS", "GEN"
        if "CLINICAL" in combined or "SKILLS" in combined or re.search(r"\bCS\d*\b", combined):
            return "IVBCP", "Clinical Skills"
        if "ETHICS" in combined or "LAW" in combined or re.search(r"\bHV\d*\b", combined) or "VALUES" in combined:
            return "IVBCP", "Human Values"
        if "4MBBS101" in combined or "4MBBS105" in combined:
            return "FBS", "MCB"
        if "4MBBS102" in combined or "4MBBS106" in combined:
            return "PAS", "FAH"
        if "4MBBS104" in combined or "4MBBS103" in combined:
            return "IVBCP", "Human Values"

        return "General", "General"

    def save_note(self, title: str, content: str, module: str = "", block: str = "", source_file: str = "", note_id: str = "") -> dict:
        self._ensure_storage()
        now_iso = datetime.now().isoformat()
        index_data = self._load_index()

        existing = None
        if note_id and "notes" in index_data and note_id in index_data["notes"]:
            existing = index_data["notes"][note_id]
            slug = note_id
            clean_title = (title or existing.get("title") or "Untitled Lecture Note").strip()
            module = module or existing.get("module") or "General"
            block = block or existing.get("block") or "General"
            source_file = source_file or existing.get("source_file", "")
            created_at = existing.get("created_at", now_iso)
        else:
            clean_title = (title or "Untitled Lecture Note").strip()
            # Clean title from any Celcat boilerplate
            clean_title = re.sub(r"^\s*Event\s*type\s*:[^:]*Description\s*:\s*", "", clean_title, flags=re.IGNORECASE).strip()
            clean_title = re.sub(r"^\s*(?:Event\s*type|Title|Topic|Subject|Description)\s*:\s*", "", clean_title, flags=re.IGNORECASE).strip()

            if not module or not block or module not in self.KCL_STRUCTURE:
                det_mod, det_blk = self.detect_module_and_block(clean_title, source_file)
                module = module or det_mod
                block = block or det_blk

            # Build stable note ID / slug
            slug = re.sub(r"[^a-zA-Z0-9_\-]+", "_", f"{module}_{block}_{clean_title}").lower().strip("_")
            if not slug:
                slug = f"note_{int(time.time())}"

            created_at = index_data.get("notes", {}).get(slug, {}).get("created_at", now_iso)

        clean_title = re.sub(r"^\s*Event\s*type\s*:[^:]*Description\s*:\s*", "", clean_title, flags=re.IGNORECASE).strip()
        clean_title = re.sub(r"^\s*(?:Event\s*type|Title|Topic|Subject|Description)\s*:\s*", "", clean_title, flags=re.IGNORECASE).strip()

        mod_info = self.KCL_STRUCTURE.get(module, self.KCL_STRUCTURE["General"])
        mod_name = mod_info["name"]
        blk_name = mod_info["blocks"].get(block, block)

        word_count = len(re.findall(r"\b\w+\b", content or ""))

        # Extract first 140 chars of text for preview snippet
        raw_text = re.sub(r"[#*_`\[\]\(\)\<\>]", " ", content or "").strip()
        snippet = " ".join(raw_text.split()[:25])
        if len(snippet) > 120:
            snippet = snippet[:117] + "..."

        if existing and existing.get("md_path") and Path(existing["md_path"]).parent.exists():
            md_file = Path(existing["md_path"])
            json_file = Path(existing.get("json_path", str(md_file).replace(".md", ".json")))
        else:
            blk_folder = self.base_dir / module / block.replace(" ", "_")
            blk_folder.mkdir(parents=True, exist_ok=True)
            md_file = blk_folder / f"{slug}.md"
            json_file = blk_folder / f"{slug}.json"

        note_meta = {
            "id": slug,
            "title": clean_title,
            "module": module,
            "module_name": mod_name,
            "block": block,
            "block_name": blk_name,
            "source_file": source_file,
            "created_at": created_at,
            "updated_at": now_iso,
            "word_count": word_count,
            "snippet": snippet,
            "search_text": raw_text[:4000].lower(),
            "md_path": str(md_file),
            "json_path": str(json_file)
        }

        # Write files
        md_file.write_text(content, encoding="utf-8")
        json_file.write_text(json.dumps({**note_meta, "content": content}, indent=2, ensure_ascii=False), encoding="utf-8")

        # Update index
        if "notes" not in index_data:
            index_data["notes"] = {}
        index_data["notes"][slug] = note_meta
        self._save_index(index_data)

        return {
            "success": True,
            "note_id": slug,
            "title": clean_title,
            "module": module,
            "module_name": mod_name,
            "block": block,
            "block_name": blk_name,
            "updated_at": now_iso,
            "word_count": word_count,
            "snippet": snippet
        }

    def get_diagrams_dir(self, slug: str, module: str = "", block: str = "", md_path: str = "") -> Path:
        """Returns the dedicated directory on disk storing cached slide diagrams for this note."""
        if md_path:
            p = Path(md_path)
            if p.parent.exists():
                d = p.parent / f"{slug}_diagrams"
                d.mkdir(parents=True, exist_ok=True)
                return d
        mod = module or "General"
        blk = (block or "General").replace(" ", "_")
        d = self.base_dir / mod / blk / f"{slug}_diagrams"
        d.mkdir(parents=True, exist_ok=True)
        return d

    def get_note(self, note_id: str) -> dict:
        index_data = self._load_index()
        meta = index_data.get("notes", {}).get(note_id)
        if not meta:
            # Fallback search by ID or title
            for k, v in index_data.get("notes", {}).items():
                if k == note_id or v.get("title") == note_id:
                    meta = v
                    break

        if not meta:
            return {"success": False, "error": f"Note '{note_id}' not found in in-app library."}

        md_path = Path(meta.get("md_path", ""))
        content = ""
        if md_path.exists():
            try:
                content = md_path.read_text(encoding="utf-8")
            except Exception as e:
                return {"success": False, "error": f"Error reading note content: {e}"}
        else:
            json_path = Path(meta.get("json_path", ""))
            if json_path.exists():
                try:
                    payload = json.loads(json_path.read_text(encoding="utf-8"))
                    content = payload.get("content", "")
                except Exception:
                    pass

        # Load persisted slide diagram images from note's diagrams directory
        diagrams: dict[int, str] = {}
        try:
            slug = meta.get("id", note_id)
            diag_dir = self.get_diagrams_dir(slug, meta.get("module", ""), meta.get("block", ""), meta.get("md_path", ""))
            if diag_dir.exists():
                for p in diag_dir.glob("slide_*.png"):
                    m = re.search(r"slide_(\d+)\.png", p.name)
                    if m:
                        try:
                            s_num = int(m.group(1))
                            b64 = base64.b64encode(p.read_bytes()).decode("ascii")
                            diagrams[s_num] = f"data:image/png;base64,{b64}"
                        except Exception as read_err:
                            print(f"[get_note] Error reading cached diagram {p}: {read_err}")
        except Exception as diag_err:
            print(f"[get_note] Error loading note diagrams: {diag_err}")

        return {
            "success": True,
            "note": {
                **meta,
                "content": content
            },
            "diagrams": diagrams
        }

    def list_notes(self, module: str = "", block: str = "", search: str = "") -> dict:
        index_data = self._load_index()
        all_notes = list(index_data.get("notes", {}).values())

        # Sort newest first
        all_notes.sort(key=lambda x: x.get("updated_at", ""), reverse=True)

        # Compute counts per module & block
        counts: dict[str, dict] = {}
        for mod, mdata in self.KCL_STRUCTURE.items():
            counts[mod] = {"total": 0, "blocks": {blk: 0 for blk in mdata["blocks"]}}

        for n in all_notes:
            m = n.get("module", "General")
            b = n.get("block", "General")
            if m in counts:
                counts[m]["total"] += 1
                if b in counts[m]["blocks"]:
                    counts[m]["blocks"][b] += 1
                else:
                    counts[m]["blocks"][b] = 1

        filtered = []
        search_lower = (search or "").strip().lower()
        for n in all_notes:
            if module and n.get("module") != module:
                continue
            if block and n.get("block") != block:
                continue
            if search_lower:
                title = n.get("title", "").lower()
                snippet = n.get("snippet", "").lower()
                blk = n.get("block", "").lower()
                mod = n.get("module", "").lower()
                stext = n.get("search_text", "").lower()
                if (search_lower not in title and search_lower not in snippet and 
                    search_lower not in blk and search_lower not in mod and 
                    search_lower not in stext):
                    continue
            filtered.append(n)

        return {
            "success": True,
            "total_count": len(all_notes),
            "filtered_count": len(filtered),
            "structure": self.KCL_STRUCTURE,
            "counts": counts,
            "notes": filtered
        }

    def delete_note(self, note_id: str) -> dict:
        index_data = self._load_index()
        if note_id in index_data.get("notes", {}):
            meta = index_data["notes"].pop(note_id)
            self._save_index(index_data)
            try:
                md_p = Path(meta.get("md_path", ""))
                md_p.unlink(missing_ok=True)
                Path(meta.get("json_path", "")).unlink(missing_ok=True)
                if md_p.parent.exists():
                    d_dir = md_p.parent / f"{note_id}_diagrams"
                    if d_dir.exists():
                        shutil.rmtree(d_dir, ignore_errors=True)
            except Exception:
                pass
            return {"success": True, "note_id": note_id, "message": f"Deleted note '{meta.get('title')}'"}
        return {"success": False, "error": f"Note '{note_id}' not found."}

# ─────────────────────────────────────────────────────────────────────────────
# 9.5 ANNOUNCEMENTS ENGINE (GitHub-Hosted JSON + Startup/Wake Refresh)
# ─────────────────────────────────────────────────────────────────────────────
ANNOUNCEMENTS_URL = "https://raw.githubusercontent.com/Synapse-Study/synapse/main/announcements.json"

_GLOBAL_ANNOUNCEMENTS_ENGINE = None
_WAKE_OBSERVER_INSTANCE = None

class AnnouncementsEngine:
    def __init__(self, config_manager: ConfigManager, webview_window=None):
        self._config = config_manager
        self._window = webview_window
        self._cached_announcements: list[dict] = []
        self._last_fetched: float = 0.0
        self._lock = threading.Lock()
        global _GLOBAL_ANNOUNCEMENTS_ENGINE
        _GLOBAL_ANNOUNCEMENTS_ENGINE = self

    def set_window(self, window):
        self._window = window

    def fetch_remote(self) -> list[dict]:
        """
        Fetches announcements.json from GitHub with a 5-second timeout
        and Cache-Control: no-cache header so offline launches never hang.
        """
        headers = {
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0",
            "User-Agent": "Synapse-App/1.0"
        }
        try:
            resp = requests.get(ANNOUNCEMENTS_URL, headers=headers, timeout=5)
            if resp.status_code == 200:
                data = resp.json()
                if isinstance(data, list):
                    with self._lock:
                        self._cached_announcements = data
                        self._last_fetched = time.time()
                    return data
        except Exception:
            # Offline launches or network timeout: fail silently without hanging
            pass
        return self._cached_announcements

    def get_active_announcements(self, force_refresh: bool = False) -> list[dict]:
        with self._lock:
            cached = list(self._cached_announcements)
            last = self._last_fetched

        # If cache is empty or older than 5 minutes (or force_refresh requested), fetch
        if force_refresh or not cached or (time.time() - last > 300):
            cached = self.fetch_remote()

        dismissed = set(self._config.get_dismissed_announcements())
        now_date = datetime.now().date()
        active = []

        for item in cached:
            if not isinstance(item, dict):
                continue
            ann_id = str(item.get("id", "")).strip()
            if not ann_id or ann_id in dismissed:
                continue
            if not item.get("active", False):
                continue
            expires_at = item.get("expires_at")
            if expires_at:
                try:
                    exp_date = datetime.fromisoformat(str(expires_at).strip().split("T")[0]).date()
                    if exp_date < now_date:
                        continue
                except Exception:
                    pass
            active.append(item)

        return active

    def fetch_async(self, callback_notify: bool = True):
        def _bg():
            self.get_active_announcements(force_refresh=True)
            if callback_notify and self._window:
                try:
                    js_code = "if (typeof loadAnnouncements === 'function') loadAnnouncements();"
                    self._window.evaluate_js(js_code)
                except Exception:
                    pass
        threading.Thread(target=_bg, daemon=True).start()

# ─────────────────────────────────────────────────────────────────────────────
# 10. PYWEBVIEW API BRIDGE CLASS
# ─────────────────────────────────────────────────────────────────────────────
class AppAPI:
    def __init__(self):
        self._window: Optional[webview.Window] = None
        self._config = ConfigManager()
        self.state   = AppState()

        # Initial Anki URL from config
        AnkiBridge.ANKI_URL = self._config.get("ANKI_URL", "http://127.0.0.1:8765")

        # Initial OCR confidence threshold
        try:
            ocr_thresh = float(self._config.get("OCR_CONFIDENCE_THRESHOLD", 0.25))
        except Exception:
            ocr_thresh = 0.25

        # Engines
        _ensure_app_data_dirs()
        self._render_engine   = SlideRenderEngine(on_ready_callback=self._on_slide_render_ready)
        self._ocr_engine      = OCREngine(on_ready_callback=self._on_ocr_ready, min_confidence=ocr_thresh)
        self._cards_engine    = GeminiCardsEngine(self._config)
        self._notes_engine    = GeminiNotesEngine(self.state, self._config, self._cards_engine)
        self._notion_engine   = NotionSyncEngine(self._config, self.state, self._render_engine)
        self._index_engine    = KCLLectureIndexEngine(LECTURES_CACHE_DB_PATH, self._config, on_update_callback=self._on_index_updated)
        self._onedrive_engine = KCLOneDriveEngine(self._config, self._index_engine)
        self._chatbot_engine  = GeminiChatbotEngine(self._config, self._onedrive_engine, self.state, self._index_engine)
        self._timetable_manager = KCLTimetableManager(APP_DATA_DIR / "kcl_timetable.json")
        self._in_app_notes_manager = InAppNotesManager(SAVED_NOTES_DIR)
        self._announcements_engine = AnnouncementsEngine(self._config)

        # Initial Target Deck from config
        self.state.target_deck = self._config.get("TARGET_DECK", _CONFIG_TEMPLATE["TARGET_DECK"])

        # Off-peak auto-prepare worker state
        self._last_user_activity: float = 0.0
        self._auto_prep_thread: Optional[threading.Thread] = None
        self._stop_requested: bool = False

    def set_window(self, window: webview.Window):
        self._window = window
        self._announcements_engine.set_window(window)
        self._announcements_engine.fetch_async(callback_notify=True)
        if not self._auto_prep_thread or not self._auto_prep_thread.is_alive():
            self._auto_prep_thread = threading.Thread(target=self._auto_prepare_loop, daemon=True)
            self._auto_prep_thread.start()

    def _notify_frontend(self, event_name: str, payload: dict):
        if self._window:
            js = f"window.dispatchEvent(new CustomEvent('{event_name}', {{ detail: {json.dumps(payload)} }}))"
            try:
                self._window.evaluate_js(js)
            except Exception:
                pass

    def _on_slide_render_ready(self, success: bool, message: str):
        if success and self.state.presentation:
            self._notify_frontend("slide-render-ready", {
                "success": True,
                "message": message,
                "slide_idx": self.state.current_slide_idx
            })
            self._notify_frontend("app-status", {"type": "success", "message": f"Rendered high-res slides: {message}"})
        elif not success:
            self._notify_frontend("app-status", {"type": "error", "message": f"Slide rendering note: {message}"})

    def is_slide_rendered(self, slide_idx: int = -1) -> bool:
        if getattr(self._render_engine, "is_rendering", False):
            return False
        idx = slide_idx if slide_idx >= 0 else self.state.current_slide_idx
        return self._render_engine.has_rendered_slide(idx)

    def _on_ocr_ready(self, success: bool, message: str):
        self.state.ocr_ready = success
        self._notify_frontend("ocr-status", {"ready": success, "message": message})

    def _on_index_updated(self, status: dict):
        self._notify_frontend("lecture-index-status", status)

    # ── Initialization ───────────────────────────────────────────────────────
    def init(self) -> dict:
        self.state.anki_connected = AnkiBridge.is_connected()
        self.state.available_decks = AnkiBridge.get_deck_names() if self.state.anki_connected else []
        # Start background index sync on startup
        threading.Thread(target=self._background_sync_index, daemon=True).start()
        # Start background update check on startup
        threading.Thread(target=self._background_check_updates, daemon=True).start()
        # Start background timetable sync if subscription URL configured
        if self._config.get("TIMETABLE_URL"):
            threading.Thread(target=self._background_sync_timetable, daemon=True).start()
        return {
            "version": CURRENT_VERSION,
            "target_deck": self.state.target_deck,
            "available_decks": self.state.available_decks,
            "anki_connected": self.state.anki_connected,
            "ocr_ready": self._ocr_engine.is_ready,
            "index_status": self._index_engine.get_status(),
            "config": self._config._data,
            "queue_summary": self.get_queue_summary()
        }

    def _background_sync_index(self):
        try:
            res = self._index_engine.sync_index(force=False)
            self._notify_frontend("lecture-index-status", res)
        except Exception as e:
            print(f"Background index sync note: {e}")

    def _background_check_updates(self):
        try:
            update_info = check_github_update()
            if update_info.get("update_available"):
                self._notify_frontend("app-update-available", update_info)
        except Exception as exc:
            print(f"Background update check note: {exc}")

    def _background_sync_timetable(self):
        try:
            url = self._config.get("TIMETABLE_URL", "").strip()
            if url:
                res = self._timetable_manager.sync_from_url(url)
                self._notify_frontend("timetable-synced", {
                    "success": True,
                    "view": self._timetable_manager.get_sessions_for_view("today")
                })
        except Exception as exc:
            print(f"Background timetable sync note: {exc}")

    def check_for_updates(self) -> dict:
        """Query GitHub Releases for the latest Synapse release."""
        return check_github_update()

    # ── Home Dashboard & KCL Timetable ───────────────────────────────────────
    def get_home_dashboard(self) -> dict:
        return {
            "timetable": self.get_timetable_sessions("today"),
            "daily_fact": self.get_daily_fledge_fact(force_refresh=False),
            "recent_lectures": self.get_recent_lectures(),
            "anki_stats": self.get_anki_study_stats(),
            "timetable_url": self._config.get("TIMETABLE_URL", "").strip(),
            "user_name": self._config.get("USER_NAME", "").strip()
        }

    def sync_timetable_url(self, url: str = "") -> dict:
        url_clean = (url or "").strip()
        if not url_clean:
            url_clean = self._config.get("TIMETABLE_URL", "").strip()
        if not url_clean:
            return {
                "success": False,
                "error": "No timetable subscription URL provided. Paste your link from mytimetable.kcl.ac.uk."
            }

        try:
            res = self._timetable_manager.sync_from_url(url_clean)
            self._config.set("TIMETABLE_URL", res["url"])
            return {
                "success": True,
                "url": res["url"],
                "total_events": res["total_events"],
                "view": self.get_timetable_sessions("today")
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def upload_timetable_dialog(self) -> dict:
        if not self._window:
            return {"success": False, "error": "Window not initialized"}

        result = self._window.create_file_dialog(
            webview.FileDialog.OPEN,
            directory=str(Path.home() / "Downloads"),
            allow_multiple=False,
            file_types=('iCalendar Timetable Files (*.ics)', 'All files (*.*)')
        )
        if not result or len(result) == 0:
            return {"cancelled": True}

        file_path = Path(result[0])
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
            events = self._timetable_manager.parse_ics_content(content, filename=file_path.name)
            return {
                "success": True,
                "filename": file_path.name,
                "total_events": len(events),
                "view": self.get_timetable_sessions("today")
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def import_timetable_content(self, ics_content: str, filename: str = "MyTimetable.ics") -> dict:
        try:
            events = self._timetable_manager.parse_ics_content(ics_content, filename=filename)
            return {
                "success": True,
                "filename": filename,
                "total_events": len(events),
                "view": self.get_timetable_sessions("today")
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def _is_off_peak_time(self) -> bool:
        now = datetime.now()
        # Off-peak: before 12:30 local time or after 20:00 local time
        return (now.hour < 12) or (now.hour == 12 and now.minute < 30) or (now.hour >= 20)

    def mark_user_activity(self):
        self._last_user_activity = time.time()

    def _is_user_busy(self) -> bool:
        if getattr(self, "_notes_generation_status", {}).get("active"):
            return True
        if time.time() - getattr(self, "_last_user_activity", 0.0) < 45.0:
            return True
        return False

    def _find_saved_note_for_session(self, session: dict) -> Optional[dict]:
        """Check if an In-App note has already been prepared for this timetable session."""
        try:
            index_data = self._in_app_notes_manager._load_index()
            saved_notes = index_data.get("notes", {})
            if not saved_notes:
                return None

            lec_code = (session.get("lec_code") or "").upper().strip()
            topic = (session.get("clean_title") or "").strip().lower()
            code = (session.get("code") or "").upper().strip()

            for note_id, note in saved_notes.items():
                title_low = (note.get("title") or "").strip().lower()
                source_file = (note.get("source_file") or "").upper()
                if lec_code and (lec_code in title_low.upper() or lec_code in source_file):
                    return {"id": note_id, "title": note.get("title", "")}
                if topic and len(topic) > 5 and topic in title_low:
                    return {"id": note_id, "title": note.get("title", "")}
                if code and code in title_low.upper():
                    return {"id": note_id, "title": note.get("title", "")}
        except Exception as e:
            print(f"[_find_saved_note_for_session] Error: {e}")
        return None

    def get_auto_prepare_status(self) -> dict:
        enabled = bool(self._config.get("AUTO_PREPARE_NOTES", True))
        is_off_peak = self._is_off_peak_time()
        label = "Auto-Prep: Active (Off-peak)" if (enabled and is_off_peak) else ("Auto-Prep: Scheduled (Off-peak)" if enabled else "Auto-Prep: Disabled")
        return {
            "enabled": enabled,
            "is_off_peak": is_off_peak,
            "label": label
        }

    def set_auto_prepare_notes(self, enabled: bool) -> dict:
        self._config.set("AUTO_PREPARE_NOTES", bool(enabled))
        status = self.get_auto_prepare_status()
        self._notify_frontend("auto-prepare-update", status)
        return {"success": True, "status": status}

    def _auto_prepare_loop(self):
        time.sleep(12)  # Delay after startup
        while not getattr(self, "_stop_requested", False):
            try:
                enabled = self._config.get("AUTO_PREPARE_NOTES", True)
                if enabled and self._is_off_peak_time():
                    if not self._is_user_busy():
                        self._run_auto_prepare_cycle()
            except Exception as e:
                print(f"[AutoPrepareNotes] Background worker note: {e}")
            time.sleep(300)  # Check every 5 minutes

    def _run_auto_prepare_cycle(self):
        today_data = self._timetable_manager.get_sessions_for_view("today")
        tomorrow_data = self._timetable_manager.get_sessions_for_view("tomorrow")
        all_sessions = (today_data.get("sessions") or []) + (tomorrow_data.get("sessions") or [])

        for s in all_sessions:
            if self._is_user_busy():
                print("[AutoPrepareNotes] User activity detected; yielding background pre-generation.")
                return
            if self._find_saved_note_for_session(s):
                continue

            lec_code = s.get("lec_code") or ""
            topic = s.get("clean_title") or s.get("raw_summary") or ""

            slide_info = self.find_slides_for_session(
                session_code=s.get("code", ""),
                session_summary=s.get("raw_summary", ""),
                module_code=s.get("module_code", ""),
                activity_title=s.get("activity_title", ""),
                topic_title=topic,
                lec_code=lec_code
            )

            if not (slide_info and slide_info.get("found")):
                continue

            f_path = Path(slide_info["file_path"])
            if not f_path.exists() or not f_path.suffix.lower().endswith(".pptx"):
                continue

            try:
                prs = Presentation(str(f_path))
                if len(prs.slides) <= 1:
                    continue
            except Exception:
                continue

            if self._is_user_busy():
                return

            lec_title = s.get("clean_title") or lec_code or f_path.stem
            print(f"[AutoPrepareNotes] Quietly pre-generating notes for '{lec_title}' ({f_path.name})...")

            try:
                def _bg_status_check(msg, model=None):
                    if self._is_user_busy():
                        raise InterruptedError("User activity detected; yielding background pre-generation.")

                notes, model_used = self._notes_engine.generate_lecture_notes(
                    prs,
                    lec_title,
                    status_callback=_bg_status_check
                )

                if self._is_user_busy():
                    return

                mod, blk = InAppNotesManager.detect_module_and_block(lec_title, str(f_path))
                save_res = self._in_app_notes_manager.save_note(
                    title=lec_title,
                    content=notes,
                    module=mod,
                    block=blk,
                    source_file=f_path.name
                )

                if save_res.get("success"):
                    note_id = save_res.get("note_id", "")
                    diag_dir = self._in_app_notes_manager.get_diagrams_dir(note_id, mod, blk)
                    slide_tags = re.findall(r"\(\s*slides?\s+([0-9\s,\-–and]+?):", notes, re.IGNORECASE)
                    all_s_nums = set()
                    for tag in slide_tags:
                        for num_match in re.findall(r"\d+", tag):
                            n = int(num_match)
                            if 1 < n <= len(prs.slides):
                                all_s_nums.add(n)
                    for s_num in all_s_nums:
                        try:
                            png_bytes = self._notion_engine.get_best_slide_image(prs, s_num, pptx_path=f_path)
                            if png_bytes and diag_dir:
                                (diag_dir / f"slide_{s_num}.png").write_bytes(png_bytes)
                        except Exception:
                            pass

                    print(f"[AutoPrepareNotes] Notes ready for '{lec_title}' saved to {mod} · {blk} ✓")
                    self._notify_frontend("timetable-update", {
                        "type": "note_ready",
                        "note_id": note_id,
                        "session_code": s.get("code")
                    })
                    time.sleep(15)  # Pace between lectures
            except InterruptedError:
                print(f"[AutoPrepareNotes] Yielded generation for '{lec_title}' due to user activity.")
                return
            except Exception as ge:
                print(f"[AutoPrepareNotes] Error pre-generating notes for '{lec_title}': {ge}")

    def get_timetable_sessions(self, view_mode: str = "today", target_date: str = "") -> dict:
        data = self._timetable_manager.get_sessions_for_view(view_mode, target_date=target_date)
        sessions = data.get("sessions", [])
        for s in sessions:
            saved = self._find_saved_note_for_session(s)
            if saved:
                s["has_saved_note"] = True
                s["saved_note_id"] = saved["id"]
                s["saved_note_title"] = saved["title"]
            else:
                s["has_saved_note"] = False
        for group in data.get("week_grouped", []):
            for s in group.get("sessions", []):
                saved = self._find_saved_note_for_session(s)
                if saved:
                    s["has_saved_note"] = True
                    s["saved_note_id"] = saved["id"]
                    s["saved_note_title"] = saved["title"]
                else:
                    s["has_saved_note"] = False
        data["auto_prepare_status"] = self.get_auto_prepare_status()
        return data

    def find_slides_for_session(self, session_code: str = "", session_summary: str = "", module_code: str = "", activity_title: str = "", topic_title: str = "", lec_code: str = "") -> dict:
        code_clean = (session_code or "").strip().upper()
        summary_clean = (session_summary or "").strip()
        act_clean = (activity_title or "").strip()
        topic_clean = (topic_title or "").strip()
        mod_clean = (module_code or "").strip().upper()
        target_lec = (lec_code or "").strip().upper()

        # If target_lec was not passed explicitly, extract it from session_code, topic_clean, or summary_clean
        if not target_lec:
            for cand in [code_clean, topic_clean, summary_clean]:
                m = re.search(r"\b([A-Z]{3,5}[-_]?\d{1,3}[A-Za-z]?)\b", cand)
                if m:
                    c_up = m.group(1).upper().replace("-", "")
                    if not c_up.startswith("4MBBS"):
                        target_lec = c_up
                        break

        # Candidate tokens for the specific lecture code (e.g. ACRS04L, ACRS04, ACRS 04, ACRS 4)
        lec_tokens = []
        if target_lec:
            lec_tokens.append(target_lec)
            m_code = re.match(r"^([A-Z]+)(\d+)([A-Z]?)$", target_lec)
            if m_code:
                prefix, num_str, suffix = m_code.groups()
                num_int = int(num_str)
                lec_tokens.extend([
                    f"{prefix}{num_str}",
                    f"{prefix} {num_str}",
                    f"{prefix}-{num_str}",
                    f"{prefix}{num_int}",
                    f"{prefix} {num_int}",
                    f"{prefix}-{num_int}",
                ])
                if suffix:
                    lec_tokens.extend([
                        f"{prefix}{num_str}{suffix}",
                        f"{prefix} {num_str}{suffix}",
                        f"{prefix}{num_int}{suffix}",
                    ])

        # Detect target module abbreviation (PAS, FBS, IVBCP)
        target_mod = ""
        if mod_clean in KCLTimetableManager.KCL_MODULE_MAP:
            target_mod = KCLTimetableManager.KCL_MODULE_MAP[mod_clean][0].upper()
        elif "PAS" in summary_clean.upper() or "PHYSIOLOGY" in summary_clean.upper():
            target_mod = "PAS"
        elif "FBS" in summary_clean.upper() or "BIOMEDICAL" in summary_clean.upper():
            target_mod = "FBS"
        elif "IVBCP" in summary_clean.upper() or "VALUES" in summary_clean.upper():
            target_mod = "IVBCP"

        # Topic keywords - filter out common noise words
        stop_words = {
            "and", "the", "for", "with", "without", "part", "lecture", "tutorial", "session",
            "campus", "theatre", "hall", "centre", "room", "kcl", "date", "group", "groups",
            "staff", "location", "additional", "information", "thursday", "friday", "monday",
            "tuesday", "wednesday", "november", "december", "january", "february", "march",
            "april", "may", "june", "july", "august", "september", "october", "system",
            "systems", "introduction", "basics", "case", "based", "discussion", "answers",
            "keats", "mbbs", "year", "clinical", "skills", "ethics", "values"
        }
        topic_words = [
            w.lower() for w in re.findall(r'[A-Za-z0-9]+', topic_clean)
            if len(w) > 3 and w.lower() not in stop_words and not re.match(r"^4mbbs", w.lower())
        ]

        display_name = topic_clean or target_lec or act_clean or "this session"

        def score_candidate(filename: str, row_mod: str, file_path_str: str) -> int:
            fn_up = filename.upper()
            fn_low = filename.lower()
            fp_up = file_path_str.upper()
            score = 0

            # 1. Primary match: Specific lecture code (e.g. ACRS04L, ACRS04, ACRS 04)
            has_lec_match = False
            for tok in lec_tokens:
                tok_up = tok.upper()
                if tok_up in fn_up:
                    score += 50
                    has_lec_match = True
                    break

            # 2. Module alignment bonus (PAS, FBS, IVBCP)
            if target_mod and (target_mod == row_mod or f"/{target_mod}/" in fp_up):
                score += 5

            # 3. Topic keyword matches
            matched_words = 0
            for tw in topic_words:
                if tw in fn_low:
                    score += 15
                    matched_words += 1

            # Strict gating:
            # - If a specific lecture code exists (e.g. ACRS04L), require matching the lecture code OR >= 2 specific topic words.
            # - Never allow generic module match or single generic word to trigger a match.
            if target_lec and not has_lec_match and matched_words < 2:
                return 0
            if not target_lec and matched_words < 1:
                return 0

            return score

        # 1. Search in SQLite index
        try:
            con = self._index_engine._get_connection()
            cur = con.cursor()
            cur.execute("SELECT filename, file_path, slide_count, module, block FROM lectures")
            rows = [dict(r) for r in cur.fetchall()]
            con.close()

            best_row = None
            best_score = 0
            for r in rows:
                score = score_candidate(r["filename"], (r.get("module") or "").upper(), r.get("file_path") or "")
                if score > best_score:
                    best_score = score
                    best_row = r

            if best_row and best_score >= 30:
                return {
                    "found": True,
                    "file_path": best_row["file_path"],
                    "filename": best_row["filename"],
                    "slide_count": best_row["slide_count"],
                    "module": best_row["module"],
                    "block": best_row["block"]
                }
        except Exception as e:
            print(f"[find_slides_for_session] DB search note: {e}")

        # 2. Fallback search in LECTURES_DIR and OneDrive directories
        candidate_dirs = [
            Path(self._config.get("LECTURES_DIR", "./Lectures_OneDrive")).expanduser(),
            Path.home() / "Library/CloudStorage/OneDrive-King'sCollegeLondon",
            Path.home() / "OneDrive - King's College London",
            Path.home() / "Downloads"
        ]

        best_fs_path = None
        best_fs_score = 0
        for base_dir in candidate_dirs:
            try:
                if not base_dir.exists():
                    continue
                for ext in ("*.pptx", "*.pdf"):
                    for p in base_dir.rglob(ext):
                        if p.name.startswith("~$"):
                            continue
                        score = score_candidate(p.name, target_mod, str(p))
                        if score > best_fs_score:
                            best_fs_score = score
                            best_fs_path = p
            except Exception as e:
                print(f"[find_slides_for_session] FS search note: {e}")

        if best_fs_path and best_fs_score >= 30:
            return {
                "found": True,
                "file_path": str(best_fs_path.resolve()),
                "filename": best_fs_path.name,
                "slide_count": 0,
                "module": target_mod,
                "block": ""
            }

        return {
            "found": False,
            "session_title": display_name,
            "message": f"No lecture slides found matching '{display_name}'.",
            "keats_note": "Note: Lecture slides must be downloaded from KEATS and placed into your connected OneDrive folder (~/OneDrive - King's College London) for Synapse to automatically locate and index them."
        }

    def get_daily_fledge_fact(self, force_refresh: bool = False) -> dict:
        return self.get_daily_pigeon_fact(force_refresh=force_refresh)

    def get_daily_pigeon_fact(self, force_refresh: bool = False) -> dict:
        cache_file = APP_DATA_DIR / "daily_fact_cache.json"
        today_str = datetime.now().strftime("%Y-%m-%d")

        if not force_refresh and cache_file.exists():
            try:
                with open(cache_file, "r", encoding="utf-8") as f:
                    cached = json.load(f)
                    if cached.get("date") == today_str and cached.get("fact_data"):
                        return {"success": True, "fact": cached["fact_data"], "date": today_str, "cached": True}
            except Exception:
                pass

        fact_obj = None
        if force_refresh:
            try:
                prompt = """You are Dr. Fledge, the mascot and expert medical tutor for MBBS Year 1 students at King's College London (Foundations of Biomedical Sciences / Physiology and Anatomy of Systems).
Generate one high-yield, memorable medical pearl or physiological/anatomical fact for today.
Topics should focus on high-yield MBBS Year 1 content:
- Anatomical relations (e.g. neurovascular pathways, fascial compartments, landmarks)
- Systemic physiology mechanisms (cardiovascular, respiratory, renal, GI)
- Foundational biochemistry rate-limiting enzymes or pharmacology receptors

Return a strictly valid JSON object matching this schema:
{
  "topic": "Anatomy · Neurovascular",
  "title": "Recurrent Laryngeal Nerve Asymmetry",
  "fact": "The left recurrent laryngeal nerve loops under the aortic arch, while the right loops under the right subclavian artery.",
  "clinical_pearl": "Left recurrent laryngeal palsy indicates mediastinal pathology, causing hoarseness and bovine cough.",
  "pigeon_quote": "Left loops low! Protect those vocal cords and lock in. 🐦🩺",
  "tags": ["Anatomy", "Neurovascular", "High-Yield"]
}
"""
                raw = self._cards_engine._call_gemini_rest(
                    prompt,
                    preferred_model="gemini-3.5-flash-lite",
                    task_label="Dr Fledge Daily Fact"
                )
                m = re.search(r"\{.*\}", raw, re.DOTALL)
                if m:
                    fact_obj = json.loads(m.group(0))
            except Exception as e:
                print(f"[DailyFact] Gemini generation note: {e}")

        if not fact_obj or not fact_obj.get("fact"):
            day_of_year = datetime.now().timetuple().tm_yday
            fact_obj = CURATED_PIGEON_FACTS[day_of_year % len(CURATED_PIGEON_FACTS)]

        try:
            cache_file.parent.mkdir(parents=True, exist_ok=True)
            with open(cache_file, "w", encoding="utf-8") as f:
                json.dump({"date": today_str, "fact_data": fact_obj}, f, indent=2)
        except Exception:
            pass

        return {"success": True, "fact": fact_obj, "date": today_str, "cached": not force_refresh}

    def get_recent_lectures(self) -> list[dict]:
        recent = self._config.get("RECENT_LECTURES", [])
        if not isinstance(recent, list):
            recent = []

        if len(recent) < 3:
            try:
                con = self._index_engine._get_connection()
                cur = con.cursor()
                cur.execute("SELECT filename, file_path, slide_count, block FROM lectures ORDER BY mtime DESC LIMIT 3")
                rows = cur.fetchall()
                con.close()
                existing_paths = {r.get("path") for r in recent}
                for r in rows:
                    if r["file_path"] not in existing_paths and len(recent) < 3:
                        recent.append({
                            "title": Path(r["filename"]).stem,
                            "filename": r["filename"],
                            "path": r["file_path"],
                            "slide_count": r["slide_count"],
                            "last_opened": "Available in OneDrive"
                        })
            except Exception:
                pass

        return recent[:3]

    def get_anki_study_stats(self) -> dict:
        connected = AnkiBridge.is_connected()
        if not connected:
            return {
                "connected": False,
                "due_today": 0,
                "reviewed_today": 0,
                "new_cards": 0,
                "total_cards": 0,
                "deck": self.state.target_deck or "MBBS Jassem A.",
                "message": "Anki desktop is offline"
            }

        target_deck = (self.state.target_deck or self._config.get("DEFAULT_PARENT_DECK", "MBBS Jassem A.") or "MBBS Jassem A.").strip()
        due_count = 0
        rev_count = 0
        new_count = 0
        total_count = 0

        # Method 1: Try AnkiConnect's native getDeckStats for target_deck and its subdecks
        try:
            available_decks = []
            res_names = AnkiBridge.invoke("deckNames")
            if res_names.get("success") and isinstance(res_names.get("result"), list):
                available_decks = res_names["result"]

            # Exact matching deck or subdecks (e.g. "MBBS Year 1" or "MBBS Year 1::...")
            matching_decks = [d for d in available_decks if d == target_deck or d.startswith(f"{target_deck}::")]

            # If target deck has no matching decks in Anki, try the root parent deck (e.g. "MBBS Year 1" or "MBBS Jassem A.")
            if not matching_decks:
                root_name = target_deck.split("::")[0].strip()
                matching_decks = [d for d in available_decks if d == root_name or d.startswith(f"{root_name}::")]

            if not matching_decks:
                default_parent = (self._config.get("DEFAULT_PARENT_DECK", "") or "MBBS Jassem A.").strip()
                default_root = default_parent.split("::")[0].strip()
                matching_decks = [d for d in available_decks if d == default_root or d.startswith(f"{default_root}::")]

            if matching_decks:
                stats_res = AnkiBridge.invoke("getDeckStats", decks=matching_decks)
                if stats_res.get("success") and isinstance(stats_res.get("result"), dict):
                    deck_stats_map = stats_res["result"]
                    for d_id, d_stat in deck_stats_map.items():
                        if isinstance(d_stat, dict):
                            due_count += int(d_stat.get("review_count", 0)) + int(d_stat.get("learn_count", 0))
                            new_count += int(d_stat.get("new_count", 0))
                            total_count += int(d_stat.get("total_in_deck", 0))
        except Exception as e:
            print(f"[AnkiStats] getDeckStats note: {e}")

        # Method 2: Wildcard findCards fallback if getDeckStats returned 0 for due cards
        if due_count == 0 or new_count == 0:
            queries_to_try = [
                f'deck:"{target_deck}*"',
                f'deck:"{target_deck.split("::")[0].strip()}*"',
                "" # global collection fallback
            ]

            for q_deck in queries_to_try:
                deck_prefix = f'{q_deck} ' if q_deck else ""

                if due_count == 0:
                    res_due = AnkiBridge.invoke("findCards", query=f'{deck_prefix}is:due')
                    if res_due.get("success") and isinstance(res_due.get("result"), list) and len(res_due["result"]) > 0:
                        due_count = len(res_due["result"])
                    else:
                        res_learn = AnkiBridge.invoke("findCards", query=f'{deck_prefix}is:learn -is:suspended')
                        if res_learn.get("success") and isinstance(res_learn.get("result"), list) and len(res_learn["result"]) > 0:
                            due_count = len(res_learn["result"])

                if new_count == 0:
                    res_new = AnkiBridge.invoke("findCards", query=f'{deck_prefix}is:new -is:suspended')
                    if res_new.get("success") and isinstance(res_new.get("result"), list) and len(res_new["result"]) > 0:
                        new_count = len(res_new["result"])

                if total_count == 0 and q_deck:
                    res_tot = AnkiBridge.invoke("findCards", query=f'{q_deck}')
                    if res_tot.get("success") and isinstance(res_tot.get("result"), list) and len(res_tot["result"]) > 0:
                        total_count = len(res_tot["result"])

                if due_count > 0 or new_count > 0:
                    break

        # Reviewed cards today across the user's active profile
        res_rev = AnkiBridge.invoke("findCards", query="rated:1")
        if res_rev.get("success") and isinstance(res_rev.get("result"), list):
            rev_count = len(res_rev["result"])

        return {
            "connected": True,
            "due_today": due_count,
            "reviewed_today": rev_count,
            "new_cards": new_count,
            "total_cards": total_count,
            "deck": target_deck
        }

    # ── Presentation Loading ─────────────────────────────────────────────────
    def open_pptx_dialog(self) -> dict:
        if not self._window:
            return {"success": False, "error": "No window"}
        configured_dir = str(self._config.get("LECTURES_DIR", "") or "").strip()
        if configured_dir and os.path.exists(os.path.expanduser(configured_dir)):
            default_dir = str(Path(configured_dir).expanduser().resolve())
        else:
            default_dir = str(Path.home())

        result = self._window.create_file_dialog(
            webview.FileDialog.OPEN,
            directory=default_dir,
            allow_multiple=False,
            file_types=('PowerPoint Presentations (*.pptx)', 'All files (*.*)')
        )
        if not result or len(result) == 0:
            return {"cancelled": True}

        pptx_path = Path(result[0])
        return self.load_pptx_file(str(pptx_path))

    def load_pptx_file(self, pptx_path_str: str) -> dict:
        try:
            pptx_path = Path(pptx_path_str)
            prs = Presentation(str(pptx_path))
            self.state.pptx_path = pptx_path
            self.state.current_lecture_name = pptx_path.stem
            self.state.presentation = prs
            self.state.total_slides = len(prs.slides)
            self.state.current_slide_idx = 0

            cand_deck = self._autodetect_deck_name(pptx_path)
            self.state.target_deck = cand_deck
            self._config.set("TARGET_DECK", cand_deck)

            # Preload slides and render with explicit presentation mapping
            self._render_engine.load_presentation(pptx_path, prs=prs)
            slide_data = self._load_slide(0)

            # Persist to RECENT_LECTURES in config
            recent = self._config.get("RECENT_LECTURES", [])
            if not isinstance(recent, list):
                recent = []
            recent = [r for r in recent if r.get("path") != str(pptx_path.resolve())]
            recent.insert(0, {
                "title": pptx_path.stem,
                "filename": pptx_path.name,
                "path": str(pptx_path.resolve()),
                "slide_count": self.state.total_slides,
                "last_opened": datetime.now().strftime("%d %b, %H:%M")
            })
            recent = recent[:10]
            self._config.set("RECENT_LECTURES", recent)
            self._config.save({"RECENT_LECTURES": recent})

            return {
                "success": True,
                "filename": pptx_path.name,
                "total_slides": self.state.total_slides,
                "target_deck": cand_deck,
                "current_slide": slide_data
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def _autodetect_deck_name(self, pptx_path: Path) -> str:
        available = AnkiBridge.get_deck_names() if AnkiBridge.is_connected() else []
        path_str = str(pptx_path)
        stem = pptx_path.stem
        parent_root = (self._config.get("DEFAULT_PARENT_DECK", "MBBS Year 1") or "MBBS Year 1").strip()

        module_subdecks = {
            "MCB":  f"{parent_root}::Year 1 Stage 1::FBS::A - MCB",
            "NAM":  f"{parent_root}::Year 1 Stage 1::FBS::B - NAM",
            "DH":   f"{parent_root}::Year 1 Stage 1::FBS::C - DH",
            "IAI":  f"{parent_root}::Year 1 Stage 1::FBS::D - IAI",
            "EAB":  f"{parent_root}::Year 1 Stage 1::FBS::E - EAB",
            "NBSS": f"{parent_root}::Year 1 Stage 1::FBS::F - NBSS",
            "GEN":  f"{parent_root}::Year 1 Stage 1::FBS::G - GEN",
            "FAH":  f"{parent_root}::Year 1 Stage 1::PAS::A - FAH",
            "FPP":  f"{parent_root}::Year 1 Stage 1::PAS::B - FPP",
            "ACRS": f"{parent_root}::Year 1 Stage 1::PAS::C - ACRS",
            "PCRS": f"{parent_root}::Year 1 Stage 1::PAS::D - PCRS",
            "AGRS": f"{parent_root}::Year 1 Stage 1::PAS::E - AGRS",
            "PGRS": f"{parent_root}::Year 1 Stage 1::PAS::F - PGRS",
            "AMS":  f"{parent_root}::Year 1 Stage 1::PAS::G - AMS",
            "CRANIAL": f"{parent_root}::Year 1 Stage 1::PAS::Z - Cranial Nerves",
            "HD":    f"{parent_root}::Year 2 Stage 2::FSTCP 1::HD",
            "FSTCP": f"{parent_root}::Year 2 Stage 2::FSTCP 1::HD",
        }

        if available:
            clean_stem = re.sub(r'^[0-9]+[.\-_ ]*', '', stem).strip().lower()
            for deck in available:
                if deck.startswith(parent_root):
                    deck_lower = deck.lower()
                    if clean_stem and len(clean_stem) > 4 and clean_stem in deck_lower:
                        return deck

        upper_text = f"{path_str} {stem}".upper()
        for mod_code, subdeck_prefix in module_subdecks.items():
            if re.search(rf"(?<![A-Z]){mod_code}(?![A-Z])", upper_text):
                if available:
                    for deck in available:
                        if deck.startswith(subdeck_prefix):
                            m_num = re.search(r'\b0?([1-9][0-9]?)\b', stem)
                            if m_num and (f"::{m_num.group(1)}." in deck or f"::0{m_num.group(1)}" in deck or f" {m_num.group(1)} " in deck):
                                return deck
                return f"{subdeck_prefix}::{stem}"

        return f"{parent_root}::Year 1 Stage 1::PAS::{stem}"

    # ── Slide & View Operations ──────────────────────────────────────────────
    def _clean_view_name(self, name: str) -> str:
        return name.replace("  [✓ Added]", "").strip()

    def _compute_slide_view_options(self) -> list[str]:
        s_idx = self.state.current_slide_idx
        options = []
        for v in getattr(self.state, "raw_slide_views", []):
            clean_v = self._clean_view_name(v)
            key = f"{s_idx}::{clean_v}"
            is_added = any(c.get("key") == key for c in self.state.queued_occlusion_cards)
            if is_added:
                options.append(f"{clean_v}  [✓ Added]")
            else:
                options.append(clean_v)
        return options or ["No Diagram Available"]

    def _create_slide_preview(self, slide, slide_idx: int) -> Image.Image:
        w, h = 1920, 1080
        im = Image.new("RGB", (w, h), (15, 23, 42))  # Slate-900 background
        draw = ImageDraw.Draw(im)

        # Top Accent gradient bar (Blue-600)
        draw.rectangle([(0, 0), (w, 10)], fill=(37, 99, 235))

        # Status Badge at top right
        is_rendering = getattr(self._render_engine, "is_rendering", False)
        badge_text = "⏳ RENDERING HIGH-RES DIAGRAM..." if is_rendering else "SLIDE TEXT PREVIEW"
        draw.rounded_rectangle([(80, 50), (480, 96)], radius=12, fill=(30, 41, 59), outline=(59, 130, 246), width=2)
        draw.text((100, 64), badge_text, fill=(96, 165, 250))

        # Extract text & title
        raw_text = SlideParsEngine.extract_slide_text(slide)
        title = ""
        body_lines = []

        lines = [line.strip() for line in raw_text.split("\n") if line.strip()]
        if lines:
            title = lines[0]
            body_lines = lines[1:]

        if not title:
            title = f"Slide {slide_idx + 1}"

        # Slide Number + Title
        draw.text((80, 130), f"Slide {slide_idx + 1}: {title[:75]}", fill=(248, 250, 252))

        # Subtitle message
        sub_msg = (
            "PowerPoint vector graphics & diagrams are exporting in background. "
            "High-resolution diagram view will appear automatically."
            if is_rendering
            else "Diagrams are being processed."
        )
        draw.text((80, 185), sub_msg, fill=(148, 163, 184))

        # Divider line
        draw.line([(80, 220), (w - 80, 220)], fill=(51, 65, 85), width=1)

        # Body text bullets
        y = 255
        for line in body_lines[:16]:
            if y > (h - 80):
                break
            clean_line = line[:110] + ("..." if len(line) > 110 else "")
            draw.text((80, y), f"•  {clean_line}", fill=(226, 232, 240))
            y += 44

        return im

    def _load_slide(self, slide_idx: int) -> dict:
        if not self.state.presentation or slide_idx < 0 or slide_idx >= self.state.total_slides:
            return {}

        self.state.undo_stack.clear()
        self.state.redo_stack.clear()
        self.state.current_slide_idx = slide_idx
        slide = self.state.presentation.slides[slide_idx]

        rendered = self._render_engine.get_slide_image(slide_idx, scale=2.0)
        self.state.rendered_slide_image = rendered

        diagrams = SlideParsEngine.extract_composite_diagrams(self.state.presentation, slide_idx, rendered)
        self.state.composite_diagrams = diagrams

        raw_imgs, raw_text = SlideParsEngine.extract_slide_data(slide)
        self.state.slide_images = raw_imgs
        self.state.slide_raw_text = raw_text

        raw_views = []
        if rendered:
            raw_views.append("Full Slide Diagram")
            for i, d in enumerate(diagrams):
                d_name = d["name"] if isinstance(d, dict) else f"Diagram {i + 1} (with Labels & Arrows)"
                raw_views.append(d_name)
        else:
            # Generate instant whole-slide preview while PowerPoint renders in background
            preview_img = self._create_slide_preview(slide, slide_idx)
            self.state.rendered_slide_image = preview_img
            raw_views.append("Full Slide Diagram (Rendering Vector Details...)")

        for i in range(len(raw_imgs)):
            raw_views.append(f"Extracted Image {i + 1}")

        self.state.raw_slide_views = raw_views
        options = self._compute_slide_view_options()

        # Always show whole slide first
        target_opt = options[0]
        return self._apply_view_change(target_opt)

    def _apply_view_change(self, mode: str) -> dict:
        clean_mode = self._clean_view_name(mode)
        self.state.current_view_mode = clean_mode

        if clean_mode.startswith("Full Slide Diagram") or clean_mode.startswith("Whole Slide"):
            self.state.active_image = self.state.rendered_slide_image
        elif clean_mode.startswith("Diagram "):
            matching_diag = None
            for d in getattr(self.state, "composite_diagrams", []):
                if isinstance(d, dict) and d.get("name") == clean_mode:
                    matching_diag = d
                    break
            if not matching_diag:
                try:
                    idx = int(clean_mode.split()[1]) - 1
                    if 0 <= idx < len(self.state.composite_diagrams):
                        matching_diag = self.state.composite_diagrams[idx]
                except Exception:
                    pass
            if matching_diag:
                self.state.active_image = matching_diag["image"] if isinstance(matching_diag, dict) else matching_diag
        elif clean_mode.startswith("Extracted Image "):
            try:
                idx = int(clean_mode.split()[2]) - 1
                if 0 <= idx < len(self.state.slide_images):
                    self.state.active_image = self.state.slide_images[idx]
            except Exception:
                pass

        if self.state.active_image is None and self.state.rendered_slide_image:
            self.state.active_image = self.state.rendered_slide_image
        elif self.state.active_image is None and self.state.slide_images:
            self.state.active_image = self.state.slide_images[0]

        # Restore from cache or auto-run OCR
        cache_key = (self.state.current_slide_idx, clean_mode)
        if cache_key in self.state.slide_ocr_cache:
            self.state.ocr_results = [dict(item) for item in self.state.slide_ocr_cache[cache_key]]
        else:
            if self._config.get("AUTO_RUN_OCR", False) and self._ocr_engine.is_ready and self.state.active_image:
                try:
                    boxes = self._ocr_engine.detect_text_boxes(self.state.active_image)
                    self.state.ocr_results = boxes
                    self.state.slide_ocr_cache[cache_key] = [dict(b) for b in boxes]
                except Exception:
                    self.state.ocr_results = []
            else:
                self.state.ocr_results = []

        self.state.selected_indices.clear()
        return self.get_slide_data(self.state.current_slide_idx)

    def get_slide_data(self, slide_idx: int = -1) -> dict:
        idx = slide_idx if slide_idx >= 0 else self.state.current_slide_idx
        img = self.state.active_image
        w, h = img.size if img else (1920, 1080)
        img_b64 = pil_to_base64_data_uri(img, format="PNG") if img else ""

        clean_view = self._clean_view_name(self.state.current_view_mode)
        note_key = f"{idx}::{clean_view}"
        is_queued = any(c.get("key") == note_key for c in self.state.queued_occlusion_cards)
        is_rendering = getattr(self._render_engine, "is_rendering", False)

        is_diagram = clean_view.startswith("Diagram ")
        diagram_box = None
        if is_diagram:
            for d in getattr(self.state, "composite_diagrams", []):
                if isinstance(d, dict) and d.get("name") == clean_view:
                    diagram_box = list(d.get("box", []))
                    break
            if not diagram_box:
                try:
                    s_idx = int(clean_view.split()[1]) - 1
                    if 0 <= s_idx < len(self.state.composite_diagrams):
                        d = self.state.composite_diagrams[s_idx]
                        if isinstance(d, dict):
                            diagram_box = list(d.get("box", []))
                except Exception:
                    pass

        full_img = self.state.rendered_slide_image
        full_w, full_h = full_img.size if full_img else (1920, 1080)

        return {
            "slide_idx": idx,
            "slide_num": idx + 1,
            "total_slides": self.state.total_slides,
            "view_options": self._compute_slide_view_options(),
            "current_view": self.state.current_view_mode,
            "image_data": img_b64,
            "image_width": w,
            "image_height": h,
            "ocr_results": self.state.ocr_results,
            "notes": self.state.slide_raw_text,
            "is_queued": is_queued,
            "is_rendering": is_rendering,
            "slide_title": f"Slide {idx + 1}",
            "status": "Rendering" if is_rendering else "Ready",
            "is_diagram_view": is_diagram,
            "diagram_box": diagram_box,
            "can_extend": is_diagram and (self.state.rendered_slide_image is not None),
            "slide_width": full_w,
            "slide_height": full_h
        }

    def change_view(self, view_mode: str) -> dict:
        return self._apply_view_change(view_mode)

    def goto_slide(self, slide_idx: int) -> dict:
        return self._load_slide(slide_idx)

    def get_slide(self, slide_idx: int = -1) -> dict:
        idx = slide_idx if slide_idx >= 0 else self.state.current_slide_idx
        return self._load_slide(idx)

    # ── OCR & Mask Editing ───────────────────────────────────────────────────
    def run_ocr(self) -> list[dict]:
        if self.state.active_image is None:
            print("[run_ocr] Warning: No active slide or diagram image loaded")
            return []
        if not self._ocr_engine.is_ready:
            err = getattr(self._ocr_engine, "_init_error", "OCR Engine still initializing")
            print(f"[run_ocr] Warning: OCR engine not ready yet ({err})")
            self._notify_frontend("ocr-status", {"ready": False, "message": err})
            return []
        boxes = self._ocr_engine.detect_text_boxes(self.state.active_image)
        self.state.ocr_results = boxes
        cache_key = (self.state.current_slide_idx, self._clean_view_name(self.state.current_view_mode))
        self.state.slide_ocr_cache[cache_key] = [dict(b) for b in boxes]
        return boxes

    def save_masks(self, masks: list[dict]) -> dict:
        self.state.ocr_results = masks
        cache_key = (self.state.current_slide_idx, self._clean_view_name(self.state.current_view_mode))
        self.state.slide_ocr_cache[cache_key] = [dict(b) for b in masks]
        return {"success": True}

    def extend_diagram_boundary(self, direction: str, amount: int = 60) -> dict:
        if not self.state.rendered_slide_image:
            return self.get_slide_data(self.state.current_slide_idx)

        clean_view = self._clean_view_name(self.state.current_view_mode)
        matching_diag = None
        for d in getattr(self.state, "composite_diagrams", []):
            if isinstance(d, dict) and d.get("name") == clean_view:
                matching_diag = d
                break

        if not matching_diag and clean_view.startswith("Diagram "):
            try:
                idx = int(clean_view.split()[1]) - 1
                if 0 <= idx < len(self.state.composite_diagrams):
                    d = self.state.composite_diagrams[idx]
                    if isinstance(d, dict):
                        matching_diag = d
            except Exception:
                pass

        if not matching_diag:
            return self.get_slide_data(self.state.current_slide_idx)

        rw, rh = self.state.rendered_slide_image.size
        rx1, ry1, rx2, ry2 = matching_diag.get("box", (0, 0, rw, rh))

        dx = 0
        dy = 0

        if direction == "left":
            new_rx1 = max(0, rx1 - amount)
            dx = rx1 - new_rx1
            rx1 = new_rx1
        elif direction == "right":
            rx2 = min(rw, rx2 + amount)
        elif direction == "top":
            new_ry1 = max(0, ry1 - amount)
            dy = ry1 - new_ry1
            ry1 = new_ry1
        elif direction == "bottom":
            ry2 = min(rh, ry2 + amount)
        elif direction == "all":
            pad_w = max(amount, int((rx2 - rx1) * 0.10))
            pad_h = max(amount, int((ry2 - ry1) * 0.10))
            new_rx1 = max(0, rx1 - pad_w)
            new_ry1 = max(0, ry1 - pad_h)
            dx = rx1 - new_rx1
            dy = ry1 - new_ry1
            rx1 = new_rx1
            ry1 = new_ry1
            rx2 = min(rw, rx2 + pad_w)
            ry2 = min(rh, ry2 + pad_h)
        elif direction == "shrink":
            pad_w = min(amount, int((rx2 - rx1) * 0.05))
            pad_h = min(amount, int((ry2 - ry1) * 0.05))
            new_rx1 = min(rx2 - 40, rx1 + pad_w)
            new_ry1 = min(ry2 - 40, ry1 + pad_h)
            dx = rx1 - new_rx1
            dy = ry1 - new_ry1
            rx1 = new_rx1
            ry1 = new_ry1
            rx2 = max(rx1 + 40, rx2 - pad_w)
            ry2 = max(ry1 + 40, ry2 - pad_h)

        if rx2 > rx1 and ry2 > ry1:
            new_crop = self.state.rendered_slide_image.crop((rx1, ry1, rx2, ry2))
            matching_diag["image"] = new_crop
            matching_diag["box"] = (rx1, ry1, rx2, ry2)
            self.state.active_image = new_crop

            # Shift existing masks so they stay aligned to anatomical structures
            if dx != 0 or dy != 0:
                for m in self.state.ocr_results:
                    b = m["box"]
                    m["box"] = (b[0] + dx, b[1] + dy, b[2] + dx, b[3] + dy)
                cache_key = (self.state.current_slide_idx, clean_view)
                self.state.slide_ocr_cache[cache_key] = [dict(b) for b in self.state.ocr_results]

        return self.get_slide_data(self.state.current_slide_idx)

    def update_diagram_bounds(self, new_box: list[int]) -> dict:
        if not self.state.rendered_slide_image or len(new_box) < 4:
            return self.get_slide_data(self.state.current_slide_idx)

        rw, rh = self.state.rendered_slide_image.size
        bx1 = max(0, min(rw, int(new_box[0])))
        by1 = max(0, min(rh, int(new_box[1])))
        bx2 = max(0, min(rw, int(new_box[2])))
        by2 = max(0, min(rh, int(new_box[3])))

        if bx2 <= bx1 or by2 <= by1:
            return self.get_slide_data(self.state.current_slide_idx)

        clean_view = self._clean_view_name(self.state.current_view_mode)
        matching_diag = None
        for d in getattr(self.state, "composite_diagrams", []):
            if isinstance(d, dict) and d.get("name") == clean_view:
                matching_diag = d
                break

        if not matching_diag and clean_view.startswith("Diagram "):
            try:
                idx = int(clean_view.split()[1]) - 1
                if 0 <= idx < len(self.state.composite_diagrams):
                    d = self.state.composite_diagrams[idx]
                    if isinstance(d, dict):
                        matching_diag = d
            except Exception:
                pass

        if matching_diag:
            old_rx1, old_ry1 = matching_diag.get("box", (0, 0))[:2]
            dx = old_rx1 - bx1
            dy = old_ry1 - by1

            new_crop = self.state.rendered_slide_image.crop((bx1, by1, bx2, by2))
            matching_diag["image"] = new_crop
            matching_diag["box"] = (bx1, by1, bx2, by2)
            self.state.active_image = new_crop

            if dx != 0 or dy != 0:
                for m in self.state.ocr_results:
                    b = m["box"]
                    m["box"] = (b[0] + dx, b[1] + dy, b[2] + dx, b[3] + dy)
                cache_key = (self.state.current_slide_idx, clean_view)
                self.state.slide_ocr_cache[cache_key] = [dict(b) for b in self.state.ocr_results]

        return self.get_slide_data(self.state.current_slide_idx)

    def create_custom_diagram(self, crop_box: list[int]) -> dict:
        if not self.state.rendered_slide_image or len(crop_box) < 4:
            return self.get_slide_data(self.state.current_slide_idx)

        rw, rh = self.state.rendered_slide_image.size
        clean_view = self._clean_view_name(self.state.current_view_mode)
        offset_x = 0
        offset_y = 0
        if clean_view.startswith("Diagram "):
            for d in getattr(self.state, "composite_diagrams", []):
                if isinstance(d, dict) and d.get("name") == clean_view:
                    offset_x, offset_y = d.get("box", (0, 0))[:2]
                    break

        x1 = max(0, min(rw, int(crop_box[0]) + offset_x))
        y1 = max(0, min(rh, int(crop_box[1]) + offset_y))
        x2 = max(0, min(rw, int(crop_box[2]) + offset_x))
        y2 = max(0, min(rh, int(crop_box[3]) + offset_y))

        if (x2 - x1) < 20 or (y2 - y1) < 20:
            return self.get_slide_data(self.state.current_slide_idx)

        cropped = self.state.rendered_slide_image.crop((x1, y1, x2, y2))

        existing_count = len(getattr(self.state, "composite_diagrams", []))
        diag_name = f"Diagram {existing_count + 1} (Selection)"

        new_diag = {
            "name": diag_name,
            "image": cropped,
            "box": (x1, y1, x2, y2),
            "is_custom": True
        }

        if not hasattr(self.state, "composite_diagrams") or self.state.composite_diagrams is None:
            self.state.composite_diagrams = []
        self.state.composite_diagrams.append(new_diag)

        # Insert into raw_slide_views right after existing Diagram options
        idx_to_insert = 1
        for i, v in enumerate(self.state.raw_slide_views):
            if v.startswith("Diagram "):
                idx_to_insert = i + 1
        self.state.raw_slide_views.insert(idx_to_insert, diag_name)

        return self._apply_view_change(diag_name)

    def get_full_slide_image(self) -> dict:
        img = self.state.rendered_slide_image
        if not img:
            return {"success": False, "image_data": "", "width": 0, "height": 0}
        return {
            "success": True,
            "image_data": pil_to_base64_data_uri(img, format="PNG"),
            "width": img.width,
            "height": img.height
        }

    def crop_diagram(self, crop_box: list[int]) -> dict:
        clean_view = self._clean_view_name(self.state.current_view_mode)
        if clean_view.startswith("Full Slide Diagram") or clean_view.startswith("Whole Slide"):
            return self.create_custom_diagram(crop_box)
        elif clean_view.startswith("Diagram "):
            offset_x, offset_y = 0, 0
            for d in getattr(self.state, "composite_diagrams", []):
                if isinstance(d, dict) and d.get("name") == clean_view:
                    offset_x, offset_y = d.get("box", (0, 0))[:2]
                    break
            abs_box = [crop_box[0] + offset_x, crop_box[1] + offset_y, crop_box[2] + offset_x, crop_box[3] + offset_y]
            return self.update_diagram_bounds(abs_box)
        else:
            return self.create_custom_diagram(crop_box)

    def reset_full_slide(self) -> dict:
        opts = getattr(self.state, "raw_slide_views", [])
        full_opt = opts[0] if opts else "Full Slide Diagram"
        return self._apply_view_change(full_opt)

    def remove_custom_diagram(self, diag_name: str, fallback_view: Optional[str] = None) -> dict:
        if hasattr(self.state, "composite_diagrams") and self.state.composite_diagrams:
            self.state.composite_diagrams = [
                d for d in self.state.composite_diagrams
                if not (isinstance(d, dict) and d.get("name") == diag_name)
            ]
        if diag_name in getattr(self.state, "raw_slide_views", []):
            try:
                self.state.raw_slide_views.remove(diag_name)
            except ValueError:
                pass

        clean_name = self._clean_view_name(diag_name)
        cache_key = (self.state.current_slide_idx, clean_name)
        self.state.slide_ocr_cache.pop(cache_key, None)

        views = getattr(self.state, "raw_slide_views", [])
        target_view = fallback_view if (fallback_view and fallback_view in views) else (views[0] if views else "Full Slide Diagram")
        return self._apply_view_change(target_view)

    def toggle_redaction(self, indices: list[int]) -> list[dict]:
        for idx in indices:
            if 0 <= idx < len(self.state.ocr_results):
                cur = self.state.ocr_results[idx].get("is_redaction", False)
                self.state.ocr_results[idx]["is_redaction"] = not cur
        self.save_masks(self.state.ocr_results)
        return self.state.ocr_results

    def merge_masks(self, indices: list[int]) -> list[dict]:
        valid = [self.state.ocr_results[i] for i in indices if 0 <= i < len(self.state.ocr_results)]
        if len(valid) < 2:
            return self.state.ocr_results

        x1 = min(b["box"][0] for b in valid)
        y1 = min(b["box"][1] for b in valid)
        x2 = max(b["box"][2] for b in valid)
        y2 = max(b["box"][3] for b in valid)
        joined_text = " ".join(b["text"] for b in sorted(valid, key=lambda b: (b["box"][1], b["box"][0])))

        new_mask = {
            "text": joined_text,
            "box": (x1, y1, x2, y2),
            "prob": 1.0,
            "checked": True,
            "is_redaction": any(b.get("is_redaction", False) for b in valid)
        }

        # Keep non-merged items and append new
        new_results = [b for i, b in enumerate(self.state.ocr_results) if i not in indices]
        new_results.append(new_mask)
        self.state.ocr_results = new_results
        self.save_masks(self.state.ocr_results)
        return self.state.ocr_results

    def auto_group_masks(self) -> list[dict]:
        boxes = list(self.state.ocr_results)
        if len(boxes) < 2:
            return boxes

        grouped = OCREngine.merge_nearby_label_boxes(boxes)
        self.state.ocr_results = grouped
        self.save_masks(grouped)
        return grouped

    def delete_masks(self, indices: list[int]) -> list[dict]:
        self.state.ocr_results = [b for i, b in enumerate(self.state.ocr_results) if i not in indices]
        self.save_masks(self.state.ocr_results)
        return self.state.ocr_results

    # ── Occlusion Queue & Export ─────────────────────────────────────────────
    def add_occlusion_to_deck(self, mode: str = "hide_all") -> dict:
        if self.state.active_image is None:
            return {"success": False, "error": "No active image"}

        checked = [r for r in self.state.ocr_results if r.get("checked", True)]
        if not checked:
            return {"success": False, "error": "No checked labels"}

        s_idx = self.state.current_slide_idx
        clean_view = self._clean_view_name(self.state.current_view_mode)
        note_key = f"{s_idx}::{clean_view}"
        lec_name = self.state.pptx_path.stem if self.state.pptx_path else "Lecture"

        cards = AnkiExporter.compile_occlusion_cards(
            base_img=self.state.active_image,
            ocr_boxes=self.state.ocr_results,
            mode=mode,
            lecture_name=lec_name,
            slide_num=s_idx + 1,
            notes=self.state.slide_raw_text
        )

        for c in cards:
            c["key"] = note_key
            c["slide_idx"] = s_idx

        # Replace or append
        self.state.queued_occlusion_cards = [c for c in self.state.queued_occlusion_cards if c.get("key") != note_key]
        self.state.queued_occlusion_cards.extend(cards)

        summary = self.get_queue_summary()
        return {
            "success": True,
            "message": f"Queued {len(checked)} labels ({mode}) ✓",
            "summary": summary
        }

    def get_queue_summary(self) -> dict:
        normal = len(self.state.normal_cards)
        occ_notes = len(self.state.queued_occlusion_cards)
        occ_cards = sum(c.get("card_count", 1) for c in self.state.queued_occlusion_cards)
        return {
            "normal_count": normal,
            "occlusion_notes_count": occ_notes,
            "occlusion_cards_count": occ_cards,
            "total_cards": normal + occ_cards
        }

    def get_queued_items(self) -> dict:
        items = []
        for c in self.state.normal_cards:
            items.append({
                "type": c.get("type", "1. Basic"),
                "header": c.get("front", ""),
                "details": c.get("back", ""),
                "slide": c.get("slide", 1),
            })
        for occ in self.state.queued_occlusion_cards:
            items.append({
                "type": f"IO ({occ.get('card_count', 1)} cards)",
                "header": occ.get("header") or f"Slide {occ.get('slide', 1)} Diagram",
                "details": f"{len(occ.get('targets', []))} occlusion labels",
                "slide": occ.get("slide", 1),
            })
        return {
            "total_count": len(items),
            "items": items
        }

    def push_to_anki(self) -> dict:
        target_deck = self.state.target_deck
        if not target_deck:
            return {"success": False, "error": "No target deck specified."}

        try:
            AnkiBridge.create_deck_if_missing(target_deck)

            # Query available models in Anki to dynamically adapt note types and field names
            available_models = []
            m_res = AnkiBridge.invoke("modelNames")
            if m_res.get("success") and isinstance(m_res.get("result"), list):
                available_models = m_res["result"]

            # Resolve Basic model
            basic_model = "1. Basic" if "1. Basic" in available_models else (
                "Basic" if "Basic" in available_models else None
            )
            if not basic_model:
                for m in available_models:
                    if "basic" in m.lower():
                        basic_model = m
                        break
            if not basic_model:
                basic_model = available_models[0] if available_models else "1. Basic"

            # Check Basic model field names
            b_fields_res = AnkiBridge.invoke("modelFieldNames", modelName=basic_model)
            b_field_names = b_fields_res.get("result", []) if b_fields_res.get("success") else ["Front", "Back"]
            b_front_field = "Front" if "Front" in b_field_names else (b_field_names[0] if b_field_names else "Front")
            b_back_field = "Back" if "Back" in b_field_names else (b_field_names[1] if len(b_field_names) > 1 else "Back")

            # Resolve Cloze model
            cloze_model = "Cloze++" if "Cloze++" in available_models else (
                "Cloze" if "Cloze" in available_models else None
            )
            if not cloze_model:
                for m in available_models:
                    if "cloze" in m.lower():
                        cloze_model = m
                        break
            if not cloze_model:
                cloze_model = "Cloze++"

            # Check Cloze model field names (e.g. "Back Extra" vs standard "Extra")
            c_fields_res = AnkiBridge.invoke("modelFieldNames", modelName=cloze_model)
            c_field_names = c_fields_res.get("result", []) if c_fields_res.get("success") else ["Text", "Back Extra"]
            c_text_field = "Text" if "Text" in c_field_names else (c_field_names[0] if c_field_names else "Text")
            c_extra_field = "Back Extra" if "Back Extra" in c_field_names else (
                "Extra" if "Extra" in c_field_names else (c_field_names[1] if len(c_field_names) > 1 else "Back Extra")
            )

            # Resolve Image Occlusion model
            io_model = "Image Occlusion++++++" if "Image Occlusion++++++" in available_models else (
                "Image Occlusion" if "Image Occlusion" in available_models else "Image Occlusion++++++"
            )

            notes_payload = []
            dup_options = {
                "allowDuplicate": True,
                "duplicateScope": "deck",
                "duplicateScopeOptions": {
                    "deckName": target_deck,
                    "checkChildren": False,
                    "checkAllModels": False
                }
            }

            for c in self.state.normal_cards:
                ctype = c.get("type", "1. Basic")
                front = re.sub(r"\s*\[Lecture:[^\]]*\]", "", c.get("front", "") or c.get("text", "")).strip()
                back = re.sub(r"\s*\[Lecture:[^\]]*\]", "", c.get("back", "") or c.get("extra", "")).strip()

                f_low = front.lower()
                if any(k in f_low for k in [
                    "why study", "why do we study", "benefit from the study",
                    "disciplines that benefit", "reasons for studying",
                    "learning objective", "learning outcome", "module code",
                    "who won", "nobel prize"
                ]):
                    continue

                if ctype in ("Cloze++", "Cloze"):
                    notes_payload.append({
                        "deckName": target_deck,
                        "modelName": cloze_model,
                        "fields": {c_text_field: front, c_extra_field: back},
                        "options": dup_options,
                        "tags": ["MBBS", "LectureCards"]
                    })
                else:
                    notes_payload.append({
                        "deckName": target_deck,
                        "modelName": basic_model,
                        "fields": {b_front_field: front, b_back_field: back},
                        "options": dup_options,
                        "tags": ["MBBS", "LectureCards"]
                    })

            uid = uuid.uuid4().hex[:8]
            for idx, occ in enumerate(self.state.queued_occlusion_cards):
                img_name = f"_io_{uid}_{idx:03d}.png"
                img_buf = io.BytesIO()
                occ["base_image"].save(img_buf, "PNG")
                AnkiBridge.store_media_file(img_name, img_buf.getvalue())

                notes_payload.append({
                    "deckName": target_deck,
                    "modelName": io_model,
                    "fields": {
                        "Occlusion": occ["occlusion_field"],
                        "Image": f'<img src="{img_name}">',
                        "Header": occ.get("header", ""),
                        "Back Extra": occ.get("back_extra", ""),
                        "Comments": "",
                    },
                    "options": {
                        "allowDuplicate": True,
                        "duplicateScope": "deck"
                    },
                    "tags": ["MBBS", "ImageOcclusion", f"Slide{occ.get('slide', 1)}"]
                })

            if not notes_payload:
                return {"success": False, "error": "No valid cards to push."}

            res = AnkiBridge.add_notes(notes_payload)
            results = res.get("result") or [] if res.get("success") else []
            created_count = sum(1 for nid in results if nid is not None)

            # If any cards were rejected by addNotes (e.g. duplicate or model mismatch),
            # retry note-by-note to add the valid ones and collect specific error reasons
            error_reasons = []
            if created_count < len(notes_payload):
                added_count = 0
                for idx, note in enumerate(notes_payload):
                    if idx < len(results) and results[idx] is not None:
                        added_count += 1
                        continue
                    single_res = AnkiBridge.invoke("addNote", note=note)
                    if single_res.get("success") and single_res.get("result"):
                        added_count += 1
                    else:
                        err = single_res.get("error", "Unknown AnkiConnect error")
                        if err and err not in error_reasons:
                            error_reasons.append(err)
                created_count = added_count

            if created_count == 0:
                err_msg = "; ".join(error_reasons[:2]) if error_reasons else (res.get("error") or "All cards were rejected by Anki.")
                return {"success": False, "error": f"Anki did not add any cards: {err_msg}"}

            # Force Anki desktop to refresh GUI deck browser so queue counts update immediately
            AnkiBridge.invoke("guiDeckBrowser")
            AnkiBridge.invoke("guiReset")
            AnkiBridge.sync()

            return {
                "success": True,
                "count": created_count,
                "total": len(notes_payload),
                "deck": target_deck
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def export_apkg(self) -> dict:
        try:
            deck_name = self.state.target_deck
            export_dir_str = str(self._config.get("EXPORTS_DIR", "") or "").strip()
            if not export_dir_str or export_dir_str in ("./anki_exports", "anki_exports"):
                exp_dir = EXPORTS_DIR
            else:
                exp_dir = Path(export_dir_str).expanduser().resolve()
            exp_dir.mkdir(parents=True, exist_ok=True)

            path = AnkiExporter.export_full_deck(
                deck_name=deck_name,
                normal_cards=self.state.normal_cards,
                occlusion_cards=self.state.queued_occlusion_cards,
                export_dir=exp_dir
            )
            return {"success": True, "path": str(path)}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    # ── Flashcard Studio AI ──────────────────────────────────────────────────
    def get_flashcards(self) -> list[dict]:
        return self.state.normal_cards

    def generate_slide_cards(self, slide_idx: int) -> list[dict]:
        if not self.state.presentation or slide_idx >= self.state.total_slides:
            return []
        if hasattr(self, "_cards_engine") and self._cards_engine:
            self._cards_engine._cards_cancel_requested = False
            self._cards_engine._cancel_requested = False
        slide = self.state.presentation.slides[slide_idx]
        _, text = SlideParsEngine.extract_slide_data(slide)
        lec = self.state.pptx_path.stem if self.state.pptx_path else "Lecture"
        cards = self._cards_engine.generate_cards_from_text(
            slide_text=text,
            slide_num=slide_idx + 1,
            lecture_name=lec,
            max_cards=3
        )
        self.state.normal_cards.extend(cards)
        return cards

    def generate_lecture_cards(self) -> list[dict]:
        if not self.state.presentation:
            return []
        if hasattr(self, "_cards_engine") and self._cards_engine:
            self._cards_engine._cards_cancel_requested = False
            self._cards_engine._cancel_requested = False
        lec = self.state.pptx_path.stem if self.state.pptx_path else "Lecture"
        slides = self.state.presentation.slides
        total_slides = len(slides)
        blueprint = self._cards_engine.build_lecture_blueprint(self.state.presentation, lec)

        # 1. Filter eligible content slides
        eligible_slides = []
        for i, slide in enumerate(slides):
            _, text = SlideParsEngine.extract_slide_data(slide)
            skip, _ = SlideParsEngine.should_skip_slide_for_cards(i, text)
            if not skip and text.strip():
                words = len(text.split())
                max_c = 4 if words > 120 else (3 if words > 60 else 2)
                eligible_slides.append({
                    "slide_num": i + 1,
                    "text": text,
                    "max_cards": max_c
                })

        if not eligible_slides:
            return self.state.normal_cards

        # 2. Batch slides into groups of 3 to stay safely under Gemini 15 RPM
        # (e.g. 30 slides -> ~7 API calls instead of 30, avoiding 429 quota exhaustion)
        batch_size = 3
        batches = [eligible_slides[j:j + batch_size] for j in range(0, len(eligible_slides), batch_size)]
        total_batches = len(batches)

        all_new = []
        self._notify_frontend("cards-progress", {
            "status": "start",
            "message": f"Analyzing {len(eligible_slides)} content slides in {total_batches} batches...",
            "current_slide": 0,
            "total_slides": total_slides,
            "generated": len(all_new)
        })

        for b_idx, batch in enumerate(batches):
            if getattr(self._cards_engine, "_cards_cancel_requested", False) or getattr(self._cards_engine, "_cancel_requested", False):
                self._notify_frontend("cards-progress", {
                    "status": "cancelled",
                    "message": f"Cancelled by user. Kept {len(all_new)} cards generated so far.",
                    "current_slide": batch[0]["slide_num"],
                    "total_slides": total_slides,
                    "generated": len(all_new)
                })
                return self.state.normal_cards

            start_s = batch[0]["slide_num"]
            end_s = batch[-1]["slide_num"]

            self._notify_frontend("cards-progress", {
                "status": "processing",
                "message": f"Generating cards for Slides {start_s}–{end_s} (Batch {b_idx + 1}/{total_batches})...",
                "current_slide": end_s,
                "total_slides": total_slides,
                "generated": len(all_new)
            })

            try:
                cards = self._cards_engine.generate_cards_from_slides_batch(
                    slides_batch=batch,
                    lecture_name=lec,
                    blueprint=blueprint,
                    status_callback=lambda msg, model=None: self._notify_frontend("cards-progress", {
                        "status": "waiting",
                        "message": msg,
                        "current_slide": end_s,
                        "total_slides": total_slides,
                        "generated": len(all_new)
                    })
                )
            except InterruptedError:
                self._notify_frontend("cards-progress", {
                    "status": "cancelled",
                    "message": f"Cancelled by user. Kept {len(all_new)} cards generated so far.",
                    "current_slide": end_s,
                    "total_slides": total_slides,
                    "generated": len(all_new)
                })
                return self.state.normal_cards

            if cards:
                all_new.extend(cards)
                self.state.normal_cards.extend(cards)
                self._notify_frontend("cards-progress", {
                    "status": "batch_done",
                    "message": f"Generated {len(all_new)} cards (Slide {end_s}/{total_slides})...",
                    "current_slide": end_s,
                    "total_slides": total_slides,
                    "generated": len(all_new),
                    "new_cards": cards
                })

            # Check again after batch completes
            if getattr(self._cards_engine, "_cards_cancel_requested", False) or getattr(self._cards_engine, "_cancel_requested", False):
                self._notify_frontend("cards-progress", {
                    "status": "cancelled",
                    "message": f"Cancelled by user. Kept {len(all_new)} cards generated so far.",
                    "current_slide": end_s,
                    "total_slides": total_slides,
                    "generated": len(all_new)
                })
                return self.state.normal_cards

            # Pacing delay between batches to stay comfortably below 15 RPM
            if b_idx < total_batches - 1:
                time.sleep(2.0)

        self._notify_frontend("cards-progress", {
            "status": "completed",
            "message": f"Completed! Generated {len(all_new)} cards.",
            "current_slide": total_slides,
            "total_slides": total_slides,
            "generated": len(all_new)
        })
        return self.state.normal_cards

    def cancel_cards_generation(self) -> dict:
        try:
            if hasattr(self, "_cards_engine") and self._cards_engine:
                self._cards_engine._cards_cancel_requested = True
            return {"success": True, "message": "Card generation cancelled"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def save_card(self, idx: int, card_dict: dict) -> dict:
        if 0 <= idx < len(self.state.normal_cards):
            self.state.normal_cards[idx] = card_dict
            return {"success": True}
        return {"success": False}

    def add_card(self, card_dict: dict) -> dict:
        self.state.normal_cards.append(card_dict)
        return {"success": True}

    def delete_card(self, idx: int) -> dict:
        if 0 <= idx < len(self.state.normal_cards):
            self.state.normal_cards.pop(idx)
            return {"success": True}
        return {"success": False}

    def clear_cards(self) -> dict:
        self.state.normal_cards.clear()
        return {"success": True}

    def chat_refine(self, instruction: str) -> list[dict]:
        updated = self._cards_engine.refine_cards_with_instruction(self.state.normal_cards, instruction)
        self.state.normal_cards = updated
        return updated

    # ── Slide Gallery ────────────────────────────────────────────────────────
    def get_gallery(self) -> list[dict]:
        if not self.state.presentation:
            return []
        items = []
        for i, slide in enumerate(self.state.presentation.slides):
            title = ""
            for sh in slide.shapes:
                if sh.has_text_frame and sh.text_frame.text.strip():
                    title = sh.text_frame.text.strip().split("\n")[0]
                    break

            has_diag = any(s.shape_type == 13 for s in slide.shapes)
            # Generate thumbnail: use vector rendered slide if ready, or instant slide preview fallback
            rendered = self._render_engine.get_slide_image(i, scale=1.0)
            if not rendered:
                rendered = self._create_slide_preview(slide, i)

            if rendered:
                thumb = rendered.copy()
                thumb.thumbnail((280, 160))
                thumb_b64 = pil_to_base64_data_uri(thumb, format="JPEG", quality=75)
            else:
                thumb_b64 = ""

            items.append({
                "slide_idx": i,
                "slide_num": i + 1,
                "title": title or f"Slide {i + 1}",
                "has_diagram": has_diag,
                "thumb_data": thumb_b64
            })
        return items

    # ── Lecture Notes ────────────────────────────────────────────────────────
    def get_lecture_notes(self) -> str:
        if not self.state.presentation:
            return ""
        notes = []
        for i, slide in enumerate(self.state.presentation.slides):
            _, text = SlideParsEngine.extract_slide_data(slide)
            if text.strip():
                notes.append(self._notes_engine.format_slide_notes(text, i + 1))
        return "\n\n---\n\n".join(notes)

    def get_notes_generation_status(self) -> dict:
        return getattr(self, "_notes_generation_status", {
            "active": False,
            "model": "",
            "message": "",
            "elapsed": 0
        })

    def cancel_ai_notes_generation(self) -> dict:
        try:
            engine = getattr(getattr(self, "_notes_engine", None), "_cards_engine", None) or getattr(self, "_cards_engine", None)
            if engine:
                engine._notes_cancel_requested = True
            self._notes_generation_status = {
                "active": False,
                "model": "",
                "message": "Cancelled by user",
                "elapsed": 0
            }
            return {"success": True, "message": "Generation cancelled"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def generate_ai_lecture_notes(self) -> dict:
        self.mark_user_activity()
        if not self.state.presentation:
            return {"success": False, "error": "No presentation loaded. Please open a lecture first."}

        engine = getattr(getattr(self, "_notes_engine", None), "_cards_engine", None) or getattr(self, "_cards_engine", None)
        if engine:
            engine._notes_cancel_requested = False
            engine._cancel_requested = False

        target_model = "gemini-4-hydrogen"
        start_time = time.time()
        self._notes_generation_status = {
            "active": True,
            "model": target_model,
            "message": f"Connecting to {target_model}...",
            "elapsed": 0
        }

        def _on_status(msg: str, model: str = None):
            cur_engine = getattr(getattr(self, "_notes_engine", None), "_cards_engine", None) or getattr(self, "_cards_engine", None)
            if cur_engine and (getattr(cur_engine, "_notes_cancel_requested", False) or getattr(cur_engine, "_cancel_requested", False)):
                raise InterruptedError("Generation cancelled by user.")
            self._notes_generation_status["message"] = msg
            if model:
                self._notes_generation_status["model"] = model
            self._notes_generation_status["elapsed"] = int(time.time() - start_time)

        try:
            lec_name = self.state.current_lecture_name or (self.state.pptx_path.stem if self.state.pptx_path else "Lecture")
            notes, model_used = self._notes_engine.generate_lecture_notes(
                self.state.presentation,
                lec_name,
                status_callback=_on_status
            )
            elapsed_sec = round(time.time() - start_time, 1)
            style_mode = self._config.get("NOTES_STYLE_MODE", "default")
            tpl_count = len(self._config.get("CUSTOM_NOTE_TEMPLATES", []) or [])
            self._notes_generation_status["active"] = False
            self._notes_generation_status["message"] = f"Finished in {elapsed_sec}s with {model_used}"

            # Automatically persist to in-app library categorized by Module & Block
            in_app_info = {}
            cached_diagrams = {}
            try:
                mod, blk = InAppNotesManager.detect_module_and_block(
                    lec_name, str(self.state.pptx_path or "")
                )
                save_res = self._in_app_notes_manager.save_note(
                    title=lec_name,
                    content=notes,
                    module=mod,
                    block=blk,
                    source_file=self.state.pptx_path.name if self.state.pptx_path else ""
                )
                if save_res.get("success"):
                    in_app_info = save_res
                    note_slug = save_res.get("note_id", "")
                    # Pre-extract all referenced diagrams and persist them to disk immediately
                    diag_dir = self._in_app_notes_manager.get_diagrams_dir(note_slug, mod, blk)
                    safe_src = self.state.pptx_path.stem if self.state.pptx_path else note_slug
                    slide_tags = re.findall(r"\(\s*slides?\s+([0-9\s,\-–and]+?):", notes, re.IGNORECASE)
                    all_s_nums = set()
                    for tag in slide_tags:
                        for num_match in re.findall(r"\d+", tag):
                            n = int(num_match)
                            if n > 1:
                                all_s_nums.add(n)
                    for s_num in all_s_nums:
                        try:
                            png_bytes = self._notion_engine.get_best_slide_image(self.state.presentation, s_num)
                            if png_bytes:
                                (diag_dir / f"slide_{s_num}.png").write_bytes(png_bytes)
                                (SLIDE_CACHE_DIR / f"{safe_src}_slide_{s_num}.png").write_bytes(png_bytes)
                                b64 = base64.b64encode(png_bytes).decode("ascii")
                                cached_diagrams[s_num] = f"data:image/png;base64,{b64}"
                        except Exception as d_err:
                            print(f"[generate_ai_notes] Diagram extraction error for slide {s_num}: {d_err}")
            except Exception as e:
                print(f"[generate_ai_notes] Local note auto-save note: {e}")

            return {
                "success": True,
                "notes": notes,
                "model_used": model_used or target_model,
                "style_mode": style_mode,
                "template_count": tpl_count,
                "elapsed": elapsed_sec,
                "in_app_saved": bool(in_app_info),
                "in_app_info": in_app_info,
                "diagrams": cached_diagrams
            }
        except InterruptedError:
            self._notes_generation_status["active"] = False
            self._notes_generation_status["message"] = "Cancelled by user"
            return {"success": False, "cancelled": True, "error": "Generation cancelled by user."}
        except Exception as exc:
            self._notes_generation_status["active"] = False
            self._notes_generation_status["message"] = f"Failed: {exc}"
            return {"success": False, "error": str(exc)}

    def save_lecture_notes(self, content: str, filename: str = "lecture-notes.md") -> dict:
        _ensure_app_data_dirs()
        out_path = APP_DATA_DIR / filename
        out_path.write_text(content, encoding="utf-8")
        return {"success": True, "filename": str(out_path.name), "path": str(out_path)}

    def save_in_app_note(self, title: str = "", content: str = "", module: str = "", block: str = "", source_file: str = "", note_id: str = "") -> dict:
        if not source_file and self.state.pptx_path:
            source_file = self.state.pptx_path.name
        if not title and not note_id:
            title = self.state.current_lecture_name or (self.state.pptx_path.stem if self.state.pptx_path else "")
        if not title and not note_id and content:
            heading_match = re.search(r"^\s*#+\s+(.+)$", content, re.MULTILINE)
            if heading_match:
                title = heading_match.group(1).strip()
        if not title and not note_id:
            title = "Lecture Note"
        return self._in_app_notes_manager.save_note(title, content, module=module, block=block, source_file=source_file, note_id=note_id)

    def get_in_app_notes(self, module: str = "", block: str = "", search: str = "") -> dict:
        return self._in_app_notes_manager.list_notes(module=module, block=block, search=search)

    def get_in_app_note(self, note_id: str) -> dict:
        return self._in_app_notes_manager.get_note(note_id)

    def delete_in_app_note(self, note_id: str) -> dict:
        return self._in_app_notes_manager.delete_note(note_id)

    def get_curriculum_structure(self) -> dict:
        return {"success": True, "structure": InAppNotesManager.KCL_STRUCTURE}

    def _find_presentation_for_note(self, note_id: str = "") -> tuple[Optional[Presentation], Optional[Path]]:
        """Try to locate and open the source PPTX presentation for a given note or current state."""
        source_file = ""
        title = ""
        if note_id:
            note_info = self._in_app_notes_manager.get_note(note_id)
            if note_info.get("success") and note_info.get("note"):
                source_file = (note_info["note"].get("source_file") or "").strip()
                title = (note_info["note"].get("title") or "").strip()

        # If active presentation matches the requested note, return it directly
        if self.state.presentation and self.state.pptx_path:
            active_name = self.state.pptx_path.name.lower()
            if not source_file or active_name == source_file.lower():
                return self.state.presentation, self.state.pptx_path
            # Extract lecture code (e.g. ACRS04L vs ACRS05L)
            act_code_m = re.search(r"\b([A-Za-z]{3,4}\d{2}[A-Za-z]?)\b", active_name)
            note_code_m = re.search(r"\b([A-Za-z]{3,4}\d{2}[A-Za-z]?)\b", f"{source_file} {title}")
            if act_code_m and note_code_m and act_code_m.group(1).lower() == note_code_m.group(1).lower():
                return self.state.presentation, self.state.pptx_path
            # Note is for a DIFFERENT presentation (e.g. note is ACRS05L while active presentation is ACRS04L).
            # Do NOT use self.state.presentation! Find the actual note's presentation on disk.
        elif self.state.presentation and not source_file and not title:
            return self.state.presentation, self.state.pptx_path

        if not source_file and self.state.pptx_path:
            source_file = self.state.pptx_path.name
        if not title:
            title = self.state.current_lecture_name or ""

        candidate_dirs = [
            Path(self._config.get("LECTURES_DIR", "./Lectures_OneDrive")).expanduser(),
            Path.home() / "Library/CloudStorage/OneDrive-King'sCollegeLondon",
            Path.home() / "Library/CloudStorage/OneDrive - King's College London",
            Path.home() / "OneDrive - King's College London",
            Path.home() / "Downloads",
            Path.home() / "Documents"
        ]
        cloud_storage = Path.home() / "Library/CloudStorage"
        if cloud_storage.exists():
            for child in cloud_storage.iterdir():
                if child.is_dir() and "onedrive" in child.name.lower() and child not in candidate_dirs:
                    candidate_dirs.append(child)

        found_path = None
        for base_dir in candidate_dirs:
            try:
                if not base_dir.exists():
                    continue
                if source_file:
                    for f in base_dir.rglob("*.pptx"):
                        if f.name.lower() == source_file.lower():
                            found_path = f
                            break
                    if found_path:
                        break
                if not found_path and title:
                    clean_title = re.sub(r"[^a-zA-Z0-9]+", " ", title).lower().strip()
                    tokens = [t for t in clean_title.split() if len(t) > 3]
                    for f in base_dir.rglob("*.pptx"):
                        f_clean = re.sub(r"[^a-zA-Z0-9]+", " ", f.stem).lower()
                        if tokens and all(tok in f_clean for tok in tokens[:3]):
                            found_path = f
                            break
                    if found_path:
                        break
            except Exception as e:
                print(f"[_find_presentation_for_note] Scan note: {e}")

        if found_path and found_path.exists():
            try:
                prs = Presentation(str(found_path))
                return prs, found_path
            except Exception as e:
                print(f"[_find_presentation_for_note] Failed opening {found_path}: {e}")

        return None, None

    def get_slide_diagram(self, slide_num: int, note_id: str = "", force: bool = False) -> dict:
        slide_int = int(slide_num)
        res = self.get_slide_diagrams([slide_int], note_id=note_id, force=force)
        if res.get("success") and res.get("diagrams") and slide_int in res["diagrams"]:
            return {"success": True, "slide_num": slide_int, "data_url": res["diagrams"][slide_int]}
        return {"success": False, "error": f"No diagram found on slide {slide_num}"}

    def get_slide_diagrams(self, slide_nums: list[int], note_id: str = "", force: bool = False) -> dict:
        res = {}
        missing_nums = []

        # 1. Tier 1: Check if already persisted on disk in note's diagrams directory or cache (unless force refresh)
        note_meta = {}
        diag_dir = None
        if note_id:
            n_res = self._in_app_notes_manager.get_note(note_id)
            if n_res.get("success") and n_res.get("note"):
                note_meta = n_res["note"]
                slug = note_meta.get("id", note_id)
                diag_dir = self._in_app_notes_manager.get_diagrams_dir(
                    slug, note_meta.get("module", ""), note_meta.get("block", ""), note_meta.get("md_path", "")
                )

        safe_src = Path(note_meta.get("source_file", "")).stem if note_meta.get("source_file") else ""
        if not safe_src and self.state.pptx_path:
            safe_src = self.state.pptx_path.stem

        for s in slide_nums:
            s_int = int(s)
            found = False
            if not force:
                # Check note diagrams folder on disk
                if diag_dir and (diag_dir / f"slide_{s_int}.png").exists():
                    try:
                        img_bytes = (diag_dir / f"slide_{s_int}.png").read_bytes()
                        b64 = base64.b64encode(img_bytes).decode("ascii")
                        res[s_int] = f"data:image/png;base64,{b64}"
                        found = True
                    except Exception:
                        pass

                # Check central slide cache on disk
                if not found and safe_src:
                    cache_file = SLIDE_CACHE_DIR / f"{safe_src}_slide_{s_int}.png"
                    if cache_file.exists():
                        try:
                            img_bytes = cache_file.read_bytes()
                            b64 = base64.b64encode(img_bytes).decode("ascii")
                            res[s_int] = f"data:image/png;base64,{b64}"
                            found = True
                            if diag_dir:
                                (diag_dir / f"slide_{s_int}.png").write_bytes(img_bytes)
                        except Exception:
                            pass

            if not found:
                missing_nums.append(s_int)

        if not missing_nums:
            return {"success": True, "diagrams": res}

        # 2. Tier 2: Locate presentation (active state or auto-found from disk)
        prs, prs_path = self._find_presentation_for_note(note_id)
        if not prs:
            return {"success": len(res) > 0, "diagrams": res}

        # Extract missing diagrams, persist to disk for eternity, and return
        for s_int in missing_nums:
            try:
                png_bytes = self._notion_engine.get_best_slide_image(prs, s_int, pptx_path=prs_path)
                if png_bytes:
                    b64 = base64.b64encode(png_bytes).decode("ascii")
                    res[s_int] = f"data:image/png;base64,{b64}"
                    # Persist to disk
                    if diag_dir:
                        (diag_dir / f"slide_{s_int}.png").write_bytes(png_bytes)
                    if safe_src:
                        (SLIDE_CACHE_DIR / f"{safe_src}_slide_{s_int}.png").write_bytes(png_bytes)
            except Exception as exc:
                print(f"[get_slide_diagrams] Error extracting slide {s_int}: {exc}")

        return {"success": True, "diagrams": res}

    def push_notes_to_notion(self, md_notes: str, target_page_id: str = "") -> dict:
        lec_name = self.state.current_lecture_name or (self.state.pptx_path.stem if self.state.pptx_path else "")
        if not lec_name and not target_page_id:
            return {"success": False, "error": "No active lecture selected. Please open a presentation first."}
        try:
            return self._notion_engine.push_notes(
                lec_name,
                md_notes,
                self.state.presentation,
                target_page_id=target_page_id
            )
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def detect_notion_target_page(self) -> dict:
        lec_name = self.state.current_lecture_name or (self.state.pptx_path.stem if self.state.pptx_path else "")
        if not lec_name:
            return {"success": False, "error": "No active lecture selected. Please open a presentation first."}
        try:
            match = self._notion_engine.search_session_page(lec_name)
            if not match:
                # Fetch candidate pages to let user choose
                fallback = self._notion_engine.search_all_notion_pages(query=lec_name, limit=10)
                return {
                    "success": True,
                    "lecture_name": lec_name,
                    "found": False,
                    "target": None,
                    "candidates": fallback
                }
            return {
                "success": True,
                "lecture_name": lec_name,
                "found": True,
                "target": {
                    "id": match["id"],
                    "title": match["title"],
                    "score": round(match.get("score", 1.0), 2),
                    "has_content": match.get("has_content", False)
                },
                "candidates": match.get("candidates", [])
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def search_notion_pages(self, query: str = "") -> dict:
        try:
            pages = self._notion_engine.search_all_notion_pages(query=query)
            return {"success": True, "pages": pages}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    # ── Lecture Notes Style & Template Personalization ───────────────────────
    def get_notes_style_info(self) -> dict:
        mode = self._config.get("NOTES_STYLE_MODE", "default")
        templates = self._config.get("CUSTOM_NOTE_TEMPLATES", []) or []
        instructions = self._config.get("CUSTOM_STYLE_INSTRUCTIONS", "") or ""
        return {
            "success": True,
            "mode": mode,
            "templates": templates,
            "instructions": instructions,
            "template_count": len(templates)
        }

    def set_notes_style_mode(self, mode: str) -> dict:
        target = "custom" if mode == "custom" else "default"
        self._config.set("NOTES_STYLE_MODE", target)
        templates = self._config.get("CUSTOM_NOTE_TEMPLATES", []) or []
        return {
            "success": True,
            "mode": target,
            "template_count": len(templates)
        }

    def pick_template_file(self) -> dict:
        if not self._window:
            return {"cancelled": True, "error": "App window not attached"}
        try:
            result = self._window.create_file_dialog(
                webview.FileDialog.OPEN,
                allow_multiple=False,
                file_types=('Note Files (*.md;*.txt;*.pdf;*.markdown)', 'All files (*.*)')
            )
            if not result or len(result) == 0:
                return {"cancelled": True}

            p = Path(result[0])
            if not p.exists():
                return {"success": False, "error": "Selected file does not exist"}

            text = _extract_text_from_file(p)
            if not text.strip():
                return {"success": False, "error": "Could not extract text from selected file or file is empty"}

            return {
                "success": True,
                "name": p.stem,
                "filename": p.name,
                "size_kb": round(p.stat().st_size / 1024, 1),
                "content": text
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def add_note_template(self, name: str, content: str, filename: str = "") -> dict:
        templates = list(self._config.get("CUSTOM_NOTE_TEMPLATES", []) or [])
        if len(templates) >= 3:
            return {"success": False, "error": "Maximum of 3 lecture templates allowed. Please delete an existing template first."}

        content_clean = content.strip()
        if not content_clean:
            return {"success": False, "error": "Template content cannot be empty."}

        tpl_id = f"tpl_{int(time.time())}_{uuid.uuid4().hex[:6]}"
        safe_name = name.strip() or f"Template {len(templates) + 1}"
        fname = filename.strip() or f"{safe_name}.md"

        # Save template text to TEMPLATES_DIR
        file_path = None
        try:
            TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
            safe_fname = re.sub(r'[^a-zA-Z0-9_\-\.]', '_', fname)
            file_path = TEMPLATES_DIR / f"{tpl_id}_{safe_fname}"
            file_path.write_text(content_clean, encoding="utf-8")
        except Exception as exc:
            print(f"[add_note_template] File save error: {exc}")

        words = len(content_clean.split())
        template_entry = {
            "id": tpl_id,
            "name": safe_name,
            "filename": fname,
            "path": str(file_path) if file_path else "",
            "size_kb": round(len(content_clean.encode("utf-8")) / 1024, 1),
            "word_count": words,
            "date_added": datetime.now().strftime("%d %b %Y, %H:%M"),
            "content": content_clean
        }

        templates.append(template_entry)
        self._config.set("CUSTOM_NOTE_TEMPLATES", templates)
        self._config.set("NOTES_STYLE_MODE", "custom")
        return {"success": True, "templates": templates, "template": template_entry, "mode": "custom"}

    def delete_note_template(self, template_id: str) -> dict:
        templates = list(self._config.get("CUSTOM_NOTE_TEMPLATES", []) or [])
        new_templates = []
        for t in templates:
            if t.get("id") == template_id:
                p_str = t.get("path")
                if p_str:
                    try:
                        p = Path(p_str)
                        if p.exists():
                            p.unlink()
                    except Exception:
                        pass
            else:
                new_templates.append(t)

        self._config.set("CUSTOM_NOTE_TEMPLATES", new_templates)
        current_mode = self._config.get("NOTES_STYLE_MODE", "default")
        if len(new_templates) == 0:
            current_mode = "default"
            self._config.set("NOTES_STYLE_MODE", "default")

        return {"success": True, "templates": new_templates, "mode": current_mode}

    def save_custom_style_instructions(self, instructions: str) -> dict:
        self._config.set("CUSTOM_STYLE_INSTRUCTIONS", instructions.strip())
        return {"success": True, "instructions": instructions.strip()}

    def analyze_template_style(self) -> dict:
        templates = self._config.get("CUSTOM_NOTE_TEMPLATES", []) or []
        if not templates:
            return {"success": False, "error": "Please upload at least 1 note template before analyzing style."}

        combined_samples = ""
        for idx, t in enumerate(templates[:3]):
            combined_samples += f"\n--- USER TEMPLATE {idx+1}: {t.get('name', 'Note')} ---\n"
            combined_samples += (t.get("content", "")[:8000]) + "\n"

        prompt = f"""You are an expert AI medical note stylist. Analyze these sample lecture notes produced by a student in Notion:

{combined_samples}

Summarize their personal note-taking style into 4 to 6 concise, actionable bullet points so that an AI can write future lecture notes in their exact personal format.
Cover:
- Title, headings, and topic formatting (e.g. use of bold/underlines, emojis, callouts)
- Structure & density (bullet lists vs short paragraphs, step-by-step pathways, tables)
- Tone & vocabulary (concise/telegrammatic vs explanatory, how jargon is handled)
- Highlighting & visual anchors

Output ONLY the bullet points, written directly as style rules."""

        try:
            cards_engine = GeminiCardsEngine(self._config)
            res = cards_engine._call_gemini_rest(
                prompt,
                timeout=45,
                preferred_model="gemini-3.5-flash-lite",
                task_label="Style Analysis"
            ).strip()
            self._config.set("CUSTOM_STYLE_INSTRUCTIONS", res)
            return {"success": True, "instructions": res}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    # ── Settings & Bridge ────────────────────────────────────────────────────
    def check_anki(self) -> dict:
        connected = AnkiBridge.is_connected()
        decks = AnkiBridge.get_deck_names() if connected else []
        self.state.anki_connected = connected
        self.state.available_decks = decks
        return {"connected": connected, "decks": decks}

    def check_ocr_status(self) -> dict:
        return {"ready": self._ocr_engine.is_ready}

    def test_anki_connection(self, url: str = "") -> dict:
        target_url = url.strip() or self._config.get("ANKI_URL", "http://127.0.0.1:8765")
        res = AnkiBridge.test_connection(target_url)
        if res.get("connected"):
            AnkiBridge.ANKI_URL = target_url
            self.state.anki_connected = True
            self.state.available_decks = AnkiBridge.get_deck_names()
            return {
                "success": True,
                "connected": True,
                "version": res.get("version"),
                "decks": self.state.available_decks
            }
        else:
            self.state.anki_connected = False
            return {
                "success": False,
                "connected": False,
                "error": res.get("error", "Could not connect to AnkiConnect")
            }

    def choose_folder(self, title: str = "Select Folder") -> dict:
        if not self._window:
            return {"cancelled": True}
        try:
            result = self._window.create_file_dialog(
                webview.FileDialog.FOLDER,
                allow_multiple=False
            )
            if not result or len(result) == 0:
                return {"cancelled": True}
            return {"success": True, "path": str(result[0])}
        except Exception as exc:
            return {"cancelled": True, "error": str(exc)}

    def set_native_theme(self, theme_mode: str, is_dark: bool = False) -> dict:
        if not HAS_COCOA:
            return {"success": True}
        def _apply_theme():
            try:
                app = AppKit.NSApplication.sharedApplication()
                if app:
                    if theme_mode == "dark":
                        app_appearance = NSAppearance.appearanceNamed_(NSAppearanceNameDarkAqua)
                    elif theme_mode == "light":
                        app_appearance = NSAppearance.appearanceNamed_(NSAppearanceNameAqua)
                    else:
                        # "system" mode: clear appearance override so NSApplication and WKWebView
                        # naturally track the macOS system appearance immediately.
                        app_appearance = None
                    app.setAppearance_(app_appearance)
            except Exception as e:
                print(f"[Theme] Cocoa appearance error: {e}")
        try:
            AppHelper.callAfter(_apply_theme)
            return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_system_is_dark(self) -> bool:
        if not HAS_COCOA:
            return False
        try:
            app = AppKit.NSApplication.sharedApplication()
            # Temporarily clear override to inspect true effective appearance if in system mode
            current_override = app.appearance()
            app.setAppearance_(None)
            eff = app.effectiveAppearance()
            matched = eff.bestMatchFromAppearancesWithNames_([
                NSAppearanceNameAqua,
                NSAppearanceNameDarkAqua
            ])
            is_dark = (matched == NSAppearanceNameDarkAqua)
            if current_override is not None:
                app.setAppearance_(current_override)
            return is_dark
        except Exception:
            try:
                defaults = AppKit.NSUserDefaults.standardUserDefaults()
                return defaults.stringForKey_("AppleInterfaceStyle") == "Dark"
            except Exception:
                return False

    def set_target_deck(self, deck_name: str) -> dict:
        self.state.target_deck = deck_name
        self._config.set("TARGET_DECK", deck_name)
        self._config.save({"TARGET_DECK": deck_name})
        return {"success": True}

    def get_config(self) -> dict:
        return self._config.load()

    def save_config(self, updates: dict) -> dict:
        if updates.get("APP_ACCENT") == "christmas" and not is_christmas_season():
            updates["APP_ACCENT"] = "cobalt"
        self._config.save(updates)
        if "TARGET_DECK" in updates:
            self.state.target_deck = updates["TARGET_DECK"]
        if "ANKI_URL" in updates:
            AnkiBridge.ANKI_URL = updates["ANKI_URL"]
        if "OCR_CONFIDENCE_THRESHOLD" in updates:
            try:
                self._ocr_engine.min_confidence = float(updates["OCR_CONFIDENCE_THRESHOLD"])
            except Exception:
                pass
        if "AUTO_PREPARE_NOTES" in updates:
            status = self.get_auto_prepare_status()
            self._notify_frontend("auto-prepare-update", status)
        return {"success": True, "config": self._config._data}

    # ── Announcements Endpoints ──────────────────────────────────────────────
    def get_announcements(self, force_refresh: bool = False) -> list[dict]:
        return self._announcements_engine.get_active_announcements(force_refresh=force_refresh)

    def dismiss_announcement(self, announcement_id: str) -> dict:
        self._config.dismiss_announcement(announcement_id)
        return {"success": True, "active": self._announcements_engine.get_active_announcements(force_refresh=False)}

    def is_christmas_season(self) -> bool:
        return is_christmas_season()

    # ── Supabase Auth & Onboarding Endpoints ─────────────────────────────────
    def login_synapse(self, email: str, password: str) -> dict:
        res = sign_in_synapse(email, password, config_path=CONFIG_PATH)
        if res.get("ok"):
            self._config.load()
        return res

    def signup_synapse(self, email: str, password: str) -> dict:
        res = sign_up_synapse(email, password, config_path=CONFIG_PATH)
        if res.get("ok"):
            self._config.load()
        return res

    def logout_synapse(self) -> dict:
        result = sign_out_synapse(config_path=CONFIG_PATH)
        self._config.load()
        return result

    def get_auth_session(self) -> dict:
        result = get_valid_session(config_path=CONFIG_PATH)
        cfg = self._config.load()
        return {**result, "onboarding_complete": bool(cfg.get("ONBOARDING_COMPLETE", False)), "questions_app_url": QUESTIONS_APP_URL}

    def accept_question_session(self, access_token: str, refresh_token: str) -> dict:
        result = accept_question_session(access_token, refresh_token, config_path=CONFIG_PATH)
        self._config.load()
        return result

    def complete_onboarding(self, updates: Optional[dict] = None) -> dict:
        payload = dict(updates or {})
        payload["ONBOARDING_COMPLETE"] = True
        return self.save_config(payload)

    def open_external_url(self, url: str) -> dict:
        target = (url or "").strip() or QUESTIONS_APP_URL
        try:
            webbrowser.open(target)
            return {"success": True, "url": target}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def quit_app(self) -> dict:
        if self._window:
            self._window.destroy()
        return {"success": True}

    # ── KCL OneDrive & AI Tutor Chat ─────────────────────────────────────────
    def get_kcl_lectures(self) -> list[dict]:
        self._index_engine.sync_index(force=False)
        return self._onedrive_engine.list_lectures()

    def get_kcl_blocks(self) -> dict:
        self._index_engine.sync_index(force=False)
        return self._onedrive_engine.get_blocks_grouped()

    def load_kcl_lecture(self, lecture_path: str) -> dict:
        return self.load_pptx_file(lecture_path)

    def chat_with_ai(self, message: str, scope: str = "slide", lecture_path: str = "", slide_idx: int = -1, history: list = None, block: str = "", mode: str = "quick") -> dict:
        return self._chatbot_engine.chat(message, scope=scope, lecture_path=lecture_path, slide_idx=slide_idx, history=history, block=block, mode=mode)

    def get_lecture_index_status(self) -> dict:
        return self._index_engine.get_status()

    def sync_lecture_index(self, force: bool = False) -> dict:
        res = self._index_engine.sync_index(force=force)
        self._notify_frontend("lecture-index-status", res)
        return res

    def open_slide_citation(self, filename: str, slide_num: int) -> dict:
        target_path = None
        p = Path(filename)
        if p.exists() and p.is_file():
            target_path = p
        else:
            lec = self._index_engine.find_lecture_by_filename(filename)
            if lec and Path(lec["file_path"]).exists():
                target_path = Path(lec["file_path"])

        if not target_path or not target_path.exists():
            return {"success": False, "error": f"Lecture presentation not found: '{filename}'"}

        target_slide_idx = max(0, int(slide_num) - 1)

        # If already loaded as active presentation
        if self.state.pptx_path and self.state.pptx_path.resolve() == target_path.resolve():
            target_slide_idx = min(target_slide_idx, self.state.total_slides - 1)
            slide_data = self._load_slide(target_slide_idx)
            return {
                "success": True,
                "filename": target_path.name,
                "total_slides": self.state.total_slides,
                "target_slide_idx": target_slide_idx,
                "target_deck": self.state.target_deck,
                "current_slide": slide_data
            }

        res = self.load_pptx_file(str(target_path))
        if res.get("success"):
            target_slide_idx = min(target_slide_idx, self.state.total_slides - 1)
            
            # If rendering in background, give PowerPoint export a few seconds to finish so the actual high-res slide appears immediately
            if getattr(self._render_engine, "is_rendering", False):
                t_end = time.time() + 6.0
                while time.time() < t_end and getattr(self._render_engine, "is_rendering", False):
                    time.sleep(0.25)

            slide_data = self._load_slide(target_slide_idx)
            res["target_slide_idx"] = target_slide_idx
            res["current_slide"] = slide_data
        return res

    def get_available_models(self) -> list[dict]:
        return [
            {"id": "gemini-4-hydrogen", "name": "gemini-4-hydrogen (Gemini 4 Deep Medical Synthesis · 20 req/day · Default for Notes)", "badge": "Gemini 4"},
            {"id": "gemini-3.8-flash", "name": "gemini-3.8-flash (High-Speed Multi-Step Reasoning · 20 req/day)", "badge": "Gemini 3.8"},
            {"id": "gemini-3-flash-preview", "name": "gemini-3-flash-preview (Gemini 3 Full Reasoning · 1,500 req/day)", "badge": "Gemini 3"},
            {"id": "gemini-3.5-flash-lite", "name": "gemini-3.5-flash-lite (Fastest · 1,500 req/day · Default for Quick Chat & Cards)", "badge": "Flash Lite 3.5"},
        ]

    def get_current_slide_screenshot(self) -> dict:
        """Returns the current slide's rendered diagram as a base64 data URI for feedback attachment."""
        try:
            if not self.state.presentation or self.state.current_slide_idx < 0:
                return {"success": False, "error": "No presentation or slide currently loaded."}

            img = self.state.rendered_slide_image
            if not img:
                img = self._render_engine.get_slide_image(self.state.current_slide_idx, scale=1.5)
            if not img and self.state.slide_images:
                img = self.state.slide_images[0]

            if not img:
                return {"success": False, "error": "Could not retrieve diagram image for current slide."}

            w, h = img.size
            if w > 1600 or h > 1200:
                img_copy = img.copy()
                img_copy.thumbnail((1600, 1200), Image.Resampling.LANCZOS)
            else:
                img_copy = img

            data_uri = pil_to_base64_data_uri(img_copy, format="PNG")
            slide_num = self.state.current_slide_idx + 1
            lec_name = self.state.pptx_path.name if self.state.pptx_path else "Presentation"
            return {
                "success": True,
                "data_uri": data_uri,
                "filename": f"{lec_name}_Slide_{slide_num}.png",
                "slide_num": slide_num,
                "lecture": lec_name
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_clipboard_image(self) -> dict:
        """Retrieves an image directly from the macOS system clipboard (e.g. from ⌘⌃⇧4, ⌘⇧4, or copied image)."""
        try:
            from PIL import ImageGrab
            content = ImageGrab.grabclipboard()
            if isinstance(content, Image.Image):
                buf = io.BytesIO()
                w, h = content.size
                if w > 1600 or h > 1200:
                    img_copy = content.copy()
                    img_copy.thumbnail((1600, 1200), Image.Resampling.LANCZOS)
                else:
                    img_copy = content
                img_copy.save(buf, format="PNG")
                b64 = base64.b64encode(buf.getvalue()).decode("ascii")
                return {
                    "success": True,
                    "data_uri": f"data:image/png;base64,{b64}",
                    "name": "clipboard_screenshot.png",
                    "size": f"{len(buf.getvalue()) / 1024:.1f} KB"
                }
            elif isinstance(content, list) and content:
                for item in content:
                    p = Path(str(item))
                    if p.exists() and p.suffix.lower() in (".png", ".jpg", ".jpeg", ".webp"):
                        img = Image.open(p)
                        buf = io.BytesIO()
                        img.save(buf, format="PNG")
                        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
                        return {
                            "success": True,
                            "data_uri": f"data:image/png;base64,{b64}",
                            "name": p.name,
                            "size": f"{p.stat().st_size / 1024:.1f} KB"
                        }
            if HAS_COCOA:
                try:
                    pb = AppKit.NSPasteboard.generalPasteboard()
                    types = pb.types() or []
                    for t in types:
                        t_str = str(t).lower()
                        if "png" in t_str or "tiff" in t_str or "pict" in t_str or "image" in t_str:
                            data = pb.dataForType_(t)
                            if data:
                                raw_bytes = bytes(data)
                                img = Image.open(io.BytesIO(raw_bytes))
                                buf = io.BytesIO()
                                img.save(buf, format="PNG")
                                b64 = base64.b64encode(buf.getvalue()).decode("ascii")
                                return {
                                    "success": True,
                                    "data_uri": f"data:image/png;base64,{b64}",
                                    "name": "clipboard_screenshot.png",
                                    "size": f"{len(buf.getvalue()) / 1024:.1f} KB"
                                }
                except Exception as cocoa_err:
                    print(f"Cocoa pasteboard fallback note: {cocoa_err}")

            return {"success": False, "error": "No image found in clipboard. Take a screenshot with ⌘⇧4 or ⌘⌃⇧4 first."}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def send_app_feedback(
        self,
        category: str = "Feature Suggestion",
        message: str = "",
        email: str = "",
        screenshot_base64: str = "",
        active_tab: str = "home",
        extra_diagnostics: dict = None
    ) -> dict:
        """Sends student feedback and optional screenshot attachment to Discord via Webhook."""
        import urllib.request
        import urllib.error
        import uuid
        from datetime import datetime, timezone

        webhook_url = (self._config.get("DISCORD_FEEDBACK_WEBHOOK") or "").strip()
        if not webhook_url:
            webhook_url = "https://discord.com/api/webhooks/1572979254625312929/3exv62rYExNtpoibQPrCCFOMkzipWFpRI5pZDzyyoprQbgNdTI3uP1M6nlH3DIKQf85K"

        msg_clean = (message or "").strip()
        if not msg_clean:
            return {"success": False, "error": "Feedback message cannot be empty."}

        category_clean = (category or "Feature Suggestion").strip()
        cat_lower = category_clean.lower()

        if "bug" in cat_lower or "error" in cat_lower or "crash" in cat_lower:
            color = 0xEF4444  # Red
            emoji = "🐛"
        elif "content" in cat_lower or "medical" in cat_lower or "card" in cat_lower:
            color = 0xF59E0B  # Amber
            emoji = "📚"
        elif "design" in cat_lower or "ui" in cat_lower or "appearance" in cat_lower:
            color = 0x8B5CF6  # Violet
            emoji = "🎨"
        else:
            color = 0x2563EB  # Cobalt Blue
            emoji = "💡"

        user_display = (email or "").strip()
        if not user_display:
            auth_email = self._config.get("AUTH_EMAIL") or ""
            user_name = self._config.get("USER_NAME") or ""
            if user_name and auth_email:
                user_display = f"{user_name} ({auth_email})"
            elif auth_email:
                user_display = auth_email
            elif user_name:
                user_display = user_name
            else:
                user_display = "Anonymous Student"

        footer_parts = [f"Synapse v{CURRENT_VERSION}", active_tab.upper()]
        if self.state.pptx_path:
            lec_name = self.state.pptx_path.name
            slide_no = self.state.current_slide_idx + 1
            footer_parts.append(f"{lec_name} (Slide {slide_no})")

        embed = {
            "title": f"{emoji} {category_clean} — {user_display}",
            "description": msg_clean[:2000],
            "color": color,
            "footer": {"text": " • ".join(footer_parts)},
            "timestamp": datetime.now(timezone.utc).isoformat(),
        }

        raw_screenshot = (screenshot_base64 or "").strip()
        img_bytes = None
        img_filename = "screenshot.png"
        mime_type = "image/png"

        if raw_screenshot:
            try:
                if "," in raw_screenshot:
                    header, b64data = raw_screenshot.split(",", 1)
                    if "image/jpeg" in header or "image/jpg" in header:
                        mime_type = "image/jpeg"
                        img_filename = "screenshot.jpg"
                    elif "image/webp" in header:
                        mime_type = "image/webp"
                        img_filename = "screenshot.webp"
                else:
                    b64data = raw_screenshot

                img_bytes = base64.b64decode(b64data)
                embed["image"] = {"url": f"attachment://{img_filename}"}
            except Exception as b64_err:
                print(f"Warning: Failed to decode feedback screenshot: {b64_err}")
                img_bytes = None

        payload = {
            "username": "Synapse Feedback",
            "embeds": [embed]
        }

        try:
            if img_bytes:
                boundary = f"----SynapseFormBoundary{uuid.uuid4().hex}"
                json_str = json.dumps(payload, ensure_ascii=False)
                body = (
                    f"--{boundary}\r\n"
                    f'Content-Disposition: form-data; name="payload_json"\r\n'
                    f'Content-Type: application/json\r\n\r\n'
                    f'{json_str}\r\n'
                    f"--{boundary}\r\n"
                    f'Content-Disposition: form-data; name="files[0]"; filename="{img_filename}"\r\n'
                    f'Content-Type: {mime_type}\r\n\r\n'
                ).encode("utf-8") + img_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")

                req = urllib.request.Request(
                    webhook_url,
                    data=body,
                    headers={
                        "Content-Type": f"multipart/form-data; boundary={boundary}",
                        "User-Agent": f"Synapse-Feedback/{CURRENT_VERSION}"
                    },
                    method="POST"
                )
            else:
                json_data = json.dumps(payload).encode("utf-8")
                req = urllib.request.Request(
                    webhook_url,
                    data=json_data,
                    headers={
                        "Content-Type": "application/json",
                        "User-Agent": f"Synapse-Feedback/{CURRENT_VERSION}"
                    },
                    method="POST"
                )

            with urllib.request.urlopen(req, timeout=15) as resp:
                if resp.status in (200, 204):
                    return {"success": True, "message": "Feedback submitted successfully! Thank you for helping shape Synapse."}
                else:
                    return {"success": False, "error": f"Discord responded with HTTP status {resp.status}"}
        except urllib.error.HTTPError as he:
            err_body = he.read().decode("utf-8", errors="ignore")
            print(f"Discord feedback HTTP error {he.code}: {err_body}")
            return {"success": False, "error": f"HTTP {he.code}: {he.reason}"}
        except Exception as e:
            print(f"Discord feedback error: {e}")
            return {"success": False, "error": str(e)}

# ─────────────────────────────────────────────────────────────────────────────
# 11. NATIVE MACOS MENU BAR & SHORTCUTS
# ─────────────────────────────────────────────────────────────────────────────
def build_native_menu(api: AppAPI) -> list[Menu]:
    def run_js(code: str):
        if api and api._window:
            try:
                api._window.evaluate_js(code)
            except Exception as err:
                print(f"Menu JS evaluation error: {err}")

    # 1. Synapse Application Menu (__app__)
    app_menu = Menu("__app__", [
        MenuAction("About Synapse", lambda: run_js("openAboutModal()")),
        MenuSeparator(),
        MenuAction("Settings…", lambda: run_js("openSettingsModal()")),
    ])

    # 2. File Menu
    file_menu = Menu("File", [
        MenuAction("Open .pptx Lecture…", lambda: run_js("openPptxFile()")),
        MenuSeparator(),
        MenuAction("Export .apkg…", lambda: run_js("exportApkgFile()")),
        MenuAction("Push Directly to Anki", lambda: run_js("pushDirectToAnki()")),
    ])

    # 3. Edit Menu (Standard macOS Edit menu routed to Cocoa First Responder for native clipboard)
    edit_menu = Menu("Edit", [
        MenuAction("Undo", lambda: run_js("smartUndo()")),
        MenuAction("Redo", lambda: run_js("smartRedo()")),
        MenuSeparator(),
        MenuAction("Cut", lambda: run_js("smartCut()")),
        MenuAction("Copy", lambda: run_js("smartCopy()")),
        MenuAction("Paste", lambda: run_js("smartPaste()")),
        MenuAction("Duplicate", lambda: run_js("duplicateSelectedMasks()")),
        MenuAction("Delete", lambda: run_js("deleteSelectedMasks()")),
        MenuSeparator(),
        MenuAction("Select All", lambda: run_js("smartSelectAll()")),
        MenuAction("Deselect All", lambda: run_js("deselectAllMasks()")),
    ])

    # 4. Canvas & OCR Edit Menu
    canvas_menu = Menu("Canvas", [
        MenuAction("Run OCR on Slide", lambda: run_js("runOcrOnDiagram()")),
        MenuSeparator(),
        MenuAction("Merge Selected Masks (M)", lambda: run_js("mergeSelectedMasks()")),
        MenuAction("Redact Selected (R)", lambda: run_js("toggleRedactionSelected()")),
        MenuAction("Auto-Group Labels (G)", lambda: run_js("autoGroupNearbyMasks()")),
        MenuSeparator(),
        MenuAction("Select All Masks", lambda: run_js("selectAllMasks(true)")),
        MenuAction("Delete Selected (⌫)", lambda: run_js("deleteSelectedMasks()")),
    ])

    # 5. View Menu
    view_menu = Menu("View", [
        MenuAction("Image Occlusion", lambda: run_js("switchTab('canvas')")),
        MenuAction("Flashcard Studio", lambda: run_js("switchTab('flashcards')")),
        MenuAction("Slide Gallery", lambda: run_js("switchTab('gallery')")),
        MenuAction("Lecture Notes", lambda: run_js("switchTab('notes')")),
        MenuAction("AI Slide & Lecture Chat", lambda: run_js("switchTab('chat')")),
        MenuAction("Questions Bank", lambda: run_js("switchTab('questions')")),
        MenuSeparator(),
        MenuAction("Toggle Dark / Light Mode", lambda: run_js("toggleThemeMode()")),
        MenuAction("Reset Zoom", lambda: run_js("zoomFit()")),
    ])

    # 6. Help Menu
    help_menu = Menu("Help", [
        MenuAction("Send Feedback & Screenshots…", lambda: run_js("openFeedbackModal()")),
        MenuSeparator(),
        MenuAction("Keyboard Shortcuts…", lambda: run_js("openShortcutsModal()")),
        MenuAction("About Synapse", lambda: run_js("openAboutModal()")),
    ])

    return [app_menu, file_menu, edit_menu, canvas_menu, view_menu, help_menu]

def setup_cocoa_shortcuts():
    if not HAS_COCOA:
        return
    try:
        app = AppKit.NSApplication.sharedApplication()
        main_menu = app.mainMenu() if app else None
        if not main_menu:
            return

        # Ensure top application menu displays Synapse
        if main_menu.numberOfItems() > 0:
            app_item = main_menu.itemAtIndex_(0)
            app_item.setTitle_("Synapse")
            app_sub = app_item.submenu()
            if app_sub:
                app_sub.setTitle_("Synapse")
                # Remove duplicate Cocoa default About panel item so only our custom in-app About modal remains
                if app_sub.numberOfItems() > 0 and str(app_sub.itemAtIndex_(0).action()) == "orderFrontStandardAboutPanel:":
                    app_sub.removeItemAtIndex_(0)
                    if app_sub.numberOfItems() > 0 and app_sub.itemAtIndex_(0).isSeparatorItem():
                        app_sub.removeItemAtIndex_(0)

        # Standard Cocoa First Responder selectors so Cmd+C, Cmd+V, Cmd+X, Cmd+A, Cmd+Z work natively in WKWebView inputs and canvas
        native_responder_actions = {
            "Undo": "undo:",
            "Redo": "redo:",
            "Cut": "cut:",
            "Copy": "copy:",
            "Paste": "paste:",
            "Select All": "selectAll:",
        }

        shortcuts = {
            "Settings…": (",", AppKit.NSEventModifierFlagCommand),
            "Open .pptx Lecture…": ("o", AppKit.NSEventModifierFlagCommand),
            "Export .apkg…": ("e", AppKit.NSEventModifierFlagCommand),
            "Push Directly to Anki": ("p", AppKit.NSEventModifierFlagCommand | AppKit.NSEventModifierFlagShift),
            "Undo": ("z", AppKit.NSEventModifierFlagCommand),
            "Redo": ("Z", AppKit.NSEventModifierFlagCommand | AppKit.NSEventModifierFlagShift),
            "Cut": ("x", AppKit.NSEventModifierFlagCommand),
            "Copy": ("c", AppKit.NSEventModifierFlagCommand),
            "Paste": ("v", AppKit.NSEventModifierFlagCommand),
            "Duplicate": ("d", AppKit.NSEventModifierFlagCommand),
            "Select All": ("a", AppKit.NSEventModifierFlagCommand),
            "Run OCR on Slide": ("r", AppKit.NSEventModifierFlagCommand),
            "Image Occlusion": ("1", AppKit.NSEventModifierFlagCommand),
            "Flashcard Studio": ("2", AppKit.NSEventModifierFlagCommand),
            "Slide Gallery": ("3", AppKit.NSEventModifierFlagCommand),
            "Lecture Notes": ("4", AppKit.NSEventModifierFlagCommand),
            "AI Slide & Lecture Chat": ("5", AppKit.NSEventModifierFlagCommand),
            "Questions Bank": ("6", AppKit.NSEventModifierFlagCommand),
            "Toggle Dark / Light Mode": ("d", AppKit.NSEventModifierFlagCommand | AppKit.NSEventModifierFlagShift),
            "Reset Zoom": ("0", AppKit.NSEventModifierFlagCommand),
            "Keyboard Shortcuts…": ("/", AppKit.NSEventModifierFlagCommand),
        }

        def walk_menu(menu):
            for item in menu.itemArray():
                title = str(item.title())
                if title in shortcuts:
                    key, mask = shortcuts[title]
                    item.setKeyEquivalent_(key)
                    item.setKeyEquivalentModifierMask_(mask)
                if title in native_responder_actions:
                    item.setTarget_(None)
                    item.setAction_(native_responder_actions[title])
                if item.hasSubmenu():
                    walk_menu(item.submenu())

        walk_menu(main_menu)

        # Register macOS wake observer to automatically refresh announcements
        global _WAKE_OBSERVER_INSTANCE
        if _WAKE_OBSERVER_INSTANCE is None:
            try:
                class SynapseWakeObserver(NSObject):
                    def workspaceDidWake_(self, notification):
                        if _GLOBAL_ANNOUNCEMENTS_ENGINE:
                            _GLOBAL_ANNOUNCEMENTS_ENGINE.fetch_async(callback_notify=True)

                _WAKE_OBSERVER_INSTANCE = SynapseWakeObserver.alloc().init()
                AppKit.NSWorkspace.sharedWorkspace().notificationCenter().addObserver_selector_name_object_(
                    _WAKE_OBSERVER_INSTANCE,
                    "workspaceDidWake:",
                    "NSWorkspaceDidWakeNotification",
                    None
                )
            except Exception as obs_err:
                pass
    except Exception as e:
        print(f"Cocoa menu shortcuts note: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# 12. MAIN APPLICATION ENTRYPOINT
# ─────────────────────────────────────────────────────────────────────────────
def main():
    if HAS_COCOA:
        try:
            NSProcessInfo.processInfo().setProcessName_("Synapse")
        except Exception:
            pass

        # Configure WebKit on macOS to allow third-party iframe storage persistence (for SBA question bank)
        try:
            from webview.platforms.cocoa import BrowserView
            _orig_bv_init = BrowserView.__init__
            def _patched_bv_init(self, window):
                _orig_bv_init(self, window)
                try:
                    if hasattr(self, "webview") and self.webview:
                        cfg = self.webview.configuration()
                        if cfg and hasattr(cfg, "preferences"):
                            prefs = cfg.preferences()
                            if hasattr(prefs, "_setStorageBlockingPolicy_"):
                                prefs._setStorageBlockingPolicy_(0)
                except Exception:
                    pass
            BrowserView.__init__ = _patched_bv_init
        except Exception:
            pass

    api = AppAPI()
    html_file = (BUNDLE_DIR / "gui" / "index.html").resolve()

    # If launched with an argument e.g. python3 anki_study_suite.py FAH05L.pptx
    if len(sys.argv) > 1 and sys.argv[1].endswith(".pptx"):
        target_path = Path(sys.argv[1]).resolve()
        if target_path.exists():
            api.load_pptx_file(str(target_path))

    window = webview.create_window(
        title="Synapse",
        url=str(html_file),
        js_api=api,
        width=1440,
        height=920,
        min_size=(1100, 750),
        background_color='#FAF9F6',
        text_select=True
    )
    api.set_window(window)

    def _ensure_webview_visible():
        try:
            from webview.platforms.cocoa import BrowserView
            for inst in BrowserView.instances.values():
                if inst and getattr(inst, "webview", None) and getattr(inst, "window", None):
                    if not inst.webview.window():
                        inst.window.setContentView_(inst.webview)
                        inst.window.makeFirstResponder_(inst.webview)
                    try:
                        cfg = inst.webview.configuration()
                        if cfg and hasattr(cfg, "preferences"):
                            prefs = cfg.preferences()
                            if hasattr(prefs, "_setStorageBlockingPolicy_"):
                                prefs._setStorageBlockingPolicy_(0)
                    except Exception:
                        pass
        except Exception:
            pass

    def on_started():
        time.sleep(0.4)
        if HAS_COCOA:
            AppHelper.callAfter(_ensure_webview_visible)
            AppHelper.callAfter(setup_cocoa_shortcuts)

    native_menus = build_native_menu(api)
    webview_storage = str(APP_DATA_DIR / "webview_data")
    Path(webview_storage).mkdir(parents=True, exist_ok=True)
    webview.start(
        func=on_started,
        menu=native_menus,
        http_server=True,
        http_port=42173,
        debug=False,
        private_mode=False,
        storage_path=webview_storage
    )

if __name__ == "__main__":
    main()
